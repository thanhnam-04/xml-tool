from __future__ import annotations

import argparse
import os
import subprocess
import tempfile
import urllib.request
import xml.etree.ElementTree as ET
from pathlib import Path

from PIL import ImageFont
from pptx import Presentation
from pptx.chart.data import ChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.util import Inches, Pt


class XMLSlideConverter:
    def __init__(
        self,
        xml_path: Path,
        output_path: Path,
        open_after_convert: bool = True,
    ) -> None:
        self.xml_path = xml_path.resolve()
        self.output_path = output_path.resolve()
        self.open_after_convert = open_after_convert
        self.base_dir = self.xml_path.parent

        self.prs = Presentation()
        self.canvas_width = 1600.0
        self.canvas_height = 900.0
        self.slide_width_inches = 13.333
        self.slide_height_inches = 7.5

        self.renderers = {
            "textbox": self.render_textbox,
            "shape": self.render_shape,
            "line": self.render_line,
            "polygon": self.render_polygon,
            "image": self.render_image,
            "svg": self.render_svg,
            "piechart": self.render_chart,
            "barchart": self.render_chart,
            "linechart": self.render_chart,
        }

    def svg_renderer_candidates(self) -> list[Path]:
        candidates = [
            Path(
                r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
            ),
            Path(
                r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"
            ),
            Path(
                r"C:\Program Files\Google\Chrome\Application\chrome.exe"
            ),
        ]

        local_appdata = os.environ.get("LOCALAPPDATA")
        if local_appdata:
            candidates.append(
                Path(local_appdata)
                / "Google/Chrome/Application/chrome.exe"
            )

        return [candidate for candidate in candidates if candidate.exists()]

    def close_open_output_presentation(self) -> None:
        if os.name != "nt":
            return

        script = r"""
$target = [System.IO.Path]::GetFullPath($env:PPTX_TO_CLOSE)

try {
    $app = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application")
} catch {
    exit 0
}

$closed = $false

foreach ($presentation in @($app.Presentations)) {
    try {
        $current = [System.IO.Path]::GetFullPath($presentation.FullName)
    } catch {
        continue
    }

    if ([string]::Equals($current, $target, [System.StringComparison]::OrdinalIgnoreCase)) {
        $presentation.Saved = $true
        $presentation.Close()
        $closed = $true
    }
}

if ($closed) {
    Write-Output "Da dong file PPTX dang mo: $target"
}
"""

        env = os.environ.copy()
        env["PPTX_TO_CLOSE"] = str(self.output_path)

        try:
            completed = subprocess.run(
                [
                    "powershell",
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-Command",
                    script,
                ],
                capture_output=True,
                env=env,
                text=True,
                timeout=10,
            )
        except (OSError, subprocess.TimeoutExpired) as exc:
            print(f"Khong the kiem tra file PPTX dang mo: {exc}")
            return

        if completed.stdout.strip():
            print(completed.stdout.strip())

        if completed.returncode != 0 and completed.stderr.strip():
            print(
                "Khong the dong file PPTX dang mo: "
                f"{completed.stderr.strip()}"
            )

    # =========================================================
    # Helpers
    # =========================================================

    @staticmethod
    def to_bool(value: str | None, default: bool = False) -> bool:
        if value is None:
            return default
        return value.strip().lower() in {"true", "1", "yes", "on"}

    @staticmethod
    def rgb(value: str | None, default: str = "#000000") -> RGBColor:
        raw = (value or default).strip().lstrip("#")

        if len(raw) == 3:
            raw = "".join(char * 2 for char in raw)

        if len(raw) != 6:
            raise ValueError(f"Màu không hợp lệ: {value}")

        return RGBColor(
            int(raw[0:2], 16),
            int(raw[2:4], 16),
            int(raw[4:6], 16),
        )

    def x(self, value: str | None) -> int:
        return round(
            float(value or 0)
            / self.canvas_width
            * self.prs.slide_width
        )

    def y(self, value: str | None) -> int:
        return round(
            float(value or 0)
            / self.canvas_height
            * self.prs.slide_height
        )

    @property
    def canvas_dpi_x(self) -> float:
        return self.canvas_width / self.slide_width_inches

    @property
    def canvas_dpi_y(self) -> float:
        return self.canvas_height / self.slide_height_inches

    def pt_to_canvas_x(self, point_size: float) -> float:
        return point_size * self.canvas_dpi_x / 72.0

    def pt_to_canvas_y(self, point_size: float) -> float:
        return point_size * self.canvas_dpi_y / 72.0

    # =========================================================
    # Font resolution and text measurement
    # =========================================================

    @staticmethod
    def normalize_font_name(name: str) -> str:
        return "".join(char.lower() for char in name if char.isalnum())

    def font_directories(self) -> list[Path]:
        directories: list[Path] = []

        windir = os.environ.get("WINDIR")
        if windir:
            directories.append(Path(windir) / "Fonts")

        directories.extend(
            [
                Path("/usr/share/fonts"),
                Path("/usr/local/share/fonts"),
                Path.home() / ".fonts",
                Path.home() / "Library/Fonts",
                Path("/Library/Fonts"),
            ]
        )

        return [path for path in directories if path.exists()]

    def resolve_font_path(
        self,
        font_name: str,
        bold: bool = False,
        italic: bool = False,
        explicit_path: str | None = None,
    ) -> Path:
        if explicit_path:
            path = Path(explicit_path)
            if not path.is_absolute():
                path = self.base_dir / path
            if not path.exists():
                raise FileNotFoundError(f"Không tìm thấy font_path: {path}")
            return path

        # Một số tên hiển thị trong PowerPoint không giống tên file font Windows.
        font_aliases = {
            "Segoe UI Symbol": "seguisym",
            "Segoe UI": "segoeui",
            "Arial": "arial",
            "Georgia": "georgia",
            "Times New Roman": "times",
            "Calibri": "calibri",
            "Aptos": "aptos",
        }

        alias = font_aliases.get(font_name)
        normalized_name = (
            self.normalize_font_name(alias)
            if alias
            else self.normalize_font_name(font_name)
        )

        style_tokens: list[str] = []
        if bold:
            style_tokens.extend(["bold", "bd", "semibold"])
        if italic:
            style_tokens.extend(["italic", "it", "oblique"])

        extensions = {".ttf", ".otf", ".ttc"}

        candidates: list[tuple[int, Path]] = []

        for directory in self.font_directories():
            try:
                font_files = directory.rglob("*")
            except OSError:
                continue

            for path in font_files:
                if path.suffix.lower() not in extensions:
                    continue

                stem = self.normalize_font_name(path.stem)
                score = 0

                if normalized_name == stem:
                    score += 100
                elif normalized_name in stem:
                    score += 70
                else:
                    continue

                if bold:
                    if any(token in stem for token in ("bold", "bd", "semibold")):
                        score += 20
                    else:
                        score -= 10

                if italic:
                    if any(token in stem for token in ("italic", "it", "oblique")):
                        score += 20
                    else:
                        score -= 10

                if not bold and any(token in stem for token in ("bold", "semibold")):
                    score -= 10

                if not italic and any(token in stem for token in ("italic", "oblique")):
                    score -= 10

                candidates.append((score, path))

        if candidates:
            candidates.sort(key=lambda item: item[0], reverse=True)
            return candidates[0][1]

        # Pillow thường đi kèm DejaVuSans; dùng làm fallback để converter không crash.
        fallback_names = [
            "DejaVuSans-BoldOblique.ttf" if bold and italic else "",
            "DejaVuSans-Bold.ttf" if bold else "",
            "DejaVuSans-Oblique.ttf" if italic else "",
            "DejaVuSans.ttf",
        ]

        for fallback in fallback_names:
            if not fallback:
                continue
            try:
                font = ImageFont.truetype(fallback, 12)
                font_path = getattr(font, "path", None)
                if font_path:
                    return Path(font_path)
            except OSError:
                continue

        raise ValueError(
            f"Không tìm thấy font '{font_name}'. "
            "Hãy cài font hoặc truyền font_path trong XML."
        )

    @staticmethod
    def load_font(font_path: Path, pixel_size: float) -> ImageFont.FreeTypeFont:
        return ImageFont.truetype(
            str(font_path),
            size=max(1, round(pixel_size)),
        )

    def measure_text_width_canvas(
        self,
        text: str,
        font_path: Path,
        font_size_pt: float,
    ) -> float:
        font = self.load_font(
            font_path,
            self.pt_to_canvas_y(font_size_pt),
        )

        if hasattr(font, "getlength"):
            return float(font.getlength(text))

        left, _, right, _ = font.getbbox(text)
        return float(right - left)

    def line_height_canvas(
        self,
        font_path: Path,
        font_size_pt: float,
        line_spacing: float,
    ) -> float:
        font = self.load_font(
            font_path,
            self.pt_to_canvas_y(font_size_pt),
        )
        ascent, descent = font.getmetrics()
        return float(ascent + descent) * line_spacing

    def break_long_word(
        self,
        word: str,
        font_path: Path,
        font_size_pt: float,
        max_width: float,
    ) -> list[str]:
        pieces: list[str] = []
        current = ""

        for char in word:
            candidate = current + char
            if (
                current
                and self.measure_text_width_canvas(
                    candidate,
                    font_path,
                    font_size_pt,
                )
                > max_width
            ):
                pieces.append(current)
                current = char
            else:
                current = candidate

        if current:
            pieces.append(current)

        return pieces or [""]

    def wrap_paragraph(
        self,
        text: str,
        font_path: Path,
        font_size_pt: float,
        max_width: float,
    ) -> list[str]:
        words = text.split()

        if not words:
            return [""]

        lines: list[str] = []
        current = ""

        for word in words:
            if (
                self.measure_text_width_canvas(
                    word,
                    font_path,
                    font_size_pt,
                )
                > max_width
            ):
                word_pieces = self.break_long_word(
                    word,
                    font_path,
                    font_size_pt,
                    max_width,
                )
            else:
                word_pieces = [word]

            for piece in word_pieces:
                candidate = piece if not current else f"{current} {piece}"

                if (
                    not current
                    or self.measure_text_width_canvas(
                        candidate,
                        font_path,
                        font_size_pt,
                    )
                    <= max_width
                ):
                    current = candidate
                else:
                    lines.append(current)
                    current = piece

        if current:
            lines.append(current)

        return lines

    def wrap_text(
        self,
        text: str,
        font_path: Path,
        font_size_pt: float,
        max_width: float,
    ) -> list[str]:
        # Giữ các xuống dòng mà XML đã chủ động đặt.
        source_paragraphs = text.splitlines() or [""]

        result: list[str] = []
        for paragraph in source_paragraphs:
            result.extend(
                self.wrap_paragraph(
                    paragraph.strip(),
                    font_path,
                    font_size_pt,
                    max_width,
                )
            )

        return result or [""]

    def calculate_text_layout(
        self,
        *,
        text: str,
        font_path: Path,
        initial_font_size: float,
        min_font_size: float,
        available_width: float,
        available_height: float,
        line_spacing: float,
        single_line: bool,
        max_lines: int | None,
        overflow: str,
        safety_ratio: float,
    ) -> tuple[list[str], float, float]:
        current_size = initial_font_size
        target_width = max(1.0, available_width * safety_ratio)
        target_height = max(1.0, available_height * safety_ratio)

        while current_size >= min_font_size:
            if single_line:
                normalized = " ".join(text.split())
                lines = [normalized]
            else:
                lines = self.wrap_text(
                    text,
                    font_path,
                    current_size,
                    target_width,
                )

            required_height = (
                len(lines)
                * self.line_height_canvas(
                    font_path,
                    current_size,
                    line_spacing,
                )
            )

            width_ok = all(
                self.measure_text_width_canvas(
                    line,
                    font_path,
                    current_size,
                )
                <= target_width
                for line in lines
            )
            height_ok = required_height <= target_height
            lines_ok = max_lines is None or len(lines) <= max_lines

            if width_ok and height_ok and lines_ok:
                return lines, current_size, required_height

            if overflow != "shrink":
                return lines, current_size, required_height

            current_size = round(current_size - 0.5, 2)

        final_size = min_font_size

        if single_line:
            lines = [" ".join(text.split())]
        else:
            lines = self.wrap_text(
                text,
                font_path,
                final_size,
                target_width,
            )

        required_height = (
            len(lines)
            * self.line_height_canvas(
                font_path,
                final_size,
                line_spacing,
            )
        )

        return lines, final_size, required_height

    # =========================================================
    # Textbox
    # =========================================================

    def render_textbox(self, slide, e: ET.Element) -> None:
        text = "".join(e.itertext()).strip()

        x = float(e.get("x", "0"))
        y = float(e.get("y", "0"))
        width = float(e.get("width", "300"))
        height = float(e.get("height", "100"))

        font_name = e.get("font", "Arial")
        font_size = float(e.get("size", "24"))
        min_font_size = float(e.get("min_font_size", "10"))

        bold = self.to_bool(e.get("bold"))
        italic = self.to_bool(e.get("italic"))
        single_line = self.to_bool(e.get("single_line"), False)
        reflow_on_resize = self.to_bool(
            e.get("reflow_on_resize"),
            False,
        )

        overflow = e.get("overflow", "shrink").lower()
        if overflow not in {"shrink", "resize", "clip"}:
            raise ValueError(
                "overflow phải là shrink, resize hoặc clip"
            )

        line_spacing = float(e.get("line_spacing", "1.0"))
        safety_ratio = float(e.get("safety_ratio", "0.97"))

        max_lines_value = e.get("max_lines")
        max_lines = (
            int(max_lines_value)
            if max_lines_value is not None
            else (1 if single_line else None)
        )

        margin_left_pt = float(e.get("margin_left", "0"))
        margin_right_pt = float(e.get("margin_right", "0"))
        margin_top_pt = float(e.get("margin_top", "0"))
        margin_bottom_pt = float(e.get("margin_bottom", "0"))

        available_width = (
            width
            - self.pt_to_canvas_x(margin_left_pt)
            - self.pt_to_canvas_x(margin_right_pt)
        )
        available_height = (
            height
            - self.pt_to_canvas_y(margin_top_pt)
            - self.pt_to_canvas_y(margin_bottom_pt)
        )

        font_path = self.resolve_font_path(
            font_name=font_name,
            bold=bold,
            italic=italic,
            explicit_path=e.get("font_path"),
        )

        lines, final_font_size, required_height = (
            self.calculate_text_layout(
                text=text,
                font_path=font_path,
                initial_font_size=font_size,
                min_font_size=min_font_size,
                available_width=available_width,
                available_height=available_height,
                line_spacing=line_spacing,
                single_line=single_line,
                max_lines=max_lines,
                overflow=overflow,
                safety_ratio=safety_ratio,
            )
        )

        final_height = height
        if overflow == "resize":
            final_height = max(
                height,
                required_height
                + self.pt_to_canvas_y(margin_top_pt)
                + self.pt_to_canvas_y(margin_bottom_pt),
            )

        box = slide.shapes.add_textbox(
            self.x(str(x)),
            self.y(str(y)),
            self.x(str(width)),
            self.y(str(final_height)),
        )

        tf = box.text_frame
        tf.clear()

        # single_line=true  -> tuyệt đối không xuống dòng.
        # single_line=false -> cho phép xuống dòng trong bounding box.
        # Converter vẫn tính trước các dòng; PowerPoint wrap là lớp dự phòng.
        tf.word_wrap = not single_line
        tf.auto_size = MSO_AUTO_SIZE.NONE

        tf.margin_left = Pt(margin_left_pt)
        tf.margin_right = Pt(margin_right_pt)
        tf.margin_top = Pt(margin_top_pt)
        tf.margin_bottom = Pt(margin_bottom_pt)

        vertical_map = {
            "top": MSO_ANCHOR.TOP,
            "middle": MSO_ANCHOR.MIDDLE,
            "center": MSO_ANCHOR.MIDDLE,
            "bottom": MSO_ANCHOR.BOTTOM,
        }
        tf.vertical_anchor = vertical_map.get(
            e.get("vertical_align", "top").lower(),
            MSO_ANCHOR.TOP,
        )

        align_map = {
            "left": PP_ALIGN.LEFT,
            "center": PP_ALIGN.CENTER,
            "right": PP_ALIGN.RIGHT,
            "justify": PP_ALIGN.JUSTIFY,
        }
        alignment = align_map.get(
            e.get("align", "left").lower(),
            PP_ALIGN.LEFT,
        )

        rendered_lines = lines
        if reflow_on_resize and not single_line:
            rendered_lines = [
                " ".join(paragraph.split())
                for paragraph in text.splitlines()
            ] or [""]

        for index, line in enumerate(rendered_lines):
            paragraph = (
                tf.paragraphs[0]
                if index == 0
                else tf.add_paragraph()
            )

            paragraph.text = line
            paragraph.alignment = alignment
            paragraph.space_before = Pt(0)
            paragraph.space_after = Pt(0)
            paragraph.line_spacing = line_spacing

            paragraph.font.name = font_name
            paragraph.font.size = Pt(final_font_size)
            paragraph.font.bold = bold
            paragraph.font.italic = italic
            paragraph.font.color.rgb = self.rgb(
                e.get("color"),
                "#000000",
            )

        print(
            f"  Textbox '{text[:35]}': "
            f"{font_size:g}pt -> {final_font_size:g}pt, "
            f"{len(lines)} dòng, "
            f"wrap={'off' if single_line else 'on'}, "
            f"reflow={'on' if reflow_on_resize else 'off'}"
        )

    # =========================================================
    # Shapes
    # =========================================================

    def render_shape(self, slide, e: ET.Element) -> None:
        mapping = {
            "rectangle": MSO_SHAPE.RECTANGLE,
            "rounded_rectangle": MSO_SHAPE.ROUNDED_RECTANGLE,
            "circle": MSO_SHAPE.OVAL,
            "ellipse": MSO_SHAPE.OVAL,
            "triangle": MSO_SHAPE.ISOSCELES_TRIANGLE,
            "diamond": MSO_SHAPE.DIAMOND,
            "hexagon": MSO_SHAPE.HEXAGON,
            "chevron": MSO_SHAPE.CHEVRON,
            "right_arrow": MSO_SHAPE.RIGHT_ARROW,
        }

        shape = slide.shapes.add_shape(
            mapping.get(
                e.get("type", "rectangle"),
                MSO_SHAPE.RECTANGLE,
            ),
            self.x(e.get("x")),
            self.y(e.get("y")),
            self.x(e.get("width")),
            self.y(e.get("height")),
        )

        fill = e.get("fill", "none")
        if fill == "none":
            shape.fill.background()
        else:
            shape.fill.solid()
            shape.fill.fore_color.rgb = self.rgb(fill)

        line_color = e.get("line_color", "none")
        if line_color == "none":
            shape.line.fill.background()
        else:
            shape.line.color.rgb = self.rgb(line_color)
            shape.line.width = Pt(
                float(e.get("line_width", "1"))
            )

        shape.rotation = float(e.get("rotation", "0"))

    def render_line(self, slide, e: ET.Element) -> None:
        line = slide.shapes.add_connector(
            MSO_CONNECTOR.STRAIGHT,
            self.x(e.get("x1")),
            self.y(e.get("y1")),
            self.x(e.get("x2")),
            self.y(e.get("y2")),
        )

        line.line.color.rgb = self.rgb(
            e.get("color"),
            "#000000",
        )
        line.line.width = Pt(float(e.get("width", "1")))

    def render_polygon(self, slide, e: ET.Element) -> None:
        points: list[tuple[int, int]] = []

        for item in e.get("points", "").split():
            px, py = item.split(",", maxsplit=1)
            points.append((self.x(px), self.y(py)))

        if len(points) < 3:
            raise ValueError("Polygon cần ít nhất 3 điểm")

        builder = slide.shapes.build_freeform(
            points[0][0],
            points[0][1],
        )
        builder.add_line_segments(points[1:], close=True)
        shape = builder.convert_to_shape()

        fill = e.get("fill", "none")
        if fill == "none":
            shape.fill.background()
        else:
            shape.fill.solid()
            shape.fill.fore_color.rgb = self.rgb(fill)

        line_color = e.get("line_color", "none")
        if line_color == "none":
            shape.line.fill.background()
        else:
            shape.line.color.rgb = self.rgb(line_color)
            shape.line.width = Pt(
                float(e.get("line_width", "1"))
            )

    # =========================================================
    # Images
    # =========================================================

    def download_image(self, url: str) -> Path:
        suffix = Path(url.split("?", maxsplit=1)[0]).suffix or ".jpg"

        temporary = tempfile.NamedTemporaryFile(
            delete=False,
            suffix=suffix,
        )
        temporary.close()
        target = Path(temporary.name)

        request = urllib.request.Request(
            url,
            headers={
                "User-Agent": (
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 Chrome/120 Safari/537.36"
                )
            },
        )

        try:
            with urllib.request.urlopen(
                request,
                timeout=30,
            ) as response:
                content_type = response.headers.get(
                    "Content-Type",
                    "",
                ).lower()

                if not content_type.startswith("image/"):
                    raise ValueError(
                        f"URL không trả về ảnh: {content_type}"
                    )

                target.write_bytes(response.read())

            return target
        except Exception:
            target.unlink(missing_ok=True)
            raise

    def render_image(self, slide, e: ET.Element) -> None:
        src = e.get("src")
        if not src:
            raise ValueError("Image thiếu src")

        temp_path: Path | None = None

        if src.startswith(("http://", "https://")):
            try:
                temp_path = self.download_image(src)
            except Exception as exc:
                print(f"  Bỏ qua ảnh tải lỗi: {src}")
                print(f"  Lý do: {exc}")
                return
            image_path = temp_path
        else:
            image_path = Path(src)
            if not image_path.is_absolute():
                image_path = self.base_dir / image_path

        try:
            if not image_path.exists():
                print(f"  Bỏ qua ảnh không tồn tại: {image_path}")
                return

            slide.shapes.add_picture(
                str(image_path),
                self.x(e.get("x")),
                self.y(e.get("y")),
                self.x(e.get("width")),
                self.y(e.get("height")),
            )
        finally:
            if temp_path is not None:
                temp_path.unlink(missing_ok=True)

    def render_svg(self, slide, e: ET.Element) -> None:
        svg_markup = "".join(e.itertext()).strip()
        if not svg_markup:
            raise ValueError("svg thieu noi dung")

        pixel_width = int(float(e.get("pixel_width", e.get("width", "1600"))))
        pixel_height = int(float(e.get("pixel_height", e.get("height", "900"))))

        browser_candidates = self.svg_renderer_candidates()
        if not browser_candidates:
            raise ValueError("Khong tim thay trinh duyet de render svg")

        svg_file = tempfile.NamedTemporaryFile(
            delete=False,
            suffix=".svg",
        )
        svg_file.close()
        svg_path = Path(svg_file.name)

        html_file = tempfile.NamedTemporaryFile(
            delete=False,
            suffix=".html",
        )
        html_file.close()
        html_path = Path(html_file.name)

        png_file = tempfile.NamedTemporaryFile(
            delete=False,
            suffix=".png",
        )
        png_file.close()
        png_path = Path(png_file.name)

        try:
            svg_path.write_text(
                svg_markup,
                encoding="utf-8",
            )

            html_path.write_text(
                (
                    "<!doctype html><html><head><meta charset=\"utf-8\">"
                    "<style>html,body{margin:0;padding:0;background:transparent;"
                    f"width:{pixel_width}px;height:{pixel_height}px;overflow:hidden;}}"
                    "img,svg{display:block;width:100%;height:100%;object-fit:fill;}"
                    "</style></head><body>"
                    f"<img src=\"file:///{svg_path.as_posix()}\" alt=\"svg\" />"
                    "</body></html>"
                ),
                encoding="utf-8",
            )

            subprocess.run(
                [
                    str(browser_candidates[0]),
                    "--headless",
                    "--disable-gpu",
                    "--default-background-color=00000000",
                    "--hide-scrollbars",
                    f"--window-size={pixel_width},{pixel_height}",
                    f"--screenshot={png_path}",
                    html_path.resolve().as_uri(),
                ],
                check=True,
                capture_output=True,
                text=True,
                timeout=30,
            )

            slide.shapes.add_picture(
                str(png_path),
                self.x(e.get("x")),
                self.y(e.get("y")),
                self.x(e.get("width")),
                self.y(e.get("height")),
            )
        finally:
            svg_path.unlink(missing_ok=True)
            html_path.unlink(missing_ok=True)
            png_path.unlink(missing_ok=True)

    # =========================================================
    # Charts
    # =========================================================

    def render_chart(self, slide, e: ET.Element) -> None:
        chart_map = {
            "piechart": XL_CHART_TYPE.PIE,
            "barchart": XL_CHART_TYPE.COLUMN_CLUSTERED,
            "linechart": XL_CHART_TYPE.LINE_MARKERS,
        }

        series_nodes = e.findall("./series")
        if not series_nodes:
            raise ValueError(f"{e.tag} thiếu series")

        first_points = series_nodes[0].findall("./point")
        categories = [
            point.get("label", "")
            for point in first_points
        ]

        data = ChartData()
        data.categories = categories

        for index, series in enumerate(series_nodes, start=1):
            values = [
                float(point.get("value", "0"))
                for point in series.findall("./point")
            ]
            data.add_series(
                series.get("name", f"Series {index}"),
                values,
            )

        frame = slide.shapes.add_chart(
            chart_map[e.tag],
            self.x(e.get("x")),
            self.y(e.get("y")),
            self.x(e.get("width")),
            self.y(e.get("height")),
            data,
        )

        chart = frame.chart
        chart.has_legend = self.to_bool(
            e.get("show_legend"),
            True,
        )

        if chart.has_legend:
            legend_map = {
                "bottom": XL_LEGEND_POSITION.BOTTOM,
                "top": XL_LEGEND_POSITION.TOP,
                "left": XL_LEGEND_POSITION.LEFT,
                "right": XL_LEGEND_POSITION.RIGHT,
            }
            chart.legend.position = legend_map.get(
                e.get("legend_position", "bottom"),
                XL_LEGEND_POSITION.BOTTOM,
            )

        plot = chart.plots[0]

        if self.to_bool(e.get("show_labels"), True):
            plot.has_data_labels = True
            labels = plot.data_labels
            labels.position = XL_LABEL_POSITION.BEST_FIT
            labels.show_percentage = (
                e.tag == "piechart"
                and self.to_bool(
                    e.get("show_percent"),
                    True,
                )
            )
            labels.show_value = self.to_bool(
                e.get("show_value"),
                e.tag != "piechart",
            )

        if e.tag == "piechart":
            for index, point_node in enumerate(first_points):
                color = point_node.get("color")
                if color:
                    point = chart.series[0].points[index]
                    point.format.fill.solid()
                    point.format.fill.fore_color.rgb = self.rgb(
                        color
                    )

    # =========================================================
    # Conversion
    # =========================================================

    def convert(self) -> None:
        if not self.xml_path.exists():
            raise FileNotFoundError(
                f"Không tìm thấy XML: {self.xml_path}"
            )

        root = ET.parse(self.xml_path).getroot()

        if root.tag != "presentation":
            raise ValueError("Thẻ gốc phải là <presentation>")

        self.canvas_width = float(root.get("width", "1600"))
        self.canvas_height = float(root.get("height", "900"))
        self.slide_width_inches = float(
            root.get("slideWidthInches", "13.333")
        )
        self.slide_height_inches = float(
            root.get("slideHeightInches", "7.5")
        )

        self.prs.slide_width = Inches(self.slide_width_inches)
        self.prs.slide_height = Inches(self.slide_height_inches)

        slide_nodes = root.findall("./slide")
        if not slide_nodes:
            raise ValueError("XML không có thẻ <slide>")

        for slide_index, slide_node in enumerate(
            slide_nodes,
            start=1,
        ):
            print(f"Đang tạo slide {slide_index}")

            slide = self.prs.slides.add_slide(
                self.prs.slide_layouts[6]
            )

            slide.background.fill.solid()
            slide.background.fill.fore_color.rgb = self.rgb(
                slide_node.get("background"),
                "#FFFFFF",
            )

            for element in slide_node:
                renderer = self.renderers.get(element.tag)
                if renderer is None:
                    raise ValueError(
                        f"Chưa hỗ trợ thẻ <{element.tag}>"
                    )
                renderer(slide, element)

        self.close_open_output_presentation()
        self.prs.save(self.output_path)
        if self.open_after_convert:
            try:
                os.startfile(str(self.output_path))
            except OSError as exc:
                print(f"Khong the mo file: {exc}")
        print(f"Đã tạo: {self.output_path}")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Chuyển XML slide sang PowerPoint"
    )
    parser.add_argument("xml")
    parser.add_argument(
        "-o",
        "--output",
        default="output.pptx",
    )
    parser.add_argument(
        "--no-open",
        action="store_true",
        help="Khong tu mo file PPTX sau khi convert",
    )

    args = parser.parse_args()

    output_path = Path(args.output)

    # python-pptx chỉ tạo định dạng .pptx.
    if output_path.suffix.lower() != ".pptx":
        output_path = output_path.with_suffix(".pptx")

    XMLSlideConverter(
        Path(args.xml),
        output_path,
        open_after_convert=not args.no_open,
    ).convert()


if __name__ == "__main__":
    main()
