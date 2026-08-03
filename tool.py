from __future__ import annotations

import os
import xml.etree.ElementTree as ET
from pathlib import Path
from xml.dom import minidom

from PIL import ImageFont


def generate_it_sales_proposal_cover_inline_vector_xml(
    brand_name: str = "FORGE",
    title: str = "IT Software Sales Proposal",
    subtitle: str = "Modern productivity for secure, high-performing teams",
    prepared_for: str = "Value Financials",
    date_text: str = "June 2035",
    background_color: str = "#1A0B3A",
    title_color: str = "#F7F2FF",
    subtitle_color: str = "#FFFFFF",
    subtitle_bar_color: str = "#6545A2",
    brand_color: str = "#F7F2FF",
    footer_color: str = "#F7F2FF",
) -> str:
    """
    Generate XML for a premium software proposal cover slide.

    The layout features an inline vector ribbon decoration on the left,
    a compact brand mark near the top-right, a large proposal title,
    a horizontal subtitle bar, and two footer labels along the bottom.
    """

    if not title.strip():
        raise ValueError("title must not be empty")

    if not subtitle.strip():
        raise ValueError("subtitle must not be empty")

    # =====================================================
    # Canvas and layout regions
    # =====================================================

    canvas_width = 1600
    canvas_height = 900

    decoration_x = 0.0
    decoration_y = 0.0
    decoration_width = 760.0
    decoration_height = 900.0

    content_x = 815.0
    content_right = 1530.0
    content_width = content_right - content_x

    brand_y = 105.0
    brand_icon_width = 58.0
    brand_gap = 18.0

    main_top = 245.0
    main_bottom = 700.0
    available_main_height = main_bottom - main_top

    title_to_subtitle_gap = 42.0
    minimum_gap = 20.0

    footer_y = 825.0
    footer_left_x = 815.0
    footer_left_width = 470.0
    footer_right_x = 1330.0
    footer_right_width = 200.0


    # =====================================================
    # Font measurement helpers
    # =====================================================

    def normalize_font_name(name: str) -> str:
        return "".join(
            character.lower()
            for character in name
            if character.isalnum()
        )

    def font_directories() -> list[Path]:
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

        return [
            directory
            for directory in directories
            if directory.exists()
        ]

    def resolve_font_path(
        font_name: str,
        bold: bool = False,
    ) -> Path:
        aliases = {
            "Arial": "arial",
            "Calibri": "calibri",
            "Aptos": "aptos",
            "Helvetica": "helvetica",
        }

        normalized = normalize_font_name(
            aliases.get(font_name, font_name)
        )

        candidates: list[tuple[int, Path]] = []

        for directory in font_directories():
            try:
                files = directory.rglob("*")
            except OSError:
                continue

            for path in files:
                if path.suffix.lower() not in {
                    ".ttf",
                    ".otf",
                    ".ttc",
                }:
                    continue

                stem = normalize_font_name(path.stem)

                if normalized not in stem:
                    continue

                score = 70

                if stem == normalized:
                    score += 30

                has_bold = any(
                    token in stem
                    for token in (
                        "bold",
                        "bd",
                        "semibold",
                    )
                )

                if bold and has_bold:
                    score += 20
                elif bold:
                    score -= 10
                elif has_bold:
                    score -= 10

                candidates.append((score, path))

        if candidates:
            candidates.sort(
                key=lambda item: item[0],
                reverse=True,
            )
            return candidates[0][1]

        fallback = (
            "DejaVuSans-Bold.ttf"
            if bold
            else "DejaVuSans.ttf"
        )

        font = ImageFont.truetype(fallback, 12)
        font_path = getattr(font, "path", None)

        if not font_path:
            raise ValueError(
                f"Font '{font_name}' was not found"
            )

        return Path(font_path)

    canvas_dpi = 120.0

    def pt_to_canvas(point_size: float) -> float:
        return point_size * canvas_dpi / 72.0

    def load_font(
        font_path: Path,
        font_size_pt: float,
    ) -> ImageFont.FreeTypeFont:
        return ImageFont.truetype(
            str(font_path),
            max(
                1,
                round(pt_to_canvas(font_size_pt)),
            ),
        )

    def measure_text_width(
        text: str,
        font_path: Path,
        font_size_pt: float,
    ) -> float:
        font = load_font(
            font_path,
            font_size_pt,
        )

        if hasattr(font, "getlength"):
            return float(font.getlength(text))

        left, _, right, _ = font.getbbox(text)
        return float(right - left)

    def measure_line_height(
        font_path: Path,
        font_size_pt: float,
        line_spacing: float,
    ) -> float:
        font = load_font(
            font_path,
            font_size_pt,
        )
        ascent, descent = font.getmetrics()

        return (
            float(ascent + descent)
            * line_spacing
        )

    def wrap_text(
        text: str,
        font_path: Path,
        font_size_pt: float,
        max_width: float,
    ) -> list[str]:
        result: list[str] = []

        for paragraph in text.splitlines() or [text]:
            words = paragraph.split()

            if not words:
                result.append("")
                continue

            current = ""

            for word in words:
                word_width = measure_text_width(
                    word,
                    font_path,
                    font_size_pt,
                )

                if word_width > max_width:
                    raise ValueError(
                        f"The word '{word}' is wider than "
                        "the available text area"
                    )

                candidate = (
                    word
                    if not current
                    else f"{current} {word}"
                )

                if (
                    not current
                    or measure_text_width(
                        candidate,
                        font_path,
                        font_size_pt,
                    )
                    <= max_width
                ):
                    current = candidate
                else:
                    result.append(current)
                    current = word

            if current:
                result.append(current)

        return result or [""]

    def layout_text(
        text: str,
        *,
        font_path: Path,
        font_size: float,
        max_width: float,
        line_spacing: float,
        padding: float = 3.0,
    ) -> dict:
        lines = wrap_text(
            text,
            font_path,
            font_size,
            max_width,
        )

        line_height = measure_line_height(
            font_path,
            font_size,
            line_spacing,
        )

        return {
            "source_text": text,
            "text": "\n".join(lines),
            "lines": lines,
            "font_size": font_size,
            "height": (
                len(lines) * line_height
                + padding
            ),
        }

    # =====================================================
    # Typography and responsive fitting
    # =====================================================

    title_font_name = "Arial"
    subtitle_font_name = "Arial"
    brand_font_name = "Arial"
    footer_font_name = "Arial"

    title_font_path = resolve_font_path(
        title_font_name,
        bold=False,
    )
    subtitle_font_path = resolve_font_path(
        subtitle_font_name,
        bold=False,
    )
    brand_font_path = resolve_font_path(
        brand_font_name,
        bold=True,
    )
    footer_font_path = resolve_font_path(
        footer_font_name,
        bold=False,
    )

    title_font_size = 70.0
    title_min_font_size = 42.0

    subtitle_font_size = 26.0
    subtitle_min_font_size = 17.0

    title_line_spacing = 0.95
    subtitle_line_spacing = 1.0

    subtitle_bar_padding_x = 34.0
    subtitle_bar_padding_y = 12.0

    def calculate_main_layout(
        current_title_size: float,
        current_subtitle_size: float,
        current_gap: float,
    ) -> dict:
        title_layout = layout_text(
            title,
            font_path=title_font_path,
            font_size=current_title_size,
            max_width=content_width,
            line_spacing=title_line_spacing,
            padding=4,
        )

        subtitle_text_width = max(
            1.0,
            content_width
            - 2 * subtitle_bar_padding_x,
        )

        subtitle_layout = layout_text(
            subtitle,
            font_path=subtitle_font_path,
            font_size=current_subtitle_size,
            max_width=subtitle_text_width,
            line_spacing=subtitle_line_spacing,
            padding=3,
        )

        subtitle_bar_height = (
            subtitle_layout["height"]
            + 2 * subtitle_bar_padding_y
        )

        total_height = (
            title_layout["height"]
            + current_gap
            + subtitle_bar_height
        )

        group_y = (
            main_top
            + max(
                0.0,
                available_main_height - total_height,
            )
            / 2
        )

        return {
            "title": title_layout,
            "subtitle": subtitle_layout,
            "subtitle_bar_height": subtitle_bar_height,
            "total_height": total_height,
            "group_y": group_y,
        }

    layouts = calculate_main_layout(
        title_font_size,
        subtitle_font_size,
        title_to_subtitle_gap,
    )

    for _ in range(220):
        if layouts["total_height"] <= available_main_height:
            break

        changed = False

        if title_to_subtitle_gap > minimum_gap:
            title_to_subtitle_gap = max(
                minimum_gap,
                title_to_subtitle_gap - 1,
            )
            changed = True

        elif subtitle_font_size > subtitle_min_font_size:
            subtitle_font_size = max(
                subtitle_min_font_size,
                subtitle_font_size - 0.5,
            )
            changed = True

        elif title_font_size > title_min_font_size:
            title_font_size = max(
                title_min_font_size,
                title_font_size - 0.5,
            )
            changed = True

        if not changed:
            break

        layouts = calculate_main_layout(
            title_font_size,
            subtitle_font_size,
            title_to_subtitle_gap,
        )

    if layouts["total_height"] > available_main_height:
        raise ValueError(
            "The proposal title and subtitle are too long "
            "to fit without overlap"
        )

    title_y = layouts["group_y"]
    subtitle_bar_y = (
        title_y
        + layouts["title"]["height"]
        + title_to_subtitle_gap
    )
    subtitle_y = (
        subtitle_bar_y
        + subtitle_bar_padding_y
    )

    # =====================================================
    # Build XML
    # =====================================================

    presentation = ET.Element(
        "presentation",
        {
            "schemaVersion": "1.0",
            "width": str(canvas_width),
            "height": str(canvas_height),
            "slideWidthInches": "13.333",
            "slideHeightInches": "7.5",
        },
    )

    slide = ET.SubElement(
        presentation,
        "slide",
        {"background": background_color},
    )

    def add(tag: str, **attrs):
        return ET.SubElement(
            slide,
            tag,
            {
                key: str(value)
                for key, value in attrs.items()
                if value is not None
            },
        )

    def textbox(text: str, **attrs):
        node = add("textbox", **attrs)
        node.text = text
        return node

    svg_node = add(
        "svg",
        x=decoration_x,
        y=decoration_y,
        width=decoration_width,
        height=decoration_height,
        pixel_width=900,
        pixel_height=900,
    )
    svg_node.text = """
<svg xmlns="http://www.w3.org/2000/svg" width="900" height="900" viewBox="0 0 900 900" style="background: transparent;">
  <defs>
    <linearGradient id="outer" x1="0" y1="0" x2="1" y2="1">
      <stop offset="0%" stop-color="#4A2C82"/>
      <stop offset="45%" stop-color="#805AD5"/>
      <stop offset="100%" stop-color="#B794F4"/>
    </linearGradient>
    <linearGradient id="middle" x1="0" y1="0" x2="1" y2="0.9">
      <stop offset="0%" stop-color="#6B46C1"/>
      <stop offset="55%" stop-color="#9F7AEA"/>
      <stop offset="100%" stop-color="#D6BCFA"/>
    </linearGradient>
    <linearGradient id="inner" x1="0" y1="0" x2="1" y2="1">
      <stop offset="0%" stop-color="#805AD5"/>
      <stop offset="50%" stop-color="#B794F4"/>
      <stop offset="100%" stop-color="#E9D8FD"/>
    </linearGradient>
    <filter id="softGlow" x="-30%" y="-30%" width="160%" height="160%">
      <feGaussianBlur stdDeviation="10" result="blur"/>
      <feMerge>
        <feMergeNode in="blur"/>
        <feMergeNode in="SourceGraphic"/>
      </feMerge>
    </filter>
  </defs>

  <g opacity="0.18">
    <rect x="360" y="58" width="250" height="190" fill="#B794F4"/>
    <rect x="360" y="384" width="250" height="190" fill="#9F7AEA"/>
    <rect x="108" y="575" width="205" height="300" fill="#8B5CF6"/>
  </g>

  <path
    d="M0,210 C185,35 420,28 620,58 C650,332 540,602 305,760 C170,850 63,880 0,882 Z"
    fill="url(#outer)"
    opacity="0.96"/>

  <path
    d="M0,410 C180,210 382,194 590,250 C548,470 438,635 258,760 C132,846 52,870 0,876 Z"
    fill="url(#middle)"
    opacity="0.98"/>

  <path
    d="M0,640 C154,410 338,350 535,384 C470,580 363,704 211,798 C112,858 44,874 0,878 Z"
    fill="url(#inner)"
    opacity="0.98"
    filter="url(#softGlow)"/>

  <path
    d="M150,900 C145,750 200,636 312,598 C390,571 457,590 510,640 C412,734 302,824 150,900 Z"
    fill="#8B5CF6"
    opacity="0.82"/>
</svg>
""".strip()

    textbox(
        "✣",
        x=content_x,
        y=brand_y,
        width=brand_icon_width,
        height=58,
        font=brand_font_name,
        font_path=str(brand_font_path),
        size=42,
        min_font_size=42,
        bold="true",
        color=brand_color,
        align="center",
        vertical_align="middle",
        single_line="true",
        max_lines=1,
        overflow="clip",
        reflow_on_resize="true",
        line_spacing=1.0,
        safety_ratio=1.0,
        margin_left=0,
        margin_right=0,
        margin_top=0,
        margin_bottom=0,
    )

    textbox(
        brand_name,
        x=content_x + brand_icon_width + brand_gap,
        y=brand_y + 5,
        width=260,
        height=50,
        font=brand_font_name,
        font_path=str(brand_font_path),
        size=30,
        min_font_size=22,
        bold="true",
        color=brand_color,
        align="left",
        vertical_align="middle",
        single_line="false",
        max_lines=2,
        overflow="shrink",
        reflow_on_resize="true",
        line_spacing=1.0,
        safety_ratio=1.0,
        margin_left=0,
        margin_right=0,
        margin_top=0,
        margin_bottom=0,
    )

    textbox(
        layouts["title"]["source_text"],
        x=content_x,
        y=title_y,
        width=content_width,
        height=layouts["title"]["height"],
        font=title_font_name,
        font_path=str(title_font_path),
        size=layouts["title"]["font_size"],
        min_font_size=layouts["title"]["font_size"],
        color=title_color,
        align="left",
        vertical_align="top",
        single_line="false",
        max_lines=len(layouts["title"]["lines"]),
        overflow="clip",
        reflow_on_resize="true",
        line_spacing=title_line_spacing,
        safety_ratio=1.0,
        margin_left=0,
        margin_right=0,
        margin_top=0,
        margin_bottom=0,
    )

    add(
        "shape",
        type="rect",
        x=content_x,
        y=subtitle_bar_y,
        width=content_width,
        height=layouts["subtitle_bar_height"],
        fill=subtitle_bar_color,
        line_color="none",
    )

    textbox(
        layouts["subtitle"]["source_text"],
        x=content_x + subtitle_bar_padding_x,
        y=subtitle_y,
        width=content_width - 2 * subtitle_bar_padding_x,
        height=layouts["subtitle"]["height"],
        font=subtitle_font_name,
        font_path=str(subtitle_font_path),
        size=layouts["subtitle"]["font_size"],
        min_font_size=layouts["subtitle"]["font_size"],
        color=subtitle_color,
        align="center",
        vertical_align="top",
        single_line="false",
        max_lines=len(layouts["subtitle"]["lines"]),
        overflow="clip",
        reflow_on_resize="true",
        line_spacing=subtitle_line_spacing,
        safety_ratio=1.0,
        margin_left=0,
        margin_right=0,
        margin_top=0,
        margin_bottom=0,
    )

    textbox(
        f"Prepared for {prepared_for}",
        x=footer_left_x,
        y=footer_y,
        width=footer_left_width,
        height=42,
        font=footer_font_name,
        font_path=str(footer_font_path),
        size=22,
        min_font_size=16,
        color=footer_color,
        align="left",
        vertical_align="middle",
        single_line="false",
        max_lines=2,
        overflow="shrink",
        reflow_on_resize="true",
        line_spacing=1.0,
        safety_ratio=1.0,
        margin_left=0,
        margin_right=0,
        margin_top=0,
        margin_bottom=0,
    )

    textbox(
        date_text,
        x=footer_right_x,
        y=footer_y,
        width=footer_right_width,
        height=42,
        font=footer_font_name,
        font_path=str(footer_font_path),
        size=22,
        min_font_size=16,
        color=footer_color,
        align="right",
        vertical_align="middle",
        single_line="false",
        max_lines=2,
        overflow="shrink",
        reflow_on_resize="true",
        line_spacing=1.0,
        safety_ratio=1.0,
        margin_left=0,
        margin_right=0,
        margin_top=0,
        margin_bottom=0,
    )

    raw_xml = ET.tostring(
        presentation,
        encoding="utf-8",
    )

    xml_content = minidom.parseString(
        raw_xml
    ).toprettyxml(
        indent="  ",
        encoding="UTF-8",
    ).decode("utf-8")

    return xml_content

if __name__ == "__main__":
    xml_content = generate_it_sales_proposal_cover_inline_vector_xml()

    Path("output.xml").write_text(
        xml_content,
        encoding="utf-8",
    )

