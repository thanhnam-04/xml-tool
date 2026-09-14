from __future__ import annotations

import os
import xml.etree.ElementTree as ET
from pathlib import Path
from xml.dom import minidom

from PIL import ImageFont


def generate_agile_values_principles_slide_xml(
    title: str = "AGILE VALUES\nAND PRINCIPLES AGILE VALUES\nAND PRINCIPLES",
    items: list[dict[str, str]] | None = None,
    background_color: str = "#F4F4F6",
    title_color: str = "#0B0B0B",
    text_color: str = "#171717",
) -> str:
    import math

    if items is None:
        items = [
          
            {"title": "Individuals and Interactions", "description": "Value the people and communication over processes and tools.", "color": "#FFC400"},
            {"title": "Working Software", "description": "Deliver functional increments over comprehensive documentation.", "color": "#9ECE45"},
        ]

    if not 2 <= len(items) <= 6:
        raise ValueError("Layout này hỗ trợ từ 2 đến 6 items.")

    presentation = ET.Element(
        "presentation",
        {
            "schemaVersion": "1.0",
            "width": "1600",
            "height": "900",
            "slideWidthInches": "13.333",
            "slideHeightInches": "7.5",
        },
    )

    slide = ET.SubElement(presentation, "slide", {"background": background_color})

    def add(tag: str, **attrs):
        return ET.SubElement(slide, tag, {k: str(v) for k, v in attrs.items() if v is not None})

    def textbox(text: str, **attrs):
        node = add("textbox", **attrs)
        node.text = str(text)
        return node

    def shorten_text(
        value: str,
        *,
        max_words: int,
        max_chars: int,
        preserve_line_breaks: bool = False,
    ) -> str:
        cleaned_lines = [
            " ".join(line.split())
            for line in str(value).splitlines()
            if line.strip()
        ]
        normalized = " ".join(cleaned_lines)
        words = normalized.split()

        if len(words) <= max_words and len(normalized) <= max_chars:
            if preserve_line_breaks and cleaned_lines:
                return "\n".join(cleaned_lines)
            return normalized

        selected: list[str] = []
        for word in words:
            candidate = " ".join([*selected, word])
            if len(selected) >= max_words or len(candidate) + 1 > max_chars:
                break
            selected.append(word)

        if not selected:
            return normalized[: max(1, max_chars - 1)].rstrip() + "…"

        return " ".join(selected).rstrip(" ,.;:-") + "…"

    add("shape", type="rectangle", x=0, y=0, width=1600, height=900, fill=background_color, line_color="none")

    # =========================================================
    # very light geometric background
    # =========================================================

    bg = add("svg", x=310, y=35, width=1000, height=830, pixel_width=1000, pixel_height=830)
    bg.text = """
<svg xmlns="http://www.w3.org/2000/svg" width="1000" height="830" viewBox="0 0 1000 830">
  <g opacity="0.55">
    <polygon points="500,415 705,210 500,40 295,210" fill="#ECEDEF"/>
    <polygon points="500,415 910,415 705,210" fill="#F0F1F3"/>
    <polygon points="500,415 705,620 910,415" fill="#F3F4F6"/>
    <polygon points="500,415 500,790 705,620" fill="#EFF0F2"/>
    <polygon points="500,415 295,620 500,790" fill="#F3F4F6"/>
    <polygon points="500,415 90,415 295,620" fill="#EFF0F2"/>
    <polygon points="500,415 295,210 90,415" fill="#F3F4F6"/>
  </g>
</svg>
"""

    # =========================================================
    # wheel
    # =========================================================

    wheel_x = 360
    wheel_y = 85
    wheel_size = 900

    cx = 450.0
    cy = 450.0
    outer_r = 350.0
    inner_r = 175.0
    arrow_r = 385.0

    def polar(r, angle_deg):
        a = math.radians(angle_deg)
        return cx + r * math.cos(a), cy + r * math.sin(a)

    def segment_path(start_angle, end_angle):
        arrow_angle = (start_angle + end_angle) / 2.0
        arrow_half = 7.0

        p1 = polar(outer_r, start_angle)
        p2 = polar(outer_r, arrow_angle - arrow_half)
        tip = polar(arrow_r, arrow_angle)
        p3 = polar(outer_r, arrow_angle + arrow_half)
        p6 = polar(outer_r, end_angle)
        i1 = polar(inner_r, end_angle)
        i2 = polar(inner_r, start_angle)

        return (
            f"M {p1[0]:.2f},{p1[1]:.2f} "
            f"A {outer_r},{outer_r} 0 0 1 {p2[0]:.2f},{p2[1]:.2f} "
            f"L {tip[0]:.2f},{tip[1]:.2f} "
            f"L {p3[0]:.2f},{p3[1]:.2f} "
            f"A {outer_r},{outer_r} 0 0 1 {p6[0]:.2f},{p6[1]:.2f} "
            f"L {i1[0]:.2f},{i1[1]:.2f} "
            f"A {inner_r},{inner_r} 0 0 0 {i2[0]:.2f},{i2[1]:.2f} Z"
        )

    # Keep a 90-degree opening at the bottom. The remaining 270 degrees are
    # divided evenly for any supported number of items.
    item_count = len(items)
    segment_angle = 270.0 / item_count
    arc_centers = [
        135.0 + (index + 0.5) * segment_angle
        for index in range(item_count)
    ]
    first_index = min(
        range(item_count),
        key=lambda index: (
            abs(arc_centers[index] - 270.0),
            arc_centers[index],
        ),
    )
    segment_centers = (
        arc_centers[first_index:]
        + arc_centers[:first_index]
    )
    segment_half_angle = segment_angle / 2.0

    def icon_collaboration(x, y):
        return f"""
<g transform="translate({x},{y})" fill="none" stroke="#111111" stroke-width="3.6" stroke-linecap="round" stroke-linejoin="round">
  <circle cx="-16" cy="10" r="5.5"/><circle cx="18" cy="-14" r="5.5"/>
  <path d="M-25 29 V21 C-25 15 -21 11 -16 11"/>
  <path d="M9 4 V-4 C9 -10 13 -14 18 -14"/>
  <path d="M-2 -8 C5 -21 18 -24 29 -19"/>
  <polyline points="19,-28 31,-20 24,-9"/>
  <path d="M4 26 C-4 37 -17 40 -28 34"/>
  <polyline points="-19,42 -31,35 -25,23"/>
</g>"""

    def icon_clipboard(x, y):
        return f"""
<g transform="translate({x},{y})" fill="none" stroke="#111111" stroke-width="3.6" stroke-linecap="round" stroke-linejoin="round">
  <rect x="-19" y="-27" width="38" height="54" rx="4"/>
  <rect x="-8" y="-34" width="16" height="10" rx="3"/>
  <path d="M-8 -3 L-2 3 L8 -9"/>
  <path d="M3 13 C10 4 14 0 17 -2"/>
  <circle cx="5" cy="14" r="2.8"/>
</g>"""

    def icon_chart(x, y):
        return f"""
<g transform="translate({x},{y})" fill="none" stroke="#111111" stroke-width="3.6" stroke-linecap="round" stroke-linejoin="round">
  <line x1="-26" y1="27" x2="28" y2="27"/>
  <rect x="-21" y="6" width="7" height="21"/>
  <rect x="-7" y="-5" width="7" height="32"/>
  <rect x="7" y="1" width="7" height="26"/>
  <rect x="21" y="-14" width="7" height="41"/>
  <polyline points="-18,-10 -4,-1 10,-14 24,-25"/>
  <circle cx="-18" cy="-10" r="2.8"/><circle cx="-4" cy="-1" r="2.8"/><circle cx="10" cy="-14" r="2.8"/><circle cx="24" cy="-25" r="2.8"/>
</g>"""

    def icon_calendar(x, y):
        return f"""
<g transform="translate({x},{y})" fill="none" stroke="#111111" stroke-width="3.6" stroke-linecap="round" stroke-linejoin="round">
  <rect x="-22" y="-20" width="44" height="40" rx="3"/>
  <line x1="-22" y1="-8" x2="22" y2="-8"/>
  <line x1="-11" y1="-28" x2="-11" y2="-15"/>
  <line x1="11" y1="-28" x2="11" y2="-15"/>
  <rect x="-12" y="0" width="6" height="6"/><rect x="0" y="0" width="6" height="6"/>
  <rect x="-12" y="11" width="6" height="6"/><rect x="0" y="11" width="6" height="6"/>
</g>"""

    def icon_people(x, y):
        return f"""
<g transform="translate({x},{y})" fill="none" stroke="#111111" stroke-width="3.6" stroke-linecap="round" stroke-linejoin="round">
  <circle cx="0" cy="-15" r="6"/><circle cx="-18" cy="-8" r="5"/><circle cx="18" cy="-8" r="5"/>
  <path d="M-10 24 V8 C-10 1 -5 -4 0 -4 C6 -4 10 1 10 8 V24"/>
  <path d="M-25 21 V10 C-25 5 -22 1 -18 1"/><path d="M25 21 V10 C25 5 22 1 18 1"/>
</g>"""

    def icon_phone(x, y):
        return f"""
<g transform="translate({x},{y})" fill="none" stroke="#111111" stroke-width="3.6" stroke-linecap="round" stroke-linejoin="round">
  <rect x="-16" y="-28" width="32" height="56" rx="4"/>
  <line x1="-16" y1="18" x2="16" y2="18"/><circle cx="0" cy="23" r="2.5"/>
</g>"""

    wheel = add("svg", x=wheel_x, y=wheel_y, width=wheel_size, height=wheel_size, pixel_width=900, pixel_height=900)

    parts = [
        '<svg xmlns="http://www.w3.org/2000/svg" width="900" height="900" viewBox="0 0 900 900">',
        '<defs><filter id="wheelShadow" x="-30%" y="-30%" width="160%" height="160%"><feDropShadow dx="0" dy="7" stdDeviation="10" flood-color="#B9BCC2" flood-opacity="0.18"/></filter></defs>',
        '<g filter="url(#wheelShadow)">',
    ]

    for i, item in enumerate(items):
        mid = segment_centers[i]
        parts.append(
            f'<path d="{segment_path(mid - segment_half_angle, mid + segment_half_angle)}" '
            f'fill="{item["color"]}"/>'
        )

    parts.append(f'<circle cx="{cx}" cy="{cy}" r="{inner_r}" fill="{background_color}"/>')
    parts.append("</g>")

    icon_r = 260
    for i in range(item_count):
        ix, iy = polar(icon_r, segment_centers[i])
        if i == 0:
            parts.append(icon_collaboration(ix, iy))
        elif i == 1:
            parts.append(icon_clipboard(ix, iy))
        elif i == 2:
            parts.append(icon_chart(ix, iy))
        elif i == 3:
            parts.append(icon_calendar(ix, iy))
        elif i == 4:
            parts.append(icon_people(ix, iy))
        else:
            parts.append(icon_phone(ix, iy))

    parts.append("</svg>")
    wheel.text = "".join(parts)

    # =========================================================
    # center title
    # =========================================================

    textbox(
        shorten_text(
            title,
            max_words=4,
            max_chars=32,
            preserve_line_breaks=True,
        ),
        x=640,
        y=470,
        width=340,
        height=130,
        font="Arial",
        size=24,
        min_font_size=20,
        color=title_color,
        bold="true",
        italic="false",
        align="center",
        vertical_align="middle",
        single_line="false",
        max_lines=2,
        overflow="shrink",
        reflow_on_resize="false",
        line_spacing=1.0,
        margin_left=0,
        margin_right=0,
        margin_top=0,
        margin_bottom=0,
    )

    # =========================================================
    # adaptive label positions for 2-6 items
    # =========================================================

    label_slots = {
        ("center", "top"): {"title_x": 585, "title_y": 18, "w": 450, "desc_y": 55, "desc_h": 62, "bar_x": 750, "bar_y": 122, "bar_w": 120, "bar_h": 6},
        ("left", "top"): {"title_x": 70, "title_y": 66, "w": 500, "desc_y": 105, "bar_x": 590, "bar_y": 62},
        ("right", "top"): {"title_x": 1052, "title_y": 66, "w": 450, "desc_y": 105, "bar_x": 1025, "bar_y": 62},
        ("left", "upper_middle"): {"title_x": 50, "title_y": 204, "w": 350, "desc_y": 249, "bar_x": 420, "bar_y": 210},
        ("right", "upper_middle"): {"title_x": 1235, "title_y": 204, "w": 300, "desc_y": 249, "bar_x": 1210, "bar_y": 210},
        ("left", "middle"): {"title_x": 50, "title_y": 335, "w": 350, "desc_y": 380, "bar_x": 420, "bar_y": 341},
        ("right", "middle"): {"title_x": 1235, "title_y": 335, "w": 300, "desc_y": 380, "bar_x": 1210, "bar_y": 341},
        ("left", "centerline"): {"title_x": 50, "title_y": 486, "w": 350, "desc_y": 531, "bar_x": 420, "bar_y": 492},
        ("right", "centerline"): {"title_x": 1235, "title_y": 486, "w": 300, "desc_y": 531, "bar_x": 1210, "bar_y": 492},
        ("left", "bottom"): {"title_x": 50, "title_y": 658, "w": 340, "desc_y": 698, "bar_x": 410, "bar_y": 653},
        ("right", "bottom"): {"title_x": 1280, "title_y": 658, "w": 255, "desc_y": 698, "bar_x": 1254, "bar_y": 653},
    }

    def slot_for_angle(angle: float) -> tuple[str, str]:
        radians = math.radians(angle)
        horizontal = math.cos(radians)
        vertical = math.sin(radians)

        # Odd item counts have one arrow pointing straight upward. Its text
        # belongs directly above it rather than being forced left or right.
        if vertical < -0.9 and abs(horizontal) < 0.2:
            return "center", "top"

        side = "left" if horizontal < 0 else "right"
        if item_count == 3 and abs(vertical) < 0.15:
            row = "centerline"
        elif item_count == 5 and vertical < -0.45:
            row = "upper_middle"
        elif vertical < -0.65:
            row = "top"
        elif vertical > 0.15:
            row = "bottom"
        else:
            row = "middle"
        return side, row

    item_slots = [
        slot_for_angle(angle)
        for angle in segment_centers
    ]
    specs = [
        {
            **label_slots[slot],
            "side": slot[0],
        }
        for slot in item_slots
    ]

    for i, item in enumerate(items):
        s = specs[i]
        if s["side"] == "left":
            align = "right"
        elif s["side"] == "right":
            align = "left"
        else:
            align = "center"
        title_char_limit = max(22, min(36, round(s["w"] / 12)))
        description_char_limit = max(65, min(105, round(s["w"] * 0.28)))
        item_title = shorten_text(
            item["title"],
            max_words=5,
            max_chars=title_char_limit,
        )
        item_description = shorten_text(
            item["description"],
            max_words=14,
            max_chars=description_char_limit,
        )

        add(
            "shape",
            type="rectangle",
            x=s["bar_x"],
            y=s["bar_y"],
            width=s.get("bar_w", 6),
            height=s.get("bar_h", 86),
            fill=item["color"],
            line_color="none",
        )

        textbox(
            item_title,
            x=s["title_x"],
            y=s["title_y"],
            width=s["w"],
            height=34,
            font="Arial",
            size=16,
            min_font_size=14,
            color=text_color,
            bold="true",
            italic="false",
            align=align,
            vertical_align="middle",
            single_line="true",
            max_lines=1,
            overflow="shrink",
            reflow_on_resize="false",
            margin_left=0,
            margin_right=0,
            margin_top=0,
            margin_bottom=0,
        )

        textbox(
            item_description,
            x=s["title_x"],
            y=s["desc_y"],
            width=s["w"],
            height=s.get("desc_h", 70),
            font="Arial",
            size=13.5,
            min_font_size=11,
            color=text_color,
            bold="false",
            italic="false",
            align=align,
            vertical_align="top",
            single_line="false",
            max_lines=3,
            overflow="shrink",
            reflow_on_resize="false",
            line_spacing=1.02,
            margin_left=0,
            margin_right=0,
            margin_top=0,
            margin_bottom=0,
        )

    raw_xml = ET.tostring(presentation, encoding="utf-8")
    xml_content = minidom.parseString(raw_xml).toprettyxml(indent="  ", encoding="UTF-8").decode("utf-8")
    return xml_content

def generate_agile_customer_journey_slide_xml(
    title: str = "AGILE AND THE\nCUSTOMER\nJOURNEY AGILE AND THE\nCUSTOMER\nJOURNEY",
    items: list[dict[str, str]] = [
                {
                    "title": "Discovery",
                    "description": "Teams work with customers to understand their specific problems.",
                    "color": "#FFC400",
                },
                {
                    "title": "Iteration",
                    "description": "New features are developed in small, usable increments.",
                    "color": "#A6D24A",
                },
                {
                    "title": "Feedback",
                    "description": "Customer feedback is continuously integrated into the development.",
                    "color": "#12CDA5",
                },
                {
                    "title": "Delivery",
                    "description": "The team releases usable products frequently.",
                    "color": "#18B7CF",
                },
            ],
    background_color: str = "#F3F4F6",
    title_color: str = "#000000",
    text_color: str = "#111111",
    connector_color: str = "#D6DAE0",
) -> str:
    import math

    def shorten_text(value: str, max_words: int, max_chars: int) -> str:
        """Keep unpredictable LLM copy inside the designed text area."""
        clean = " ".join(str(value or "").split())
        words = clean.split()
        if len(words) > max_words:
            clean = " ".join(words[:max_words])
        if len(clean) > max_chars:
            clean = clean[:max_chars].rsplit(" ", 1)[0]
        if clean != " ".join(str(value or "").split()):
            clean = clean.rstrip(".,;:!?") + "…"
        return clean

   

    raw_items = items if isinstance(items, (list, tuple)) else [items]
    if not 2 <= len(raw_items) <= 6:
        raise ValueError("Layout này hỗ trợ từ 2 đến 6 items.")

    normalized_items = []

    default_colors = [
        "#FFC400",
        "#A6D24A",
        "#12CDA5",
        "#18B7CF",
        "#4B91E2",
        "#8273D9",
    ]
    default_icons = [
        "discovery",
        "iteration",
        "feedback",
        "delivery",
        "collaboration",
        "improvement",
    ]

    item_count = len(raw_items)
    description_limits = {
        2: (32, 190),
        3: (28, 170),
        4: (24, 145),
        5: (19, 120),
        6: (15, 96),
    }
    item_title_limits = {
        2: (11, 82),
        3: (10, 76),
        4: (9, 70),
        5: (8, 62),
        6: (7, 54),
    }
    description_max_words, description_max_chars = description_limits[item_count]
    item_title_max_words, item_title_max_chars = item_title_limits[item_count]

    for index, raw_item in enumerate(raw_items):
        item = raw_item if isinstance(raw_item, dict) else {}
        color = str(item.get("color") or default_colors[index]).strip()
        if not color.startswith("#"):
            color = "#" + color

        normalized_items.append(
            {
                "title": shorten_text(
                    item.get("title") or f"Step {index + 1}",
                    max_words=item_title_max_words,
                    max_chars=item_title_max_chars,
                ),
                "description": shorten_text(
                    item.get("description") or "",
                    max_words=description_max_words,
                    max_chars=description_max_chars,
                ),
                "color": color,
                "icon": str(item.get("icon") or default_icons[index]).strip().lower(),
            }
        )

    # Scale the title hub together with long copy instead of only shrinking
    # the font. The outer journey ring grows too, preserving its thickness.
    center_title = shorten_text(title, max_words=18, max_chars=105)
    center_title_length = len(center_title)
    if center_title_length <= 34:
        inner_r = 180.0
        outer_r = 300.0
        center_title_size = 27
        center_title_min_size = 19
        center_title_max_lines = 3
    elif center_title_length <= 62:
        inner_r = 205.0
        outer_r = 325.0
        center_title_size = 23
        center_title_min_size = 17
        center_title_max_lines = 4
    else:
        inner_r = 225.0
        outer_r = 345.0
        center_title_size = 20
        center_title_min_size = 15
        center_title_max_lines = 5

    arrow_r = outer_r + 26.0

    # =========================================================
    # Canvas
    # =========================================================

    canvas_width = 1600
    canvas_height = 900

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
        {
            "background": background_color,
        },
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
        node.text = str(text)
        return node

    # =========================================================
    # Background
    # =========================================================

    add(
        "shape",
        type="rectangle",
        x=0,
        y=0,
        width=canvas_width,
        height=canvas_height,
        fill=background_color,
        line_color="none",
    )

    # =========================================================
    # Left journey ring
    # =========================================================

    ring_x = 40
    ring_y = 90
    ring_size = 720

    ring_svg = add(
        "svg",
        x=ring_x,
        y=ring_y,
        width=ring_size,
        height=ring_size,
        pixel_width=900,
        pixel_height=900,
    )

    # The SVG maps to 720 x 720 on the slide (scale = 0.8).
    # Its centre therefore lands at slide coordinate (400, 450).
    cx = 450.0
    cy = 450.0

    def polar(radius: float, angle_deg: float):
        rad = math.radians(angle_deg)
        return (
            cx + radius * math.cos(rad),
            cy + radius * math.sin(rad),
        )

    def annular_arrow_path(start_angle: float, end_angle: float):
        """Equal annular segment with a centred outward pointer."""
        centre_angle = (start_angle + end_angle) / 2
        pointer_half_angle = min(5.5, (end_angle - start_angle) * 0.16)
        outer_start = polar(outer_r, start_angle)
        pointer_start = polar(outer_r, centre_angle - pointer_half_angle)
        pointer_tip = polar(arrow_r, centre_angle)
        pointer_end = polar(outer_r, centre_angle + pointer_half_angle)
        outer_end = polar(outer_r, end_angle)
        inner_end = polar(inner_r, end_angle)
        inner_start = polar(inner_r, start_angle)

        return (
            f"M {outer_start[0]:.2f},{outer_start[1]:.2f} "
            f"A {outer_r:.2f},{outer_r:.2f} 0 0 1 {pointer_start[0]:.2f},{pointer_start[1]:.2f} "
            f"L {pointer_tip[0]:.2f},{pointer_tip[1]:.2f} "
            f"L {pointer_end[0]:.2f},{pointer_end[1]:.2f} "
            f"A {outer_r:.2f},{outer_r:.2f} 0 0 1 {outer_end[0]:.2f},{outer_end[1]:.2f} "
            f"L {inner_end[0]:.2f},{inner_end[1]:.2f} "
            f"A {inner_r:.2f},{inner_r:.2f} 0 0 0 {inner_start[0]:.2f},{inner_start[1]:.2f} Z"
        )

    # The right-side 180-degree journey arc is divided into equal-area
    # segments for any supported item count.
    segment_span = 180.0 / item_count
    segment_angles = [
        (-90.0 + index * segment_span, -90.0 + (index + 1) * segment_span)
        for index in range(item_count)
    ]

    svg_parts = [
        '<svg xmlns="http://www.w3.org/2000/svg" width="900" height="900" viewBox="0 0 900 900">',
        '<defs><filter id="hubShadow" x="-30%" y="-30%" width="160%" height="160%">'
        '<feDropShadow dx="0" dy="8" stdDeviation="12" flood-color="#253044" flood-opacity="0.12"/>'
        '</filter></defs>',
    ]

    for idx, item in enumerate(normalized_items):
        start_a, end_a = segment_angles[idx]
        svg_parts.append(
            f'<path d="{annular_arrow_path(start_a, end_a)}" '
            f'fill="{item["color"]}" stroke="{background_color}" stroke-width="4"/>'
        )

    # White hub above a soft shadow.
    svg_parts.append(
        f'<circle cx="{cx}" cy="{cy}" r="{inner_r + 2}" fill="#FFFFFF" filter="url(#hubShadow)"/>'
    )

    svg_parts.append("</svg>")
    ring_svg.text = "".join(svg_parts)

    # =========================================================
    # Center title inside circle
    # =========================================================

    hub_diameter_slide = (inner_r + 2.0) * 2.0 * ring_size / 900.0
    title_box_width = hub_diameter_slide - 10.0
    title_box_height = hub_diameter_slide - 36.0
    hub_center_x = ring_x + cx * ring_size / 900.0
    hub_center_y = ring_y + cy * ring_size / 900.0

    textbox(
        center_title,
        x=hub_center_x - title_box_width / 2.0,
        y=hub_center_y - title_box_height / 2.0,
        width=title_box_width,
        height=title_box_height,
        font="Arial",
        size=center_title_size,
        min_font_size=center_title_min_size,
        color=title_color,
        bold="true",
        italic="false",
        align="center",
        vertical_align="middle",
        single_line="false",
        max_lines=center_title_max_lines,
        overflow="shrink",
        reflow_on_resize="true",
        line_spacing=1.05,
        margin_left=0,
        margin_right=0,
        margin_top=0,
        margin_bottom=0,
    )

    # =========================================================
    # Connector lines
    # =========================================================

    card_x = 850
    card_width = 650
    card_height, card_gap = {
        2: (150, 80),
        3: (140, 70),
        4: (120, 80),
        5: (110, 50),
        6: (96, 34),
    }[item_count]
    cards_total_height = item_count * card_height + (item_count - 1) * card_gap
    cards_top = (canvas_height - cards_total_height) / 2
    card_specs = [
        {
            "x": card_x,
            "y": cards_top + index * (card_height + card_gap),
            "w": card_width,
            "h": card_height,
        }
        for index in range(item_count)
    ]
    card_centres_y = [spec["y"] + spec["h"] / 2 for spec in card_specs]

    # Each connector begins at the matching segment's pointer tip.
    connector_starts = []
    for start_a, end_a in segment_angles:
        centre_angle = (start_a + end_a) / 2
        tip_x_svg, tip_y_svg = polar(arrow_r, centre_angle)
        connector_starts.append(
            (
                ring_x + tip_x_svg * ring_size / 900,
                ring_y + tip_y_svg * ring_size / 900,
            )
        )

    for i, (start_x, start_y) in enumerate(connector_starts):
        add(
            "line",
            x1=round(start_x, 2),
            y1=round(start_y, 2),
            x2=card_x,
            y2=card_centres_y[i],
            color=connector_color,
            width=1.6,
        )

        add(
            "shape",
            type="circle",
            x=card_x - 6,
            y=card_centres_y[i] - 6,
            width=12,
            height=12,
            fill=normalized_items[i]["color"],
            line_color=background_color,
            line_width=2,
        )

    # =========================================================
    # Icon helpers
    # =========================================================

    def make_icon_svg(icon_name: str, color: str) -> str:
        """Return a consistent, modern 64px line icon."""
        aliases = {
            "search": "discovery",
            "research": "discovery",
            "repeat": "iteration",
            "cycle": "iteration",
            "chat": "feedback",
            "message": "feedback",
            "launch": "delivery",
            "release": "delivery",
            "people": "collaboration",
            "team": "collaboration",
            "growth": "improvement",
            "chart": "improvement",
        }
        icon_name = aliases.get(icon_name, icon_name)

        paths = {
            "discovery": """
  <circle cx="25" cy="23" r="8"/>
  <path d="M12 45c1-9 6-14 13-14 5 0 9 2 11 6"/>
  <circle cx="43" cy="41" r="9"/>
  <path d="M50 48l8 8"/>
  <path d="M39 41h8M43 37v8"/>
""",
            "iteration": """
  <path d="M49 22A20 20 0 0 0 15 18"/>
  <path d="M15 18v-9M15 18h9"/>
  <path d="M15 42a20 20 0 0 0 34 4"/>
  <path d="M49 46v9M49 46h-9"/>
  <path d="M25 25h14v14H25z"/>
""",
            "feedback": """
  <path d="M9 15h33a5 5 0 0 1 5 5v18a5 5 0 0 1-5 5H25L14 52v-9H9a5 5 0 0 1-5-5V20a5 5 0 0 1 5-5z"/>
  <path d="M16 28h20M16 35h13"/>
  <path d="M45 48h5l8 7v-20a5 5 0 0 0-5-5h-2"/>
""",
            "delivery": """
  <path d="M9 21l23-12 23 12-23 12z"/>
  <path d="M9 21v25l23 12 23-12V21"/>
  <path d="M32 33v25"/>
  <path d="M21 18l23 12"/>
  <path d="M42 43l5 5 10-12"/>
""",
            "collaboration": """
  <circle cx="32" cy="20" r="7"/>
  <circle cx="13" cy="27" r="5"/>
  <circle cx="51" cy="27" r="5"/>
  <path d="M20 54V42c0-8 5-13 12-13s12 5 12 13v12"/>
  <path d="M4 52v-9c0-6 4-10 9-10 3 0 5 1 7 3"/>
  <path d="M60 52v-9c0-6-4-10-9-10-3 0-5 1-7 3"/>
""",
            "improvement": """
  <path d="M9 54V35h10v19M27 54V25h10v29M45 54V14h10v40"/>
  <path d="M7 54h51"/>
  <path d="M11 27l14-10 10 5L53 7"/>
  <path d="M44 7h9v9"/>
""",
        }
        icon_paths = paths.get(icon_name, paths["improvement"])
        return f"""
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64">
  <g fill="none" stroke="{color}" stroke-width="2.8" stroke-linecap="round" stroke-linejoin="round">
    {icon_paths}
  </g>
</svg>
"""

    # =========================================================
    # Right cards
    # =========================================================

    title_font_size = {2: 19, 3: 18.5, 4: 18, 5: 16.5, 6: 15.5}[item_count]
    description_font_size = {2: 14.5, 3: 14, 4: 14, 5: 12.5, 6: 11.5}[item_count]
    title_box_height = {2: 48, 3: 44, 4: 40, 5: 34, 6: 30}[item_count]
    text_padding_y = max(10.0, (card_height - 88.0) / 2)

    for i, item in enumerate(normalized_items):
        spec = card_specs[i]
        badge_size = min(80.0, spec["h"] - 24.0)
        badge_x = spec["x"] + 28.0
        badge_y = spec["y"] + (spec["h"] - badge_size) / 2
        icon_size = badge_size * 0.56
        text_x = badge_x + badge_size + 28.0
        title_y = spec["y"] + text_padding_y
        description_y = title_y + title_box_height + 2.0
        description_height = max(32.0, spec["y"] + spec["h"] - 11.0 - description_y)

        # Offset background creates a subtle, renderer-independent shadow.
        add(
            "shape",
            type="rounded_rectangle",
            x=spec["x"] + 5,
            y=spec["y"] + 7,
            width=spec["w"],
            height=spec["h"],
            fill="#DDE1E6",
            line_color="none",
        )

        add(
            "shape",
            type="rounded_rectangle",
            x=spec["x"],
            y=spec["y"],
            width=spec["w"],
            height=spec["h"],
            fill="#FFFFFF",
            line_color="#E2E5E9",
            line_width=1,
        )

        # Colour is used as an accent, keeping the card light and readable.
        add(
            "shape",
            type="rounded_rectangle",
            x=spec["x"],
            y=spec["y"],
            width=10,
            height=spec["h"],
            fill=item["color"],
            line_color="none",
        )

        add(
            "shape",
            type="circle",
            x=badge_x,
            y=badge_y,
            width=badge_size,
            height=badge_size,
            fill="#F8FAFC",
            line_color=item["color"],
            line_width=3,
        )

        icon_node = add(
            "svg",
            x=badge_x + (badge_size - icon_size) / 2,
            y=badge_y + (badge_size - icon_size) / 2,
            width=icon_size,
            height=icon_size,
            pixel_width=120,
            pixel_height=120,
        )
        icon_node.text = make_icon_svg(item["icon"], item["color"])

        textbox(
            item["title"],
            x=text_x,
            y=title_y,
            width=spec["x"] + spec["w"] - 28.0 - text_x,
            height=title_box_height,
            font="Arial",
            size=title_font_size,
            min_font_size=13,
            color="#000000",
            bold="true",
            italic="false",
            align="left",
            vertical_align="middle",
            single_line="false",
            max_lines=2,
            overflow="shrink",
            reflow_on_resize="true",
            margin_left=0,
            margin_right=0,
            margin_top=0,
            margin_bottom=0,
        )

        textbox(
            item["description"],
            x=text_x,
            y=description_y,
            width=spec["x"] + spec["w"] - 28.0 - text_x,
            height=description_height,
            font="Arial",
            size=description_font_size,
            min_font_size=9.5,
            color=text_color,
            bold="false",
            italic="false",
            align="left",
            vertical_align="top",
            single_line="false",
            max_lines=3 if item_count <= 4 else 2,
            overflow="shrink",
            reflow_on_resize="true",
            line_spacing=1.02,
            margin_left=0,
            margin_right=0,
            margin_top=0,
            margin_bottom=0,
        )

    raw_xml = ET.tostring(presentation, encoding="utf-8")
    xml_content = minidom.parseString(raw_xml).toprettyxml(indent="  ", encoding="UTF-8").decode("utf-8")
    return xml_content
    
if __name__ == "__main__":
    xml_content = generate_agile_customer_journey_slide_xml()

    Path("output.xml").write_text(
        xml_content,
        encoding="utf-8",
    )
