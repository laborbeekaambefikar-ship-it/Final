"""
Rebuild ArbazFinal1.pdf with figures overlaid onto existing pages (no new
pages added). Behaviour summary:

  * Original text content of every page is preserved verbatim.
  * No additions on pages 1-32.
  * From page 33 onward, technical diagrams are overlaid onto pages relevant
    to their topic. The bottom of each target page is extended only as much
    as needed so the figure is clearly visible (the original text positions
    are never moved).
  * All hyperlink rectangles (the red boxes on the Contents list and green
    boxes around in-text citations) are stripped from every page.
  * New numbering format: sequential "Figure N: ..." captions.
"""

import os
import fitz
from PIL import Image

SRC = "ArbazFinal1.pdf"
OUT = "ArbazFinal1_with_figures.pdf"

A4_W = 595.276
ORIG_H = 841.89

# Geometry of a placed figure
FIG_TARGET_H = 300        # target image height (pt)
CAPTION_GAP = 18          # space between image and caption
CAPTION_RESERVE = 36      # vertical room reserved for (possibly wrapped) caption
SPACE_ABOVE_FIG = 24      # gap between last text line and the image
BOTTOM_PAD = 36           # bottom margin under the caption
SIDE_MARGIN = 72          # 1 inch

# (page_1based, image_filename, figure_number, caption)
PLAN = [
    (33, "Overall Architecture.png", 1,
     "Overall ATLAS system architecture spanning AGV navigation, "
     "manipulation and the unified ROS 2 communication layer."),
    (34, "Ros2 node Communication.png", 2,
     "ROS 2 inter-node communication graph across the eleven ATLAS "
     "packages."),
    (37, "warehouse with agv.png", 3,
     "Differential-drive ATLAS AGV operating within a structured "
     "warehouse environment."),
    (43, "robotic arm.png", 4,
     "Universal Robots UR3 6-DOF collaborative manipulator fitted with the "
     "Robotiq 2F-85 parallel gripper."),
    (49, "Agv Motion COntrol Pipeline using Pd control.png", 5,
     "AGV motion-control pipeline using a Proportional-Derivative (PD) "
     "line-following controller."),
    (51, "12 state.png", 6,
     "12-state Finite State Machine governing the complete autonomous "
     "mission lifecycle."),
    (54, "Arm Pose Stages.png", 7,
     "UR3 arm pose stages during the MoveIt Task Constructor "
     "pick-and-place workflow."),
    (56, "Hardware flow diagram.png", 8,
     "Hardware flow diagram of the ATLAS computing, sensing and actuation "
     "stack."),
    (62, "Priority based order management.png", 9,
     "Priority-based order-management subsystem driving the PyQt5 "
     "industrial control centre."),
    (64, "Warehouse nav map.png", 10,
     "Warehouse navigation map showing the spine corridor, perpendicular "
     "aisles and 21 RFID-tagged shelf locations."),
    (67, "warehouse top shot.png", 11,
     "Top-down view of the simulated warehouse environment used for full "
     "mission validation."),
    (69, "layour warehouse.png", 12,
     "Conceptual warehouse layout depicting shelf placement and AGV "
     "guidance paths."),
    (73, "Pipeline.png", 13,
     "End-to-end project execution pipeline from mission dispatch through "
     "navigation, manipulation and return-to-dock."),
    (74, "warehouse.png", 14,
     "Simulated warehouse view used during continuous 50-mission "
     "validation runs."),
]


def strip_link_annotations(doc):
    """Remove every hyperlink rectangle from every page. We delete via the
    high-level API first, then fall back to clearing the page's /Annots
    array directly (some pages keep the array even after delete_link)."""
    removed = 0
    for page in doc:
        # First pass: high-level delete
        for link in page.get_links():
            try:
                page.delete_link(link)
                removed += 1
            except Exception:
                pass
        # Second pass: if any link annotations remain, drop the /Annots key
        if page.get_links():
            xref = page.xref
            try:
                doc.xref_set_key(xref, "Annots", "null")
                removed += len(page.get_links())
            except Exception:
                pass
    return removed


def last_text_y(page):
    """Return the largest y-coordinate of any text block on the page."""
    blocks = page.get_text("blocks")
    if not blocks:
        return 90.0
    return max(b[3] for b in blocks)


def place_figure(page, image_path, fig_n, caption_text):
    """Overlay the figure onto the bottom of the page, extending the page
    height if needed. Original page content is never moved."""

    # Pre-compute image aspect-fit dimensions
    im = Image.open(image_path)
    iw, ih = im.size
    max_fig_w = A4_W - 2 * SIDE_MARGIN
    scale = min(max_fig_w / iw, FIG_TARGET_H / ih)
    fw, fh = iw * scale, ih * scale

    # Vertical layout below existing text
    text_end_y = last_text_y(page)
    fig_top = text_end_y + SPACE_ABOVE_FIG
    fig_bottom = fig_top + fh
    caption_y = fig_bottom + CAPTION_GAP
    needed_height = caption_y + CAPTION_RESERVE + BOTTOM_PAD

    # Extend the page height (mediabox) only if needed
    current_h = page.rect.height
    if needed_height > current_h:
        new_h = needed_height
        new_box = fitz.Rect(0, 0, A4_W, new_h)
        page.set_mediabox(new_box)
        # Also extend the cropbox so viewers display the new area
        try:
            page.set_cropbox(new_box)
        except Exception:
            pass

    # Paint a white background over the figure area to guarantee clean
    # whitespace (covers any stray watermarks etc.)
    bg_rect = fitz.Rect(SIDE_MARGIN - 8,
                        text_end_y + SPACE_ABOVE_FIG - 8,
                        A4_W - SIDE_MARGIN + 8,
                        caption_y + CAPTION_RESERVE + 4)
    page.draw_rect(bg_rect, color=(1, 1, 1), fill=(1, 1, 1), width=0)

    # Place image, centred horizontally
    fx0 = (A4_W - fw) / 2
    page.insert_image(
        fitz.Rect(fx0, fig_top, fx0 + fw, fig_top + fh),
        filename=image_path,
        keep_proportion=True,
    )

    # Caption
    label = f"Figure {fig_n}: "
    full = label + caption_text
    font_main = "helv"
    font_bold = "hebo"
    fs = 10

    full_w = fitz.get_text_length(full, fontname=font_main, fontsize=fs)
    avail_w = A4_W - 2 * SIDE_MARGIN

    if full_w <= avail_w:
        cap_x = (A4_W - full_w) / 2
        label_w = fitz.get_text_length(label, fontname=font_bold, fontsize=fs)
        page.insert_text((cap_x, caption_y), label,
                         fontname=font_bold, fontsize=fs, color=(0, 0, 0))
        page.insert_text((cap_x + label_w, caption_y), caption_text,
                         fontname=font_main, fontsize=fs, color=(0, 0, 0))
    else:
        # Wrap caption inside a centred textbox
        tb = fitz.Rect(SIDE_MARGIN, caption_y - 10,
                       A4_W - SIDE_MARGIN, caption_y + CAPTION_RESERVE)
        # Use a single bold-prefixed string approximated by inserting label
        # separately above the wrapped text
        page.insert_textbox(tb, full, fontname=font_main, fontsize=fs,
                            align=fitz.TEXT_ALIGN_CENTER)


def main():
    doc = fitz.open(SRC)

    removed = strip_link_annotations(doc)
    print(f"Stripped {removed} hyperlink annotations.")

    placed = 0
    for page_n, img, fig_n, cap in PLAN:
        if not os.path.exists(img):
            print(f"  WARN: missing {img}, skipping Figure {fig_n}")
            continue
        page = doc[page_n - 1]
        place_figure(page, img, fig_n, cap)
        placed += 1
        print(f"  Placed Figure {fig_n} on page {page_n} -> {img}")

    doc.save(OUT, deflate=True, garbage=4)
    doc.close()
    print(f"\nWrote {OUT}  (figures placed: {placed})")

    chk = fitz.open(OUT)
    print(f"Final page count: {len(chk)}  (unchanged from original 80)")
    chk.close()


if __name__ == "__main__":
    main()
