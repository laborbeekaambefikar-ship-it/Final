"""
Build a new PDF from ArbazFinal1.pdf with all repo images placed in suitable
positions. Original pages are copied verbatim (no formatting changes).
Figure pages are inserted right after the original page that introduces the
relevant topic. The cover page gets the logo overlaid in the lower portion.
"""
import fitz  # PyMuPDF
from PIL import Image
import os

SRC = "ArbazFinal1.pdf"
OUT = "ArbazFinal1_with_figures.pdf"

# A4 page dimensions in points (PyMuPDF default)
A4_W, A4_H = 595.276, 841.89

# Margin used by the original document (approx)
MARGIN_X = 72       # ~1 inch left/right
TOP_Y = 90          # top margin
BOTTOM_Y = 90       # bottom margin

# Image placement plan:
#   key   = original 1-based page number after which the figure page is inserted
#   value = list of (image_filename, figure_label, caption) tuples
PLAN = {
    15: [("warehouse top shot.png",
          "Figure 1.1",
          "A modern warehouse environment in which the ATLAS Smart AGV system is intended to operate.")],
    19: [("layour warehouse.png",
          "Figure 1.2",
          "Conceptual warehouse layout showing aisles, shelves and the spine path used by the AGV.")],
    20: [("Pipeline.png",
          "Figure 1.3",
          "End-to-end project development pipeline from mission dispatch to return-to-dock.")],
    21: [("warehouse with agv.png",
          "Figure 1.4",
          "Differential-drive AGV operating within a structured warehouse environment.")],
    29: [("Warehouse nav map.png",
          "Figure 2.1",
          "Warehouse navigation map: spine path, aisles and 21 RFID-tagged shelf locations.")],
    33: [("Overall Architecture.png",
          "Figure 3.1",
          "Overall system architecture of the ATLAS Smart Warehouse AGV with integrated UR3 arm.")],
    34: [("Ros2 node Communication.png",
          "Figure 3.2",
          "ROS 2 node communication graph across the eleven ATLAS packages.")],
    37: [("warehouse.png",
          "Figure 3.3",
          "Simulated warehouse environment used for full mission validation in Gazebo."),
         ("warehouse with agv.png",
          "Figure 3.4",
          "Real-world ATLAS AGV chassis derived from the URDF model.")],
    43: [("robotic arm.png",
          "Figure 3.5",
          "Universal Robots UR3 6-DOF collaborative manipulator with Robotiq 2F-85 gripper.")],
    49: [("Agv Motion COntrol Pipeline using Pd control.png",
          "Figure 3.6",
          "AGV motion-control pipeline using a Proportional-Derivative (PD) line-following controller.")],
    51: [("12 state.png",
          "Figure 3.7",
          "12-state Finite State Machine governing the complete autonomous mission lifecycle.")],
    54: [("Arm Pose Stages.png",
          "Figure 3.8",
          "UR3 arm pose stages during the MoveIt Task Constructor pick-and-place workflow.")],
    56: [("Hardware flow diagram.png",
          "Figure 3.9",
          "Hardware flow diagram of the ATLAS AGV computing, sensing and actuation stack.")],
    62: [("Priority based order management.png",
          "Figure 3.10",
          "Priority-based order-management subsystem driving the PyQt5 industrial control centre.")],
}


def fit_rect(img_w, img_h, max_w, max_h):
    """Return (w, h) preserving aspect ratio inside (max_w, max_h)."""
    scale = min(max_w / img_w, max_h / img_h)
    return img_w * scale, img_h * scale


def add_figure_page(out_doc, image_path, label, caption):
    """Append a new A4 page with a centred figure and caption below."""
    page = out_doc.new_page(width=A4_W, height=A4_H)

    # Page header (matches original document running header style)
    header = "Smart AGV Warehouse System with Integrated Robotic Arm"
    page.insert_text(
        (MARGIN_X, 50),
        header,
        fontname="helv",
        fontsize=10,
        color=(0, 0, 0),
    )
    # thin rule under header
    page.draw_line((MARGIN_X, 58), (A4_W - MARGIN_X, 58),
                   color=(0, 0, 0), width=0.5)

    # Compute image rect
    im = Image.open(image_path)
    iw, ih = im.size

    avail_w = A4_W - 2 * MARGIN_X
    # Reserve room for caption (about 60 pts) + header (about 70 pts)
    avail_h = A4_H - TOP_Y - BOTTOM_Y - 80

    fw, fh = fit_rect(iw, ih, avail_w, avail_h)
    x0 = (A4_W - fw) / 2
    y0 = TOP_Y + 10
    rect = fitz.Rect(x0, y0, x0 + fw, y0 + fh)
    page.insert_image(rect, filename=image_path, keep_proportion=True)

    # Caption below the image
    cap_y = y0 + fh + 18
    label_text = f"{label}: "
    # Measure label width to position caption
    label_w = fitz.get_text_length(label_text, fontname="hebo", fontsize=10)
    caption_full = label_text + caption

    # Center the caption block
    full_w = fitz.get_text_length(caption_full, fontname="helv", fontsize=10)
    if full_w <= avail_w:
        cap_x = (A4_W - full_w) / 2
        page.insert_text((cap_x, cap_y), label_text,
                         fontname="hebo", fontsize=10, color=(0, 0, 0))
        page.insert_text((cap_x + label_w, cap_y), caption,
                         fontname="helv", fontsize=10, color=(0, 0, 0))
    else:
        # Wrap caption into a textbox under the image
        tb_rect = fitz.Rect(MARGIN_X, cap_y - 10,
                            A4_W - MARGIN_X, cap_y + 60)
        page.insert_textbox(
            tb_rect,
            caption_full,
            fontname="helv",
            fontsize=10,
            align=fitz.TEXT_ALIGN_CENTER,
        )


def overlay_logo_on_cover(page, logo_path):
    """Place the logo on the cover page in the lower-centre area
    without disturbing the existing text."""
    im = Image.open(logo_path)
    iw, ih = im.size
    # Target width ~ 130 pt, keep aspect
    target_w = 130.0
    scale = target_w / iw
    target_h = ih * scale
    # Position: horizontally centred, vertical between author block and
    # institution name (around y ~ 600)
    x0 = (A4_W - target_w) / 2
    y0 = 600
    rect = fitz.Rect(x0, y0, x0 + target_w, y0 + target_h)
    page.insert_image(rect, filename=logo_path, keep_proportion=True,
                      overlay=True)


def main():
    src = fitz.open(SRC)
    out = fitz.open()

    for i in range(len(src)):
        # Copy original page byte-for-byte
        out.insert_pdf(src, from_page=i, to_page=i)
        page_num = i + 1

        # Logo overlay on cover
        if page_num == 1:
            overlay_logo_on_cover(out[-1], "image-removebg-preview.png")

        # Insert figure pages after this original page if planned
        if page_num in PLAN:
            for img, label, cap in PLAN[page_num]:
                if not os.path.exists(img):
                    print(f"WARN: missing image {img}")
                    continue
                add_figure_page(out, img, label, cap)
                print(f"Inserted {label} ({img}) after original page {page_num}")

    out.save(OUT, deflate=True, garbage=4)
    out.close()
    src.close()
    print(f"\nWrote {OUT}")
    new_doc = fitz.open(OUT)
    print(f"Final page count: {len(new_doc)}")
    new_doc.close()


if __name__ == "__main__":
    main()
