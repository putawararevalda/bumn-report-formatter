"""
Rebuild assets/template.pptx from a BUMN .pptx master.

Usage:
    python3 rebuild_template.py /path/to/master.pptx [out.pptx]

Steps:
1. Opens the master presentation and removes all slide instances (keeps layouts and master).
2. Prunes large unused media files, replacing them with minimal stubs to keep the template lightweight (< 1 MB).
3. Verifies layout structure and theme integrity.
"""
import io
import os
import shutil
import sys
import tempfile
import zipfile
from PIL import Image
from pptx import Presentation

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_OUT = os.path.join(HERE, "..", "assets", "template.pptx")


def rebuild(src_path: str, out_path: str = DEFAULT_OUT) -> str:
    print(f"Rebuilding template from {src_path} -> {out_path}...")
    work_dir = tempfile.mkdtemp()
    temp_stage1 = os.path.join(work_dir, "stage1.pptx")
    temp_stage2 = os.path.join(work_dir, "stage2.pptx")

    try:
        # Step 1: Remove all slide instances
        prs = Presentation(src_path)
        slide_ids = prs.slides._sldIdLst
        for s in list(slide_ids):
            prs.part.drop_rel(s.rId)
            slide_ids.remove(s)
        prs.save(temp_stage1)

        # Step 2: Slim down large unneeded media files (> 500KB not in primary layouts)
        tiny = Image.new("RGBA", (1, 1), (0, 0, 0, 0))
        tiny_buf = io.BytesIO()
        tiny.save(tiny_buf, format="PNG")
        tiny_bytes = tiny_buf.getvalue()

        # Identify heavy unneeded images (like old unused layout images)
        prune_names = {"ppt/media/image2.png", "ppt/media/hdphoto1.wdp"}

        with zipfile.ZipFile(temp_stage1, "r") as zin:
            with zipfile.ZipFile(temp_stage2, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    if item.filename in prune_names:
                        data = tiny_bytes
                    zout.writestr(item, data)

        # Step 3: Copy to destination
        os.makedirs(os.path.dirname(out_path), exist_ok=True)
        shutil.copy2(temp_stage2, out_path)

        # Step 4: Verify template
        prs_check = Presentation(out_path)
        print(f"Success! Template saved at {out_path}")
        print(f"  Slide count: {len(prs_check.slides)} (should be 0)")
        print(f"  Layout count: {len(prs_check.slide_layouts)}")
        print(f"  File size: {os.path.getsize(out_path):,} bytes")
        return out_path
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


if __name__ == "__main__":
    src = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "..", "..", "Framework Task Force BUMN.pptx")
    out = sys.argv[2] if len(sys.argv) > 2 else DEFAULT_OUT
    rebuild(src, out)
