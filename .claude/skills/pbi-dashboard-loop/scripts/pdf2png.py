# Chuyen PDF render cua Power BI thanh PNG tung trang de review bang mat
# Dung: python pdf2png.py [render.pdf] [thu_muc_png]   (mac dinh ./review/render.pdf -> ./review/render_png)
import sys
from pathlib import Path

import pypdfium2 as pdfium

pdf_path = sys.argv[1] if len(sys.argv) > 1 else str(Path("review") / "render.pdf")
out_dir = Path(sys.argv[2]) if len(sys.argv) > 2 else Path(pdf_path).parent / "render_png"
out_dir.mkdir(parents=True, exist_ok=True)

doc = pdfium.PdfDocument(pdf_path)
for i, page in enumerate(doc):
    bmp = page.render(scale=1.6)  # ~1440px wide voi trang 16:10
    img = bmp.to_pil()
    p = out_dir / f"page_{i+1}.png"
    img.save(p)
    print(f"OK page {i+1} -> {p} ({img.width}x{img.height})")
doc.close()
