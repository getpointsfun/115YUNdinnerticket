"""把 QR 套進設計師的憑證模板(向量 PDF)。

用法: python build_cards.py 模板.pdf 代號清單.txt 輸出.pdf
  - 代號清單:每行一個代號(例如民眾的 PB15SK)
  - 白框位置沿用貴賓版驗證過的座標;QR 為向量,內容就是代號本身(大寫)
  - 需要:pymupdf、segno;檢查掃描另需 zxing-cpp、pillow
"""
import sys, pymupdf, segno
tpl, codes_file, out_file = sys.argv[1:4]
codes = [l.strip().upper() for l in open(codes_file, encoding='utf-8') if l.strip()]
FONT = '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'
src = pymupdf.open(tpl); W, H = src[0].rect.width, src[0].rect.height
BOX = pymupdf.Rect(57.1937, 156.0146, 197.9237, 286.3736)   # 模板中的白色圓角框
k = W / 1385.0; MOD = 21 * k; FS = 58 * k                    # 每個 QR 方塊 21px(1385px 寬基準)
font = pymupdf.Font(fontfile=FONT)
out = pymupdf.open()
for code in codes:
    page = out.new_page(width=W, height=H)
    page.show_pdf_page(page.rect, src, 0)
    m = [list(r) for r in segno.make(code, error='m', micro=False).matrix]; n = len(m)
    qx = (BOX.x0 + BOX.x1) / 2 - n * MOD / 2; qy = BOX.y0 + 4 * MOD
    sh = page.new_shape()
    for y, row in enumerate(m):
        for x, v in enumerate(row):
            if v: sh.draw_rect(pymupdf.Rect(qx + x * MOD, qy + y * MOD, qx + (x + 1) * MOD, qy + (y + 1) * MOD))
    sh.finish(fill=(0, 0, 0), color=None, width=0); sh.commit()
    tw = font.text_length(code, fontsize=FS)
    page.insert_text(((BOX.x0 + BOX.x1) / 2 - tw / 2, qy + n * MOD + 3 * MOD + 0.73 * FS), code, fontsize=FS, fontname='dvb', fontfile=FONT, color=(0, 0, 0))
out.save(out_file, garbage=4, deflate=True, clean=True)
print('頁數', len(out))
