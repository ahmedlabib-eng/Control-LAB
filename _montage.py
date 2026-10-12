from PIL import Image, ImageDraw
import glob, os

files = sorted(glob.glob('_media/*.png'), key=lambda p: int(''.join(c for c in os.path.basename(p) if c.isdigit())))
CELL_W, CELL_H = 800, 620
LABEL_H = 34
per_sheet = 4
os.makedirs('_sheets', exist_ok=True)

for s in range(0, len(files), per_sheet):
    chunk = files[s:s + per_sheet]
    cols, rows = 2, 2
    sheet = Image.new('RGB', (cols * CELL_W, rows * (CELL_H + LABEL_H)), 'white')
    d = ImageDraw.Draw(sheet)
    for i, f in enumerate(chunk):
        im = Image.open(f).convert('RGB')
        im.thumbnail((CELL_W - 20, CELL_H - 20))
        cx = (i % cols) * CELL_W
        cy = (i // cols) * (CELL_H + LABEL_H)
        sheet.paste(im, (cx + 10, cy + LABEL_H + 10))
        d.rectangle([cx, cy, cx + CELL_W - 1, cy + LABEL_H], fill='black')
        d.text((cx + 8, cy + 10), f'{os.path.basename(f)}  ({Image.open(f).size[0]}x{Image.open(f).size[1]})', fill='white')
        d.rectangle([cx, cy, cx + CELL_W - 1, cy + CELL_H + LABEL_H - 1], outline='gray')
    out = f'_sheets/sheet{s // per_sheet + 1}.png'
    sheet.save(out)
    print(out, sheet.size)
