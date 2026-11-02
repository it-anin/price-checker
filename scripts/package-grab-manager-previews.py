"""Package the ten browser screenshots after desktop/mobile interaction checks."""
from pathlib import Path
import json
import re
import zipfile
from PIL import Image, ImageDraw, ImageFont

base = Path(__file__).resolve().parents[1] / 'public/design-previews/grab-manager'
designs = json.loads((base / 'concepts.json').read_text())
latin = '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'
latin_bold = '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'
heading = ImageFont.truetype(latin_bold, 36)
subtitle = ImageFont.truetype(latin, 17)
name_font = ImageFont.truetype(latin_bold, 25)
thai = ImageFont.truetype(str(base / 'fonts/NotoSansThai-Regular.ttf'), 21)

width, thumb_width, thumb_height, row_height, top = 1600, 752, 499, 590, 160
sheet = Image.new('RGB', (width, top + row_height * 5 + 60), '#f8f6f0')
draw = ImageDraw.Draw(sheet)
draw.text((38, 30), 'GRAB CSV MANAGER / 10 DESIGN CONCEPTS', font=heading, fill='#353029')
draw.text((40, 95), 'ROW CHANGES / BEFORE & AFTER / CSV EXPORT / 10 UNIQUE LAYOUTS', font=subtitle, fill='#947d62')
pages = []
for index, design in enumerate(designs):
    number = f'{design["id"]:02d}'
    source = Image.open(base / 'images' / (number + '.jpg')).convert('RGB')
    pages.append(source)
    thumbnail = source.resize((thumb_width, thumb_height), Image.Resampling.LANCZOS)
    x, y = 32 + index % 2 * 784, top + index // 2 * row_height
    sheet.paste(thumbnail, (x, y))
    draw.rectangle((x, y, x + thumb_width, y + thumb_height), outline='#d8cebe', width=1)
    draw.text((x + 5, y + thumb_height + 13), number + ' / ' + design['name'], font=name_font, fill=design['color'])
    caption_x = x + 5
    # Noto Sans Thai is subsetted; use the Latin font for non-Thai runs.
    for run in re.findall(r'[\u0e00-\u0e7f]+|[^\u0e00-\u0e7f]+', design['thai']):
        font = thai if '\u0e00' <= run[0] <= '\u0e7f' else subtitle
        draw.text((caption_x, y + thumb_height + 48), run, font=font, fill='#8b7e6b')
        caption_x += draw.textlength(run, font=font)
sheet.save(base / 'overview.jpg', quality=90, optimize=True)
pdf = base / 'grab-manager-10-designs.pdf'
pages[0].save(pdf, save_all=True, append_images=pages[1:], resolution=144,
              title='GRAB CSV MANAGER - 10 Design Concepts', author='Price Checker')
assert len(re.findall(rb'/Type /Page\b', pdf.read_bytes())) == 10
archive = base / 'grab-manager-designs.zip'
with zipfile.ZipFile(archive, 'w', zipfile.ZIP_DEFLATED) as bundle:
    for file in sorted(base.rglob('*')):
        if file.is_file() and file != archive:
            bundle.write(file, file.relative_to(base))
with zipfile.ZipFile(archive) as bundle:
    assert bundle.testzip() is None
    assert len([name for name in bundle.namelist() if name.startswith('concepts/') and name.endswith('.html')]) == 10
print('Packaged 10-page PDF, overview image and verified preview ZIP.')
