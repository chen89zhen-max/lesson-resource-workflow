"""Check explicit text sizes, bounds, notes and package paths; render for human review.

Exit 0 = automated checks and rendering complete, not visual/teaching approval.
Exit 1 = detected error. Exit 2 = rendering unavailable or intentionally omitted.
Dependencies: python-pptx, Pillow, lxml; LibreOffice + PyMuPDF or Poppler.
"""
import argparse
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import tempfile
from zipfile import ZipFile
from lxml import etree
from PIL import Image, ImageDraw
from pptx import Presentation

NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
LOCAL_PATH = re.compile(r'(?:[A-Za-z]:[\\/](?:Users|Documents and Settings)[\\/]|/(?:Users|home)/[^/\s]+/|file:/)', re.I)


def shapes_in(shapes):
    for shape in shapes:
        yield shape
        if hasattr(shape, 'shapes'):
            yield from shapes_in(shape.shapes)


def text_size(paragraph, run):
    if run.font.size is not None:
        return run.font.size.pt
    if paragraph.font.size is not None:
        return paragraph.font.size.pt
    return None


def inspect(pptx, template=False):
    errors, warnings, stats = [], [], []
    with ZipFile(pptx) as archive:
        for name in archive.namelist():
            if name.endswith(('.xml', '.rels')):
                value = archive.read(name).decode('utf-8')
                if LOCAL_PATH.search(value):
                    errors.append(f'{name}: 存在本机个人路径或文件链接')
        if 'docProps/core.xml' in archive.namelist():
            root = etree.fromstring(archive.read('docProps/core.xml'))
            if any((el.text or '').strip() for el in root if etree.QName(el).localname in ('creator', 'lastModifiedBy')):
                warnings.append('作者属性非空：公开发布前核对是否含个人信息。')
    prs = Presentation(str(pptx))
    if abs(prs.slide_width / prs.slide_height - 16/9) > .005:
        errors.append('画幅应为16:9。')
    if not prs.slides:
        errors.append('课件没有页面。')
    for index, slide in enumerate(prs.slides, 1):
        body_sizes = []
        for shape in shapes_in(slide.shapes):
            if any(v is None for v in (shape.left, shape.top, shape.width, shape.height)):
                continue
            tolerance = 28575  # 3 px at 96 dpi
            if shape.left < -tolerance or shape.top < -tolerance or shape.left + shape.width > prs.slide_width+tolerance or shape.top + shape.height > prs.slide_height+tolerance:
                errors.append(f'第{index}页对象越界：{shape.name}')
            frames = []
            if shape.has_text_frame:
                frames.append(shape.text_frame)
            if shape.has_table:
                frames.extend(cell.text_frame for row in shape.table.rows for cell in row.cells if not cell.is_spanned)
            for frame in frames:
                text = frame.text.strip()
                if not text:
                    continue
                if not template and re.search(r'【[^】]+】', text) and shape.name.startswith('slot:'):
                    errors.append(f'第{index}页仍有模板占位文字：{shape.name}')
                # Only explicit role names exempt navigation metadata; position is not a role.
                metadata = shape.name.startswith('meta:')
                for paragraph in frame.paragraphs:
                    for run in paragraph.runs:
                        if not run.text.strip():
                            continue
                        size = text_size(paragraph, run)
                        if size is None:
                            warnings.append(f'第{index}页字号继承未解析：{shape.name}，须人工核对。')
                        elif not metadata:
                            # Two-digit question counters are navigation metadata, not teaching text.
                            if re.fullmatch(r'0[1-9]', text) and size >= 18:
                                continue
                            body_sizes.append(size)
                            if size < 24 - .01:
                                errors.append(f'第{index}页主要文字小于24磅（{size:g}）：{shape.name}')
        notes = slide.notes_slide.notes_text_frame.text.strip() if slide.has_notes_slide else ''
        if not notes:
            errors.append(f'第{index}页缺少教师备注。')
        elif not template and notes.startswith('模板页：'):
            errors.append(f'第{index}页教师备注尚未替换。')
        stats.append({'slide': index, 'min_core_text_pt': min(body_sizes) if body_sizes else None, 'notes_present': bool(notes)})
    return {'slide_count': len(prs.slides), 'errors': list(dict.fromkeys(errors)), 'warnings': list(dict.fromkeys(warnings)), 'slides': stats}


def inspect_teaching(pptx, manifest=None):
    """Audit the declared teaching pages; semantic adequacy still needs review."""
    errors = []
    path = Path(manifest) if manifest else Path(pptx).with_suffix('.teaching-pages.json')
    result = {'manifest': str(path), 'effective_teaching_pages': 0,
              'minimum_effective_pages': 12, 'excluded_pages': [],
              'teaching_review': 'pending: 逐页核对实际内容、任务、难点展开与反馈，清单不能自证教学合格。',
              'errors': errors}
    prs = Presentation(str(pptx))
    if len(prs.slides) < 12:
        errors.append('实际总页数不足12，不可能满足12页有效教学内容。')
    try:
        data = json.loads(path.read_text(encoding='utf-8-sig'))
    except (OSError, ValueError) as exc:
        errors.append('缺失或无法读取有效教学页清单：' + str(exc))
        return result
    if not isinstance(data, dict) or data.get('schema_version') != 1 or not isinstance(data.get('pages'), list):
        errors.append('有效教学页清单须为schema_version=1且包含pages数组。')
        return result
    pages = data['pages']
    numbers = [p.get('page') if isinstance(p, dict) else None for p in pages]
    if any(type(n) is not int for n in numbers) or numbers != list(range(1, len(prs.slides)+1)):
        errors.append('清单必须按实际顺序覆盖所有页面，page为连续且唯一的整数。')
        return result
    seen = {}
    for item, slide in zip(pages, prs.slides):
        number, role = item['page'], item.get('role')
        if role not in {'teaching', 'cover', 'agenda', 'transition', 'closing'}:
            errors.append(f'第{number}页role无效。')
            continue
        if role != 'teaching':
            result['excluded_pages'].append(number)
            continue
        if any(not isinstance(item.get(key), str) or not item[key].strip()
               for key in ['purpose', 'task_or_evidence']):
            errors.append(f'第{number}页缺少具体purpose或task_or_evidence。')
            continue
        content = []
        substantive = False
        for shape in shapes_in(slide.shapes):
            if shape.name.startswith('meta:'):
                continue
            text = shape.text_frame.text.strip() if shape.has_text_frame else ''
            if shape.has_table:
                text += '|'.join(cell.text for row in shape.table.rows for cell in row.cells)
            if text or shape.shape_type == 13:
                substantive = True
            image_hash = hashlib.sha256(shape.image.blob).hexdigest() if shape.shape_type == 13 else ''
            content.append((str(shape.shape_type), text, image_hash,
                            shape.left, shape.top, shape.width, shape.height))
        if not substantive:
            # Pure diagram pages require human checking; don't count an unlabelled blank.
            errors.append(f'第{number}页缺少可识别教学文字或图片，须补充可核对标签及教学内容。')
            continue
        signature = repr(content)
        if signature in seen:
            errors.append(f'第{number}页与第{seen[signature]}页教学内容完全重复，不能重复计数。')
            continue
        seen[signature] = number
        result['effective_teaching_pages'] += 1
    if result['effective_teaching_pages'] < 12:
        errors.append(f"有效教学页仅{result['effective_teaching_pages']}，每次连堂课至少12页。")
    return result


def executable(explicit, name):
    if explicit:
        found = Path(explicit).expanduser()
        if found.is_file():
            return str(found.resolve())
        raise RuntimeError(f'指定程序不存在：{explicit}')
    found = shutil.which(name)
    if found:
        return found
    if name == 'soffice':
        for root in [os.environ.get('PROGRAMFILES'), os.environ.get('PROGRAMFILES(X86)')]:
            if root:
                for suffix in ['soffice.com', 'soffice.exe']:
                    candidate = Path(root) / 'LibreOffice/program' / suffix
                    if candidate.is_file():
                        return str(candidate)
        candidate = Path('/Applications/LibreOffice.app/Contents/MacOS/soffice')
        if candidate.is_file():
            return str(candidate)
    return None


def run(command):
    return subprocess.run(command, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                          timeout=120, creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))


def render(pptx, directory, soffice=None, pdftoppm=None):
    office = executable(soffice, 'soffice')
    if not office:
        raise RuntimeError('未找到LibreOffice；可用 --soffice 指定。可另用宿主渲染工具，但须记录实际检查范围。')
    directory.mkdir(parents=True, exist_ok=True)
    # Every run uses a new directory and isolated profile to avoid stale output.
    output = Path(tempfile.mkdtemp(prefix='render-', dir=str(directory.resolve())))
    profile = output / 'profile'
    run([office, '-env:UserInstallation='+profile.as_uri(), '--headless', '--convert-to', 'pdf', '--outdir', str(output), str(pptx.resolve())])
    pdf = output / (pptx.stem + '.pdf')
    if not pdf.is_file():
        raise RuntimeError('LibreOffice未生成本次PDF。')
    images = []
    try:
        import fitz
    except ImportError:
        fitz = None
    if fitz is not None:
        with fitz.open(pdf) as doc:
            for number, page in enumerate(doc, 1):
                target = output / f'page-{number}.png'
                page.get_pixmap(dpi=144, alpha=False).save(target)
                images.append(target)
    else:
        poppler = executable(pdftoppm, 'pdftoppm')
        if not poppler:
            raise RuntimeError('需要PyMuPDF或Poppler；可用 --pdftoppm 指定程序。')
        run([poppler, '-scale-to', '1600', '-png', str(pdf), str(output/'page')])
        images = sorted(output.glob('page-*.png'), key=lambda p: int(p.stem.split('-')[-1]))
    if not images:
        raise RuntimeError('未产生逐页预览。')
    sheets = []
    for first in range(0, len(images), 6):
        sheet = Image.new('RGB', (1920, 830), '#E7EBE8')
        draw = ImageDraw.Draw(sheet)
        for k, source in enumerate(images[first:first+6]):
            with Image.open(source) as image:
                thumb = image.convert('RGB');thumb.thumbnail((620, 349))
                x, y = 20+(k % 3)*638, 30+(k//3)*397
                sheet.paste(thumb, (x, y))
                draw.text((x, y+355), f'P{first+k+1:02}', fill='#172D36')
        path = output / f'sheet-{first//6+1}.png';sheet.save(path);sheets.append(str(path))
    return [str(p) for p in images], sheets


def main():
    parser = argparse.ArgumentParser(description='PPTX自动检查与实际渲染；通过不等于视觉/知识审查通过。')
    parser.add_argument('pptx', type=Path)
    parser.add_argument('--render-dir', type=Path)
    parser.add_argument('--soffice')
    parser.add_argument('--pdftoppm')
    parser.add_argument('--teaching-manifest', type=Path, help='有效教学页清单；默认同名.teaching-pages.json')
    parser.add_argument('--template', action='store_true', help='只用于通用模板自身，允许保留占位文字与备注')
    parser.add_argument('--structural-only', action='store_true', help='仅检查结构，退出码2表示渲染尚未完成')
    args = parser.parse_args()
    pptx = args.pptx.expanduser().resolve()
    directory = (args.render_dir or pptx.parent/'_render_check'/pptx.stem).resolve()
    directory.mkdir(parents=True, exist_ok=True)
    result = inspect(pptx, args.template)
    if not args.template:
        result['teaching_content'] = inspect_teaching(pptx, args.teaching_manifest)
        result['errors'].extend(result['teaching_content']['errors'])
    else:
        result['teaching_content'] = {'status': 'not_applicable: 仅模板维护，不能用作课堂成品验收。'}
    result['render_complete'] = False
    result['visual_review'] = 'pending: 必须另行逐页看图，核对美学与专业图示。'
    if not args.structural_only:
        try:
            result['slide_pngs'], result['contact_sheets'] = render(pptx, directory, args.soffice, args.pdftoppm)
            if len(result['slide_pngs']) != result['slide_count']:
                result['errors'].append('渲染页数与PPTX页数不一致。')
            else:
                result['render_complete'] = True
        except (RuntimeError, subprocess.SubprocessError, OSError) as error:
            result['render_error'] = str(error)
    result['automated_pass'] = not result['errors'] and result['render_complete']
    (directory/'report.json').write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 1 if result['errors'] else (0 if result['render_complete'] else 2)


if __name__ == '__main__':
    raise SystemExit(main())
