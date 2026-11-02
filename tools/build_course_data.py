#!/usr/bin/env python3
"""
Rebuild js/course-data.js and images/lessons/ from the Word source of the course book.

Usage (from the project root):
    python3 tools/build_course_data.py "Computer for Office Application (Office 2021).docx"

The Word file is the single source of truth for the lesson text, headings, lists,
tables, formulas and screenshots. Site-only settings that the book doesn't contain
(lesson icons, chapter ids, logo/cover paths, institution names) are kept from the
current js/course-data.js.

Screenshots are written as WebP (smaller = faster pages). Requires Pillow with WebP
support:  python3 -m pip install Pillow
"""
import io
import json
import os
import re
import shutil
import sys
import zipfile
import xml.etree.ElementTree as ET

try:
    from PIL import Image, features
    if not features.check('webp'):
        raise ImportError('Pillow was built without WebP support')
except ImportError as e:
    sys.exit(f'Pillow with WebP support is required ({e}). Install it with: python3 -m pip install Pillow')

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
V = '{urn:schemas-microsoft-com:vml}'

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_PATH = os.path.join(ROOT, 'js', 'course-data.js')
IMG_DIR = os.path.join(ROOT, 'images', 'lessons')
MAX_IMG_WIDTH = 1600

KH2AR = str.maketrans('០១២៣៤៥៦៧៨៩', '0123456789')
AR2KH = str.maketrans('0123456789', '០១២៣៤៥៦៧៨៩')
ZW = re.compile(r'[​‌‍﻿]')


# --------------------------------------------------------------------------- text helpers
def plain(t):
    """Text without zero-width spaces (for matching / search)."""
    return ZW.sub('', t or '').strip()


def clean(t):
    """Display text: keep Khmer zero-width word breaks, tidy everything else."""
    t = (t or '').replace(' ', ' ').replace('\t', ' ')
    t = re.sub(r'<([A-Za-z][A-Za-z0-9 ]{0,15})>', r'\1', t)        # <Ctrl>+<S> -> Ctrl+S
    t = re.sub(r'[ ​]*\n[ ​]*', '\n', t)
    t = re.sub(r' {2,}', ' ', t)
    t = re.sub(r'​*( )​*', r'\1', t)                          # no ZWSP next to spaces
    return t.strip(' ​\n')


def has_khmer(t):
    return re.search(r'[ក-៿]', t or '') is not None


# --------------------------------------------------------------------------- docx reading
class Docx:
    def __init__(self, path):
        self.z = zipfile.ZipFile(path)
        self.doc = ET.fromstring(self.z.read('word/document.xml'))
        rels = ET.fromstring(self.z.read('word/_rels/document.xml.rels'))
        self.rels = {r.get('Id'): r.get('Target') for r in rels}
        self.styles = {}
        st = ET.fromstring(self.z.read('word/styles.xml'))
        for s in st.iter(W + 'style'):
            n = s.find(W + 'name')
            self.styles[s.get(W + 'styleId')] = n.get(W + 'val') if n is not None else s.get(W + 'styleId')
        self._load_numbering()

    def _load_numbering(self):
        self.num_fmt = {}      # numId -> {ilvl: (fmt, start)}
        try:
            num = ET.fromstring(self.z.read('word/numbering.xml'))
        except KeyError:
            return
        abstract = {}
        for a in num.findall(W + 'abstractNum'):
            lv = {}
            for l in a.findall(W + 'lvl'):
                f = l.find(W + 'numFmt'); s = l.find(W + 'start')
                lv[l.get(W + 'ilvl')] = (f.get(W + 'val') if f is not None else 'decimal',
                                         int(s.get(W + 'val')) if s is not None else 1)
            abstract[a.get(W + 'abstractNumId')] = lv
        for n in num.findall(W + 'num'):
            lv = dict(abstract.get(n.find(W + 'abstractNumId').get(W + 'val'), {}))
            for o in n.findall(W + 'lvlOverride'):
                so = o.find(W + 'startOverride')
                il = o.get(W + 'ilvl')
                if so is not None and il in lv:
                    lv[il] = (lv[il][0], int(so.get(W + 'val')))
            self.num_fmt[n.get(W + 'numId')] = lv

    def media(self, target):
        return self.z.read('word/' + target if not target.startswith('word/') else target)

    def para(self, p):
        texts, images, bold_chars, all_chars = [], [], 0, 0
        for r in p.iter(W + 'r'):
            rpr = r.find(W + 'rPr')
            is_bold = rpr is not None and rpr.find(W + 'b') is not None and \
                rpr.find(W + 'b').get(W + 'val') not in ('0', 'false')
            for node in r.iter():
                if node.tag == W + 't':
                    texts.append(node.text or '')
                    n = len((node.text or '').strip())
                    all_chars += n
                    bold_chars += n if is_bold else 0
                elif node.tag == W + 'tab':
                    texts.append('\t')
                elif node.tag in (W + 'br', W + 'cr'):
                    texts.append('\n')
                elif node.tag == A + 'blip':
                    rid = node.get(R + 'embed')
                    if rid in self.rels:
                        images.append(self.rels[rid])
                elif node.tag == V + 'imagedata':
                    rid = node.get(R + 'id')
                    if rid in self.rels:
                        images.append(self.rels[rid])
        ppr = p.find(W + 'pPr')
        style, num = None, None
        if ppr is not None:
            ps = ppr.find(W + 'pStyle')
            if ps is not None:
                style = self.styles.get(ps.get(W + 'val'), ps.get(W + 'val'))
            npr = ppr.find(W + 'numPr')
            if npr is not None and npr.find(W + 'numId') is not None:
                nid = npr.find(W + 'numId').get(W + 'val')
                il = npr.find(W + 'ilvl')
                if nid != '0':
                    num = (nid, il.get(W + 'val') if il is not None else '0')
        return {'k': 'p', 'text': ''.join(texts), 'style': style or '', 'num': num, 'images': images,
                'bold': all_chars > 0 and bold_chars == all_chars}

    def blocks(self):
        body = self.doc.find(W + 'body')
        for child in body:
            if child.tag == W + 'p':
                yield self.para(child)
            elif child.tag == W + 'tbl':
                rows = []
                for tr in child.findall(W + 'tr'):
                    cells = []
                    for tc in tr.findall(W + 'tc'):
                        ps = [self.para(p) for p in tc.iter(W + 'p')]
                        span = tc.find(W + 'tcPr/' + W + 'gridSpan')
                        cells.append({'text': '\n'.join(p['text'] for p in ps if plain(p['text'])),
                                      'span': int(span.get(W + 'val')) if span is not None else 1,
                                      'images': [i for p in ps for i in p['images']],
                                      'bold': all(p['bold'] for p in ps if plain(p['text'])) and any(plain(p['text']) for p in ps)})
                    rows.append(cells)
                yield {'k': 'tbl', 'rows': rows}


# --------------------------------------------------------------------------- images
WEBP_QUALITY = 82          # lossy quality for photos/screenshots (text stays sharp at 82)


def to_webp(data, max_width=MAX_IMG_WIDTH):
    """Encode image bytes as WebP. Tries lossy and lossless and keeps the smaller one:
    lossy wins for big screenshots, lossless for small icons/diagrams. Transparency is kept."""
    img = Image.open(io.BytesIO(data))
    img.load()
    has_alpha = img.mode in ('RGBA', 'LA') or (img.mode == 'P' and 'transparency' in img.info)
    img = img.convert('RGBA' if has_alpha else 'RGB')
    if img.width > max_width:
        img = img.resize((max_width, round(img.height * max_width / img.width)), Image.LANCZOS)
    lossy = io.BytesIO()
    img.save(lossy, 'WEBP', quality=WEBP_QUALITY, method=6, alpha_quality=90)
    candidates = [lossy.getvalue()]
    # Lossless only pays off for small, flat images; on big screenshots it is slow and larger.
    if img.width * img.height <= 400_000:
        lossless = io.BytesIO()
        img.save(lossless, 'WEBP', lossless=True, quality=80, method=4)
        candidates.append(lossless.getvalue())
    return min(candidates, key=len)


class ImageWriter:
    def __init__(self, docx):
        self.docx = docx
        self.done = {}          # media target -> relative path
        self.counters = {}
        self.bytes_in = self.bytes_out = 0

    def write(self, target, lesson_id):
        if target in self.done:
            return self.done[target]
        data = self.docx.media(target)
        self.bytes_in += len(data)
        self.counters[lesson_id] = self.counters.get(lesson_id, 0) + 1
        rel = f'images/lessons/{lesson_id}/fig-{self.counters[lesson_id]:02d}.webp'
        os.makedirs(os.path.join(IMG_DIR, lesson_id), exist_ok=True)
        webp = to_webp(data)
        open(os.path.join(ROOT, rel), 'wb').write(webp)
        self.bytes_out += len(webp)
        self.done[target] = rel
        return rel


# --------------------------------------------------------------------------- lesson builder
CHAPTER_DIVIDER = re.compile(r'^(កុំព្យូទ័រ|ជំពូកទី[០-៩]+\s+Microsoft\s+\S+\s+2021)$')
MANUAL_BULLET = re.compile(r'^\s*(o|\+|🖌|Y|•|-)\t')
MANUAL_LETTER = re.compile(r'^\s*([a-z])\.\t')


def heading_for_caption(t):
    t = plain(re.sub(r'<[^>]+>', '', t))
    t = re.sub(r'^[\s➢❖✓•▪\-\*]+', '', t)
    t = re.sub(r'^[០-៩0-9]+(?:\.[០-៩0-9]+)*\.?\s*', '', t)
    return t.strip().rstrip('៖:').strip()


class LessonBuilder:
    def __init__(self, docx, images, lesson_id, promote_subheadings=False):
        self.docx, self.images, self.lesson_id = docx, images, lesson_id
        self.promote_subheadings = promote_subheadings
        self.objectives = []
        self.sections = []
        self.cur = None
        self.list_block = None      # (key, item dict)
        self.counters = {}
        self.in_objectives = False

    def section(self, title):
        self.cur = {'title': clean(title), 'items': []}
        self.sections.append(self.cur)
        self.list_block = None

    def add(self, item):
        if self.cur is None:
            self.section('សេចក្តីផ្តើមនៃមេរៀន')
        self.cur['items'].append(item)
        if item['type'] not in ('steps', 'bullets'):
            self.list_block = None

    def list_item(self, kind, key, text):
        if self.list_block and self.list_block[0] == (kind, key) and self.cur['items'] and \
                self.cur['items'][-1] is self.list_block[1]:
            blk = self.list_block[1]
        else:
            blk = {'type': kind, ('steps' if kind == 'steps' else 'items'): []}
            self.add(blk)
            self.list_block = ((kind, key), blk)
        blk['steps' if kind == 'steps' else 'items'].append(text)

    def number_label(self, nid, ilvl):
        fmt, start = self.docx.num_fmt.get(nid, {}).get(ilvl, ('decimal', 1))
        lvl_counts = self.counters.setdefault(nid, {})
        lvl_counts[ilvl] = lvl_counts.get(ilvl, start - 1) + 1
        for deeper in [k for k in lvl_counts if int(k) > int(ilvl)]:
            del lvl_counts[deeper]
        n = lvl_counts[ilvl]
        if fmt == 'bullet':
            return 'bullet', None
        if fmt in ('lowerLetter', 'upperLetter'):
            c = chr(ord('a') + n - 1)
            return 'steps', c.upper() if fmt == 'upperLetter' else c
        return 'steps', str(n)

    def image(self, target):
        rel = self.images.write(target, self.lesson_id)
        self.add({'type': 'image', 'src': rel, 'caption': ''})

    def paragraph(self, b):
        raw = b['text']
        text = clean(raw)
        p = plain(raw)
        style = b['style']
        if not p and not b['images']:
            return
        if CHAPTER_DIVIDER.match(p):
            return
        if p.startswith('គោលបំណងនៃមេរៀន'):
            self.in_objectives = True
            return
        if self.in_objectives:
            if style == 'List Paragraph' or b['num']:
                self.objectives.append(text)
                return
            self.in_objectives = False
        if style == 'heading 2':
            self.section(text)
            return
        if style == 'heading 3':
            self.add({'type': 'subsection', 'title': text})
            return
        if style == 'Kh Sub Heading':
            if self.promote_subheadings:
                # lesson has no main headings in the book: its sub-headings become the sections
                self.section(re.sub(r'^[\s➢❖✓•▪]+', '', text))
            else:
                self.add({'type': 'subtopic', 'title': text})
            return
        if p:
            if b['num']:
                nid, ilvl = b['num']
                kind, label = self.number_label(nid, ilvl)
                if kind == 'bullet':
                    self.list_item('bullets', nid + ':' + ilvl, text)
                else:
                    self.list_item('steps', nid + ':' + ilvl, f'{label}. {text}')
            elif MANUAL_LETTER.match(raw.replace('​', '')):
                m = MANUAL_LETTER.match(raw.replace('​', ''))
                self.list_item('steps', 'manual-letter', f'{m.group(1)}. {clean(raw.replace(chr(0x200b), "")[m.end():])}')
            elif MANUAL_BULLET.match(raw.replace('​', '')):
                m = MANUAL_BULLET.match(raw.replace('​', ''))
                self.list_item('bullets', 'manual', clean(raw.replace('​', '')[m.end():]))
            elif p.startswith('=') and not has_khmer(p.split('"')[0]):
                self.add({'type': 'formula', 'text': plain(raw)})
            elif p.startswith('ចំណាំ') or p.startswith('សម្គាល់'):
                self.add({'type': 'callout', 'kind': 'note', 'text': text})
            else:
                self.add({'type': 'paragraph', 'text': text})
        for img in b['images']:
            self.image(img)

    def table(self, t):
        rows = t['rows']
        cells = [c for r in rows for c in r]
        images = [i for c in cells for i in c['images']]
        if len(rows) == 1 and len(rows[0]) == 1:
            txt = rows[0][0]['text']
            p = plain(txt)
            if p:
                looks_formula = (not has_khmer(p)) and re.search(r'\w\(|^=', p)
                if looks_formula:
                    self.add({'type': 'formula', 'text': plain(txt)})
                else:
                    self.add({'type': 'callout', 'kind': 'note', 'text': clean(txt)})
            for img in images:
                self.image(img)
            return
        def is_header_row(row):
            texts = [plain(c['text']) for c in row]
            if not any(texts):
                return False
            if texts[0] in ('ល.រ', 'ល.រ.', 'លរ'):
                return True
            return all(c['bold'] for c in row if plain(c['text']))

        def cell(c):
            # merged cells become {"text", "colspan"}; plain cells stay strings
            t = clean(c['text'])
            return {'text': t, 'colspan': c['span']} if c.get('span', 1) > 1 else t

        header = None
        if len(rows) > 1 and is_header_row(rows[0]):
            header = [cell(c) for c in rows[0]]
            body = rows[1:]
        else:
            body = rows
        body_rows = [[cell(c) for c in r] for r in body]
        ncols = max(sum(c.get('span', 1) for c in r) for r in rows)

        # Word splits long tables at page breaks (and sometimes puts the heading row in its own
        # one-row table). Join a table onto the one right before it when the columns match.
        def width(cells):
            return sum(c['colspan'] if isinstance(c, dict) else 1 for c in cells)

        if header is None and len(rows) == 1 and is_header_row(rows[0]):
            # a heading row on its own: start a new table that the next part will join
            self.add({'type': 'table', 'headers': [cell(c) for c in rows[0]], 'rows': []})
        else:
            items = self.cur['items'] if self.cur else []
            prev = items[-1] if items else None
            # "header-only table, short label (e.g. «ពីមុន»), table body" -> label first, one table
            if prev and prev['type'] == 'paragraph' and len(plain(prev['text'])) <= 25 and len(items) >= 2 \
                    and items[-2]['type'] == 'table' and not items[-2]['rows'] and header is None \
                    and width(items[-2]['headers']) == ncols:
                label, head = items.pop(), items.pop()
                items.append(label)
                items.append(head)
                prev = head
            if prev and prev['type'] == 'table' and header is None and \
                    max([width(r) for r in prev['rows']] + [width(prev['headers'])]) == ncols:
                prev['rows'].extend(body_rows)
            else:
                self.add({'type': 'table', 'headers': header or [], 'rows': body_rows})
        for img in images:
            self.image(img)


# --------------------------------------------------------------------------- main
def main(docx_path):
    docx = Docx(docx_path)
    old = json.loads(open(DATA_PATH, encoding='utf-8').read().split('window.COURSE_DATA =', 1)[1].strip().rstrip(';'))
    old_lessons = {l['number']: l for l in old['lessons']}
    blocks = list(docx.blocks())

    # table of contents -> page numbers
    toc_pages, ref_page = {}, None
    for b in blocks:
        if b['k'] == 'p' and b['style'] == 'toc 1':
            t = plain(b['text'])
            m = re.match(r'មេរៀនទី([០-៩0-9]+).*?\t\s*([0-9]+)\s*$', t, re.S)
            if m:
                toc_pages[int(m.group(1).translate(KH2AR))] = int(m.group(2))
            elif t.startswith('ឯកសារយោង'):
                ref_page = int(re.search(r'([0-9]+)\s*$', t).group(1))

    h1 = [i for i, b in enumerate(blocks) if b['k'] == 'p' and b['style'] == 'heading 1']
    first_h1 = h1[0]

    # ---- front matter: preface ----
    front = [b for b in blocks[:first_h1] if b['k'] == 'p' and not b['style'].startswith('toc')]
    texts = [plain(b['text']) for b in front]
    i_pref = texts.index('អារម្ភកថា')
    i_place = next(i for i, t in enumerate(texts) if t.startswith('សៀមរាប'))
    pref_paras, chapter_desc = [], {}
    for b in front[i_pref + 1:i_place]:
        t = clean(b['text'])
        if b['num'] or b['style'] == 'List Paragraph':
            m = re.match(r'ជំពូក​?ទី​?([០-៩]+)', t.replace('​', ''))
            if m:
                chapter_desc[int(m.group(1).translate(KH2AR))] = clean(t.split('៖', 1)[1]) if '៖' in t else ''
            pref_paras[-1] += '\n- ' + t
        else:
            pref_paras.append(t)
    preface = {'title': 'អារម្ភកថា', 'content': '\n\n'.join(pref_paras),
               'signOff': {'location': texts[i_place].split(',')[0].strip(),
                           'date': texts[i_place].split(',', 1)[1].strip() if ',' in texts[i_place] else '',
                           'role': texts[i_place + 1], 'author': texts[i_place + 2]}}

    # ---- course introduction (first heading 1) ----
    intro_parts = []
    for b in blocks[h1[0] + 1:h1[1]]:
        if b['k'] != 'p':
            continue
        p = plain(b['text'])
        if not p or CHAPTER_DIVIDER.match(p):
            continue
        t = clean(b['text'])
        if b['style'] == 'heading 2':
            intro_parts.append('## ' + t)
        elif b['num'] or b['style'] == 'List Paragraph':
            intro_parts[-1] += '\n- ' + t
        else:
            intro_parts.append(t)
    course_intro = {'title': old['courseIntro']['title'], 'content': '\n\n'.join(intro_parts)}

    # ---- lessons ----
    if os.path.isdir(IMG_DIR):
        shutil.rmtree(IMG_DIR)
    images = ImageWriter(docx)
    lessons = []
    ref_index = None
    for n, start in enumerate(h1):
        title = plain(blocks[start]['text'])
        m = re.match(r'មេរៀនទី([០-៩]+)\s*\n?(.*)', title, re.S)
        if not m:
            if title.startswith('ឯកសារយោង'):
                ref_index = start
            continue
        num = int(m.group(1).translate(KH2AR))
        lesson_id = f'lesson-{num:02d}'
        end = h1[n + 1] if n + 1 < len(h1) else len(blocks)
        has_h2 = any(b['k'] == 'p' and b['style'] == 'heading 2' for b in blocks[start + 1:end])
        lb = LessonBuilder(docx, images, lesson_id, promote_subheadings=not has_h2)
        for b in blocks[start + 1:end]:
            if b['k'] == 'p':
                lb.paragraph(b)
            else:
                lb.table(b)
        # a lesson with no headings at all is one block of content, not an "introduction"
        if len(lb.sections) == 1 and lb.sections[0]['title'] == 'សេចក្តីផ្តើមនៃមេរៀន':
            lb.sections[0]['title'] = 'ខ្លឹមសារមេរៀន'
        # captions from the nearest heading
        for sec in lb.sections:
            heading, seen, per = heading_for_caption(sec['title']), {}, {}
            imgs = []
            for it in sec['items']:
                if it['type'] in ('subsection', 'subtopic'):
                    heading = heading_for_caption(it['title'])
                elif it['type'] == 'image':
                    imgs.append((it, heading)); per[heading] = per.get(heading, 0) + 1
            for it, h in imgs:
                seen[h] = seen.get(h, 0) + 1
                it['caption'] = h + (f' រូបទី{str(seen[h]).translate(AR2KH)}' if per[h] > 1 else '')
        o = old_lessons[num]
        heading_title = blocks[start]['text']
        clean_title = clean(heading_title.split('\n', 1)[1] if '\n' in heading_title else m.group(2))
        chap = next(c for c in old['chapters'] if lesson_id in c['lessons'])
        lessons.append({
            'id': lesson_id, 'number': num, 'chapterId': chap['id'],
            'chapterTitle': f"{chap['number']}៖ {chap['title']}",
            'title': clean_title,
            'pageStart': toc_pages.get(num), 'pageEnd': None,
            'icon': o.get('icon', 'file-text'),
            'objectives': lb.objectives, 'sections': lb.sections,
        })
    for i, l in enumerate(lessons):
        nxt = lessons[i + 1]['pageStart'] if i + 1 < len(lessons) else ref_page
        l['pageEnd'] = (nxt - 1) if nxt else l['pageStart']

    # ---- references ----
    references = []
    if ref_index is not None:
        for b in blocks[ref_index + 1:]:
            if b['k'] != 'p':
                continue
            t = plain(b['text'])
            m = re.match(r'(.+?),\s*(\d{4})\s*["“](.+?)["”],\s*Published by\s*(.+?)\.?$', t)
            if m:
                lccn = re.search(r'Control Number:\s*(\d+)', m.group(4))
                references.append({'author': m.group(1).strip(), 'year': m.group(2), 'title': m.group(3).strip(),
                                   'publisher': re.sub(r',?\s*Library of Congress.*$', '', m.group(4)).strip(),
                                   **({'lccn': lccn.group(1)} if lccn else {})})

    chapters = []
    for c in old['chapters']:
        idx = int(c['number'].replace('ជំពូកទី', '').translate(KH2AR))
        chapters.append(dict(c, description=chapter_desc.get(idx, c['description']),
                             lessons=[l['id'] for l in lessons if l['chapterId'] == c['id']]))

    meta = dict(old['meta'])
    # page images use WebP; images/logo.png stays only for the favicon / app icons
    meta['logoImage'] = 'images/logo.webp'
    meta['coverImage'] = 'images/cover.webp'
    meta['totalLessons'] = len(lessons)
    meta['totalChapters'] = len(chapters)
    if ref_page:
        meta['totalPages'] = ref_page

    data = {'meta': meta, 'preface': preface, 'courseIntro': course_intro, 'chapters': chapters,
            'lessons': lessons, 'references': references or old['references']}
    header = ('/* Course content generated from the Word book by tools/build_course_data.py.\n'
              '   Edit the Word file and re-run the script instead of editing this file by hand. */\n')
    open(DATA_PATH, 'w', encoding='utf-8').write(
        header + 'window.COURSE_DATA = ' + json.dumps(data, ensure_ascii=False, indent=2) + ';\n')

    print(f'lessons: {len(lessons)}  images: {len(images.done)} '
          f'({images.bytes_in / 1e6:.1f} MB -> {images.bytes_out / 1e6:.1f} MB)  references: {len(references)}')


if __name__ == '__main__':
    main(sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, 'Computer for Office Application (Office 2021).docx'))
