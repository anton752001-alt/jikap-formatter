 """
JIKAP Manuscript Formatter v4 - Streamlit Web App
Jurnal Informasi dan Komunikasi Administrasi Perkantoran
PAP FKIP Universitas Sebelas Maret
"""

import streamlit as st
import io, re
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

FONT_ARIAL = "Arial"
FONT_TNR   = "Times New Roman"

SECTIONS_H1 = {
    "introduction","pendahuluan","research methods","metode penelitian",
    "metodologi penelitian","method","research method",
    "results and discussion","hasil dan pembahasan","hasil penelitian dan pembahasan",
    "conclusion","kesimpulan","conclusions","references","daftar pustaka",
    "acknowledgments","acknowledgement","ucapan terima kasih",
}
SECTIONS_H2 = {"research results","hasil penelitian","discussion","pembahasan"}

# ── Sentence case converter ───────────────────────────────────────────────────

def title_to_sentence_case(text):
    """
    Sentence case: first word capitalized, rest lowercase.
    Preserve all-caps abbreviations (SPSS, UNS, PLP, IEQ, UMKM, dll.)
    """
    if not text.strip(): return text
    words = text.split(' ')
    result = []
    for i, word in enumerate(words):
        if not word: result.append(word); continue
        core = re.sub(r'[^a-zA-Z0-9]', '', word)
        if not core: result.append(word); continue
        # All-caps abbreviation (2+ uppercase letters, optional trailing digits)
        if re.match(r'^[A-Z]{2,}[0-9]*$', core):
            result.append(word)
        elif i == 0:
            result.append(word[0].upper() + word[1:].lower())
        else:
            result.append(word.lower())
    return ' '.join(result)

# ── Sub/superscript auto-detection ───────────────────────────────────────────

# Superscript patterns: R2, Y2 (single uppercase letter + digit, not part of longer word)
SUPER_RE = re.compile(r'\b([RYryCc])(\d+)\b')
# Subscript patterns: X1, X2, H1, H2, β1, β2 (variable + digit)
SUB_RE   = re.compile(r'\b([A-Za-zβαγμσ])(\d+)\b')

def split_run_with_sub_super(text):
    """
    Split text into segments: (text, is_subscript, is_superscript)
    Rules:
    - R2, Y2, C2 etc → superscript on digit
    - X1, X2, H1, H2, β1 etc → subscript on digit
    - Patterns only matched at word boundaries
    """
    segments = []
    pos = 0

    # Merge both patterns, process left to right
    all_matches = []
    for m in SUPER_RE.finditer(text):
        all_matches.append((m.start(), m.end(), m.group(1), m.group(2), 'super'))
    for m in SUB_RE.finditer(text):
        # Skip if already captured as super
        already = any(am[0]==m.start() and am[4]=='super' for am in all_matches)
        if not already:
            all_matches.append((m.start(), m.end(), m.group(1), m.group(2), 'sub'))

    all_matches.sort(key=lambda x: x[0])

    # Remove overlaps (keep first)
    filtered = []
    last_end = 0
    for m in all_matches:
        if m[0] >= last_end:
            filtered.append(m)
            last_end = m[1]

    for (start, end, var, num, mode) in filtered:
        if pos < start:
            segments.append((text[pos:start], False, False))
        segments.append((var, False, False))
        if mode == 'super':
            segments.append((num, False, True))
        else:
            segments.append((num, True, False))
        pos = end

    if pos < len(text):
        segments.append((text[pos:], False, False))

    return segments if segments else [(text, False, False)]

def rebuild_runs_with_sub_super(para, fname, fsize, bold=False, italic=False):
    """
    Rebuild all runs in para with sub/superscript detection applied.
    Preserves existing explicit sub/super from input.
    """
    # Collect original runs with their explicit sub/super
    original = []
    for run in para.runs:
        original.append({
            'text': run.text,
            'sub': run.font.subscript,
            'super': run.font.superscript,
        })

    # Clear all runs
    for run in para.runs:
        run.text = ''

    first_run = para.runs[0] if para.runs else None

    def add_run(text, sub, sup):
        nonlocal first_run
        if first_run is not None:
            r = first_run; r.text = text; first_run = None
        else:
            r = para.add_run(text)
        set_run_font(r, fname, fsize, bold=bold, italic=italic,
                     subscript=sub, superscript=sup)

    for orig in original:
        t = orig['text']
        if not t: continue
        # If run already had explicit sub/super, preserve it
        if orig['sub'] or orig['super']:
            add_run(t, orig['sub'], orig['super'])
            continue
        # Otherwise apply auto-detection
        segs = split_run_with_sub_super(t)
        for seg_text, is_sub, is_sup in segs:
            if seg_text:
                add_run(seg_text, is_sub, is_sup)

# ── Reference parser ──────────────────────────────────────────────────────────

def clean_text(t): return re.sub(r'\s+', ' ', t).strip()

def parse_author_string(raw):
    raw = re.sub(r'\s*&\s*$','',raw.strip()).strip()
    parts = re.split(r',\s*&\s*|\s*&\s*|;\s*', raw)
    authors=[]; i=0
    while i < len(parts):
        part = parts[i].strip()
        if not part: i+=1; continue
        if i+1 < len(parts) and re.match(r'^([A-Z]\.[\s-]?)+$', parts[i+1].strip()):
            authors.append((part, parts[i+1].strip())); i+=2
        elif ',' in part:
            ci = part.index(',')
            authors.append((part[:ci].strip(), part[ci+1:].strip())); i+=1
        else:
            authors.append((part,'')); i+=1
    return authors

def format_authors_apa7(authors):
    if not authors: return ''
    fmt=[]
    for last,inits in authors:
        ic = re.sub(r'([A-Z])(?!\.)(?!\s*[A-Z]\.)',r'\1.',inits)
        ic = re.sub(r'\s+',' ',ic).strip()
        fmt.append(f"{last}, {ic}" if ic else last)
    if len(fmt)==1: return fmt[0]
    if len(fmt)<=20: return ', '.join(fmt[:-1])+', & '+fmt[-1]
    return ', '.join(fmt[:19])+', . . . '+fmt[-1]

def detect_ref_type(t):
    tl=t.lower()
    if 'http' in tl or 'doi' in tl or 'retrieved from' in tl:
        if any(x in tl for x in ['skripsi','thesis','tesis','dissertation','disertasi']): return 'thesis_online'
        if any(x in tl for x in ['journal','jurnal']) or re.search(r'\d+\(\d+\)',tl): return 'journal'
        return 'webpage'
    if any(x in tl for x in ['skripsi','thesis','tesis','dissertation','disertasi']): return 'thesis'
    if re.search(r',\s*\d+\s*\(\d+\)',t): return 'journal'
    return 'book'

def parse_reference(ref_text):
    ref_text=clean_text(ref_text)
    r={'raw':ref_text,'authors':'','year':'','title':'','source':'','volume':'','issue':'',
       'pages':'','doi':'','publisher':'','ref_type':'journal','url':'','thesis_info':''}
    dm=re.search(r'https?://doi\.org/\S+|doi:\s*\S+',ref_text,re.IGNORECASE)
    if dm: r['doi']=dm.group().strip().rstrip('.'); ref_text=ref_text[:dm.start()].strip()
    um=re.search(r'https?://\S+',ref_text)
    if um and not r['doi']: r['url']=um.group().strip().rstrip('.'); ref_text=ref_text[:um.start()].strip()
    ym=re.search(r'\((\d{4}[a-z]?)\)',ref_text)
    if ym:
        r['year']=ym.group(1); r['authors']=ref_text[:ym.start()].strip().rstrip('.,')
        rest=ref_text[ym.end():].strip().lstrip('.,').strip()
    else: r['authors']=ref_text; rest=''
    rt=detect_ref_type(ref_text); r['ref_type']=rt
    if rt=='journal':
        parts=[p.strip() for p in rest.split('.') if p.strip()]
        if parts: r['title']=parts[0]
        if len(parts)>=2:
            jp=parts[1]
            vm=re.search(r',\s*(\d+)\s*\((\d+)\)\s*,\s*([\d\u2013\-]+)',jp)
            if vm: r['source']=jp[:vm.start()].strip(); r['volume']=vm.group(1); r['issue']=vm.group(2); r['pages']=vm.group(3)
            else: r['source']=jp
    elif rt=='book':
        parts=[p.strip() for p in rest.split('.') if p.strip()]
        if parts: r['title']=parts[0]
        if len(parts)>=2: r['publisher']=parts[1]
    elif rt in ('thesis','thesis_online'):
        parts=[p.strip() for p in rest.split('.') if p.strip()]
        if parts: r['title']=parts[0]
        tm=re.search(r'\(([^)]+(?:skripsi|thesis|tesis|dissertation|disertasi)[^)]*)\)',rest,re.IGNORECASE)
        if tm: r['thesis_info']=tm.group(1)
        im=re.search(r'(?:Universitas|University|Institut|Institute)[^.,)]+',rest,re.IGNORECASE)
        if im: r['source']=im.group().strip()
    elif rt=='webpage':
        parts=[p.strip() for p in rest.split('.') if p.strip()]
        if parts: r['title']=parts[0]
        if len(parts)>=2: r['source']=parts[1]
    return r

def format_reference_apa7(parsed):
    r=parsed
    try: authors_str=format_authors_apa7(parse_author_string(r['authors']))
    except: authors_str=r['authors']
    year=f"({r['year']})" if r['year'] else ''
    segs=[]
    if r['ref_type']=='journal':
        vi=f"{r['volume']}({r['issue']})" if r['volume'] and r['issue'] else r['volume']
        pages=r['pages'].strip().rstrip('.') if r['pages'] else ''
        segs.append((f"{authors_str} {year}. {r['title'].strip().rstrip('.')}. ",False))
        jp=r['source'].strip().rstrip('.')
        if vi: jp+=f", {vi}"
        if pages: jp+=f", {pages}"
        segs.append((jp,True))
        if r['doi']: segs.append((f". {r['doi']}",False))
        elif r['url']: segs.append((f". {r['url']}",False))
        else: segs.append((".",False))
    elif r['ref_type']=='book':
        segs.append((f"{authors_str} {year}. ",False))
        segs.append((f"{r['title'].strip().rstrip('.')}. ",True))
        segs.append((f"{r['publisher'].strip().rstrip('.')}.",False))
    elif r['ref_type'] in ('thesis','thesis_online'):
        ti=r['thesis_info'] if r['thesis_info'] else 'Skripsi'
        src=r['source'].strip().rstrip('.') if r['source'] else ''
        segs.append((f"{authors_str} {year}. ",False))
        segs.append((r['title'].strip().rstrip('.'),True))
        segs.append((f" ({ti}, {src})." if src else f" ({ti}).",False))
        if r['url']: segs.append((f" Retrieved from {r['url']}",False))
        if r['doi']: segs.append((f" {r['doi']}",False))
    elif r['ref_type']=='webpage':
        segs.append((f"{authors_str} {year}. ",False))
        segs.append((f"{r['title'].strip().rstrip('.')}. ",True))
        if r['source']: segs.append((f"{r['source'].strip().rstrip('.')}.",False))
        if r['url']: segs.append((f" {r['url']}",False))
        if r['doi']: segs.append((f" {r['doi']}",False))
    else: segs.append((r['raw'],False))
    return segs

# ── Font / paragraph helpers ─────────────────────────────────────────────────

def set_page_margins(doc):
    for section in doc.sections:
        section.page_width=Cm(21.0); section.page_height=Cm(29.7)
        section.top_margin=Cm(3.0); section.bottom_margin=Cm(3.0)
        section.left_margin=Cm(3.0); section.right_margin=Cm(3.0)

def set_run_font(run, fname, fsize, bold=False, italic=False,
                 subscript=None, superscript=None):
    run.font.name=fname; run.font.size=Pt(fsize)
    run.font.bold=bold; run.font.italic=italic
    if subscript is not None: run.font.subscript=subscript
    if superscript is not None: run.font.superscript=superscript
    rpr=run._r.get_or_add_rPr()
    rFonts=rpr.find(qn('w:rFonts'))
    if rFonts is None: rFonts=OxmlElement('w:rFonts'); rpr.insert(0,rFonts)
    rFonts.set(qn('w:ascii'),fname); rFonts.set(qn('w:hAnsi'),fname); rFonts.set(qn('w:cs'),fname)

def clear_pf(para):
    """Clear paragraph formatting and enforce space before/after = 0 (point 1)."""
    pf=para.paragraph_format
    pf.space_before=Pt(0); pf.space_after=Pt(0)  # FIX #1
    pf.first_line_indent=None; pf.left_indent=None; pf.right_indent=None

def get_full_text(para): return ''.join(r.text for r in para.runs).strip()

def apply_font_simple(para, fname, fsize, bold=False, italic=False):
    """Apply font preserving existing explicit sub/superscript."""
    for run in para.runs:
        sub = run.font.subscript; sup = run.font.superscript
        set_run_font(run, fname, fsize, bold=bold, italic=italic,
                     subscript=sub, superscript=sup)

# ── Tab-delimited → Word Table converter ─────────────────────────────────────

def is_tab_row(text):
    return '\t' in text and len(text.strip()) > 0

def insert_table_after(doc, ref_para, rows_data):
    if not rows_data: return None
    n_cols = max(len(row) for row in rows_data)
    rows_data = [row + ['']*(n_cols-len(row)) for row in rows_data]
    page_width_dxa = 8505
    col_width = page_width_dxa // n_cols

    table = doc.add_table(rows=len(rows_data), cols=n_cols)

    for row in table.rows:
        for cell in row.cells:
            tc=cell._tc; tcPr=tc.find(qn('w:tcPr'))
            if tcPr is None: tcPr=OxmlElement('w:tcPr'); tc.insert(0,tcPr)
            tcW=OxmlElement('w:tcW'); tcW.set(qn('w:w'),str(col_width)); tcW.set(qn('w:type'),'dxa')
            old=tcPr.find(qn('w:tcW'))
            if old is not None: tcPr.remove(old)
            tcPr.append(tcW)

    for ri, row_data in enumerate(rows_data):
        for ci, cell_text in enumerate(row_data):
            cell=table.cell(ri,ci); para=cell.paragraphs[0]; para.clear()
            run=para.add_run(cell_text)
            # FIX #7: table header NOT bold — all cells TNR 10pt normal
            set_run_font(run,FONT_TNR,10,bold=False)
            para.alignment=WD_ALIGN_PARAGRAPH.LEFT
            para.paragraph_format.space_before=Pt(0); para.paragraph_format.space_after=Pt(0)

    # Borders: top, header-bottom, table-bottom only
    tbl=table._tbl; tblPr=tbl.find(qn('w:tblPr'))
    if tblPr is None: tblPr=OxmlElement('w:tblPr'); tbl.insert(0,tblPr)
    tb=OxmlElement('w:tblBorders')
    for side in ['top','left','bottom','right','insideH','insideV']:
        b=OxmlElement(f'w:{side}')
        if side in ('top','bottom'): b.set(qn('w:val'),'single'); b.set(qn('w:sz'),'6'); b.set(qn('w:color'),'000000')
        else: b.set(qn('w:val'),'none'); b.set(qn('w:sz'),'0'); b.set(qn('w:color'),'auto')
        tb.append(b)
    old=tblPr.find(qn('w:tblBorders'))
    if old is not None: tblPr.remove(old)
    tblPr.append(tb)
    # Header row bottom border
    for cell in table.rows[0].cells:
        tc=cell._tc; tcPr=tc.find(qn('w:tcPr'))
        if tcPr is None: tcPr=OxmlElement('w:tcPr'); tc.insert(0,tcPr)
        tcB=OxmlElement('w:tcBorders'); b=OxmlElement('w:bottom')
        b.set(qn('w:val'),'single'); b.set(qn('w:sz'),'6'); b.set(qn('w:color'),'000000')
        tcB.append(b)
        old_tb=tcPr.find(qn('w:tcBorders'))
        if old_tb is not None: tcPr.remove(old_tb)
        tcPr.append(tcB)

    ref_para._p.addnext(table._tbl)
    return table

# ── Paragraph classifiers ─────────────────────────────────────────────────────

def classify_paragraph(para):
    sn=(para.style.name or '').lower()
    text=get_full_text(para); tl=text.lower().strip()
    if not tl: return 'empty'
    if 'heading 1' in sn: return 'h1'
    if 'heading 2' in sn: return 'h2'
    if 'heading 3' in sn: return 'h3'
    if 'title' in sn: return 'title'
    if re.match(r'^(abstract|abstrak)\s*$',tl): return 'abstract_label'
    if re.match(r'^(kata kunci|keywords?)\s*:',tl): return 'keywords'
    if re.match(r'^received|^diterima',tl): return 'received_line'
    if tl in SECTIONS_H1: return 'h1'
    if re.match(r'^[A-Z][a-zA-Z\-]+,\s+[A-Z]\.',text): return 'reference_entry'
    if re.match(r'^(table|tabel)\s+\d+',tl): return 'table_title'
    if re.match(r'^(figure|gambar|image)\s+\d+',tl): return 'figure_caption'
    if re.match(r'^email:',tl) or ('@' in text and len(text)<80): return 'email'
    if is_tab_row(text): return 'tab_row'
    all_bold=all(r.bold for r in para.runs if r.text.strip())
    if all_bold and len(text)<80:
        if tl in SECTIONS_H1: return 'h1'
        if tl in SECTIONS_H2: return 'h2'
        if len(text)<50: return 'h2'
    return 'body'

# ── Paragraph formatters ──────────────────────────────────────────────────────

def fmt_title(p):
    """FIX #10: sentence case. Arial 14pt bold."""
    clear_pf(p); p.alignment=WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_after=Pt(0)  # FIX #1
    full_text = get_full_text(p)
    new_text = title_to_sentence_case(full_text)
    for run in p.runs: run.text=''
    if p.runs: r=p.runs[0]; r.text=new_text
    else: r=p.add_run(new_text)
    set_run_font(r,FONT_ARIAL,14,bold=True)

def fmt_author(p):
    """FIX #4: TNR 10pt, not bold."""
    clear_pf(p); p.alignment=WD_ALIGN_PARAGRAPH.LEFT
    apply_font_simple(p,FONT_TNR,10)

def fmt_affil(p):
    clear_pf(p); p.alignment=WD_ALIGN_PARAGRAPH.LEFT
    apply_font_simple(p,FONT_TNR,10)

def fmt_email(p):
    clear_pf(p); p.alignment=WD_ALIGN_PARAGRAPH.LEFT
    apply_font_simple(p,FONT_TNR,10)

def fmt_abstract_label(p):
    clear_pf(p); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before=Pt(0)  # FIX #1
    apply_font_simple(p,FONT_TNR,10,bold=True)

def fmt_abstract_body_id(p):
    """
    FIX #3: Abstrak Indonesia = ITALIC (not tegak).
    FIX #3 of prev batch: indent kiri-kanan 0.9cm.
    """
    clear_pf(p); p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.first_line_indent=Cm(0)
    p.paragraph_format.left_indent=Cm(0.9)
    p.paragraph_format.right_indent=Cm(0.9)
    apply_font_simple(p,FONT_TNR,10,italic=True)   # FIX #3: italic

def fmt_abstract_body_en(p):
    """
    FIX #4: Abstrak Inggris = TEGAK (not italic).
    Indent kiri-kanan 0.9cm.
    """
    clear_pf(p); p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.first_line_indent=Cm(0)
    p.paragraph_format.left_indent=Cm(0.9)
    p.paragraph_format.right_indent=Cm(0.9)
    apply_font_simple(p,FONT_TNR,10,italic=False)  # FIX #4: tegak

def fmt_keywords(p):
    clear_pf(p); p.alignment=WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.left_indent=Cm(0.9)
    p.paragraph_format.right_indent=Cm(0.9)
    apply_font_simple(p,FONT_TNR,10,bold=True,italic=True)

def fmt_h1(p):
    clear_pf(p); p.alignment=WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before=Pt(10)
    apply_font_simple(p,FONT_ARIAL,12,bold=True)

def fmt_h2(p):
    clear_pf(p); p.alignment=WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before=Pt(10)
    apply_font_simple(p,FONT_ARIAL,10,bold=True)

def fmt_h3(p):
    clear_pf(p); p.alignment=WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before=Pt(10)
    apply_font_simple(p,FONT_ARIAL,10,italic=True)

def fmt_body(p):
    """Body text with auto sub/superscript detection (FIX #8 & #11)."""
    clear_pf(p); p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.first_line_indent=Cm(0.9)
    rebuild_runs_with_sub_super(p,FONT_TNR,10)

def fmt_table_title(p):
    clear_pf(p); p.alignment=WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before=Pt(10)
    apply_font_simple(p,FONT_TNR,10,bold=True)

def fmt_fig_caption(p):
    clear_pf(p); p.alignment=WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_after=Pt(0)
    apply_font_simple(p,FONT_TNR,10,italic=True)

def fmt_ref_entry(para, segments):
    clear_pf(para); para.alignment=WD_ALIGN_PARAGRAPH.LEFT
    para.paragraph_format.left_indent=Cm(0.9); para.paragraph_format.first_line_indent=Cm(-0.9)
    if segments:
        for r in para.runs: r.text=''
        first=True
        for text,is_italic in segments:
            if first and para.runs: r=para.runs[0]; r.text=text; first=False
            else: r=para.add_run(text)
            set_run_font(r,FONT_TNR,10,italic=is_italic)
    else:
        apply_font_simple(para,FONT_TNR,10)

def fmt_table_existing(table):
    """Format existing Word tables. FIX #7: no bold in header."""
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                p.alignment=WD_ALIGN_PARAGRAPH.LEFT
                p.paragraph_format.space_before=Pt(0); p.paragraph_format.space_after=Pt(0)
                # FIX #7: NOT bold even for header row
                apply_font_simple(p,FONT_TNR,10,bold=False)
    tbl=table._tbl; tblPr=tbl.find(qn('w:tblPr'))
    if tblPr is None: tblPr=OxmlElement('w:tblPr'); tbl.insert(0,tblPr)
    tb=OxmlElement('w:tblBorders')
    for side in ['top','left','bottom','right','insideH','insideV']:
        b=OxmlElement(f'w:{side}')
        if side in ('top','bottom'): b.set(qn('w:val'),'single'); b.set(qn('w:sz'),'6'); b.set(qn('w:color'),'000000')
        else: b.set(qn('w:val'),'none'); b.set(qn('w:sz'),'0'); b.set(qn('w:color'),'auto')
        tb.append(b)
    old=tblPr.find(qn('w:tblBorders'))
    if old is not None: tblPr.remove(old)
    tblPr.append(tb)
    if len(table.rows)>1:
        for cell in table.rows[0].cells:
            tc=cell._tc; tcPr=tc.find(qn('w:tcPr'))
            if tcPr is None: tcPr=OxmlElement('w:tcPr'); tc.insert(0,tcPr)
            tcB=OxmlElement('w:tcBorders'); b=OxmlElement('w:bottom')
            b.set(qn('w:val'),'single'); b.set(qn('w:sz'),'6'); b.set(qn('w:color'),'000000')
            tcB.append(b)
            old_tb=tcPr.find(qn('w:tcBorders'))
            if old_tb is not None: tcPr.remove(old_tb)
            tcPr.append(tcB)

# ── Header / Footer ──────────────────────────────────────────────────────────

def make_run_rpr(fname, fsize, bold=False, italic=False):
    rpr=OxmlElement('w:rPr')
    rFonts=OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'),fname); rFonts.set(qn('w:hAnsi'),fname); rFonts.set(qn('w:cs'),fname)
    rpr.append(rFonts)
    sz=OxmlElement('w:sz'); sz.set(qn('w:val'),str(int(fsize*2))); rpr.append(sz)
    szCs=OxmlElement('w:szCs'); szCs.set(qn('w:val'),str(int(fsize*2))); rpr.append(szCs)
    if bold: rpr.append(OxmlElement('w:b')); rpr.append(OxmlElement('w:bCs'))
    if italic: rpr.append(OxmlElement('w:i')); rpr.append(OxmlElement('w:iCs'))
    return rpr

def add_page_number_run(para, suffix='', fname=FONT_TNR, fsize=10, italic=True):
    p=para._p; NS='{http://www.w3.org/XML/1998/namespace}'
    for tag in ['begin','PAGE_INSTR','separate','VALUE','end']:
        r=OxmlElement('w:r'); r.append(make_run_rpr(fname,fsize,italic=italic))
        if tag=='PAGE_INSTR':
            it=OxmlElement('w:instrText'); it.text='PAGE'; it.set(f'{NS}space','preserve'); r.append(it)
        elif tag=='VALUE':
            t=OxmlElement('w:t'); t.text='1'; r.append(t)
        else:
            fc=OxmlElement('w:fldChar'); fc.set(qn('w:fldCharType'),tag); r.append(fc)
        p.append(r)
    if suffix:
        r=OxmlElement('w:r'); r.append(make_run_rpr(fname,fsize,italic=italic))
        t=OxmlElement('w:t'); t.text=suffix; t.set(f'{NS}space','preserve'); r.append(t); p.append(r)

def add_page_number_start_field(section, start_num):
    """FIX #9: Set page number to start at page_start value."""
    pgNumType = OxmlElement('w:pgNumType')
    pgNumType.set(qn('w:start'), str(start_num))
    sectPr = section._sectPr
    old = sectPr.find(qn('w:pgNumType'))
    if old is not None: sectPr.remove(old)
    sectPr.append(pgNumType)

def hf_add_para(container, text='', fname=FONT_TNR, fsize=10, bold=False, italic=False,
                align=WD_ALIGN_PARAGRAPH.LEFT):
    p=container.add_paragraph()
    p.alignment=align
    p.paragraph_format.space_before=Pt(0); p.paragraph_format.space_after=Pt(0)
    if text:
        run=p.add_run(text); set_run_font(run,fname,fsize,bold=bold,italic=italic)
    return p

def build_headers_footers(doc, meta):
    vol=meta.get('vol',''); no=meta.get('no',''); year=meta.get('year','')
    page_start=meta.get('page_start',''); page_end=meta.get('page_end','')
    doi=meta.get('doi','')
    authors_cite=meta.get('authors_cite','')
    title_cite=meta.get('title_cite','')

    section=doc.sections[0]
    section.different_first_page_header_footer=True

    # FIX #9: page numbering starts at page_start
    try:
        start_num = int(page_start) if page_start else 1
        add_page_number_start_field(section, start_num)
    except: pass

    # ── First page header: RIGHT aligned (FIX #12) ──
    fph=section.first_page_header
    for p in fph.paragraphs: p.clear()
    hf_add_para(fph,'Jurnal Informasi dan Komunikasi Administrasi Perkantoran',
                bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT)  # FIX #12
    hf_add_para(fph,f'Vol. {vol}, No. {no}, {year}',
                align=WD_ALIGN_PARAGRAPH.RIGHT)
    hf_add_para(fph,f'Hlm. {page_start}',
                align=WD_ALIGN_PARAGRAPH.RIGHT)

    # ── Default header (page 2+): {PAGE} – Jurnal..., year, vol(no). ──
    dh=section.header
    for p in dh.paragraphs: p.clear()
    hp=dh.paragraphs[0]
    hp.alignment=WD_ALIGN_PARAGRAPH.LEFT
    hp.paragraph_format.space_before=Pt(0); hp.paragraph_format.space_after=Pt(0)
    suffix=f'  \u2013  Jurnal Informasi dan Komunikasi Administrasi Perkantoran, {year}, {vol}({no}).'
    add_page_number_run(hp,suffix=suffix,fname=FONT_TNR,fsize=10,italic=True)

    # ── First page footer ──
    fpf=section.first_page_footer
    for p in fpf.paragraphs: p.clear()
    hf_add_para(fpf,'____________________',italic=True)
    hf_add_para(fpf,'* Corresponding author',italic=True)
    p_cite=fpf.add_paragraph()
    p_cite.paragraph_format.space_before=Pt(0); p_cite.paragraph_format.space_after=Pt(0)
    r1=p_cite.add_run('Citation in APA style'); set_run_font(r1,FONT_TNR,10,bold=True)
    r2=p_cite.add_run(f': {authors_cite} ({year}). {title_cite}. '); set_run_font(r2,FONT_TNR,10)
    r3=p_cite.add_run(f'Jurnal Informasi dan Komunikasi Administrasi Perkantoran, {vol}({no}), {page_start}-{page_end}.')
    set_run_font(r3,FONT_TNR,10,italic=True)
    if doi: r4=p_cite.add_run(f' {doi}'); set_run_font(r4,FONT_TNR,10)

    # ── Default footer: empty ──
    df=section.footer
    for p in df.paragraphs: p.clear()

# ── Received / DOI line insertion ─────────────────────────────────────────────

def insert_doi_received(doc, meta):
    """
    FIX #5: Received line indented 0.9cm left and right.
    FIX #6: DOI line indented 0.9cm left and right.
    """
    received=meta.get('received',''); revised=meta.get('revised','')
    accepted=meta.get('accepted',''); published=meta.get('published','')
    doi=meta.get('doi','')
    received_text=f"Received {received}; Revised {revised}; Accepted {accepted}; Published Online {published}"

    insert_after_idx=None
    for i,para in enumerate(doc.paragraphs):
        tl=get_full_text(para).lower()
        if re.match(r'^keywords?\s*:',tl): insert_after_idx=i

    if insert_after_idx is not None:
        ref_para=doc.paragraphs[insert_after_idx]

        def insert_after(ref_p_elem, text, bold=False, italic=False, indent=True):
            new_p=OxmlElement('w:p'); ref_p_elem.addnext(new_p)
            for p in doc.paragraphs:
                if p._p is new_p:
                    p.paragraph_format.space_before=Pt(0); p.paragraph_format.space_after=Pt(0)
                    if indent:
                        p.paragraph_format.left_indent=Cm(0.9)   # FIX #5 & #6
                        p.paragraph_format.right_indent=Cm(0.9)
                    run=p.add_run(text)
                    set_run_font(run,FONT_TNR,10,bold=bold,italic=italic)
                    return p
            return None

        p1=insert_after(ref_para._p, received_text, italic=True)
        if p1 and doi:
            insert_after(p1._p, doi, italic=False)

# ── Main format function ──────────────────────────────────────────────────────

def format_document(input_bytes, meta):
    doc=Document(io.BytesIO(input_bytes))
    set_page_margins(doc)

    # Convert tab-delimited rows to tables first
    paragraphs=list(doc.paragraphs)
    n=len(paragraphs)
    classifications=[classify_paragraph(p) for p in paragraphs]

    tab_groups=[]
    i=0
    while i < n:
        if classifications[i]=='tab_row':
            start=i
            while i < n and classifications[i]=='tab_row': i+=1
            tab_groups.append((start,i))
        else: i+=1

    for (start,end) in reversed(tab_groups):
        rows_data=[]
        for idx in range(start,end):
            text=''.join(r.text for r in paragraphs[idx].runs)
            cells=[c.strip() for c in text.split('\t')]
            rows_data.append(cells)
        anchor=paragraphs[start-1] if start>0 else paragraphs[start]
        insert_table_after(doc,anchor,rows_data)
        for idx in range(start,end):
            p=paragraphs[idx]._p; p.getparent().remove(p)

    # Re-read after table insertion
    paragraphs=list(doc.paragraphs)
    classifications=[classify_paragraph(p) for p in paragraphs]

    state='preamble'; id_ab=False; en_ab=False; ref_active=False

    for i,para in enumerate(paragraphs):
        text=get_full_text(para); tl=text.lower().strip(); cls=classifications[i]

        if tl=='abstrak': state='abstract_id_label'; id_ab=True; en_ab=False
        elif tl=='abstract':
            if state in ('abstract_id_label','abstract_id_body','keywords_id'):
                state='abstract_en_label'; id_ab=False; en_ab=True
            else: state='body'; en_ab=False
        elif re.match(r'^(kata kunci|keywords?)\s*:',tl):
            if id_ab: state='keywords_id'
            elif en_ab: state='keywords_en'; en_ab=False
        elif re.match(r'^received|^diterima',tl):
            state='received'; en_ab=False; id_ab=False
        elif tl in SECTIONS_H1:
            if 'reference' in tl or 'pustaka' in tl: ref_active=True; state='references'
            else: state='body'; ref_active=False
        elif ref_active and re.match(r'^[A-Z][a-zA-Z\-]+,?\s+[A-Z]',text):
            cls='reference_entry'; classifications[i]=cls

        # FIX #2: empty paragraphs always TNR 10pt, space 0
        if cls=='empty':
            pf=para.paragraph_format
            pf.space_before=Pt(0); pf.space_after=Pt(0)
            for r in para.runs: set_run_font(r,FONT_TNR,10)
            if not para.runs:
                r=para.add_run(' '); set_run_font(r,FONT_TNR,10)
            continue

        if cls=='tab_row': continue

        if state=='abstract_id_label' and tl=='abstrak': fmt_abstract_label(para)
        elif state=='abstract_en_label' and tl=='abstract': fmt_abstract_label(para)
        elif id_ab and tl!='abstrak' and cls not in ('keywords',):
            state='abstract_id_body'; fmt_abstract_body_id(para)   # FIX #3
        elif en_ab and tl!='abstract' and cls not in ('keywords',):
            state='abstract_en_body'; fmt_abstract_body_en(para)   # FIX #4
        elif state=='keywords_id': fmt_keywords(para)
        elif state=='keywords_en': fmt_keywords(para)
        elif state=='received': para.clear(); state='body'
        elif cls=='h1': fmt_h1(para)
        elif cls=='h2': fmt_h2(para)
        elif cls=='h3': fmt_h3(para)
        elif cls=='table_title': fmt_table_title(para)
        elif cls=='figure_caption': fmt_fig_caption(para)
        elif cls=='reference_entry' or (ref_active and re.match(r'^[A-Z]',text) and len(text)>30):
            fmt_ref_entry(para,format_reference_apa7(parse_reference(text)))
        elif cls=='title' or (i<4 and len(text)>20): fmt_title(para)   # FIX #10
        elif cls=='author': fmt_author(para)
        elif cls=='affiliation': fmt_affil(para)
        elif cls=='email': fmt_email(para)
        else: fmt_body(para)   # FIX #8 & #11 applied inside fmt_body

    for table in doc.tables: fmt_table_existing(table)   # FIX #7

    insert_doi_received(doc,meta)      # FIX #5 & #6
    build_headers_footers(doc,meta)    # FIX #9 & #12

    output=io.BytesIO(); doc.save(output); output.seek(0)
    return output.getvalue()

# ── Streamlit UI ──────────────────────────────────────────────────────────────

st.set_page_config(page_title="JIKAP Manuscript Formatter", page_icon="📄", layout="centered")
st.title("📄 JIKAP Manuscript Formatter")
st.markdown("*Jurnal Informasi dan Komunikasi Administrasi Perkantoran · PAP FKIP UNS*")
st.divider()

st.subheader("1. Upload Manuscript")
uploaded_file=st.file_uploader("Pilih file manuscript (.docx)",type=["docx"])
if uploaded_file:
    st.success(f"✅ **{uploaded_file.name}** ({uploaded_file.size/1024:.1f} KB)")

st.subheader("2. Metadata Jurnal")
st.caption("Isi sesuai informasi edisi jurnal untuk artikel ini.")
col1,col2,col3=st.columns(3)
with col1: vol=st.text_input("Volume",placeholder="mis. 9")
with col2: no=st.text_input("Nomor",placeholder="mis. 2")
with col3: year=st.text_input("Tahun",placeholder="mis. 2025")
col4,col5,col6=st.columns(3)
with col4: page_start=st.text_input("Halaman awal",placeholder="mis. 146")
with col5: page_end=st.text_input("Halaman akhir",placeholder="mis. 154")
with col6: doi=st.text_input("DOI",placeholder="mis. https://dx.doi.org/...")
st.caption("Citation (untuk footer halaman pertama)")
authors_cite=st.text_input("Nama penulis (format APA)",placeholder="mis. Febriana, D.P., & Sawiji, H.")
title_cite=st.text_input("Judul artikel",placeholder="mis. Self-Efficacy and Interpersonal Communication...")
st.caption("Tanggal (format bebas, mis. March 05, 2025)")
col7,col8,col9,col10=st.columns(4)
with col7: received=st.text_input("Received",placeholder="March 05, 2025")
with col8: revised=st.text_input("Revised",placeholder="March 13, 2025")
with col9: accepted=st.text_input("Accepted",placeholder="March 16, 2025")
with col10: published=st.text_input("Published Online",placeholder="March 02, 2025")

st.subheader("3. Format & Download")
if st.button("▶  Format Manuscript",type="primary",use_container_width=True):
    if not uploaded_file:
        st.warning("⚠️ Upload file manuscript terlebih dahulu.")
    else:
        meta={
            'vol':vol,'no':no,'year':year,
            'page_start':page_start,'page_end':page_end,'doi':doi,
            'authors_cite':authors_cite,'title_cite':title_cite,
            'received':received,'revised':revised,'accepted':accepted,'published':published,
        }
        with st.spinner("Memformat manuscript..."):
            try:
                output_bytes=format_document(uploaded_file.getvalue(),meta)
                st.success("✅ Manuscript berhasil diformat!")
                out_name=uploaded_file.name.replace(".docx","_JIKAP_formatted.docx")
                st.download_button(
                    label="⬇️  Download Manuscript Terformat",
                    data=output_bytes, file_name=out_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True, type="primary"
                )
            except Exception as e:
                st.error(f"❌ Error: {str(e)}")
                st.info("Pastikan file adalah dokumen Word (.docx) yang valid.")

with st.expander("ℹ️ Format yang diterapkan (v4)"):
    st.markdown("""
    | # | Elemen | Format |
    |---|---|---|
    | 1 | Line spacing | Before & after = 0 semua paragraf |
    | 2 | Baris kosong | TNR 10pt |
    | 3 | Abstrak Indonesia | TNR 10pt **italic**, indent kiri-kanan |
    | 4 | Abstrak Inggris | TNR 10pt tegak, indent kiri-kanan |
    | 5 | Received/Revised/Accepted | Indent kiri-kanan 5 spasi |
    | 6 | DOI | Indent kiri-kanan 5 spasi |
    | 7 | Judul kolom tabel | Tidak bold |
    | 8 | Subscript | X1, X2, H1 → angka subscript otomatis |
    | 9 | Penomoran halaman | Mulai dari halaman awal artikel |
    | 10 | Judul artikel | Sentence case (huruf kecil kecuali awal & singkatan) |
    | 11 | Superscript | R2, Y2 → angka superscript otomatis |
    | 12 | Header hal. 1 | Rata kanan |
    """)

st.divider()
st.caption("JIKAP Manuscript Formatter v4 · PAP FKIP Universitas Sebelas Maret · 2025")
