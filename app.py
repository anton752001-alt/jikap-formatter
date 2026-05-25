"""
JIKAP Manuscript Formatter - Streamlit Web App
Jurnal Informasi dan Komunikasi Administrasi Perkantoran
PAP FKIP Universitas Sebelas Maret
"""

import streamlit as st
import io, re
from docx import Document
from docx.shared import Pt, Cm, RGBColor
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
        ic = re.sub(r'([A-Z])(?!\.)(?!\s*[A-Z]\.)' ,r'\1.',inits)
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

# ── Font helpers ──────────────────────────────────────────────────────────────

def set_page_margins(doc):
    for section in doc.sections:
        section.page_width=Cm(21.0); section.page_height=Cm(29.7)
        section.top_margin=Cm(3.0); section.bottom_margin=Cm(3.0)
        section.left_margin=Cm(3.0); section.right_margin=Cm(3.0)

def clear_pf(para):
    pf=para.paragraph_format
    pf.space_before=Pt(0); pf.space_after=Pt(0)
    pf.first_line_indent=None; pf.left_indent=None; pf.right_indent=None

def set_run_font(run, fname, fsize, bold=False, italic=False):
    run.font.name=fname; run.font.size=Pt(fsize)
    run.font.bold=bold; run.font.italic=italic
    rpr=run._r.get_or_add_rPr()
    rFonts=rpr.find(qn('w:rFonts'))
    if rFonts is None: rFonts=OxmlElement('w:rFonts'); rpr.insert(0,rFonts)
    rFonts.set(qn('w:ascii'),fname); rFonts.set(qn('w:hAnsi'),fname); rFonts.set(qn('w:cs'),fname)

def get_full_text(para): return ''.join(r.text for r in para.runs).strip()

# ── Paragraph formatters ──────────────────────────────────────────────────────

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
    all_bold=all(r.bold for r in para.runs if r.text.strip())
    if all_bold and len(text)<80:
        if tl in SECTIONS_H1: return 'h1'
        if tl in SECTIONS_H2: return 'h2'
        if len(text)<50: return 'h2'
    return 'body'

def fmt_title(p):
    clear_pf(p); p.alignment=WD_ALIGN_PARAGRAPH.LEFT; p.paragraph_format.space_after=Pt(20)
    for r in p.runs: set_run_font(r,FONT_ARIAL,14,bold=True)
def fmt_author(p):
    clear_pf(p); p.alignment=WD_ALIGN_PARAGRAPH.LEFT
    for r in p.runs: set_run_font(r,FONT_TNR,10,bold=True)
def fmt_affil(p):
    clear_pf(p); p.alignment=WD_ALIGN_PARAGRAPH.LEFT
    for r in p.runs: set_run_font(r,FONT_TNR,10)
def fmt_email(p):
    clear_pf(p); p.alignment=WD_ALIGN_PARAGRAPH.LEFT; p.paragraph_format.space_after=Pt(10)
    for r in p.runs: set_run_font(r,FONT_TNR,10)
def fmt_abstract_label(p):
    clear_pf(p); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_before=Pt(10)
    for r in p.runs: set_run_font(r,FONT_TNR,10,bold=True)
def fmt_abstract_body(p,en=False):
    clear_pf(p); p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY; p.paragraph_format.first_line_indent=Cm(0)
    for r in p.runs: set_run_font(r,FONT_TNR,10,italic=en)
def fmt_keywords(p):
    clear_pf(p); p.alignment=WD_ALIGN_PARAGRAPH.LEFT
    for r in p.runs: set_run_font(r,FONT_TNR,10,bold=True,italic=True)
def fmt_h1(p):
    clear_pf(p); p.alignment=WD_ALIGN_PARAGRAPH.LEFT; p.paragraph_format.space_before=Pt(10)
    for r in p.runs: set_run_font(r,FONT_ARIAL,12,bold=True)
def fmt_h2(p):
    clear_pf(p); p.alignment=WD_ALIGN_PARAGRAPH.LEFT; p.paragraph_format.space_before=Pt(10)
    for r in p.runs: set_run_font(r,FONT_ARIAL,10,bold=True)
def fmt_h3(p):
    clear_pf(p); p.alignment=WD_ALIGN_PARAGRAPH.LEFT; p.paragraph_format.space_before=Pt(10)
    for r in p.runs: set_run_font(r,FONT_ARIAL,10,italic=True)
def fmt_body(p):
    clear_pf(p); p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY; p.paragraph_format.first_line_indent=Cm(0.9)
    for r in p.runs: set_run_font(r,FONT_TNR,10)
def fmt_table_title(p):
    clear_pf(p); p.alignment=WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before=Pt(10); p.paragraph_format.space_after=Pt(2)
    for r in p.runs: set_run_font(r,FONT_TNR,10,bold=True)
def fmt_fig_caption(p):
    clear_pf(p); p.alignment=WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before=Pt(2); p.paragraph_format.space_after=Pt(10)
    for r in p.runs: set_run_font(r,FONT_TNR,10,italic=True)
def fmt_ref_entry(para,segments):
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
        for r in para.runs: set_run_font(r,FONT_TNR,10)

def fmt_table(table):
    n=len(table.rows)
    for ri,row in enumerate(table.rows):
        for cell in row.cells:
            for p in cell.paragraphs:
                p.alignment=WD_ALIGN_PARAGRAPH.LEFT
                for r in p.runs: set_run_font(r,FONT_TNR,10,bold=(ri==0))
    tbl=table._tbl
    tblPr=tbl.find(qn('w:tblPr'))
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
    if n>1:
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
    if bold: b=OxmlElement('w:b'); rpr.append(b); bCs=OxmlElement('w:bCs'); rpr.append(bCs)
    if italic: i=OxmlElement('w:i'); rpr.append(i); iCs=OxmlElement('w:iCs'); rpr.append(iCs)
    return rpr

def add_page_number_run(para, suffix='', fname=FONT_TNR, fsize=10, italic=True):
    """Add PAGE field + suffix text to existing paragraph."""
    p=para._p
    XML_SPACE='http://www.w3.org/XML/1998/namespace'
    r1=OxmlElement('w:r'); r1.append(make_run_rpr(fname,fsize,italic=italic))
    fc1=OxmlElement('w:fldChar'); fc1.set(qn('w:fldCharType'),'begin'); r1.append(fc1); p.append(r1)
    r2=OxmlElement('w:r'); r2.append(make_run_rpr(fname,fsize,italic=italic))
    it=OxmlElement('w:instrText'); it.text='PAGE'; it.set(f'{{{XML_SPACE}}}space','preserve'); r2.append(it); p.append(r2)
    r3=OxmlElement('w:r'); r3.append(make_run_rpr(fname,fsize,italic=italic))
    fc3=OxmlElement('w:fldChar'); fc3.set(qn('w:fldCharType'),'separate'); r3.append(fc3); p.append(r3)
    r4=OxmlElement('w:r'); r4.append(make_run_rpr(fname,fsize,italic=italic))
    t4=OxmlElement('w:t'); t4.text='1'; r4.append(t4); p.append(r4)
    r5=OxmlElement('w:r'); r5.append(make_run_rpr(fname,fsize,italic=italic))
    fc5=OxmlElement('w:fldChar'); fc5.set(qn('w:fldCharType'),'end'); r5.append(fc5); p.append(r5)
    if suffix:
        r6=OxmlElement('w:r'); r6.append(make_run_rpr(fname,fsize,italic=italic))
        t6=OxmlElement('w:t'); t6.text=suffix; t6.set(f'{{{XML_SPACE}}}space','preserve'); r6.append(t6); p.append(r6)

def hf_add_para(container, text, fname=FONT_TNR, fsize=10, bold=False, italic=False, align=WD_ALIGN_PARAGRAPH.LEFT):
    p=container.add_paragraph()
    p.alignment=align
    p.paragraph_format.space_before=Pt(0); p.paragraph_format.space_after=Pt(0)
    if text:
        run=p.add_run(text)
        set_run_font(run,fname,fsize,bold=bold,italic=italic)
    return p

def build_headers_footers(doc, meta):
    vol=meta.get('vol',''); no=meta.get('no',''); year=meta.get('year','')
    page_start=meta.get('page_start',''); page_end=meta.get('page_end','')
    doi=meta.get('doi','')
    authors_cite=meta.get('authors_cite','')
    title_cite=meta.get('title_cite','')

    section=doc.sections[0]
    section.different_first_page_header_footer=True

    # ── First page header: journal name / vol info / hlm ──
    fph=section.first_page_header
    for p in fph.paragraphs: p.clear()
    hf_add_para(fph,'Jurnal Informasi dan Komunikasi Administrasi Perkantoran',bold=True)
    hf_add_para(fph,f'Vol. {vol}, No. {no}, {year}')
    hf_add_para(fph,f'Hlm. {page_start}')

    # ── Default header (page 2+): {PAGE}  –  Jurnal..., year, vol(no). ──
    dh=section.header
    for p in dh.paragraphs: p.clear()
    hp=dh.paragraphs[0]
    hp.alignment=WD_ALIGN_PARAGRAPH.LEFT
    hp.paragraph_format.space_before=Pt(0); hp.paragraph_format.space_after=Pt(0)
    suffix=f'  \u2013  Jurnal Informasi dan Komunikasi Administrasi Perkantoran, {year}, {vol}({no}).'
    add_page_number_run(hp, suffix=suffix, fname=FONT_TNR, fsize=10, italic=True)

    # ── First page footer: ___ / * Corresponding author / Citation ──
    fpf=section.first_page_footer
    for p in fpf.paragraphs: p.clear()
    hf_add_para(fpf,'____________________',italic=True)
    hf_add_para(fpf,'* Corresponding author',italic=True)
    # Citation paragraph (mixed bold/normal/italic)
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

# ── DOI / Received line insertion ────────────────────────────────────────────

def insert_doi_received(doc, meta):
    received=meta.get('received',''); revised=meta.get('revised','')
    accepted=meta.get('accepted',''); published=meta.get('published','')
    doi=meta.get('doi','')
    received_text=f"Received {received}; Revised {revised}; Accepted {accepted}; Published Online {published}"

    paragraphs=doc.paragraphs
    insert_after_idx=None
    for i,para in enumerate(paragraphs):
        tl=get_full_text(para).lower()
        if re.match(r'^keywords?\s*:',tl): insert_after_idx=i

    if insert_after_idx is not None:
        ref_para=paragraphs[insert_after_idx]
        def insert_after(ref_p_elem, text_segs):
            new_p=OxmlElement('w:p'); ref_p_elem.addnext(new_p)
            for p in doc.paragraphs:
                if p._p is new_p:
                    p.paragraph_format.space_before=Pt(0); p.paragraph_format.space_after=Pt(0)
                    for text,bold,italic in text_segs:
                        run=p.add_run(text); set_run_font(run,FONT_TNR,10,bold=bold,italic=italic)
                    return p
            return None
        p1=insert_after(ref_para._p,[(received_text,False,True)])
        if p1 and doi: insert_after(p1._p,[(doi,False,False)])

# ── Main format function ──────────────────────────────────────────────────────

def format_document(input_bytes, meta):
    doc=Document(io.BytesIO(input_bytes))
    set_page_margins(doc)
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

        if cls=='empty': clear_pf(para); continue

        if state=='abstract_id_label' and tl=='abstrak': fmt_abstract_label(para)
        elif state=='abstract_en_label' and tl=='abstract': fmt_abstract_label(para)
        elif id_ab and tl!='abstrak' and cls not in ('keywords',): state='abstract_id_body'; fmt_abstract_body(para,en=False)
        elif en_ab and tl!='abstract' and cls not in ('keywords',): state='abstract_en_body'; fmt_abstract_body(para,en=True)
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
        elif cls=='title' or (i<4 and len(text)>20): fmt_title(para)
        elif cls=='author': fmt_author(para)
        elif cls=='affiliation': fmt_affil(para)
        elif cls=='email': fmt_email(para)
        else: fmt_body(para)

    for table in doc.tables: fmt_table(table)

    insert_doi_received(doc, meta)
    build_headers_footers(doc, meta)

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
                    data=output_bytes,file_name=out_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,type="primary"
                )
            except Exception as e:
                st.error(f"❌ Error: {str(e)}")
                st.info("Pastikan file adalah dokumen Word (.docx) yang valid.")

with st.expander("ℹ️ Format yang diterapkan"):
    st.markdown("""
    | Elemen | Format JIKAP |
    |---|---|
    | Halaman | A4, margin 3 cm semua sisi |
    | Header hal. 1 | Nama jurnal, Vol/No/Tahun, Hlm |
    | Header hal. 2+ | {No halaman} – Jurnal..., tahun, vol(no) |
    | Footer hal. 1 | Garis / * Corresponding author / Citation APA |
    | Judul artikel | Arial 14pt Bold |
    | Body teks | TNR 10pt, justified, indent 5 spasi |
    | Heading 1 | Arial 12pt Bold |
    | Heading 2 | Arial 10pt Bold |
    | Abstrak Indonesia | TNR 10pt tegak |
    | Abstrak Inggris | TNR 10pt italic |
    | Referensi | TNR 10pt, hanging indent, APA 7 |
    """)

st.divider()
st.caption("JIKAP Manuscript Formatter · PAP FKIP Universitas Sebelas Maret · 2025")
