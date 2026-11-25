"""Add a real Word footnote to the first occurrence of a term in a DOCX.

The input is a freshly generated document without an existing footnotes part.
The reference follows the term, keeping the rest of the run intact.
"""
from copy import deepcopy
from pathlib import Path
from tempfile import NamedTemporaryFile
from zipfile import ZipFile, ZIP_DEFLATED
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
Q = lambda ns, name: f'{{{ns}}}{name}'


def add_footnote(path, term, definition):
    path = Path(path)
    with ZipFile(path) as source:
        parts = {item.filename: source.read(item.filename) for item in source.infolist()}
    if 'word/footnotes.xml' in parts:
        raise ValueError('Document already has footnotes; refusing to overwrite')
    doc = etree.fromstring(parts['word/document.xml'])
    matches = doc.xpath('//w:body//w:t[contains(text(), $term)]', namespaces={'w': W}, term=term)
    if not matches:
        raise ValueError(f'Footnote term not found in body: {term}')
    text = matches[0]
    run = text.getparent()
    if run.tag != Q(W, 'r') or len(run.findall(Q(W, 't'))) != 1:
        raise ValueError('Footnote term must be in a simple run')
    before, after = text.text.split(term, 1)
    text.text = before + term
    reference = etree.Element(Q(W, 'r'))
    props = etree.SubElement(reference, Q(W, 'rPr'))
    etree.SubElement(props, Q(W, 'rStyle')).set(Q(W, 'val'), 'FootnoteReference')
    etree.SubElement(props, Q(W, 'vertAlign')).set(Q(W, 'val'), 'superscript')
    etree.SubElement(reference, Q(W, 'footnoteReference')).set(Q(W, 'id'), '1')
    run.addnext(reference)
    if after:
        tail = deepcopy(run)
        tail.find(Q(W, 't')).text = after
        reference.addnext(tail)
    parts['word/document.xml'] = etree.tostring(doc, xml_declaration=True, encoding='UTF-8', standalone=True)

    footnotes = etree.Element(Q(W, 'footnotes'), nsmap={'w': W})
    for i, typ, tag in [(-1, 'separator', 'separator'), (0, 'continuationSeparator', 'continuationSeparator')]:
        n = etree.SubElement(footnotes, Q(W, 'footnote'))
        n.set(Q(W, 'type'), typ); n.set(Q(W, 'id'), str(i))
        etree.SubElement(etree.SubElement(n, Q(W, 'p')), Q(W, 'r')).append(etree.Element(Q(W, tag)))
    n = etree.SubElement(footnotes, Q(W, 'footnote')); n.set(Q(W, 'id'), '1')
    p = etree.SubElement(n, Q(W, 'p')); ppr = etree.SubElement(p, Q(W, 'pPr'))
    etree.SubElement(ppr, Q(W, 'pStyle')).set(Q(W, 'val'), 'FootnoteText')
    etree.SubElement(ppr, Q(W, 'bidi')).set(Q(W, 'val'), '1')
    etree.SubElement(ppr, Q(W, 'jc')).set(Q(W, 'val'), 'right')
    mark = etree.SubElement(p, Q(W, 'r')); markpr = etree.SubElement(mark, Q(W, 'rPr'))
    etree.SubElement(markpr, Q(W, 'rStyle')).set(Q(W, 'val'), 'FootnoteReference')
    etree.SubElement(markpr, Q(W, 'vertAlign')).set(Q(W, 'val'), 'superscript')
    etree.SubElement(mark, Q(W, 'footnoteRef'))
    r = etree.SubElement(p, Q(W, 'r')); rp = etree.SubElement(r, Q(W, 'rPr'))
    fonts = etree.SubElement(rp, Q(W, 'rFonts'))
    for key, value in [('ascii','B Nazanin'),('hAnsi','B Nazanin'),('cs','B Zar')]: fonts.set(Q(W, key), value)
    etree.SubElement(rp, Q(W, 'sz')).set(Q(W, 'val'), '20')
    etree.SubElement(rp, Q(W, 'szCs')).set(Q(W, 'val'), '20')
    etree.SubElement(rp, Q(W, 'rtl')).set(Q(W, 'val'), '1')
    t = etree.SubElement(r, Q(W, 't')); t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve'); t.text = ' ' + definition
    parts['word/footnotes.xml'] = etree.tostring(footnotes, xml_declaration=True, encoding='UTF-8', standalone=True)

    rels = etree.fromstring(parts['word/_rels/document.xml.rels'])
    ids = {e.get('Id') for e in rels}; idx = 1
    while f'rId{idx}' in ids: idx += 1
    rel = etree.SubElement(rels, Q(R, 'Relationship'))
    rel.set('Id', f'rId{idx}'); rel.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes'); rel.set('Target', 'footnotes.xml')
    parts['word/_rels/document.xml.rels'] = etree.tostring(rels, xml_declaration=True, encoding='UTF-8', standalone=True)
    types = etree.fromstring(parts['[Content_Types].xml'])
    node = etree.SubElement(types, Q(CT, 'Override'))
    node.set('PartName', '/word/footnotes.xml'); node.set('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml')
    parts['[Content_Types].xml'] = etree.tostring(types, xml_declaration=True, encoding='UTF-8', standalone=True)
    with NamedTemporaryFile(dir=path.parent, suffix='.docx', delete=False) as tmp: temp = Path(tmp.name)
    try:
        with ZipFile(temp, 'w', ZIP_DEFLATED) as output:
            for name, data in parts.items(): output.writestr(name, data)
        temp.replace(path)
    finally:
        temp.unlink(missing_ok=True)
