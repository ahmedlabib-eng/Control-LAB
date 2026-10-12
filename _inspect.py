import zipfile, re, os
from lxml import etree

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
WP = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'

path = '0_GetFam added.docx'
z = zipfile.ZipFile(path)
rels = etree.fromstring(z.read('word/_rels/document.xml.rels'))
rid2target = {}
for rel in rels:
    rid2target[rel.get('Id')] = rel.get('Target')

doc = etree.fromstring(z.read('word/document.xml'))
body = doc.find(W + 'body')

idx = 0
for el in body:
    tag = etree.QName(el).localname
    if tag == 'p':
        texts = ''.join(t.text or '' for t in el.iter(W + 't'))
        imgs = []
        for d in el.iter():
            ln = etree.QName(d).localname
            if ln == 'blip':
                rid = d.get(R + 'embed')
                imgs.append(('INLINE' if False else 'blip', rid2target.get(rid)))
        # detect anchor vs inline
        kinds = []
        for d in el.iter():
            if etree.QName(d).localname == 'drawing':
                has_anchor = any(etree.QName(x).localname == 'anchor' for x in d.iter())
                kinds.append('ANCHOR' if has_anchor else 'INLINE')
        print(f'P{idx:03d} imgs={imgs} kinds={kinds} text={texts!r}')
        idx += 1
    else:
        print(f'--- {tag} ---')
