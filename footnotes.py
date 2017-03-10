#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Note: the file /anaconda/lib/python3.6/site-package/docx/oxml/text/run.py
has also been modified to add footnote references in the text.
The following code was added in the text property decorator:

elif child.tag == qn("w:footnoteReference"):
    f_id = child.get(qn('w:id'))
    text += "<sup><a href=\"#fn" + f_id + "\"  id=\"ref" + f_id + "\">"+ f_id + "</a></sup>"

"""
import xml.etree.ElementTree as ET
import zipfile

nsmap = {
    'w':   ('http://schemas.openxmlformats.org/wordprocessingml/2006/main'),
    'wp':  ('http://schemas.openxmlformats.org/drawingml/2006/wordprocessing'
            'Drawing'),
}

def qn(tag):
    """
    Source: https://github.com/python-openxml/python-docx/
    """
    prefix, tagroot = tag.split(':')
    uri = nsmap[prefix]
    return '{{{}}}{}'.format(uri, tagroot)

def get_footnotes(xml):
    """
    Get a .xml file, return a dictionary of footnote's id and content
    """
    footdict= {}
    root = ET.fromstring(xml)
    for footnote in root.findall(qn('w:footnote')):
        f_id = footnote.get(qn('w:id'))
        text = u""
        for child in footnote.iter():            
            if child.tag == qn('w:t'):
                t_text = child.text
                text += t_text if t_text is not None else ''
            if text is not '': 
                footdict[f_id] = text
    return footdict

def get_html_footnotes(file):
    """
    Get a .docx file, return a list of footnotes in HTML
    """
    # Unzip the docx in memory
    zipf = zipfile.ZipFile(file)
    # Get the footnote dictionary
    footdict = get_footnotes(zipf.read('word/footnotes.xml'))   
    # Generate HTML from the dictionary
    footnotes = ["<p><sup id=\"fn" + key +  "\">" + key + "." + footdict[key] + "<a href=\"#ref" + key +"\" title=\"Retour au texte.\">↩</a></sup></p>"
                for key in footdict]
    return footnotes