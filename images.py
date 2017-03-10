#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import docx2txt
import xml.etree.ElementTree as ET
import zipfile
import re

nsmap = {
    'a':   ('http://schemas.openxmlformats.org/drawingml/2006/main'),
    'r':   ('http://schemas.openxmlformats.org/officeDocument/2006/relations'
            'hips'),
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

def get_images(xml):
    """
    Get .xml, return full text with image ids at the right place
    """
    text = u''
    root = ET.fromstring(xml)
    for child in root.iter():
        if child.tag == qn('w:t'):
            t_text = child.text
            text += t_text if t_text is not None else ''
        elif child.tag == qn('w:tab'):
            text += '\t'
        elif child.tag in (qn('w:br'), qn('w:cr')):
            text += '\n'
        elif child.tag == qn("w:p"):
            text += '\n\n'
        elif child.tag == qn("a:blip"):
            i_id = child.get(qn("r:embed"))
            text += "{{" + i_id + "}}"
    return text

def build_image_dict(text):
    """
    Get full text, return a dictionary of image's id and before text
    """
    image_dict = {}
    image_list = re.findall(r'\{\{(.*?)\}}', text)
    for image in image_list:
        text_position = text.find("{{" + image + "}}")
        text_before = text[text_position-200 : text_position-2]
        text_before = re.sub('\{\{[^}]*}}', '', text_before)
        text_before = re.sub('.*?(?=\}})', '', text_before)
        text_before = re.sub('[^a-zA-Z]+', '', text_before)
        image_dict[image] = text_before
    return image_dict

def build_relation_dict(xml):
    """
    Get .xml, return a dictionary of image's id and target 
    """
    relation_dict = {}
    root = ET.fromstring(xml)
    for child in root:
        relation_dict[child.get('Id')] = child.get('Target')
    return relation_dict

def get_images_dicts(file):
    """
    Get a .docx file, return two dictionnaries (and extract the media in a folder)
    """
    # Unzip the docx in memory
    zipf = zipfile.ZipFile(file)
    # TODO: Image extraction shounldn't be in this function
    # Write images in /media
    docx2txt.process(file, "media") 
    # Build image dict
    text = u''
    text = get_images(zipf.read('word/document.xml'))
    image_dict = build_image_dict(text)
    # Build relation dict    
    relation_dict = build_relation_dict(zipf.read('word/_rels/document.xml.rels'))
    return image_dict, relation_dict