# -*- coding: utf-8 -*-

import pandas as pd
from docx import Document

document = Document('examples/RPT-ETAT-ACTIONNAIRE.docx')

# Get the structure of the docx file text
#text_df = pd.DataFrame(data = [(para.text, para.style.name) for para in document.paragraphs], columns = ("Text", "Style"))    
#print("They are %s different text styles in the document." %text_df["Style"].nunique())

def table_to_html(table):
    """
    Get a docx table, generate a html table
    """
    html_table = "<table class = \"table\">"
    for row in table.rows:
        html_table += "<tr>"
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                html_table += "<td>" + paragraph.text + "</td>"
        html_table += "</tr>"
    html_table += "</table>"
    return(html_table)
                
# Get the structure of the docx file table
#table_df = pd.DataFrame(data = [(table_to_html(table), table.style.name) for table in document.tables], columns = ("Table","Style"))
#print("They are %s different text styles in the document." %table_df["Style"].nunique())

from docx.document import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph

def iter_block_items(parent):
    """
    Yield each paragraph and table child within *parent*, in document order.
    Each returned value is an instance of either Table or Paragraph. *parent*
    would most commonly be a reference to a main Document object, but
    also works for a _Cell object, which itself can contain paragraphs and tables.
    """
    if isinstance(parent, Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise ValueError("something's not right")

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)

def text_or_table(block):
    """
    If the block is a paragraph, return text.
    If the block is a table, return an html table
    """
    try:
        return(block.text)
    except AttributeError:
        return(table_to_html(block))

# Look mom, I can do list comprehension ;)
structure = pd.DataFrame(data = [(text_or_table(block),block.style.name) for block in iter_block_items(document)], columns = ["content", "style"])


#style = pd.DataFrame(data = structure["Style"].unique(), columns = ["Style"])


style_correspondance = {
"Titre général" : "h1"
"Heading 1" : "h1",
"Heading 2" : "h2",
"Heading 3" : "h3",
"Heading 4" : "h4",
"Heading 5" : "h5",
"Heading 6" : "h6",
"Titre de partie" : "h1",
"Titre Annexe" : "h1",
"Titre (Sommaire des réponses)" : "h1",
"Réponse" : "i",
"Puce Reco + Italique" : "i"
"ENCADRÉ - Titre" : 
"ENCADRÉ - Texte"

}

default_tag = "p"

structure["HTML"] = structure["style"].map(style_correspondance).fillna(default_tag)

structure["web content"] = "<div><" + structure["HTML"] + ">" + structure["content"] + "</" + structure["HTML"] + "></div>"



structure["web content"].to_frame().to_csv("index2.html")


