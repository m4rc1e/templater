from docx import Document
import pandas as pd
import re
from copy import copy
from datetime import datetime


def _date(timestamp):
    today = datetime.now()
    date_format = timestamp.split('date ')[-1]
    date = date_format.replace('yyyy', str(today.year))
    date = date.replace('mm', str(today.month))
    date = date.replace('dd', str(today.day))
    return date


def _tmpl_key_2_xl_key(matchobj):
    """convert the Word template key '{{ key }}' to the Excel key
    'key'"""
    # Remove '"{{ " and " }}""
    key = matchobj.group(0)[3:-3]
    if key in XL_KEYS:
        return str(XL_KEYS[key])
    elif 'date' in key:
        return _date(key)
    else:
        return 'KEY MISSING'


def populate_template(template):
    tmpl_key = re.compile(r'\{\{ [a-zA-Z\t \-\_]* \}\}')
    for paragraph in template.paragraphs:
        inline = paragraph.runs
        for i in range(len(inline)):
            inline[i].text = re.sub(tmpl_key, _tmpl_key_2_xl_key, inline[i].text)
    return template


table = pd.read_excel('example1.xlsx')

for row in range(len(table)):
    document_template = Document('example1.docx')
    doc = copy(document_template)
    XL_KEYS = dict(table.iloc[row].to_dict())
    p = populate_template(doc)
    p.save('foo' + str(row) + '.docx')
