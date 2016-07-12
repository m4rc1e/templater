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




class WordTemplate(object):

    def __init__(self, template, excel_file):
        self.template = template
        self.excel_file = pd.read_excel(excel_file)
        self.XL_KEYS = {}
        self._docs = []

        self._get_docs()

    def get_docs(self):
        for row in range(len(self.excel_file)):
            document_template = Document(self.template)
            doc = copy(document_template)
            XL_KEYS = dict(self.excel_file.iloc[row].to_dict())
            p = self._populate_template(doc)
            self._docs.append(p)

    def export_single_file(self):
        pass

    def export_multiple_files(self):
        pass

    def _date(self, timestamp):
        today = datetime.now()
        date_format = timestamp.split('date ')[-1]
        date = date_format.replace('yyyy', str(today.year))
        date = date.replace('mm', str(today.month))
        date = date.replace('dd', str(today.day))
        return date

    def _tmpl_key_2_xl_key(self, matchobj):
        """convert the Word template key '{{ key }}' to the Excel key
        'key'"""
        # Remove '"{{ " and " }}""
        key = matchobj.group(0)[3:-3]
        if key in self.XL_KEYS:
            return str(self.XL_KEYS[key])
        elif 'date' in key:
            return self._date(key)
        else:
            return 'KEY MISSING'

    def _populate_template(self, template):
        tmpl_key = re.compile(r'\{\{ [a-zA-Z\t \-\_]* \}\}')
        for paragraph in template.paragraphs:
            inline = paragraph.runs
            for i in range(len(inline)):
                inline[i].text = re.sub(tmpl_key, self._tmpl_key_2_xl_key, inline[i].text)
        return template
