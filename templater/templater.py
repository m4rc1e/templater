import pandas as pd
import re
from copy import copy
from datetime import datetime
from .docxml import Document
import lxml

class WordTemplate(object):

    def __init__(self, template_file, excel_file):
        self.template_file = Document(template_file)
        self.excel_file = pd.read_excel(excel_file)
        self._xl_keys = {}
        self._docs = []

        self._get_docs()

    def _get_docs(self):
        for row in range(len(self.excel_file)):
            doc = copy(self.template_file)
            self._xl_keys = dict(self.excel_file.iloc[row].to_dict())
            doc = self._populate_template(doc)
            self._docs.append(doc)

    def _populate_template(self, template):
        '''Find a template tag and replace it with the tags key'''
        tmpl_key = re.compile(r'\{\{ [a-zA-Z\t \-\_]* \}\}')  # {{ key }}
        template.replace_text(tmpl_key, self._tmpl_key_2_xl_key)  # {{ key }} -> key
        return template

    def _tmpl_key_2_xl_key(self, matchobj):
        """convert the Word template key '{{ key }}' to the Excel key
        'key'"""
        # Remove '"{{ " and " }}""
        key = matchobj.group(0)[3:-3]
        if key in self._xl_keys:
            return str(self._xl_keys[key])
        elif 'date' in key:
            return self._date(key)
        else:
            return 'KEY MISSING'

    def _date(self, d_format):
        today = datetime.now()
        date_format = d_format.split('date ')[-1]
        date = date_format.replace('yyyy', str(today.year))
        date = date.replace('mm', str(today.month))
        date = date.replace('dd', str(today.day))
        return date

    def export_single_file(self, file_output):
        pass

    def export_multiple_files(self, file_output):
        for i, doc in enumerate(self._docs):
            filename = file_output[:-4] + str(i) + '.docx'
            doc.save(filename)

    @property
    def docs(self):
        '''Return the document list.'''
        return self._docs
