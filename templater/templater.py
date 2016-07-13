from docx import Document
import pandas as pd
import re
from copy import copy
from datetime import datetime


class WordTemplate(object):

    def __init__(self, template_file, excel_file):
        self.template_file = template_file
        self.excel_file = pd.read_excel(excel_file)
        self._xl_keys = {}
        self._docs = []

        self._get_docs()

    def _get_docs(self):
        for row in range(len(self.excel_file)):
            document_template = Document(self.template_file)
            doc = copy(document_template)
            self._xl_keys = dict(self.excel_file.iloc[row].to_dict())
            p = self._populate_template(doc)
            self._docs.append(p)

    def _populate_template(self, template):
        tmpl_key = re.compile(r'\{\{ [a-zA-Z\t \-\_]* \}\}')
        for paragraph in template.paragraphs:
            inline = paragraph.runs
            for i in range(len(inline)):
                inline[i].text = re.sub(tmpl_key, self._tmpl_key_2_xl_key,
                                        inline[i].text)
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

    def _date(self, timestamp):
        today = datetime.now()
        date_format = timestamp.split('date ')[-1]
        date = date_format.replace('yyyy', str(today.year))
        date = date.replace('mm', str(today.month))
        date = date.replace('dd', str(today.day))
        return date

    def export_single_file(self, file_output):
        pass

    def export_multiple_files(self):
        count = 0
        for doc in self._docs:
            doc.save('example_' + str(count) + '.docx')
            count += 1

    @property
    def docs(self):
        '''Return the document list.'''
        return self._docs
