# -*- coding: utf-8 -*-

import pandas as pd
import re
from copy import deepcopy
from datetime import datetime
from lxml import etree
from .docxml import Document
from . import components


class WordTemplate(object):
    '''Write either a single Word file of multiple word files from
    a Word template and Excel file.'''
    def __init__(self, template_file, excel_file):
        self.template_file = template_file
        self.excel_file = pd.read_excel(excel_file)
        self._xl_keys = {}
        self._docs = []

        self._get_docs()

    def _get_docs(self):
        '''Create a Word document for each row in the Excel file.'''
        for row in range(len(self.excel_file)):
            doc = Document(self.template_file)
            self._xl_keys = dict(self.excel_file.iloc[row].to_dict())
            doc = self._replace_template_keys(doc)
            self._docs.append(doc)

    def _replace_template_keys(self, template):
        '''Find a template tag and replace it with the tags key'''
        tmpl_key = re.compile(r'\{\{ [a-zA-Z\t \-\_]* \}\}')  # {{ key }}
        # {{ key }} -> key
        template.replace_text(tmpl_key, self._tmpl_key_2_xl_key)
        return template

    def _tmpl_key_2_xl_key(self, matchobj):
        """convert the Word template key '{{ key }}' to the Excel key
        'key'"""
        excel_key = matchobj.group(0)[3:-3]  # Remove '"{{ " and " }}""
        if excel_key in self._xl_keys:
            return str(self._xl_keys[excel_key])
        elif 'date' in excel_key:
            return self._date(excel_key)
        else:
            return 'KEY MISSING'

    def _date(self, d_format):
        '''Convert date template tags to locale based date. e.g:

        {{ date yyyy/mm/dd }} => 2016/12/12
        {{ date dd-mm-yyyy }} => 12-12-2016
        '''
        today = datetime.now()
        date_format = d_format.split('date ')[-1]
        date = date_format.replace('yyyy', str(today.year))
        date = date.replace('mm', str(today.month))
        date = date.replace('dd', str(today.day))
        return date

    def export_single_file(self, filename):
        '''Collate all files into one docx. A pagebreak is inserted
        between each doc'''
        master = deepcopy(self.docs[0])
        doc_xml = etree.XML(master.files['word/document.xml'])

        for doc in self.docs[1:]:
            #  Insert a pagebreak for each document
            doc_xml[0].append(deepcopy(components.page_break))
            xml = etree.XML(doc.files['word/document.xml'])
            for element in xml[0]:
                # Append inside the <document> tag
                doc_xml[0].append(element)
        master.files['word/document.xml'] = etree.tostring(doc_xml)
        master.save(filename)

    def export_multiple_files(self, file_output):
        '''Export each file in self._docs. Filenames are numerically
        increased.'''
        for i, doc in enumerate(self.docs):
            filename = file_output[:-4] + str(i) + '.docx'
            doc.save(filename)

    @property
    def docs(self):
        '''Return all the documents spawned from template and Excel file'''
        return self._docs
