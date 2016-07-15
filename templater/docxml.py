'''
Primitive docx parser.


Why not use python-docx?
------------------------
As of 2016-07-15, this library cannot replace text. It is a fantastic library
to create new documents. We could rebuild the document using this library.
However, this will change the data significantly and we may lose formatting
and other features not yet supported. For this reason we decided to work on
the raw xml.
'''
import shutil
import os
from zipfile import ZipFile
import re
from os.path import basename


class Document:
    '''
    Docx file is loaded via ZipFile and its files are stored in a dictionary
    which can be accessed by using Document.files.

    Structure:

    Example1.docx
    ├── [Content_Types].xml
    ├── _rels
    └── word
        ├── _rels
        │   └── document.xml.rels
        ├── document.xml
        ├── fontTable.xml
        ├── numbering.xml
        ├── settings.xml
        └── styles.xml

    Eg: To Access the document's text:
        >>>Document.files['word/document.xml']
    '''

    def __init__(self, file_path=None):
        with ZipFile(file_path) as self._docx:
            self._files = {name: self._docx.read(name) for name
                           in self._docx.namelist()}

    def replace_text(self, current, new):
        text = self._files['word/document.xml'].decode('utf-8')
        text = re.sub(current, new, text)
        self._files['word/document.xml'] = bytes(text, 'utf-8')

    def save(self, path):
        with ZipFile(path, 'w') as doc:
            for file in self._files:
                doc.writestr(file, self._files[file])

    @property
    def files(self):
        return self._files
