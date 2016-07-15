import shutil
import os
from zipfile import ZipFile
import re
from os.path import basename


class Document:
    '''Primitive docx parser.'''

    def __init__(self, file_path=None):
        with ZipFile(file_path) as self._docx:
            self._files = {name: self._docx.read(name) for name in self._docx.namelist()}

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
