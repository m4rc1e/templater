# -*- coding: utf-8 -*-

from __future__ import print_function, unicode_literals, with_statement

'''
Word Templates:

Write a template in Word and populate it with tabular data from Excel.

For each Row in the Excel file, a new Word document is created.

'''


__author__ = 'Marc Foley'
__version__ = '0.1'

from .docxml import Document
from . import components
from .templates import WordTemplate
