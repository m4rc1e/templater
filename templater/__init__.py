# -*- coding: utf-8 -*-

from __future__ import print_function, unicode_literals, with_statement

'''
Word Templates:

Write a template in Word and populate it with tabular data from Excel.

For each Row in the Excel file, a new Word document is created.


Templating language documentation:
----------------------------------
A simple templating language exists to pair the Word template keys to Excel's
column keys.


A key in the Word template is defined by placing the name of the Excel column
between two curly braces, e.g:

{{ Name }} -> Name

Pairing is case sensitive so '{{ Name }}' will not match an Excel column
called 'name'.

Dates:
There is a special key {{ date yyyy-mm-dd }}. This key will produce the
current date, e.g:

{{ date yyyy-mm-dd }} -> 2016-12-12

The user is free to change the yyyy-mm-dd to any combination to suite their
locatilty. '-' can also be replaced by '/'.
e.g:

{{ date dd/mm/yyyy }} -> 2016/12/12



Complete Example:
-----------------

from templater import WordTemplate

w = WordTemplate('word-file.docx', 'excel_file.xlsx')
w.export_multiple_files('output.docx')


____________________________________________________________
Excel file:

      Forename   Surname    Amount Due
    0     Marc     Foley    100
    1      Sam     Smith    230
____________________________________________________________

____________________________________________________________
Word Template:

Hello {{ Forename }} {{ Surname }},

Your invoice for {{ date yyyy-mm-dd }} is ${{ Amount Due }}.

Kind regards,
Tim
____________________________________________________________

____________________________________________________________
Output: (same for next Excel row but with Sam's row data)

Hello Marc Foley,

Your invoice for 2016-12-12 is $100.

Kind regards,
Tim
____________________________________________________________

'''


__author__ = 'Marc Foley'
__version__ = '0.1'

from .docxml import Document
from . import components
from .templater import WordTemplate
