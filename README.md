# Word Templater:

Write a template in Word and populate it with tabular data from Excel.

## Features:
* Generate Word documents from an Excel spreadsheet's rows
* Export documents into one .docx file or seperate .docx files for each row
* Template formatting is preserved

## Installation

    pip install git+https://github.com/m4rc1e/templater

## Templating Language:

A key in the Word template is defined by placing the name of the Excel column
between two curly braces, e.g:

    {{ Name }} ->  Name 


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



## Example:
```
from wordtemplater import WordTemplate

w = WordTemplate('word-file.docx', 'excel_file.xlsx')
w.export_single_file('output.docx')
```


### Excel file:

| Forename | Surname | Amount Due |
| -------- | ------- | ---------- |
| Ben      | Jones   | 100        |
| Sam      | Smith   | 230        |



### Word Template:
---
Hello {{ Forename }} {{ Surname }},

Your invoice for {{ date yyyy-mm-dd }} is **${{ Amount Due }}**.

Regards,
Tim

### Output .docx
---
Hello Ben Jones,

Your invoice for 2016-12-12 is **$100**.

Regards,
Tim


# Testing

```
py.test tests/test.py
```