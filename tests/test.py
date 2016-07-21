# from templater.docxml import Document
# from templater.templater import WordTemplate
from datetime import datetime
import os
import shutil
from wordtemplater import Document, WordTemplate, components
from zipfile import ZipFile


TEMPLATE = './tests/example1.docx'
EXCEL = './tests/example1.xlsx'
EXPORT_LOC = './tests/_exports/'


def test_docs():
    '''Test the WordTemplate returns a collection of Word docs'''
    w = WordTemplate(TEMPLATE, EXCEL)
    assert(len(w.docs)) != 0


def test_date():
    date = datetime.now()
    w = WordTemplate(TEMPLATE, EXCEL)
    assert w._date('date dd-mm-yyyy') == '%s-%s-%s' % (date.day, date.month, date.year)
    #  Change '-' to '/'
    assert w._date('date dd/mm/yyyy') == '%s/%s/%s' % (date.day, date.month, date.year)
    # Change locality order
    assert w._date('date yyyy/mm/dd') == '%s/%s/%s' % (date.year, date.month, date.day)


def test_document_file_dict():
    '''Test .docx has succesfully unzipped it's files into a python dictioanry'''
    doc = Document(TEMPLATE)
    assert 'word/document.xml' in doc.files
    assert 'word/styles.xml' in doc.files


def test_convert_template_key_2_excel_column_name():
    w = WordTemplate(TEMPLATE, EXCEL)
    # Check if the {{ Forename }} key is replaced with the first Name Entry.
    assert 'Marc' in str(w.docs[0].files['word/document.xml'])
    # Check {{ Surname }} key is replaced by the second row
    assert 'Smith' in str(w.docs[1].files['word/document.xml'])


def test_single_doc_export():
    '''Create temporary export folder and output a single document containing all
    templated documents'''
    os.mkdir(EXPORT_LOC)
    w = WordTemplate(TEMPLATE, EXCEL)
    w.export_single_file(EXPORT_LOC + 'master.docx')
    assert 'master.docx' in os.listdir(EXPORT_LOC)
    shutil.rmtree(EXPORT_LOC)


def test_multiple_doc_export():
    '''Create temprary export folder and output each templated document'''
    os.mkdir(EXPORT_LOC)
    w = WordTemplate(TEMPLATE, EXCEL)
    w.export_multiple_files(EXPORT_LOC + 'file.zip')
    
    with ZipFile(EXPORT_LOC + 'file.zip', 'r') as zip_file:
        zip_contents = [f.filename for f in zip_file.filelist]
        assert len(zip_contents) >= 3
    [os.remove(EXPORT_LOC + f) for f in os.listdir(EXPORT_LOC)]
    shutil.rmtree(EXPORT_LOC)
