import os
from setuptools import setup

# Utility function to read the README file.
# Used for the long_description.  It's nice, because now 1) we have a top level
# README file and 2) it's easier to type in the README file than to put a raw
# string in below ...
def read(fname):
    return open(os.path.join(os.path.dirname(__file__), fname)).read()

setup(
    name = "wordtemplater",
    version = "0.0.1",
    author = "Marc Foley",
    author_email = "m.foley.88@gmail.com",
    description = ("Write a template in Word and populate it with \
                    tabular data from Excel."),
    license = "BSD",
    keywords = "Word Excel templating generating",
    packages=['wordtemplater'],
    url='https://github.com/m4rc1e/templater',
    long_description=read('README.md'),
    classifiers=[
        "Development Status :: 3 - Alpha",
        "Topic :: Utilities",
        "License :: OSI Approved :: BSD License",
    ],
    install_requires=[
        'lxml==3.6.0',
        'pandas==0.18.1',
        'pytest==2.9.2',
    ],
)