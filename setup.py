import os
import re
from setuptools import find_packages, setup


def text_of(relpath):
    """
    Return string containing the contents of the file at *relpath* relative to
    this file.
    """
    thisdir = os.path.dirname(__file__)
    file_path = os.path.join(thisdir, os.path.normpath(relpath))
    with open(file_path, encoding='utf-8') as f:
        text = f.read()
    return text

# Read the version from docx.__version__ without importing the package
# (and thus attempting to import packages it depends on that may not be
# installed yet)
version = re.search(
    "__version__ = '([^']+)'", text_of('mswordtree/__init__.py')
).group(1)


NAME = 'mswordtree'
VERSION = version
DESCRIPTION = 'Get the parsed microsoft word document in a hierarchical tree structure.'
KEYWORDS = 'docx office openxml word tree microsoft headings tables'
AUTHOR = 'Ali Asad'
AUTHOR_EMAIL = 'imaliasad@outlook.com'
URL = 'https://github.com/imAliAsad/mswordtree'
LICENSE = 'MIT License' #text_of('LICENSE')
PACKAGES = find_packages()
INSTALL_REQUIRES = ['pandas', 'python-docx', 'uuid']
LONG_DESCRIPTION = text_of('README.md')
LONG_DESCRIPTION_Content_Type = 'text/markdown'

CLASSIFIERS = [ 'Development Status :: 3 - Alpha', 'Environment :: Console', 'Intended Audience :: Developers', 'License :: OSI Approved :: MIT License', 'Operating System :: OS Independent', 'Programming Language :: Python', 'Programming Language :: Python :: 2', 'Programming Language :: Python :: 2.6', 'Programming Language :: Python :: 2.7', 'Programming Language :: Python :: 3', 'Programming Language :: Python :: 3.3', 'Programming Language :: Python :: 3.4', 'Topic :: Office/Business :: Office Suites', 'Topic :: Software Development :: Libraries' ]
params = {
    'name':             NAME,
    'version':          VERSION,
    'description':      DESCRIPTION,
    'keywords':         KEYWORDS,
    'long_description': LONG_DESCRIPTION,
    'long_description_content_type': LONG_DESCRIPTION_Content_Type,
    'author':           AUTHOR,
    'author_email':     AUTHOR_EMAIL,
    'url':              URL,
    'license':          LICENSE,
    'packages':         PACKAGES,    
    'install_requires': INSTALL_REQUIRES,
    'classifiers':      CLASSIFIERS,
}


setup(**params)