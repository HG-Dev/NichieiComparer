"""
NichieiComparer
-------

NichieiComparer's methods allow one to compare Japanese text
in a previously translated document to Japanese text in a
yet untranslated document to discover terms and phrases
they have in common, fostering consistency between translations.
"""

__author__ = 'Ian M. <ian.hg.dev@gmail.com>'
__version__ = '0.0.0'

from .utils import FileIO, CollectionUtils, LangUtils, ExcelUtils
from .data_analysis import ExcelDoc
