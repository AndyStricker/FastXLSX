#
# Copyright (c) 2011 Andreas Stricker <andy@knitter.ch>
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in
# all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.
#

"""
Fast XLSX reader package. It use an event based parser to
keep memory footprint low.
"""

from fastxlsx import reader
from fastxlsx import csvconverter

__major__ = 0  # for major interface/format changes
__minor__ = 2  # for minor interface/format changes
__release__ = 0  # for tweaks, bug-fixes, or development

__version__ = '%d.%d.%d' % (__major__, __minor__, __release__)

__author__ = 'Andreas Stricker'
__license__ = 'BSD'
__author_email__ = 'andy@knitter.ch'
__maintainer_email__ = 'andy@knitter.ch'
__url__ = 'http://github.com/AndyStricker/FastXLSX'
__downloadUrl__ = "http://github.com/AndyStricker/FastXSLX/downloads"

__all__ = ('reader', 'csvconverter',)
