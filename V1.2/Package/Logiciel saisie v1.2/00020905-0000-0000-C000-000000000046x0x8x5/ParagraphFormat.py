# -*- coding: mbcs -*-
# Created by makepy.py version 0.5.01
# By python version 3.3.1 (v3.3.1:d9893d13c628, Apr  6 2013, 20:25:12) [MSC v.1600 32 bit (Intel)]
# From type library '{00020905-0000-0000-C000-000000000046}'
# On Fri Mar  7 14:57:46 2014
'Microsoft Word 14.0 Object Library'
makepy_version = '0.5.01'
python_version = 0x30301f0

import win32com.client.CLSIDToClass, pythoncom, pywintypes
import win32com.client.util
from pywintypes import IID
from win32com.client import Dispatch

# The following 3 lines may need tweaking for the particular server
# Candidates are pythoncom.Missing, .Empty and .ArgNotFound
defaultNamedOptArg=pythoncom.Empty
defaultNamedNotOptArg=pythoncom.Empty
defaultUnnamedArg=pythoncom.Empty

CLSID = IID('{00020905-0000-0000-C000-000000000046}')
MajorVersion = 8
MinorVersion = 5
LibraryFlags = 8
LCID = 0x0

from win32com.client import CoClassBaseClass
import sys
__import__('win32com.gen_py.00020905-0000-0000-C000-000000000046x0x8x5._ParagraphFormat')
_ParagraphFormat = sys.modules['win32com.gen_py.00020905-0000-0000-C000-000000000046x0x8x5._ParagraphFormat']._ParagraphFormat
class ParagraphFormat(CoClassBaseClass): # A CoClass
	CLSID = IID('{000209F4-0000-0000-C000-000000000046}')
	coclass_sources = [
	]
	coclass_interfaces = [
		_ParagraphFormat,
	]
	default_interface = _ParagraphFormat

win32com.client.CLSIDToClass.RegisterCLSID( "{000209F4-0000-0000-C000-000000000046}", ParagraphFormat )
