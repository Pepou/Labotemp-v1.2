# -*- coding: mbcs -*-
# Created by makepy.py version 0.5.01
# By python version 3.3.1 (v3.3.1:d9893d13c628, Apr  6 2013, 20:25:12) [MSC v.1600 32 bit (Intel)]
# From type library '{00020905-0000-0000-C000-000000000046}'
# On Mon Apr 15 10:41:20 2013
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
__import__('win32com.gen_py.00020905-0000-0000-C000-000000000046x0x8x5.DocumentEvents2')
DocumentEvents2 = sys.modules['win32com.gen_py.00020905-0000-0000-C000-000000000046x0x8x5.DocumentEvents2'].DocumentEvents2
__import__('win32com.gen_py.00020905-0000-0000-C000-000000000046x0x8x5.DocumentEvents')
DocumentEvents = sys.modules['win32com.gen_py.00020905-0000-0000-C000-000000000046x0x8x5.DocumentEvents'].DocumentEvents
__import__('win32com.gen_py.00020905-0000-0000-C000-000000000046x0x8x5._Document')
_Document = sys.modules['win32com.gen_py.00020905-0000-0000-C000-000000000046x0x8x5._Document']._Document
# This CoClass is known by the name 'Word.Document.8'
class Document(CoClassBaseClass): # A CoClass
	CLSID = IID('{00020906-0000-0000-C000-000000000046}')
	coclass_sources = [
		DocumentEvents2,
		DocumentEvents,
	]
	default_source = DocumentEvents2
	coclass_interfaces = [
		_Document,
	]
	default_interface = _Document

win32com.client.CLSIDToClass.RegisterCLSID( "{00020906-0000-0000-C000-000000000046}", Document )
