# -*- coding: mbcs -*-
# Created by makepy.py version 0.5.01
# By python version 3.3.1 (v3.3.1:d9893d13c628, Apr  6 2013, 20:25:12) [MSC v.1600 32 bit (Intel)]
# From type library '{00020813-0000-0000-C000-000000000046}'
# On Sat Apr  5 18:15:38 2014
'Microsoft Excel 14.0 Object Library'
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

CLSID = IID('{00020813-0000-0000-C000-000000000046}')
MajorVersion = 1
MinorVersion = 7
LibraryFlags = 8
LCID = 0x0

from win32com.client import CoClassBaseClass
import sys
__import__('win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x7.DocEvents')
DocEvents = sys.modules['win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x7.DocEvents'].DocEvents
__import__('win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x7._Worksheet')
_Worksheet = sys.modules['win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x7._Worksheet']._Worksheet
# This CoClass is known by the name 'Excel.Sheet.8'
class Worksheet(CoClassBaseClass): # A CoClass
	CLSID = IID('{00020820-0000-0000-C000-000000000046}')
	coclass_sources = [
		DocEvents,
	]
	default_source = DocEvents
	coclass_interfaces = [
		_Worksheet,
	]
	default_interface = _Worksheet

win32com.client.CLSIDToClass.RegisterCLSID( "{00020820-0000-0000-C000-000000000046}", Worksheet )
