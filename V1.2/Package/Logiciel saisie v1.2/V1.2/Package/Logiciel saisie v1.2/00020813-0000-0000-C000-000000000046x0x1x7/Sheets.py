# -*- coding: mbcs -*-
# Created by makepy.py version 0.5.01
# By python version 3.3.1 (v3.3.1:d9893d13c628, Apr  6 2013, 20:25:12) [MSC v.1600 32 bit (Intel)]
# From type library '{00020813-0000-0000-C000-000000000046}'
# On Sat Apr  5 18:19:55 2014
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

from win32com.client import DispatchBaseClass
class Sheets(DispatchBaseClass):
	CLSID = IID('{000208D7-0000-0000-C000-000000000046}')
	coclass_clsid = None

	def Add(self, Before=defaultNamedOptArg, After=defaultNamedOptArg, Count=defaultNamedOptArg, Type=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(181, LCID, 1, (9, 0), ((12, 17), (12, 17), (12, 17), (12, 17)),Before
			, After, Count, Type)
		if ret is not None:
			ret = Dispatch(ret, 'Add', None)
		return ret

	def Copy(self, Before=defaultNamedOptArg, After=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(551, LCID, 1, (24, 0), ((12, 17), (12, 17)),Before
			, After)

	def Delete(self):
		return self._oleobj_.InvokeTypes(117, LCID, 1, (24, 0), (),)

	def FillAcrossSheets(self, Range=defaultNamedNotOptArg, Type=-4104):
		return self._oleobj_.InvokeTypes(469, LCID, 1, (24, 0), ((9, 1), (3, 49)),Range
			, Type)

	# The method Item is actually a property, but must be used as a method to correctly pass the arguments
	def Item(self, Index=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(170, LCID, 2, (9, 0), ((12, 1),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'Item', None)
		return ret

	def Move(self, Before=defaultNamedOptArg, After=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(637, LCID, 1, (24, 0), ((12, 17), (12, 17)),Before
			, After)

	def PrintOut(self, From=defaultNamedOptArg, To=defaultNamedOptArg, Copies=defaultNamedOptArg, Preview=defaultNamedOptArg
			, ActivePrinter=defaultNamedOptArg, PrintToFile=defaultNamedOptArg, Collate=defaultNamedOptArg, PrToFileName=defaultNamedOptArg, IgnorePrintAreas=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2361, LCID, 1, (24, 0), ((12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)),From
			, To, Copies, Preview, ActivePrinter, PrintToFile
			, Collate, PrToFileName, IgnorePrintAreas)

	def PrintPreview(self, EnableChanges=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(281, LCID, 1, (24, 0), ((12, 17),),EnableChanges
			)

	def Select(self, Replace=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(235, LCID, 1, (24, 0), ((12, 17),),Replace
			)

	# The method _Default is actually a property, but must be used as a method to correctly pass the arguments
	def _Default(self, Index=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(0, LCID, 2, (9, 0), ((12, 1),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, '_Default', None)
		return ret

	def _PrintOut(self, From=defaultNamedOptArg, To=defaultNamedOptArg, Copies=defaultNamedOptArg, Preview=defaultNamedOptArg
			, ActivePrinter=defaultNamedOptArg, PrintToFile=defaultNamedOptArg, Collate=defaultNamedOptArg, PrToFileName=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1772, LCID, 1, (24, 0), ((12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)),From
			, To, Copies, Preview, ActivePrinter, PrintToFile
			, Collate, PrToFileName)

	def _PrintOut_(self, From=defaultNamedOptArg, To=defaultNamedOptArg, Copies=defaultNamedOptArg, Preview=defaultNamedOptArg
			, ActivePrinter=defaultNamedOptArg, PrintToFile=defaultNamedOptArg, Collate=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(905, LCID, 1, (24, 0), ((12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)),From
			, To, Copies, Preview, ActivePrinter, PrintToFile
			, Collate)

	_prop_map_get_ = {
		# Method 'Application' returns object of type 'Application'
		"Application": (148, 2, (13, 0), (), "Application", '{00024500-0000-0000-C000-000000000046}'),
		"Count": (118, 2, (3, 0), (), "Count", None),
		"Creator": (149, 2, (3, 0), (), "Creator", None),
		# Method 'HPageBreaks' returns object of type 'HPageBreaks'
		"HPageBreaks": (1418, 2, (9, 0), (), "HPageBreaks", '{00024404-0000-0000-C000-000000000046}'),
		"Parent": (150, 2, (9, 0), (), "Parent", None),
		# Method 'VPageBreaks' returns object of type 'VPageBreaks'
		"VPageBreaks": (1419, 2, (9, 0), (), "VPageBreaks", '{00024405-0000-0000-C000-000000000046}'),
		"Visible": (558, 2, (12, 0), (), "Visible", None),
	}
	_prop_map_put_ = {
		"Visible": ((558, LCID, 4, 0),()),
	}
	# Default method for this class is '_Default'
	def __call__(self, Index=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(0, LCID, 2, (9, 0), ((12, 1),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, '__call__', None)
		return ret

	def __str__(self, *args):
		return str(self.__call__(*args))
	def __int__(self, *args):
		return int(self.__call__(*args))
	def __iter__(self):
		"Return a Python iterator for this object"
		try:
			ob = self._oleobj_.InvokeTypes(-4,LCID,2,(13, 10),())
		except pythoncom.error:
			raise TypeError("This object does not support enumeration")
		return win32com.client.util.Iterator(ob, None)
	#This class has Item property/method which allows indexed access with the object[key] syntax.
	#Some objects will accept a string or other type of key in addition to integers.
	#Note that many Office objects do not use zero-based indexing.
	def __getitem__(self, key):
		return self._get_good_object_(self._oleobj_.Invoke(*(170, LCID, 2, 1, key)), "Item", None)
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		return self._ApplyTypes_(*(118, 2, (3, 0), (), "Count", None))
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

win32com.client.CLSIDToClass.RegisterCLSID( "{000208D7-0000-0000-C000-000000000046}", Sheets )
# -*- coding: mbcs -*-
# Created by makepy.py version 0.5.01
# By python version 3.3.1 (v3.3.1:d9893d13c628, Apr  6 2013, 20:25:12) [MSC v.1600 32 bit (Intel)]
# From type library '{00020813-0000-0000-C000-000000000046}'
# On Sat Apr  5 18:19:55 2014
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

Sheets_vtables_dispatch_ = 1
Sheets_vtables_ = [
	(( 'Application' , 'RHS' , ), 148, (148, (), [ (16397, 10, None, "IID('{00024500-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 28 , (3, 0, None, None) , 0 , )),
	(( 'Creator' , 'RHS' , ), 149, (149, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 32 , (3, 0, None, None) , 0 , )),
	(( 'Parent' , 'RHS' , ), 150, (150, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 36 , (3, 0, None, None) , 0 , )),
	(( 'Add' , 'Before' , 'After' , 'Count' , 'Type' , 
			 'lcid' , 'RHS' , ), 181, (181, (), [ (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 4 , 40 , (3, 0, None, None) , 0 , )),
	(( 'Copy' , 'Before' , 'After' , 'lcid' , ), 551, (551, (), [ 
			 (12, 17, None, None) , (12, 17, None, None) , (3, 5, None, None) , ], 1 , 1 , 4 , 2 , 44 , (3, 0, None, None) , 0 , )),
	(( 'Count' , 'RHS' , ), 118, (118, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 48 , (3, 0, None, None) , 0 , )),
	(( 'Delete' , 'lcid' , ), 117, (117, (), [ (3, 5, None, None) , ], 1 , 1 , 4 , 0 , 52 , (3, 0, None, None) , 0 , )),
	(( 'FillAcrossSheets' , 'Range' , 'Type' , 'lcid' , ), 469, (469, (), [ 
			 (9, 1, None, "IID('{00020846-0000-0000-C000-000000000046}')") , (3, 49, '-4104', None) , (3, 5, None, None) , ], 1 , 1 , 4 , 0 , 56 , (3, 0, None, None) , 0 , )),
	(( 'Item' , 'Index' , 'RHS' , ), 170, (170, (), [ (12, 1, None, None) , 
			 (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 60 , (3, 0, None, None) , 0 , )),
	(( 'Move' , 'Before' , 'After' , 'lcid' , ), 637, (637, (), [ 
			 (12, 17, None, None) , (12, 17, None, None) , (3, 5, None, None) , ], 1 , 1 , 4 , 2 , 64 , (3, 0, None, None) , 0 , )),
	(( '_NewEnum' , 'RHS' , ), -4, (-4, (), [ (16397, 10, None, None) , ], 1 , 2 , 4 , 0 , 68 , (3, 0, None, None) , 1024 , )),
	(( '__PrintOut' , 'From' , 'To' , 'Copies' , 'Preview' , 
			 'ActivePrinter' , 'PrintToFile' , 'Collate' , 'lcid' , ), 905, (905, (), [ 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (3, 5, None, None) , ], 1 , 1 , 4 , 7 , 72 , (3, 0, None, None) , 1088 , )),
	(( 'PrintPreview' , 'EnableChanges' , 'lcid' , ), 281, (281, (), [ (12, 17, None, None) , 
			 (3, 5, None, None) , ], 1 , 1 , 4 , 1 , 76 , (3, 0, None, None) , 0 , )),
	(( 'Select' , 'Replace' , 'lcid' , ), 235, (235, (), [ (12, 17, None, None) , 
			 (3, 5, None, None) , ], 1 , 1 , 4 , 1 , 80 , (3, 0, None, None) , 0 , )),
	(( 'HPageBreaks' , 'RHS' , ), 1418, (1418, (), [ (16393, 10, None, "IID('{00024404-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 84 , (3, 0, None, None) , 0 , )),
	(( 'VPageBreaks' , 'RHS' , ), 1419, (1419, (), [ (16393, 10, None, "IID('{00024405-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 88 , (3, 0, None, None) , 0 , )),
	(( 'Visible' , 'lcid' , 'RHS' , ), 558, (558, (), [ (3, 5, None, None) , 
			 (16396, 10, None, None) , ], 1 , 2 , 4 , 0 , 92 , (3, 0, None, None) , 0 , )),
	(( 'Visible' , 'lcid' , 'RHS' , ), 558, (558, (), [ (3, 5, None, None) , 
			 (12, 1, None, None) , ], 1 , 4 , 4 , 0 , 96 , (3, 0, None, None) , 0 , )),
	(( '_Default' , 'Index' , 'RHS' , ), 0, (0, (), [ (12, 1, None, None) , 
			 (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 100 , (3, 0, None, None) , 1024 , )),
	(( '_PrintOut' , 'From' , 'To' , 'Copies' , 'Preview' , 
			 'ActivePrinter' , 'PrintToFile' , 'Collate' , 'PrToFileName' , 'lcid' , 
			 ), 1772, (1772, (), [ (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (3, 5, None, None) , ], 1 , 1 , 4 , 8 , 104 , (3, 0, None, None) , 1088 , )),
	(( 'PrintOut' , 'From' , 'To' , 'Copies' , 'Preview' , 
			 'ActivePrinter' , 'PrintToFile' , 'Collate' , 'PrToFileName' , 'IgnorePrintAreas' , 
			 'lcid' , ), 2361, (2361, (), [ (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (3, 5, None, None) , ], 1 , 1 , 4 , 9 , 108 , (3, 0, None, None) , 0 , )),
]

win32com.client.CLSIDToClass.RegisterCLSID( "{000208D7-0000-0000-C000-000000000046}", Sheets )
