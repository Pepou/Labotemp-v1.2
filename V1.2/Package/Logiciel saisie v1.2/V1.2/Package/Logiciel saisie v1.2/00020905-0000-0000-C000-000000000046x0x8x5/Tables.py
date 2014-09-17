# -*- coding: mbcs -*-
# Created by makepy.py version 0.5.01
# By python version 3.3.1 (v3.3.1:d9893d13c628, Apr  6 2013, 20:25:12) [MSC v.1600 32 bit (Intel)]
# From type library '{00020905-0000-0000-C000-000000000046}'
# On Thu Mar  6 21:58:03 2014
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

from win32com.client import DispatchBaseClass
class Tables(DispatchBaseClass):
	CLSID = IID('{0002094D-0000-0000-C000-000000000046}')
	coclass_clsid = None

	# Result is of type Table
	def Add(self, Range=defaultNamedNotOptArg, NumRows=defaultNamedNotOptArg, NumColumns=defaultNamedNotOptArg, DefaultTableBehavior=defaultNamedOptArg
			, AutoFitBehavior=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(200, LCID, 1, (9, 0), ((9, 1), (3, 1), (3, 1), (16396, 17), (16396, 17)),Range
			, NumRows, NumColumns, DefaultTableBehavior, AutoFitBehavior)
		if ret is not None:
			ret = Dispatch(ret, 'Add', '{00020951-0000-0000-C000-000000000046}')
		return ret

	# Result is of type Table
	def AddOld(self, Range=defaultNamedNotOptArg, NumRows=defaultNamedNotOptArg, NumColumns=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(4, LCID, 1, (9, 0), ((9, 1), (3, 1), (3, 1)),Range
			, NumRows, NumColumns)
		if ret is not None:
			ret = Dispatch(ret, 'AddOld', '{00020951-0000-0000-C000-000000000046}')
		return ret

	# Result is of type Table
	def Item(self, Index=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(0, LCID, 1, (9, 0), ((3, 1),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'Item', '{00020951-0000-0000-C000-000000000046}')
		return ret

	_prop_map_get_ = {
		# Method 'Application' returns object of type 'Application'
		"Application": (1000, 2, (13, 0), (), "Application", '{000209FF-0000-0000-C000-000000000046}'),
		"Count": (2, 2, (3, 0), (), "Count", None),
		"Creator": (1001, 2, (3, 0), (), "Creator", None),
		"NestingLevel": (100, 2, (3, 0), (), "NestingLevel", None),
		"Parent": (1002, 2, (9, 0), (), "Parent", None),
	}
	_prop_map_put_ = {
	}
	# Default method for this class is 'Item'
	def __call__(self, Index=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(0, LCID, 1, (9, 0), ((3, 1),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, '__call__', '{00020951-0000-0000-C000-000000000046}')
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
		return win32com.client.util.Iterator(ob, '{00020951-0000-0000-C000-000000000046}')
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		return self._ApplyTypes_(*(2, 2, (3, 0), (), "Count", None))
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

win32com.client.CLSIDToClass.RegisterCLSID( "{0002094D-0000-0000-C000-000000000046}", Tables )
# -*- coding: mbcs -*-
# Created by makepy.py version 0.5.01
# By python version 3.3.1 (v3.3.1:d9893d13c628, Apr  6 2013, 20:25:12) [MSC v.1600 32 bit (Intel)]
# From type library '{00020905-0000-0000-C000-000000000046}'
# On Thu Mar  6 21:58:03 2014
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

Tables_vtables_dispatch_ = 1
Tables_vtables_ = [
	(( '_NewEnum' , 'prop' , ), -4, (-4, (), [ (16397, 10, None, None) , ], 1 , 2 , 4 , 0 , 28 , (3, 0, None, None) , 1024 , )),
	(( 'Count' , 'prop' , ), 2, (2, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 32 , (3, 0, None, None) , 0 , )),
	(( 'Application' , 'prop' , ), 1000, (1000, (), [ (16397, 10, None, "IID('{000209FF-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 36 , (3, 0, None, None) , 0 , )),
	(( 'Creator' , 'prop' , ), 1001, (1001, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 40 , (3, 0, None, None) , 0 , )),
	(( 'Parent' , 'prop' , ), 1002, (1002, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 44 , (3, 0, None, None) , 0 , )),
	(( 'Item' , 'Index' , 'prop' , ), 0, (0, (), [ (3, 1, None, None) , 
			 (16393, 10, None, "IID('{00020951-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 0 , 48 , (3, 0, None, None) , 0 , )),
	(( 'AddOld' , 'Range' , 'NumRows' , 'NumColumns' , 'prop' , 
			 ), 4, (4, (), [ (9, 1, None, "IID('{0002095E-0000-0000-C000-000000000046}')") , (3, 1, None, None) , (3, 1, None, None) , (16393, 10, None, "IID('{00020951-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 0 , 52 , (3, 0, None, None) , 64 , )),
	(( 'Add' , 'Range' , 'NumRows' , 'NumColumns' , 'DefaultTableBehavior' , 
			 'AutoFitBehavior' , 'prop' , ), 200, (200, (), [ (9, 1, None, "IID('{0002095E-0000-0000-C000-000000000046}')") , (3, 1, None, None) , 
			 (3, 1, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16393, 10, None, "IID('{00020951-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 2 , 56 , (3, 0, None, None) , 0 , )),
	(( 'NestingLevel' , 'prop' , ), 100, (100, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 60 , (3, 0, None, None) , 0 , )),
]

win32com.client.CLSIDToClass.RegisterCLSID( "{0002094D-0000-0000-C000-000000000046}", Tables )
