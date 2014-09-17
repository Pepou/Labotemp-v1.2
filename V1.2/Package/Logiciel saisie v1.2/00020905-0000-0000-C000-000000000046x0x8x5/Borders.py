# -*- coding: mbcs -*-
# Created by makepy.py version 0.5.01
# By python version 3.3.1 (v3.3.1:d9893d13c628, Apr  6 2013, 20:25:12) [MSC v.1600 32 bit (Intel)]
# From type library '{00020905-0000-0000-C000-000000000046}'
# On Fri Mar  7 15:00:53 2014
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
class Borders(DispatchBaseClass):
	CLSID = IID('{0002093C-0000-0000-C000-000000000046}')
	coclass_clsid = None

	def ApplyPageBordersToAllSections(self):
		return self._oleobj_.InvokeTypes(2000, LCID, 1, (24, 0), (),)

	# Result is of type Border
	def Item(self, Index=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(0, LCID, 1, (9, 0), ((3, 1),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'Item', '{0002093B-0000-0000-C000-000000000046}')
		return ret

	_prop_map_get_ = {
		"AlwaysInFront": (23, 2, (11, 0), (), "AlwaysInFront", None),
		# Method 'Application' returns object of type 'Application'
		"Application": (1000, 2, (13, 0), (), "Application", '{000209FF-0000-0000-C000-000000000046}'),
		"Count": (1, 2, (3, 0), (), "Count", None),
		"Creator": (1001, 2, (3, 0), (), "Creator", None),
		"DistanceFrom": (29, 2, (3, 0), (), "DistanceFrom", None),
		"DistanceFromBottom": (21, 2, (3, 0), (), "DistanceFromBottom", None),
		"DistanceFromLeft": (20, 2, (3, 0), (), "DistanceFromLeft", None),
		"DistanceFromRight": (22, 2, (3, 0), (), "DistanceFromRight", None),
		"DistanceFromTop": (4, 2, (3, 0), (), "DistanceFromTop", None),
		"Enable": (2, 2, (3, 0), (), "Enable", None),
		"EnableFirstPageInSection": (30, 2, (11, 0), (), "EnableFirstPageInSection", None),
		"EnableOtherPagesInSection": (31, 2, (11, 0), (), "EnableOtherPagesInSection", None),
		"HasHorizontal": (27, 2, (11, 0), (), "HasHorizontal", None),
		"HasVertical": (28, 2, (11, 0), (), "HasVertical", None),
		"InsideColor": (32, 2, (3, 0), (), "InsideColor", None),
		"InsideColorIndex": (10, 2, (3, 0), (), "InsideColorIndex", None),
		"InsideLineStyle": (6, 2, (3, 0), (), "InsideLineStyle", None),
		"InsideLineWidth": (8, 2, (3, 0), (), "InsideLineWidth", None),
		"JoinBorders": (26, 2, (11, 0), (), "JoinBorders", None),
		"OutsideColor": (33, 2, (3, 0), (), "OutsideColor", None),
		"OutsideColorIndex": (11, 2, (3, 0), (), "OutsideColorIndex", None),
		"OutsideLineStyle": (7, 2, (3, 0), (), "OutsideLineStyle", None),
		"OutsideLineWidth": (9, 2, (3, 0), (), "OutsideLineWidth", None),
		"Parent": (1002, 2, (9, 0), (), "Parent", None),
		"Shadow": (5, 2, (11, 0), (), "Shadow", None),
		"SurroundFooter": (25, 2, (11, 0), (), "SurroundFooter", None),
		"SurroundHeader": (24, 2, (11, 0), (), "SurroundHeader", None),
	}
	_prop_map_put_ = {
		"AlwaysInFront": ((23, LCID, 4, 0),()),
		"DistanceFrom": ((29, LCID, 4, 0),()),
		"DistanceFromBottom": ((21, LCID, 4, 0),()),
		"DistanceFromLeft": ((20, LCID, 4, 0),()),
		"DistanceFromRight": ((22, LCID, 4, 0),()),
		"DistanceFromTop": ((4, LCID, 4, 0),()),
		"Enable": ((2, LCID, 4, 0),()),
		"EnableFirstPageInSection": ((30, LCID, 4, 0),()),
		"EnableOtherPagesInSection": ((31, LCID, 4, 0),()),
		"InsideColor": ((32, LCID, 4, 0),()),
		"InsideColorIndex": ((10, LCID, 4, 0),()),
		"InsideLineStyle": ((6, LCID, 4, 0),()),
		"InsideLineWidth": ((8, LCID, 4, 0),()),
		"JoinBorders": ((26, LCID, 4, 0),()),
		"OutsideColor": ((33, LCID, 4, 0),()),
		"OutsideColorIndex": ((11, LCID, 4, 0),()),
		"OutsideLineStyle": ((7, LCID, 4, 0),()),
		"OutsideLineWidth": ((9, LCID, 4, 0),()),
		"Shadow": ((5, LCID, 4, 0),()),
		"SurroundFooter": ((25, LCID, 4, 0),()),
		"SurroundHeader": ((24, LCID, 4, 0),()),
	}
	# Default method for this class is 'Item'
	def __call__(self, Index=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(0, LCID, 1, (9, 0), ((3, 1),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, '__call__', '{0002093B-0000-0000-C000-000000000046}')
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
		return win32com.client.util.Iterator(ob, '{0002093B-0000-0000-C000-000000000046}')
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		return self._ApplyTypes_(*(1, 2, (3, 0), (), "Count", None))
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

win32com.client.CLSIDToClass.RegisterCLSID( "{0002093C-0000-0000-C000-000000000046}", Borders )
# -*- coding: mbcs -*-
# Created by makepy.py version 0.5.01
# By python version 3.3.1 (v3.3.1:d9893d13c628, Apr  6 2013, 20:25:12) [MSC v.1600 32 bit (Intel)]
# From type library '{00020905-0000-0000-C000-000000000046}'
# On Fri Mar  7 15:00:53 2014
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

Borders_vtables_dispatch_ = 1
Borders_vtables_ = [
	(( 'Application' , 'prop' , ), 1000, (1000, (), [ (16397, 10, None, "IID('{000209FF-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 28 , (3, 0, None, None) , 0 , )),
	(( 'Creator' , 'prop' , ), 1001, (1001, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 32 , (3, 0, None, None) , 0 , )),
	(( 'Parent' , 'prop' , ), 1002, (1002, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 36 , (3, 0, None, None) , 0 , )),
	(( '_NewEnum' , 'prop' , ), -4, (-4, (), [ (16397, 10, None, None) , ], 1 , 2 , 4 , 0 , 40 , (3, 0, None, None) , 1024 , )),
	(( 'Count' , 'prop' , ), 1, (1, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 44 , (3, 0, None, None) , 0 , )),
	(( 'Enable' , 'prop' , ), 2, (2, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 48 , (3, 0, None, None) , 0 , )),
	(( 'Enable' , 'prop' , ), 2, (2, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 52 , (3, 0, None, None) , 0 , )),
	(( 'DistanceFromTop' , 'prop' , ), 4, (4, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 56 , (3, 0, None, None) , 0 , )),
	(( 'DistanceFromTop' , 'prop' , ), 4, (4, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 60 , (3, 0, None, None) , 0 , )),
	(( 'Shadow' , 'prop' , ), 5, (5, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 64 , (3, 0, None, None) , 0 , )),
	(( 'Shadow' , 'prop' , ), 5, (5, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 68 , (3, 0, None, None) , 0 , )),
	(( 'InsideLineStyle' , 'prop' , ), 6, (6, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 72 , (3, 0, None, None) , 0 , )),
	(( 'InsideLineStyle' , 'prop' , ), 6, (6, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 76 , (3, 0, None, None) , 0 , )),
	(( 'OutsideLineStyle' , 'prop' , ), 7, (7, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 80 , (3, 0, None, None) , 0 , )),
	(( 'OutsideLineStyle' , 'prop' , ), 7, (7, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 84 , (3, 0, None, None) , 0 , )),
	(( 'InsideLineWidth' , 'prop' , ), 8, (8, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 88 , (3, 0, None, None) , 0 , )),
	(( 'InsideLineWidth' , 'prop' , ), 8, (8, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 92 , (3, 0, None, None) , 0 , )),
	(( 'OutsideLineWidth' , 'prop' , ), 9, (9, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 96 , (3, 0, None, None) , 0 , )),
	(( 'OutsideLineWidth' , 'prop' , ), 9, (9, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 100 , (3, 0, None, None) , 0 , )),
	(( 'InsideColorIndex' , 'prop' , ), 10, (10, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 104 , (3, 0, None, None) , 0 , )),
	(( 'InsideColorIndex' , 'prop' , ), 10, (10, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 108 , (3, 0, None, None) , 0 , )),
	(( 'OutsideColorIndex' , 'prop' , ), 11, (11, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 112 , (3, 0, None, None) , 0 , )),
	(( 'OutsideColorIndex' , 'prop' , ), 11, (11, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 116 , (3, 0, None, None) , 0 , )),
	(( 'DistanceFromLeft' , 'prop' , ), 20, (20, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 120 , (3, 0, None, None) , 0 , )),
	(( 'DistanceFromLeft' , 'prop' , ), 20, (20, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 124 , (3, 0, None, None) , 0 , )),
	(( 'DistanceFromBottom' , 'prop' , ), 21, (21, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 128 , (3, 0, None, None) , 0 , )),
	(( 'DistanceFromBottom' , 'prop' , ), 21, (21, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 132 , (3, 0, None, None) , 0 , )),
	(( 'DistanceFromRight' , 'prop' , ), 22, (22, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 136 , (3, 0, None, None) , 0 , )),
	(( 'DistanceFromRight' , 'prop' , ), 22, (22, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 140 , (3, 0, None, None) , 0 , )),
	(( 'AlwaysInFront' , 'prop' , ), 23, (23, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 144 , (3, 0, None, None) , 0 , )),
	(( 'AlwaysInFront' , 'prop' , ), 23, (23, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 148 , (3, 0, None, None) , 0 , )),
	(( 'SurroundHeader' , 'prop' , ), 24, (24, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 152 , (3, 0, None, None) , 0 , )),
	(( 'SurroundHeader' , 'prop' , ), 24, (24, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 156 , (3, 0, None, None) , 0 , )),
	(( 'SurroundFooter' , 'prop' , ), 25, (25, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 160 , (3, 0, None, None) , 0 , )),
	(( 'SurroundFooter' , 'prop' , ), 25, (25, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 164 , (3, 0, None, None) , 0 , )),
	(( 'JoinBorders' , 'prop' , ), 26, (26, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 168 , (3, 0, None, None) , 0 , )),
	(( 'JoinBorders' , 'prop' , ), 26, (26, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 172 , (3, 0, None, None) , 0 , )),
	(( 'HasHorizontal' , 'prop' , ), 27, (27, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 176 , (3, 0, None, None) , 0 , )),
	(( 'HasVertical' , 'prop' , ), 28, (28, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 180 , (3, 0, None, None) , 0 , )),
	(( 'DistanceFrom' , 'prop' , ), 29, (29, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 184 , (3, 0, None, None) , 0 , )),
	(( 'DistanceFrom' , 'prop' , ), 29, (29, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 188 , (3, 0, None, None) , 0 , )),
	(( 'EnableFirstPageInSection' , 'prop' , ), 30, (30, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 192 , (3, 0, None, None) , 0 , )),
	(( 'EnableFirstPageInSection' , 'prop' , ), 30, (30, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 196 , (3, 0, None, None) , 0 , )),
	(( 'EnableOtherPagesInSection' , 'prop' , ), 31, (31, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 200 , (3, 0, None, None) , 0 , )),
	(( 'EnableOtherPagesInSection' , 'prop' , ), 31, (31, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 204 , (3, 0, None, None) , 0 , )),
	(( 'Item' , 'Index' , 'prop' , ), 0, (0, (), [ (3, 1, None, None) , 
			 (16393, 10, None, "IID('{0002093B-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 0 , 208 , (3, 0, None, None) , 0 , )),
	(( 'ApplyPageBordersToAllSections' , ), 2000, (2000, (), [ ], 1 , 1 , 4 , 0 , 212 , (3, 0, None, None) , 0 , )),
	(( 'InsideColor' , 'prop' , ), 32, (32, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 216 , (3, 0, None, None) , 0 , )),
	(( 'InsideColor' , 'prop' , ), 32, (32, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 220 , (3, 0, None, None) , 0 , )),
	(( 'OutsideColor' , 'prop' , ), 33, (33, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 224 , (3, 0, None, None) , 0 , )),
	(( 'OutsideColor' , 'prop' , ), 33, (33, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 228 , (3, 0, None, None) , 0 , )),
]

win32com.client.CLSIDToClass.RegisterCLSID( "{0002093C-0000-0000-C000-000000000046}", Borders )
