# -*- coding: mbcs -*-
# Created by makepy.py version 0.5.01
# By python version 3.3.1 (v3.3.1:d9893d13c628, Apr  6 2013, 20:25:12) [MSC v.1600 32 bit (Intel)]
# From type library '{00020905-0000-0000-C000-000000000046}'
# On Thu Mar  6 22:05:15 2014
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
class Cell(DispatchBaseClass):
	CLSID = IID('{0002094E-0000-0000-C000-000000000046}')
	coclass_clsid = None

	def AutoSum(self):
		return self._oleobj_.InvokeTypes(206, LCID, 1, (24, 0), (),)

	def Delete(self, ShiftCells=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(200, LCID, 1, (24, 0), ((16396, 17),),ShiftCells
			)

	def Formula(self, Formula=defaultNamedOptArg, NumFormat=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(201, LCID, 1, (24, 0), ((16396, 17), (16396, 17)),Formula
			, NumFormat)

	def Merge(self, MergeTo=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(204, LCID, 1, (24, 0), ((9, 1),),MergeTo
			)

	def Select(self):
		return self._oleobj_.InvokeTypes(65535, LCID, 1, (24, 0), (),)

	def SetHeight(self, RowHeight=defaultNamedNotOptArg, HeightRule=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(203, LCID, 1, (24, 0), ((16396, 1), (3, 1)),RowHeight
			, HeightRule)

	def SetWidth(self, ColumnWidth=defaultNamedNotOptArg, RulerStyle=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(202, LCID, 1, (24, 0), ((4, 1), (3, 1)),ColumnWidth
			, RulerStyle)

	def Split(self, NumRows=defaultNamedOptArg, NumColumns=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(205, LCID, 1, (24, 0), ((16396, 17), (16396, 17)),NumRows
			, NumColumns)

	_prop_map_get_ = {
		# Method 'Application' returns object of type 'Application'
		"Application": (1000, 2, (13, 0), (), "Application", '{000209FF-0000-0000-C000-000000000046}'),
		# Method 'Borders' returns object of type 'Borders'
		"Borders": (1100, 2, (9, 0), (), "Borders", '{0002093C-0000-0000-C000-000000000046}'),
		"BottomPadding": (112, 2, (4, 0), (), "BottomPadding", None),
		# Method 'Column' returns object of type 'Column'
		"Column": (101, 2, (9, 0), (), "Column", '{0002094F-0000-0000-C000-000000000046}'),
		"ColumnIndex": (5, 2, (3, 0), (), "ColumnIndex", None),
		"Creator": (1001, 2, (3, 0), (), "Creator", None),
		"FitText": (110, 2, (11, 0), (), "FitText", None),
		"Height": (7, 2, (4, 0), (), "Height", None),
		"HeightRule": (8, 2, (3, 0), (), "HeightRule", None),
		"ID": (115, 2, (8, 0), (), "ID", None),
		"LeftPadding": (113, 2, (4, 0), (), "LeftPadding", None),
		"NestingLevel": (107, 2, (3, 0), (), "NestingLevel", None),
		# Method 'Next' returns object of type 'Cell'
		"Next": (103, 2, (9, 0), (), "Next", '{0002094E-0000-0000-C000-000000000046}'),
		"Parent": (1002, 2, (9, 0), (), "Parent", None),
		"PreferredWidth": (109, 2, (4, 0), (), "PreferredWidth", None),
		"PreferredWidthType": (116, 2, (3, 0), (), "PreferredWidthType", None),
		# Method 'Previous' returns object of type 'Cell'
		"Previous": (104, 2, (9, 0), (), "Previous", '{0002094E-0000-0000-C000-000000000046}'),
		# Method 'Range' returns object of type 'Range'
		"Range": (0, 2, (9, 0), (), "Range", '{0002095E-0000-0000-C000-000000000046}'),
		"RightPadding": (114, 2, (4, 0), (), "RightPadding", None),
		# Method 'Row' returns object of type 'Row'
		"Row": (102, 2, (9, 0), (), "Row", '{00020950-0000-0000-C000-000000000046}'),
		"RowIndex": (4, 2, (3, 0), (), "RowIndex", None),
		# Method 'Shading' returns object of type 'Shading'
		"Shading": (105, 2, (9, 0), (), "Shading", '{0002093A-0000-0000-C000-000000000046}'),
		# Method 'Tables' returns object of type 'Tables'
		"Tables": (106, 2, (9, 0), (), "Tables", '{0002094D-0000-0000-C000-000000000046}'),
		"TopPadding": (111, 2, (4, 0), (), "TopPadding", None),
		"VerticalAlignment": (1104, 2, (3, 0), (), "VerticalAlignment", None),
		"Width": (6, 2, (4, 0), (), "Width", None),
		"WordWrap": (108, 2, (11, 0), (), "WordWrap", None),
	}
	_prop_map_put_ = {
		"Borders": ((1100, LCID, 4, 0),()),
		"BottomPadding": ((112, LCID, 4, 0),()),
		"FitText": ((110, LCID, 4, 0),()),
		"Height": ((7, LCID, 4, 0),()),
		"HeightRule": ((8, LCID, 4, 0),()),
		"ID": ((115, LCID, 4, 0),()),
		"LeftPadding": ((113, LCID, 4, 0),()),
		"PreferredWidth": ((109, LCID, 4, 0),()),
		"PreferredWidthType": ((116, LCID, 4, 0),()),
		"RightPadding": ((114, LCID, 4, 0),()),
		"TopPadding": ((111, LCID, 4, 0),()),
		"VerticalAlignment": ((1104, LCID, 4, 0),()),
		"Width": ((6, LCID, 4, 0),()),
		"WordWrap": ((108, LCID, 4, 0),()),
	}
	# Default property for this class is 'Range'
	def __call__(self):
		return self._ApplyTypes_(*(0, 2, (9, 0), (), "Range", '{0002095E-0000-0000-C000-000000000046}'))
	def __str__(self, *args):
		return str(self.__call__(*args))
	def __int__(self, *args):
		return int(self.__call__(*args))
	def __iter__(self):
		"Return a Python iterator for this object"
		try:
			ob = self._oleobj_.InvokeTypes(-4,LCID,3,(13, 10),())
		except pythoncom.error:
			raise TypeError("This object does not support enumeration")
		return win32com.client.util.Iterator(ob, None)

win32com.client.CLSIDToClass.RegisterCLSID( "{0002094E-0000-0000-C000-000000000046}", Cell )
# -*- coding: mbcs -*-
# Created by makepy.py version 0.5.01
# By python version 3.3.1 (v3.3.1:d9893d13c628, Apr  6 2013, 20:25:12) [MSC v.1600 32 bit (Intel)]
# From type library '{00020905-0000-0000-C000-000000000046}'
# On Thu Mar  6 22:05:15 2014
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

Cell_vtables_dispatch_ = 1
Cell_vtables_ = [
	(( 'Range' , 'prop' , ), 0, (0, (), [ (16393, 10, None, "IID('{0002095E-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 28 , (3, 0, None, None) , 0 , )),
	(( 'Application' , 'prop' , ), 1000, (1000, (), [ (16397, 10, None, "IID('{000209FF-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 32 , (3, 0, None, None) , 0 , )),
	(( 'Creator' , 'prop' , ), 1001, (1001, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 36 , (3, 0, None, None) , 0 , )),
	(( 'Parent' , 'prop' , ), 1002, (1002, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 40 , (3, 0, None, None) , 0 , )),
	(( 'RowIndex' , 'prop' , ), 4, (4, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 44 , (3, 0, None, None) , 0 , )),
	(( 'ColumnIndex' , 'prop' , ), 5, (5, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 48 , (3, 0, None, None) , 0 , )),
	(( 'Width' , 'prop' , ), 6, (6, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 52 , (3, 0, None, None) , 0 , )),
	(( 'Width' , 'prop' , ), 6, (6, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 56 , (3, 0, None, None) , 0 , )),
	(( 'Height' , 'prop' , ), 7, (7, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 60 , (3, 0, None, None) , 0 , )),
	(( 'Height' , 'prop' , ), 7, (7, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 64 , (3, 0, None, None) , 0 , )),
	(( 'HeightRule' , 'prop' , ), 8, (8, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 68 , (3, 0, None, None) , 0 , )),
	(( 'HeightRule' , 'prop' , ), 8, (8, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 72 , (3, 0, None, None) , 0 , )),
	(( 'VerticalAlignment' , 'prop' , ), 1104, (1104, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 76 , (3, 0, None, None) , 0 , )),
	(( 'VerticalAlignment' , 'prop' , ), 1104, (1104, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 80 , (3, 0, None, None) , 0 , )),
	(( 'Column' , 'prop' , ), 101, (101, (), [ (16393, 10, None, "IID('{0002094F-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 84 , (3, 0, None, None) , 0 , )),
	(( 'Row' , 'prop' , ), 102, (102, (), [ (16393, 10, None, "IID('{00020950-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 88 , (3, 0, None, None) , 0 , )),
	(( 'Next' , 'prop' , ), 103, (103, (), [ (16393, 10, None, "IID('{0002094E-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 92 , (3, 0, None, None) , 0 , )),
	(( 'Previous' , 'prop' , ), 104, (104, (), [ (16393, 10, None, "IID('{0002094E-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 96 , (3, 0, None, None) , 0 , )),
	(( 'Shading' , 'prop' , ), 105, (105, (), [ (16393, 10, None, "IID('{0002093A-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 100 , (3, 0, None, None) , 0 , )),
	(( 'Borders' , 'prop' , ), 1100, (1100, (), [ (16393, 10, None, "IID('{0002093C-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 104 , (3, 0, None, None) , 0 , )),
	(( 'Borders' , 'prop' , ), 1100, (1100, (), [ (9, 1, None, "IID('{0002093C-0000-0000-C000-000000000046}')") , ], 1 , 4 , 4 , 0 , 108 , (3, 0, None, None) , 0 , )),
	(( 'Select' , ), 65535, (65535, (), [ ], 1 , 1 , 4 , 0 , 112 , (3, 0, None, None) , 0 , )),
	(( 'Delete' , 'ShiftCells' , ), 200, (200, (), [ (16396, 17, None, None) , ], 1 , 1 , 4 , 1 , 116 , (3, 0, None, None) , 0 , )),
	(( 'Formula' , 'Formula' , 'NumFormat' , ), 201, (201, (), [ (16396, 17, None, None) , 
			 (16396, 17, None, None) , ], 1 , 1 , 4 , 2 , 120 , (3, 0, None, None) , 0 , )),
	(( 'SetWidth' , 'ColumnWidth' , 'RulerStyle' , ), 202, (202, (), [ (4, 1, None, None) , 
			 (3, 1, None, None) , ], 1 , 1 , 4 , 0 , 124 , (3, 0, None, None) , 0 , )),
	(( 'SetHeight' , 'RowHeight' , 'HeightRule' , ), 203, (203, (), [ (16396, 1, None, None) , 
			 (3, 1, None, None) , ], 1 , 1 , 4 , 0 , 128 , (3, 0, None, None) , 0 , )),
	(( 'Merge' , 'MergeTo' , ), 204, (204, (), [ (9, 1, None, "IID('{0002094E-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 0 , 132 , (3, 0, None, None) , 0 , )),
	(( 'Split' , 'NumRows' , 'NumColumns' , ), 205, (205, (), [ (16396, 17, None, None) , 
			 (16396, 17, None, None) , ], 1 , 1 , 4 , 2 , 136 , (3, 0, None, None) , 0 , )),
	(( 'AutoSum' , ), 206, (206, (), [ ], 1 , 1 , 4 , 0 , 140 , (3, 0, None, None) , 0 , )),
	(( 'Tables' , 'prop' , ), 106, (106, (), [ (16393, 10, None, "IID('{0002094D-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 144 , (3, 0, None, None) , 0 , )),
	(( 'NestingLevel' , 'prop' , ), 107, (107, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 148 , (3, 0, None, None) , 0 , )),
	(( 'WordWrap' , 'prop' , ), 108, (108, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 152 , (3, 0, None, None) , 0 , )),
	(( 'WordWrap' , 'prop' , ), 108, (108, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 156 , (3, 0, None, None) , 0 , )),
	(( 'PreferredWidth' , 'prop' , ), 109, (109, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 160 , (3, 0, None, None) , 0 , )),
	(( 'PreferredWidth' , 'prop' , ), 109, (109, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 164 , (3, 0, None, None) , 0 , )),
	(( 'FitText' , 'prop' , ), 110, (110, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 168 , (3, 0, None, None) , 0 , )),
	(( 'FitText' , 'prop' , ), 110, (110, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 172 , (3, 0, None, None) , 0 , )),
	(( 'TopPadding' , 'prop' , ), 111, (111, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 176 , (3, 0, None, None) , 0 , )),
	(( 'TopPadding' , 'prop' , ), 111, (111, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 180 , (3, 0, None, None) , 0 , )),
	(( 'BottomPadding' , 'prop' , ), 112, (112, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 184 , (3, 0, None, None) , 0 , )),
	(( 'BottomPadding' , 'prop' , ), 112, (112, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 188 , (3, 0, None, None) , 0 , )),
	(( 'LeftPadding' , 'prop' , ), 113, (113, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 192 , (3, 0, None, None) , 0 , )),
	(( 'LeftPadding' , 'prop' , ), 113, (113, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 196 , (3, 0, None, None) , 0 , )),
	(( 'RightPadding' , 'prop' , ), 114, (114, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 200 , (3, 0, None, None) , 0 , )),
	(( 'RightPadding' , 'prop' , ), 114, (114, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 204 , (3, 0, None, None) , 0 , )),
	(( 'ID' , 'prop' , ), 115, (115, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 208 , (3, 0, None, None) , 0 , )),
	(( 'ID' , 'prop' , ), 115, (115, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 212 , (3, 0, None, None) , 0 , )),
	(( 'PreferredWidthType' , 'prop' , ), 116, (116, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 216 , (3, 0, None, None) , 0 , )),
	(( 'PreferredWidthType' , 'prop' , ), 116, (116, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 220 , (3, 0, None, None) , 0 , )),
]

win32com.client.CLSIDToClass.RegisterCLSID( "{0002094E-0000-0000-C000-000000000046}", Cell )
