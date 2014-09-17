# -*- coding: mbcs -*-
# Created by makepy.py version 0.5.01
# By python version 3.3.1 (v3.3.1:d9893d13c628, Apr  6 2013, 20:25:12) [MSC v.1600 32 bit (Intel)]
# From type library '{00020905-0000-0000-C000-000000000046}'
# On Thu Mar  6 21:59:44 2014
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
class Table(DispatchBaseClass):
	CLSID = IID('{00020951-0000-0000-C000-000000000046}')
	coclass_clsid = None

	def ApplyStyleDirectFormatting(self, StyleName=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(208, LCID, 1, (24, 0), ((8, 1),),StyleName
			)

	def AutoFitBehavior(self, Behavior=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(20, LCID, 1, (24, 0), ((3, 1),),Behavior
			)

	def AutoFormat(self, Format=defaultNamedOptArg, ApplyBorders=defaultNamedOptArg, ApplyShading=defaultNamedOptArg, ApplyFont=defaultNamedOptArg
			, ApplyColor=defaultNamedOptArg, ApplyHeadingRows=defaultNamedOptArg, ApplyLastRow=defaultNamedOptArg, ApplyFirstColumn=defaultNamedOptArg, ApplyLastColumn=defaultNamedOptArg
			, AutoFit=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(14, LCID, 1, (24, 0), ((16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),Format
			, ApplyBorders, ApplyShading, ApplyFont, ApplyColor, ApplyHeadingRows
			, ApplyLastRow, ApplyFirstColumn, ApplyLastColumn, AutoFit)

	# Result is of type Cell
	def Cell(self, Row=defaultNamedNotOptArg, Column=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(17, LCID, 1, (9, 0), ((3, 1), (3, 1)),Row
			, Column)
		if ret is not None:
			ret = Dispatch(ret, 'Cell', '{0002094E-0000-0000-C000-000000000046}')
		return ret

	# Result is of type Range
	def ConvertToText(self, Separator=defaultNamedOptArg, NestedTables=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(19, LCID, 1, (9, 0), ((16396, 17), (16396, 17)),Separator
			, NestedTables)
		if ret is not None:
			ret = Dispatch(ret, 'ConvertToText', '{0002095E-0000-0000-C000-000000000046}')
		return ret

	# Result is of type Range
	def ConvertToTextOld(self, Separator=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(16, LCID, 1, (9, 0), ((16396, 17),),Separator
			)
		if ret is not None:
			ret = Dispatch(ret, 'ConvertToTextOld', '{0002095E-0000-0000-C000-000000000046}')
		return ret

	def Delete(self):
		return self._oleobj_.InvokeTypes(9, LCID, 1, (24, 0), (),)

	def Select(self):
		return self._oleobj_.InvokeTypes(200, LCID, 1, (24, 0), (),)

	def Sort(self, ExcludeHeader=defaultNamedOptArg, FieldNumber=defaultNamedOptArg, SortFieldType=defaultNamedOptArg, SortOrder=defaultNamedOptArg
			, FieldNumber2=defaultNamedOptArg, SortFieldType2=defaultNamedOptArg, SortOrder2=defaultNamedOptArg, FieldNumber3=defaultNamedOptArg, SortFieldType3=defaultNamedOptArg
			, SortOrder3=defaultNamedOptArg, CaseSensitive=defaultNamedOptArg, BidiSort=defaultNamedOptArg, IgnoreThe=defaultNamedOptArg, IgnoreKashida=defaultNamedOptArg
			, IgnoreDiacritics=defaultNamedOptArg, IgnoreHe=defaultNamedOptArg, LanguageID=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(23, LCID, 1, (24, 0), ((16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),ExcludeHeader
			, FieldNumber, SortFieldType, SortOrder, FieldNumber2, SortFieldType2
			, SortOrder2, FieldNumber3, SortFieldType3, SortOrder3, CaseSensitive
			, BidiSort, IgnoreThe, IgnoreKashida, IgnoreDiacritics, IgnoreHe
			, LanguageID)

	def SortAscending(self):
		return self._oleobj_.InvokeTypes(12, LCID, 1, (24, 0), (),)

	def SortDescending(self):
		return self._oleobj_.InvokeTypes(13, LCID, 1, (24, 0), (),)

	def SortOld(self, ExcludeHeader=defaultNamedOptArg, FieldNumber=defaultNamedOptArg, SortFieldType=defaultNamedOptArg, SortOrder=defaultNamedOptArg
			, FieldNumber2=defaultNamedOptArg, SortFieldType2=defaultNamedOptArg, SortOrder2=defaultNamedOptArg, FieldNumber3=defaultNamedOptArg, SortFieldType3=defaultNamedOptArg
			, SortOrder3=defaultNamedOptArg, CaseSensitive=defaultNamedOptArg, LanguageID=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(10, LCID, 1, (24, 0), ((16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),ExcludeHeader
			, FieldNumber, SortFieldType, SortOrder, FieldNumber2, SortFieldType2
			, SortOrder2, FieldNumber3, SortFieldType3, SortOrder3, CaseSensitive
			, LanguageID)

	# Result is of type Table
	def Split(self, BeforeRow=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(18, LCID, 1, (9, 0), ((16396, 1),),BeforeRow
			)
		if ret is not None:
			ret = Dispatch(ret, 'Split', '{00020951-0000-0000-C000-000000000046}')
		return ret

	def UpdateAutoFormat(self):
		return self._oleobj_.InvokeTypes(15, LCID, 1, (24, 0), (),)

	_prop_map_get_ = {
		"AllowAutoFit": (110, 2, (11, 0), (), "AllowAutoFit", None),
		"AllowPageBreaks": (109, 2, (11, 0), (), "AllowPageBreaks", None),
		# Method 'Application' returns object of type 'Application'
		"Application": (1000, 2, (13, 0), (), "Application", '{000209FF-0000-0000-C000-000000000046}'),
		"ApplyStyleColumnBands": (207, 2, (11, 0), (), "ApplyStyleColumnBands", None),
		"ApplyStyleFirstColumn": (204, 2, (11, 0), (), "ApplyStyleFirstColumn", None),
		"ApplyStyleHeadingRows": (202, 2, (11, 0), (), "ApplyStyleHeadingRows", None),
		"ApplyStyleLastColumn": (205, 2, (11, 0), (), "ApplyStyleLastColumn", None),
		"ApplyStyleLastRow": (203, 2, (11, 0), (), "ApplyStyleLastRow", None),
		"ApplyStyleRowBands": (206, 2, (11, 0), (), "ApplyStyleRowBands", None),
		"AutoFormatType": (106, 2, (3, 0), (), "AutoFormatType", None),
		# Method 'Borders' returns object of type 'Borders'
		"Borders": (1100, 2, (9, 0), (), "Borders", '{0002093C-0000-0000-C000-000000000046}'),
		"BottomPadding": (114, 2, (4, 0), (), "BottomPadding", None),
		# Method 'Columns' returns object of type 'Columns'
		"Columns": (100, 2, (9, 0), (), "Columns", '{0002094B-0000-0000-C000-000000000046}'),
		"Creator": (1001, 2, (3, 0), (), "Creator", None),
		"Descr": (210, 2, (8, 0), (), "Descr", None),
		"ID": (119, 2, (8, 0), (), "ID", None),
		"LeftPadding": (115, 2, (4, 0), (), "LeftPadding", None),
		"NestingLevel": (108, 2, (3, 0), (), "NestingLevel", None),
		"Parent": (1002, 2, (9, 0), (), "Parent", None),
		"PreferredWidth": (111, 2, (4, 0), (), "PreferredWidth", None),
		"PreferredWidthType": (112, 2, (3, 0), (), "PreferredWidthType", None),
		# Method 'Range' returns object of type 'Range'
		"Range": (0, 2, (9, 0), (), "Range", '{0002095E-0000-0000-C000-000000000046}'),
		"RightPadding": (116, 2, (4, 0), (), "RightPadding", None),
		# Method 'Rows' returns object of type 'Rows'
		"Rows": (101, 2, (9, 0), (), "Rows", '{0002094C-0000-0000-C000-000000000046}'),
		# Method 'Shading' returns object of type 'Shading'
		"Shading": (104, 2, (9, 0), (), "Shading", '{0002093A-0000-0000-C000-000000000046}'),
		"Spacing": (117, 2, (4, 0), (), "Spacing", None),
		"Style": (201, 2, (12, 0), (), "Style", None),
		"TableDirection": (118, 2, (3, 0), (), "TableDirection", None),
		# Method 'Tables' returns object of type 'Tables'
		"Tables": (107, 2, (9, 0), (), "Tables", '{0002094D-0000-0000-C000-000000000046}'),
		"Title": (209, 2, (8, 0), (), "Title", None),
		"TopPadding": (113, 2, (4, 0), (), "TopPadding", None),
		"Uniform": (105, 2, (11, 0), (), "Uniform", None),
	}
	_prop_map_put_ = {
		"AllowAutoFit": ((110, LCID, 4, 0),()),
		"AllowPageBreaks": ((109, LCID, 4, 0),()),
		"ApplyStyleColumnBands": ((207, LCID, 4, 0),()),
		"ApplyStyleFirstColumn": ((204, LCID, 4, 0),()),
		"ApplyStyleHeadingRows": ((202, LCID, 4, 0),()),
		"ApplyStyleLastColumn": ((205, LCID, 4, 0),()),
		"ApplyStyleLastRow": ((203, LCID, 4, 0),()),
		"ApplyStyleRowBands": ((206, LCID, 4, 0),()),
		"Borders": ((1100, LCID, 4, 0),()),
		"BottomPadding": ((114, LCID, 4, 0),()),
		"Descr": ((210, LCID, 4, 0),()),
		"ID": ((119, LCID, 4, 0),()),
		"LeftPadding": ((115, LCID, 4, 0),()),
		"PreferredWidth": ((111, LCID, 4, 0),()),
		"PreferredWidthType": ((112, LCID, 4, 0),()),
		"RightPadding": ((116, LCID, 4, 0),()),
		"Spacing": ((117, LCID, 4, 0),()),
		"Style": ((201, LCID, 4, 0),()),
		"TableDirection": ((118, LCID, 4, 0),()),
		"Title": ((209, LCID, 4, 0),()),
		"TopPadding": ((113, LCID, 4, 0),()),
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

win32com.client.CLSIDToClass.RegisterCLSID( "{00020951-0000-0000-C000-000000000046}", Table )
# -*- coding: mbcs -*-
# Created by makepy.py version 0.5.01
# By python version 3.3.1 (v3.3.1:d9893d13c628, Apr  6 2013, 20:25:12) [MSC v.1600 32 bit (Intel)]
# From type library '{00020905-0000-0000-C000-000000000046}'
# On Thu Mar  6 21:59:44 2014
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

Table_vtables_dispatch_ = 1
Table_vtables_ = [
	(( 'Range' , 'prop' , ), 0, (0, (), [ (16393, 10, None, "IID('{0002095E-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 28 , (3, 0, None, None) , 0 , )),
	(( 'Application' , 'prop' , ), 1000, (1000, (), [ (16397, 10, None, "IID('{000209FF-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 32 , (3, 0, None, None) , 0 , )),
	(( 'Creator' , 'prop' , ), 1001, (1001, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 36 , (3, 0, None, None) , 0 , )),
	(( 'Parent' , 'prop' , ), 1002, (1002, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 40 , (3, 0, None, None) , 0 , )),
	(( 'Columns' , 'prop' , ), 100, (100, (), [ (16393, 10, None, "IID('{0002094B-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 44 , (3, 0, None, None) , 0 , )),
	(( 'Rows' , 'prop' , ), 101, (101, (), [ (16393, 10, None, "IID('{0002094C-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 48 , (3, 0, None, None) , 0 , )),
	(( 'Borders' , 'prop' , ), 1100, (1100, (), [ (16393, 10, None, "IID('{0002093C-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 52 , (3, 0, None, None) , 0 , )),
	(( 'Borders' , 'prop' , ), 1100, (1100, (), [ (9, 1, None, "IID('{0002093C-0000-0000-C000-000000000046}')") , ], 1 , 4 , 4 , 0 , 56 , (3, 0, None, None) , 0 , )),
	(( 'Shading' , 'prop' , ), 104, (104, (), [ (16393, 10, None, "IID('{0002093A-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 60 , (3, 0, None, None) , 0 , )),
	(( 'Uniform' , 'prop' , ), 105, (105, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 64 , (3, 0, None, None) , 0 , )),
	(( 'AutoFormatType' , 'prop' , ), 106, (106, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 68 , (3, 0, None, None) , 0 , )),
	(( 'Select' , ), 200, (200, (), [ ], 1 , 1 , 4 , 0 , 72 , (3, 0, None, None) , 0 , )),
	(( 'Delete' , ), 9, (9, (), [ ], 1 , 1 , 4 , 0 , 76 , (3, 0, None, None) , 0 , )),
	(( 'SortOld' , 'ExcludeHeader' , 'FieldNumber' , 'SortFieldType' , 'SortOrder' , 
			 'FieldNumber2' , 'SortFieldType2' , 'SortOrder2' , 'FieldNumber3' , 'SortFieldType3' , 
			 'SortOrder3' , 'CaseSensitive' , 'LanguageID' , ), 10, (10, (), [ (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , ], 1 , 1 , 4 , 12 , 80 , (3, 0, None, None) , 64 , )),
	(( 'SortAscending' , ), 12, (12, (), [ ], 1 , 1 , 4 , 0 , 84 , (3, 0, None, None) , 0 , )),
	(( 'SortDescending' , ), 13, (13, (), [ ], 1 , 1 , 4 , 0 , 88 , (3, 0, None, None) , 0 , )),
	(( 'AutoFormat' , 'Format' , 'ApplyBorders' , 'ApplyShading' , 'ApplyFont' , 
			 'ApplyColor' , 'ApplyHeadingRows' , 'ApplyLastRow' , 'ApplyFirstColumn' , 'ApplyLastColumn' , 
			 'AutoFit' , ), 14, (14, (), [ (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , ], 1 , 1 , 4 , 10 , 92 , (3, 0, None, None) , 0 , )),
	(( 'UpdateAutoFormat' , ), 15, (15, (), [ ], 1 , 1 , 4 , 0 , 96 , (3, 0, None, None) , 0 , )),
	(( 'ConvertToTextOld' , 'Separator' , 'prop' , ), 16, (16, (), [ (16396, 17, None, None) , 
			 (16393, 10, None, "IID('{0002095E-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 1 , 100 , (3, 0, None, None) , 64 , )),
	(( 'Cell' , 'Row' , 'Column' , 'prop' , ), 17, (17, (), [ 
			 (3, 1, None, None) , (3, 1, None, None) , (16393, 10, None, "IID('{0002094E-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 0 , 104 , (3, 0, None, None) , 0 , )),
	(( 'Split' , 'BeforeRow' , 'prop' , ), 18, (18, (), [ (16396, 1, None, None) , 
			 (16393, 10, None, "IID('{00020951-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 0 , 108 , (3, 0, None, None) , 0 , )),
	(( 'ConvertToText' , 'Separator' , 'NestedTables' , 'prop' , ), 19, (19, (), [ 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16393, 10, None, "IID('{0002095E-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 2 , 112 , (3, 0, None, None) , 0 , )),
	(( 'AutoFitBehavior' , 'Behavior' , ), 20, (20, (), [ (3, 1, None, None) , ], 1 , 1 , 4 , 0 , 116 , (3, 0, None, None) , 0 , )),
	(( 'Sort' , 'ExcludeHeader' , 'FieldNumber' , 'SortFieldType' , 'SortOrder' , 
			 'FieldNumber2' , 'SortFieldType2' , 'SortOrder2' , 'FieldNumber3' , 'SortFieldType3' , 
			 'SortOrder3' , 'CaseSensitive' , 'BidiSort' , 'IgnoreThe' , 'IgnoreKashida' , 
			 'IgnoreDiacritics' , 'IgnoreHe' , 'LanguageID' , ), 23, (23, (), [ (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , ], 1 , 1 , 4 , 17 , 120 , (3, 0, None, None) , 0 , )),
	(( 'Tables' , 'prop' , ), 107, (107, (), [ (16393, 10, None, "IID('{0002094D-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 124 , (3, 0, None, None) , 0 , )),
	(( 'NestingLevel' , 'prop' , ), 108, (108, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 128 , (3, 0, None, None) , 0 , )),
	(( 'AllowPageBreaks' , 'prop' , ), 109, (109, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 132 , (3, 0, None, None) , 64 , )),
	(( 'AllowPageBreaks' , 'prop' , ), 109, (109, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 136 , (3, 0, None, None) , 64 , )),
	(( 'AllowAutoFit' , 'prop' , ), 110, (110, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 140 , (3, 0, None, None) , 0 , )),
	(( 'AllowAutoFit' , 'prop' , ), 110, (110, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 144 , (3, 0, None, None) , 0 , )),
	(( 'PreferredWidth' , 'prop' , ), 111, (111, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 148 , (3, 0, None, None) , 0 , )),
	(( 'PreferredWidth' , 'prop' , ), 111, (111, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 152 , (3, 0, None, None) , 0 , )),
	(( 'PreferredWidthType' , 'prop' , ), 112, (112, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 156 , (3, 0, None, None) , 0 , )),
	(( 'PreferredWidthType' , 'prop' , ), 112, (112, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 160 , (3, 0, None, None) , 0 , )),
	(( 'TopPadding' , 'prop' , ), 113, (113, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 164 , (3, 0, None, None) , 0 , )),
	(( 'TopPadding' , 'prop' , ), 113, (113, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 168 , (3, 0, None, None) , 0 , )),
	(( 'BottomPadding' , 'prop' , ), 114, (114, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 172 , (3, 0, None, None) , 0 , )),
	(( 'BottomPadding' , 'prop' , ), 114, (114, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 176 , (3, 0, None, None) , 0 , )),
	(( 'LeftPadding' , 'prop' , ), 115, (115, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 180 , (3, 0, None, None) , 0 , )),
	(( 'LeftPadding' , 'prop' , ), 115, (115, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 184 , (3, 0, None, None) , 0 , )),
	(( 'RightPadding' , 'prop' , ), 116, (116, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 188 , (3, 0, None, None) , 0 , )),
	(( 'RightPadding' , 'prop' , ), 116, (116, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 192 , (3, 0, None, None) , 0 , )),
	(( 'Spacing' , 'prop' , ), 117, (117, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 196 , (3, 0, None, None) , 0 , )),
	(( 'Spacing' , 'prop' , ), 117, (117, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 200 , (3, 0, None, None) , 0 , )),
	(( 'TableDirection' , 'prop' , ), 118, (118, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 204 , (3, 0, None, None) , 0 , )),
	(( 'TableDirection' , 'prop' , ), 118, (118, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 208 , (3, 0, None, None) , 0 , )),
	(( 'ID' , 'prop' , ), 119, (119, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 212 , (3, 0, None, None) , 0 , )),
	(( 'ID' , 'prop' , ), 119, (119, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 216 , (3, 0, None, None) , 0 , )),
	(( 'Style' , 'prop' , ), 201, (201, (), [ (16396, 10, None, None) , ], 1 , 2 , 4 , 0 , 220 , (3, 0, None, None) , 0 , )),
	(( 'Style' , 'prop' , ), 201, (201, (), [ (16396, 1, None, None) , ], 1 , 4 , 4 , 0 , 224 , (3, 0, None, None) , 0 , )),
	(( 'ApplyStyleHeadingRows' , 'prop' , ), 202, (202, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 228 , (3, 0, None, None) , 0 , )),
	(( 'ApplyStyleHeadingRows' , 'prop' , ), 202, (202, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 232 , (3, 0, None, None) , 0 , )),
	(( 'ApplyStyleLastRow' , 'prop' , ), 203, (203, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 236 , (3, 0, None, None) , 0 , )),
	(( 'ApplyStyleLastRow' , 'prop' , ), 203, (203, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 240 , (3, 0, None, None) , 0 , )),
	(( 'ApplyStyleFirstColumn' , 'prop' , ), 204, (204, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 244 , (3, 0, None, None) , 0 , )),
	(( 'ApplyStyleFirstColumn' , 'prop' , ), 204, (204, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 248 , (3, 0, None, None) , 0 , )),
	(( 'ApplyStyleLastColumn' , 'prop' , ), 205, (205, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 252 , (3, 0, None, None) , 0 , )),
	(( 'ApplyStyleLastColumn' , 'prop' , ), 205, (205, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 256 , (3, 0, None, None) , 0 , )),
	(( 'ApplyStyleRowBands' , 'prop' , ), 206, (206, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 260 , (3, 0, None, None) , 0 , )),
	(( 'ApplyStyleRowBands' , 'prop' , ), 206, (206, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 264 , (3, 0, None, None) , 0 , )),
	(( 'ApplyStyleColumnBands' , 'prop' , ), 207, (207, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 268 , (3, 0, None, None) , 0 , )),
	(( 'ApplyStyleColumnBands' , 'prop' , ), 207, (207, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 272 , (3, 0, None, None) , 0 , )),
	(( 'ApplyStyleDirectFormatting' , 'StyleName' , ), 208, (208, (), [ (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 276 , (3, 0, None, None) , 0 , )),
	(( 'Title' , 'prop' , ), 209, (209, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 280 , (3, 0, None, None) , 0 , )),
	(( 'Title' , 'prop' , ), 209, (209, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 284 , (3, 0, None, None) , 0 , )),
	(( 'Descr' , 'prop' , ), 210, (210, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 288 , (3, 0, None, None) , 0 , )),
	(( 'Descr' , 'prop' , ), 210, (210, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 292 , (3, 0, None, None) , 0 , )),
]

win32com.client.CLSIDToClass.RegisterCLSID( "{00020951-0000-0000-C000-000000000046}", Table )
