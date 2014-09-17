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

from win32com.client import DispatchBaseClass
class _ParagraphFormat(DispatchBaseClass):
	CLSID = IID('{00020953-0000-0000-C000-000000000046}')
	coclass_clsid = IID('{000209F4-0000-0000-C000-000000000046}')

	def CloseUp(self):
		return self._oleobj_.InvokeTypes(301, LCID, 1, (24, 0), (),)

	def IndentCharWidth(self, Count=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(320, LCID, 1, (24, 0), ((2, 1),),Count
			)

	def IndentFirstLineCharWidth(self, Count=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(322, LCID, 1, (24, 0), ((2, 1),),Count
			)

	def OpenOrCloseUp(self):
		return self._oleobj_.InvokeTypes(303, LCID, 1, (24, 0), (),)

	def OpenUp(self):
		return self._oleobj_.InvokeTypes(302, LCID, 1, (24, 0), (),)

	def Reset(self):
		return self._oleobj_.InvokeTypes(312, LCID, 1, (24, 0), (),)

	def Space1(self):
		return self._oleobj_.InvokeTypes(313, LCID, 1, (24, 0), (),)

	def Space15(self):
		return self._oleobj_.InvokeTypes(314, LCID, 1, (24, 0), (),)

	def Space2(self):
		return self._oleobj_.InvokeTypes(315, LCID, 1, (24, 0), (),)

	def TabHangingIndent(self, Count=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(304, LCID, 1, (24, 0), ((2, 1),),Count
			)

	def TabIndent(self, Count=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(306, LCID, 1, (24, 0), ((2, 1),),Count
			)

	_prop_map_get_ = {
		"AddSpaceBetweenFarEastAndAlpha": (121, 2, (3, 0), (), "AddSpaceBetweenFarEastAndAlpha", None),
		"AddSpaceBetweenFarEastAndDigit": (122, 2, (3, 0), (), "AddSpaceBetweenFarEastAndDigit", None),
		"Alignment": (101, 2, (3, 0), (), "Alignment", None),
		# Method 'Application' returns object of type 'Application'
		"Application": (1000, 2, (13, 0), (), "Application", '{000209FF-0000-0000-C000-000000000046}'),
		"AutoAdjustRightIndent": (124, 2, (3, 0), (), "AutoAdjustRightIndent", None),
		"BaseLineAlignment": (123, 2, (3, 0), (), "BaseLineAlignment", None),
		# Method 'Borders' returns object of type 'Borders'
		"Borders": (1100, 2, (9, 0), (), "Borders", '{0002093C-0000-0000-C000-000000000046}'),
		"CharacterUnitFirstLineIndent": (128, 2, (4, 0), (), "CharacterUnitFirstLineIndent", None),
		"CharacterUnitLeftIndent": (127, 2, (4, 0), (), "CharacterUnitLeftIndent", None),
		"CharacterUnitRightIndent": (126, 2, (4, 0), (), "CharacterUnitRightIndent", None),
		"Creator": (1001, 2, (3, 0), (), "Creator", None),
		"DisableLineHeightGrid": (125, 2, (3, 0), (), "DisableLineHeightGrid", None),
		# Method 'Duplicate' returns object of type 'ParagraphFormat'
		"Duplicate": (10, 2, (13, 0), (), "Duplicate", '{000209F4-0000-0000-C000-000000000046}'),
		"FarEastLineBreakControl": (117, 2, (3, 0), (), "FarEastLineBreakControl", None),
		"FirstLineIndent": (108, 2, (4, 0), (), "FirstLineIndent", None),
		"HalfWidthPunctuationOnTopOfLine": (120, 2, (3, 0), (), "HalfWidthPunctuationOnTopOfLine", None),
		"HangingPunctuation": (119, 2, (3, 0), (), "HangingPunctuation", None),
		"Hyphenation": (113, 2, (3, 0), (), "Hyphenation", None),
		"KeepTogether": (102, 2, (3, 0), (), "KeepTogether", None),
		"KeepWithNext": (103, 2, (3, 0), (), "KeepWithNext", None),
		"LeftIndent": (107, 2, (4, 0), (), "LeftIndent", None),
		"LineSpacing": (109, 2, (4, 0), (), "LineSpacing", None),
		"LineSpacingRule": (110, 2, (3, 0), (), "LineSpacingRule", None),
		"LineUnitAfter": (130, 2, (4, 0), (), "LineUnitAfter", None),
		"LineUnitBefore": (129, 2, (4, 0), (), "LineUnitBefore", None),
		"MirrorIndents": (134, 2, (3, 0), (), "MirrorIndents", None),
		"NoLineNumber": (105, 2, (3, 0), (), "NoLineNumber", None),
		"OutlineLevel": (202, 2, (3, 0), (), "OutlineLevel", None),
		"PageBreakBefore": (104, 2, (3, 0), (), "PageBreakBefore", None),
		"Parent": (1002, 2, (9, 0), (), "Parent", None),
		"ReadingOrder": (131, 2, (3, 0), (), "ReadingOrder", None),
		"RightIndent": (106, 2, (4, 0), (), "RightIndent", None),
		# Method 'Shading' returns object of type 'Shading'
		"Shading": (1101, 2, (9, 0), (), "Shading", '{0002093A-0000-0000-C000-000000000046}'),
		"SpaceAfter": (112, 2, (4, 0), (), "SpaceAfter", None),
		"SpaceAfterAuto": (133, 2, (3, 0), (), "SpaceAfterAuto", None),
		"SpaceBefore": (111, 2, (4, 0), (), "SpaceBefore", None),
		"SpaceBeforeAuto": (132, 2, (3, 0), (), "SpaceBeforeAuto", None),
		"Style": (100, 2, (12, 0), (), "Style", None),
		# Method 'TabStops' returns object of type 'TabStops'
		"TabStops": (1103, 2, (9, 0), (), "TabStops", '{00020955-0000-0000-C000-000000000046}'),
		"TextboxTightWrap": (135, 2, (3, 0), (), "TextboxTightWrap", None),
		"WidowControl": (114, 2, (3, 0), (), "WidowControl", None),
		"WordWrap": (118, 2, (3, 0), (), "WordWrap", None),
	}
	_prop_map_put_ = {
		"AddSpaceBetweenFarEastAndAlpha": ((121, LCID, 4, 0),()),
		"AddSpaceBetweenFarEastAndDigit": ((122, LCID, 4, 0),()),
		"Alignment": ((101, LCID, 4, 0),()),
		"AutoAdjustRightIndent": ((124, LCID, 4, 0),()),
		"BaseLineAlignment": ((123, LCID, 4, 0),()),
		"Borders": ((1100, LCID, 4, 0),()),
		"CharacterUnitFirstLineIndent": ((128, LCID, 4, 0),()),
		"CharacterUnitLeftIndent": ((127, LCID, 4, 0),()),
		"CharacterUnitRightIndent": ((126, LCID, 4, 0),()),
		"DisableLineHeightGrid": ((125, LCID, 4, 0),()),
		"FarEastLineBreakControl": ((117, LCID, 4, 0),()),
		"FirstLineIndent": ((108, LCID, 4, 0),()),
		"HalfWidthPunctuationOnTopOfLine": ((120, LCID, 4, 0),()),
		"HangingPunctuation": ((119, LCID, 4, 0),()),
		"Hyphenation": ((113, LCID, 4, 0),()),
		"KeepTogether": ((102, LCID, 4, 0),()),
		"KeepWithNext": ((103, LCID, 4, 0),()),
		"LeftIndent": ((107, LCID, 4, 0),()),
		"LineSpacing": ((109, LCID, 4, 0),()),
		"LineSpacingRule": ((110, LCID, 4, 0),()),
		"LineUnitAfter": ((130, LCID, 4, 0),()),
		"LineUnitBefore": ((129, LCID, 4, 0),()),
		"MirrorIndents": ((134, LCID, 4, 0),()),
		"NoLineNumber": ((105, LCID, 4, 0),()),
		"OutlineLevel": ((202, LCID, 4, 0),()),
		"PageBreakBefore": ((104, LCID, 4, 0),()),
		"ReadingOrder": ((131, LCID, 4, 0),()),
		"RightIndent": ((106, LCID, 4, 0),()),
		"SpaceAfter": ((112, LCID, 4, 0),()),
		"SpaceAfterAuto": ((133, LCID, 4, 0),()),
		"SpaceBefore": ((111, LCID, 4, 0),()),
		"SpaceBeforeAuto": ((132, LCID, 4, 0),()),
		"Style": ((100, LCID, 4, 0),()),
		"TabStops": ((1103, LCID, 4, 0),()),
		"TextboxTightWrap": ((135, LCID, 4, 0),()),
		"WidowControl": ((114, LCID, 4, 0),()),
		"WordWrap": ((118, LCID, 4, 0),()),
	}
	def __iter__(self):
		"Return a Python iterator for this object"
		try:
			ob = self._oleobj_.InvokeTypes(-4,LCID,3,(13, 10),())
		except pythoncom.error:
			raise TypeError("This object does not support enumeration")
		return win32com.client.util.Iterator(ob, None)

win32com.client.CLSIDToClass.RegisterCLSID( "{00020953-0000-0000-C000-000000000046}", _ParagraphFormat )
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

_ParagraphFormat_vtables_dispatch_ = 1
_ParagraphFormat_vtables_ = [
	(( 'Application' , 'prop' , ), 1000, (1000, (), [ (16397, 10, None, "IID('{000209FF-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 28 , (3, 0, None, None) , 0 , )),
	(( 'Creator' , 'prop' , ), 1001, (1001, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 32 , (3, 0, None, None) , 0 , )),
	(( 'Parent' , 'prop' , ), 1002, (1002, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 36 , (3, 0, None, None) , 0 , )),
	(( 'Duplicate' , 'prop' , ), 10, (10, (), [ (16397, 10, None, "IID('{000209F4-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 40 , (3, 0, None, None) , 0 , )),
	(( 'Style' , 'prop' , ), 100, (100, (), [ (16396, 10, None, None) , ], 1 , 2 , 4 , 0 , 44 , (3, 0, None, None) , 0 , )),
	(( 'Style' , 'prop' , ), 100, (100, (), [ (16396, 1, None, None) , ], 1 , 4 , 4 , 0 , 48 , (3, 0, None, None) , 0 , )),
	(( 'Alignment' , 'prop' , ), 101, (101, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 52 , (3, 0, None, None) , 0 , )),
	(( 'Alignment' , 'prop' , ), 101, (101, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 56 , (3, 0, None, None) , 0 , )),
	(( 'KeepTogether' , 'prop' , ), 102, (102, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 60 , (3, 0, None, None) , 0 , )),
	(( 'KeepTogether' , 'prop' , ), 102, (102, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 64 , (3, 0, None, None) , 0 , )),
	(( 'KeepWithNext' , 'prop' , ), 103, (103, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 68 , (3, 0, None, None) , 0 , )),
	(( 'KeepWithNext' , 'prop' , ), 103, (103, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 72 , (3, 0, None, None) , 0 , )),
	(( 'PageBreakBefore' , 'prop' , ), 104, (104, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 76 , (3, 0, None, None) , 0 , )),
	(( 'PageBreakBefore' , 'prop' , ), 104, (104, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 80 , (3, 0, None, None) , 0 , )),
	(( 'NoLineNumber' , 'prop' , ), 105, (105, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 84 , (3, 0, None, None) , 0 , )),
	(( 'NoLineNumber' , 'prop' , ), 105, (105, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 88 , (3, 0, None, None) , 0 , )),
	(( 'RightIndent' , 'prop' , ), 106, (106, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 92 , (3, 0, None, None) , 0 , )),
	(( 'RightIndent' , 'prop' , ), 106, (106, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 96 , (3, 0, None, None) , 0 , )),
	(( 'LeftIndent' , 'prop' , ), 107, (107, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 100 , (3, 0, None, None) , 0 , )),
	(( 'LeftIndent' , 'prop' , ), 107, (107, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 104 , (3, 0, None, None) , 0 , )),
	(( 'FirstLineIndent' , 'prop' , ), 108, (108, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 108 , (3, 0, None, None) , 0 , )),
	(( 'FirstLineIndent' , 'prop' , ), 108, (108, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 112 , (3, 0, None, None) , 0 , )),
	(( 'LineSpacing' , 'prop' , ), 109, (109, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 116 , (3, 0, None, None) , 0 , )),
	(( 'LineSpacing' , 'prop' , ), 109, (109, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 120 , (3, 0, None, None) , 0 , )),
	(( 'LineSpacingRule' , 'prop' , ), 110, (110, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 124 , (3, 0, None, None) , 0 , )),
	(( 'LineSpacingRule' , 'prop' , ), 110, (110, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 128 , (3, 0, None, None) , 0 , )),
	(( 'SpaceBefore' , 'prop' , ), 111, (111, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 132 , (3, 0, None, None) , 0 , )),
	(( 'SpaceBefore' , 'prop' , ), 111, (111, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 136 , (3, 0, None, None) , 0 , )),
	(( 'SpaceAfter' , 'prop' , ), 112, (112, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 140 , (3, 0, None, None) , 0 , )),
	(( 'SpaceAfter' , 'prop' , ), 112, (112, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 144 , (3, 0, None, None) , 0 , )),
	(( 'Hyphenation' , 'prop' , ), 113, (113, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 148 , (3, 0, None, None) , 0 , )),
	(( 'Hyphenation' , 'prop' , ), 113, (113, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 152 , (3, 0, None, None) , 0 , )),
	(( 'WidowControl' , 'prop' , ), 114, (114, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 156 , (3, 0, None, None) , 0 , )),
	(( 'WidowControl' , 'prop' , ), 114, (114, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 160 , (3, 0, None, None) , 0 , )),
	(( 'FarEastLineBreakControl' , 'prop' , ), 117, (117, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 164 , (3, 0, None, None) , 0 , )),
	(( 'FarEastLineBreakControl' , 'prop' , ), 117, (117, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 168 , (3, 0, None, None) , 0 , )),
	(( 'WordWrap' , 'prop' , ), 118, (118, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 172 , (3, 0, None, None) , 0 , )),
	(( 'WordWrap' , 'prop' , ), 118, (118, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 176 , (3, 0, None, None) , 0 , )),
	(( 'HangingPunctuation' , 'prop' , ), 119, (119, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 180 , (3, 0, None, None) , 0 , )),
	(( 'HangingPunctuation' , 'prop' , ), 119, (119, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 184 , (3, 0, None, None) , 0 , )),
	(( 'HalfWidthPunctuationOnTopOfLine' , 'prop' , ), 120, (120, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 188 , (3, 0, None, None) , 0 , )),
	(( 'HalfWidthPunctuationOnTopOfLine' , 'prop' , ), 120, (120, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 192 , (3, 0, None, None) , 0 , )),
	(( 'AddSpaceBetweenFarEastAndAlpha' , 'prop' , ), 121, (121, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 196 , (3, 0, None, None) , 0 , )),
	(( 'AddSpaceBetweenFarEastAndAlpha' , 'prop' , ), 121, (121, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 200 , (3, 0, None, None) , 0 , )),
	(( 'AddSpaceBetweenFarEastAndDigit' , 'prop' , ), 122, (122, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 204 , (3, 0, None, None) , 0 , )),
	(( 'AddSpaceBetweenFarEastAndDigit' , 'prop' , ), 122, (122, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 208 , (3, 0, None, None) , 0 , )),
	(( 'BaseLineAlignment' , 'prop' , ), 123, (123, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 212 , (3, 0, None, None) , 0 , )),
	(( 'BaseLineAlignment' , 'prop' , ), 123, (123, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 216 , (3, 0, None, None) , 0 , )),
	(( 'AutoAdjustRightIndent' , 'prop' , ), 124, (124, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 220 , (3, 0, None, None) , 0 , )),
	(( 'AutoAdjustRightIndent' , 'prop' , ), 124, (124, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 224 , (3, 0, None, None) , 0 , )),
	(( 'DisableLineHeightGrid' , 'prop' , ), 125, (125, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 228 , (3, 0, None, None) , 0 , )),
	(( 'DisableLineHeightGrid' , 'prop' , ), 125, (125, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 232 , (3, 0, None, None) , 0 , )),
	(( 'TabStops' , 'prop' , ), 1103, (1103, (), [ (16393, 10, None, "IID('{00020955-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 236 , (3, 0, None, None) , 0 , )),
	(( 'TabStops' , 'prop' , ), 1103, (1103, (), [ (9, 1, None, "IID('{00020955-0000-0000-C000-000000000046}')") , ], 1 , 4 , 4 , 0 , 240 , (3, 0, None, None) , 0 , )),
	(( 'Borders' , 'prop' , ), 1100, (1100, (), [ (16393, 10, None, "IID('{0002093C-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 244 , (3, 0, None, None) , 0 , )),
	(( 'Borders' , 'prop' , ), 1100, (1100, (), [ (9, 1, None, "IID('{0002093C-0000-0000-C000-000000000046}')") , ], 1 , 4 , 4 , 0 , 248 , (3, 0, None, None) , 0 , )),
	(( 'Shading' , 'prop' , ), 1101, (1101, (), [ (16393, 10, None, "IID('{0002093A-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 252 , (3, 0, None, None) , 0 , )),
	(( 'OutlineLevel' , 'prop' , ), 202, (202, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 256 , (3, 0, None, None) , 0 , )),
	(( 'OutlineLevel' , 'prop' , ), 202, (202, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 260 , (3, 0, None, None) , 0 , )),
	(( 'CloseUp' , ), 301, (301, (), [ ], 1 , 1 , 4 , 0 , 264 , (3, 0, None, None) , 0 , )),
	(( 'OpenUp' , ), 302, (302, (), [ ], 1 , 1 , 4 , 0 , 268 , (3, 0, None, None) , 0 , )),
	(( 'OpenOrCloseUp' , ), 303, (303, (), [ ], 1 , 1 , 4 , 0 , 272 , (3, 0, None, None) , 0 , )),
	(( 'TabHangingIndent' , 'Count' , ), 304, (304, (), [ (2, 1, None, None) , ], 1 , 1 , 4 , 0 , 276 , (3, 0, None, None) , 0 , )),
	(( 'TabIndent' , 'Count' , ), 306, (306, (), [ (2, 1, None, None) , ], 1 , 1 , 4 , 0 , 280 , (3, 0, None, None) , 0 , )),
	(( 'Reset' , ), 312, (312, (), [ ], 1 , 1 , 4 , 0 , 284 , (3, 0, None, None) , 0 , )),
	(( 'Space1' , ), 313, (313, (), [ ], 1 , 1 , 4 , 0 , 288 , (3, 0, None, None) , 0 , )),
	(( 'Space15' , ), 314, (314, (), [ ], 1 , 1 , 4 , 0 , 292 , (3, 0, None, None) , 0 , )),
	(( 'Space2' , ), 315, (315, (), [ ], 1 , 1 , 4 , 0 , 296 , (3, 0, None, None) , 0 , )),
	(( 'IndentCharWidth' , 'Count' , ), 320, (320, (), [ (2, 1, None, None) , ], 1 , 1 , 4 , 0 , 300 , (3, 0, None, None) , 0 , )),
	(( 'IndentFirstLineCharWidth' , 'Count' , ), 322, (322, (), [ (2, 1, None, None) , ], 1 , 1 , 4 , 0 , 304 , (3, 0, None, None) , 0 , )),
	(( 'CharacterUnitRightIndent' , 'prop' , ), 126, (126, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 308 , (3, 0, None, None) , 0 , )),
	(( 'CharacterUnitRightIndent' , 'prop' , ), 126, (126, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 312 , (3, 0, None, None) , 0 , )),
	(( 'CharacterUnitLeftIndent' , 'prop' , ), 127, (127, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 316 , (3, 0, None, None) , 0 , )),
	(( 'CharacterUnitLeftIndent' , 'prop' , ), 127, (127, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 320 , (3, 0, None, None) , 0 , )),
	(( 'CharacterUnitFirstLineIndent' , 'prop' , ), 128, (128, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 324 , (3, 0, None, None) , 0 , )),
	(( 'CharacterUnitFirstLineIndent' , 'prop' , ), 128, (128, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 328 , (3, 0, None, None) , 0 , )),
	(( 'LineUnitBefore' , 'prop' , ), 129, (129, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 332 , (3, 0, None, None) , 0 , )),
	(( 'LineUnitBefore' , 'prop' , ), 129, (129, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 336 , (3, 0, None, None) , 0 , )),
	(( 'LineUnitAfter' , 'prop' , ), 130, (130, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 340 , (3, 0, None, None) , 0 , )),
	(( 'LineUnitAfter' , 'prop' , ), 130, (130, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 344 , (3, 0, None, None) , 0 , )),
	(( 'ReadingOrder' , 'prop' , ), 131, (131, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 348 , (3, 0, None, None) , 0 , )),
	(( 'ReadingOrder' , 'prop' , ), 131, (131, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 352 , (3, 0, None, None) , 0 , )),
	(( 'SpaceBeforeAuto' , 'prop' , ), 132, (132, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 356 , (3, 0, None, None) , 0 , )),
	(( 'SpaceBeforeAuto' , 'prop' , ), 132, (132, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 360 , (3, 0, None, None) , 0 , )),
	(( 'SpaceAfterAuto' , 'prop' , ), 133, (133, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 364 , (3, 0, None, None) , 0 , )),
	(( 'SpaceAfterAuto' , 'prop' , ), 133, (133, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 368 , (3, 0, None, None) , 0 , )),
	(( 'MirrorIndents' , 'prop' , ), 134, (134, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 372 , (3, 0, None, None) , 0 , )),
	(( 'MirrorIndents' , 'prop' , ), 134, (134, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 376 , (3, 0, None, None) , 0 , )),
	(( 'TextboxTightWrap' , 'prop' , ), 135, (135, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 380 , (3, 0, None, None) , 0 , )),
	(( 'TextboxTightWrap' , 'prop' , ), 135, (135, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 384 , (3, 0, None, None) , 0 , )),
]

win32com.client.CLSIDToClass.RegisterCLSID( "{00020953-0000-0000-C000-000000000046}", _ParagraphFormat )
