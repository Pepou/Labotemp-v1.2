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

from win32com.client import DispatchBaseClass
class _Worksheet(DispatchBaseClass):
	CLSID = IID('{000208D8-0000-0000-C000-000000000046}')
	coclass_clsid = IID('{00020820-0000-0000-C000-000000000046}')

	def Activate(self):
		return self._oleobj_.InvokeTypes(304, LCID, 1, (24, 0), (),)

	def Arcs(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(760, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'Arcs', None)
		return ret

	def Buttons(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(557, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'Buttons', None)
		return ret

	def Calculate(self):
		return self._oleobj_.InvokeTypes(279, LCID, 1, (24, 0), (),)

	def ChartObjects(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(1060, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'ChartObjects', None)
		return ret

	def CheckBoxes(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(824, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'CheckBoxes', None)
		return ret

	def CheckSpelling(self, CustomDictionary=defaultNamedOptArg, IgnoreUppercase=defaultNamedOptArg, AlwaysSuggest=defaultNamedOptArg, SpellLang=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(505, LCID, 1, (24, 0), ((12, 17), (12, 17), (12, 17), (12, 17)),CustomDictionary
			, IgnoreUppercase, AlwaysSuggest, SpellLang)

	def CircleInvalid(self):
		return self._oleobj_.InvokeTypes(1437, LCID, 1, (24, 0), (),)

	def ClearArrows(self):
		return self._oleobj_.InvokeTypes(970, LCID, 1, (24, 0), (),)

	def ClearCircles(self):
		return self._oleobj_.InvokeTypes(1436, LCID, 1, (24, 0), (),)

	def Copy(self, Before=defaultNamedOptArg, After=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(551, LCID, 1, (24, 0), ((12, 17), (12, 17)),Before
			, After)

	def Delete(self):
		return self._oleobj_.InvokeTypes(117, LCID, 1, (24, 0), (),)

	def DrawingObjects(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(88, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'DrawingObjects', None)
		return ret

	def Drawings(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(772, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'Drawings', None)
		return ret

	def DropDowns(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(836, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'DropDowns', None)
		return ret

	def Evaluate(self, Name=defaultNamedNotOptArg):
		return self._ApplyTypes_(1, 1, (12, 0), ((12, 1),), 'Evaluate', None,Name
			)

	def ExportAsFixedFormat(self, Type=defaultNamedNotOptArg, Filename=defaultNamedOptArg, Quality=defaultNamedOptArg, IncludeDocProperties=defaultNamedOptArg
			, IgnorePrintAreas=defaultNamedOptArg, From=defaultNamedOptArg, To=defaultNamedOptArg, OpenAfterPublish=defaultNamedOptArg, FixedFormatExtClassPtr=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2493, LCID, 1, (24, 0), ((3, 1), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)),Type
			, Filename, Quality, IncludeDocProperties, IgnorePrintAreas, From
			, To, OpenAfterPublish, FixedFormatExtClassPtr)

	def GroupBoxes(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(834, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'GroupBoxes', None)
		return ret

	def GroupObjects(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(1113, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'GroupObjects', None)
		return ret

	def Labels(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(841, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'Labels', None)
		return ret

	def Lines(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(767, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'Lines', None)
		return ret

	def ListBoxes(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(832, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'ListBoxes', None)
		return ret

	def Move(self, Before=defaultNamedOptArg, After=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(637, LCID, 1, (24, 0), ((12, 17), (12, 17)),Before
			, After)

	def OLEObjects(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(799, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'OLEObjects', None)
		return ret

	def OptionButtons(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(826, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'OptionButtons', None)
		return ret

	def Ovals(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(801, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'Ovals', None)
		return ret

	def Paste(self, Destination=defaultNamedOptArg, Link=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(211, LCID, 1, (24, 0), ((12, 17), (12, 17)),Destination
			, Link)

	def PasteSpecial(self, Format=defaultNamedOptArg, Link=defaultNamedOptArg, DisplayAsIcon=defaultNamedOptArg, IconFileName=defaultNamedOptArg
			, IconIndex=defaultNamedOptArg, IconLabel=defaultNamedOptArg, NoHTMLFormatting=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1928, LCID, 1, (24, 0), ((12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)),Format
			, Link, DisplayAsIcon, IconFileName, IconIndex, IconLabel
			, NoHTMLFormatting)

	def Pictures(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(771, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'Pictures', None)
		return ret

	# Result is of type PivotTable
	def PivotTableWizard(self, SourceType=defaultNamedOptArg, SourceData=defaultNamedOptArg, TableDestination=defaultNamedOptArg, TableName=defaultNamedOptArg
			, RowGrand=defaultNamedOptArg, ColumnGrand=defaultNamedOptArg, SaveData=defaultNamedOptArg, HasAutoFormat=defaultNamedOptArg, AutoPage=defaultNamedOptArg
			, Reserved=defaultNamedOptArg, BackgroundQuery=defaultNamedOptArg, OptimizeCache=defaultNamedOptArg, PageFieldOrder=defaultNamedOptArg, PageFieldWrapCount=defaultNamedOptArg
			, ReadData=defaultNamedOptArg, Connection=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(684, LCID, 1, (9, 0), ((12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)),SourceType
			, SourceData, TableDestination, TableName, RowGrand, ColumnGrand
			, SaveData, HasAutoFormat, AutoPage, Reserved, BackgroundQuery
			, OptimizeCache, PageFieldOrder, PageFieldWrapCount, ReadData, Connection
			)
		if ret is not None:
			ret = Dispatch(ret, 'PivotTableWizard', '{00020872-0000-0000-C000-000000000046}')
		return ret

	def PivotTables(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(690, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'PivotTables', None)
		return ret

	def PrintOut(self, From=defaultNamedOptArg, To=defaultNamedOptArg, Copies=defaultNamedOptArg, Preview=defaultNamedOptArg
			, ActivePrinter=defaultNamedOptArg, PrintToFile=defaultNamedOptArg, Collate=defaultNamedOptArg, PrToFileName=defaultNamedOptArg, IgnorePrintAreas=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2361, LCID, 1, (24, 0), ((12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)),From
			, To, Copies, Preview, ActivePrinter, PrintToFile
			, Collate, PrToFileName, IgnorePrintAreas)

	def PrintPreview(self, EnableChanges=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(281, LCID, 1, (24, 0), ((12, 17),),EnableChanges
			)

	def Protect(self, Password=defaultNamedOptArg, DrawingObjects=defaultNamedOptArg, Contents=defaultNamedOptArg, Scenarios=defaultNamedOptArg
			, UserInterfaceOnly=defaultNamedOptArg, AllowFormattingCells=defaultNamedOptArg, AllowFormattingColumns=defaultNamedOptArg, AllowFormattingRows=defaultNamedOptArg, AllowInsertingColumns=defaultNamedOptArg
			, AllowInsertingRows=defaultNamedOptArg, AllowInsertingHyperlinks=defaultNamedOptArg, AllowDeletingColumns=defaultNamedOptArg, AllowDeletingRows=defaultNamedOptArg, AllowSorting=defaultNamedOptArg
			, AllowFiltering=defaultNamedOptArg, AllowUsingPivotTables=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2029, LCID, 1, (24, 0), ((12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)),Password
			, DrawingObjects, Contents, Scenarios, UserInterfaceOnly, AllowFormattingCells
			, AllowFormattingColumns, AllowFormattingRows, AllowInsertingColumns, AllowInsertingRows, AllowInsertingHyperlinks
			, AllowDeletingColumns, AllowDeletingRows, AllowSorting, AllowFiltering, AllowUsingPivotTables
			)

	# Result is of type Range
	# The method Range is actually a property, but must be used as a method to correctly pass the arguments
	def Range(self, Cell1=defaultNamedNotOptArg, Cell2=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(197, LCID, 2, (9, 0), ((12, 1), (12, 17)),Cell1
			, Cell2)
		if ret is not None:
			ret = Dispatch(ret, 'Range', '{00020846-0000-0000-C000-000000000046}')
		return ret

	def Rectangles(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(774, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'Rectangles', None)
		return ret

	def ResetAllPageBreaks(self):
		return self._oleobj_.InvokeTypes(1426, LCID, 1, (24, 0), (),)

	def SaveAs(self, Filename=defaultNamedNotOptArg, FileFormat=defaultNamedOptArg, Password=defaultNamedOptArg, WriteResPassword=defaultNamedOptArg
			, ReadOnlyRecommended=defaultNamedOptArg, CreateBackup=defaultNamedOptArg, AddToMru=defaultNamedOptArg, TextCodepage=defaultNamedOptArg, TextVisualLayout=defaultNamedOptArg
			, Local=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1925, LCID, 1, (24, 0), ((8, 1), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)),Filename
			, FileFormat, Password, WriteResPassword, ReadOnlyRecommended, CreateBackup
			, AddToMru, TextCodepage, TextVisualLayout, Local)

	def Scenarios(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(908, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'Scenarios', None)
		return ret

	def ScrollBars(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(830, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'ScrollBars', None)
		return ret

	def Select(self, Replace=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(235, LCID, 1, (24, 0), ((12, 17),),Replace
			)

	def SetBackgroundPicture(self, Filename=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(1188, LCID, 1, (24, 0), ((8, 1),),Filename
			)

	def ShowAllData(self):
		return self._oleobj_.InvokeTypes(794, LCID, 1, (24, 0), (),)

	def ShowDataForm(self):
		return self._oleobj_.InvokeTypes(409, LCID, 1, (24, 0), (),)

	def Spinners(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(838, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'Spinners', None)
		return ret

	def TextBoxes(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(777, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'TextBoxes', None)
		return ret

	def Unprotect(self, Password=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(285, LCID, 1, (24, 0), ((12, 17),),Password
			)

	# Result is of type Range
	def XmlDataQuery(self, XPath=defaultNamedNotOptArg, SelectionNamespaces=defaultNamedOptArg, Map=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(2260, LCID, 1, (9, 0), ((8, 1), (12, 17), (12, 17)),XPath
			, SelectionNamespaces, Map)
		if ret is not None:
			ret = Dispatch(ret, 'XmlDataQuery', '{00020846-0000-0000-C000-000000000046}')
		return ret

	# Result is of type Range
	def XmlMapQuery(self, XPath=defaultNamedNotOptArg, SelectionNamespaces=defaultNamedOptArg, Map=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(2263, LCID, 1, (9, 0), ((8, 1), (12, 17), (12, 17)),XPath
			, SelectionNamespaces, Map)
		if ret is not None:
			ret = Dispatch(ret, 'XmlMapQuery', '{00020846-0000-0000-C000-000000000046}')
		return ret

	def _CheckSpelling(self, CustomDictionary=defaultNamedOptArg, IgnoreUppercase=defaultNamedOptArg, AlwaysSuggest=defaultNamedOptArg, SpellLang=defaultNamedOptArg
			, IgnoreFinalYaa=defaultNamedOptArg, SpellScript=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1817, LCID, 1, (24, 0), ((12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)),CustomDictionary
			, IgnoreUppercase, AlwaysSuggest, SpellLang, IgnoreFinalYaa, SpellScript
			)

	def _Evaluate(self, Name=defaultNamedNotOptArg):
		return self._ApplyTypes_(-5, 1, (12, 0), ((12, 1),), '_Evaluate', None,Name
			)

	def _PasteSpecial(self, Format=defaultNamedOptArg, Link=defaultNamedOptArg, DisplayAsIcon=defaultNamedOptArg, IconFileName=defaultNamedOptArg
			, IconIndex=defaultNamedOptArg, IconLabel=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1027, LCID, 1, (24, 0), ((12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)),Format
			, Link, DisplayAsIcon, IconFileName, IconIndex, IconLabel
			)

	def _PrintOut(self, From=defaultNamedOptArg, To=defaultNamedOptArg, Copies=defaultNamedOptArg, Preview=defaultNamedOptArg
			, ActivePrinter=defaultNamedOptArg, PrintToFile=defaultNamedOptArg, Collate=defaultNamedOptArg, PrToFileName=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1772, LCID, 1, (24, 0), ((12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)),From
			, To, Copies, Preview, ActivePrinter, PrintToFile
			, Collate, PrToFileName)

	def _Protect(self, Password=defaultNamedOptArg, DrawingObjects=defaultNamedOptArg, Contents=defaultNamedOptArg, Scenarios=defaultNamedOptArg
			, UserInterfaceOnly=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(282, LCID, 1, (24, 0), ((12, 17), (12, 17), (12, 17), (12, 17), (12, 17)),Password
			, DrawingObjects, Contents, Scenarios, UserInterfaceOnly)

	def _SaveAs(self, Filename=defaultNamedNotOptArg, FileFormat=defaultNamedOptArg, Password=defaultNamedOptArg, WriteResPassword=defaultNamedOptArg
			, ReadOnlyRecommended=defaultNamedOptArg, CreateBackup=defaultNamedOptArg, AddToMru=defaultNamedOptArg, TextCodepage=defaultNamedOptArg, TextVisualLayout=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(284, LCID, 1, (24, 0), ((8, 1), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)),Filename
			, FileFormat, Password, WriteResPassword, ReadOnlyRecommended, CreateBackup
			, AddToMru, TextCodepage, TextVisualLayout)

	def _PrintOut_(self, From=defaultNamedOptArg, To=defaultNamedOptArg, Copies=defaultNamedOptArg, Preview=defaultNamedOptArg
			, ActivePrinter=defaultNamedOptArg, PrintToFile=defaultNamedOptArg, Collate=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(905, LCID, 1, (24, 0), ((12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)),From
			, To, Copies, Preview, ActivePrinter, PrintToFile
			, Collate)

	_prop_map_get_ = {
		# Method 'Application' returns object of type 'Application'
		"Application": (148, 2, (13, 0), (), "Application", '{00024500-0000-0000-C000-000000000046}'),
		# Method 'AutoFilter' returns object of type 'AutoFilter'
		"AutoFilter": (793, 2, (9, 0), (), "AutoFilter", '{00024432-0000-0000-C000-000000000046}'),
		"AutoFilterMode": (792, 2, (11, 0), (), "AutoFilterMode", None),
		# Method 'Cells' returns object of type 'Range'
		"Cells": (238, 2, (9, 0), (), "Cells", '{00020846-0000-0000-C000-000000000046}'),
		# Method 'CircularReference' returns object of type 'Range'
		"CircularReference": (1069, 2, (9, 0), (), "CircularReference", '{00020846-0000-0000-C000-000000000046}'),
		"CodeName": (1373, 2, (8, 0), (), "CodeName", None),
		# Method 'Columns' returns object of type 'Range'
		"Columns": (241, 2, (9, 0), (), "Columns", '{00020846-0000-0000-C000-000000000046}'),
		# Method 'Comments' returns object of type 'Comments'
		"Comments": (575, 2, (9, 0), (), "Comments", '{00024426-0000-0000-C000-000000000046}'),
		"ConsolidationFunction": (789, 2, (3, 0), (), "ConsolidationFunction", None),
		"ConsolidationOptions": (790, 2, (12, 0), (), "ConsolidationOptions", None),
		"ConsolidationSources": (791, 2, (12, 0), (), "ConsolidationSources", None),
		"Creator": (149, 2, (3, 0), (), "Creator", None),
		# Method 'CustomProperties' returns object of type 'CustomProperties'
		"CustomProperties": (2030, 2, (9, 0), (), "CustomProperties", '{00024452-0000-0000-C000-000000000046}'),
		"DisplayAutomaticPageBreaks": (643, 2, (11, 0), (), "DisplayAutomaticPageBreaks", None),
		"DisplayPageBreaks": (1435, 2, (11, 0), (), "DisplayPageBreaks", None),
		"DisplayRightToLeft": (1774, 2, (11, 0), (), "DisplayRightToLeft", None),
		"EnableAutoFilter": (1156, 2, (11, 0), (), "EnableAutoFilter", None),
		"EnableCalculation": (1424, 2, (11, 0), (), "EnableCalculation", None),
		"EnableFormatConditionsCalculation": (2511, 2, (11, 0), (), "EnableFormatConditionsCalculation", None),
		"EnableOutlining": (1157, 2, (11, 0), (), "EnableOutlining", None),
		"EnablePivotTable": (1158, 2, (11, 0), (), "EnablePivotTable", None),
		"EnableSelection": (1425, 2, (3, 0), (), "EnableSelection", None),
		"FilterMode": (800, 2, (11, 0), (), "FilterMode", None),
		# Method 'HPageBreaks' returns object of type 'HPageBreaks'
		"HPageBreaks": (1418, 2, (9, 0), (), "HPageBreaks", '{00024404-0000-0000-C000-000000000046}'),
		# Method 'Hyperlinks' returns object of type 'Hyperlinks'
		"Hyperlinks": (1393, 2, (9, 0), (), "Hyperlinks", '{00024430-0000-0000-C000-000000000046}'),
		"Index": (486, 2, (3, 0), (), "Index", None),
		# Method 'ListObjects' returns object of type 'ListObjects'
		"ListObjects": (2259, 2, (9, 0), (), "ListObjects", '{00024470-0000-0000-C000-000000000046}'),
		# Method 'MailEnvelope' returns object of type 'MsoEnvelope'
		"MailEnvelope": (2021, 2, (13, 0), (), "MailEnvelope", '{0006F01A-0000-0000-C000-000000000046}'),
		"Name": (110, 2, (8, 0), (), "Name", None),
		# Method 'Names' returns object of type 'Names'
		"Names": (442, 2, (9, 0), (), "Names", '{000208B8-0000-0000-C000-000000000046}'),
		"Next": (502, 2, (9, 0), (), "Next", None),
		"OnCalculate": (625, 2, (8, 0), (), "OnCalculate", None),
		"OnData": (629, 2, (8, 0), (), "OnData", None),
		"OnDoubleClick": (628, 2, (8, 0), (), "OnDoubleClick", None),
		"OnEntry": (627, 2, (8, 0), (), "OnEntry", None),
		"OnSheetActivate": (1031, 2, (8, 0), (), "OnSheetActivate", None),
		"OnSheetDeactivate": (1081, 2, (8, 0), (), "OnSheetDeactivate", None),
		# Method 'Outline' returns object of type 'Outline'
		"Outline": (102, 2, (9, 0), (), "Outline", '{000208AB-0000-0000-C000-000000000046}'),
		# Method 'PageSetup' returns object of type 'PageSetup'
		"PageSetup": (998, 2, (9, 0), (), "PageSetup", '{000208B4-0000-0000-C000-000000000046}'),
		"Parent": (150, 2, (9, 0), (), "Parent", None),
		"Previous": (503, 2, (9, 0), (), "Previous", None),
		"PrintedCommentPages": (2857, 2, (3, 0), (), "PrintedCommentPages", None),
		"ProtectContents": (292, 2, (11, 0), (), "ProtectContents", None),
		"ProtectDrawingObjects": (293, 2, (11, 0), (), "ProtectDrawingObjects", None),
		"ProtectScenarios": (294, 2, (11, 0), (), "ProtectScenarios", None),
		# Method 'Protection' returns object of type 'Protection'
		"Protection": (176, 2, (9, 0), (), "Protection", '{00024467-0000-0000-C000-000000000046}'),
		"ProtectionMode": (1159, 2, (11, 0), (), "ProtectionMode", None),
		# Method 'QueryTables' returns object of type 'QueryTables'
		"QueryTables": (1434, 2, (9, 0), (), "QueryTables", '{00024429-0000-0000-C000-000000000046}'),
		# Method 'Rows' returns object of type 'Range'
		"Rows": (258, 2, (9, 0), (), "Rows", '{00020846-0000-0000-C000-000000000046}'),
		# Method 'Scripts' returns object of type 'Scripts'
		"Scripts": (1816, 2, (9, 0), (), "Scripts", '{000C0340-0000-0000-C000-000000000046}'),
		"ScrollArea": (1433, 2, (8, 0), (), "ScrollArea", None),
		# Method 'Shapes' returns object of type 'Shapes'
		"Shapes": (1377, 2, (9, 0), (), "Shapes", '{0002443A-0000-0000-C000-000000000046}'),
		# Method 'SmartTags' returns object of type 'SmartTags'
		"SmartTags": (2016, 2, (9, 0), (), "SmartTags", '{00024461-0000-0000-C000-000000000046}'),
		# Method 'Sort' returns object of type 'Sort'
		"Sort": (880, 2, (9, 0), (), "Sort", '{000244AB-0000-0000-C000-000000000046}'),
		"StandardHeight": (407, 2, (5, 0), (), "StandardHeight", None),
		"StandardWidth": (408, 2, (5, 0), (), "StandardWidth", None),
		# Method 'Tab' returns object of type 'Tab'
		"Tab": (1041, 2, (9, 0), (), "Tab", '{00024469-0000-0000-C000-000000000046}'),
		"TransitionExpEval": (401, 2, (11, 0), (), "TransitionExpEval", None),
		"TransitionFormEntry": (402, 2, (11, 0), (), "TransitionFormEntry", None),
		"Type": (108, 2, (3, 0), (), "Type", None),
		# Method 'UsedRange' returns object of type 'Range'
		"UsedRange": (412, 2, (9, 0), (), "UsedRange", '{00020846-0000-0000-C000-000000000046}'),
		# Method 'VPageBreaks' returns object of type 'VPageBreaks'
		"VPageBreaks": (1419, 2, (9, 0), (), "VPageBreaks", '{00024405-0000-0000-C000-000000000046}'),
		"Visible": (558, 2, (3, 0), (), "Visible", None),
		"_CodeName": (-2147418112, 2, (8, 0), (), "_CodeName", None),
		"_DisplayRightToLeft": (648, 2, (3, 0), (), "_DisplayRightToLeft", None),
	}
	_prop_map_put_ = {
		"AutoFilterMode": ((792, LCID, 4, 0),()),
		"DisplayAutomaticPageBreaks": ((643, LCID, 4, 0),()),
		"DisplayPageBreaks": ((1435, LCID, 4, 0),()),
		"DisplayRightToLeft": ((1774, LCID, 4, 0),()),
		"EnableAutoFilter": ((1156, LCID, 4, 0),()),
		"EnableCalculation": ((1424, LCID, 4, 0),()),
		"EnableFormatConditionsCalculation": ((2511, LCID, 4, 0),()),
		"EnableOutlining": ((1157, LCID, 4, 0),()),
		"EnablePivotTable": ((1158, LCID, 4, 0),()),
		"EnableSelection": ((1425, LCID, 4, 0),()),
		"Name": ((110, LCID, 4, 0),()),
		"OnCalculate": ((625, LCID, 4, 0),()),
		"OnData": ((629, LCID, 4, 0),()),
		"OnDoubleClick": ((628, LCID, 4, 0),()),
		"OnEntry": ((627, LCID, 4, 0),()),
		"OnSheetActivate": ((1031, LCID, 4, 0),()),
		"OnSheetDeactivate": ((1081, LCID, 4, 0),()),
		"ScrollArea": ((1433, LCID, 4, 0),()),
		"StandardWidth": ((408, LCID, 4, 0),()),
		"TransitionExpEval": ((401, LCID, 4, 0),()),
		"TransitionFormEntry": ((402, LCID, 4, 0),()),
		"Visible": ((558, LCID, 4, 0),()),
		"_CodeName": ((-2147418112, LCID, 4, 0),()),
		"_DisplayRightToLeft": ((648, LCID, 4, 0),()),
	}
	def __iter__(self):
		"Return a Python iterator for this object"
		try:
			ob = self._oleobj_.InvokeTypes(-4,LCID,3,(13, 10),())
		except pythoncom.error:
			raise TypeError("This object does not support enumeration")
		return win32com.client.util.Iterator(ob, None)

win32com.client.CLSIDToClass.RegisterCLSID( "{000208D8-0000-0000-C000-000000000046}", _Worksheet )
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

_Worksheet_vtables_dispatch_ = 1
_Worksheet_vtables_ = [
	(( 'Application' , 'RHS' , ), 148, (148, (), [ (16397, 10, None, "IID('{00024500-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 28 , (3, 0, None, None) , 0 , )),
	(( 'Creator' , 'RHS' , ), 149, (149, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 32 , (3, 0, None, None) , 0 , )),
	(( 'Parent' , 'RHS' , ), 150, (150, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 36 , (3, 0, None, None) , 0 , )),
	(( 'Activate' , 'lcid' , ), 304, (304, (), [ (3, 5, None, None) , ], 1 , 1 , 4 , 0 , 40 , (3, 0, None, None) , 0 , )),
	(( 'Copy' , 'Before' , 'After' , 'lcid' , ), 551, (551, (), [ 
			 (12, 17, None, None) , (12, 17, None, None) , (3, 5, None, None) , ], 1 , 1 , 4 , 2 , 44 , (3, 0, None, None) , 0 , )),
	(( 'Delete' , 'lcid' , ), 117, (117, (), [ (3, 5, None, None) , ], 1 , 1 , 4 , 0 , 48 , (3, 0, None, None) , 0 , )),
	(( 'CodeName' , 'RHS' , ), 1373, (1373, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 52 , (3, 0, None, None) , 0 , )),
	(( '_CodeName' , 'RHS' , ), -2147418112, (-2147418112, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 56 , (3, 0, None, None) , 1024 , )),
	(( '_CodeName' , 'RHS' , ), -2147418112, (-2147418112, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 60 , (3, 0, None, None) , 1024 , )),
	(( 'Index' , 'lcid' , 'RHS' , ), 486, (486, (), [ (3, 5, None, None) , 
			 (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 64 , (3, 0, None, None) , 0 , )),
	(( 'Move' , 'Before' , 'After' , 'lcid' , ), 637, (637, (), [ 
			 (12, 17, None, None) , (12, 17, None, None) , (3, 5, None, None) , ], 1 , 1 , 4 , 2 , 68 , (3, 0, None, None) , 0 , )),
	(( 'Name' , 'RHS' , ), 110, (110, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 72 , (3, 0, None, None) , 0 , )),
	(( 'Name' , 'RHS' , ), 110, (110, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 76 , (3, 0, None, None) , 0 , )),
	(( 'Next' , 'RHS' , ), 502, (502, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 80 , (3, 0, None, None) , 0 , )),
	(( 'OnDoubleClick' , 'lcid' , 'RHS' , ), 628, (628, (), [ (3, 5, None, None) , 
			 (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 84 , (3, 0, None, None) , 64 , )),
	(( 'OnDoubleClick' , 'lcid' , 'RHS' , ), 628, (628, (), [ (3, 5, None, None) , 
			 (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 88 , (3, 0, None, None) , 64 , )),
	(( 'OnSheetActivate' , 'lcid' , 'RHS' , ), 1031, (1031, (), [ (3, 5, None, None) , 
			 (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 92 , (3, 0, None, None) , 64 , )),
	(( 'OnSheetActivate' , 'lcid' , 'RHS' , ), 1031, (1031, (), [ (3, 5, None, None) , 
			 (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 96 , (3, 0, None, None) , 64 , )),
	(( 'OnSheetDeactivate' , 'lcid' , 'RHS' , ), 1081, (1081, (), [ (3, 5, None, None) , 
			 (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 100 , (3, 0, None, None) , 64 , )),
	(( 'OnSheetDeactivate' , 'lcid' , 'RHS' , ), 1081, (1081, (), [ (3, 5, None, None) , 
			 (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 104 , (3, 0, None, None) , 64 , )),
	(( 'PageSetup' , 'RHS' , ), 998, (998, (), [ (16393, 10, None, "IID('{000208B4-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 108 , (3, 0, None, None) , 0 , )),
	(( 'Previous' , 'RHS' , ), 503, (503, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 112 , (3, 0, None, None) , 0 , )),
	(( '__PrintOut' , 'From' , 'To' , 'Copies' , 'Preview' , 
			 'ActivePrinter' , 'PrintToFile' , 'Collate' , 'lcid' , ), 905, (905, (), [ 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (3, 5, None, None) , ], 1 , 1 , 4 , 7 , 116 , (3, 0, None, None) , 1088 , )),
	(( 'PrintPreview' , 'EnableChanges' , 'lcid' , ), 281, (281, (), [ (12, 17, None, None) , 
			 (3, 5, None, None) , ], 1 , 1 , 4 , 1 , 120 , (3, 0, None, None) , 0 , )),
	(( '_Protect' , 'Password' , 'DrawingObjects' , 'Contents' , 'Scenarios' , 
			 'UserInterfaceOnly' , 'lcid' , ), 282, (282, (), [ (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (3, 5, None, None) , ], 1 , 1 , 4 , 5 , 124 , (3, 0, None, None) , 1088 , )),
	(( 'ProtectContents' , 'lcid' , 'RHS' , ), 292, (292, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 128 , (3, 0, None, None) , 0 , )),
	(( 'ProtectDrawingObjects' , 'lcid' , 'RHS' , ), 293, (293, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 132 , (3, 0, None, None) , 0 , )),
	(( 'ProtectionMode' , 'lcid' , 'RHS' , ), 1159, (1159, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 136 , (3, 0, None, None) , 0 , )),
	(( 'ProtectScenarios' , 'lcid' , 'RHS' , ), 294, (294, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 140 , (3, 0, None, None) , 0 , )),
	(( '_SaveAs' , 'Filename' , 'FileFormat' , 'Password' , 'WriteResPassword' , 
			 'ReadOnlyRecommended' , 'CreateBackup' , 'AddToMru' , 'TextCodepage' , 'TextVisualLayout' , 
			 'lcid' , ), 284, (284, (), [ (8, 1, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (3, 5, None, None) , ], 1 , 1 , 4 , 8 , 144 , (3, 0, None, None) , 1088 , )),
	(( 'Select' , 'Replace' , 'lcid' , ), 235, (235, (), [ (12, 17, None, None) , 
			 (3, 5, None, None) , ], 1 , 1 , 4 , 1 , 148 , (3, 0, None, None) , 0 , )),
	(( 'Unprotect' , 'Password' , 'lcid' , ), 285, (285, (), [ (12, 17, None, None) , 
			 (3, 5, None, None) , ], 1 , 1 , 4 , 1 , 152 , (3, 0, None, None) , 0 , )),
	(( 'Visible' , 'lcid' , 'RHS' , ), 558, (558, (), [ (3, 5, None, None) , 
			 (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 156 , (3, 0, None, None) , 0 , )),
	(( 'Visible' , 'lcid' , 'RHS' , ), 558, (558, (), [ (3, 5, None, None) , 
			 (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 160 , (3, 0, None, None) , 0 , )),
	(( 'Shapes' , 'RHS' , ), 1377, (1377, (), [ (16393, 10, None, "IID('{0002443A-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 164 , (3, 0, None, None) , 0 , )),
	(( 'TransitionExpEval' , 'lcid' , 'RHS' , ), 401, (401, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 168 , (3, 0, None, None) , 1024 , )),
	(( 'TransitionExpEval' , 'lcid' , 'RHS' , ), 401, (401, (), [ (3, 5, None, None) , 
			 (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 172 , (3, 0, None, None) , 1024 , )),
	(( 'Arcs' , 'Index' , 'lcid' , 'RHS' , ), 760, (760, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 1 , 176 , (3, 0, None, None) , 64 , )),
	(( 'AutoFilterMode' , 'lcid' , 'RHS' , ), 792, (792, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 180 , (3, 0, None, None) , 1024 , )),
	(( 'AutoFilterMode' , 'lcid' , 'RHS' , ), 792, (792, (), [ (3, 5, None, None) , 
			 (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 184 , (3, 0, None, None) , 1024 , )),
	(( 'SetBackgroundPicture' , 'Filename' , ), 1188, (1188, (), [ (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 188 , (3, 0, None, None) , 0 , )),
	(( 'Buttons' , 'Index' , 'lcid' , 'RHS' , ), 557, (557, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 1 , 192 , (3, 0, None, None) , 64 , )),
	(( 'Calculate' , 'lcid' , ), 279, (279, (), [ (3, 5, None, None) , ], 1 , 1 , 4 , 0 , 196 , (3, 0, None, None) , 0 , )),
	(( 'EnableCalculation' , 'RHS' , ), 1424, (1424, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 200 , (3, 0, None, None) , 0 , )),
	(( 'EnableCalculation' , 'RHS' , ), 1424, (1424, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 204 , (3, 0, None, None) , 0 , )),
	(( 'Cells' , 'RHS' , ), 238, (238, (), [ (16393, 10, None, "IID('{00020846-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 208 , (3, 0, None, None) , 0 , )),
	(( 'ChartObjects' , 'Index' , 'lcid' , 'RHS' , ), 1060, (1060, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 1 , 212 , (3, 0, None, None) , 0 , )),
	(( 'CheckBoxes' , 'Index' , 'lcid' , 'RHS' , ), 824, (824, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 1 , 216 , (3, 0, None, None) , 64 , )),
	(( 'CheckSpelling' , 'CustomDictionary' , 'IgnoreUppercase' , 'AlwaysSuggest' , 'SpellLang' , 
			 'lcid' , ), 505, (505, (), [ (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (3, 5, None, None) , ], 1 , 1 , 4 , 4 , 220 , (3, 0, None, None) , 0 , )),
	(( 'CircularReference' , 'lcid' , 'RHS' , ), 1069, (1069, (), [ (3, 5, None, None) , 
			 (16393, 10, None, "IID('{00020846-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 224 , (3, 0, None, None) , 0 , )),
	(( 'ClearArrows' , 'lcid' , ), 970, (970, (), [ (3, 5, None, None) , ], 1 , 1 , 4 , 0 , 228 , (3, 0, None, None) , 0 , )),
	(( 'Columns' , 'RHS' , ), 241, (241, (), [ (16393, 10, None, "IID('{00020846-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 232 , (3, 0, None, None) , 1024 , )),
	(( 'ConsolidationFunction' , 'lcid' , 'RHS' , ), 789, (789, (), [ (3, 5, None, None) , 
			 (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 236 , (3, 0, None, None) , 0 , )),
	(( 'ConsolidationOptions' , 'lcid' , 'RHS' , ), 790, (790, (), [ (3, 5, None, None) , 
			 (16396, 10, None, None) , ], 1 , 2 , 4 , 0 , 240 , (3, 0, None, None) , 0 , )),
	(( 'ConsolidationSources' , 'lcid' , 'RHS' , ), 791, (791, (), [ (3, 5, None, None) , 
			 (16396, 10, None, None) , ], 1 , 2 , 4 , 0 , 244 , (3, 0, None, None) , 0 , )),
	(( 'DisplayAutomaticPageBreaks' , 'lcid' , 'RHS' , ), 643, (643, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 248 , (3, 0, None, None) , 64 , )),
	(( 'DisplayAutomaticPageBreaks' , 'lcid' , 'RHS' , ), 643, (643, (), [ (3, 5, None, None) , 
			 (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 252 , (3, 0, None, None) , 64 , )),
	(( 'Drawings' , 'Index' , 'lcid' , 'RHS' , ), 772, (772, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 1 , 256 , (3, 0, None, None) , 64 , )),
	(( 'DrawingObjects' , 'Index' , 'lcid' , 'RHS' , ), 88, (88, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 1 , 260 , (3, 0, None, None) , 64 , )),
	(( 'DropDowns' , 'Index' , 'lcid' , 'RHS' , ), 836, (836, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 1 , 264 , (3, 0, None, None) , 64 , )),
	(( 'EnableAutoFilter' , 'lcid' , 'RHS' , ), 1156, (1156, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 268 , (3, 0, None, None) , 0 , )),
	(( 'EnableAutoFilter' , 'lcid' , 'RHS' , ), 1156, (1156, (), [ (3, 5, None, None) , 
			 (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 272 , (3, 0, None, None) , 0 , )),
	(( 'EnableSelection' , 'RHS' , ), 1425, (1425, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 276 , (3, 0, None, None) , 0 , )),
	(( 'EnableSelection' , 'RHS' , ), 1425, (1425, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 280 , (3, 0, None, None) , 0 , )),
	(( 'EnableOutlining' , 'lcid' , 'RHS' , ), 1157, (1157, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 284 , (3, 0, None, None) , 0 , )),
	(( 'EnableOutlining' , 'lcid' , 'RHS' , ), 1157, (1157, (), [ (3, 5, None, None) , 
			 (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 288 , (3, 0, None, None) , 0 , )),
	(( 'EnablePivotTable' , 'lcid' , 'RHS' , ), 1158, (1158, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 292 , (3, 0, None, None) , 0 , )),
	(( 'EnablePivotTable' , 'lcid' , 'RHS' , ), 1158, (1158, (), [ (3, 5, None, None) , 
			 (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 296 , (3, 0, None, None) , 0 , )),
	(( 'Evaluate' , 'Name' , 'lcid' , 'RHS' , ), 1, (1, (), [ 
			 (12, 1, None, None) , (3, 5, None, None) , (16396, 10, None, None) , ], 1 , 1 , 4 , 0 , 300 , (3, 0, None, None) , 0 , )),
	(( '_Evaluate' , 'Name' , 'lcid' , 'RHS' , ), -5, (-5, (), [ 
			 (12, 1, None, None) , (3, 5, None, None) , (16396, 10, None, None) , ], 1 , 1 , 4 , 0 , 304 , (3, 0, None, None) , 1024 , )),
	(( 'FilterMode' , 'lcid' , 'RHS' , ), 800, (800, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 308 , (3, 0, None, None) , 0 , )),
	(( 'ResetAllPageBreaks' , ), 1426, (1426, (), [ ], 1 , 1 , 4 , 0 , 312 , (3, 0, None, None) , 0 , )),
	(( 'GroupBoxes' , 'Index' , 'lcid' , 'RHS' , ), 834, (834, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 1 , 316 , (3, 0, None, None) , 64 , )),
	(( 'GroupObjects' , 'Index' , 'lcid' , 'RHS' , ), 1113, (1113, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 1 , 320 , (3, 0, None, None) , 64 , )),
	(( 'Labels' , 'Index' , 'lcid' , 'RHS' , ), 841, (841, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 1 , 324 , (3, 0, None, None) , 64 , )),
	(( 'Lines' , 'Index' , 'lcid' , 'RHS' , ), 767, (767, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 1 , 328 , (3, 0, None, None) , 64 , )),
	(( 'ListBoxes' , 'Index' , 'lcid' , 'RHS' , ), 832, (832, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 1 , 332 , (3, 0, None, None) , 64 , )),
	(( 'Names' , 'RHS' , ), 442, (442, (), [ (16393, 10, None, "IID('{000208B8-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 336 , (3, 0, None, None) , 0 , )),
	(( 'OLEObjects' , 'Index' , 'lcid' , 'RHS' , ), 799, (799, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 1 , 340 , (3, 0, None, None) , 0 , )),
	(( 'OnCalculate' , 'lcid' , 'RHS' , ), 625, (625, (), [ (3, 5, None, None) , 
			 (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 344 , (3, 0, None, None) , 64 , )),
	(( 'OnCalculate' , 'lcid' , 'RHS' , ), 625, (625, (), [ (3, 5, None, None) , 
			 (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 348 , (3, 0, None, None) , 64 , )),
	(( 'OnData' , 'lcid' , 'RHS' , ), 629, (629, (), [ (3, 5, None, None) , 
			 (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 352 , (3, 0, None, None) , 64 , )),
	(( 'OnData' , 'lcid' , 'RHS' , ), 629, (629, (), [ (3, 5, None, None) , 
			 (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 356 , (3, 0, None, None) , 64 , )),
	(( 'OnEntry' , 'lcid' , 'RHS' , ), 627, (627, (), [ (3, 5, None, None) , 
			 (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 360 , (3, 0, None, None) , 64 , )),
	(( 'OnEntry' , 'lcid' , 'RHS' , ), 627, (627, (), [ (3, 5, None, None) , 
			 (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 364 , (3, 0, None, None) , 64 , )),
	(( 'OptionButtons' , 'Index' , 'lcid' , 'RHS' , ), 826, (826, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 1 , 368 , (3, 0, None, None) , 64 , )),
	(( 'Outline' , 'RHS' , ), 102, (102, (), [ (16393, 10, None, "IID('{000208AB-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 372 , (3, 0, None, None) , 0 , )),
	(( 'Ovals' , 'Index' , 'lcid' , 'RHS' , ), 801, (801, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 1 , 376 , (3, 0, None, None) , 64 , )),
	(( 'Paste' , 'Destination' , 'Link' , 'lcid' , ), 211, (211, (), [ 
			 (12, 17, None, None) , (12, 17, None, None) , (3, 5, None, None) , ], 1 , 1 , 4 , 2 , 380 , (3, 0, None, None) , 0 , )),
	(( '_PasteSpecial' , 'Format' , 'Link' , 'DisplayAsIcon' , 'IconFileName' , 
			 'IconIndex' , 'IconLabel' , 'lcid' , ), 1027, (1027, (), [ (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (3, 5, None, None) , ], 1 , 1 , 4 , 6 , 384 , (3, 0, None, None) , 1088 , )),
	(( 'Pictures' , 'Index' , 'lcid' , 'RHS' , ), 771, (771, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 1 , 388 , (3, 0, None, None) , 64 , )),
	(( 'PivotTables' , 'Index' , 'lcid' , 'RHS' , ), 690, (690, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 1 , 392 , (3, 0, None, None) , 0 , )),
	(( 'PivotTableWizard' , 'SourceType' , 'SourceData' , 'TableDestination' , 'TableName' , 
			 'RowGrand' , 'ColumnGrand' , 'SaveData' , 'HasAutoFormat' , 'AutoPage' , 
			 'Reserved' , 'BackgroundQuery' , 'OptimizeCache' , 'PageFieldOrder' , 'PageFieldWrapCount' , 
			 'ReadData' , 'Connection' , 'lcid' , 'RHS' , ), 684, (684, (), [ 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, "IID('{00020872-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 16 , 396 , (3, 0, None, None) , 0 , )),
	(( 'Range' , 'Cell1' , 'Cell2' , 'RHS' , ), 197, (197, (), [ 
			 (12, 1, None, None) , (12, 17, None, None) , (16393, 10, None, "IID('{00020846-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 1 , 400 , (3, 0, None, None) , 0 , )),
	(( 'Rectangles' , 'Index' , 'lcid' , 'RHS' , ), 774, (774, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 1 , 404 , (3, 0, None, None) , 64 , )),
	(( 'Rows' , 'RHS' , ), 258, (258, (), [ (16393, 10, None, "IID('{00020846-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 408 , (3, 0, None, None) , 1024 , )),
	(( 'Scenarios' , 'Index' , 'lcid' , 'RHS' , ), 908, (908, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 1 , 412 , (3, 0, None, None) , 0 , )),
	(( 'ScrollArea' , 'RHS' , ), 1433, (1433, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 416 , (3, 0, None, None) , 0 , )),
	(( 'ScrollArea' , 'RHS' , ), 1433, (1433, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 420 , (3, 0, None, None) , 0 , )),
	(( 'ScrollBars' , 'Index' , 'lcid' , 'RHS' , ), 830, (830, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 1 , 424 , (3, 0, None, None) , 64 , )),
	(( 'ShowAllData' , 'lcid' , ), 794, (794, (), [ (3, 5, None, None) , ], 1 , 1 , 4 , 0 , 428 , (3, 0, None, None) , 0 , )),
	(( 'ShowDataForm' , 'lcid' , ), 409, (409, (), [ (3, 5, None, None) , ], 1 , 1 , 4 , 0 , 432 , (3, 0, None, None) , 0 , )),
	(( 'Spinners' , 'Index' , 'lcid' , 'RHS' , ), 838, (838, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 1 , 436 , (3, 0, None, None) , 64 , )),
	(( 'StandardHeight' , 'lcid' , 'RHS' , ), 407, (407, (), [ (3, 5, None, None) , 
			 (16389, 10, None, None) , ], 1 , 2 , 4 , 0 , 440 , (3, 0, None, None) , 0 , )),
	(( 'StandardWidth' , 'lcid' , 'RHS' , ), 408, (408, (), [ (3, 5, None, None) , 
			 (16389, 10, None, None) , ], 1 , 2 , 4 , 0 , 444 , (3, 0, None, None) , 0 , )),
	(( 'StandardWidth' , 'lcid' , 'RHS' , ), 408, (408, (), [ (3, 5, None, None) , 
			 (5, 1, None, None) , ], 1 , 4 , 4 , 0 , 448 , (3, 0, None, None) , 0 , )),
	(( 'TextBoxes' , 'Index' , 'lcid' , 'RHS' , ), 777, (777, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 1 , 452 , (3, 0, None, None) , 64 , )),
	(( 'TransitionFormEntry' , 'lcid' , 'RHS' , ), 402, (402, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 456 , (3, 0, None, None) , 1024 , )),
	(( 'TransitionFormEntry' , 'lcid' , 'RHS' , ), 402, (402, (), [ (3, 5, None, None) , 
			 (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 460 , (3, 0, None, None) , 1024 , )),
	(( 'Type' , 'lcid' , 'RHS' , ), 108, (108, (), [ (3, 5, None, None) , 
			 (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 464 , (3, 0, None, None) , 0 , )),
	(( 'UsedRange' , 'lcid' , 'RHS' , ), 412, (412, (), [ (3, 5, None, None) , 
			 (16393, 10, None, "IID('{00020846-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 468 , (3, 0, None, None) , 0 , )),
	(( 'HPageBreaks' , 'RHS' , ), 1418, (1418, (), [ (16393, 10, None, "IID('{00024404-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 472 , (3, 0, None, None) , 0 , )),
	(( 'VPageBreaks' , 'RHS' , ), 1419, (1419, (), [ (16393, 10, None, "IID('{00024405-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 476 , (3, 0, None, None) , 0 , )),
	(( 'QueryTables' , 'RHS' , ), 1434, (1434, (), [ (16393, 10, None, "IID('{00024429-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 480 , (3, 0, None, None) , 0 , )),
	(( 'DisplayPageBreaks' , 'RHS' , ), 1435, (1435, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 484 , (3, 0, None, None) , 0 , )),
	(( 'DisplayPageBreaks' , 'RHS' , ), 1435, (1435, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 488 , (3, 0, None, None) , 0 , )),
	(( 'Comments' , 'RHS' , ), 575, (575, (), [ (16393, 10, None, "IID('{00024426-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 492 , (3, 0, None, None) , 0 , )),
	(( 'Hyperlinks' , 'RHS' , ), 1393, (1393, (), [ (16393, 10, None, "IID('{00024430-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 496 , (3, 0, None, None) , 0 , )),
	(( 'ClearCircles' , ), 1436, (1436, (), [ ], 1 , 1 , 4 , 0 , 500 , (3, 0, None, None) , 0 , )),
	(( 'CircleInvalid' , ), 1437, (1437, (), [ ], 1 , 1 , 4 , 0 , 504 , (3, 0, None, None) , 0 , )),
	(( '_DisplayRightToLeft' , 'lcid' , 'RHS' , ), 648, (648, (), [ (3, 5, None, None) , 
			 (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 508 , (3, 0, None, None) , 1088 , )),
	(( '_DisplayRightToLeft' , 'lcid' , 'RHS' , ), 648, (648, (), [ (3, 5, None, None) , 
			 (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 512 , (3, 0, None, None) , 1088 , )),
	(( 'AutoFilter' , 'RHS' , ), 793, (793, (), [ (16393, 10, None, "IID('{00024432-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 516 , (3, 0, None, None) , 0 , )),
	(( 'DisplayRightToLeft' , 'lcid' , 'RHS' , ), 1774, (1774, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 520 , (3, 0, None, None) , 0 , )),
	(( 'DisplayRightToLeft' , 'lcid' , 'RHS' , ), 1774, (1774, (), [ (3, 5, None, None) , 
			 (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 524 , (3, 0, None, None) , 0 , )),
	(( 'Scripts' , 'RHS' , ), 1816, (1816, (), [ (16393, 10, None, "IID('{000C0340-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 528 , (3, 0, None, None) , 64 , )),
	(( '_PrintOut' , 'From' , 'To' , 'Copies' , 'Preview' , 
			 'ActivePrinter' , 'PrintToFile' , 'Collate' , 'PrToFileName' , 'lcid' , 
			 ), 1772, (1772, (), [ (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (3, 5, None, None) , ], 1 , 1 , 4 , 8 , 532 , (3, 0, None, None) , 1088 , )),
	(( '_CheckSpelling' , 'CustomDictionary' , 'IgnoreUppercase' , 'AlwaysSuggest' , 'SpellLang' , 
			 'IgnoreFinalYaa' , 'SpellScript' , 'lcid' , ), 1817, (1817, (), [ (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (3, 5, None, None) , ], 1 , 1 , 4 , 6 , 536 , (3, 0, None, None) , 1088 , )),
	(( 'Tab' , 'RHS' , ), 1041, (1041, (), [ (16393, 10, None, "IID('{00024469-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 540 , (3, 0, None, None) , 0 , )),
	(( 'MailEnvelope' , 'RHS' , ), 2021, (2021, (), [ (16397, 10, None, "IID('{0006F01A-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 544 , (3, 0, None, None) , 0 , )),
	(( 'SaveAs' , 'Filename' , 'FileFormat' , 'Password' , 'WriteResPassword' , 
			 'ReadOnlyRecommended' , 'CreateBackup' , 'AddToMru' , 'TextCodepage' , 'TextVisualLayout' , 
			 'Local' , ), 1925, (1925, (), [ (8, 1, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , ], 1 , 1 , 4 , 9 , 548 , (3, 0, None, None) , 0 , )),
	(( 'CustomProperties' , 'RHS' , ), 2030, (2030, (), [ (16393, 10, None, "IID('{00024452-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 552 , (3, 0, None, None) , 0 , )),
	(( 'SmartTags' , 'RHS' , ), 2016, (2016, (), [ (16393, 10, None, "IID('{00024461-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 556 , (3, 0, None, None) , 64 , )),
	(( 'Protection' , 'RHS' , ), 176, (176, (), [ (16393, 10, None, "IID('{00024467-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 560 , (3, 0, None, None) , 0 , )),
	(( 'PasteSpecial' , 'Format' , 'Link' , 'DisplayAsIcon' , 'IconFileName' , 
			 'IconIndex' , 'IconLabel' , 'NoHTMLFormatting' , 'lcid' , ), 1928, (1928, (), [ 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (3, 5, None, None) , ], 1 , 1 , 4 , 7 , 564 , (3, 0, None, None) , 0 , )),
	(( 'Protect' , 'Password' , 'DrawingObjects' , 'Contents' , 'Scenarios' , 
			 'UserInterfaceOnly' , 'AllowFormattingCells' , 'AllowFormattingColumns' , 'AllowFormattingRows' , 'AllowInsertingColumns' , 
			 'AllowInsertingRows' , 'AllowInsertingHyperlinks' , 'AllowDeletingColumns' , 'AllowDeletingRows' , 'AllowSorting' , 
			 'AllowFiltering' , 'AllowUsingPivotTables' , ), 2029, (2029, (), [ (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , ], 1 , 1 , 4 , 16 , 568 , (3, 0, None, None) , 0 , )),
	(( 'ListObjects' , 'RHS' , ), 2259, (2259, (), [ (16393, 10, None, "IID('{00024470-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 572 , (3, 0, None, None) , 0 , )),
	(( 'XmlDataQuery' , 'XPath' , 'SelectionNamespaces' , 'Map' , 'RHS' , 
			 ), 2260, (2260, (), [ (8, 1, None, None) , (12, 17, None, None) , (12, 17, None, None) , (16393, 10, None, "IID('{00020846-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 2 , 576 , (3, 0, None, None) , 0 , )),
	(( 'XmlMapQuery' , 'XPath' , 'SelectionNamespaces' , 'Map' , 'RHS' , 
			 ), 2263, (2263, (), [ (8, 1, None, None) , (12, 17, None, None) , (12, 17, None, None) , (16393, 10, None, "IID('{00020846-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 2 , 580 , (3, 0, None, None) , 0 , )),
	(( 'PrintOut' , 'From' , 'To' , 'Copies' , 'Preview' , 
			 'ActivePrinter' , 'PrintToFile' , 'Collate' , 'PrToFileName' , 'IgnorePrintAreas' , 
			 'lcid' , ), 2361, (2361, (), [ (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (3, 5, None, None) , ], 1 , 1 , 4 , 9 , 584 , (3, 0, None, None) , 0 , )),
	(( 'EnableFormatConditionsCalculation' , 'RHS' , ), 2511, (2511, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 588 , (3, 0, None, None) , 0 , )),
	(( 'EnableFormatConditionsCalculation' , 'RHS' , ), 2511, (2511, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 592 , (3, 0, None, None) , 0 , )),
	(( 'Sort' , 'RHS' , ), 880, (880, (), [ (16393, 10, None, "IID('{000244AB-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 596 , (3, 0, None, None) , 0 , )),
	(( 'ExportAsFixedFormat' , 'Type' , 'Filename' , 'Quality' , 'IncludeDocProperties' , 
			 'IgnorePrintAreas' , 'From' , 'To' , 'OpenAfterPublish' , 'FixedFormatExtClassPtr' , 
			 ), 2493, (2493, (), [ (3, 1, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , ], 1 , 1 , 4 , 8 , 600 , (3, 0, None, None) , 0 , )),
	(( 'PrintedCommentPages' , 'RHS' , ), 2857, (2857, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 604 , (3, 0, None, None) , 0 , )),
]

win32com.client.CLSIDToClass.RegisterCLSID( "{000208D8-0000-0000-C000-000000000046}", _Worksheet )
