# -*- coding: mbcs -*-
# Created by makepy.py version 0.5.01
# By python version 3.3.1 (v3.3.1:d9893d13c628, Apr  6 2013, 20:25:12) [MSC v.1600 32 bit (Intel)]
# From type library '{00020813-0000-0000-C000-000000000046}'
# On Sat Apr  5 18:15:39 2014
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
class Range(DispatchBaseClass):
	CLSID = IID('{00020846-0000-0000-C000-000000000046}')
	coclass_clsid = None

	def Activate(self):
		return self._ApplyTypes_(304, 1, (12, 0), (), 'Activate', None,)

	# Result is of type Comment
	def AddComment(self, Text=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(1389, LCID, 1, (9, 0), ((12, 17),),Text
			)
		if ret is not None:
			ret = Dispatch(ret, 'AddComment', '{00024427-0000-0000-C000-000000000046}')
		return ret

	def AdvancedFilter(self, Action=defaultNamedNotOptArg, CriteriaRange=defaultNamedOptArg, CopyToRange=defaultNamedOptArg, Unique=defaultNamedOptArg):
		return self._ApplyTypes_(876, 1, (12, 0), ((3, 1), (12, 17), (12, 17), (12, 17)), 'AdvancedFilter', None,Action
			, CriteriaRange, CopyToRange, Unique)

	def AllocateChanges(self):
		return self._oleobj_.InvokeTypes(2855, LCID, 1, (24, 0), (),)

	def ApplyNames(self, Names=defaultNamedNotOptArg, IgnoreRelativeAbsolute=defaultNamedNotOptArg, UseRowColumnNames=defaultNamedNotOptArg, OmitColumn=defaultNamedNotOptArg
			, OmitRow=defaultNamedNotOptArg, Order=1, AppendLast=defaultNamedOptArg):
		return self._ApplyTypes_(441, 1, (12, 0), ((12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (3, 49), (12, 17)), 'ApplyNames', None,Names
			, IgnoreRelativeAbsolute, UseRowColumnNames, OmitColumn, OmitRow, Order
			, AppendLast)

	def ApplyOutlineStyles(self):
		return self._ApplyTypes_(448, 1, (12, 0), (), 'ApplyOutlineStyles', None,)

	def AutoComplete(self, String=defaultNamedNotOptArg):
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(1185, LCID, 1, (8, 0), ((8, 1),),String
			)

	def AutoFill(self, Destination=defaultNamedNotOptArg, Type=0):
		return self._ApplyTypes_(449, 1, (12, 0), ((9, 1), (3, 49)), 'AutoFill', None,Destination
			, Type)

	def AutoFilter(self, Field=defaultNamedNotOptArg, Criteria1=defaultNamedNotOptArg, Operator=1, Criteria2=defaultNamedOptArg
			, VisibleDropDown=defaultNamedOptArg):
		return self._ApplyTypes_(793, 1, (12, 0), ((12, 17), (12, 17), (3, 49), (12, 17), (12, 17)), 'AutoFilter', None,Field
			, Criteria1, Operator, Criteria2, VisibleDropDown)

	def AutoFit(self):
		return self._ApplyTypes_(237, 1, (12, 0), (), 'AutoFit', None,)

	def AutoFormat(self, Format=1, Number=defaultNamedOptArg, Font=defaultNamedOptArg, Alignment=defaultNamedOptArg
			, Border=defaultNamedOptArg, Pattern=defaultNamedOptArg, Width=defaultNamedOptArg):
		return self._ApplyTypes_(114, 1, (12, 0), ((3, 49), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)), 'AutoFormat', None,Format
			, Number, Font, Alignment, Border, Pattern
			, Width)

	def AutoOutline(self):
		return self._ApplyTypes_(1036, 1, (12, 0), (), 'AutoOutline', None,)

	def BorderAround(self, LineStyle=defaultNamedNotOptArg, Weight=2, ColorIndex=-4105, Color=defaultNamedOptArg
			, ThemeColor=defaultNamedOptArg):
		return self._ApplyTypes_(2771, 1, (12, 0), ((12, 17), (3, 49), (3, 49), (12, 17), (12, 17)), 'BorderAround', None,LineStyle
			, Weight, ColorIndex, Color, ThemeColor)

	def Calculate(self):
		return self._ApplyTypes_(279, 1, (12, 0), (), 'Calculate', None,)

	def CalculateRowMajorOrder(self):
		return self._ApplyTypes_(2364, 1, (12, 0), (), 'CalculateRowMajorOrder', None,)

	def CheckSpelling(self, CustomDictionary=defaultNamedOptArg, IgnoreUppercase=defaultNamedOptArg, AlwaysSuggest=defaultNamedOptArg, SpellLang=defaultNamedOptArg):
		return self._ApplyTypes_(505, 1, (12, 0), ((12, 17), (12, 17), (12, 17), (12, 17)), 'CheckSpelling', None,CustomDictionary
			, IgnoreUppercase, AlwaysSuggest, SpellLang)

	def Clear(self):
		return self._ApplyTypes_(111, 1, (12, 0), (), 'Clear', None,)

	def ClearComments(self):
		return self._oleobj_.InvokeTypes(1390, LCID, 1, (24, 0), (),)

	def ClearContents(self):
		return self._ApplyTypes_(113, 1, (12, 0), (), 'ClearContents', None,)

	def ClearFormats(self):
		return self._ApplyTypes_(112, 1, (12, 0), (), 'ClearFormats', None,)

	def ClearHyperlinks(self):
		return self._oleobj_.InvokeTypes(2854, LCID, 1, (24, 0), (),)

	def ClearNotes(self):
		return self._ApplyTypes_(239, 1, (12, 0), (), 'ClearNotes', None,)

	def ClearOutline(self):
		return self._ApplyTypes_(1037, 1, (12, 0), (), 'ClearOutline', None,)

	# Result is of type Range
	def ColumnDifferences(self, Comparison=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(510, LCID, 1, (9, 0), ((12, 1),),Comparison
			)
		if ret is not None:
			ret = Dispatch(ret, 'ColumnDifferences', '{00020846-0000-0000-C000-000000000046}')
		return ret

	def Consolidate(self, Sources=defaultNamedOptArg, Function=defaultNamedOptArg, TopRow=defaultNamedOptArg, LeftColumn=defaultNamedOptArg
			, CreateLinks=defaultNamedOptArg):
		return self._ApplyTypes_(482, 1, (12, 0), ((12, 17), (12, 17), (12, 17), (12, 17), (12, 17)), 'Consolidate', None,Sources
			, Function, TopRow, LeftColumn, CreateLinks)

	def Copy(self, Destination=defaultNamedOptArg):
		return self._ApplyTypes_(551, 1, (12, 0), ((12, 17),), 'Copy', None,Destination
			)

	def CopyFromRecordset(self, Data=defaultNamedNotOptArg, MaxRows=defaultNamedOptArg, MaxColumns=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1152, LCID, 1, (3, 0), ((13, 1), (12, 17), (12, 17)),Data
			, MaxRows, MaxColumns)

	def CopyPicture(self, Appearance=1, Format=-4147):
		return self._ApplyTypes_(213, 1, (12, 0), ((3, 49), (3, 49)), 'CopyPicture', None,Appearance
			, Format)

	def CreateNames(self, Top=defaultNamedOptArg, Left=defaultNamedOptArg, Bottom=defaultNamedOptArg, Right=defaultNamedOptArg):
		return self._ApplyTypes_(457, 1, (12, 0), ((12, 17), (12, 17), (12, 17), (12, 17)), 'CreateNames', None,Top
			, Left, Bottom, Right)

	def CreatePublisher(self, Edition=defaultNamedNotOptArg, Appearance=1, ContainsPICT=defaultNamedOptArg, ContainsBIFF=defaultNamedOptArg
			, ContainsRTF=defaultNamedOptArg, ContainsVALU=defaultNamedOptArg):
		return self._ApplyTypes_(458, 1, (12, 0), ((12, 17), (3, 49), (12, 17), (12, 17), (12, 17), (12, 17)), 'CreatePublisher', None,Edition
			, Appearance, ContainsPICT, ContainsBIFF, ContainsRTF, ContainsVALU
			)

	def Cut(self, Destination=defaultNamedOptArg):
		return self._ApplyTypes_(565, 1, (12, 0), ((12, 17),), 'Cut', None,Destination
			)

	def DataSeries(self, Rowcol=defaultNamedNotOptArg, Type=-4132, Date=1, Step=defaultNamedOptArg
			, Stop=defaultNamedOptArg, Trend=defaultNamedOptArg):
		return self._ApplyTypes_(464, 1, (12, 0), ((12, 17), (3, 49), (3, 49), (12, 17), (12, 17), (12, 17)), 'DataSeries', None,Rowcol
			, Type, Date, Step, Stop, Trend
			)

	def Delete(self, Shift=defaultNamedOptArg):
		return self._ApplyTypes_(117, 1, (12, 0), ((12, 17),), 'Delete', None,Shift
			)

	def DialogBox(self):
		return self._ApplyTypes_(245, 1, (12, 0), (), 'DialogBox', None,)

	def Dirty(self):
		return self._oleobj_.InvokeTypes(2014, LCID, 1, (24, 0), (),)

	def DiscardChanges(self):
		return self._oleobj_.InvokeTypes(2856, LCID, 1, (24, 0), (),)

	def EditionOptions(self, Type=defaultNamedNotOptArg, Option=defaultNamedNotOptArg, Name=defaultNamedNotOptArg, Reference=defaultNamedNotOptArg
			, Appearance=1, ChartSize=1, Format=defaultNamedOptArg):
		return self._ApplyTypes_(1131, 1, (12, 0), ((3, 1), (3, 1), (12, 17), (12, 17), (3, 49), (3, 49), (12, 17)), 'EditionOptions', None,Type
			, Option, Name, Reference, Appearance, ChartSize
			, Format)

	# Result is of type Range
	# The method End is actually a property, but must be used as a method to correctly pass the arguments
	def End(self, Direction=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(500, LCID, 2, (9, 0), ((3, 1),),Direction
			)
		if ret is not None:
			ret = Dispatch(ret, 'End', '{00020846-0000-0000-C000-000000000046}')
		return ret

	def ExportAsFixedFormat(self, Type=defaultNamedNotOptArg, Filename=defaultNamedOptArg, Quality=defaultNamedOptArg, IncludeDocProperties=defaultNamedOptArg
			, IgnorePrintAreas=defaultNamedOptArg, From=defaultNamedOptArg, To=defaultNamedOptArg, OpenAfterPublish=defaultNamedOptArg, FixedFormatExtClassPtr=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2493, LCID, 1, (24, 0), ((3, 1), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)),Type
			, Filename, Quality, IncludeDocProperties, IgnorePrintAreas, From
			, To, OpenAfterPublish, FixedFormatExtClassPtr)

	def FillDown(self):
		return self._ApplyTypes_(248, 1, (12, 0), (), 'FillDown', None,)

	def FillLeft(self):
		return self._ApplyTypes_(249, 1, (12, 0), (), 'FillLeft', None,)

	def FillRight(self):
		return self._ApplyTypes_(250, 1, (12, 0), (), 'FillRight', None,)

	def FillUp(self):
		return self._ApplyTypes_(251, 1, (12, 0), (), 'FillUp', None,)

	# Result is of type Range
	def Find(self, What=defaultNamedNotOptArg, After=defaultNamedNotOptArg, LookIn=defaultNamedNotOptArg, LookAt=defaultNamedNotOptArg
			, SearchOrder=defaultNamedNotOptArg, SearchDirection=1, MatchCase=defaultNamedOptArg, MatchByte=defaultNamedOptArg, SearchFormat=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(398, LCID, 1, (9, 0), ((12, 1), (12, 17), (12, 17), (12, 17), (12, 17), (3, 49), (12, 17), (12, 17), (12, 17)),What
			, After, LookIn, LookAt, SearchOrder, SearchDirection
			, MatchCase, MatchByte, SearchFormat)
		if ret is not None:
			ret = Dispatch(ret, 'Find', '{00020846-0000-0000-C000-000000000046}')
		return ret

	# Result is of type Range
	def FindNext(self, After=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(399, LCID, 1, (9, 0), ((12, 17),),After
			)
		if ret is not None:
			ret = Dispatch(ret, 'FindNext', '{00020846-0000-0000-C000-000000000046}')
		return ret

	# Result is of type Range
	def FindPrevious(self, After=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(400, LCID, 1, (9, 0), ((12, 17),),After
			)
		if ret is not None:
			ret = Dispatch(ret, 'FindPrevious', '{00020846-0000-0000-C000-000000000046}')
		return ret

	def FunctionWizard(self):
		return self._ApplyTypes_(571, 1, (12, 0), (), 'FunctionWizard', None,)

	# The method GetAddress is actually a property, but must be used as a method to correctly pass the arguments
	def GetAddress(self, RowAbsolute=defaultNamedNotOptArg, ColumnAbsolute=defaultNamedNotOptArg, ReferenceStyle=1, External=defaultNamedOptArg
			, RelativeTo=defaultNamedOptArg):
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(236, LCID, 2, (8, 0), ((12, 17), (12, 17), (3, 49), (12, 17), (12, 17)),RowAbsolute
			, ColumnAbsolute, ReferenceStyle, External, RelativeTo)

	# The method GetAddressLocal is actually a property, but must be used as a method to correctly pass the arguments
	def GetAddressLocal(self, RowAbsolute=defaultNamedNotOptArg, ColumnAbsolute=defaultNamedNotOptArg, ReferenceStyle=1, External=defaultNamedOptArg
			, RelativeTo=defaultNamedOptArg):
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(437, LCID, 2, (8, 0), ((12, 17), (12, 17), (3, 49), (12, 17), (12, 17)),RowAbsolute
			, ColumnAbsolute, ReferenceStyle, External, RelativeTo)

	# Result is of type Characters
	# The method GetCharacters is actually a property, but must be used as a method to correctly pass the arguments
	def GetCharacters(self, Start=defaultNamedOptArg, Length=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(603, LCID, 2, (9, 0), ((12, 17), (12, 17)),Start
			, Length)
		if ret is not None:
			ret = Dispatch(ret, 'GetCharacters', '{00020878-0000-0000-C000-000000000046}')
		return ret

	# Result is of type Range
	# The method GetOffset is actually a property, but must be used as a method to correctly pass the arguments
	def GetOffset(self, RowOffset=defaultNamedOptArg, ColumnOffset=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(254, LCID, 2, (9, 0), ((12, 17), (12, 17)),RowOffset
			, ColumnOffset)
		if ret is not None:
			ret = Dispatch(ret, 'GetOffset', '{00020846-0000-0000-C000-000000000046}')
		return ret

	# Result is of type Range
	# The method GetResize is actually a property, but must be used as a method to correctly pass the arguments
	def GetResize(self, RowSize=defaultNamedOptArg, ColumnSize=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(256, LCID, 2, (9, 0), ((12, 17), (12, 17)),RowSize
			, ColumnSize)
		if ret is not None:
			ret = Dispatch(ret, 'GetResize', '{00020846-0000-0000-C000-000000000046}')
		return ret

	# The method GetValue is actually a property, but must be used as a method to correctly pass the arguments
	def GetValue(self, RangeValueDataType=defaultNamedOptArg):
		return self._ApplyTypes_(6, 2, (12, 0), ((12, 17),), 'GetValue', None,RangeValueDataType
			)

	# The method Get_Default is actually a property, but must be used as a method to correctly pass the arguments
	def Get_Default(self, RowIndex=defaultNamedOptArg, ColumnIndex=defaultNamedOptArg):
		return self._ApplyTypes_(0, 2, (12, 0), ((12, 17), (12, 17)), 'Get_Default', None,RowIndex
			, ColumnIndex)

	def GoalSeek(self, Goal=defaultNamedNotOptArg, ChangingCell=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(472, LCID, 1, (11, 0), ((12, 1), (9, 1)),Goal
			, ChangingCell)

	def Group(self, Start=defaultNamedOptArg, End=defaultNamedOptArg, By=defaultNamedOptArg, Periods=defaultNamedOptArg):
		return self._ApplyTypes_(46, 1, (12, 0), ((12, 17), (12, 17), (12, 17), (12, 17)), 'Group', None,Start
			, End, By, Periods)

	def Insert(self, Shift=defaultNamedOptArg, CopyOrigin=defaultNamedOptArg):
		return self._ApplyTypes_(252, 1, (12, 0), ((12, 17), (12, 17)), 'Insert', None,Shift
			, CopyOrigin)

	def InsertIndent(self, InsertAmount=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(1381, LCID, 1, (24, 0), ((3, 1),),InsertAmount
			)

	# The method Item is actually a property, but must be used as a method to correctly pass the arguments
	def Item(self, RowIndex=defaultNamedNotOptArg, ColumnIndex=defaultNamedOptArg):
		return self._ApplyTypes_(170, 2, (12, 0), ((12, 1), (12, 17)), 'Item', None,RowIndex
			, ColumnIndex)

	def Justify(self):
		return self._ApplyTypes_(495, 1, (12, 0), (), 'Justify', None,)

	def ListNames(self):
		return self._ApplyTypes_(253, 1, (12, 0), (), 'ListNames', None,)

	def Merge(self, Across=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(564, LCID, 1, (24, 0), ((12, 17),),Across
			)

	def NavigateArrow(self, TowardPrecedent=defaultNamedOptArg, ArrowNumber=defaultNamedOptArg, LinkNumber=defaultNamedOptArg):
		return self._ApplyTypes_(1032, 1, (12, 0), ((12, 17), (12, 17), (12, 17)), 'NavigateArrow', None,TowardPrecedent
			, ArrowNumber, LinkNumber)

	def NoteText(self, Text=defaultNamedOptArg, Start=defaultNamedOptArg, Length=defaultNamedOptArg):
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(1127, LCID, 1, (8, 0), ((12, 17), (12, 17), (12, 17)),Text
			, Start, Length)

	def Parse(self, ParseLine=defaultNamedOptArg, Destination=defaultNamedOptArg):
		return self._ApplyTypes_(477, 1, (12, 0), ((12, 17), (12, 17)), 'Parse', None,ParseLine
			, Destination)

	def PasteSpecial(self, Paste=-4104, Operation=-4142, SkipBlanks=defaultNamedOptArg, Transpose=defaultNamedOptArg):
		return self._ApplyTypes_(1928, 1, (12, 0), ((3, 49), (3, 49), (12, 17), (12, 17)), 'PasteSpecial', None,Paste
			, Operation, SkipBlanks, Transpose)

	def PrintOut(self, From=defaultNamedOptArg, To=defaultNamedOptArg, Copies=defaultNamedOptArg, Preview=defaultNamedOptArg
			, ActivePrinter=defaultNamedOptArg, PrintToFile=defaultNamedOptArg, Collate=defaultNamedOptArg, PrToFileName=defaultNamedOptArg):
		return self._ApplyTypes_(2361, 1, (12, 0), ((12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)), 'PrintOut', None,From
			, To, Copies, Preview, ActivePrinter, PrintToFile
			, Collate, PrToFileName)

	def PrintPreview(self, EnableChanges=defaultNamedOptArg):
		return self._ApplyTypes_(281, 1, (12, 0), ((12, 17),), 'PrintPreview', None,EnableChanges
			)

	# Result is of type Range
	# The method Range is actually a property, but must be used as a method to correctly pass the arguments
	def Range(self, Cell1=defaultNamedNotOptArg, Cell2=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(197, LCID, 2, (9, 0), ((12, 1), (12, 17)),Cell1
			, Cell2)
		if ret is not None:
			ret = Dispatch(ret, 'Range', '{00020846-0000-0000-C000-000000000046}')
		return ret

	def RemoveDuplicates(self, Columns=defaultNamedNotOptArg, Header=2):
		return self._oleobj_.InvokeTypes(2492, LCID, 1, (24, 0), ((12, 17), (3, 49)),Columns
			, Header)

	def RemoveSubtotal(self):
		return self._ApplyTypes_(883, 1, (12, 0), (), 'RemoveSubtotal', None,)

	def Replace(self, What=defaultNamedNotOptArg, Replacement=defaultNamedNotOptArg, LookAt=defaultNamedOptArg, SearchOrder=defaultNamedOptArg
			, MatchCase=defaultNamedOptArg, MatchByte=defaultNamedOptArg, SearchFormat=defaultNamedOptArg, ReplaceFormat=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(226, LCID, 1, (11, 0), ((12, 1), (12, 1), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)),What
			, Replacement, LookAt, SearchOrder, MatchCase, MatchByte
			, SearchFormat, ReplaceFormat)

	# Result is of type Range
	def RowDifferences(self, Comparison=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(511, LCID, 1, (9, 0), ((12, 1),),Comparison
			)
		if ret is not None:
			ret = Dispatch(ret, 'RowDifferences', '{00020846-0000-0000-C000-000000000046}')
		return ret

	def Run(self, Arg1=defaultNamedOptArg, Arg2=defaultNamedOptArg, Arg3=defaultNamedOptArg, Arg4=defaultNamedOptArg
			, Arg5=defaultNamedOptArg, Arg6=defaultNamedOptArg, Arg7=defaultNamedOptArg, Arg8=defaultNamedOptArg, Arg9=defaultNamedOptArg
			, Arg10=defaultNamedOptArg, Arg11=defaultNamedOptArg, Arg12=defaultNamedOptArg, Arg13=defaultNamedOptArg, Arg14=defaultNamedOptArg
			, Arg15=defaultNamedOptArg, Arg16=defaultNamedOptArg, Arg17=defaultNamedOptArg, Arg18=defaultNamedOptArg, Arg19=defaultNamedOptArg
			, Arg20=defaultNamedOptArg, Arg21=defaultNamedOptArg, Arg22=defaultNamedOptArg, Arg23=defaultNamedOptArg, Arg24=defaultNamedOptArg
			, Arg25=defaultNamedOptArg, Arg26=defaultNamedOptArg, Arg27=defaultNamedOptArg, Arg28=defaultNamedOptArg, Arg29=defaultNamedOptArg
			, Arg30=defaultNamedOptArg):
		return self._ApplyTypes_(259, 1, (12, 0), ((12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)), 'Run', None,Arg1
			, Arg2, Arg3, Arg4, Arg5, Arg6
			, Arg7, Arg8, Arg9, Arg10, Arg11
			, Arg12, Arg13, Arg14, Arg15, Arg16
			, Arg17, Arg18, Arg19, Arg20, Arg21
			, Arg22, Arg23, Arg24, Arg25, Arg26
			, Arg27, Arg28, Arg29, Arg30)

	def Select(self):
		return self._ApplyTypes_(235, 1, (12, 0), (), 'Select', None,)

	# The method SetItem is actually a property, but must be used as a method to correctly pass the arguments
	def SetItem(self, RowIndex=defaultNamedNotOptArg, ColumnIndex=defaultNamedNotOptArg, arg2=defaultUnnamedArg):
		return self._oleobj_.InvokeTypes(170, LCID, 4, (24, 0), ((12, 1), (12, 17), (12, 1)),RowIndex
			, ColumnIndex, arg2)

	def SetPhonetic(self):
		return self._oleobj_.InvokeTypes(1812, LCID, 1, (24, 0), (),)

	# The method SetValue is actually a property, but must be used as a method to correctly pass the arguments
	def SetValue(self, RangeValueDataType=defaultNamedNotOptArg, arg1=defaultUnnamedArg):
		return self._oleobj_.InvokeTypes(6, LCID, 4, (24, 0), ((12, 17), (12, 1)),RangeValueDataType
			, arg1)

	# The method Set_Default is actually a property, but must be used as a method to correctly pass the arguments
	def Set_Default(self, RowIndex=defaultNamedNotOptArg, ColumnIndex=defaultNamedOptArg, arg2=defaultUnnamedArg):
		return self._oleobj_.InvokeTypes(0, LCID, 4, (24, 0), ((12, 17), (12, 17), (12, 1)),RowIndex
			, ColumnIndex, arg2)

	def Show(self):
		return self._ApplyTypes_(496, 1, (12, 0), (), 'Show', None,)

	def ShowDependents(self, Remove=defaultNamedOptArg):
		return self._ApplyTypes_(877, 1, (12, 0), ((12, 17),), 'ShowDependents', None,Remove
			)

	def ShowErrors(self):
		return self._ApplyTypes_(878, 1, (12, 0), (), 'ShowErrors', None,)

	def ShowPrecedents(self, Remove=defaultNamedOptArg):
		return self._ApplyTypes_(879, 1, (12, 0), ((12, 17),), 'ShowPrecedents', None,Remove
			)

	def Sort(self, Key1=defaultNamedNotOptArg, Order1=1, Key2=defaultNamedNotOptArg, Type=defaultNamedNotOptArg
			, Order2=1, Key3=defaultNamedNotOptArg, Order3=1, Header=2, OrderCustom=defaultNamedNotOptArg
			, MatchCase=defaultNamedNotOptArg, Orientation=2, SortMethod=1, DataOption1=0, DataOption2=0
			, DataOption3=0):
		return self._ApplyTypes_(880, 1, (12, 0), ((12, 17), (3, 49), (12, 17), (12, 17), (3, 49), (12, 17), (3, 49), (3, 49), (12, 17), (12, 17), (3, 49), (3, 49), (3, 49), (3, 49), (3, 49)), 'Sort', None,Key1
			, Order1, Key2, Type, Order2, Key3
			, Order3, Header, OrderCustom, MatchCase, Orientation
			, SortMethod, DataOption1, DataOption2, DataOption3)

	def SortSpecial(self, SortMethod=1, Key1=defaultNamedNotOptArg, Order1=1, Type=defaultNamedNotOptArg
			, Key2=defaultNamedNotOptArg, Order2=1, Key3=defaultNamedNotOptArg, Order3=1, Header=2
			, OrderCustom=defaultNamedNotOptArg, MatchCase=defaultNamedNotOptArg, Orientation=2, DataOption1=0, DataOption2=0
			, DataOption3=0):
		return self._ApplyTypes_(881, 1, (12, 0), ((3, 49), (12, 17), (3, 49), (12, 17), (12, 17), (3, 49), (12, 17), (3, 49), (3, 49), (12, 17), (12, 17), (3, 49), (3, 49), (3, 49), (3, 49)), 'SortSpecial', None,SortMethod
			, Key1, Order1, Type, Key2, Order2
			, Key3, Order3, Header, OrderCustom, MatchCase
			, Orientation, DataOption1, DataOption2, DataOption3)

	def Speak(self, SpeakDirection=defaultNamedOptArg, SpeakFormulas=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2017, LCID, 1, (24, 0), ((12, 17), (12, 17)),SpeakDirection
			, SpeakFormulas)

	# Result is of type Range
	def SpecialCells(self, Type=defaultNamedNotOptArg, Value=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(410, LCID, 1, (9, 0), ((3, 1), (12, 17)),Type
			, Value)
		if ret is not None:
			ret = Dispatch(ret, 'SpecialCells', '{00020846-0000-0000-C000-000000000046}')
		return ret

	def SubscribeTo(self, Edition=defaultNamedNotOptArg, Format=-4158):
		return self._ApplyTypes_(481, 1, (12, 0), ((8, 1), (3, 49)), 'SubscribeTo', None,Edition
			, Format)

	def Subtotal(self, GroupBy=defaultNamedNotOptArg, Function=defaultNamedNotOptArg, TotalList=defaultNamedNotOptArg, Replace=defaultNamedNotOptArg
			, PageBreaks=defaultNamedNotOptArg, SummaryBelowData=1):
		return self._ApplyTypes_(882, 1, (12, 0), ((3, 1), (3, 1), (12, 1), (12, 17), (12, 17), (3, 49)), 'Subtotal', None,GroupBy
			, Function, TotalList, Replace, PageBreaks, SummaryBelowData
			)

	def Table(self, RowInput=defaultNamedOptArg, ColumnInput=defaultNamedOptArg):
		return self._ApplyTypes_(497, 1, (12, 0), ((12, 17), (12, 17)), 'Table', None,RowInput
			, ColumnInput)

	def TextToColumns(self, Destination=defaultNamedNotOptArg, DataType=1, TextQualifier=1, ConsecutiveDelimiter=defaultNamedOptArg
			, Tab=defaultNamedOptArg, Semicolon=defaultNamedOptArg, Comma=defaultNamedOptArg, Space=defaultNamedOptArg, Other=defaultNamedOptArg
			, OtherChar=defaultNamedOptArg, FieldInfo=defaultNamedOptArg, DecimalSeparator=defaultNamedOptArg, ThousandsSeparator=defaultNamedOptArg, TrailingMinusNumbers=defaultNamedOptArg):
		return self._ApplyTypes_(1040, 1, (12, 0), ((12, 17), (3, 49), (3, 49), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)), 'TextToColumns', None,Destination
			, DataType, TextQualifier, ConsecutiveDelimiter, Tab, Semicolon
			, Comma, Space, Other, OtherChar, FieldInfo
			, DecimalSeparator, ThousandsSeparator, TrailingMinusNumbers)

	def UnMerge(self):
		return self._oleobj_.InvokeTypes(1384, LCID, 1, (24, 0), (),)

	def Ungroup(self):
		return self._ApplyTypes_(244, 1, (12, 0), (), 'Ungroup', None,)

	def _BorderAround(self, LineStyle=defaultNamedNotOptArg, Weight=2, ColorIndex=-4105, Color=defaultNamedOptArg):
		return self._ApplyTypes_(1067, 1, (12, 0), ((12, 17), (3, 49), (3, 49), (12, 17)), '_BorderAround', None,LineStyle
			, Weight, ColorIndex, Color)

	def _PasteSpecial(self, Paste=-4104, Operation=-4142, SkipBlanks=defaultNamedOptArg, Transpose=defaultNamedOptArg):
		return self._ApplyTypes_(1027, 1, (12, 0), ((3, 49), (3, 49), (12, 17), (12, 17)), '_PasteSpecial', None,Paste
			, Operation, SkipBlanks, Transpose)

	def _PrintOut(self, From=defaultNamedOptArg, To=defaultNamedOptArg, Copies=defaultNamedOptArg, Preview=defaultNamedOptArg
			, ActivePrinter=defaultNamedOptArg, PrintToFile=defaultNamedOptArg, Collate=defaultNamedOptArg, PrToFileName=defaultNamedOptArg):
		return self._ApplyTypes_(1772, 1, (12, 0), ((12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)), '_PrintOut', None,From
			, To, Copies, Preview, ActivePrinter, PrintToFile
			, Collate, PrToFileName)

	def _PrintOut_(self, From=defaultNamedOptArg, To=defaultNamedOptArg, Copies=defaultNamedOptArg, Preview=defaultNamedOptArg
			, ActivePrinter=defaultNamedOptArg, PrintToFile=defaultNamedOptArg, Collate=defaultNamedOptArg):
		return self._ApplyTypes_(905, 1, (12, 0), ((12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)), '_PrintOut_', None,From
			, To, Copies, Preview, ActivePrinter, PrintToFile
			, Collate)

	_prop_map_get_ = {
		"AddIndent": (1063, 2, (12, 0), (), "AddIndent", None),
		"Address": (236, 2, (8, 0), ((12, 17), (12, 17), (3, 49), (12, 17), (12, 17)), "Address", None),
		"AddressLocal": (437, 2, (8, 0), ((12, 17), (12, 17), (3, 49), (12, 17), (12, 17)), "AddressLocal", None),
		"AllowEdit": (2020, 2, (11, 0), (), "AllowEdit", None),
		# Method 'Application' returns object of type 'Application'
		"Application": (148, 2, (13, 0), (), "Application", '{00024500-0000-0000-C000-000000000046}'),
		# Method 'Areas' returns object of type 'Areas'
		"Areas": (568, 2, (9, 0), (), "Areas", '{00020860-0000-0000-C000-000000000046}'),
		# Method 'Borders' returns object of type 'Borders'
		"Borders": (435, 2, (9, 0), (), "Borders", '{00020855-0000-0000-C000-000000000046}'),
		# Method 'Cells' returns object of type 'Range'
		"Cells": (238, 2, (9, 0), (), "Cells", '{00020846-0000-0000-C000-000000000046}'),
		# Method 'Characters' returns object of type 'Characters'
		"Characters": (603, 2, (9, 0), ((12, 17), (12, 17)), "Characters", '{00020878-0000-0000-C000-000000000046}'),
		"Column": (240, 2, (3, 0), (), "Column", None),
		"ColumnWidth": (242, 2, (12, 0), (), "ColumnWidth", None),
		# Method 'Columns' returns object of type 'Range'
		"Columns": (241, 2, (9, 0), (), "Columns", '{00020846-0000-0000-C000-000000000046}'),
		# Method 'Comment' returns object of type 'Comment'
		"Comment": (910, 2, (9, 0), (), "Comment", '{00024427-0000-0000-C000-000000000046}'),
		"Count": (118, 2, (3, 0), (), "Count", None),
		"CountLarge": (2499, 2, (12, 0), (), "CountLarge", None),
		"Creator": (149, 2, (3, 0), (), "Creator", None),
		# Method 'CurrentArray' returns object of type 'Range'
		"CurrentArray": (501, 2, (9, 0), (), "CurrentArray", '{00020846-0000-0000-C000-000000000046}'),
		# Method 'CurrentRegion' returns object of type 'Range'
		"CurrentRegion": (243, 2, (9, 0), (), "CurrentRegion", '{00020846-0000-0000-C000-000000000046}'),
		# Method 'Dependents' returns object of type 'Range'
		"Dependents": (543, 2, (9, 0), (), "Dependents", '{00020846-0000-0000-C000-000000000046}'),
		# Method 'DirectDependents' returns object of type 'Range'
		"DirectDependents": (545, 2, (9, 0), (), "DirectDependents", '{00020846-0000-0000-C000-000000000046}'),
		# Method 'DirectPrecedents' returns object of type 'Range'
		"DirectPrecedents": (546, 2, (9, 0), (), "DirectPrecedents", '{00020846-0000-0000-C000-000000000046}'),
		# Method 'DisplayFormat' returns object of type 'DisplayFormat'
		"DisplayFormat": (666, 2, (9, 0), (), "DisplayFormat", '{000244C2-0000-0000-C000-000000000046}'),
		# Method 'EntireColumn' returns object of type 'Range'
		"EntireColumn": (246, 2, (9, 0), (), "EntireColumn", '{00020846-0000-0000-C000-000000000046}'),
		# Method 'EntireRow' returns object of type 'Range'
		"EntireRow": (247, 2, (9, 0), (), "EntireRow", '{00020846-0000-0000-C000-000000000046}'),
		# Method 'Errors' returns object of type 'Errors'
		"Errors": (2015, 2, (9, 0), (), "Errors", '{0002445C-0000-0000-C000-000000000046}'),
		# Method 'Font' returns object of type 'Font'
		"Font": (146, 2, (9, 0), (), "Font", '{0002084D-0000-0000-C000-000000000046}'),
		# Method 'FormatConditions' returns object of type 'FormatConditions'
		"FormatConditions": (1392, 2, (9, 0), (), "FormatConditions", '{00024424-0000-0000-C000-000000000046}'),
		"Formula": (261, 2, (12, 0), (), "Formula", None),
		"FormulaArray": (586, 2, (12, 0), (), "FormulaArray", None),
		"FormulaHidden": (262, 2, (12, 0), (), "FormulaHidden", None),
		"FormulaLabel": (1380, 2, (3, 0), (), "FormulaLabel", None),
		"FormulaLocal": (263, 2, (12, 0), (), "FormulaLocal", None),
		"FormulaR1C1": (264, 2, (12, 0), (), "FormulaR1C1", None),
		"FormulaR1C1Local": (265, 2, (12, 0), (), "FormulaR1C1Local", None),
		"HasArray": (266, 2, (12, 0), (), "HasArray", None),
		"HasFormula": (267, 2, (12, 0), (), "HasFormula", None),
		"Height": (123, 2, (12, 0), (), "Height", None),
		"Hidden": (268, 2, (12, 0), (), "Hidden", None),
		"HorizontalAlignment": (136, 2, (12, 0), (), "HorizontalAlignment", None),
		# Method 'Hyperlinks' returns object of type 'Hyperlinks'
		"Hyperlinks": (1393, 2, (9, 0), (), "Hyperlinks", '{00024430-0000-0000-C000-000000000046}'),
		"ID": (1813, 2, (8, 0), (), "ID", None),
		"IndentLevel": (201, 2, (12, 0), (), "IndentLevel", None),
		# Method 'Interior' returns object of type 'Interior'
		"Interior": (129, 2, (9, 0), (), "Interior", '{00020870-0000-0000-C000-000000000046}'),
		"Left": (127, 2, (12, 0), (), "Left", None),
		"ListHeaderRows": (1187, 2, (3, 0), (), "ListHeaderRows", None),
		# Method 'ListObject' returns object of type 'ListObject'
		"ListObject": (2257, 2, (9, 0), (), "ListObject", '{00024471-0000-0000-C000-000000000046}'),
		"LocationInTable": (691, 2, (3, 0), (), "LocationInTable", None),
		"Locked": (269, 2, (12, 0), (), "Locked", None),
		"MDX": (2123, 2, (8, 0), (), "MDX", None),
		# Method 'MergeArea' returns object of type 'Range'
		"MergeArea": (1385, 2, (9, 0), (), "MergeArea", '{00020846-0000-0000-C000-000000000046}'),
		"MergeCells": (208, 2, (12, 0), (), "MergeCells", None),
		"Name": (110, 2, (12, 0), (), "Name", None),
		# Method 'Next' returns object of type 'Range'
		"Next": (502, 2, (9, 0), (), "Next", '{00020846-0000-0000-C000-000000000046}'),
		"NumberFormat": (193, 2, (12, 0), (), "NumberFormat", None),
		"NumberFormatLocal": (1097, 2, (12, 0), (), "NumberFormatLocal", None),
		# Method 'Offset' returns object of type 'Range'
		"Offset": (254, 2, (9, 0), ((12, 17), (12, 17)), "Offset", '{00020846-0000-0000-C000-000000000046}'),
		"Orientation": (134, 2, (12, 0), (), "Orientation", None),
		"OutlineLevel": (271, 2, (12, 0), (), "OutlineLevel", None),
		"PageBreak": (255, 2, (3, 0), (), "PageBreak", None),
		"Parent": (150, 2, (9, 0), (), "Parent", None),
		# Method 'Phonetic' returns object of type 'Phonetic'
		"Phonetic": (1391, 2, (9, 0), (), "Phonetic", '{00024438-0000-0000-C000-000000000046}'),
		# Method 'Phonetics' returns object of type 'Phonetics'
		"Phonetics": (1811, 2, (9, 0), (), "Phonetics", '{00024447-0000-0000-C000-000000000046}'),
		# Method 'PivotCell' returns object of type 'PivotCell'
		"PivotCell": (2013, 2, (9, 0), (), "PivotCell", '{00024458-0000-0000-C000-000000000046}'),
		# Method 'PivotField' returns object of type 'PivotField'
		"PivotField": (731, 2, (9, 0), (), "PivotField", '{00020874-0000-0000-C000-000000000046}'),
		# Method 'PivotItem' returns object of type 'PivotItem'
		"PivotItem": (740, 2, (9, 0), (), "PivotItem", '{00020876-0000-0000-C000-000000000046}'),
		# Method 'PivotTable' returns object of type 'PivotTable'
		"PivotTable": (716, 2, (9, 0), (), "PivotTable", '{00020872-0000-0000-C000-000000000046}'),
		# Method 'Precedents' returns object of type 'Range'
		"Precedents": (544, 2, (9, 0), (), "Precedents", '{00020846-0000-0000-C000-000000000046}'),
		"PrefixCharacter": (504, 2, (12, 0), (), "PrefixCharacter", None),
		# Method 'Previous' returns object of type 'Range'
		"Previous": (503, 2, (9, 0), (), "Previous", '{00020846-0000-0000-C000-000000000046}'),
		# Method 'QueryTable' returns object of type 'QueryTable'
		"QueryTable": (1386, 2, (13, 0), (), "QueryTable", '{59191DA1-EA47-11CE-A51F-00AA0061507F}'),
		"ReadingOrder": (975, 2, (3, 0), (), "ReadingOrder", None),
		# Method 'Resize' returns object of type 'Range'
		"Resize": (256, 2, (9, 0), ((12, 17), (12, 17)), "Resize", '{00020846-0000-0000-C000-000000000046}'),
		"Row": (257, 2, (3, 0), (), "Row", None),
		"RowHeight": (272, 2, (12, 0), (), "RowHeight", None),
		# Method 'Rows' returns object of type 'Range'
		"Rows": (258, 2, (9, 0), (), "Rows", '{00020846-0000-0000-C000-000000000046}'),
		# Method 'ServerActions' returns object of type 'Actions'
		"ServerActions": (2491, 2, (9, 0), (), "ServerActions", '{00024490-0000-0000-C000-000000000046}'),
		"ShowDetail": (585, 2, (12, 0), (), "ShowDetail", None),
		"ShrinkToFit": (209, 2, (12, 0), (), "ShrinkToFit", None),
		# Method 'SmartTags' returns object of type 'SmartTags'
		"SmartTags": (2016, 2, (9, 0), (), "SmartTags", '{00024461-0000-0000-C000-000000000046}'),
		# Method 'SoundNote' returns object of type 'SoundNote'
		"SoundNote": (916, 2, (9, 0), (), "SoundNote", '{0002087B-0000-0000-C000-000000000046}'),
		# Method 'SparklineGroups' returns object of type 'SparklineGroups'
		"SparklineGroups": (2853, 2, (9, 0), (), "SparklineGroups", '{000244B6-0000-0000-C000-000000000046}'),
		"Style": (260, 2, (12, 0), (), "Style", None),
		"Summary": (273, 2, (12, 0), (), "Summary", None),
		"Text": (138, 2, (12, 0), (), "Text", None),
		"Top": (126, 2, (12, 0), (), "Top", None),
		"UseStandardHeight": (274, 2, (12, 0), (), "UseStandardHeight", None),
		"UseStandardWidth": (275, 2, (12, 0), (), "UseStandardWidth", None),
		# Method 'Validation' returns object of type 'Validation'
		"Validation": (1387, 2, (9, 0), (), "Validation", '{0002442F-0000-0000-C000-000000000046}'),
		"Value": (6, 2, (12, 0), ((12, 17),), "Value", None),
		"Value2": (1388, 2, (12, 0), (), "Value2", None),
		"VerticalAlignment": (137, 2, (12, 0), (), "VerticalAlignment", None),
		"Width": (122, 2, (12, 0), (), "Width", None),
		# Method 'Worksheet' returns object of type 'Worksheet'
		"Worksheet": (348, 2, (13, 0), (), "Worksheet", '{00020820-0000-0000-C000-000000000046}'),
		"WrapText": (276, 2, (12, 0), (), "WrapText", None),
		# Method 'XPath' returns object of type 'XPath'
		"XPath": (2258, 2, (9, 0), (), "XPath", '{0002447E-0000-0000-C000-000000000046}'),
		"_Default": (0, 2, (12, 0), ((12, 17), (12, 17)), "_Default", None),
	}
	_prop_map_put_ = {
		"AddIndent": ((1063, LCID, 4, 0),()),
		"ColumnWidth": ((242, LCID, 4, 0),()),
		"Formula": ((261, LCID, 4, 0),()),
		"FormulaArray": ((586, LCID, 4, 0),()),
		"FormulaHidden": ((262, LCID, 4, 0),()),
		"FormulaLabel": ((1380, LCID, 4, 0),()),
		"FormulaLocal": ((263, LCID, 4, 0),()),
		"FormulaR1C1": ((264, LCID, 4, 0),()),
		"FormulaR1C1Local": ((265, LCID, 4, 0),()),
		"Hidden": ((268, LCID, 4, 0),()),
		"HorizontalAlignment": ((136, LCID, 4, 0),()),
		"ID": ((1813, LCID, 4, 0),()),
		"IndentLevel": ((201, LCID, 4, 0),()),
		"Locked": ((269, LCID, 4, 0),()),
		"MergeCells": ((208, LCID, 4, 0),()),
		"Name": ((110, LCID, 4, 0),()),
		"NumberFormat": ((193, LCID, 4, 0),()),
		"NumberFormatLocal": ((1097, LCID, 4, 0),()),
		"Orientation": ((134, LCID, 4, 0),()),
		"OutlineLevel": ((271, LCID, 4, 0),()),
		"PageBreak": ((255, LCID, 4, 0),()),
		"ReadingOrder": ((975, LCID, 4, 0),()),
		"RowHeight": ((272, LCID, 4, 0),()),
		"ShowDetail": ((585, LCID, 4, 0),()),
		"ShrinkToFit": ((209, LCID, 4, 0),()),
		"Style": ((260, LCID, 4, 0),()),
		"UseStandardHeight": ((274, LCID, 4, 0),()),
		"UseStandardWidth": ((275, LCID, 4, 0),()),
		"Value": ((6, LCID, 4, 0),()),
		"Value2": ((1388, LCID, 4, 0),()),
		"VerticalAlignment": ((137, LCID, 4, 0),()),
		"WrapText": ((276, LCID, 4, 0),()),
		"_Default": ((0, LCID, 4, 0),()),
	}
	# Default method for this class is '_Default'
	def __call__(self, RowIndex=defaultNamedOptArg, ColumnIndex=defaultNamedOptArg):
		return self._ApplyTypes_(0, 2, (12, 0), ((12, 17), (12, 17)), '__call__', None,RowIndex
			, ColumnIndex)

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

win32com.client.CLSIDToClass.RegisterCLSID( "{00020846-0000-0000-C000-000000000046}", Range )
