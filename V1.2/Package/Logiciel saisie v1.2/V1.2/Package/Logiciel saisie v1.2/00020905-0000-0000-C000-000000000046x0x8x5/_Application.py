# -*- coding: mbcs -*-
# Created by makepy.py version 0.5.01
# By python version 3.3.1 (v3.3.1:d9893d13c628, Apr  6 2013, 20:25:12) [MSC v.1600 32 bit (Intel)]
# From type library '{00020905-0000-0000-C000-000000000046}'
# On Mon Apr 15 10:41:19 2013
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
class _Application(DispatchBaseClass):
	CLSID = IID('{00020970-0000-0000-C000-000000000046}')
	coclass_clsid = IID('{000209FF-0000-0000-C000-000000000046}')

	def Activate(self):
		return self._oleobj_.InvokeTypes(385, LCID, 1, (24, 0), (),)

	def AddAddress(self, TagID=defaultNamedNotOptArg, Value=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(321, LCID, 1, (24, 0), ((24584, 1), (24584, 1)),TagID
			, Value)

	def AutomaticChange(self):
		return self._oleobj_.InvokeTypes(330, LCID, 1, (24, 0), (),)

	def BuildKeyCode(self, Arg1=defaultNamedNotOptArg, Arg2=defaultNamedOptArg, Arg3=defaultNamedOptArg, Arg4=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(316, LCID, 1, (3, 0), ((3, 1), (16396, 17), (16396, 17), (16396, 17)),Arg1
			, Arg2, Arg3, Arg4)

	def CentimetersToPoints(self, Centimeters=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(371, LCID, 1, (4, 0), ((4, 1),),Centimeters
			)

	def ChangeFileOpenDirectory(self, Path=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(357, LCID, 1, (24, 0), ((8, 1),),Path
			)

	def CheckGrammar(self, String=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(323, LCID, 1, (11, 0), ((8, 1),),String
			)

	def CheckSpelling(self, Word=defaultNamedNotOptArg, CustomDictionary=defaultNamedOptArg, IgnoreUppercase=defaultNamedOptArg, MainDictionary=defaultNamedOptArg
			, CustomDictionary2=defaultNamedOptArg, CustomDictionary3=defaultNamedOptArg, CustomDictionary4=defaultNamedOptArg, CustomDictionary5=defaultNamedOptArg, CustomDictionary6=defaultNamedOptArg
			, CustomDictionary7=defaultNamedOptArg, CustomDictionary8=defaultNamedOptArg, CustomDictionary9=defaultNamedOptArg, CustomDictionary10=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(324, LCID, 1, (11, 0), ((8, 1), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),Word
			, CustomDictionary, IgnoreUppercase, MainDictionary, CustomDictionary2, CustomDictionary3
			, CustomDictionary4, CustomDictionary5, CustomDictionary6, CustomDictionary7, CustomDictionary8
			, CustomDictionary9, CustomDictionary10)

	def CleanString(self, String=defaultNamedNotOptArg):
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(354, LCID, 1, (8, 0), ((8, 1),),String
			)

	# Result is of type Document
	def CompareDocuments(self, OriginalDocument=defaultNamedNotOptArg, RevisedDocument=defaultNamedNotOptArg, Destination=2, Granularity=1
			, CompareFormatting=True, CompareCaseChanges=True, CompareWhitespace=True, CompareTables=True, CompareHeaders=True
			, CompareFootnotes=True, CompareTextboxes=True, CompareFields=True, CompareComments=True, CompareMoves=True
			, RevisedAuthor='', IgnoreAllComparisonWarnings=False):
		return self._ApplyTypes_(470, 1, (13, 32), ((13, 1), (13, 1), (3, 49), (3, 49), (11, 49), (11, 49), (11, 49), (11, 49), (11, 49), (11, 49), (11, 49), (11, 49), (11, 49), (11, 49), (8, 49), (11, 49)), 'CompareDocuments', '{00020906-0000-0000-C000-000000000046}',OriginalDocument
			, RevisedDocument, Destination, Granularity, CompareFormatting, CompareCaseChanges
			, CompareWhitespace, CompareTables, CompareHeaders, CompareFootnotes, CompareTextboxes
			, CompareFields, CompareComments, CompareMoves, RevisedAuthor, IgnoreAllComparisonWarnings
			)

	def DDEExecute(self, Channel=defaultNamedNotOptArg, Command=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(310, LCID, 1, (24, 0), ((3, 1), (8, 1)),Channel
			, Command)

	def DDEInitiate(self, App=defaultNamedNotOptArg, Topic=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(311, LCID, 1, (3, 0), ((8, 1), (8, 1)),App
			, Topic)

	def DDEPoke(self, Channel=defaultNamedNotOptArg, Item=defaultNamedNotOptArg, Data=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(312, LCID, 1, (24, 0), ((3, 1), (8, 1), (8, 1)),Channel
			, Item, Data)

	def DDERequest(self, Channel=defaultNamedNotOptArg, Item=defaultNamedNotOptArg):
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(313, LCID, 1, (8, 0), ((3, 1), (8, 1)),Channel
			, Item)

	def DDETerminate(self, Channel=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(314, LCID, 1, (24, 0), ((3, 1),),Channel
			)

	def DDETerminateAll(self):
		return self._oleobj_.InvokeTypes(315, LCID, 1, (24, 0), (),)

	# Result is of type DefaultWebOptions
	def DefaultWebOptions(self):
		ret = self._oleobj_.InvokeTypes(405, LCID, 1, (9, 0), (),)
		if ret is not None:
			ret = Dispatch(ret, 'DefaultWebOptions', '{000209E3-0000-0000-C000-000000000046}')
		return ret

	def DiscussionSupport(self, Range=defaultNamedNotOptArg, cid=defaultNamedNotOptArg, piCSE=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(407, LCID, 1, (24, 0), ((16396, 1), (16396, 1), (16396, 1)),Range
			, cid, piCSE)

	def Dummy2(self):
		return self._oleobj_.InvokeTypes(458, LCID, 1, (11, 0), (),)

	def Dummy4(self):
		return self._oleobj_.InvokeTypes(485, LCID, 1, (24, 0), (),)

	# Result is of type FileDialog
	# The method FileDialog is actually a property, but must be used as a method to correctly pass the arguments
	def FileDialog(self, FileDialogType=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(450, LCID, 2, (9, 0), ((3, 1),),FileDialogType
			)
		if ret is not None:
			ret = Dispatch(ret, 'FileDialog', '{000C0362-0000-0000-C000-000000000046}')
		return ret

	# Result is of type KeyBinding
	# The method FindKey is actually a property, but must be used as a method to correctly pass the arguments
	def FindKey(self, KeyCode=defaultNamedNotOptArg, KeyCode2=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(71, LCID, 2, (9, 0), ((3, 1), (16396, 17)),KeyCode
			, KeyCode2)
		if ret is not None:
			ret = Dispatch(ret, 'FindKey', '{00020998-0000-0000-C000-000000000046}')
		return ret

	def GetAddress(self, Name=defaultNamedOptArg, AddressProperties=defaultNamedOptArg, UseAutoText=defaultNamedOptArg, DisplaySelectDialog=defaultNamedOptArg
			, SelectDialog=defaultNamedOptArg, CheckNamesDialog=defaultNamedOptArg, RecentAddressesChoice=defaultNamedOptArg, UpdateRecentAddresses=defaultNamedOptArg):
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(322, LCID, 1, (8, 0), ((16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),Name
			, AddressProperties, UseAutoText, DisplaySelectDialog, SelectDialog, CheckNamesDialog
			, RecentAddressesChoice, UpdateRecentAddresses)

	def GetDefaultTheme(self, DocumentType=defaultNamedNotOptArg):
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(416, LCID, 1, (8, 0), ((3, 1),),DocumentType
			)

	# Result is of type SpellingSuggestions
	def GetSpellingSuggestions(self, Word=defaultNamedNotOptArg, CustomDictionary=defaultNamedOptArg, IgnoreUppercase=defaultNamedOptArg, MainDictionary=defaultNamedOptArg
			, SuggestionMode=defaultNamedOptArg, CustomDictionary2=defaultNamedOptArg, CustomDictionary3=defaultNamedOptArg, CustomDictionary4=defaultNamedOptArg, CustomDictionary5=defaultNamedOptArg
			, CustomDictionary6=defaultNamedOptArg, CustomDictionary7=defaultNamedOptArg, CustomDictionary8=defaultNamedOptArg, CustomDictionary9=defaultNamedOptArg, CustomDictionary10=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(327, LCID, 1, (9, 0), ((8, 1), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),Word
			, CustomDictionary, IgnoreUppercase, MainDictionary, SuggestionMode, CustomDictionary2
			, CustomDictionary3, CustomDictionary4, CustomDictionary5, CustomDictionary6, CustomDictionary7
			, CustomDictionary8, CustomDictionary9, CustomDictionary10)
		if ret is not None:
			ret = Dispatch(ret, 'GetSpellingSuggestions', '{000209AA-0000-0000-C000-000000000046}')
		return ret

	def GoBack(self):
		return self._oleobj_.InvokeTypes(328, LCID, 1, (24, 0), (),)

	def GoForward(self):
		return self._oleobj_.InvokeTypes(359, LCID, 1, (24, 0), (),)

	def Help(self, HelpType=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(329, LCID, 1, (24, 0), ((16396, 1),),HelpType
			)

	def HelpTool(self):
		return self._oleobj_.InvokeTypes(332, LCID, 1, (24, 0), (),)

	def InchesToPoints(self, Inches=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(370, LCID, 1, (4, 0), ((4, 1),),Inches
			)

	# The method International is actually a property, but must be used as a method to correctly pass the arguments
	def International(self, Index=defaultNamedNotOptArg):
		return self._ApplyTypes_(46, 2, (12, 0), ((3, 1),), 'International', None,Index
			)

	# The method IsObjectValid is actually a property, but must be used as a method to correctly pass the arguments
	def IsObjectValid(self, Object=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(109, LCID, 2, (11, 0), ((9, 1),),Object
			)

	def KeyString(self, KeyCode=defaultNamedNotOptArg, KeyCode2=defaultNamedOptArg):
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(317, LCID, 1, (8, 0), ((3, 1), (16396, 17)),KeyCode
			, KeyCode2)

	def Keyboard(self, LangId=0):
		return self._oleobj_.InvokeTypes(446, LCID, 1, (3, 0), ((3, 49),),LangId
			)

	def KeyboardBidi(self):
		return self._oleobj_.InvokeTypes(401, LCID, 1, (24, 0), (),)

	def KeyboardLatin(self):
		return self._oleobj_.InvokeTypes(400, LCID, 1, (24, 0), (),)

	# Result is of type KeysBoundTo
	# The method KeysBoundTo is actually a property, but must be used as a method to correctly pass the arguments
	def KeysBoundTo(self, KeyCategory=defaultNamedNotOptArg, Command=defaultNamedNotOptArg, CommandParameter=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(70, LCID, 2, (9, 0), ((3, 1), (8, 1), (16396, 17)),KeyCategory
			, Command, CommandParameter)
		if ret is not None:
			ret = Dispatch(ret, 'KeysBoundTo', '{00020997-0000-0000-C000-000000000046}')
		return ret

	def LinesToPoints(self, Lines=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(374, LCID, 1, (4, 0), ((4, 1),),Lines
			)

	def ListCommands(self, ListAllCommands=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(346, LCID, 1, (24, 0), ((11, 1),),ListAllCommands
			)

	def LoadMasterList(self, FileName=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(469, LCID, 1, (24, 0), ((8, 1),),FileName
			)

	def LookupNameProperties(self, Name=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(303, LCID, 1, (24, 0), ((8, 1),),Name
			)

	# Result is of type Document
	def MergeDocuments(self, OriginalDocument=defaultNamedNotOptArg, RevisedDocument=defaultNamedNotOptArg, Destination=2, Granularity=1
			, CompareFormatting=True, CompareCaseChanges=True, CompareWhitespace=True, CompareTables=True, CompareHeaders=True
			, CompareFootnotes=True, CompareTextboxes=True, CompareFields=True, CompareComments=True, CompareMoves=True
			, OriginalAuthor='', RevisedAuthor='', FormatFrom=2):
		return self._ApplyTypes_(471, 1, (13, 32), ((13, 1), (13, 1), (3, 49), (3, 49), (11, 49), (11, 49), (11, 49), (11, 49), (11, 49), (11, 49), (11, 49), (11, 49), (11, 49), (11, 49), (8, 49), (8, 49), (3, 49)), 'MergeDocuments', '{00020906-0000-0000-C000-000000000046}',OriginalDocument
			, RevisedDocument, Destination, Granularity, CompareFormatting, CompareCaseChanges
			, CompareWhitespace, CompareTables, CompareHeaders, CompareFootnotes, CompareTextboxes
			, CompareFields, CompareComments, CompareMoves, OriginalAuthor, RevisedAuthor
			, FormatFrom)

	def MillimetersToPoints(self, Millimeters=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(372, LCID, 1, (4, 0), ((4, 1),),Millimeters
			)

	def MountVolume(self, Zone=defaultNamedNotOptArg, Server=defaultNamedNotOptArg, Volume=defaultNamedNotOptArg, User=defaultNamedOptArg
			, UserPassword=defaultNamedOptArg, VolumePassword=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(353, LCID, 1, (2, 0), ((8, 1), (8, 1), (8, 1), (16396, 17), (16396, 17), (16396, 17)),Zone
			, Server, Volume, User, UserPassword, VolumePassword
			)

	def Move(self, Left=defaultNamedNotOptArg, Top=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(360, LCID, 1, (24, 0), ((3, 1), (3, 1)),Left
			, Top)

	# Result is of type Window
	def NewWindow(self):
		ret = self._oleobj_.InvokeTypes(345, LCID, 1, (9, 0), (),)
		if ret is not None:
			ret = Dispatch(ret, 'NewWindow', '{00020962-0000-0000-C000-000000000046}')
		return ret

	def NextLetter(self):
		return self._oleobj_.InvokeTypes(351, LCID, 1, (24, 0), (),)

	def OnTime(self, When=defaultNamedNotOptArg, Name=defaultNamedNotOptArg, Tolerance=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(350, LCID, 1, (24, 0), ((16396, 1), (8, 1), (16396, 17)),When
			, Name, Tolerance)

	def OrganizerCopy(self, Source=defaultNamedNotOptArg, Destination=defaultNamedNotOptArg, Name=defaultNamedNotOptArg, Object=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(318, LCID, 1, (24, 0), ((8, 1), (8, 1), (8, 1), (3, 1)),Source
			, Destination, Name, Object)

	def OrganizerDelete(self, Source=defaultNamedNotOptArg, Name=defaultNamedNotOptArg, Object=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(319, LCID, 1, (24, 0), ((8, 1), (8, 1), (3, 1)),Source
			, Name, Object)

	def OrganizerRename(self, Source=defaultNamedNotOptArg, Name=defaultNamedNotOptArg, NewName=defaultNamedNotOptArg, Object=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(320, LCID, 1, (24, 0), ((8, 1), (8, 1), (8, 1), (3, 1)),Source
			, Name, NewName, Object)

	def PicasToPoints(self, Picas=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(373, LCID, 1, (4, 0), ((4, 1),),Picas
			)

	def PixelsToPoints(self, Pixels=defaultNamedNotOptArg, fVertical=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(388, LCID, 1, (4, 0), ((4, 1), (16396, 17)),Pixels
			, fVertical)

	def PointsToCentimeters(self, Points=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(381, LCID, 1, (4, 0), ((4, 1),),Points
			)

	def PointsToInches(self, Points=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(380, LCID, 1, (4, 0), ((4, 1),),Points
			)

	def PointsToLines(self, Points=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(384, LCID, 1, (4, 0), ((4, 1),),Points
			)

	def PointsToMillimeters(self, Points=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(382, LCID, 1, (4, 0), ((4, 1),),Points
			)

	def PointsToPicas(self, Points=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(383, LCID, 1, (4, 0), ((4, 1),),Points
			)

	def PointsToPixels(self, Points=defaultNamedNotOptArg, fVertical=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(387, LCID, 1, (4, 0), ((4, 1), (16396, 17)),Points
			, fVertical)

	def PrintOut(self, Background=defaultNamedOptArg, Append=defaultNamedOptArg, Range=defaultNamedOptArg, OutputFileName=defaultNamedOptArg
			, From=defaultNamedOptArg, To=defaultNamedOptArg, Item=defaultNamedOptArg, Copies=defaultNamedOptArg, Pages=defaultNamedOptArg
			, PageType=defaultNamedOptArg, PrintToFile=defaultNamedOptArg, Collate=defaultNamedOptArg, FileName=defaultNamedOptArg, ActivePrinterMacGX=defaultNamedOptArg
			, ManualDuplexPrint=defaultNamedOptArg, PrintZoomColumn=defaultNamedOptArg, PrintZoomRow=defaultNamedOptArg, PrintZoomPaperWidth=defaultNamedOptArg, PrintZoomPaperHeight=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(448, LCID, 1, (24, 0), ((16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),Background
			, Append, Range, OutputFileName, From, To
			, Item, Copies, Pages, PageType, PrintToFile
			, Collate, FileName, ActivePrinterMacGX, ManualDuplexPrint, PrintZoomColumn
			, PrintZoomRow, PrintZoomPaperWidth, PrintZoomPaperHeight)

	def PrintOut2000(self, Background=defaultNamedOptArg, Append=defaultNamedOptArg, Range=defaultNamedOptArg, OutputFileName=defaultNamedOptArg
			, From=defaultNamedOptArg, To=defaultNamedOptArg, Item=defaultNamedOptArg, Copies=defaultNamedOptArg, Pages=defaultNamedOptArg
			, PageType=defaultNamedOptArg, PrintToFile=defaultNamedOptArg, Collate=defaultNamedOptArg, FileName=defaultNamedOptArg, ActivePrinterMacGX=defaultNamedOptArg
			, ManualDuplexPrint=defaultNamedOptArg, PrintZoomColumn=defaultNamedOptArg, PrintZoomRow=defaultNamedOptArg, PrintZoomPaperWidth=defaultNamedOptArg, PrintZoomPaperHeight=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(444, LCID, 1, (24, 0), ((16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),Background
			, Append, Range, OutputFileName, From, To
			, Item, Copies, Pages, PageType, PrintToFile
			, Collate, FileName, ActivePrinterMacGX, ManualDuplexPrint, PrintZoomColumn
			, PrintZoomRow, PrintZoomPaperWidth, PrintZoomPaperHeight)

	def PrintOutOld(self, Background=defaultNamedOptArg, Append=defaultNamedOptArg, Range=defaultNamedOptArg, OutputFileName=defaultNamedOptArg
			, From=defaultNamedOptArg, To=defaultNamedOptArg, Item=defaultNamedOptArg, Copies=defaultNamedOptArg, Pages=defaultNamedOptArg
			, PageType=defaultNamedOptArg, PrintToFile=defaultNamedOptArg, Collate=defaultNamedOptArg, FileName=defaultNamedOptArg, ActivePrinterMacGX=defaultNamedOptArg
			, ManualDuplexPrint=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(302, LCID, 1, (24, 0), ((16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),Background
			, Append, Range, OutputFileName, From, To
			, Item, Copies, Pages, PageType, PrintToFile
			, Collate, FileName, ActivePrinterMacGX, ManualDuplexPrint)

	def ProductCode(self):
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(404, LCID, 1, (8, 0), (),)

	def PutFocusInMailHeader(self):
		return self._oleobj_.InvokeTypes(464, LCID, 1, (24, 0), (),)

	def Quit(self, SaveChanges=defaultNamedOptArg, OriginalFormat=defaultNamedOptArg, RouteDocument=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1105, LCID, 1, (24, 0), ((16396, 17), (16396, 17), (16396, 17)),SaveChanges
			, OriginalFormat, RouteDocument)

	def Repeat(self, Times=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(305, LCID, 1, (11, 0), ((16396, 17),),Times
			)

	def ResetIgnoreAll(self):
		return self._oleobj_.InvokeTypes(326, LCID, 1, (24, 0), (),)

	def Resize(self, Width=defaultNamedNotOptArg, Height=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(361, LCID, 1, (24, 0), ((3, 1), (3, 1)),Width
			, Height)

	def Run(self, MacroName=defaultNamedNotOptArg, varg1=defaultNamedOptArg, varg2=defaultNamedOptArg, varg3=defaultNamedOptArg
			, varg4=defaultNamedOptArg, varg5=defaultNamedOptArg, varg6=defaultNamedOptArg, varg7=defaultNamedOptArg, varg8=defaultNamedOptArg
			, varg9=defaultNamedOptArg, varg10=defaultNamedOptArg, varg11=defaultNamedOptArg, varg12=defaultNamedOptArg, varg13=defaultNamedOptArg
			, varg14=defaultNamedOptArg, varg15=defaultNamedOptArg, varg16=defaultNamedOptArg, varg17=defaultNamedOptArg, varg18=defaultNamedOptArg
			, varg19=defaultNamedOptArg, varg20=defaultNamedOptArg, varg21=defaultNamedOptArg, varg22=defaultNamedOptArg, varg23=defaultNamedOptArg
			, varg24=defaultNamedOptArg, varg25=defaultNamedOptArg, varg26=defaultNamedOptArg, varg27=defaultNamedOptArg, varg28=defaultNamedOptArg
			, varg29=defaultNamedOptArg, varg30=defaultNamedOptArg):
		return self._ApplyTypes_(445, 1, (12, 0), ((8, 1), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17)), 'Run', None,MacroName
			, varg1, varg2, varg3, varg4, varg5
			, varg6, varg7, varg8, varg9, varg10
			, varg11, varg12, varg13, varg14, varg15
			, varg16, varg17, varg18, varg19, varg20
			, varg21, varg22, varg23, varg24, varg25
			, varg26, varg27, varg28, varg29, varg30
			)

	def RunOld(self, MacroName=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(358, LCID, 1, (24, 0), ((8, 1),),MacroName
			)

	def ScreenRefresh(self):
		return self._oleobj_.InvokeTypes(301, LCID, 1, (24, 0), (),)

	def SendFax(self):
		return self._oleobj_.InvokeTypes(356, LCID, 1, (24, 0), (),)

	def SetDefaultTheme(self, Name=defaultNamedNotOptArg, DocumentType=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(414, LCID, 1, (24, 0), ((8, 1), (3, 1)),Name
			, DocumentType)

	def ShowClipboard(self):
		return self._oleobj_.InvokeTypes(349, LCID, 1, (24, 0), (),)

	def ShowMe(self):
		return self._oleobj_.InvokeTypes(331, LCID, 1, (24, 0), (),)

	def SubstituteFont(self, UnavailableFont=defaultNamedNotOptArg, SubstituteFont=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(304, LCID, 1, (24, 0), ((8, 1), (8, 1)),UnavailableFont
			, SubstituteFont)

	# Result is of type SynonymInfo
	# The method SynonymInfo is actually a property, but must be used as a method to correctly pass the arguments
	def SynonymInfo(self, Word=defaultNamedNotOptArg, LanguageID=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(59, LCID, 2, (9, 0), ((8, 1), (16396, 17)),Word
			, LanguageID)
		if ret is not None:
			ret = Dispatch(ret, 'SynonymInfo', '{0002099B-0000-0000-C000-000000000046}')
		return ret

	def ThreeWayMerge(self, LocalDocument=defaultNamedNotOptArg, ServerDocument=defaultNamedNotOptArg, BaseDocument=defaultNamedNotOptArg, FavorSource=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(484, LCID, 1, (24, 0), ((13, 1), (13, 1), (13, 1), (11, 1)),LocalDocument
			, ServerDocument, BaseDocument, FavorSource)

	def ToggleKeyboard(self):
		return self._oleobj_.InvokeTypes(402, LCID, 1, (24, 0), (),)

	_prop_map_get_ = {
		# Method 'ActiveDocument' returns object of type 'Document'
		"ActiveDocument": (3, 2, (13, 0), (), "ActiveDocument", '{00020906-0000-0000-C000-000000000046}'),
		"ActiveEncryptionSession": (479, 2, (3, 0), (), "ActiveEncryptionSession", None),
		"ActivePrinter": (66, 2, (8, 0), (), "ActivePrinter", None),
		# Method 'ActiveProtectedViewWindow' returns object of type 'ProtectedViewWindow'
		"ActiveProtectedViewWindow": (491, 2, (9, 0), (), "ActiveProtectedViewWindow", '{F743EDD0-9B97-4B09-89CC-77BE19B51481}'),
		# Method 'ActiveWindow' returns object of type 'Window'
		"ActiveWindow": (4, 2, (9, 0), (), "ActiveWindow", '{00020962-0000-0000-C000-000000000046}'),
		# Method 'AddIns' returns object of type 'AddIns'
		"AddIns": (22, 2, (9, 0), (), "AddIns", '{0002097F-0000-0000-C000-000000000046}'),
		# Method 'AnswerWizard' returns object of type 'AnswerWizard'
		"AnswerWizard": (409, 2, (9, 0), (), "AnswerWizard", '{000C0360-0000-0000-C000-000000000046}'),
		# Method 'Application' returns object of type 'Application'
		"Application": (1000, 2, (13, 0), (), "Application", '{000209FF-0000-0000-C000-000000000046}'),
		"ArbitraryXMLSupportAvailable": (465, 2, (11, 0), (), "ArbitraryXMLSupportAvailable", None),
		# Method 'Assistance' returns object of type 'IAssistance'
		"Assistance": (477, 2, (9, 0), (), "Assistance", '{4291224C-DEFE-485B-8E69-6CF8AA85CB76}'),
		# Method 'Assistant' returns object of type 'Assistant'
		"Assistant": (15, 2, (9, 0), (), "Assistant", '{000C0322-0000-0000-C000-000000000046}'),
		# Method 'AutoCaptions' returns object of type 'AutoCaptions'
		"AutoCaptions": (21, 2, (9, 0), (), "AutoCaptions", '{0002097A-0000-0000-C000-000000000046}'),
		# Method 'AutoCorrect' returns object of type 'AutoCorrect'
		"AutoCorrect": (10, 2, (9, 0), (), "AutoCorrect", '{00020949-0000-0000-C000-000000000046}'),
		# Method 'AutoCorrectEmail' returns object of type 'AutoCorrect'
		"AutoCorrectEmail": (456, 2, (9, 0), (), "AutoCorrectEmail", '{00020949-0000-0000-C000-000000000046}'),
		"AutomationSecurity": (449, 2, (3, 0), (), "AutomationSecurity", None),
		"BackgroundPrintingStatus": (86, 2, (3, 0), (), "BackgroundPrintingStatus", None),
		"BackgroundSavingStatus": (85, 2, (3, 0), (), "BackgroundSavingStatus", None),
		# Method 'Bibliography' returns object of type 'Bibliography'
		"Bibliography": (472, 2, (9, 0), (), "Bibliography", '{3834F60F-EE8C-455D-A441-D766675D6D3B}'),
		"BrowseExtraFileTypes": (108, 2, (8, 0), (), "BrowseExtraFileTypes", None),
		# Method 'Browser' returns object of type 'Browser'
		"Browser": (16, 2, (9, 0), (), "Browser", '{0002092E-0000-0000-C000-000000000046}'),
		"Build": (47, 2, (8, 0), (), "Build", None),
		"BuildFeatureCrew": (467, 2, (8, 0), (), "BuildFeatureCrew", None),
		"BuildFull": (466, 2, (8, 0), (), "BuildFull", None),
		# Method 'COMAddIns' returns object of type 'COMAddIns'
		"COMAddIns": (111, 2, (9, 0), (), "COMAddIns", '{000C0339-0000-0000-C000-000000000046}'),
		"CapsLock": (48, 2, (11, 0), (), "CapsLock", None),
		"Caption": (80, 2, (8, 0), (), "Caption", None),
		# Method 'CaptionLabels' returns object of type 'CaptionLabels'
		"CaptionLabels": (20, 2, (9, 0), (), "CaptionLabels", '{00020978-0000-0000-C000-000000000046}'),
		"CheckLanguage": (112, 2, (11, 0), (), "CheckLanguage", None),
		# Method 'CommandBars' returns object of type 'CommandBars'
		"CommandBars": (57, 2, (13, 0), (), "CommandBars", '{55F88893-7708-11D1-ACEB-006008961DA5}'),
		"Creator": (1001, 2, (3, 0), (), "Creator", None),
		# Method 'CustomDictionaries' returns object of type 'Dictionaries'
		"CustomDictionaries": (95, 2, (9, 0), (), "CustomDictionaries", '{000209AC-0000-0000-C000-000000000046}'),
		"CustomizationContext": (68, 2, (9, 0), (), "CustomizationContext", None),
		"DefaultLegalBlackline": (459, 2, (11, 0), (), "DefaultLegalBlackline", None),
		"DefaultSaveFormat": (64, 2, (8, 0), (), "DefaultSaveFormat", None),
		"DefaultTableSeparator": (105, 2, (8, 0), (), "DefaultTableSeparator", None),
		# Method 'Dialogs' returns object of type 'Dialogs'
		"Dialogs": (19, 2, (9, 0), (), "Dialogs", '{00020910-0000-0000-C000-000000000046}'),
		"DisplayAlerts": (94, 2, (3, 0), (), "DisplayAlerts", None),
		"DisplayAutoCompleteTips": (92, 2, (11, 0), (), "DisplayAutoCompleteTips", None),
		"DisplayDocumentInformationPanel": (476, 2, (11, 0), (), "DisplayDocumentInformationPanel", None),
		"DisplayRecentFiles": (56, 2, (11, 0), (), "DisplayRecentFiles", None),
		"DisplayScreenTips": (99, 2, (11, 0), (), "DisplayScreenTips", None),
		"DisplayScrollBars": (82, 2, (11, 0), (), "DisplayScrollBars", None),
		"DisplayStatusBar": (29, 2, (11, 0), (), "DisplayStatusBar", None),
		# Method 'Documents' returns object of type 'Documents'
		"Documents": (6, 2, (9, 0), (), "Documents", '{0002096C-0000-0000-C000-000000000046}'),
		"DontResetInsertionPointProperties": (480, 2, (11, 0), (), "DontResetInsertionPointProperties", None),
		"Dummy1": (406, 2, (11, 0), (), "Dummy1", None),
		# Method 'EmailOptions' returns object of type 'EmailOptions'
		"EmailOptions": (389, 2, (9, 0), (), "EmailOptions", '{000209DB-0000-0000-C000-000000000046}'),
		"EmailTemplate": (451, 2, (8, 0), (), "EmailTemplate", None),
		"EnableCancelKey": (100, 2, (3, 0), (), "EnableCancelKey", None),
		"FeatureInstall": (447, 2, (3, 0), (), "FeatureInstall", None),
		# Method 'FileConverters' returns object of type 'FileConverters'
		"FileConverters": (17, 2, (9, 0), (), "FileConverters", '{0002099A-0000-0000-C000-000000000046}'),
		# Method 'FileSearch' returns object of type 'FileSearch'
		"FileSearch": (103, 2, (9, 0), (), "FileSearch", '{000C0332-0000-0000-C000-000000000046}'),
		"FileValidation": (493, 2, (3, 0), (), "FileValidation", None),
		"FocusInMailHeader": (386, 2, (11, 0), (), "FocusInMailHeader", None),
		# Method 'FontNames' returns object of type 'FontNames'
		"FontNames": (11, 2, (9, 0), (), "FontNames", '{0002096F-0000-0000-C000-000000000046}'),
		# Method 'HangulHanjaDictionaries' returns object of type 'HangulHanjaConversionDictionaries'
		"HangulHanjaDictionaries": (110, 2, (9, 0), (), "HangulHanjaDictionaries", '{000209E0-0000-0000-C000-000000000046}'),
		"Height": (90, 2, (3, 0), (), "Height", None),
		"IsSandboxed": (492, 2, (11, 0), (), "IsSandboxed", None),
		# Method 'KeyBindings' returns object of type 'KeyBindings'
		"KeyBindings": (69, 2, (9, 0), (), "KeyBindings", '{00020996-0000-0000-C000-000000000046}'),
		# Method 'LandscapeFontNames' returns object of type 'FontNames'
		"LandscapeFontNames": (12, 2, (9, 0), (), "LandscapeFontNames", '{0002096F-0000-0000-C000-000000000046}'),
		"Language": (391, 2, (3, 0), (), "Language", None),
		# Method 'LanguageSettings' returns object of type 'LanguageSettings'
		"LanguageSettings": (403, 2, (9, 0), (), "LanguageSettings", '{000C0353-0000-0000-C000-000000000046}'),
		# Method 'Languages' returns object of type 'Languages'
		"Languages": (14, 2, (9, 0), (), "Languages", '{0002096E-0000-0000-C000-000000000046}'),
		"Left": (87, 2, (3, 0), (), "Left", None),
		# Method 'ListGalleries' returns object of type 'ListGalleries'
		"ListGalleries": (65, 2, (9, 0), (), "ListGalleries", '{00020995-0000-0000-C000-000000000046}'),
		"MAPIAvailable": (98, 2, (11, 0), (), "MAPIAvailable", None),
		"MacroContainer": (55, 2, (9, 0), (), "MacroContainer", None),
		# Method 'MailMessage' returns object of type 'MailMessage'
		"MailMessage": (348, 2, (9, 0), (), "MailMessage", '{000209BA-0000-0000-C000-000000000046}'),
		"MailSystem": (104, 2, (3, 0), (), "MailSystem", None),
		# Method 'MailingLabel' returns object of type 'MailingLabel'
		"MailingLabel": (18, 2, (9, 0), (), "MailingLabel", '{00020917-0000-0000-C000-000000000046}'),
		"MathCoprocessorAvailable": (36, 2, (11, 0), (), "MathCoprocessorAvailable", None),
		"MouseAvailable": (37, 2, (11, 0), (), "MouseAvailable", None),
		"Name": (0, 2, (8, 0), (), "Name", None),
		# Method 'NewDocument' returns object of type 'NewFile'
		"NewDocument": (454, 2, (9, 0), (), "NewDocument", '{000C0936-0000-0000-C000-000000000046}'),
		# Method 'NormalTemplate' returns object of type 'Template'
		"NormalTemplate": (8, 2, (9, 0), (), "NormalTemplate", '{0002096A-0000-0000-C000-000000000046}'),
		"NumLock": (49, 2, (11, 0), (), "NumLock", None),
		# Method 'OMathAutoCorrect' returns object of type 'OMathAutoCorrect'
		"OMathAutoCorrect": (475, 2, (9, 0), (), "OMathAutoCorrect", '{6F9D1F68-06F7-49EF-8902-185E54EB5E87}'),
		"OpenAttachmentsInFullScreen": (478, 2, (11, 0), (), "OpenAttachmentsInFullScreen", None),
		# Method 'Options' returns object of type 'Options'
		"Options": (93, 2, (9, 0), (), "Options", '{000209B7-0000-0000-C000-000000000046}'),
		"Parent": (1002, 2, (9, 0), (), "Parent", None),
		"Path": (81, 2, (8, 0), (), "Path", None),
		"PathSeparator": (96, 2, (8, 0), (), "PathSeparator", None),
		# Method 'PickerDialog' returns object of type 'PickerDialog'
		"PickerDialog": (489, 2, (9, 0), (), "PickerDialog", '{000C03E6-0000-0000-C000-000000000046}'),
		# Method 'PortraitFontNames' returns object of type 'FontNames'
		"PortraitFontNames": (13, 2, (9, 0), (), "PortraitFontNames", '{0002096F-0000-0000-C000-000000000046}'),
		"PrintPreview": (27, 2, (11, 0), (), "PrintPreview", None),
		# Method 'ProtectedViewWindows' returns object of type 'ProtectedViewWindows'
		"ProtectedViewWindows": (490, 2, (9, 0), (), "ProtectedViewWindows", '{FD0A74E8-C719-49F6-BA1B-F6D9839D1AB9}'),
		# Method 'RecentFiles' returns object of type 'RecentFiles'
		"RecentFiles": (7, 2, (9, 0), (), "RecentFiles", '{00020963-0000-0000-C000-000000000046}'),
		"RestrictLinkedStyles": (474, 2, (11, 0), (), "RestrictLinkedStyles", None),
		"ScreenUpdating": (26, 2, (11, 0), (), "ScreenUpdating", None),
		# Method 'Selection' returns object of type 'Selection'
		"Selection": (5, 2, (9, 0), (), "Selection", '{00020975-0000-0000-C000-000000000046}'),
		"ShowStartupDialog": (455, 2, (11, 0), (), "ShowStartupDialog", None),
		"ShowStylePreviews": (473, 2, (11, 0), (), "ShowStylePreviews", None),
		"ShowVisualBasicEditor": (106, 2, (11, 0), (), "ShowVisualBasicEditor", None),
		"ShowWindowsInTaskbar": (452, 2, (11, 0), (), "ShowWindowsInTaskbar", None),
		# Method 'SmartArtColors' returns object of type 'SmartArtColors'
		"SmartArtColors": (483, 2, (9, 0), (), "SmartArtColors", '{000C03CD-0000-0000-C000-000000000046}'),
		# Method 'SmartArtLayouts' returns object of type 'SmartArtLayouts'
		"SmartArtLayouts": (481, 2, (9, 0), (), "SmartArtLayouts", '{000C03C9-0000-0000-C000-000000000046}'),
		# Method 'SmartArtQuickStyles' returns object of type 'SmartArtQuickStyles'
		"SmartArtQuickStyles": (482, 2, (9, 0), (), "SmartArtQuickStyles", '{000C03CB-0000-0000-C000-000000000046}'),
		# Method 'SmartTagRecognizers' returns object of type 'SmartTagRecognizers'
		"SmartTagRecognizers": (460, 2, (9, 0), (), "SmartTagRecognizers", '{F2B60A10-DED5-46FB-A914-3C6F4EBB6451}'),
		# Method 'SmartTagTypes' returns object of type 'SmartTagTypes'
		"SmartTagTypes": (461, 2, (9, 0), (), "SmartTagTypes", '{DB8E3072-E106-4453-8E7C-53056F1D30DC}'),
		"SpecialMode": (30, 2, (11, 0), (), "SpecialMode", None),
		"StartupPath": (83, 2, (8, 0), (), "StartupPath", None),
		# Method 'System' returns object of type 'System'
		"System": (9, 2, (9, 0), (), "System", '{00020935-0000-0000-C000-000000000046}'),
		# Method 'TaskPanes' returns object of type 'TaskPanes'
		"TaskPanes": (457, 2, (9, 0), (), "TaskPanes", '{E6AAEC05-E543-4085-BA92-9BF7D2474F5C}'),
		# Method 'Tasks' returns object of type 'Tasks'
		"Tasks": (28, 2, (9, 0), (), "Tasks", '{00020983-0000-0000-C000-000000000046}'),
		# Method 'Templates' returns object of type 'Templates'
		"Templates": (67, 2, (9, 0), (), "Templates", '{000209A2-0000-0000-C000-000000000046}'),
		"Top": (88, 2, (3, 0), (), "Top", None),
		# Method 'UndoRecord' returns object of type 'UndoRecord'
		"UndoRecord": (486, 2, (9, 0), (), "UndoRecord", '{E598E358-2852-42D4-8775-160BD91B7244}'),
		"UsableHeight": (34, 2, (3, 0), (), "UsableHeight", None),
		"UsableWidth": (33, 2, (3, 0), (), "UsableWidth", None),
		"UserAddress": (54, 2, (8, 0), (), "UserAddress", None),
		"UserControl": (101, 2, (11, 0), (), "UserControl", None),
		"UserInitials": (53, 2, (8, 0), (), "UserInitials", None),
		"UserName": (52, 2, (8, 0), (), "UserName", None),
		# Method 'VBE' returns object of type 'VBE'
		"VBE": (61, 2, (9, 0), (), "VBE", '{0002E166-0000-0000-C000-000000000046}'),
		"Version": (24, 2, (8, 0), (), "Version", None),
		"Visible": (23, 2, (11, 0), (), "Visible", None),
		"Width": (89, 2, (3, 0), (), "Width", None),
		"WindowState": (91, 2, (3, 0), (), "WindowState", None),
		# Method 'Windows' returns object of type 'Windows'
		"Windows": (2, 2, (9, 0), (), "Windows", '{00020961-0000-0000-C000-000000000046}'),
		"WordBasic": (1, 2, (9, 0), (), "WordBasic", None),
		# Method 'XMLNamespaces' returns object of type 'XMLNamespaces'
		"XMLNamespaces": (463, 2, (9, 0), (), "XMLNamespaces", '{656BBED7-E82D-4B0A-8F97-EC742BA11FFA}'),
	}
	_prop_map_put_ = {
		"ActivePrinter": ((66, LCID, 4, 0),()),
		"AutomationSecurity": ((449, LCID, 4, 0),()),
		"BrowseExtraFileTypes": ((108, LCID, 4, 0),()),
		"Caption": ((80, LCID, 4, 0),()),
		"CheckLanguage": ((112, LCID, 4, 0),()),
		"CustomizationContext": ((68, LCID, 4, 0),()),
		"DefaultLegalBlackline": ((459, LCID, 4, 0),()),
		"DefaultSaveFormat": ((64, LCID, 4, 0),()),
		"DefaultTableSeparator": ((105, LCID, 4, 0),()),
		"DisplayAlerts": ((94, LCID, 4, 0),()),
		"DisplayAutoCompleteTips": ((92, LCID, 4, 0),()),
		"DisplayDocumentInformationPanel": ((476, LCID, 4, 0),()),
		"DisplayRecentFiles": ((56, LCID, 4, 0),()),
		"DisplayScreenTips": ((99, LCID, 4, 0),()),
		"DisplayScrollBars": ((82, LCID, 4, 0),()),
		"DisplayStatusBar": ((29, LCID, 4, 0),()),
		"DontResetInsertionPointProperties": ((480, LCID, 4, 0),()),
		"EmailTemplate": ((451, LCID, 4, 0),()),
		"EnableCancelKey": ((100, LCID, 4, 0),()),
		"FeatureInstall": ((447, LCID, 4, 0),()),
		"FileValidation": ((493, LCID, 4, 0),()),
		"Height": ((90, LCID, 4, 0),()),
		"Left": ((87, LCID, 4, 0),()),
		"OpenAttachmentsInFullScreen": ((478, LCID, 4, 0),()),
		"PrintPreview": ((27, LCID, 4, 0),()),
		"RestrictLinkedStyles": ((474, LCID, 4, 0),()),
		"ScreenUpdating": ((26, LCID, 4, 0),()),
		"ShowStartupDialog": ((455, LCID, 4, 0),()),
		"ShowStylePreviews": ((473, LCID, 4, 0),()),
		"ShowVisualBasicEditor": ((106, LCID, 4, 0),()),
		"ShowWindowsInTaskbar": ((452, LCID, 4, 0),()),
		"StartupPath": ((83, LCID, 4, 0),()),
		"StatusBar": ((97, LCID, 4, 0),()),
		"Top": ((88, LCID, 4, 0),()),
		"UserAddress": ((54, LCID, 4, 0),()),
		"UserInitials": ((53, LCID, 4, 0),()),
		"UserName": ((52, LCID, 4, 0),()),
		"Visible": ((23, LCID, 4, 0),()),
		"Width": ((89, LCID, 4, 0),()),
		"WindowState": ((91, LCID, 4, 0),()),
	}
	# Default property for this class is 'Name'
	def __call__(self):
		return self._ApplyTypes_(*(0, 2, (8, 0), (), "Name", None))
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

win32com.client.CLSIDToClass.RegisterCLSID( "{00020970-0000-0000-C000-000000000046}", _Application )
# -*- coding: mbcs -*-
# Created by makepy.py version 0.5.01
# By python version 3.3.1 (v3.3.1:d9893d13c628, Apr  6 2013, 20:25:12) [MSC v.1600 32 bit (Intel)]
# From type library '{00020905-0000-0000-C000-000000000046}'
# On Mon Apr 15 10:41:19 2013
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

_Application_vtables_dispatch_ = 1
_Application_vtables_ = [
	(( 'Application' , 'prop' , ), 1000, (1000, (), [ (16397, 10, None, "IID('{000209FF-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 28 , (3, 0, None, None) , 0 , )),
	(( 'Creator' , 'prop' , ), 1001, (1001, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 32 , (3, 0, None, None) , 0 , )),
	(( 'Parent' , 'prop' , ), 1002, (1002, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 36 , (3, 0, None, None) , 0 , )),
	(( 'Name' , 'prop' , ), 0, (0, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 40 , (3, 0, None, None) , 0 , )),
	(( 'Documents' , 'prop' , ), 6, (6, (), [ (16393, 10, None, "IID('{0002096C-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 44 , (3, 0, None, None) , 0 , )),
	(( 'Windows' , 'prop' , ), 2, (2, (), [ (16393, 10, None, "IID('{00020961-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 48 , (3, 0, None, None) , 0 , )),
	(( 'ActiveDocument' , 'prop' , ), 3, (3, (), [ (16397, 10, None, "IID('{00020906-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 52 , (3, 0, None, None) , 0 , )),
	(( 'ActiveWindow' , 'prop' , ), 4, (4, (), [ (16393, 10, None, "IID('{00020962-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 56 , (3, 0, None, None) , 0 , )),
	(( 'Selection' , 'prop' , ), 5, (5, (), [ (16393, 10, None, "IID('{00020975-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 60 , (3, 0, None, None) , 0 , )),
	(( 'WordBasic' , 'prop' , ), 1, (1, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 64 , (3, 0, None, None) , 0 , )),
	(( 'RecentFiles' , 'prop' , ), 7, (7, (), [ (16393, 10, None, "IID('{00020963-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 68 , (3, 0, None, None) , 0 , )),
	(( 'NormalTemplate' , 'prop' , ), 8, (8, (), [ (16393, 10, None, "IID('{0002096A-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 72 , (3, 0, None, None) , 0 , )),
	(( 'System' , 'prop' , ), 9, (9, (), [ (16393, 10, None, "IID('{00020935-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 76 , (3, 0, None, None) , 0 , )),
	(( 'AutoCorrect' , 'prop' , ), 10, (10, (), [ (16393, 10, None, "IID('{00020949-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 80 , (3, 0, None, None) , 0 , )),
	(( 'FontNames' , 'prop' , ), 11, (11, (), [ (16393, 10, None, "IID('{0002096F-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 84 , (3, 0, None, None) , 0 , )),
	(( 'LandscapeFontNames' , 'prop' , ), 12, (12, (), [ (16393, 10, None, "IID('{0002096F-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 88 , (3, 0, None, None) , 0 , )),
	(( 'PortraitFontNames' , 'prop' , ), 13, (13, (), [ (16393, 10, None, "IID('{0002096F-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 92 , (3, 0, None, None) , 0 , )),
	(( 'Languages' , 'prop' , ), 14, (14, (), [ (16393, 10, None, "IID('{0002096E-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 96 , (3, 0, None, None) , 0 , )),
	(( 'Assistant' , 'prop' , ), 15, (15, (), [ (16393, 10, None, "IID('{000C0322-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 100 , (3, 0, None, None) , 64 , )),
	(( 'Browser' , 'prop' , ), 16, (16, (), [ (16393, 10, None, "IID('{0002092E-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 104 , (3, 0, None, None) , 0 , )),
	(( 'FileConverters' , 'prop' , ), 17, (17, (), [ (16393, 10, None, "IID('{0002099A-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 108 , (3, 0, None, None) , 0 , )),
	(( 'MailingLabel' , 'prop' , ), 18, (18, (), [ (16393, 10, None, "IID('{00020917-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 112 , (3, 0, None, None) , 0 , )),
	(( 'Dialogs' , 'prop' , ), 19, (19, (), [ (16393, 10, None, "IID('{00020910-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 116 , (3, 0, None, None) , 0 , )),
	(( 'CaptionLabels' , 'prop' , ), 20, (20, (), [ (16393, 10, None, "IID('{00020978-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 120 , (3, 0, None, None) , 0 , )),
	(( 'AutoCaptions' , 'prop' , ), 21, (21, (), [ (16393, 10, None, "IID('{0002097A-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 124 , (3, 0, None, None) , 0 , )),
	(( 'AddIns' , 'prop' , ), 22, (22, (), [ (16393, 10, None, "IID('{0002097F-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 128 , (3, 0, None, None) , 0 , )),
	(( 'Visible' , 'prop' , ), 23, (23, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 132 , (3, 0, None, None) , 0 , )),
	(( 'Visible' , 'prop' , ), 23, (23, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 136 , (3, 0, None, None) , 0 , )),
	(( 'Version' , 'prop' , ), 24, (24, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 140 , (3, 0, None, None) , 0 , )),
	(( 'ScreenUpdating' , 'prop' , ), 26, (26, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 144 , (3, 0, None, None) , 0 , )),
	(( 'ScreenUpdating' , 'prop' , ), 26, (26, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 148 , (3, 0, None, None) , 0 , )),
	(( 'PrintPreview' , 'prop' , ), 27, (27, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 152 , (3, 0, None, None) , 0 , )),
	(( 'PrintPreview' , 'prop' , ), 27, (27, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 156 , (3, 0, None, None) , 0 , )),
	(( 'Tasks' , 'prop' , ), 28, (28, (), [ (16393, 10, None, "IID('{00020983-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 160 , (3, 0, None, None) , 0 , )),
	(( 'DisplayStatusBar' , 'prop' , ), 29, (29, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 164 , (3, 0, None, None) , 64 , )),
	(( 'DisplayStatusBar' , 'prop' , ), 29, (29, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 168 , (3, 0, None, None) , 64 , )),
	(( 'SpecialMode' , 'prop' , ), 30, (30, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 172 , (3, 0, None, None) , 0 , )),
	(( 'UsableWidth' , 'prop' , ), 33, (33, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 176 , (3, 0, None, None) , 0 , )),
	(( 'UsableHeight' , 'prop' , ), 34, (34, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 180 , (3, 0, None, None) , 0 , )),
	(( 'MathCoprocessorAvailable' , 'prop' , ), 36, (36, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 184 , (3, 0, None, None) , 0 , )),
	(( 'MouseAvailable' , 'prop' , ), 37, (37, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 188 , (3, 0, None, None) , 0 , )),
	(( 'International' , 'Index' , 'prop' , ), 46, (46, (), [ (3, 1, None, None) , 
			 (16396, 10, None, None) , ], 1 , 2 , 4 , 0 , 192 , (3, 0, None, None) , 0 , )),
	(( 'Build' , 'prop' , ), 47, (47, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 196 , (3, 0, None, None) , 0 , )),
	(( 'CapsLock' , 'prop' , ), 48, (48, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 200 , (3, 0, None, None) , 0 , )),
	(( 'NumLock' , 'prop' , ), 49, (49, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 204 , (3, 0, None, None) , 0 , )),
	(( 'UserName' , 'prop' , ), 52, (52, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 208 , (3, 0, None, None) , 0 , )),
	(( 'UserName' , 'prop' , ), 52, (52, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 212 , (3, 0, None, None) , 0 , )),
	(( 'UserInitials' , 'prop' , ), 53, (53, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 216 , (3, 0, None, None) , 0 , )),
	(( 'UserInitials' , 'prop' , ), 53, (53, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 220 , (3, 0, None, None) , 0 , )),
	(( 'UserAddress' , 'prop' , ), 54, (54, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 224 , (3, 0, None, None) , 0 , )),
	(( 'UserAddress' , 'prop' , ), 54, (54, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 228 , (3, 0, None, None) , 0 , )),
	(( 'MacroContainer' , 'prop' , ), 55, (55, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 232 , (3, 0, None, None) , 0 , )),
	(( 'DisplayRecentFiles' , 'prop' , ), 56, (56, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 236 , (3, 0, None, None) , 0 , )),
	(( 'DisplayRecentFiles' , 'prop' , ), 56, (56, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 240 , (3, 0, None, None) , 0 , )),
	(( 'CommandBars' , 'prop' , ), 57, (57, (), [ (16397, 10, None, "IID('{55F88893-7708-11D1-ACEB-006008961DA5}')") , ], 1 , 2 , 4 , 0 , 244 , (3, 0, None, None) , 0 , )),
	(( 'SynonymInfo' , 'Word' , 'LanguageID' , 'prop' , ), 59, (59, (), [ 
			 (8, 1, None, None) , (16396, 17, None, None) , (16393, 10, None, "IID('{0002099B-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 1 , 248 , (3, 0, None, None) , 0 , )),
	(( 'VBE' , 'prop' , ), 61, (61, (), [ (16393, 10, None, "IID('{0002E166-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 252 , (3, 0, None, None) , 0 , )),
	(( 'DefaultSaveFormat' , 'prop' , ), 64, (64, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 256 , (3, 0, None, None) , 0 , )),
	(( 'DefaultSaveFormat' , 'prop' , ), 64, (64, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 260 , (3, 0, None, None) , 0 , )),
	(( 'ListGalleries' , 'prop' , ), 65, (65, (), [ (16393, 10, None, "IID('{00020995-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 264 , (3, 0, None, None) , 0 , )),
	(( 'ActivePrinter' , 'prop' , ), 66, (66, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 268 , (3, 0, None, None) , 0 , )),
	(( 'ActivePrinter' , 'prop' , ), 66, (66, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 272 , (3, 0, None, None) , 0 , )),
	(( 'Templates' , 'prop' , ), 67, (67, (), [ (16393, 10, None, "IID('{000209A2-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 276 , (3, 0, None, None) , 0 , )),
	(( 'CustomizationContext' , 'prop' , ), 68, (68, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 280 , (3, 0, None, None) , 0 , )),
	(( 'CustomizationContext' , 'prop' , ), 68, (68, (), [ (9, 1, None, None) , ], 1 , 4 , 4 , 0 , 284 , (3, 0, None, None) , 0 , )),
	(( 'KeyBindings' , 'prop' , ), 69, (69, (), [ (16393, 10, None, "IID('{00020996-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 288 , (3, 0, None, None) , 0 , )),
	(( 'KeysBoundTo' , 'KeyCategory' , 'Command' , 'CommandParameter' , 'prop' , 
			 ), 70, (70, (), [ (3, 1, None, None) , (8, 1, None, None) , (16396, 17, None, None) , (16393, 10, None, "IID('{00020997-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 1 , 292 , (3, 0, None, None) , 0 , )),
	(( 'FindKey' , 'KeyCode' , 'KeyCode2' , 'prop' , ), 71, (71, (), [ 
			 (3, 1, None, None) , (16396, 17, None, None) , (16393, 10, None, "IID('{00020998-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 1 , 296 , (3, 0, None, None) , 0 , )),
	(( 'Caption' , 'prop' , ), 80, (80, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 300 , (3, 0, None, None) , 0 , )),
	(( 'Caption' , 'prop' , ), 80, (80, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 304 , (3, 0, None, None) , 0 , )),
	(( 'Path' , 'prop' , ), 81, (81, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 308 , (3, 0, None, None) , 0 , )),
	(( 'DisplayScrollBars' , 'prop' , ), 82, (82, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 312 , (3, 0, None, None) , 0 , )),
	(( 'DisplayScrollBars' , 'prop' , ), 82, (82, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 316 , (3, 0, None, None) , 0 , )),
	(( 'StartupPath' , 'prop' , ), 83, (83, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 320 , (3, 0, None, None) , 0 , )),
	(( 'StartupPath' , 'prop' , ), 83, (83, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 324 , (3, 0, None, None) , 0 , )),
	(( 'BackgroundSavingStatus' , 'prop' , ), 85, (85, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 328 , (3, 0, None, None) , 0 , )),
	(( 'BackgroundPrintingStatus' , 'prop' , ), 86, (86, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 332 , (3, 0, None, None) , 0 , )),
	(( 'Left' , 'prop' , ), 87, (87, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 336 , (3, 0, None, None) , 0 , )),
	(( 'Left' , 'prop' , ), 87, (87, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 340 , (3, 0, None, None) , 0 , )),
	(( 'Top' , 'prop' , ), 88, (88, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 344 , (3, 0, None, None) , 0 , )),
	(( 'Top' , 'prop' , ), 88, (88, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 348 , (3, 0, None, None) , 0 , )),
	(( 'Width' , 'prop' , ), 89, (89, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 352 , (3, 0, None, None) , 0 , )),
	(( 'Width' , 'prop' , ), 89, (89, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 356 , (3, 0, None, None) , 0 , )),
	(( 'Height' , 'prop' , ), 90, (90, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 360 , (3, 0, None, None) , 0 , )),
	(( 'Height' , 'prop' , ), 90, (90, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 364 , (3, 0, None, None) , 0 , )),
	(( 'WindowState' , 'prop' , ), 91, (91, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 368 , (3, 0, None, None) , 0 , )),
	(( 'WindowState' , 'prop' , ), 91, (91, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 372 , (3, 0, None, None) , 0 , )),
	(( 'DisplayAutoCompleteTips' , 'prop' , ), 92, (92, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 376 , (3, 0, None, None) , 0 , )),
	(( 'DisplayAutoCompleteTips' , 'prop' , ), 92, (92, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 380 , (3, 0, None, None) , 0 , )),
	(( 'Options' , 'prop' , ), 93, (93, (), [ (16393, 10, None, "IID('{000209B7-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 384 , (3, 0, None, None) , 0 , )),
	(( 'DisplayAlerts' , 'prop' , ), 94, (94, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 388 , (3, 0, None, None) , 0 , )),
	(( 'DisplayAlerts' , 'prop' , ), 94, (94, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 392 , (3, 0, None, None) , 0 , )),
	(( 'CustomDictionaries' , 'prop' , ), 95, (95, (), [ (16393, 10, None, "IID('{000209AC-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 396 , (3, 0, None, None) , 0 , )),
	(( 'PathSeparator' , 'prop' , ), 96, (96, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 400 , (3, 0, None, None) , 0 , )),
	(( 'StatusBar' , ), 97, (97, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 404 , (3, 0, None, None) , 0 , )),
	(( 'MAPIAvailable' , 'prop' , ), 98, (98, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 408 , (3, 0, None, None) , 0 , )),
	(( 'DisplayScreenTips' , 'prop' , ), 99, (99, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 412 , (3, 0, None, None) , 0 , )),
	(( 'DisplayScreenTips' , 'prop' , ), 99, (99, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 416 , (3, 0, None, None) , 0 , )),
	(( 'EnableCancelKey' , 'prop' , ), 100, (100, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 420 , (3, 0, None, None) , 0 , )),
	(( 'EnableCancelKey' , 'prop' , ), 100, (100, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 424 , (3, 0, None, None) , 0 , )),
	(( 'UserControl' , 'prop' , ), 101, (101, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 428 , (3, 0, None, None) , 0 , )),
	(( 'FileSearch' , 'prop' , ), 103, (103, (), [ (16393, 10, None, "IID('{000C0332-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 432 , (3, 0, None, None) , 64 , )),
	(( 'MailSystem' , 'prop' , ), 104, (104, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 436 , (3, 0, None, None) , 0 , )),
	(( 'DefaultTableSeparator' , 'prop' , ), 105, (105, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 440 , (3, 0, None, None) , 0 , )),
	(( 'DefaultTableSeparator' , 'prop' , ), 105, (105, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 444 , (3, 0, None, None) , 0 , )),
	(( 'ShowVisualBasicEditor' , 'prop' , ), 106, (106, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 448 , (3, 0, None, None) , 0 , )),
	(( 'ShowVisualBasicEditor' , 'prop' , ), 106, (106, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 452 , (3, 0, None, None) , 0 , )),
	(( 'BrowseExtraFileTypes' , 'prop' , ), 108, (108, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 456 , (3, 0, None, None) , 0 , )),
	(( 'BrowseExtraFileTypes' , 'prop' , ), 108, (108, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 460 , (3, 0, None, None) , 0 , )),
	(( 'IsObjectValid' , 'Object' , 'prop' , ), 109, (109, (), [ (9, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 464 , (3, 0, None, None) , 0 , )),
	(( 'HangulHanjaDictionaries' , 'prop' , ), 110, (110, (), [ (16393, 10, None, "IID('{000209E0-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 468 , (3, 0, None, None) , 0 , )),
	(( 'MailMessage' , 'prop' , ), 348, (348, (), [ (16393, 10, None, "IID('{000209BA-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 472 , (3, 0, None, None) , 0 , )),
	(( 'FocusInMailHeader' , 'prop' , ), 386, (386, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 476 , (3, 0, None, None) , 0 , )),
	(( 'Quit' , 'SaveChanges' , 'OriginalFormat' , 'RouteDocument' , ), 1105, (1105, (), [ 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , ], 1 , 1 , 4 , 3 , 480 , (3, 0, None, None) , 0 , )),
	(( 'ScreenRefresh' , ), 301, (301, (), [ ], 1 , 1 , 4 , 0 , 484 , (3, 0, None, None) , 0 , )),
	(( 'PrintOutOld' , 'Background' , 'Append' , 'Range' , 'OutputFileName' , 
			 'From' , 'To' , 'Item' , 'Copies' , 'Pages' , 
			 'PageType' , 'PrintToFile' , 'Collate' , 'FileName' , 'ActivePrinterMacGX' , 
			 'ManualDuplexPrint' , ), 302, (302, (), [ (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , ], 1 , 1 , 4 , 15 , 488 , (3, 0, None, None) , 64 , )),
	(( 'LookupNameProperties' , 'Name' , ), 303, (303, (), [ (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 492 , (3, 0, None, None) , 0 , )),
	(( 'SubstituteFont' , 'UnavailableFont' , 'SubstituteFont' , ), 304, (304, (), [ (8, 1, None, None) , 
			 (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 496 , (3, 0, None, None) , 0 , )),
	(( 'Repeat' , 'Times' , 'prop' , ), 305, (305, (), [ (16396, 17, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 1 , 500 , (3, 0, None, None) , 0 , )),
	(( 'DDEExecute' , 'Channel' , 'Command' , ), 310, (310, (), [ (3, 1, None, None) , 
			 (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 504 , (3, 0, None, None) , 0 , )),
	(( 'DDEInitiate' , 'App' , 'Topic' , 'prop' , ), 311, (311, (), [ 
			 (8, 1, None, None) , (8, 1, None, None) , (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 508 , (3, 0, None, None) , 0 , )),
	(( 'DDEPoke' , 'Channel' , 'Item' , 'Data' , ), 312, (312, (), [ 
			 (3, 1, None, None) , (8, 1, None, None) , (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 512 , (3, 0, None, None) , 0 , )),
	(( 'DDERequest' , 'Channel' , 'Item' , 'prop' , ), 313, (313, (), [ 
			 (3, 1, None, None) , (8, 1, None, None) , (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 516 , (3, 0, None, None) , 0 , )),
	(( 'DDETerminate' , 'Channel' , ), 314, (314, (), [ (3, 1, None, None) , ], 1 , 1 , 4 , 0 , 520 , (3, 0, None, None) , 0 , )),
	(( 'DDETerminateAll' , ), 315, (315, (), [ ], 1 , 1 , 4 , 0 , 524 , (3, 0, None, None) , 0 , )),
	(( 'BuildKeyCode' , 'Arg1' , 'Arg2' , 'Arg3' , 'Arg4' , 
			 'prop' , ), 316, (316, (), [ (3, 1, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16387, 10, None, None) , ], 1 , 1 , 4 , 3 , 528 , (3, 0, None, None) , 0 , )),
	(( 'KeyString' , 'KeyCode' , 'KeyCode2' , 'prop' , ), 317, (317, (), [ 
			 (3, 1, None, None) , (16396, 17, None, None) , (16392, 10, None, None) , ], 1 , 1 , 4 , 1 , 532 , (3, 0, None, None) , 0 , )),
	(( 'OrganizerCopy' , 'Source' , 'Destination' , 'Name' , 'Object' , 
			 ), 318, (318, (), [ (8, 1, None, None) , (8, 1, None, None) , (8, 1, None, None) , (3, 1, None, None) , ], 1 , 1 , 4 , 0 , 536 , (3, 0, None, None) , 0 , )),
	(( 'OrganizerDelete' , 'Source' , 'Name' , 'Object' , ), 319, (319, (), [ 
			 (8, 1, None, None) , (8, 1, None, None) , (3, 1, None, None) , ], 1 , 1 , 4 , 0 , 540 , (3, 0, None, None) , 0 , )),
	(( 'OrganizerRename' , 'Source' , 'Name' , 'NewName' , 'Object' , 
			 ), 320, (320, (), [ (8, 1, None, None) , (8, 1, None, None) , (8, 1, None, None) , (3, 1, None, None) , ], 1 , 1 , 4 , 0 , 544 , (3, 0, None, None) , 0 , )),
	(( 'AddAddress' , 'TagID' , 'Value' , ), 321, (321, (), [ (24584, 1, None, None) , 
			 (24584, 1, None, None) , ], 1 , 1 , 4 , 0 , 548 , (3, 0, None, None) , 0 , )),
	(( 'GetAddress' , 'Name' , 'AddressProperties' , 'UseAutoText' , 'DisplaySelectDialog' , 
			 'SelectDialog' , 'CheckNamesDialog' , 'RecentAddressesChoice' , 'UpdateRecentAddresses' , 'prop' , 
			 ), 322, (322, (), [ (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16392, 10, None, None) , ], 1 , 1 , 4 , 8 , 552 , (3, 0, None, None) , 0 , )),
	(( 'CheckGrammar' , 'String' , 'prop' , ), 323, (323, (), [ (8, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 556 , (3, 0, None, None) , 0 , )),
	(( 'CheckSpelling' , 'Word' , 'CustomDictionary' , 'IgnoreUppercase' , 'MainDictionary' , 
			 'CustomDictionary2' , 'CustomDictionary3' , 'CustomDictionary4' , 'CustomDictionary5' , 'CustomDictionary6' , 
			 'CustomDictionary7' , 'CustomDictionary8' , 'CustomDictionary9' , 'CustomDictionary10' , 'prop' , 
			 ), 324, (324, (), [ (8, 1, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 12 , 560 , (3, 0, None, None) , 0 , )),
	(( 'ResetIgnoreAll' , ), 326, (326, (), [ ], 1 , 1 , 4 , 0 , 564 , (3, 0, None, None) , 0 , )),
	(( 'GetSpellingSuggestions' , 'Word' , 'CustomDictionary' , 'IgnoreUppercase' , 'MainDictionary' , 
			 'SuggestionMode' , 'CustomDictionary2' , 'CustomDictionary3' , 'CustomDictionary4' , 'CustomDictionary5' , 
			 'CustomDictionary6' , 'CustomDictionary7' , 'CustomDictionary8' , 'CustomDictionary9' , 'CustomDictionary10' , 
			 'prop' , ), 327, (327, (), [ (8, 1, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16393, 10, None, "IID('{000209AA-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 13 , 568 , (3, 0, None, None) , 0 , )),
	(( 'GoBack' , ), 328, (328, (), [ ], 1 , 1 , 4 , 0 , 572 , (3, 0, None, None) , 0 , )),
	(( 'Help' , 'HelpType' , ), 329, (329, (), [ (16396, 1, None, None) , ], 1 , 1 , 4 , 0 , 576 , (3, 0, None, None) , 0 , )),
	(( 'AutomaticChange' , ), 330, (330, (), [ ], 1 , 1 , 4 , 0 , 580 , (3, 0, None, None) , 0 , )),
	(( 'ShowMe' , ), 331, (331, (), [ ], 1 , 1 , 4 , 0 , 584 , (3, 0, None, None) , 0 , )),
	(( 'HelpTool' , ), 332, (332, (), [ ], 1 , 1 , 4 , 0 , 588 , (3, 0, None, None) , 0 , )),
	(( 'NewWindow' , 'prop' , ), 345, (345, (), [ (16393, 10, None, "IID('{00020962-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 0 , 592 , (3, 0, None, None) , 0 , )),
	(( 'ListCommands' , 'ListAllCommands' , ), 346, (346, (), [ (11, 1, None, None) , ], 1 , 1 , 4 , 0 , 596 , (3, 0, None, None) , 0 , )),
	(( 'ShowClipboard' , ), 349, (349, (), [ ], 1 , 1 , 4 , 0 , 600 , (3, 0, None, None) , 0 , )),
	(( 'OnTime' , 'When' , 'Name' , 'Tolerance' , ), 350, (350, (), [ 
			 (16396, 1, None, None) , (8, 1, None, None) , (16396, 17, None, None) , ], 1 , 1 , 4 , 1 , 604 , (3, 0, None, None) , 0 , )),
	(( 'NextLetter' , ), 351, (351, (), [ ], 1 , 1 , 4 , 0 , 608 , (3, 0, None, None) , 0 , )),
	(( 'MountVolume' , 'Zone' , 'Server' , 'Volume' , 'User' , 
			 'UserPassword' , 'VolumePassword' , 'prop' , ), 353, (353, (), [ (8, 1, None, None) , 
			 (8, 1, None, None) , (8, 1, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16386, 10, None, None) , ], 1 , 1 , 4 , 3 , 612 , (3, 0, None, None) , 64 , )),
	(( 'CleanString' , 'String' , 'prop' , ), 354, (354, (), [ (8, 1, None, None) , 
			 (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 616 , (3, 0, None, None) , 0 , )),
	(( 'SendFax' , ), 356, (356, (), [ ], 1 , 1 , 4 , 0 , 620 , (3, 0, None, None) , 64 , )),
	(( 'ChangeFileOpenDirectory' , 'Path' , ), 357, (357, (), [ (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 624 , (3, 0, None, None) , 0 , )),
	(( 'RunOld' , 'MacroName' , ), 358, (358, (), [ (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 628 , (3, 0, None, None) , 64 , )),
	(( 'GoForward' , ), 359, (359, (), [ ], 1 , 1 , 4 , 0 , 632 , (3, 0, None, None) , 0 , )),
	(( 'Move' , 'Left' , 'Top' , ), 360, (360, (), [ (3, 1, None, None) , 
			 (3, 1, None, None) , ], 1 , 1 , 4 , 0 , 636 , (3, 0, None, None) , 0 , )),
	(( 'Resize' , 'Width' , 'Height' , ), 361, (361, (), [ (3, 1, None, None) , 
			 (3, 1, None, None) , ], 1 , 1 , 4 , 0 , 640 , (3, 0, None, None) , 0 , )),
	(( 'InchesToPoints' , 'Inches' , 'prop' , ), 370, (370, (), [ (4, 1, None, None) , 
			 (16388, 10, None, None) , ], 1 , 1 , 4 , 0 , 644 , (3, 0, None, None) , 0 , )),
	(( 'CentimetersToPoints' , 'Centimeters' , 'prop' , ), 371, (371, (), [ (4, 1, None, None) , 
			 (16388, 10, None, None) , ], 1 , 1 , 4 , 0 , 648 , (3, 0, None, None) , 0 , )),
	(( 'MillimetersToPoints' , 'Millimeters' , 'prop' , ), 372, (372, (), [ (4, 1, None, None) , 
			 (16388, 10, None, None) , ], 1 , 1 , 4 , 0 , 652 , (3, 0, None, None) , 0 , )),
	(( 'PicasToPoints' , 'Picas' , 'prop' , ), 373, (373, (), [ (4, 1, None, None) , 
			 (16388, 10, None, None) , ], 1 , 1 , 4 , 0 , 656 , (3, 0, None, None) , 0 , )),
	(( 'LinesToPoints' , 'Lines' , 'prop' , ), 374, (374, (), [ (4, 1, None, None) , 
			 (16388, 10, None, None) , ], 1 , 1 , 4 , 0 , 660 , (3, 0, None, None) , 0 , )),
	(( 'PointsToInches' , 'Points' , 'prop' , ), 380, (380, (), [ (4, 1, None, None) , 
			 (16388, 10, None, None) , ], 1 , 1 , 4 , 0 , 664 , (3, 0, None, None) , 0 , )),
	(( 'PointsToCentimeters' , 'Points' , 'prop' , ), 381, (381, (), [ (4, 1, None, None) , 
			 (16388, 10, None, None) , ], 1 , 1 , 4 , 0 , 668 , (3, 0, None, None) , 0 , )),
	(( 'PointsToMillimeters' , 'Points' , 'prop' , ), 382, (382, (), [ (4, 1, None, None) , 
			 (16388, 10, None, None) , ], 1 , 1 , 4 , 0 , 672 , (3, 0, None, None) , 0 , )),
	(( 'PointsToPicas' , 'Points' , 'prop' , ), 383, (383, (), [ (4, 1, None, None) , 
			 (16388, 10, None, None) , ], 1 , 1 , 4 , 0 , 676 , (3, 0, None, None) , 0 , )),
	(( 'PointsToLines' , 'Points' , 'prop' , ), 384, (384, (), [ (4, 1, None, None) , 
			 (16388, 10, None, None) , ], 1 , 1 , 4 , 0 , 680 , (3, 0, None, None) , 0 , )),
	(( 'Activate' , ), 385, (385, (), [ ], 1 , 1 , 4 , 0 , 684 , (3, 0, None, None) , 0 , )),
	(( 'PointsToPixels' , 'Points' , 'fVertical' , 'prop' , ), 387, (387, (), [ 
			 (4, 1, None, None) , (16396, 17, None, None) , (16388, 10, None, None) , ], 1 , 1 , 4 , 1 , 688 , (3, 0, None, None) , 0 , )),
	(( 'PixelsToPoints' , 'Pixels' , 'fVertical' , 'prop' , ), 388, (388, (), [ 
			 (4, 1, None, None) , (16396, 17, None, None) , (16388, 10, None, None) , ], 1 , 1 , 4 , 1 , 692 , (3, 0, None, None) , 0 , )),
	(( 'KeyboardLatin' , ), 400, (400, (), [ ], 1 , 1 , 4 , 0 , 696 , (3, 0, None, None) , 0 , )),
	(( 'KeyboardBidi' , ), 401, (401, (), [ ], 1 , 1 , 4 , 0 , 700 , (3, 0, None, None) , 0 , )),
	(( 'ToggleKeyboard' , ), 402, (402, (), [ ], 1 , 1 , 4 , 0 , 704 , (3, 0, None, None) , 0 , )),
	(( 'Keyboard' , 'LangId' , 'prop' , ), 446, (446, (), [ (3, 49, '0', None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 708 , (3, 0, None, None) , 0 , )),
	(( 'ProductCode' , 'prop' , ), 404, (404, (), [ (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 712 , (3, 0, None, None) , 0 , )),
	(( 'DefaultWebOptions' , 'prop' , ), 405, (405, (), [ (16393, 10, None, "IID('{000209E3-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 0 , 716 , (3, 0, None, None) , 0 , )),
	(( 'DiscussionSupport' , 'Range' , 'cid' , 'piCSE' , ), 407, (407, (), [ 
			 (16396, 1, None, None) , (16396, 1, None, None) , (16396, 1, None, None) , ], 1 , 1 , 4 , 0 , 720 , (3, 0, None, None) , 64 , )),
	(( 'SetDefaultTheme' , 'Name' , 'DocumentType' , ), 414, (414, (), [ (8, 1, None, None) , 
			 (3, 1, None, None) , ], 1 , 1 , 4 , 0 , 724 , (3, 0, None, None) , 0 , )),
	(( 'GetDefaultTheme' , 'DocumentType' , 'prop' , ), 416, (416, (), [ (3, 1, None, None) , 
			 (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 728 , (3, 0, None, None) , 0 , )),
	(( 'EmailOptions' , 'prop' , ), 389, (389, (), [ (16393, 10, None, "IID('{000209DB-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 732 , (3, 0, None, None) , 0 , )),
	(( 'Language' , 'prop' , ), 391, (391, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 736 , (3, 0, None, None) , 0 , )),
	(( 'COMAddIns' , 'prop' , ), 111, (111, (), [ (16393, 10, None, "IID('{000C0339-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 740 , (3, 0, None, None) , 0 , )),
	(( 'CheckLanguage' , 'prop' , ), 112, (112, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 744 , (3, 0, None, None) , 0 , )),
	(( 'CheckLanguage' , 'prop' , ), 112, (112, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 748 , (3, 0, None, None) , 0 , )),
	(( 'LanguageSettings' , 'prop' , ), 403, (403, (), [ (16393, 10, None, "IID('{000C0353-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 752 , (3, 0, None, None) , 0 , )),
	(( 'Dummy1' , 'prop' , ), 406, (406, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 756 , (3, 0, None, None) , 64 , )),
	(( 'AnswerWizard' , 'prop' , ), 409, (409, (), [ (16393, 10, None, "IID('{000C0360-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 760 , (3, 0, None, None) , 64 , )),
	(( 'FeatureInstall' , 'prop' , ), 447, (447, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 764 , (3, 0, None, None) , 0 , )),
	(( 'FeatureInstall' , 'prop' , ), 447, (447, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 768 , (3, 0, None, None) , 0 , )),
	(( 'PrintOut2000' , 'Background' , 'Append' , 'Range' , 'OutputFileName' , 
			 'From' , 'To' , 'Item' , 'Copies' , 'Pages' , 
			 'PageType' , 'PrintToFile' , 'Collate' , 'FileName' , 'ActivePrinterMacGX' , 
			 'ManualDuplexPrint' , 'PrintZoomColumn' , 'PrintZoomRow' , 'PrintZoomPaperWidth' , 'PrintZoomPaperHeight' , 
			 ), 444, (444, (), [ (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , ], 1 , 1 , 4 , 19 , 772 , (3, 0, None, None) , 64 , )),
	(( 'Run' , 'MacroName' , 'varg1' , 'varg2' , 'varg3' , 
			 'varg4' , 'varg5' , 'varg6' , 'varg7' , 'varg8' , 
			 'varg9' , 'varg10' , 'varg11' , 'varg12' , 'varg13' , 
			 'varg14' , 'varg15' , 'varg16' , 'varg17' , 'varg18' , 
			 'varg19' , 'varg20' , 'varg21' , 'varg22' , 'varg23' , 
			 'varg24' , 'varg25' , 'varg26' , 'varg27' , 'varg28' , 
			 'varg29' , 'varg30' , 'prop' , ), 445, (445, (), [ (8, 1, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 10, None, None) , ], 1 , 1 , 4 , 30 , 776 , (3, 0, None, None) , 0 , )),
	(( 'PrintOut' , 'Background' , 'Append' , 'Range' , 'OutputFileName' , 
			 'From' , 'To' , 'Item' , 'Copies' , 'Pages' , 
			 'PageType' , 'PrintToFile' , 'Collate' , 'FileName' , 'ActivePrinterMacGX' , 
			 'ManualDuplexPrint' , 'PrintZoomColumn' , 'PrintZoomRow' , 'PrintZoomPaperWidth' , 'PrintZoomPaperHeight' , 
			 ), 448, (448, (), [ (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , ], 1 , 1 , 4 , 19 , 780 , (3, 0, None, None) , 0 , )),
	(( 'AutomationSecurity' , 'prop' , ), 449, (449, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 784 , (3, 0, None, None) , 0 , )),
	(( 'AutomationSecurity' , 'prop' , ), 449, (449, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 788 , (3, 0, None, None) , 0 , )),
	(( 'FileDialog' , 'FileDialogType' , 'prop' , ), 450, (450, (), [ (3, 1, None, None) , 
			 (16393, 10, None, "IID('{000C0362-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 792 , (3, 0, None, None) , 0 , )),
	(( 'EmailTemplate' , 'prop' , ), 451, (451, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 796 , (3, 0, None, None) , 0 , )),
	(( 'EmailTemplate' , 'prop' , ), 451, (451, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 800 , (3, 0, None, None) , 0 , )),
	(( 'ShowWindowsInTaskbar' , 'prop' , ), 452, (452, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 804 , (3, 0, None, None) , 0 , )),
	(( 'ShowWindowsInTaskbar' , 'prop' , ), 452, (452, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 808 , (3, 0, None, None) , 0 , )),
	(( 'NewDocument' , 'prop' , ), 454, (454, (), [ (16393, 10, None, "IID('{000C0936-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 812 , (3, 0, None, None) , 0 , )),
	(( 'ShowStartupDialog' , 'prop' , ), 455, (455, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 816 , (3, 0, None, None) , 0 , )),
	(( 'ShowStartupDialog' , 'prop' , ), 455, (455, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 820 , (3, 0, None, None) , 0 , )),
	(( 'AutoCorrectEmail' , 'prop' , ), 456, (456, (), [ (16393, 10, None, "IID('{00020949-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 824 , (3, 0, None, None) , 0 , )),
	(( 'TaskPanes' , 'prop' , ), 457, (457, (), [ (16393, 10, None, "IID('{E6AAEC05-E543-4085-BA92-9BF7D2474F5C}')") , ], 1 , 2 , 4 , 0 , 828 , (3, 0, None, None) , 0 , )),
	(( 'DefaultLegalBlackline' , 'prop' , ), 459, (459, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 832 , (3, 0, None, None) , 0 , )),
	(( 'DefaultLegalBlackline' , 'prop' , ), 459, (459, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 836 , (3, 0, None, None) , 0 , )),
	(( 'Dummy2' , 'prop' , ), 458, (458, (), [ (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 840 , (3, 0, None, None) , 64 , )),
	(( 'SmartTagRecognizers' , 'prop' , ), 460, (460, (), [ (16393, 10, None, "IID('{F2B60A10-DED5-46FB-A914-3C6F4EBB6451}')") , ], 1 , 2 , 4 , 0 , 844 , (3, 0, None, None) , 64 , )),
	(( 'SmartTagTypes' , 'prop' , ), 461, (461, (), [ (16393, 10, None, "IID('{DB8E3072-E106-4453-8E7C-53056F1D30DC}')") , ], 1 , 2 , 4 , 0 , 848 , (3, 0, None, None) , 64 , )),
	(( 'XMLNamespaces' , 'prop' , ), 463, (463, (), [ (16393, 10, None, "IID('{656BBED7-E82D-4B0A-8F97-EC742BA11FFA}')") , ], 1 , 2 , 4 , 0 , 852 , (3, 0, None, None) , 0 , )),
	(( 'PutFocusInMailHeader' , ), 464, (464, (), [ ], 1 , 1 , 4 , 0 , 856 , (3, 0, None, None) , 0 , )),
	(( 'ArbitraryXMLSupportAvailable' , 'prop' , ), 465, (465, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 860 , (3, 0, None, None) , 0 , )),
	(( 'BuildFull' , 'prop' , ), 466, (466, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 864 , (3, 0, None, None) , 64 , )),
	(( 'BuildFeatureCrew' , 'prop' , ), 467, (467, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 868 , (3, 0, None, None) , 64 , )),
	(( 'LoadMasterList' , 'FileName' , ), 469, (469, (), [ (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 872 , (3, 0, None, None) , 0 , )),
	(( 'CompareDocuments' , 'OriginalDocument' , 'RevisedDocument' , 'Destination' , 'Granularity' , 
			 'CompareFormatting' , 'CompareCaseChanges' , 'CompareWhitespace' , 'CompareTables' , 'CompareHeaders' , 
			 'CompareFootnotes' , 'CompareTextboxes' , 'CompareFields' , 'CompareComments' , 'CompareMoves' , 
			 'RevisedAuthor' , 'IgnoreAllComparisonWarnings' , 'prop' , ), 470, (470, (), [ (13, 1, None, "IID('{00020906-0000-0000-C000-000000000046}')") , 
			 (13, 1, None, "IID('{00020906-0000-0000-C000-000000000046}')") , (3, 49, '2', None) , (3, 49, '1', None) , (11, 49, 'True', None) , (11, 49, 'True', None) , 
			 (11, 49, 'True', None) , (11, 49, 'True', None) , (11, 49, 'True', None) , (11, 49, 'True', None) , (11, 49, 'True', None) , 
			 (11, 49, 'True', None) , (11, 49, 'True', None) , (11, 49, 'True', None) , (8, 49, "''", None) , (11, 49, 'False', None) , 
			 (16397, 10, None, "IID('{00020906-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 0 , 876 , (3, 32, None, None) , 0 , )),
	(( 'MergeDocuments' , 'OriginalDocument' , 'RevisedDocument' , 'Destination' , 'Granularity' , 
			 'CompareFormatting' , 'CompareCaseChanges' , 'CompareWhitespace' , 'CompareTables' , 'CompareHeaders' , 
			 'CompareFootnotes' , 'CompareTextboxes' , 'CompareFields' , 'CompareComments' , 'CompareMoves' , 
			 'OriginalAuthor' , 'RevisedAuthor' , 'FormatFrom' , 'prop' , ), 471, (471, (), [ 
			 (13, 1, None, "IID('{00020906-0000-0000-C000-000000000046}')") , (13, 1, None, "IID('{00020906-0000-0000-C000-000000000046}')") , (3, 49, '2', None) , (3, 49, '1', None) , (11, 49, 'True', None) , 
			 (11, 49, 'True', None) , (11, 49, 'True', None) , (11, 49, 'True', None) , (11, 49, 'True', None) , (11, 49, 'True', None) , 
			 (11, 49, 'True', None) , (11, 49, 'True', None) , (11, 49, 'True', None) , (11, 49, 'True', None) , (8, 49, "''", None) , 
			 (8, 49, "''", None) , (3, 49, '2', None) , (16397, 10, None, "IID('{00020906-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 0 , 880 , (3, 32, None, None) , 0 , )),
	(( 'Bibliography' , 'prop' , ), 472, (472, (), [ (16393, 10, None, "IID('{3834F60F-EE8C-455D-A441-D766675D6D3B}')") , ], 1 , 2 , 4 , 0 , 884 , (3, 0, None, None) , 0 , )),
	(( 'ShowStylePreviews' , 'prop' , ), 473, (473, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 888 , (3, 0, None, None) , 0 , )),
	(( 'ShowStylePreviews' , 'prop' , ), 473, (473, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 892 , (3, 0, None, None) , 0 , )),
	(( 'RestrictLinkedStyles' , 'prop' , ), 474, (474, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 896 , (3, 0, None, None) , 0 , )),
	(( 'RestrictLinkedStyles' , 'prop' , ), 474, (474, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 900 , (3, 0, None, None) , 0 , )),
	(( 'OMathAutoCorrect' , 'prop' , ), 475, (475, (), [ (16393, 10, None, "IID('{6F9D1F68-06F7-49EF-8902-185E54EB5E87}')") , ], 1 , 2 , 4 , 0 , 904 , (3, 0, None, None) , 0 , )),
	(( 'DisplayDocumentInformationPanel' , 'prop' , ), 476, (476, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 908 , (3, 0, None, None) , 0 , )),
	(( 'DisplayDocumentInformationPanel' , 'prop' , ), 476, (476, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 912 , (3, 0, None, None) , 0 , )),
	(( 'Assistance' , 'prop' , ), 477, (477, (), [ (16393, 10, None, "IID('{4291224C-DEFE-485B-8E69-6CF8AA85CB76}')") , ], 1 , 2 , 4 , 0 , 916 , (3, 0, None, None) , 0 , )),
	(( 'OpenAttachmentsInFullScreen' , 'prop' , ), 478, (478, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 920 , (3, 0, None, None) , 0 , )),
	(( 'OpenAttachmentsInFullScreen' , 'prop' , ), 478, (478, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 924 , (3, 0, None, None) , 0 , )),
	(( 'ActiveEncryptionSession' , 'prop' , ), 479, (479, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 928 , (3, 0, None, None) , 0 , )),
	(( 'DontResetInsertionPointProperties' , 'prop' , ), 480, (480, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 932 , (3, 0, None, None) , 0 , )),
	(( 'DontResetInsertionPointProperties' , 'prop' , ), 480, (480, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 936 , (3, 0, None, None) , 0 , )),
	(( 'SmartArtLayouts' , 'prop' , ), 481, (481, (), [ (16393, 10, None, "IID('{000C03C9-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 940 , (3, 0, None, None) , 0 , )),
	(( 'SmartArtQuickStyles' , 'prop' , ), 482, (482, (), [ (16393, 10, None, "IID('{000C03CB-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 944 , (3, 0, None, None) , 0 , )),
	(( 'SmartArtColors' , 'prop' , ), 483, (483, (), [ (16393, 10, None, "IID('{000C03CD-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 948 , (3, 0, None, None) , 0 , )),
	(( 'ThreeWayMerge' , 'LocalDocument' , 'ServerDocument' , 'BaseDocument' , 'FavorSource' , 
			 ), 484, (484, (), [ (13, 1, None, "IID('{00020906-0000-0000-C000-000000000046}')") , (13, 1, None, "IID('{00020906-0000-0000-C000-000000000046}')") , (13, 1, None, "IID('{00020906-0000-0000-C000-000000000046}')") , (11, 1, None, None) , ], 1 , 1 , 4 , 0 , 952 , (3, 0, None, None) , 64 , )),
	(( 'Dummy4' , ), 485, (485, (), [ ], 1 , 1 , 4 , 0 , 956 , (3, 0, None, None) , 64 , )),
	(( 'UndoRecord' , 'prop' , ), 486, (486, (), [ (16393, 10, None, "IID('{E598E358-2852-42D4-8775-160BD91B7244}')") , ], 1 , 2 , 4 , 0 , 960 , (3, 0, None, None) , 0 , )),
	(( 'PickerDialog' , 'prop' , ), 489, (489, (), [ (16393, 10, None, "IID('{000C03E6-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 964 , (3, 0, None, None) , 0 , )),
	(( 'ProtectedViewWindows' , 'prop' , ), 490, (490, (), [ (16393, 10, None, "IID('{FD0A74E8-C719-49F6-BA1B-F6D9839D1AB9}')") , ], 1 , 2 , 4 , 0 , 968 , (3, 0, None, None) , 0 , )),
	(( 'ActiveProtectedViewWindow' , 'prop' , ), 491, (491, (), [ (16393, 10, None, "IID('{F743EDD0-9B97-4B09-89CC-77BE19B51481}')") , ], 1 , 2 , 4 , 0 , 972 , (3, 0, None, None) , 0 , )),
	(( 'IsSandboxed' , 'prop' , ), 492, (492, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 976 , (3, 0, None, None) , 0 , )),
	(( 'FileValidation' , 'prop' , ), 493, (493, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 980 , (3, 0, None, None) , 0 , )),
	(( 'FileValidation' , 'prop' , ), 493, (493, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 984 , (3, 0, None, None) , 0 , )),
]

win32com.client.CLSIDToClass.RegisterCLSID( "{00020970-0000-0000-C000-000000000046}", _Application )
