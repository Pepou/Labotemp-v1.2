# -*- coding: mbcs -*-
# Created by makepy.py version 0.5.01
# By python version 3.3.1 (v3.3.1:d9893d13c628, Apr  6 2013, 20:25:12) [MSC v.1600 32 bit (Intel)]
# From type library '{00020905-0000-0000-C000-000000000046}'
# On Mon Apr 15 10:41:20 2013
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
class _Document(DispatchBaseClass):
	CLSID = IID('{0002096B-0000-0000-C000-000000000046}')
	coclass_clsid = IID('{00020906-0000-0000-C000-000000000046}')

	def AcceptAllRevisions(self):
		return self._oleobj_.InvokeTypes(317, LCID, 1, (24, 0), (),)

	def AcceptAllRevisionsShown(self):
		return self._oleobj_.InvokeTypes(372, LCID, 1, (24, 0), (),)

	def Activate(self):
		return self._oleobj_.InvokeTypes(113, LCID, 1, (24, 0), (),)

	# The method ActiveWritingStyle is actually a property, but must be used as a method to correctly pass the arguments
	def ActiveWritingStyle(self, LanguageID=defaultNamedNotOptArg):
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(90, LCID, 2, (8, 0), ((16396, 1),),LanguageID
			)

	def AddDocumentWorkspaceHeader(self, RichFormat=defaultNamedNotOptArg, Url=defaultNamedNotOptArg, Title=defaultNamedNotOptArg, Description=defaultNamedNotOptArg
			, ID=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(482, LCID, 1, (24, 0), ((11, 1), (8, 1), (8, 1), (8, 1), (8, 1)),RichFormat
			, Url, Title, Description, ID)

	def AddMeetingWorkspaceHeader(self, SkipIfAbsent=defaultNamedNotOptArg, Url=defaultNamedNotOptArg, Title=defaultNamedNotOptArg, Description=defaultNamedNotOptArg
			, ID=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(515, LCID, 1, (24, 0), ((11, 1), (8, 1), (8, 1), (8, 1), (8, 1)),SkipIfAbsent
			, Url, Title, Description, ID)

	def AddToFavorites(self):
		return self._oleobj_.InvokeTypes(136, LCID, 1, (24, 0), (),)

	def ApplyDocumentTheme(self, FileName=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(546, LCID, 1, (24, 0), ((8, 1),),FileName
			)

	def ApplyQuickStyleSet(self, Name=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(525, LCID, 1, (24, 0), ((8, 1),),Name
			)

	def ApplyQuickStyleSet2(self, Style=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(566, LCID, 1, (24, 0), ((16396, 1),),Style
			)

	def ApplyTheme(self, Name=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(322, LCID, 1, (24, 0), ((8, 1),),Name
			)

	def AutoFormat(self):
		return self._oleobj_.InvokeTypes(148, LCID, 1, (24, 0), (),)

	# Result is of type Range
	def AutoSummarize(self, Length=defaultNamedOptArg, Mode=defaultNamedOptArg, UpdateProperties=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(138, LCID, 1, (9, 0), ((16396, 17), (16396, 17), (16396, 17)),Length
			, Mode, UpdateProperties)
		if ret is not None:
			ret = Dispatch(ret, 'AutoSummarize', '{0002095E-0000-0000-C000-000000000046}')
		return ret

	def CanCheckin(self):
		return self._oleobj_.InvokeTypes(351, LCID, 1, (11, 0), (),)

	def CheckConsistency(self):
		return self._oleobj_.InvokeTypes(259, LCID, 1, (24, 0), (),)

	def CheckGrammar(self):
		return self._oleobj_.InvokeTypes(131, LCID, 1, (24, 0), (),)

	def CheckIn(self, SaveChanges=True, Comments=defaultNamedNotOptArg, MakePublic=False):
		return self._oleobj_.InvokeTypes(349, LCID, 1, (24, 0), ((11, 49), (16396, 17), (11, 49)),SaveChanges
			, Comments, MakePublic)

	def CheckInWithVersion(self, SaveChanges=True, Comments=defaultNamedNotOptArg, MakePublic=False, VersionType=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(501, LCID, 1, (24, 0), ((11, 49), (16396, 17), (11, 49), (16396, 17)),SaveChanges
			, Comments, MakePublic, VersionType)

	def CheckNewSmartTags(self):
		return self._oleobj_.InvokeTypes(378, LCID, 1, (24, 0), (),)

	def CheckSpelling(self, CustomDictionary=defaultNamedOptArg, IgnoreUppercase=defaultNamedOptArg, AlwaysSuggest=defaultNamedOptArg, CustomDictionary2=defaultNamedOptArg
			, CustomDictionary3=defaultNamedOptArg, CustomDictionary4=defaultNamedOptArg, CustomDictionary5=defaultNamedOptArg, CustomDictionary6=defaultNamedOptArg, CustomDictionary7=defaultNamedOptArg
			, CustomDictionary8=defaultNamedOptArg, CustomDictionary9=defaultNamedOptArg, CustomDictionary10=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(132, LCID, 1, (24, 0), ((16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),CustomDictionary
			, IgnoreUppercase, AlwaysSuggest, CustomDictionary2, CustomDictionary3, CustomDictionary4
			, CustomDictionary5, CustomDictionary6, CustomDictionary7, CustomDictionary8, CustomDictionary9
			, CustomDictionary10)

	def Close(self, SaveChanges=defaultNamedOptArg, OriginalFormat=defaultNamedOptArg, RouteDocument=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1105, LCID, 1, (24, 0), ((16396, 17), (16396, 17), (16396, 17)),SaveChanges
			, OriginalFormat, RouteDocument)

	def ClosePrintPreview(self):
		return self._oleobj_.InvokeTypes(258, LCID, 1, (24, 0), (),)

	def Compare(self, Name=defaultNamedNotOptArg, AuthorName=defaultNamedOptArg, CompareTarget=defaultNamedOptArg, DetectFormatChanges=defaultNamedOptArg
			, IgnoreAllComparisonWarnings=defaultNamedOptArg, AddToRecentFiles=defaultNamedOptArg, RemovePersonalInformation=defaultNamedOptArg, RemoveDateAndTime=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(485, LCID, 1, (24, 0), ((8, 1), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),Name
			, AuthorName, CompareTarget, DetectFormatChanges, IgnoreAllComparisonWarnings, AddToRecentFiles
			, RemovePersonalInformation, RemoveDateAndTime)

	def Compare2000(self, Name=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(145, LCID, 1, (24, 0), ((8, 1),),Name
			)

	def Compare2002(self, Name=defaultNamedNotOptArg, AuthorName=defaultNamedOptArg, CompareTarget=defaultNamedOptArg, DetectFormatChanges=defaultNamedOptArg
			, IgnoreAllComparisonWarnings=defaultNamedOptArg, AddToRecentFiles=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(345, LCID, 1, (24, 0), ((8, 1), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),Name
			, AuthorName, CompareTarget, DetectFormatChanges, IgnoreAllComparisonWarnings, AddToRecentFiles
			)

	# The method Compatibility is actually a property, but must be used as a method to correctly pass the arguments
	def Compatibility(self, Type=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(55, LCID, 2, (11, 0), ((3, 1),),Type
			)

	def ComputeStatistics(self, Statistic=defaultNamedNotOptArg, IncludeFootnotesAndEndnotes=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(118, LCID, 1, (3, 0), ((3, 1), (16396, 17)),Statistic
			, IncludeFootnotesAndEndnotes)

	def Convert(self):
		return self._oleobj_.InvokeTypes(561, LCID, 1, (24, 0), (),)

	def ConvertAutoHyphens(self):
		return self._oleobj_.InvokeTypes(650, LCID, 1, (24, 0), (),)

	def ConvertNumbersToText(self, NumberType=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(141, LCID, 1, (24, 0), ((16396, 17),),NumberType
			)

	def ConvertVietDoc(self, CodePageOrigin=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(447, LCID, 1, (24, 0), ((3, 1),),CodePageOrigin
			)

	def CopyStylesFromTemplate(self, Template=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(126, LCID, 1, (24, 0), ((8, 1),),Template
			)

	def CountNumberedItems(self, NumberType=defaultNamedOptArg, Level=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(142, LCID, 1, (3, 0), ((16396, 17), (16396, 17)),NumberType
			, Level)

	# Result is of type LetterContent
	def CreateLetterContent(self, DateFormat=defaultNamedNotOptArg, IncludeHeaderFooter=defaultNamedNotOptArg, PageDesign=defaultNamedNotOptArg, LetterStyle=defaultNamedNotOptArg
			, Letterhead=defaultNamedNotOptArg, LetterheadLocation=defaultNamedNotOptArg, LetterheadSize=defaultNamedNotOptArg, RecipientName=defaultNamedNotOptArg, RecipientAddress=defaultNamedNotOptArg
			, Salutation=defaultNamedNotOptArg, SalutationType=defaultNamedNotOptArg, RecipientReference=defaultNamedNotOptArg, MailingInstructions=defaultNamedNotOptArg, AttentionLine=defaultNamedNotOptArg
			, Subject=defaultNamedNotOptArg, CCList=defaultNamedNotOptArg, ReturnAddress=defaultNamedNotOptArg, SenderName=defaultNamedNotOptArg, Closing=defaultNamedNotOptArg
			, SenderCompany=defaultNamedNotOptArg, SenderJobTitle=defaultNamedNotOptArg, SenderInitials=defaultNamedNotOptArg, EnclosureNumber=defaultNamedNotOptArg, InfoBlock=defaultNamedOptArg
			, RecipientCode=defaultNamedOptArg, RecipientGender=defaultNamedOptArg, ReturnAddressShortForm=defaultNamedOptArg, SenderCity=defaultNamedOptArg, SenderCode=defaultNamedOptArg
			, SenderGender=defaultNamedOptArg, SenderReference=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(260, LCID, 1, (13, 0), ((8, 1), (11, 1), (8, 1), (3, 1), (11, 1), (3, 1), (4, 1), (8, 1), (8, 1), (8, 1), (3, 1), (8, 1), (8, 1), (8, 1), (8, 1), (8, 1), (8, 1), (8, 1), (8, 1), (8, 1), (8, 1), (8, 1), (3, 1), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),DateFormat
			, IncludeHeaderFooter, PageDesign, LetterStyle, Letterhead, LetterheadLocation
			, LetterheadSize, RecipientName, RecipientAddress, Salutation, SalutationType
			, RecipientReference, MailingInstructions, AttentionLine, Subject, CCList
			, ReturnAddress, SenderName, Closing, SenderCompany, SenderJobTitle
			, SenderInitials, EnclosureNumber, InfoBlock, RecipientCode, RecipientGender
			, ReturnAddressShortForm, SenderCity, SenderCode, SenderGender, SenderReference
			)
		if ret is not None:
			# See if this IUnknown is really an IDispatch
			try:
				ret = ret.QueryInterface(pythoncom.IID_IDispatch)
			except pythoncom.error:
				return ret
			ret = Dispatch(ret, 'CreateLetterContent', '{000209F1-0000-0000-C000-000000000046}')
		return ret

	def DataForm(self):
		return self._oleobj_.InvokeTypes(106, LCID, 1, (24, 0), (),)

	def DeleteAllComments(self):
		return self._oleobj_.InvokeTypes(371, LCID, 1, (24, 0), (),)

	def DeleteAllCommentsShown(self):
		return self._oleobj_.InvokeTypes(374, LCID, 1, (24, 0), (),)

	def DeleteAllEditableRanges(self, EditorID=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(469, LCID, 1, (24, 0), ((16396, 17),),EditorID
			)

	def DeleteAllInkAnnotations(self):
		return self._oleobj_.InvokeTypes(479, LCID, 1, (24, 0), (),)

	def DetectLanguage(self):
		return self._oleobj_.InvokeTypes(151, LCID, 1, (24, 0), (),)

	def DowngradeDocument(self):
		return self._oleobj_.InvokeTypes(558, LCID, 1, (24, 0), (),)

	def Dummy2(self):
		return self._oleobj_.InvokeTypes(505, LCID, 1, (24, 0), (),)

	def Dummy4(self):
		return self._oleobj_.InvokeTypes(514, LCID, 1, (24, 0), (),)

	def EditionOptions(self, Type=defaultNamedNotOptArg, Option=defaultNamedNotOptArg, Name=defaultNamedNotOptArg, Format=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(122, LCID, 1, (24, 0), ((3, 1), (3, 1), (8, 1), (16396, 17)),Type
			, Option, Name, Format)

	def EndReview(self):
		return self._oleobj_.InvokeTypes(356, LCID, 1, (24, 0), (),)

	def ExportAsFixedFormat(self, OutputFileName=defaultNamedNotOptArg, ExportFormat=defaultNamedNotOptArg, OpenAfterExport=False, OptimizeFor=0
			, Range=0, From=1, To=1, Item=0, IncludeDocProps=False
			, KeepIRM=True, CreateBookmarks=0, DocStructureTags=True, BitmapMissingFonts=True, UseISO19005_1=False
			, FixedFormatExtClassPtr=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(552, LCID, 1, (24, 0), ((8, 1), (3, 1), (11, 49), (3, 49), (3, 49), (3, 49), (3, 49), (3, 49), (11, 49), (11, 49), (3, 49), (11, 49), (11, 49), (11, 49), (16396, 17)),OutputFileName
			, ExportFormat, OpenAfterExport, OptimizeFor, Range, From
			, To, Item, IncludeDocProps, KeepIRM, CreateBookmarks
			, DocStructureTags, BitmapMissingFonts, UseISO19005_1, FixedFormatExtClassPtr)

	def FitToPages(self):
		return self._oleobj_.InvokeTypes(104, LCID, 1, (24, 0), (),)

	def FollowHyperlink(self, Address=defaultNamedOptArg, SubAddress=defaultNamedOptArg, NewWindow=defaultNamedOptArg, AddHistory=defaultNamedOptArg
			, ExtraInfo=defaultNamedOptArg, Method=defaultNamedOptArg, HeaderInfo=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(135, LCID, 1, (24, 0), ((16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),Address
			, SubAddress, NewWindow, AddHistory, ExtraInfo, Method
			, HeaderInfo)

	def ForwardMailer(self):
		return self._oleobj_.InvokeTypes(250, LCID, 1, (24, 0), (),)

	def FreezeLayout(self):
		return self._oleobj_.InvokeTypes(553, LCID, 1, (24, 0), (),)

	def GetCrossReferenceItems(self, ReferenceType=defaultNamedNotOptArg):
		return self._ApplyTypes_(147, 1, (12, 0), ((16396, 1),), 'GetCrossReferenceItems', None,ReferenceType
			)

	# Result is of type LetterContent
	def GetLetterContent(self):
		ret = self._oleobj_.InvokeTypes(124, LCID, 1, (13, 0), (),)
		if ret is not None:
			# See if this IUnknown is really an IDispatch
			try:
				ret = ret.QueryInterface(pythoncom.IID_IDispatch)
			except pythoncom.error:
				return ret
			ret = Dispatch(ret, 'GetLetterContent', '{000209F1-0000-0000-C000-000000000046}')
		return ret

	# Result is of type WorkflowTasks
	def GetWorkflowTasks(self):
		ret = self._oleobj_.InvokeTypes(511, LCID, 1, (9, 0), (),)
		if ret is not None:
			ret = Dispatch(ret, 'GetWorkflowTasks', '{000CD901-0000-0000-C000-000000000046}')
		return ret

	# Result is of type WorkflowTemplates
	def GetWorkflowTemplates(self):
		ret = self._oleobj_.InvokeTypes(512, LCID, 1, (9, 0), (),)
		if ret is not None:
			ret = Dispatch(ret, 'GetWorkflowTemplates', '{000CD903-0000-0000-C000-000000000046}')
		return ret

	# Result is of type Range
	def GoTo(self, What=defaultNamedOptArg, Which=defaultNamedOptArg, Count=defaultNamedOptArg, Name=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(115, LCID, 1, (9, 0), ((16396, 17), (16396, 17), (16396, 17), (16396, 17)),What
			, Which, Count, Name)
		if ret is not None:
			ret = Dispatch(ret, 'GoTo', '{0002095E-0000-0000-C000-000000000046}')
		return ret

	def LockServerFile(self):
		return self._oleobj_.InvokeTypes(509, LCID, 1, (24, 0), (),)

	def MakeCompatibilityDefault(self):
		return self._oleobj_.InvokeTypes(119, LCID, 1, (24, 0), (),)

	def ManualHyphenation(self):
		return self._oleobj_.InvokeTypes(105, LCID, 1, (24, 0), (),)

	def Merge(self, FileName=defaultNamedNotOptArg, MergeTarget=defaultNamedOptArg, DetectFormatChanges=defaultNamedOptArg, UseFormattingFrom=defaultNamedOptArg
			, AddToRecentFiles=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(362, LCID, 1, (24, 0), ((8, 1), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),FileName
			, MergeTarget, DetectFormatChanges, UseFormattingFrom, AddToRecentFiles)

	def Merge2000(self, FileName=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(257, LCID, 1, (24, 0), ((8, 1),),FileName
			)

	def Post(self):
		return self._oleobj_.InvokeTypes(143, LCID, 1, (24, 0), (),)

	def PresentIt(self):
		return self._oleobj_.InvokeTypes(255, LCID, 1, (24, 0), (),)

	def PrintOut(self, Background=defaultNamedOptArg, Append=defaultNamedOptArg, Range=defaultNamedOptArg, OutputFileName=defaultNamedOptArg
			, From=defaultNamedOptArg, To=defaultNamedOptArg, Item=defaultNamedOptArg, Copies=defaultNamedOptArg, Pages=defaultNamedOptArg
			, PageType=defaultNamedOptArg, PrintToFile=defaultNamedOptArg, Collate=defaultNamedOptArg, ActivePrinterMacGX=defaultNamedOptArg, ManualDuplexPrint=defaultNamedOptArg
			, PrintZoomColumn=defaultNamedOptArg, PrintZoomRow=defaultNamedOptArg, PrintZoomPaperWidth=defaultNamedOptArg, PrintZoomPaperHeight=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(446, LCID, 1, (24, 0), ((16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),Background
			, Append, Range, OutputFileName, From, To
			, Item, Copies, Pages, PageType, PrintToFile
			, Collate, ActivePrinterMacGX, ManualDuplexPrint, PrintZoomColumn, PrintZoomRow
			, PrintZoomPaperWidth, PrintZoomPaperHeight)

	def PrintOut2000(self, Background=defaultNamedOptArg, Append=defaultNamedOptArg, Range=defaultNamedOptArg, OutputFileName=defaultNamedOptArg
			, From=defaultNamedOptArg, To=defaultNamedOptArg, Item=defaultNamedOptArg, Copies=defaultNamedOptArg, Pages=defaultNamedOptArg
			, PageType=defaultNamedOptArg, PrintToFile=defaultNamedOptArg, Collate=defaultNamedOptArg, ActivePrinterMacGX=defaultNamedOptArg, ManualDuplexPrint=defaultNamedOptArg
			, PrintZoomColumn=defaultNamedOptArg, PrintZoomRow=defaultNamedOptArg, PrintZoomPaperWidth=defaultNamedOptArg, PrintZoomPaperHeight=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(444, LCID, 1, (24, 0), ((16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),Background
			, Append, Range, OutputFileName, From, To
			, Item, Copies, Pages, PageType, PrintToFile
			, Collate, ActivePrinterMacGX, ManualDuplexPrint, PrintZoomColumn, PrintZoomRow
			, PrintZoomPaperWidth, PrintZoomPaperHeight)

	def PrintOutOld(self, Background=defaultNamedOptArg, Append=defaultNamedOptArg, Range=defaultNamedOptArg, OutputFileName=defaultNamedOptArg
			, From=defaultNamedOptArg, To=defaultNamedOptArg, Item=defaultNamedOptArg, Copies=defaultNamedOptArg, Pages=defaultNamedOptArg
			, PageType=defaultNamedOptArg, PrintToFile=defaultNamedOptArg, Collate=defaultNamedOptArg, ActivePrinterMacGX=defaultNamedOptArg, ManualDuplexPrint=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(109, LCID, 1, (24, 0), ((16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),Background
			, Append, Range, OutputFileName, From, To
			, Item, Copies, Pages, PageType, PrintToFile
			, Collate, ActivePrinterMacGX, ManualDuplexPrint)

	def PrintPreview(self):
		return self._oleobj_.InvokeTypes(114, LCID, 1, (24, 0), (),)

	def Protect(self, Type=defaultNamedNotOptArg, NoReset=defaultNamedOptArg, Password=defaultNamedOptArg, UseIRM=defaultNamedOptArg
			, EnforceStyleLock=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(467, LCID, 1, (24, 0), ((3, 1), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),Type
			, NoReset, Password, UseIRM, EnforceStyleLock)

	def Protect2002(self, Type=defaultNamedNotOptArg, NoReset=defaultNamedOptArg, Password=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(120, LCID, 1, (24, 0), ((3, 1), (16396, 17), (16396, 17)),Type
			, NoReset, Password)

	# Result is of type Range
	def Range(self, Start=defaultNamedOptArg, End=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(2000, LCID, 1, (9, 0), ((16396, 17), (16396, 17)),Start
			, End)
		if ret is not None:
			ret = Dispatch(ret, 'Range', '{0002095E-0000-0000-C000-000000000046}')
		return ret

	def RecheckSmartTags(self):
		return self._oleobj_.InvokeTypes(363, LCID, 1, (24, 0), (),)

	def Redo(self, Times=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(117, LCID, 1, (11, 0), ((16396, 17),),Times
			)

	def RejectAllRevisions(self):
		return self._oleobj_.InvokeTypes(318, LCID, 1, (24, 0), (),)

	def RejectAllRevisionsShown(self):
		return self._oleobj_.InvokeTypes(373, LCID, 1, (24, 0), (),)

	def Reload(self):
		return self._oleobj_.InvokeTypes(137, LCID, 1, (24, 0), (),)

	def ReloadAs(self, Encoding=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(331, LCID, 1, (24, 0), ((3, 1),),Encoding
			)

	def RemoveDocumentInformation(self, RemoveDocInfoType=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(495, LCID, 1, (24, 0), ((3, 1),),RemoveDocInfoType
			)

	def RemoveDocumentWorkspaceHeader(self, ID=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(483, LCID, 1, (24, 0), ((8, 1),),ID
			)

	def RemoveLockedStyles(self):
		return self._oleobj_.InvokeTypes(487, LCID, 1, (24, 0), (),)

	def RemoveNumbers(self, NumberType=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(140, LCID, 1, (24, 0), ((16396, 17),),NumberType
			)

	def RemoveSmartTags(self):
		return self._oleobj_.InvokeTypes(364, LCID, 1, (24, 0), (),)

	def RemoveTheme(self):
		return self._oleobj_.InvokeTypes(323, LCID, 1, (24, 0), (),)

	def Repaginate(self):
		return self._oleobj_.InvokeTypes(103, LCID, 1, (24, 0), (),)

	def Reply(self):
		return self._oleobj_.InvokeTypes(251, LCID, 1, (24, 0), (),)

	def ReplyAll(self):
		return self._oleobj_.InvokeTypes(252, LCID, 1, (24, 0), (),)

	def ReplyWithChanges(self, ShowMessage=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(354, LCID, 1, (24, 0), ((16396, 17),),ShowMessage
			)

	def ResetFormFields(self):
		return self._oleobj_.InvokeTypes(375, LCID, 1, (24, 0), (),)

	def Route(self):
		return self._oleobj_.InvokeTypes(107, LCID, 1, (24, 0), (),)

	def RunAutoMacro(self, Which=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(112, LCID, 1, (24, 0), ((3, 1),),Which
			)

	def RunLetterWizard(self, LetterContent=defaultNamedOptArg, WizardMode=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(123, LCID, 1, (24, 0), ((16396, 17), (16396, 17)),LetterContent
			, WizardMode)

	def Save(self):
		return self._oleobj_.InvokeTypes(108, LCID, 1, (24, 0), (),)

	def SaveAs(self, FileName=defaultNamedOptArg, FileFormat=defaultNamedOptArg, LockComments=defaultNamedOptArg, Password=defaultNamedOptArg
			, AddToRecentFiles=defaultNamedOptArg, WritePassword=defaultNamedOptArg, ReadOnlyRecommended=defaultNamedOptArg, EmbedTrueTypeFonts=defaultNamedOptArg, SaveNativePictureFormat=defaultNamedOptArg
			, SaveFormsData=defaultNamedOptArg, SaveAsAOCELetter=defaultNamedOptArg, Encoding=defaultNamedOptArg, InsertLineBreaks=defaultNamedOptArg, AllowSubstitutions=defaultNamedOptArg
			, LineEnding=defaultNamedOptArg, AddBiDiMarks=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(376, LCID, 1, (24, 0), ((16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),FileName
			, FileFormat, LockComments, Password, AddToRecentFiles, WritePassword
			, ReadOnlyRecommended, EmbedTrueTypeFonts, SaveNativePictureFormat, SaveFormsData, SaveAsAOCELetter
			, Encoding, InsertLineBreaks, AllowSubstitutions, LineEnding, AddBiDiMarks
			)

	def SaveAs2(self, FileName=defaultNamedOptArg, FileFormat=defaultNamedOptArg, LockComments=defaultNamedOptArg, Password=defaultNamedOptArg
			, AddToRecentFiles=defaultNamedOptArg, WritePassword=defaultNamedOptArg, ReadOnlyRecommended=defaultNamedOptArg, EmbedTrueTypeFonts=defaultNamedOptArg, SaveNativePictureFormat=defaultNamedOptArg
			, SaveFormsData=defaultNamedOptArg, SaveAsAOCELetter=defaultNamedOptArg, Encoding=defaultNamedOptArg, InsertLineBreaks=defaultNamedOptArg, AllowSubstitutions=defaultNamedOptArg
			, LineEnding=defaultNamedOptArg, AddBiDiMarks=defaultNamedOptArg, CompatibilityMode=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(568, LCID, 1, (24, 0), ((16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),FileName
			, FileFormat, LockComments, Password, AddToRecentFiles, WritePassword
			, ReadOnlyRecommended, EmbedTrueTypeFonts, SaveNativePictureFormat, SaveFormsData, SaveAsAOCELetter
			, Encoding, InsertLineBreaks, AllowSubstitutions, LineEnding, AddBiDiMarks
			, CompatibilityMode)

	def SaveAs2000(self, FileName=defaultNamedOptArg, FileFormat=defaultNamedOptArg, LockComments=defaultNamedOptArg, Password=defaultNamedOptArg
			, AddToRecentFiles=defaultNamedOptArg, WritePassword=defaultNamedOptArg, ReadOnlyRecommended=defaultNamedOptArg, EmbedTrueTypeFonts=defaultNamedOptArg, SaveNativePictureFormat=defaultNamedOptArg
			, SaveFormsData=defaultNamedOptArg, SaveAsAOCELetter=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(102, LCID, 1, (24, 0), ((16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),FileName
			, FileFormat, LockComments, Password, AddToRecentFiles, WritePassword
			, ReadOnlyRecommended, EmbedTrueTypeFonts, SaveNativePictureFormat, SaveFormsData, SaveAsAOCELetter
			)

	def SaveAsQuickStyleSet(self, FileName=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(524, LCID, 1, (24, 0), ((8, 1),),FileName
			)

	def Select(self):
		return self._oleobj_.InvokeTypes(65535, LCID, 1, (24, 0), (),)

	def SelectAllEditableRanges(self, EditorID=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(468, LCID, 1, (24, 0), ((16396, 17),),EditorID
			)

	# Result is of type ContentControls
	def SelectContentControlsByTag(self, Tag=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(562, LCID, 1, (9, 0), ((8, 1),),Tag
			)
		if ret is not None:
			ret = Dispatch(ret, 'SelectContentControlsByTag', '{804CD967-F83B-432D-9446-C61A45CFEFF0}')
		return ret

	# Result is of type ContentControls
	def SelectContentControlsByTitle(self, Title=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(551, LCID, 1, (9, 0), ((8, 1),),Title
			)
		if ret is not None:
			ret = Dispatch(ret, 'SelectContentControlsByTitle', '{804CD967-F83B-432D-9446-C61A45CFEFF0}')
		return ret

	# Result is of type ContentControls
	def SelectLinkedControls(self, Node=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(549, LCID, 1, (9, 0), ((9, 1),),Node
			)
		if ret is not None:
			ret = Dispatch(ret, 'SelectLinkedControls', '{804CD967-F83B-432D-9446-C61A45CFEFF0}')
		return ret

	# Result is of type XMLNodes
	def SelectNodes(self, XPath=defaultNamedNotOptArg, PrefixMapping='', FastSearchSkippingTextNodes=True):
		return self._ApplyTypes_(489, 1, (9, 32), ((8, 1), (8, 49), (11, 49)), 'SelectNodes', '{D36C1F42-7044-4B9E-9CA3-85919454DB04}',XPath
			, PrefixMapping, FastSearchSkippingTextNodes)

	# Result is of type XMLNode
	def SelectSingleNode(self, XPath=defaultNamedNotOptArg, PrefixMapping='', FastSearchSkippingTextNodes=True):
		return self._ApplyTypes_(488, 1, (9, 32), ((8, 1), (8, 49), (11, 49)), 'SelectSingleNode', '{09760240-0B89-49F7-A79D-479F24723F56}',XPath
			, PrefixMapping, FastSearchSkippingTextNodes)

	# Result is of type ContentControls
	def SelectUnlinkedControls(self, Stream=0):
		ret = self._oleobj_.InvokeTypes(550, LCID, 1, (9, 0), ((13, 49),),Stream
			)
		if ret is not None:
			ret = Dispatch(ret, 'SelectUnlinkedControls', '{804CD967-F83B-432D-9446-C61A45CFEFF0}')
		return ret

	def SendFax(self, Address=defaultNamedNotOptArg, Subject=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(256, LCID, 1, (24, 0), ((8, 1), (16396, 17)),Address
			, Subject)

	def SendFaxOverInternet(self, Recipients=defaultNamedOptArg, Subject=defaultNamedOptArg, ShowMessage=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(464, LCID, 1, (24, 0), ((16396, 17), (16396, 17), (16396, 17)),Recipients
			, Subject, ShowMessage)

	def SendForReview(self, Recipients=defaultNamedOptArg, Subject=defaultNamedOptArg, ShowMessage=defaultNamedOptArg, IncludeAttachment=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(353, LCID, 1, (24, 0), ((16396, 17), (16396, 17), (16396, 17), (16396, 17)),Recipients
			, Subject, ShowMessage, IncludeAttachment)

	def SendMail(self):
		return self._oleobj_.InvokeTypes(110, LCID, 1, (24, 0), (),)

	def SendMailer(self, FileFormat=defaultNamedOptArg, Priority=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(253, LCID, 1, (24, 0), ((16396, 17), (16396, 17)),FileFormat
			, Priority)

	# The method SetActiveWritingStyle is actually a property, but must be used as a method to correctly pass the arguments
	def SetActiveWritingStyle(self, LanguageID=defaultNamedNotOptArg, arg1=defaultUnnamedArg):
		return self._oleobj_.InvokeTypes(90, LCID, 4, (24, 0), ((16396, 1), (8, 1)),LanguageID
			, arg1)

	# The method SetCompatibility is actually a property, but must be used as a method to correctly pass the arguments
	def SetCompatibility(self, Type=defaultNamedNotOptArg, arg1=defaultUnnamedArg):
		return self._oleobj_.InvokeTypes(55, LCID, 4, (24, 0), ((3, 1), (11, 1)),Type
			, arg1)

	def SetCompatibilityMode(self, Mode=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(571, LCID, 1, (24, 0), ((3, 1),),Mode
			)

	def SetDefaultTableStyle(self, Style=defaultNamedNotOptArg, SetInTemplate=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(366, LCID, 1, (24, 0), ((16396, 1), (11, 1)),Style
			, SetInTemplate)

	def SetLetterContent(self, LetterContent=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(125, LCID, 1, (24, 0), ((16396, 1),),LetterContent
			)

	def SetPasswordEncryptionOptions(self, PasswordEncryptionProvider=defaultNamedNotOptArg, PasswordEncryptionAlgorithm=defaultNamedNotOptArg, PasswordEncryptionKeyLength=defaultNamedNotOptArg, PasswordEncryptionFileProperties=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(361, LCID, 1, (24, 0), ((8, 1), (8, 1), (3, 1), (16396, 17)),PasswordEncryptionProvider
			, PasswordEncryptionAlgorithm, PasswordEncryptionKeyLength, PasswordEncryptionFileProperties)

	def ToggleFormsDesign(self):
		return self._oleobj_.InvokeTypes(144, LCID, 1, (24, 0), (),)

	def TransformDocument(self, Path=defaultNamedNotOptArg, DataOnly=True):
		return self._oleobj_.InvokeTypes(500, LCID, 1, (24, 0), ((8, 1), (11, 49)),Path
			, DataOnly)

	def Undo(self, Times=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(116, LCID, 1, (11, 0), ((16396, 17),),Times
			)

	def UndoClear(self):
		return self._oleobj_.InvokeTypes(254, LCID, 1, (24, 0), (),)

	def UnfreezeLayout(self):
		return self._oleobj_.InvokeTypes(554, LCID, 1, (24, 0), (),)

	def Unprotect(self, Password=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(121, LCID, 1, (24, 0), ((16396, 17),),Password
			)

	def UpdateStyles(self):
		return self._oleobj_.InvokeTypes(127, LCID, 1, (24, 0), (),)

	def UpdateSummaryProperties(self):
		return self._oleobj_.InvokeTypes(146, LCID, 1, (24, 0), (),)

	def ViewCode(self):
		return self._oleobj_.InvokeTypes(149, LCID, 1, (24, 0), (),)

	def ViewPropertyBrowser(self):
		return self._oleobj_.InvokeTypes(150, LCID, 1, (24, 0), (),)

	def WebPagePreview(self):
		return self._oleobj_.InvokeTypes(325, LCID, 1, (24, 0), (),)

	def sblt(self, s=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(445, LCID, 1, (24, 0), ((8, 1),),s
			)

	_prop_map_get_ = {
		"ActiveTheme": (540, 2, (8, 0), (), "ActiveTheme", None),
		"ActiveThemeDisplayName": (541, 2, (8, 0), (), "ActiveThemeDisplayName", None),
		# Method 'ActiveWindow' returns object of type 'Window'
		"ActiveWindow": (42, 2, (9, 0), (), "ActiveWindow", '{00020962-0000-0000-C000-000000000046}'),
		# Method 'Application' returns object of type 'Application'
		"Application": (1, 2, (13, 0), (), "Application", '{000209FF-0000-0000-C000-000000000046}'),
		"AttachedTemplate": (67, 2, (12, 0), (), "AttachedTemplate", None),
		"AutoFormatOverride": (472, 2, (11, 0), (), "AutoFormatOverride", None),
		"AutoHyphenation": (11, 2, (11, 0), (), "AutoHyphenation", None),
		# Method 'Background' returns object of type 'Shape'
		"Background": (69, 2, (9, 0), (), "Background", '{000209A0-0000-0000-C000-000000000046}'),
		# Method 'Bibliography' returns object of type 'Bibliography'
		"Bibliography": (516, 2, (9, 0), (), "Bibliography", '{3834F60F-EE8C-455D-A441-D766675D6D3B}'),
		# Method 'Bookmarks' returns object of type 'Bookmarks'
		"Bookmarks": (4, 2, (9, 0), (), "Bookmarks", '{00020967-0000-0000-C000-000000000046}'),
		"BuiltInDocumentProperties": (1000, 2, (9, 0), (), "BuiltInDocumentProperties", None),
		# Method 'Characters' returns object of type 'Characters'
		"Characters": (19, 2, (9, 0), (), "Characters", '{0002095D-0000-0000-C000-000000000046}'),
		# Method 'ChildNodeSuggestions' returns object of type 'XMLChildNodeSuggestions'
		"ChildNodeSuggestions": (486, 2, (9, 0), (), "ChildNodeSuggestions", '{DE63B5AC-CA4F-46FE-9184-A5719AB9ED5B}'),
		"ClickAndTypeParagraphStyle": (328, 2, (12, 0), (), "ClickAndTypeParagraphStyle", None),
		# Method 'CoAuthoring' returns object of type 'CoAuthoring'
		"CoAuthoring": (600, 2, (9, 0), (), "CoAuthoring", '{65DF9F31-B1E3-4651-87E8-51D55F302161}'),
		"CodeName": (262, 2, (8, 0), (), "CodeName", None),
		# Method 'CommandBars' returns object of type 'CommandBars'
		"CommandBars": (57, 2, (13, 0), (), "CommandBars", '{55F88893-7708-11D1-ACEB-006008961DA5}'),
		# Method 'Comments' returns object of type 'Comments'
		"Comments": (9, 2, (9, 0), (), "Comments", '{00020940-0000-0000-C000-000000000046}'),
		"CompatibilityMode": (567, 2, (3, 0), (), "CompatibilityMode", None),
		"ConsecutiveHyphensLimit": (14, 2, (3, 0), (), "ConsecutiveHyphensLimit", None),
		"Container": (82, 2, (9, 0), (), "Container", None),
		# Method 'Content' returns object of type 'Range'
		"Content": (41, 2, (9, 0), (), "Content", '{0002095E-0000-0000-C000-000000000046}'),
		# Method 'ContentControls' returns object of type 'ContentControls'
		"ContentControls": (508, 2, (9, 0), (), "ContentControls", '{804CD967-F83B-432D-9446-C61A45CFEFF0}'),
		# Method 'ContentTypeProperties' returns object of type 'MetaProperties'
		"ContentTypeProperties": (496, 2, (9, 0), (), "ContentTypeProperties", '{000C038E-0000-0000-C000-000000000046}'),
		"Creator": (1001, 2, (3, 0), (), "Creator", None),
		"CurrentRsid": (563, 2, (3, 0), (), "CurrentRsid", None),
		"CustomDocumentProperties": (2, 2, (9, 0), (), "CustomDocumentProperties", None),
		# Method 'CustomXMLParts' returns object of type 'CustomXMLParts'
		"CustomXMLParts": (521, 2, (13, 0), (), "CustomXMLParts", '{000CDB0C-0000-0000-C000-000000000046}'),
		"DefaultTabStop": (48, 2, (4, 0), (), "DefaultTabStop", None),
		"DefaultTableStyle": (365, 2, (12, 0), (), "DefaultTableStyle", None),
		"DefaultTargetFrame": (340, 2, (8, 0), (), "DefaultTargetFrame", None),
		"DisableFeatures": (337, 2, (11, 0), (), "DisableFeatures", None),
		"DisableFeaturesIntroducedAfter": (343, 2, (3, 0), (), "DisableFeaturesIntroducedAfter", None),
		"DoNotEmbedSystemFonts": (338, 2, (11, 0), (), "DoNotEmbedSystemFonts", None),
		"DocID": (564, 2, (3, 0), (), "DocID", None),
		# Method 'DocumentInspectors' returns object of type 'DocumentInspectors'
		"DocumentInspectors": (510, 2, (9, 0), (), "DocumentInspectors", '{000C0392-0000-0000-C000-000000000046}'),
		# Method 'DocumentLibraryVersions' returns object of type 'DocumentLibraryVersions'
		"DocumentLibraryVersions": (476, 2, (9, 0), (), "DocumentLibraryVersions", '{000C0388-0000-0000-C000-000000000046}'),
		# Method 'DocumentTheme' returns object of type 'OfficeTheme'
		"DocumentTheme": (545, 2, (9, 0), (), "DocumentTheme", '{000C03A0-0000-0000-C000-000000000046}'),
		"Dummy1": (503, 2, (24, 0), (), "Dummy1", None),
		"Dummy3": (506, 2, (24, 0), (), "Dummy3", None),
		# Method 'Email' returns object of type 'Email'
		"Email": (319, 2, (9, 0), (), "Email", '{000209DD-0000-0000-C000-000000000046}'),
		"EmbedLinguisticData": (377, 2, (11, 0), (), "EmbedLinguisticData", None),
		"EmbedSmartTags": (347, 2, (11, 0), (), "EmbedSmartTags", None),
		"EmbedTrueTypeFonts": (50, 2, (11, 0), (), "EmbedTrueTypeFonts", None),
		"EncryptionProvider": (559, 2, (8, 0), (), "EncryptionProvider", None),
		# Method 'Endnotes' returns object of type 'Endnotes'
		"Endnotes": (8, 2, (9, 0), (), "Endnotes", '{00020941-0000-0000-C000-000000000046}'),
		"EnforceStyle": (471, 2, (11, 0), (), "EnforceStyle", None),
		# Method 'Envelope' returns object of type 'Envelope'
		"Envelope": (28, 2, (9, 0), (), "Envelope", '{00020918-0000-0000-C000-000000000046}'),
		"FarEastLineBreakLanguage": (326, 2, (3, 0), (), "FarEastLineBreakLanguage", None),
		"FarEastLineBreakLevel": (311, 2, (3, 0), (), "FarEastLineBreakLevel", None),
		# Method 'Fields' returns object of type 'Fields'
		"Fields": (20, 2, (9, 0), (), "Fields", '{00020930-0000-0000-C000-000000000046}'),
		"Final": (527, 2, (11, 0), (), "Final", None),
		# Method 'Footnotes' returns object of type 'Footnotes'
		"Footnotes": (7, 2, (9, 0), (), "Footnotes", '{00020942-0000-0000-C000-000000000046}'),
		# Method 'FormFields' returns object of type 'FormFields'
		"FormFields": (21, 2, (9, 0), (), "FormFields", '{00020929-0000-0000-C000-000000000046}'),
		"FormattingShowClear": (449, 2, (11, 0), (), "FormattingShowClear", None),
		"FormattingShowFilter": (452, 2, (3, 0), (), "FormattingShowFilter", None),
		"FormattingShowFont": (448, 2, (11, 0), (), "FormattingShowFont", None),
		"FormattingShowNextLevel": (522, 2, (11, 0), (), "FormattingShowNextLevel", None),
		"FormattingShowNumbering": (451, 2, (11, 0), (), "FormattingShowNumbering", None),
		"FormattingShowParagraph": (450, 2, (11, 0), (), "FormattingShowParagraph", None),
		"FormattingShowUserStyleName": (523, 2, (11, 0), (), "FormattingShowUserStyleName", None),
		"FormsDesign": (100, 2, (11, 0), (), "FormsDesign", None),
		# Method 'Frames' returns object of type 'Frames'
		"Frames": (23, 2, (9, 0), (), "Frames", '{0002092B-0000-0000-C000-000000000046}'),
		# Method 'Frameset' returns object of type 'Frameset'
		"Frameset": (327, 2, (9, 0), (), "Frameset", '{000209E2-0000-0000-C000-000000000046}'),
		"FullName": (29, 2, (8, 0), (), "FullName", None),
		"GrammarChecked": (70, 2, (11, 0), (), "GrammarChecked", None),
		# Method 'GrammaticalErrors' returns object of type 'ProofreadingErrors'
		"GrammaticalErrors": (97, 2, (9, 0), (), "GrammaticalErrors", '{000209BB-0000-0000-C000-000000000046}'),
		"GridDistanceHorizontal": (302, 2, (4, 0), (), "GridDistanceHorizontal", None),
		"GridDistanceVertical": (303, 2, (4, 0), (), "GridDistanceVertical", None),
		"GridOriginFromMargin": (308, 2, (11, 0), (), "GridOriginFromMargin", None),
		"GridOriginHorizontal": (304, 2, (4, 0), (), "GridOriginHorizontal", None),
		"GridOriginVertical": (305, 2, (4, 0), (), "GridOriginVertical", None),
		"GridSpaceBetweenHorizontalLines": (306, 2, (3, 0), (), "GridSpaceBetweenHorizontalLines", None),
		"GridSpaceBetweenVerticalLines": (307, 2, (3, 0), (), "GridSpaceBetweenVerticalLines", None),
		# Method 'HTMLDivisions' returns object of type 'HTMLDivisions'
		"HTMLDivisions": (342, 2, (9, 0), (), "HTMLDivisions", '{000209E8-0000-0000-C000-000000000046}'),
		# Method 'HTMLProject' returns object of type 'HTMLProject'
		"HTMLProject": (329, 2, (9, 0), (), "HTMLProject", '{000C0356-0000-0000-C000-000000000046}'),
		"HasMailer": (93, 2, (11, 0), (), "HasMailer", None),
		"HasPassword": (87, 2, (11, 0), (), "HasPassword", None),
		"HasRoutingSlip": (35, 2, (11, 0), (), "HasRoutingSlip", None),
		"HasVBProject": (548, 2, (11, 0), (), "HasVBProject", None),
		# Method 'Hyperlinks' returns object of type 'Hyperlinks'
		"Hyperlinks": (61, 2, (9, 0), (), "Hyperlinks", '{0002099C-0000-0000-C000-000000000046}'),
		"HyphenateCaps": (12, 2, (11, 0), (), "HyphenateCaps", None),
		"HyphenationZone": (13, 2, (3, 0), (), "HyphenationZone", None),
		# Method 'Indexes' returns object of type 'Indexes'
		"Indexes": (39, 2, (9, 0), (), "Indexes", '{0002097C-0000-0000-C000-000000000046}'),
		# Method 'InlineShapes' returns object of type 'InlineShapes'
		"InlineShapes": (68, 2, (9, 0), (), "InlineShapes", '{000209A9-0000-0000-C000-000000000046}'),
		"IsMasterDocument": (46, 2, (11, 0), (), "IsMasterDocument", None),
		"IsSubdocument": (58, 2, (11, 0), (), "IsSubdocument", None),
		"JustificationMode": (310, 2, (3, 0), (), "JustificationMode", None),
		"KerningByAlgorithm": (309, 2, (11, 0), (), "KerningByAlgorithm", None),
		"Kind": (43, 2, (3, 0), (), "Kind", None),
		"LanguageDetected": (321, 2, (11, 0), (), "LanguageDetected", None),
		# Method 'ListParagraphs' returns object of type 'ListParagraphs'
		"ListParagraphs": (84, 2, (9, 0), (), "ListParagraphs", '{00020991-0000-0000-C000-000000000046}'),
		# Method 'ListTemplates' returns object of type 'ListTemplates'
		"ListTemplates": (63, 2, (9, 0), (), "ListTemplates", '{00020990-0000-0000-C000-000000000046}'),
		# Method 'Lists' returns object of type 'Lists'
		"Lists": (64, 2, (9, 0), (), "Lists", '{00020993-0000-0000-C000-000000000046}'),
		"LockQuickStyleSet": (518, 2, (11, 0), (), "LockQuickStyleSet", None),
		"LockTheme": (517, 2, (11, 0), (), "LockTheme", None),
		# Method 'MailEnvelope' returns object of type 'MsoEnvelope'
		"MailEnvelope": (336, 2, (13, 0), (), "MailEnvelope", '{0006F01A-0000-0000-C000-000000000046}'),
		# Method 'MailMerge' returns object of type 'MailMerge'
		"MailMerge": (27, 2, (9, 0), (), "MailMerge", '{00020920-0000-0000-C000-000000000046}'),
		# Method 'Mailer' returns object of type 'Mailer'
		"Mailer": (94, 2, (9, 0), (), "Mailer", '{000209BD-0000-0000-C000-000000000046}'),
		"Name": (0, 2, (8, 0), (), "Name", None),
		"NoLineBreakAfter": (313, 2, (8, 0), (), "NoLineBreakAfter", None),
		"NoLineBreakBefore": (312, 2, (8, 0), (), "NoLineBreakBefore", None),
		"OMathBreakBin": (528, 2, (3, 0), (), "OMathBreakBin", None),
		"OMathBreakSub": (529, 2, (3, 0), (), "OMathBreakSub", None),
		"OMathFontName": (555, 2, (8, 0), (), "OMathFontName", None),
		"OMathIntSubSupLim": (536, 2, (11, 0), (), "OMathIntSubSupLim", None),
		"OMathJc": (530, 2, (3, 0), (), "OMathJc", None),
		"OMathLeftMargin": (531, 2, (4, 0), (), "OMathLeftMargin", None),
		"OMathNarySupSubLim": (537, 2, (11, 0), (), "OMathNarySupSubLim", None),
		"OMathRightMargin": (532, 2, (4, 0), (), "OMathRightMargin", None),
		"OMathSmallFrac": (539, 2, (11, 0), (), "OMathSmallFrac", None),
		"OMathWrap": (535, 2, (4, 0), (), "OMathWrap", None),
		# Method 'OMaths' returns object of type 'OMaths'
		"OMaths": (504, 2, (9, 0), (), "OMaths", '{873E774B-926A-4CB1-878D-635A45187595}'),
		"OpenEncoding": (332, 2, (3, 0), (), "OpenEncoding", None),
		"OptimizeForWord97": (334, 2, (11, 0), (), "OptimizeForWord97", None),
		"OriginalDocumentTitle": (519, 2, (8, 0), (), "OriginalDocumentTitle", None),
		# Method 'PageSetup' returns object of type 'PageSetup'
		"PageSetup": (1101, 2, (9, 0), (), "PageSetup", '{00020971-0000-0000-C000-000000000046}'),
		# Method 'Paragraphs' returns object of type 'Paragraphs'
		"Paragraphs": (16, 2, (9, 0), (), "Paragraphs", '{00020958-0000-0000-C000-000000000046}'),
		"Parent": (1002, 2, (9, 0), (), "Parent", None),
		"PasswordEncryptionAlgorithm": (368, 2, (8, 0), (), "PasswordEncryptionAlgorithm", None),
		"PasswordEncryptionFileProperties": (370, 2, (11, 0), (), "PasswordEncryptionFileProperties", None),
		"PasswordEncryptionKeyLength": (369, 2, (3, 0), (), "PasswordEncryptionKeyLength", None),
		"PasswordEncryptionProvider": (367, 2, (8, 0), (), "PasswordEncryptionProvider", None),
		"Path": (3, 2, (8, 0), (), "Path", None),
		# Method 'Permission' returns object of type 'Permission'
		"Permission": (453, 2, (9, 0), (), "Permission", '{000C0376-0000-0000-C000-000000000046}'),
		"PrintFormsData": (83, 2, (11, 0), (), "PrintFormsData", None),
		"PrintFractionalWidths": (79, 2, (11, 0), (), "PrintFractionalWidths", None),
		"PrintPostScriptOverText": (80, 2, (11, 0), (), "PrintPostScriptOverText", None),
		"PrintRevisions": (315, 2, (11, 0), (), "PrintRevisions", None),
		"ProtectionType": (60, 2, (3, 0), (), "ProtectionType", None),
		"ReadOnly": (44, 2, (11, 0), (), "ReadOnly", None),
		"ReadOnlyRecommended": (52, 2, (11, 0), (), "ReadOnlyRecommended", None),
		# Method 'ReadabilityStatistics' returns object of type 'ReadabilityStatistics'
		"ReadabilityStatistics": (96, 2, (9, 0), (), "ReadabilityStatistics", '{000209AE-0000-0000-C000-000000000046}'),
		"ReadingLayoutSizeX": (491, 2, (3, 0), (), "ReadingLayoutSizeX", None),
		"ReadingLayoutSizeY": (492, 2, (3, 0), (), "ReadingLayoutSizeY", None),
		"ReadingModeLayoutFrozen": (481, 2, (11, 0), (), "ReadingModeLayoutFrozen", None),
		"RemoveDateAndTime": (484, 2, (11, 0), (), "RemoveDateAndTime", None),
		"RemovePersonalInformation": (344, 2, (11, 0), (), "RemovePersonalInformation", None),
		# Method 'Research' returns object of type 'Research'
		"Research": (526, 2, (9, 0), (), "Research", '{E6AAEC05-E543-4085-BA92-9BF7D2474F51}'),
		"RevisedDocumentTitle": (520, 2, (8, 0), (), "RevisedDocumentTitle", None),
		# Method 'Revisions' returns object of type 'Revisions'
		"Revisions": (30, 2, (9, 0), (), "Revisions", '{00020980-0000-0000-C000-000000000046}'),
		"Routed": (37, 2, (11, 0), (), "Routed", None),
		# Method 'RoutingSlip' returns object of type 'RoutingSlip'
		"RoutingSlip": (36, 2, (9, 0), (), "RoutingSlip", '{00020969-0000-0000-C000-000000000046}'),
		"SaveEncoding": (333, 2, (3, 0), (), "SaveEncoding", None),
		"SaveFormat": (59, 2, (3, 0), (), "SaveFormat", None),
		"SaveFormsData": (51, 2, (11, 0), (), "SaveFormsData", None),
		"SaveSubsetFonts": (53, 2, (11, 0), (), "SaveSubsetFonts", None),
		"Saved": (40, 2, (11, 0), (), "Saved", None),
		# Method 'Scripts' returns object of type 'Scripts'
		"Scripts": (320, 2, (9, 0), (), "Scripts", '{000C0340-0000-0000-C000-000000000046}'),
		# Method 'Sections' returns object of type 'Sections'
		"Sections": (15, 2, (9, 0), (), "Sections", '{0002095A-0000-0000-C000-000000000046}'),
		# Method 'Sentences' returns object of type 'Sentences'
		"Sentences": (18, 2, (9, 0), (), "Sentences", '{0002095B-0000-0000-C000-000000000046}'),
		# Method 'ServerPolicy' returns object of type 'ServerPolicy'
		"ServerPolicy": (507, 2, (9, 0), (), "ServerPolicy", '{000C0390-0000-0000-C000-000000000046}'),
		# Method 'Shapes' returns object of type 'Shapes'
		"Shapes": (62, 2, (9, 0), (), "Shapes", '{0002099F-0000-0000-C000-000000000046}'),
		# Method 'SharedWorkspace' returns object of type 'SharedWorkspace'
		"SharedWorkspace": (463, 2, (9, 0), (), "SharedWorkspace", '{000C0385-0000-0000-C000-000000000046}'),
		"ShowGrammaticalErrors": (72, 2, (11, 0), (), "ShowGrammaticalErrors", None),
		"ShowRevisions": (316, 2, (11, 0), (), "ShowRevisions", None),
		"ShowSpellingErrors": (73, 2, (11, 0), (), "ShowSpellingErrors", None),
		"ShowSummary": (76, 2, (11, 0), (), "ShowSummary", None),
		# Method 'Signatures' returns object of type 'SignatureSet'
		"Signatures": (339, 2, (9, 0), (), "Signatures", '{000C0410-0000-0000-C000-000000000046}'),
		# Method 'SmartDocument' returns object of type 'SmartDocument'
		"SmartDocument": (462, 2, (9, 0), (), "SmartDocument", '{000C0377-0000-0000-C000-000000000046}'),
		# Method 'SmartTags' returns object of type 'SmartTags'
		"SmartTags": (346, 2, (9, 0), (), "SmartTags", '{000209EE-0000-0000-C000-000000000046}'),
		"SmartTagsAsXMLProps": (348, 2, (11, 0), (), "SmartTagsAsXMLProps", None),
		"SnapToGrid": (300, 2, (11, 0), (), "SnapToGrid", None),
		"SnapToShapes": (301, 2, (11, 0), (), "SnapToShapes", None),
		"SpellingChecked": (71, 2, (11, 0), (), "SpellingChecked", None),
		# Method 'SpellingErrors' returns object of type 'ProofreadingErrors'
		"SpellingErrors": (98, 2, (9, 0), (), "SpellingErrors", '{000209BB-0000-0000-C000-000000000046}'),
		# Method 'StoryRanges' returns object of type 'StoryRanges'
		"StoryRanges": (56, 2, (9, 0), (), "StoryRanges", '{0002098C-0000-0000-C000-000000000046}'),
		# Method 'StyleSheets' returns object of type 'StyleSheets'
		"StyleSheets": (360, 2, (9, 0), (), "StyleSheets", '{07B7CC7E-E66C-11D3-9454-00105AA31A08}'),
		"StyleSortMethod": (493, 2, (3, 0), (), "StyleSortMethod", None),
		# Method 'Styles' returns object of type 'Styles'
		"Styles": (22, 2, (9, 0), (), "Styles", '{0002092D-0000-0000-C000-000000000046}'),
		# Method 'Subdocuments' returns object of type 'Subdocuments'
		"Subdocuments": (45, 2, (9, 0), (), "Subdocuments", '{00020988-0000-0000-C000-000000000046}'),
		"SummaryLength": (78, 2, (3, 0), (), "SummaryLength", None),
		"SummaryViewMode": (77, 2, (3, 0), (), "SummaryViewMode", None),
		# Method 'Sync' returns object of type 'Sync'
		"Sync": (466, 2, (9, 0), (), "Sync", '{000C0386-0000-0000-C000-000000000046}'),
		# Method 'Tables' returns object of type 'Tables'
		"Tables": (6, 2, (9, 0), (), "Tables", '{0002094D-0000-0000-C000-000000000046}'),
		# Method 'TablesOfAuthorities' returns object of type 'TablesOfAuthorities'
		"TablesOfAuthorities": (32, 2, (9, 0), (), "TablesOfAuthorities", '{00020912-0000-0000-C000-000000000046}'),
		# Method 'TablesOfAuthoritiesCategories' returns object of type 'TablesOfAuthoritiesCategories'
		"TablesOfAuthoritiesCategories": (38, 2, (9, 0), (), "TablesOfAuthoritiesCategories", '{00020976-0000-0000-C000-000000000046}'),
		# Method 'TablesOfContents' returns object of type 'TablesOfContents'
		"TablesOfContents": (31, 2, (9, 0), (), "TablesOfContents", '{00020914-0000-0000-C000-000000000046}'),
		# Method 'TablesOfFigures' returns object of type 'TablesOfFigures'
		"TablesOfFigures": (25, 2, (9, 0), (), "TablesOfFigures", '{00020922-0000-0000-C000-000000000046}'),
		"TextEncoding": (357, 2, (3, 0), (), "TextEncoding", None),
		"TextLineEnding": (358, 2, (3, 0), (), "TextLineEnding", None),
		"TrackFormatting": (502, 2, (11, 0), (), "TrackFormatting", None),
		"TrackMoves": (499, 2, (11, 0), (), "TrackMoves", None),
		"TrackRevisions": (314, 2, (11, 0), (), "TrackRevisions", None),
		"Type": (10, 2, (3, 0), (), "Type", None),
		"UpdateStylesOnOpen": (66, 2, (11, 0), (), "UpdateStylesOnOpen", None),
		"UseMathDefaults": (560, 2, (11, 0), (), "UseMathDefaults", None),
		"UserControl": (92, 2, (11, 0), (), "UserControl", None),
		"VBASigned": (335, 2, (11, 0), (), "VBASigned", None),
		# Method 'VBProject' returns object of type 'VBProject'
		"VBProject": (99, 2, (13, 0), (), "VBProject", '{0002E169-0000-0000-C000-000000000046}'),
		# Method 'Variables' returns object of type 'Variables'
		"Variables": (26, 2, (9, 0), (), "Variables", '{00020965-0000-0000-C000-000000000046}'),
		# Method 'Versions' returns object of type 'Versions'
		"Versions": (75, 2, (9, 0), (), "Versions", '{000209B3-0000-0000-C000-000000000046}'),
		# Method 'WebOptions' returns object of type 'WebOptions'
		"WebOptions": (330, 2, (9, 0), (), "WebOptions", '{000209E4-0000-0000-C000-000000000046}'),
		# Method 'Windows' returns object of type 'Windows'
		"Windows": (34, 2, (9, 0), (), "Windows", '{00020961-0000-0000-C000-000000000046}'),
		"WordOpenXML": (542, 2, (8, 0), (), "WordOpenXML", None),
		# Method 'Words' returns object of type 'Words'
		"Words": (17, 2, (9, 0), (), "Words", '{0002095C-0000-0000-C000-000000000046}'),
		"WriteReserved": (88, 2, (11, 0), (), "WriteReserved", None),
		"XMLHideNamespaces": (477, 2, (11, 0), (), "XMLHideNamespaces", None),
		# Method 'XMLNodes' returns object of type 'XMLNodes'
		"XMLNodes": (460, 2, (9, 0), (), "XMLNodes", '{D36C1F42-7044-4B9E-9CA3-85919454DB04}'),
		"XMLSaveDataOnly": (473, 2, (11, 0), (), "XMLSaveDataOnly", None),
		"XMLSaveThroughXSLT": (475, 2, (8, 0), (), "XMLSaveThroughXSLT", None),
		# Method 'XMLSchemaReferences' returns object of type 'XMLSchemaReferences'
		"XMLSchemaReferences": (461, 2, (9, 0), (), "XMLSchemaReferences", '{356B06EC-4908-42A4-81FC-4B5A51F3483B}'),
		# Method 'XMLSchemaViolations' returns object of type 'XMLNodes'
		"XMLSchemaViolations": (490, 2, (9, 0), (), "XMLSchemaViolations", '{D36C1F42-7044-4B9E-9CA3-85919454DB04}'),
		"XMLShowAdvancedErrors": (478, 2, (11, 0), (), "XMLShowAdvancedErrors", None),
		"XMLUseXSLTWhenSaving": (474, 2, (11, 0), (), "XMLUseXSLTWhenSaving", None),
		"_CodeName": (-2147418112, 2, (8, 0), (), "_CodeName", None),
	}
	_prop_map_put_ = {
		"AttachedTemplate": ((67, LCID, 4, 0),()),
		"AutoFormatOverride": ((472, LCID, 4, 0),()),
		"AutoHyphenation": ((11, LCID, 4, 0),()),
		"Background": ((69, LCID, 4, 0),()),
		"ClickAndTypeParagraphStyle": ((328, LCID, 4, 0),()),
		"ConsecutiveHyphensLimit": ((14, LCID, 4, 0),()),
		"DefaultTabStop": ((48, LCID, 4, 0),()),
		"DefaultTargetFrame": ((340, LCID, 4, 0),()),
		"DisableFeatures": ((337, LCID, 4, 0),()),
		"DisableFeaturesIntroducedAfter": ((343, LCID, 4, 0),()),
		"DoNotEmbedSystemFonts": ((338, LCID, 4, 0),()),
		"EmbedLinguisticData": ((377, LCID, 4, 0),()),
		"EmbedSmartTags": ((347, LCID, 4, 0),()),
		"EmbedTrueTypeFonts": ((50, LCID, 4, 0),()),
		"EncryptionProvider": ((559, LCID, 4, 0),()),
		"EnforceStyle": ((471, LCID, 4, 0),()),
		"FarEastLineBreakLanguage": ((326, LCID, 4, 0),()),
		"FarEastLineBreakLevel": ((311, LCID, 4, 0),()),
		"Final": ((527, LCID, 4, 0),()),
		"FormattingShowClear": ((449, LCID, 4, 0),()),
		"FormattingShowFilter": ((452, LCID, 4, 0),()),
		"FormattingShowFont": ((448, LCID, 4, 0),()),
		"FormattingShowNextLevel": ((522, LCID, 4, 0),()),
		"FormattingShowNumbering": ((451, LCID, 4, 0),()),
		"FormattingShowParagraph": ((450, LCID, 4, 0),()),
		"FormattingShowUserStyleName": ((523, LCID, 4, 0),()),
		"GrammarChecked": ((70, LCID, 4, 0),()),
		"GridDistanceHorizontal": ((302, LCID, 4, 0),()),
		"GridDistanceVertical": ((303, LCID, 4, 0),()),
		"GridOriginFromMargin": ((308, LCID, 4, 0),()),
		"GridOriginHorizontal": ((304, LCID, 4, 0),()),
		"GridOriginVertical": ((305, LCID, 4, 0),()),
		"GridSpaceBetweenHorizontalLines": ((306, LCID, 4, 0),()),
		"GridSpaceBetweenVerticalLines": ((307, LCID, 4, 0),()),
		"HasMailer": ((93, LCID, 4, 0),()),
		"HasRoutingSlip": ((35, LCID, 4, 0),()),
		"HyphenateCaps": ((12, LCID, 4, 0),()),
		"HyphenationZone": ((13, LCID, 4, 0),()),
		"JustificationMode": ((310, LCID, 4, 0),()),
		"KerningByAlgorithm": ((309, LCID, 4, 0),()),
		"Kind": ((43, LCID, 4, 0),()),
		"LanguageDetected": ((321, LCID, 4, 0),()),
		"LockQuickStyleSet": ((518, LCID, 4, 0),()),
		"LockTheme": ((517, LCID, 4, 0),()),
		"NoLineBreakAfter": ((313, LCID, 4, 0),()),
		"NoLineBreakBefore": ((312, LCID, 4, 0),()),
		"OMathBreakBin": ((528, LCID, 4, 0),()),
		"OMathBreakSub": ((529, LCID, 4, 0),()),
		"OMathFontName": ((555, LCID, 4, 0),()),
		"OMathIntSubSupLim": ((536, LCID, 4, 0),()),
		"OMathJc": ((530, LCID, 4, 0),()),
		"OMathLeftMargin": ((531, LCID, 4, 0),()),
		"OMathNarySupSubLim": ((537, LCID, 4, 0),()),
		"OMathRightMargin": ((532, LCID, 4, 0),()),
		"OMathSmallFrac": ((539, LCID, 4, 0),()),
		"OMathWrap": ((535, LCID, 4, 0),()),
		"OptimizeForWord97": ((334, LCID, 4, 0),()),
		"PageSetup": ((1101, LCID, 4, 0),()),
		"Password": ((85, LCID, 4, 0),()),
		"PrintFormsData": ((83, LCID, 4, 0),()),
		"PrintFractionalWidths": ((79, LCID, 4, 0),()),
		"PrintPostScriptOverText": ((80, LCID, 4, 0),()),
		"PrintRevisions": ((315, LCID, 4, 0),()),
		"ReadOnlyRecommended": ((52, LCID, 4, 0),()),
		"ReadingLayoutSizeX": ((491, LCID, 4, 0),()),
		"ReadingLayoutSizeY": ((492, LCID, 4, 0),()),
		"ReadingModeLayoutFrozen": ((481, LCID, 4, 0),()),
		"RemoveDateAndTime": ((484, LCID, 4, 0),()),
		"RemovePersonalInformation": ((344, LCID, 4, 0),()),
		"SaveEncoding": ((333, LCID, 4, 0),()),
		"SaveFormsData": ((51, LCID, 4, 0),()),
		"SaveSubsetFonts": ((53, LCID, 4, 0),()),
		"Saved": ((40, LCID, 4, 0),()),
		"ShowGrammaticalErrors": ((72, LCID, 4, 0),()),
		"ShowRevisions": ((316, LCID, 4, 0),()),
		"ShowSpellingErrors": ((73, LCID, 4, 0),()),
		"ShowSummary": ((76, LCID, 4, 0),()),
		"SmartTagsAsXMLProps": ((348, LCID, 4, 0),()),
		"SnapToGrid": ((300, LCID, 4, 0),()),
		"SnapToShapes": ((301, LCID, 4, 0),()),
		"SpellingChecked": ((71, LCID, 4, 0),()),
		"StyleSortMethod": ((493, LCID, 4, 0),()),
		"SummaryLength": ((78, LCID, 4, 0),()),
		"SummaryViewMode": ((77, LCID, 4, 0),()),
		"TextEncoding": ((357, LCID, 4, 0),()),
		"TextLineEnding": ((358, LCID, 4, 0),()),
		"TrackFormatting": ((502, LCID, 4, 0),()),
		"TrackMoves": ((499, LCID, 4, 0),()),
		"TrackRevisions": ((314, LCID, 4, 0),()),
		"UpdateStylesOnOpen": ((66, LCID, 4, 0),()),
		"UseMathDefaults": ((560, LCID, 4, 0),()),
		"UserControl": ((92, LCID, 4, 0),()),
		"WritePassword": ((86, LCID, 4, 0),()),
		"XMLHideNamespaces": ((477, LCID, 4, 0),()),
		"XMLSaveDataOnly": ((473, LCID, 4, 0),()),
		"XMLSaveThroughXSLT": ((475, LCID, 4, 0),()),
		"XMLShowAdvancedErrors": ((478, LCID, 4, 0),()),
		"XMLUseXSLTWhenSaving": ((474, LCID, 4, 0),()),
		"_CodeName": ((-2147418112, LCID, 4, 0),()),
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

win32com.client.CLSIDToClass.RegisterCLSID( "{0002096B-0000-0000-C000-000000000046}", _Document )
# -*- coding: mbcs -*-
# Created by makepy.py version 0.5.01
# By python version 3.3.1 (v3.3.1:d9893d13c628, Apr  6 2013, 20:25:12) [MSC v.1600 32 bit (Intel)]
# From type library '{00020905-0000-0000-C000-000000000046}'
# On Mon Apr 15 10:41:20 2013
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

_Document_vtables_dispatch_ = 1
_Document_vtables_ = [
	(( 'Name' , 'prop' , ), 0, (0, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 28 , (3, 0, None, None) , 0 , )),
	(( 'Application' , 'prop' , ), 1, (1, (), [ (16397, 10, None, "IID('{000209FF-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 32 , (3, 0, None, None) , 0 , )),
	(( 'Creator' , 'prop' , ), 1001, (1001, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 36 , (3, 0, None, None) , 0 , )),
	(( 'Parent' , 'prop' , ), 1002, (1002, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 40 , (3, 0, None, None) , 0 , )),
	(( 'BuiltInDocumentProperties' , 'prop' , ), 1000, (1000, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 44 , (3, 0, None, None) , 0 , )),
	(( 'CustomDocumentProperties' , 'prop' , ), 2, (2, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 48 , (3, 0, None, None) , 0 , )),
	(( 'Path' , 'prop' , ), 3, (3, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 52 , (3, 0, None, None) , 0 , )),
	(( 'Bookmarks' , 'prop' , ), 4, (4, (), [ (16393, 10, None, "IID('{00020967-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 56 , (3, 0, None, None) , 0 , )),
	(( 'Tables' , 'prop' , ), 6, (6, (), [ (16393, 10, None, "IID('{0002094D-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 60 , (3, 0, None, None) , 0 , )),
	(( 'Footnotes' , 'prop' , ), 7, (7, (), [ (16393, 10, None, "IID('{00020942-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 64 , (3, 0, None, None) , 0 , )),
	(( 'Endnotes' , 'prop' , ), 8, (8, (), [ (16393, 10, None, "IID('{00020941-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 68 , (3, 0, None, None) , 0 , )),
	(( 'Comments' , 'prop' , ), 9, (9, (), [ (16393, 10, None, "IID('{00020940-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 72 , (3, 0, None, None) , 0 , )),
	(( 'Type' , 'prop' , ), 10, (10, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 76 , (3, 0, None, None) , 0 , )),
	(( 'AutoHyphenation' , 'prop' , ), 11, (11, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 80 , (3, 0, None, None) , 0 , )),
	(( 'AutoHyphenation' , 'prop' , ), 11, (11, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 84 , (3, 0, None, None) , 0 , )),
	(( 'HyphenateCaps' , 'prop' , ), 12, (12, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 88 , (3, 0, None, None) , 0 , )),
	(( 'HyphenateCaps' , 'prop' , ), 12, (12, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 92 , (3, 0, None, None) , 0 , )),
	(( 'HyphenationZone' , 'prop' , ), 13, (13, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 96 , (3, 0, None, None) , 0 , )),
	(( 'HyphenationZone' , 'prop' , ), 13, (13, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 100 , (3, 0, None, None) , 0 , )),
	(( 'ConsecutiveHyphensLimit' , 'prop' , ), 14, (14, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 104 , (3, 0, None, None) , 0 , )),
	(( 'ConsecutiveHyphensLimit' , 'prop' , ), 14, (14, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 108 , (3, 0, None, None) , 0 , )),
	(( 'Sections' , 'prop' , ), 15, (15, (), [ (16393, 10, None, "IID('{0002095A-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 112 , (3, 0, None, None) , 0 , )),
	(( 'Paragraphs' , 'prop' , ), 16, (16, (), [ (16393, 10, None, "IID('{00020958-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 116 , (3, 0, None, None) , 0 , )),
	(( 'Words' , 'prop' , ), 17, (17, (), [ (16393, 10, None, "IID('{0002095C-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 120 , (3, 0, None, None) , 0 , )),
	(( 'Sentences' , 'prop' , ), 18, (18, (), [ (16393, 10, None, "IID('{0002095B-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 124 , (3, 0, None, None) , 0 , )),
	(( 'Characters' , 'prop' , ), 19, (19, (), [ (16393, 10, None, "IID('{0002095D-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 128 , (3, 0, None, None) , 0 , )),
	(( 'Fields' , 'prop' , ), 20, (20, (), [ (16393, 10, None, "IID('{00020930-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 132 , (3, 0, None, None) , 0 , )),
	(( 'FormFields' , 'prop' , ), 21, (21, (), [ (16393, 10, None, "IID('{00020929-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 136 , (3, 0, None, None) , 0 , )),
	(( 'Styles' , 'prop' , ), 22, (22, (), [ (16393, 10, None, "IID('{0002092D-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 140 , (3, 0, None, None) , 0 , )),
	(( 'Frames' , 'prop' , ), 23, (23, (), [ (16393, 10, None, "IID('{0002092B-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 144 , (3, 0, None, None) , 0 , )),
	(( 'TablesOfFigures' , 'prop' , ), 25, (25, (), [ (16393, 10, None, "IID('{00020922-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 148 , (3, 0, None, None) , 0 , )),
	(( 'Variables' , 'prop' , ), 26, (26, (), [ (16393, 10, None, "IID('{00020965-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 152 , (3, 0, None, None) , 0 , )),
	(( 'MailMerge' , 'prop' , ), 27, (27, (), [ (16393, 10, None, "IID('{00020920-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 156 , (3, 0, None, None) , 0 , )),
	(( 'Envelope' , 'prop' , ), 28, (28, (), [ (16393, 10, None, "IID('{00020918-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 160 , (3, 0, None, None) , 0 , )),
	(( 'FullName' , 'prop' , ), 29, (29, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 164 , (3, 0, None, None) , 0 , )),
	(( 'Revisions' , 'prop' , ), 30, (30, (), [ (16393, 10, None, "IID('{00020980-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 168 , (3, 0, None, None) , 0 , )),
	(( 'TablesOfContents' , 'prop' , ), 31, (31, (), [ (16393, 10, None, "IID('{00020914-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 172 , (3, 0, None, None) , 0 , )),
	(( 'TablesOfAuthorities' , 'prop' , ), 32, (32, (), [ (16393, 10, None, "IID('{00020912-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 176 , (3, 0, None, None) , 0 , )),
	(( 'PageSetup' , 'prop' , ), 1101, (1101, (), [ (16393, 10, None, "IID('{00020971-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 180 , (3, 0, None, None) , 0 , )),
	(( 'PageSetup' , 'prop' , ), 1101, (1101, (), [ (9, 1, None, "IID('{00020971-0000-0000-C000-000000000046}')") , ], 1 , 4 , 4 , 0 , 184 , (3, 0, None, None) , 0 , )),
	(( 'Windows' , 'prop' , ), 34, (34, (), [ (16393, 10, None, "IID('{00020961-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 188 , (3, 0, None, None) , 0 , )),
	(( 'HasRoutingSlip' , 'prop' , ), 35, (35, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 192 , (3, 0, None, None) , 64 , )),
	(( 'HasRoutingSlip' , 'prop' , ), 35, (35, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 196 , (3, 0, None, None) , 64 , )),
	(( 'RoutingSlip' , 'prop' , ), 36, (36, (), [ (16393, 10, None, "IID('{00020969-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 200 , (3, 0, None, None) , 64 , )),
	(( 'Routed' , 'prop' , ), 37, (37, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 204 , (3, 0, None, None) , 64 , )),
	(( 'TablesOfAuthoritiesCategories' , 'prop' , ), 38, (38, (), [ (16393, 10, None, "IID('{00020976-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 208 , (3, 0, None, None) , 0 , )),
	(( 'Indexes' , 'prop' , ), 39, (39, (), [ (16393, 10, None, "IID('{0002097C-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 212 , (3, 0, None, None) , 0 , )),
	(( 'Saved' , 'prop' , ), 40, (40, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 216 , (3, 0, None, None) , 0 , )),
	(( 'Saved' , 'prop' , ), 40, (40, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 220 , (3, 0, None, None) , 0 , )),
	(( 'Content' , 'prop' , ), 41, (41, (), [ (16393, 10, None, "IID('{0002095E-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 224 , (3, 0, None, None) , 0 , )),
	(( 'ActiveWindow' , 'prop' , ), 42, (42, (), [ (16393, 10, None, "IID('{00020962-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 228 , (3, 0, None, None) , 0 , )),
	(( 'Kind' , 'prop' , ), 43, (43, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 232 , (3, 0, None, None) , 0 , )),
	(( 'Kind' , 'prop' , ), 43, (43, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 236 , (3, 0, None, None) , 0 , )),
	(( 'ReadOnly' , 'prop' , ), 44, (44, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 240 , (3, 0, None, None) , 0 , )),
	(( 'Subdocuments' , 'prop' , ), 45, (45, (), [ (16393, 10, None, "IID('{00020988-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 244 , (3, 0, None, None) , 0 , )),
	(( 'IsMasterDocument' , 'prop' , ), 46, (46, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 248 , (3, 0, None, None) , 0 , )),
	(( 'DefaultTabStop' , 'prop' , ), 48, (48, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 252 , (3, 0, None, None) , 0 , )),
	(( 'DefaultTabStop' , 'prop' , ), 48, (48, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 256 , (3, 0, None, None) , 0 , )),
	(( 'EmbedTrueTypeFonts' , 'prop' , ), 50, (50, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 260 , (3, 0, None, None) , 0 , )),
	(( 'EmbedTrueTypeFonts' , 'prop' , ), 50, (50, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 264 , (3, 0, None, None) , 0 , )),
	(( 'SaveFormsData' , 'prop' , ), 51, (51, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 268 , (3, 0, None, None) , 0 , )),
	(( 'SaveFormsData' , 'prop' , ), 51, (51, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 272 , (3, 0, None, None) , 0 , )),
	(( 'ReadOnlyRecommended' , 'prop' , ), 52, (52, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 276 , (3, 0, None, None) , 0 , )),
	(( 'ReadOnlyRecommended' , 'prop' , ), 52, (52, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 280 , (3, 0, None, None) , 0 , )),
	(( 'SaveSubsetFonts' , 'prop' , ), 53, (53, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 284 , (3, 0, None, None) , 0 , )),
	(( 'SaveSubsetFonts' , 'prop' , ), 53, (53, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 288 , (3, 0, None, None) , 0 , )),
	(( 'Compatibility' , 'Type' , 'prop' , ), 55, (55, (), [ (3, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 292 , (3, 0, None, None) , 0 , )),
	(( 'Compatibility' , 'Type' , 'prop' , ), 55, (55, (), [ (3, 1, None, None) , 
			 (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 296 , (3, 0, None, None) , 0 , )),
	(( 'StoryRanges' , 'prop' , ), 56, (56, (), [ (16393, 10, None, "IID('{0002098C-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 300 , (3, 0, None, None) , 0 , )),
	(( 'CommandBars' , 'prop' , ), 57, (57, (), [ (16397, 10, None, "IID('{55F88893-7708-11D1-ACEB-006008961DA5}')") , ], 1 , 2 , 4 , 0 , 304 , (3, 0, None, None) , 0 , )),
	(( 'IsSubdocument' , 'prop' , ), 58, (58, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 308 , (3, 0, None, None) , 0 , )),
	(( 'SaveFormat' , 'prop' , ), 59, (59, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 312 , (3, 0, None, None) , 0 , )),
	(( 'ProtectionType' , 'prop' , ), 60, (60, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 316 , (3, 0, None, None) , 0 , )),
	(( 'Hyperlinks' , 'prop' , ), 61, (61, (), [ (16393, 10, None, "IID('{0002099C-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 320 , (3, 0, None, None) , 0 , )),
	(( 'Shapes' , 'prop' , ), 62, (62, (), [ (16393, 10, None, "IID('{0002099F-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 324 , (3, 0, None, None) , 0 , )),
	(( 'ListTemplates' , 'prop' , ), 63, (63, (), [ (16393, 10, None, "IID('{00020990-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 328 , (3, 0, None, None) , 0 , )),
	(( 'Lists' , 'prop' , ), 64, (64, (), [ (16393, 10, None, "IID('{00020993-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 332 , (3, 0, None, None) , 0 , )),
	(( 'UpdateStylesOnOpen' , 'prop' , ), 66, (66, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 336 , (3, 0, None, None) , 0 , )),
	(( 'UpdateStylesOnOpen' , 'prop' , ), 66, (66, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 340 , (3, 0, None, None) , 0 , )),
	(( 'AttachedTemplate' , 'prop' , ), 67, (67, (), [ (16396, 10, None, None) , ], 1 , 2 , 4 , 0 , 344 , (3, 0, None, None) , 1024 , )),
	(( 'AttachedTemplate' , 'prop' , ), 67, (67, (), [ (16396, 1, None, None) , ], 1 , 4 , 4 , 0 , 348 , (3, 0, None, None) , 1024 , )),
	(( 'InlineShapes' , 'prop' , ), 68, (68, (), [ (16393, 10, None, "IID('{000209A9-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 352 , (3, 0, None, None) , 0 , )),
	(( 'Background' , 'prop' , ), 69, (69, (), [ (16393, 10, None, "IID('{000209A0-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 356 , (3, 0, None, None) , 0 , )),
	(( 'Background' , 'prop' , ), 69, (69, (), [ (9, 1, None, "IID('{000209A0-0000-0000-C000-000000000046}')") , ], 1 , 4 , 4 , 0 , 360 , (3, 0, None, None) , 0 , )),
	(( 'GrammarChecked' , 'prop' , ), 70, (70, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 364 , (3, 0, None, None) , 0 , )),
	(( 'GrammarChecked' , 'prop' , ), 70, (70, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 368 , (3, 0, None, None) , 0 , )),
	(( 'SpellingChecked' , 'prop' , ), 71, (71, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 372 , (3, 0, None, None) , 0 , )),
	(( 'SpellingChecked' , 'prop' , ), 71, (71, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 376 , (3, 0, None, None) , 0 , )),
	(( 'ShowGrammaticalErrors' , 'prop' , ), 72, (72, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 380 , (3, 0, None, None) , 0 , )),
	(( 'ShowGrammaticalErrors' , 'prop' , ), 72, (72, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 384 , (3, 0, None, None) , 0 , )),
	(( 'ShowSpellingErrors' , 'prop' , ), 73, (73, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 388 , (3, 0, None, None) , 0 , )),
	(( 'ShowSpellingErrors' , 'prop' , ), 73, (73, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 392 , (3, 0, None, None) , 0 , )),
	(( 'Versions' , 'prop' , ), 75, (75, (), [ (16393, 10, None, "IID('{000209B3-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 396 , (3, 0, None, None) , 64 , )),
	(( 'ShowSummary' , 'prop' , ), 76, (76, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 400 , (3, 0, None, None) , 64 , )),
	(( 'ShowSummary' , 'prop' , ), 76, (76, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 404 , (3, 0, None, None) , 64 , )),
	(( 'SummaryViewMode' , 'prop' , ), 77, (77, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 408 , (3, 0, None, None) , 64 , )),
	(( 'SummaryViewMode' , 'prop' , ), 77, (77, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 412 , (3, 0, None, None) , 64 , )),
	(( 'SummaryLength' , 'prop' , ), 78, (78, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 416 , (3, 0, None, None) , 64 , )),
	(( 'SummaryLength' , 'prop' , ), 78, (78, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 420 , (3, 0, None, None) , 64 , )),
	(( 'PrintFractionalWidths' , 'prop' , ), 79, (79, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 424 , (3, 0, None, None) , 0 , )),
	(( 'PrintFractionalWidths' , 'prop' , ), 79, (79, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 428 , (3, 0, None, None) , 0 , )),
	(( 'PrintPostScriptOverText' , 'prop' , ), 80, (80, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 432 , (3, 0, None, None) , 0 , )),
	(( 'PrintPostScriptOverText' , 'prop' , ), 80, (80, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 436 , (3, 0, None, None) , 0 , )),
	(( 'Container' , 'prop' , ), 82, (82, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 440 , (3, 0, None, None) , 0 , )),
	(( 'PrintFormsData' , 'prop' , ), 83, (83, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 444 , (3, 0, None, None) , 0 , )),
	(( 'PrintFormsData' , 'prop' , ), 83, (83, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 448 , (3, 0, None, None) , 0 , )),
	(( 'ListParagraphs' , 'prop' , ), 84, (84, (), [ (16393, 10, None, "IID('{00020991-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 452 , (3, 0, None, None) , 0 , )),
	(( 'Password' , ), 85, (85, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 456 , (3, 0, None, None) , 0 , )),
	(( 'WritePassword' , ), 86, (86, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 460 , (3, 0, None, None) , 0 , )),
	(( 'HasPassword' , 'prop' , ), 87, (87, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 464 , (3, 0, None, None) , 0 , )),
	(( 'WriteReserved' , 'prop' , ), 88, (88, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 468 , (3, 0, None, None) , 0 , )),
	(( 'ActiveWritingStyle' , 'LanguageID' , 'prop' , ), 90, (90, (), [ (16396, 1, None, None) , 
			 (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 472 , (3, 0, None, None) , 0 , )),
	(( 'ActiveWritingStyle' , 'LanguageID' , 'prop' , ), 90, (90, (), [ (16396, 1, None, None) , 
			 (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 476 , (3, 0, None, None) , 0 , )),
	(( 'UserControl' , 'prop' , ), 92, (92, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 480 , (3, 0, None, None) , 0 , )),
	(( 'UserControl' , 'prop' , ), 92, (92, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 484 , (3, 0, None, None) , 0 , )),
	(( 'HasMailer' , 'prop' , ), 93, (93, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 488 , (3, 0, None, None) , 0 , )),
	(( 'HasMailer' , 'prop' , ), 93, (93, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 492 , (3, 0, None, None) , 0 , )),
	(( 'Mailer' , 'prop' , ), 94, (94, (), [ (16393, 10, None, "IID('{000209BD-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 496 , (3, 0, None, None) , 0 , )),
	(( 'ReadabilityStatistics' , 'prop' , ), 96, (96, (), [ (16393, 10, None, "IID('{000209AE-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 500 , (3, 0, None, None) , 0 , )),
	(( 'GrammaticalErrors' , 'prop' , ), 97, (97, (), [ (16393, 10, None, "IID('{000209BB-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 504 , (3, 0, None, None) , 0 , )),
	(( 'SpellingErrors' , 'prop' , ), 98, (98, (), [ (16393, 10, None, "IID('{000209BB-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 508 , (3, 0, None, None) , 0 , )),
	(( 'VBProject' , 'prop' , ), 99, (99, (), [ (16397, 10, None, "IID('{0002E169-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 512 , (3, 0, None, None) , 0 , )),
	(( 'FormsDesign' , 'prop' , ), 100, (100, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 516 , (3, 0, None, None) , 0 , )),
	(( '_CodeName' , 'prop' , ), -2147418112, (-2147418112, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 520 , (3, 0, None, None) , 1024 , )),
	(( '_CodeName' , 'prop' , ), -2147418112, (-2147418112, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 524 , (3, 0, None, None) , 1024 , )),
	(( 'CodeName' , 'prop' , ), 262, (262, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 528 , (3, 0, None, None) , 0 , )),
	(( 'SnapToGrid' , 'prop' , ), 300, (300, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 532 , (3, 0, None, None) , 0 , )),
	(( 'SnapToGrid' , 'prop' , ), 300, (300, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 536 , (3, 0, None, None) , 0 , )),
	(( 'SnapToShapes' , 'prop' , ), 301, (301, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 540 , (3, 0, None, None) , 0 , )),
	(( 'SnapToShapes' , 'prop' , ), 301, (301, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 544 , (3, 0, None, None) , 0 , )),
	(( 'GridDistanceHorizontal' , 'prop' , ), 302, (302, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 548 , (3, 0, None, None) , 0 , )),
	(( 'GridDistanceHorizontal' , 'prop' , ), 302, (302, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 552 , (3, 0, None, None) , 0 , )),
	(( 'GridDistanceVertical' , 'prop' , ), 303, (303, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 556 , (3, 0, None, None) , 0 , )),
	(( 'GridDistanceVertical' , 'prop' , ), 303, (303, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 560 , (3, 0, None, None) , 0 , )),
	(( 'GridOriginHorizontal' , 'prop' , ), 304, (304, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 564 , (3, 0, None, None) , 0 , )),
	(( 'GridOriginHorizontal' , 'prop' , ), 304, (304, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 568 , (3, 0, None, None) , 0 , )),
	(( 'GridOriginVertical' , 'prop' , ), 305, (305, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 572 , (3, 0, None, None) , 0 , )),
	(( 'GridOriginVertical' , 'prop' , ), 305, (305, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 576 , (3, 0, None, None) , 0 , )),
	(( 'GridSpaceBetweenHorizontalLines' , 'prop' , ), 306, (306, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 580 , (3, 0, None, None) , 0 , )),
	(( 'GridSpaceBetweenHorizontalLines' , 'prop' , ), 306, (306, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 584 , (3, 0, None, None) , 0 , )),
	(( 'GridSpaceBetweenVerticalLines' , 'prop' , ), 307, (307, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 588 , (3, 0, None, None) , 0 , )),
	(( 'GridSpaceBetweenVerticalLines' , 'prop' , ), 307, (307, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 592 , (3, 0, None, None) , 0 , )),
	(( 'GridOriginFromMargin' , 'prop' , ), 308, (308, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 596 , (3, 0, None, None) , 0 , )),
	(( 'GridOriginFromMargin' , 'prop' , ), 308, (308, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 600 , (3, 0, None, None) , 0 , )),
	(( 'KerningByAlgorithm' , 'prop' , ), 309, (309, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 604 , (3, 0, None, None) , 0 , )),
	(( 'KerningByAlgorithm' , 'prop' , ), 309, (309, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 608 , (3, 0, None, None) , 0 , )),
	(( 'JustificationMode' , 'prop' , ), 310, (310, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 612 , (3, 0, None, None) , 0 , )),
	(( 'JustificationMode' , 'prop' , ), 310, (310, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 616 , (3, 0, None, None) , 0 , )),
	(( 'FarEastLineBreakLevel' , 'prop' , ), 311, (311, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 620 , (3, 0, None, None) , 0 , )),
	(( 'FarEastLineBreakLevel' , 'prop' , ), 311, (311, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 624 , (3, 0, None, None) , 0 , )),
	(( 'NoLineBreakBefore' , 'prop' , ), 312, (312, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 628 , (3, 0, None, None) , 0 , )),
	(( 'NoLineBreakBefore' , 'prop' , ), 312, (312, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 632 , (3, 0, None, None) , 0 , )),
	(( 'NoLineBreakAfter' , 'prop' , ), 313, (313, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 636 , (3, 0, None, None) , 0 , )),
	(( 'NoLineBreakAfter' , 'prop' , ), 313, (313, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 640 , (3, 0, None, None) , 0 , )),
	(( 'TrackRevisions' , 'prop' , ), 314, (314, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 644 , (3, 0, None, None) , 0 , )),
	(( 'TrackRevisions' , 'prop' , ), 314, (314, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 648 , (3, 0, None, None) , 0 , )),
	(( 'PrintRevisions' , 'prop' , ), 315, (315, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 652 , (3, 0, None, None) , 0 , )),
	(( 'PrintRevisions' , 'prop' , ), 315, (315, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 656 , (3, 0, None, None) , 0 , )),
	(( 'ShowRevisions' , 'prop' , ), 316, (316, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 660 , (3, 0, None, None) , 0 , )),
	(( 'ShowRevisions' , 'prop' , ), 316, (316, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 664 , (3, 0, None, None) , 0 , )),
	(( 'Close' , 'SaveChanges' , 'OriginalFormat' , 'RouteDocument' , ), 1105, (1105, (), [ 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , ], 1 , 1 , 4 , 3 , 668 , (3, 0, None, None) , 0 , )),
	(( 'SaveAs2000' , 'FileName' , 'FileFormat' , 'LockComments' , 'Password' , 
			 'AddToRecentFiles' , 'WritePassword' , 'ReadOnlyRecommended' , 'EmbedTrueTypeFonts' , 'SaveNativePictureFormat' , 
			 'SaveFormsData' , 'SaveAsAOCELetter' , ), 102, (102, (), [ (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , ], 1 , 1 , 4 , 11 , 672 , (3, 0, None, None) , 64 , )),
	(( 'Repaginate' , ), 103, (103, (), [ ], 1 , 1 , 4 , 0 , 676 , (3, 0, None, None) , 0 , )),
	(( 'FitToPages' , ), 104, (104, (), [ ], 1 , 1 , 4 , 0 , 680 , (3, 0, None, None) , 0 , )),
	(( 'ManualHyphenation' , ), 105, (105, (), [ ], 1 , 1 , 4 , 0 , 684 , (3, 0, None, None) , 0 , )),
	(( 'Select' , ), 65535, (65535, (), [ ], 1 , 1 , 4 , 0 , 688 , (3, 0, None, None) , 0 , )),
	(( 'DataForm' , ), 106, (106, (), [ ], 1 , 1 , 4 , 0 , 692 , (3, 0, None, None) , 0 , )),
	(( 'Route' , ), 107, (107, (), [ ], 1 , 1 , 4 , 0 , 696 , (3, 0, None, None) , 64 , )),
	(( 'Save' , ), 108, (108, (), [ ], 1 , 1 , 4 , 0 , 700 , (3, 0, None, None) , 0 , )),
	(( 'PrintOutOld' , 'Background' , 'Append' , 'Range' , 'OutputFileName' , 
			 'From' , 'To' , 'Item' , 'Copies' , 'Pages' , 
			 'PageType' , 'PrintToFile' , 'Collate' , 'ActivePrinterMacGX' , 'ManualDuplexPrint' , 
			 ), 109, (109, (), [ (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , ], 1 , 1 , 4 , 14 , 704 , (3, 0, None, None) , 64 , )),
	(( 'SendMail' , ), 110, (110, (), [ ], 1 , 1 , 4 , 0 , 708 , (3, 0, None, None) , 0 , )),
	(( 'Range' , 'Start' , 'End' , 'prop' , ), 2000, (2000, (), [ 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16393, 10, None, "IID('{0002095E-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 2 , 712 , (3, 0, None, None) , 0 , )),
	(( 'RunAutoMacro' , 'Which' , ), 112, (112, (), [ (3, 1, None, None) , ], 1 , 1 , 4 , 0 , 716 , (3, 0, None, None) , 0 , )),
	(( 'Activate' , ), 113, (113, (), [ ], 1 , 1 , 4 , 0 , 720 , (3, 0, None, None) , 0 , )),
	(( 'PrintPreview' , ), 114, (114, (), [ ], 1 , 1 , 4 , 0 , 724 , (3, 0, None, None) , 0 , )),
	(( 'GoTo' , 'What' , 'Which' , 'Count' , 'Name' , 
			 'prop' , ), 115, (115, (), [ (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16393, 10, None, "IID('{0002095E-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 4 , 728 , (3, 0, None, None) , 0 , )),
	(( 'Undo' , 'Times' , 'prop' , ), 116, (116, (), [ (16396, 17, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 1 , 732 , (3, 0, None, None) , 0 , )),
	(( 'Redo' , 'Times' , 'prop' , ), 117, (117, (), [ (16396, 17, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 1 , 736 , (3, 0, None, None) , 0 , )),
	(( 'ComputeStatistics' , 'Statistic' , 'IncludeFootnotesAndEndnotes' , 'prop' , ), 118, (118, (), [ 
			 (3, 1, None, None) , (16396, 17, None, None) , (16387, 10, None, None) , ], 1 , 1 , 4 , 1 , 740 , (3, 0, None, None) , 0 , )),
	(( 'MakeCompatibilityDefault' , ), 119, (119, (), [ ], 1 , 1 , 4 , 0 , 744 , (3, 0, None, None) , 0 , )),
	(( 'Protect2002' , 'Type' , 'NoReset' , 'Password' , ), 120, (120, (), [ 
			 (3, 1, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , ], 1 , 1 , 4 , 2 , 748 , (3, 0, None, None) , 64 , )),
	(( 'Unprotect' , 'Password' , ), 121, (121, (), [ (16396, 17, None, None) , ], 1 , 1 , 4 , 1 , 752 , (3, 0, None, None) , 0 , )),
	(( 'EditionOptions' , 'Type' , 'Option' , 'Name' , 'Format' , 
			 ), 122, (122, (), [ (3, 1, None, None) , (3, 1, None, None) , (8, 1, None, None) , (16396, 17, None, None) , ], 1 , 1 , 4 , 1 , 756 , (3, 0, None, None) , 0 , )),
	(( 'RunLetterWizard' , 'LetterContent' , 'WizardMode' , ), 123, (123, (), [ (16396, 17, None, None) , 
			 (16396, 17, None, None) , ], 1 , 1 , 4 , 2 , 760 , (3, 0, None, None) , 0 , )),
	(( 'GetLetterContent' , 'prop' , ), 124, (124, (), [ (16397, 10, None, "IID('{000209F1-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 0 , 764 , (3, 0, None, None) , 0 , )),
	(( 'SetLetterContent' , 'LetterContent' , ), 125, (125, (), [ (16396, 1, None, None) , ], 1 , 1 , 4 , 0 , 768 , (3, 0, None, None) , 0 , )),
	(( 'CopyStylesFromTemplate' , 'Template' , ), 126, (126, (), [ (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 772 , (3, 0, None, None) , 0 , )),
	(( 'UpdateStyles' , ), 127, (127, (), [ ], 1 , 1 , 4 , 0 , 776 , (3, 0, None, None) , 0 , )),
	(( 'CheckGrammar' , ), 131, (131, (), [ ], 1 , 1 , 4 , 0 , 780 , (3, 0, None, None) , 0 , )),
	(( 'CheckSpelling' , 'CustomDictionary' , 'IgnoreUppercase' , 'AlwaysSuggest' , 'CustomDictionary2' , 
			 'CustomDictionary3' , 'CustomDictionary4' , 'CustomDictionary5' , 'CustomDictionary6' , 'CustomDictionary7' , 
			 'CustomDictionary8' , 'CustomDictionary9' , 'CustomDictionary10' , ), 132, (132, (), [ (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , ], 1 , 1 , 4 , 12 , 784 , (3, 0, None, None) , 0 , )),
	(( 'FollowHyperlink' , 'Address' , 'SubAddress' , 'NewWindow' , 'AddHistory' , 
			 'ExtraInfo' , 'Method' , 'HeaderInfo' , ), 135, (135, (), [ (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , ], 1 , 1 , 4 , 7 , 788 , (3, 0, None, None) , 0 , )),
	(( 'AddToFavorites' , ), 136, (136, (), [ ], 1 , 1 , 4 , 0 , 792 , (3, 0, None, None) , 0 , )),
	(( 'Reload' , ), 137, (137, (), [ ], 1 , 1 , 4 , 0 , 796 , (3, 0, None, None) , 0 , )),
	(( 'AutoSummarize' , 'Length' , 'Mode' , 'UpdateProperties' , 'prop' , 
			 ), 138, (138, (), [ (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16393, 10, None, "IID('{0002095E-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 3 , 800 , (3, 0, None, None) , 64 , )),
	(( 'RemoveNumbers' , 'NumberType' , ), 140, (140, (), [ (16396, 17, None, None) , ], 1 , 1 , 4 , 1 , 804 , (3, 0, None, None) , 0 , )),
	(( 'ConvertNumbersToText' , 'NumberType' , ), 141, (141, (), [ (16396, 17, None, None) , ], 1 , 1 , 4 , 1 , 808 , (3, 0, None, None) , 0 , )),
	(( 'CountNumberedItems' , 'NumberType' , 'Level' , 'prop' , ), 142, (142, (), [ 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16387, 10, None, None) , ], 1 , 1 , 4 , 2 , 812 , (3, 0, None, None) , 0 , )),
	(( 'Post' , ), 143, (143, (), [ ], 1 , 1 , 4 , 0 , 816 , (3, 0, None, None) , 0 , )),
	(( 'ToggleFormsDesign' , ), 144, (144, (), [ ], 1 , 1 , 4 , 0 , 820 , (3, 0, None, None) , 0 , )),
	(( 'Compare2000' , 'Name' , ), 145, (145, (), [ (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 824 , (3, 0, None, None) , 64 , )),
	(( 'UpdateSummaryProperties' , ), 146, (146, (), [ ], 1 , 1 , 4 , 0 , 828 , (3, 0, None, None) , 0 , )),
	(( 'GetCrossReferenceItems' , 'ReferenceType' , 'prop' , ), 147, (147, (), [ (16396, 1, None, None) , 
			 (16396, 10, None, None) , ], 1 , 1 , 4 , 0 , 832 , (3, 0, None, None) , 0 , )),
	(( 'AutoFormat' , ), 148, (148, (), [ ], 1 , 1 , 4 , 0 , 836 , (3, 0, None, None) , 0 , )),
	(( 'ViewCode' , ), 149, (149, (), [ ], 1 , 1 , 4 , 0 , 840 , (3, 0, None, None) , 0 , )),
	(( 'ViewPropertyBrowser' , ), 150, (150, (), [ ], 1 , 1 , 4 , 0 , 844 , (3, 0, None, None) , 0 , )),
	(( 'ForwardMailer' , ), 250, (250, (), [ ], 1 , 1 , 4 , 0 , 848 , (3, 0, None, None) , 0 , )),
	(( 'Reply' , ), 251, (251, (), [ ], 1 , 1 , 4 , 0 , 852 , (3, 0, None, None) , 0 , )),
	(( 'ReplyAll' , ), 252, (252, (), [ ], 1 , 1 , 4 , 0 , 856 , (3, 0, None, None) , 0 , )),
	(( 'SendMailer' , 'FileFormat' , 'Priority' , ), 253, (253, (), [ (16396, 17, None, None) , 
			 (16396, 17, None, None) , ], 1 , 1 , 4 , 2 , 860 , (3, 0, None, None) , 0 , )),
	(( 'UndoClear' , ), 254, (254, (), [ ], 1 , 1 , 4 , 0 , 864 , (3, 0, None, None) , 0 , )),
	(( 'PresentIt' , ), 255, (255, (), [ ], 1 , 1 , 4 , 0 , 868 , (3, 0, None, None) , 0 , )),
	(( 'SendFax' , 'Address' , 'Subject' , ), 256, (256, (), [ (8, 1, None, None) , 
			 (16396, 17, None, None) , ], 1 , 1 , 4 , 1 , 872 , (3, 0, None, None) , 0 , )),
	(( 'Merge2000' , 'FileName' , ), 257, (257, (), [ (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 876 , (3, 0, None, None) , 64 , )),
	(( 'ClosePrintPreview' , ), 258, (258, (), [ ], 1 , 1 , 4 , 0 , 880 , (3, 0, None, None) , 0 , )),
	(( 'CheckConsistency' , ), 259, (259, (), [ ], 1 , 1 , 4 , 0 , 884 , (3, 0, None, None) , 0 , )),
	(( 'CreateLetterContent' , 'DateFormat' , 'IncludeHeaderFooter' , 'PageDesign' , 'LetterStyle' , 
			 'Letterhead' , 'LetterheadLocation' , 'LetterheadSize' , 'RecipientName' , 'RecipientAddress' , 
			 'Salutation' , 'SalutationType' , 'RecipientReference' , 'MailingInstructions' , 'AttentionLine' , 
			 'Subject' , 'CCList' , 'ReturnAddress' , 'SenderName' , 'Closing' , 
			 'SenderCompany' , 'SenderJobTitle' , 'SenderInitials' , 'EnclosureNumber' , 'InfoBlock' , 
			 'RecipientCode' , 'RecipientGender' , 'ReturnAddressShortForm' , 'SenderCity' , 'SenderCode' , 
			 'SenderGender' , 'SenderReference' , 'prop' , ), 260, (260, (), [ (8, 1, None, None) , 
			 (11, 1, None, None) , (8, 1, None, None) , (3, 1, None, None) , (11, 1, None, None) , (3, 1, None, None) , 
			 (4, 1, None, None) , (8, 1, None, None) , (8, 1, None, None) , (8, 1, None, None) , (3, 1, None, None) , 
			 (8, 1, None, None) , (8, 1, None, None) , (8, 1, None, None) , (8, 1, None, None) , (8, 1, None, None) , 
			 (8, 1, None, None) , (8, 1, None, None) , (8, 1, None, None) , (8, 1, None, None) , (8, 1, None, None) , 
			 (8, 1, None, None) , (3, 1, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16397, 10, None, "IID('{000209F1-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 8 , 888 , (3, 0, None, None) , 0 , )),
	(( 'AcceptAllRevisions' , ), 317, (317, (), [ ], 1 , 1 , 4 , 0 , 892 , (3, 0, None, None) , 0 , )),
	(( 'RejectAllRevisions' , ), 318, (318, (), [ ], 1 , 1 , 4 , 0 , 896 , (3, 0, None, None) , 0 , )),
	(( 'DetectLanguage' , ), 151, (151, (), [ ], 1 , 1 , 4 , 0 , 900 , (3, 0, None, None) , 0 , )),
	(( 'ApplyTheme' , 'Name' , ), 322, (322, (), [ (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 904 , (3, 0, None, None) , 0 , )),
	(( 'RemoveTheme' , ), 323, (323, (), [ ], 1 , 1 , 4 , 0 , 908 , (3, 0, None, None) , 0 , )),
	(( 'WebPagePreview' , ), 325, (325, (), [ ], 1 , 1 , 4 , 0 , 912 , (3, 0, None, None) , 0 , )),
	(( 'ReloadAs' , 'Encoding' , ), 331, (331, (), [ (3, 1, None, None) , ], 1 , 1 , 4 , 0 , 916 , (3, 0, None, None) , 0 , )),
	(( 'ActiveTheme' , 'prop' , ), 540, (540, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 920 , (3, 0, None, None) , 0 , )),
	(( 'ActiveThemeDisplayName' , 'prop' , ), 541, (541, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 924 , (3, 0, None, None) , 0 , )),
	(( 'Email' , 'prop' , ), 319, (319, (), [ (16393, 10, None, "IID('{000209DD-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 928 , (3, 0, None, None) , 0 , )),
	(( 'Scripts' , 'prop' , ), 320, (320, (), [ (16393, 10, None, "IID('{000C0340-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 932 , (3, 0, None, None) , 0 , )),
	(( 'LanguageDetected' , 'prop' , ), 321, (321, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 936 , (3, 0, None, None) , 0 , )),
	(( 'LanguageDetected' , 'prop' , ), 321, (321, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 940 , (3, 0, None, None) , 0 , )),
	(( 'FarEastLineBreakLanguage' , 'prop' , ), 326, (326, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 944 , (3, 0, None, None) , 0 , )),
	(( 'FarEastLineBreakLanguage' , 'prop' , ), 326, (326, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 948 , (3, 0, None, None) , 0 , )),
	(( 'Frameset' , 'prop' , ), 327, (327, (), [ (16393, 10, None, "IID('{000209E2-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 952 , (3, 0, None, None) , 0 , )),
	(( 'ClickAndTypeParagraphStyle' , 'prop' , ), 328, (328, (), [ (16396, 10, None, None) , ], 1 , 2 , 4 , 0 , 956 , (3, 0, None, None) , 1024 , )),
	(( 'ClickAndTypeParagraphStyle' , 'prop' , ), 328, (328, (), [ (16396, 1, None, None) , ], 1 , 4 , 4 , 0 , 960 , (3, 0, None, None) , 1024 , )),
	(( 'HTMLProject' , 'prop' , ), 329, (329, (), [ (16393, 10, None, "IID('{000C0356-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 964 , (3, 0, None, None) , 64 , )),
	(( 'WebOptions' , 'prop' , ), 330, (330, (), [ (16393, 10, None, "IID('{000209E4-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 968 , (3, 0, None, None) , 0 , )),
	(( 'OpenEncoding' , 'prop' , ), 332, (332, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 972 , (3, 0, None, None) , 0 , )),
	(( 'SaveEncoding' , 'prop' , ), 333, (333, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 976 , (3, 0, None, None) , 0 , )),
	(( 'SaveEncoding' , 'prop' , ), 333, (333, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 980 , (3, 0, None, None) , 0 , )),
	(( 'OptimizeForWord97' , 'prop' , ), 334, (334, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 984 , (3, 0, None, None) , 0 , )),
	(( 'OptimizeForWord97' , 'prop' , ), 334, (334, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 988 , (3, 0, None, None) , 0 , )),
	(( 'VBASigned' , 'prop' , ), 335, (335, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 992 , (3, 0, None, None) , 0 , )),
	(( 'PrintOut2000' , 'Background' , 'Append' , 'Range' , 'OutputFileName' , 
			 'From' , 'To' , 'Item' , 'Copies' , 'Pages' , 
			 'PageType' , 'PrintToFile' , 'Collate' , 'ActivePrinterMacGX' , 'ManualDuplexPrint' , 
			 'PrintZoomColumn' , 'PrintZoomRow' , 'PrintZoomPaperWidth' , 'PrintZoomPaperHeight' , ), 444, (444, (), [ 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , ], 1 , 1 , 4 , 18 , 996 , (3, 0, None, None) , 64 , )),
	(( 'sblt' , 's' , ), 445, (445, (), [ (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 1000 , (3, 0, None, None) , 64 , )),
	(( 'ConvertVietDoc' , 'CodePageOrigin' , ), 447, (447, (), [ (3, 1, None, None) , ], 1 , 1 , 4 , 0 , 1004 , (3, 0, None, None) , 0 , )),
	(( 'PrintOut' , 'Background' , 'Append' , 'Range' , 'OutputFileName' , 
			 'From' , 'To' , 'Item' , 'Copies' , 'Pages' , 
			 'PageType' , 'PrintToFile' , 'Collate' , 'ActivePrinterMacGX' , 'ManualDuplexPrint' , 
			 'PrintZoomColumn' , 'PrintZoomRow' , 'PrintZoomPaperWidth' , 'PrintZoomPaperHeight' , ), 446, (446, (), [ 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , ], 1 , 1 , 4 , 18 , 1008 , (3, 0, None, None) , 0 , )),
	(( 'MailEnvelope' , 'prop' , ), 336, (336, (), [ (16397, 10, None, "IID('{0006F01A-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 1012 , (3, 0, None, None) , 0 , )),
	(( 'DisableFeatures' , 'prop' , ), 337, (337, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1016 , (3, 0, None, None) , 0 , )),
	(( 'DisableFeatures' , 'prop' , ), 337, (337, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1020 , (3, 0, None, None) , 0 , )),
	(( 'DoNotEmbedSystemFonts' , 'prop' , ), 338, (338, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1024 , (3, 0, None, None) , 0 , )),
	(( 'DoNotEmbedSystemFonts' , 'prop' , ), 338, (338, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1028 , (3, 0, None, None) , 0 , )),
	(( 'Signatures' , 'prop' , ), 339, (339, (), [ (16393, 10, None, "IID('{000C0410-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 1032 , (3, 0, None, None) , 0 , )),
	(( 'DefaultTargetFrame' , 'prop' , ), 340, (340, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 1036 , (3, 0, None, None) , 0 , )),
	(( 'DefaultTargetFrame' , 'prop' , ), 340, (340, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 1040 , (3, 0, None, None) , 0 , )),
	(( 'HTMLDivisions' , 'prop' , ), 342, (342, (), [ (16393, 10, None, "IID('{000209E8-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 1044 , (3, 0, None, None) , 0 , )),
	(( 'DisableFeaturesIntroducedAfter' , 'prop' , ), 343, (343, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 1048 , (3, 0, None, None) , 0 , )),
	(( 'DisableFeaturesIntroducedAfter' , 'prop' , ), 343, (343, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 1052 , (3, 0, None, None) , 0 , )),
	(( 'RemovePersonalInformation' , 'prop' , ), 344, (344, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1056 , (3, 0, None, None) , 0 , )),
	(( 'RemovePersonalInformation' , 'prop' , ), 344, (344, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1060 , (3, 0, None, None) , 0 , )),
	(( 'SmartTags' , 'prop' , ), 346, (346, (), [ (16393, 10, None, "IID('{000209EE-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 1064 , (3, 0, None, None) , 64 , )),
	(( 'Compare2002' , 'Name' , 'AuthorName' , 'CompareTarget' , 'DetectFormatChanges' , 
			 'IgnoreAllComparisonWarnings' , 'AddToRecentFiles' , ), 345, (345, (), [ (8, 1, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , ], 1 , 1 , 4 , 5 , 1068 , (3, 0, None, None) , 64 , )),
	(( 'CheckIn' , 'SaveChanges' , 'Comments' , 'MakePublic' , ), 349, (349, (), [ 
			 (11, 49, 'True', None) , (16396, 17, None, None) , (11, 49, 'False', None) , ], 1 , 1 , 4 , 0 , 1072 , (3, 0, None, None) , 0 , )),
	(( 'CanCheckin' , 'prop' , ), 351, (351, (), [ (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 1076 , (3, 0, None, None) , 0 , )),
	(( 'Merge' , 'FileName' , 'MergeTarget' , 'DetectFormatChanges' , 'UseFormattingFrom' , 
			 'AddToRecentFiles' , ), 362, (362, (), [ (8, 1, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , ], 1 , 1 , 4 , 4 , 1080 , (3, 0, None, None) , 0 , )),
	(( 'EmbedSmartTags' , 'prop' , ), 347, (347, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1084 , (3, 0, None, None) , 64 , )),
	(( 'EmbedSmartTags' , 'prop' , ), 347, (347, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1088 , (3, 0, None, None) , 64 , )),
	(( 'SmartTagsAsXMLProps' , 'prop' , ), 348, (348, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1092 , (3, 0, None, None) , 64 , )),
	(( 'SmartTagsAsXMLProps' , 'prop' , ), 348, (348, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1096 , (3, 0, None, None) , 64 , )),
	(( 'TextEncoding' , 'prop' , ), 357, (357, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 1100 , (3, 0, None, None) , 0 , )),
	(( 'TextEncoding' , 'prop' , ), 357, (357, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 1104 , (3, 0, None, None) , 0 , )),
	(( 'TextLineEnding' , 'prop' , ), 358, (358, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 1108 , (3, 0, None, None) , 0 , )),
	(( 'TextLineEnding' , 'prop' , ), 358, (358, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 1112 , (3, 0, None, None) , 0 , )),
	(( 'SendForReview' , 'Recipients' , 'Subject' , 'ShowMessage' , 'IncludeAttachment' , 
			 ), 353, (353, (), [ (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , ], 1 , 1 , 4 , 4 , 1116 , (3, 0, None, None) , 0 , )),
	(( 'ReplyWithChanges' , 'ShowMessage' , ), 354, (354, (), [ (16396, 17, None, None) , ], 1 , 1 , 4 , 1 , 1120 , (3, 0, None, None) , 0 , )),
	(( 'EndReview' , ), 356, (356, (), [ ], 1 , 1 , 4 , 0 , 1124 , (3, 0, None, None) , 0 , )),
	(( 'StyleSheets' , 'prop' , ), 360, (360, (), [ (16393, 10, None, "IID('{07B7CC7E-E66C-11D3-9454-00105AA31A08}')") , ], 1 , 2 , 4 , 0 , 1128 , (3, 0, None, None) , 0 , )),
	(( 'DefaultTableStyle' , 'prop' , ), 365, (365, (), [ (16396, 10, None, None) , ], 1 , 2 , 4 , 0 , 1132 , (3, 0, None, None) , 0 , )),
	(( 'PasswordEncryptionProvider' , 'prop' , ), 367, (367, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 1136 , (3, 0, None, None) , 0 , )),
	(( 'PasswordEncryptionAlgorithm' , 'prop' , ), 368, (368, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 1140 , (3, 0, None, None) , 0 , )),
	(( 'PasswordEncryptionKeyLength' , 'prop' , ), 369, (369, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 1144 , (3, 0, None, None) , 0 , )),
	(( 'PasswordEncryptionFileProperties' , 'prop' , ), 370, (370, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1148 , (3, 0, None, None) , 0 , )),
	(( 'SetPasswordEncryptionOptions' , 'PasswordEncryptionProvider' , 'PasswordEncryptionAlgorithm' , 'PasswordEncryptionKeyLength' , 'PasswordEncryptionFileProperties' , 
			 ), 361, (361, (), [ (8, 1, None, None) , (8, 1, None, None) , (3, 1, None, None) , (16396, 17, None, None) , ], 1 , 1 , 4 , 1 , 1152 , (3, 0, None, None) , 0 , )),
	(( 'RecheckSmartTags' , ), 363, (363, (), [ ], 1 , 1 , 4 , 0 , 1156 , (3, 0, None, None) , 64 , )),
	(( 'RemoveSmartTags' , ), 364, (364, (), [ ], 1 , 1 , 4 , 0 , 1160 , (3, 0, None, None) , 64 , )),
	(( 'SetDefaultTableStyle' , 'Style' , 'SetInTemplate' , ), 366, (366, (), [ (16396, 1, None, None) , 
			 (11, 1, None, None) , ], 1 , 1 , 4 , 0 , 1164 , (3, 0, None, None) , 0 , )),
	(( 'DeleteAllComments' , ), 371, (371, (), [ ], 1 , 1 , 4 , 0 , 1168 , (3, 0, None, None) , 0 , )),
	(( 'AcceptAllRevisionsShown' , ), 372, (372, (), [ ], 1 , 1 , 4 , 0 , 1172 , (3, 0, None, None) , 0 , )),
	(( 'RejectAllRevisionsShown' , ), 373, (373, (), [ ], 1 , 1 , 4 , 0 , 1176 , (3, 0, None, None) , 0 , )),
	(( 'DeleteAllCommentsShown' , ), 374, (374, (), [ ], 1 , 1 , 4 , 0 , 1180 , (3, 0, None, None) , 0 , )),
	(( 'ResetFormFields' , ), 375, (375, (), [ ], 1 , 1 , 4 , 0 , 1184 , (3, 0, None, None) , 0 , )),
	(( 'SaveAs' , 'FileName' , 'FileFormat' , 'LockComments' , 'Password' , 
			 'AddToRecentFiles' , 'WritePassword' , 'ReadOnlyRecommended' , 'EmbedTrueTypeFonts' , 'SaveNativePictureFormat' , 
			 'SaveFormsData' , 'SaveAsAOCELetter' , 'Encoding' , 'InsertLineBreaks' , 'AllowSubstitutions' , 
			 'LineEnding' , 'AddBiDiMarks' , ), 376, (376, (), [ (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , ], 1 , 1 , 4 , 16 , 1188 , (3, 0, None, None) , 64 , )),
	(( 'EmbedLinguisticData' , 'prop' , ), 377, (377, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1192 , (3, 0, None, None) , 0 , )),
	(( 'EmbedLinguisticData' , 'prop' , ), 377, (377, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1196 , (3, 0, None, None) , 0 , )),
	(( 'FormattingShowFont' , 'prop' , ), 448, (448, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1200 , (3, 0, None, None) , 0 , )),
	(( 'FormattingShowFont' , 'prop' , ), 448, (448, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1204 , (3, 0, None, None) , 0 , )),
	(( 'FormattingShowClear' , 'prop' , ), 449, (449, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1208 , (3, 0, None, None) , 0 , )),
	(( 'FormattingShowClear' , 'prop' , ), 449, (449, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1212 , (3, 0, None, None) , 0 , )),
	(( 'FormattingShowParagraph' , 'prop' , ), 450, (450, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1216 , (3, 0, None, None) , 0 , )),
	(( 'FormattingShowParagraph' , 'prop' , ), 450, (450, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1220 , (3, 0, None, None) , 0 , )),
	(( 'FormattingShowNumbering' , 'prop' , ), 451, (451, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1224 , (3, 0, None, None) , 0 , )),
	(( 'FormattingShowNumbering' , 'prop' , ), 451, (451, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1228 , (3, 0, None, None) , 0 , )),
	(( 'FormattingShowFilter' , 'prop' , ), 452, (452, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 1232 , (3, 0, None, None) , 0 , )),
	(( 'FormattingShowFilter' , 'prop' , ), 452, (452, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 1236 , (3, 0, None, None) , 0 , )),
	(( 'CheckNewSmartTags' , ), 378, (378, (), [ ], 1 , 1 , 4 , 0 , 1240 , (3, 0, None, None) , 64 , )),
	(( 'Permission' , 'prop' , ), 453, (453, (), [ (16393, 10, None, "IID('{000C0376-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 1244 , (3, 0, None, None) , 0 , )),
	(( 'XMLNodes' , 'prop' , ), 460, (460, (), [ (16393, 10, None, "IID('{D36C1F42-7044-4B9E-9CA3-85919454DB04}')") , ], 1 , 2 , 4 , 0 , 1248 , (3, 0, None, None) , 0 , )),
	(( 'XMLSchemaReferences' , 'prop' , ), 461, (461, (), [ (16393, 10, None, "IID('{356B06EC-4908-42A4-81FC-4B5A51F3483B}')") , ], 1 , 2 , 4 , 0 , 1252 , (3, 0, None, None) , 0 , )),
	(( 'SmartDocument' , 'prop' , ), 462, (462, (), [ (16393, 10, None, "IID('{000C0377-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 1256 , (3, 0, None, None) , 0 , )),
	(( 'SharedWorkspace' , 'prop' , ), 463, (463, (), [ (16393, 10, None, "IID('{000C0385-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 1260 , (3, 0, None, None) , 64 , )),
	(( 'Sync' , 'prop' , ), 466, (466, (), [ (16393, 10, None, "IID('{000C0386-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 1264 , (3, 0, None, None) , 0 , )),
	(( 'EnforceStyle' , 'prop' , ), 471, (471, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1268 , (3, 0, None, None) , 0 , )),
	(( 'EnforceStyle' , 'prop' , ), 471, (471, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1272 , (3, 0, None, None) , 0 , )),
	(( 'AutoFormatOverride' , 'prop' , ), 472, (472, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1276 , (3, 0, None, None) , 0 , )),
	(( 'AutoFormatOverride' , 'prop' , ), 472, (472, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1280 , (3, 0, None, None) , 0 , )),
	(( 'XMLSaveDataOnly' , 'prop' , ), 473, (473, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1284 , (3, 0, None, None) , 0 , )),
	(( 'XMLSaveDataOnly' , 'prop' , ), 473, (473, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1288 , (3, 0, None, None) , 0 , )),
	(( 'XMLHideNamespaces' , 'prop' , ), 477, (477, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1292 , (3, 0, None, None) , 0 , )),
	(( 'XMLHideNamespaces' , 'prop' , ), 477, (477, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1296 , (3, 0, None, None) , 0 , )),
	(( 'XMLShowAdvancedErrors' , 'prop' , ), 478, (478, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1300 , (3, 0, None, None) , 0 , )),
	(( 'XMLShowAdvancedErrors' , 'prop' , ), 478, (478, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1304 , (3, 0, None, None) , 0 , )),
	(( 'XMLUseXSLTWhenSaving' , 'prop' , ), 474, (474, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1308 , (3, 0, None, None) , 0 , )),
	(( 'XMLUseXSLTWhenSaving' , 'prop' , ), 474, (474, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1312 , (3, 0, None, None) , 0 , )),
	(( 'XMLSaveThroughXSLT' , 'prop' , ), 475, (475, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 1316 , (3, 0, None, None) , 0 , )),
	(( 'XMLSaveThroughXSLT' , 'prop' , ), 475, (475, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 1320 , (3, 0, None, None) , 0 , )),
	(( 'DocumentLibraryVersions' , 'prop' , ), 476, (476, (), [ (16393, 10, None, "IID('{000C0388-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 1324 , (3, 0, None, None) , 0 , )),
	(( 'ReadingModeLayoutFrozen' , 'prop' , ), 481, (481, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1328 , (3, 0, None, None) , 0 , )),
	(( 'ReadingModeLayoutFrozen' , 'prop' , ), 481, (481, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1332 , (3, 0, None, None) , 0 , )),
	(( 'RemoveDateAndTime' , 'prop' , ), 484, (484, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1336 , (3, 0, None, None) , 0 , )),
	(( 'RemoveDateAndTime' , 'prop' , ), 484, (484, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1340 , (3, 0, None, None) , 0 , )),
	(( 'SendFaxOverInternet' , 'Recipients' , 'Subject' , 'ShowMessage' , ), 464, (464, (), [ 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , ], 1 , 1 , 4 , 3 , 1344 , (3, 0, None, None) , 0 , )),
	(( 'TransformDocument' , 'Path' , 'DataOnly' , ), 500, (500, (), [ (8, 1, None, None) , 
			 (11, 49, 'True', None) , ], 1 , 1 , 4 , 0 , 1348 , (3, 0, None, None) , 0 , )),
	(( 'Protect' , 'Type' , 'NoReset' , 'Password' , 'UseIRM' , 
			 'EnforceStyleLock' , ), 467, (467, (), [ (3, 1, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , ], 1 , 1 , 4 , 4 , 1352 , (3, 0, None, None) , 0 , )),
	(( 'SelectAllEditableRanges' , 'EditorID' , ), 468, (468, (), [ (16396, 17, None, None) , ], 1 , 1 , 4 , 1 , 1356 , (3, 0, None, None) , 0 , )),
	(( 'DeleteAllEditableRanges' , 'EditorID' , ), 469, (469, (), [ (16396, 17, None, None) , ], 1 , 1 , 4 , 1 , 1360 , (3, 0, None, None) , 0 , )),
	(( 'DeleteAllInkAnnotations' , ), 479, (479, (), [ ], 1 , 1 , 4 , 0 , 1364 , (3, 0, None, None) , 0 , )),
	(( 'AddDocumentWorkspaceHeader' , 'RichFormat' , 'Url' , 'Title' , 'Description' , 
			 'ID' , ), 482, (482, (), [ (11, 1, None, None) , (8, 1, None, None) , (8, 1, None, None) , 
			 (8, 1, None, None) , (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 1368 , (3, 0, None, None) , 64 , )),
	(( 'RemoveDocumentWorkspaceHeader' , 'ID' , ), 483, (483, (), [ (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 1372 , (3, 0, None, None) , 64 , )),
	(( 'Compare' , 'Name' , 'AuthorName' , 'CompareTarget' , 'DetectFormatChanges' , 
			 'IgnoreAllComparisonWarnings' , 'AddToRecentFiles' , 'RemovePersonalInformation' , 'RemoveDateAndTime' , ), 485, (485, (), [ 
			 (8, 1, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , ], 1 , 1 , 4 , 7 , 1376 , (3, 0, None, None) , 0 , )),
	(( 'RemoveLockedStyles' , ), 487, (487, (), [ ], 1 , 1 , 4 , 0 , 1380 , (3, 0, None, None) , 0 , )),
	(( 'ChildNodeSuggestions' , 'prop' , ), 486, (486, (), [ (16393, 10, None, "IID('{DE63B5AC-CA4F-46FE-9184-A5719AB9ED5B}')") , ], 1 , 2 , 4 , 0 , 1384 , (3, 0, None, None) , 0 , )),
	(( 'SelectSingleNode' , 'XPath' , 'PrefixMapping' , 'FastSearchSkippingTextNodes' , 'prop' , 
			 ), 488, (488, (), [ (8, 1, None, None) , (8, 49, "''", None) , (11, 49, 'True', None) , (16393, 10, None, "IID('{09760240-0B89-49F7-A79D-479F24723F56}')") , ], 1 , 1 , 4 , 0 , 1388 , (3, 32, None, None) , 0 , )),
	(( 'SelectNodes' , 'XPath' , 'PrefixMapping' , 'FastSearchSkippingTextNodes' , 'prop' , 
			 ), 489, (489, (), [ (8, 1, None, None) , (8, 49, "''", None) , (11, 49, 'True', None) , (16393, 10, None, "IID('{D36C1F42-7044-4B9E-9CA3-85919454DB04}')") , ], 1 , 1 , 4 , 0 , 1392 , (3, 32, None, None) , 0 , )),
	(( 'XMLSchemaViolations' , 'prop' , ), 490, (490, (), [ (16393, 10, None, "IID('{D36C1F42-7044-4B9E-9CA3-85919454DB04}')") , ], 1 , 2 , 4 , 0 , 1396 , (3, 0, None, None) , 0 , )),
	(( 'ReadingLayoutSizeX' , 'prop' , ), 491, (491, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 1400 , (3, 0, None, None) , 0 , )),
	(( 'ReadingLayoutSizeX' , 'prop' , ), 491, (491, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 1404 , (3, 0, None, None) , 0 , )),
	(( 'ReadingLayoutSizeY' , 'prop' , ), 492, (492, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 1408 , (3, 0, None, None) , 0 , )),
	(( 'ReadingLayoutSizeY' , 'prop' , ), 492, (492, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 1412 , (3, 0, None, None) , 0 , )),
	(( 'StyleSortMethod' , 'prop' , ), 493, (493, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 1416 , (3, 0, None, None) , 0 , )),
	(( 'StyleSortMethod' , 'prop' , ), 493, (493, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 1420 , (3, 0, None, None) , 0 , )),
	(( 'ContentTypeProperties' , 'prop' , ), 496, (496, (), [ (16393, 10, None, "IID('{000C038E-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 1424 , (3, 0, None, None) , 0 , )),
	(( 'TrackMoves' , 'prop' , ), 499, (499, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1428 , (3, 0, None, None) , 0 , )),
	(( 'TrackMoves' , 'prop' , ), 499, (499, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1432 , (3, 0, None, None) , 0 , )),
	(( 'TrackFormatting' , 'prop' , ), 502, (502, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1436 , (3, 0, None, None) , 0 , )),
	(( 'TrackFormatting' , 'prop' , ), 502, (502, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1440 , (3, 0, None, None) , 0 , )),
	(( 'Dummy1' , ), 503, (503, (), [ ], 1 , 2 , 4 , 0 , 1444 , (3, 0, None, None) , 64 , )),
	(( 'OMaths' , 'prop' , ), 504, (504, (), [ (16393, 10, None, "IID('{873E774B-926A-4CB1-878D-635A45187595}')") , ], 1 , 2 , 4 , 0 , 1448 , (3, 0, None, None) , 0 , )),
	(( 'RemoveDocumentInformation' , 'RemoveDocInfoType' , ), 495, (495, (), [ (3, 1, None, None) , ], 1 , 1 , 4 , 0 , 1452 , (3, 0, None, None) , 0 , )),
	(( 'CheckInWithVersion' , 'SaveChanges' , 'Comments' , 'MakePublic' , 'VersionType' , 
			 ), 501, (501, (), [ (11, 49, 'True', None) , (16396, 17, None, None) , (11, 49, 'False', None) , (16396, 17, None, None) , ], 1 , 1 , 4 , 1 , 1456 , (3, 0, None, None) , 0 , )),
	(( 'Dummy2' , ), 505, (505, (), [ ], 1 , 1 , 4 , 0 , 1460 , (3, 0, None, None) , 64 , )),
	(( 'Dummy3' , ), 506, (506, (), [ ], 1 , 2 , 4 , 0 , 1464 , (3, 0, None, None) , 64 , )),
	(( 'ServerPolicy' , 'prop' , ), 507, (507, (), [ (16393, 10, None, "IID('{000C0390-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 1468 , (3, 0, None, None) , 0 , )),
	(( 'ContentControls' , 'prop' , ), 508, (508, (), [ (16393, 10, None, "IID('{804CD967-F83B-432D-9446-C61A45CFEFF0}')") , ], 1 , 2 , 4 , 0 , 1472 , (3, 0, None, None) , 0 , )),
	(( 'DocumentInspectors' , 'prop' , ), 510, (510, (), [ (16393, 10, None, "IID('{000C0392-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 1476 , (3, 0, None, None) , 0 , )),
	(( 'LockServerFile' , ), 509, (509, (), [ ], 1 , 1 , 4 , 0 , 1480 , (3, 0, None, None) , 0 , )),
	(( 'GetWorkflowTasks' , 'prop' , ), 511, (511, (), [ (16393, 10, None, "IID('{000CD901-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 0 , 1484 , (3, 0, None, None) , 0 , )),
	(( 'GetWorkflowTemplates' , 'prop' , ), 512, (512, (), [ (16393, 10, None, "IID('{000CD903-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 0 , 1488 , (3, 0, None, None) , 0 , )),
	(( 'Dummy4' , ), 514, (514, (), [ ], 1 , 1 , 4 , 0 , 1492 , (3, 0, None, None) , 64 , )),
	(( 'AddMeetingWorkspaceHeader' , 'SkipIfAbsent' , 'Url' , 'Title' , 'Description' , 
			 'ID' , ), 515, (515, (), [ (11, 1, None, None) , (8, 1, None, None) , (8, 1, None, None) , 
			 (8, 1, None, None) , (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 1496 , (3, 0, None, None) , 64 , )),
	(( 'Bibliography' , 'prop' , ), 516, (516, (), [ (16393, 10, None, "IID('{3834F60F-EE8C-455D-A441-D766675D6D3B}')") , ], 1 , 2 , 4 , 0 , 1500 , (3, 0, None, None) , 0 , )),
	(( 'LockTheme' , 'prop' , ), 517, (517, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1504 , (3, 0, None, None) , 0 , )),
	(( 'LockTheme' , 'prop' , ), 517, (517, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1508 , (3, 0, None, None) , 0 , )),
	(( 'LockQuickStyleSet' , 'prop' , ), 518, (518, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1512 , (3, 0, None, None) , 0 , )),
	(( 'LockQuickStyleSet' , 'prop' , ), 518, (518, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1516 , (3, 0, None, None) , 0 , )),
	(( 'OriginalDocumentTitle' , 'prop' , ), 519, (519, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 1520 , (3, 0, None, None) , 0 , )),
	(( 'RevisedDocumentTitle' , 'prop' , ), 520, (520, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 1524 , (3, 0, None, None) , 0 , )),
	(( 'CustomXMLParts' , 'prop' , ), 521, (521, (), [ (16397, 10, None, "IID('{000CDB0C-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 1528 , (3, 0, None, None) , 0 , )),
	(( 'FormattingShowNextLevel' , 'prop' , ), 522, (522, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1532 , (3, 0, None, None) , 0 , )),
	(( 'FormattingShowNextLevel' , 'prop' , ), 522, (522, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1536 , (3, 0, None, None) , 0 , )),
	(( 'FormattingShowUserStyleName' , 'prop' , ), 523, (523, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1540 , (3, 0, None, None) , 0 , )),
	(( 'FormattingShowUserStyleName' , 'prop' , ), 523, (523, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1544 , (3, 0, None, None) , 0 , )),
	(( 'SaveAsQuickStyleSet' , 'FileName' , ), 524, (524, (), [ (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 1548 , (3, 0, None, None) , 0 , )),
	(( 'ApplyQuickStyleSet' , 'Name' , ), 525, (525, (), [ (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 1552 , (3, 0, None, None) , 64 , )),
	(( 'Research' , 'prop' , ), 526, (526, (), [ (16393, 10, None, "IID('{E6AAEC05-E543-4085-BA92-9BF7D2474F51}')") , ], 1 , 2 , 4 , 0 , 1556 , (3, 0, None, None) , 0 , )),
	(( 'Final' , 'prop' , ), 527, (527, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1560 , (3, 0, None, None) , 0 , )),
	(( 'Final' , 'prop' , ), 527, (527, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1564 , (3, 0, None, None) , 0 , )),
	(( 'OMathBreakBin' , 'prop' , ), 528, (528, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 1568 , (3, 0, None, None) , 0 , )),
	(( 'OMathBreakBin' , 'prop' , ), 528, (528, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 1572 , (3, 0, None, None) , 0 , )),
	(( 'OMathBreakSub' , 'prop' , ), 529, (529, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 1576 , (3, 0, None, None) , 0 , )),
	(( 'OMathBreakSub' , 'prop' , ), 529, (529, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 1580 , (3, 0, None, None) , 0 , )),
	(( 'OMathJc' , 'prop' , ), 530, (530, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 1584 , (3, 0, None, None) , 0 , )),
	(( 'OMathJc' , 'prop' , ), 530, (530, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 1588 , (3, 0, None, None) , 0 , )),
	(( 'OMathLeftMargin' , 'prop' , ), 531, (531, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 1592 , (3, 0, None, None) , 0 , )),
	(( 'OMathLeftMargin' , 'prop' , ), 531, (531, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 1596 , (3, 0, None, None) , 0 , )),
	(( 'OMathRightMargin' , 'prop' , ), 532, (532, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 1600 , (3, 0, None, None) , 0 , )),
	(( 'OMathRightMargin' , 'prop' , ), 532, (532, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 1604 , (3, 0, None, None) , 0 , )),
	(( 'OMathWrap' , 'prop' , ), 535, (535, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 1608 , (3, 0, None, None) , 0 , )),
	(( 'OMathWrap' , 'prop' , ), 535, (535, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 1612 , (3, 0, None, None) , 0 , )),
	(( 'OMathIntSubSupLim' , 'prop' , ), 536, (536, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1616 , (3, 0, None, None) , 0 , )),
	(( 'OMathIntSubSupLim' , 'prop' , ), 536, (536, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1620 , (3, 0, None, None) , 0 , )),
	(( 'OMathNarySupSubLim' , 'prop' , ), 537, (537, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1624 , (3, 0, None, None) , 0 , )),
	(( 'OMathNarySupSubLim' , 'prop' , ), 537, (537, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1628 , (3, 0, None, None) , 0 , )),
	(( 'OMathSmallFrac' , 'prop' , ), 539, (539, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1632 , (3, 0, None, None) , 0 , )),
	(( 'OMathSmallFrac' , 'prop' , ), 539, (539, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1636 , (3, 0, None, None) , 0 , )),
	(( 'WordOpenXML' , 'prop' , ), 542, (542, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 1640 , (3, 0, None, None) , 0 , )),
	(( 'DocumentTheme' , 'prop' , ), 545, (545, (), [ (16393, 10, None, "IID('{000C03A0-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 1644 , (3, 0, None, None) , 0 , )),
	(( 'ApplyDocumentTheme' , 'FileName' , ), 546, (546, (), [ (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 1648 , (3, 0, None, None) , 0 , )),
	(( 'HasVBProject' , 'prop' , ), 548, (548, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1652 , (3, 0, None, None) , 0 , )),
	(( 'SelectLinkedControls' , 'Node' , 'prop' , ), 549, (549, (), [ (9, 1, None, "IID('{000CDB04-0000-0000-C000-000000000046}')") , 
			 (16393, 10, None, "IID('{804CD967-F83B-432D-9446-C61A45CFEFF0}')") , ], 1 , 1 , 4 , 0 , 1656 , (3, 0, None, None) , 0 , )),
	(( 'SelectUnlinkedControls' , 'Stream' , 'prop' , ), 550, (550, (), [ (13, 49, '0', "IID('{000CDB08-0000-0000-C000-000000000046}')") , 
			 (16393, 10, None, "IID('{804CD967-F83B-432D-9446-C61A45CFEFF0}')") , ], 1 , 1 , 4 , 0 , 1660 , (3, 0, None, None) , 0 , )),
	(( 'SelectContentControlsByTitle' , 'Title' , 'prop' , ), 551, (551, (), [ (8, 1, None, None) , 
			 (16393, 10, None, "IID('{804CD967-F83B-432D-9446-C61A45CFEFF0}')") , ], 1 , 1 , 4 , 0 , 1664 , (3, 0, None, None) , 0 , )),
	(( 'ExportAsFixedFormat' , 'OutputFileName' , 'ExportFormat' , 'OpenAfterExport' , 'OptimizeFor' , 
			 'Range' , 'From' , 'To' , 'Item' , 'IncludeDocProps' , 
			 'KeepIRM' , 'CreateBookmarks' , 'DocStructureTags' , 'BitmapMissingFonts' , 'UseISO19005_1' , 
			 'FixedFormatExtClassPtr' , ), 552, (552, (), [ (8, 1, None, None) , (3, 1, None, None) , (11, 49, 'False', None) , 
			 (3, 49, '0', None) , (3, 49, '0', None) , (3, 49, '1', None) , (3, 49, '1', None) , (3, 49, '0', None) , 
			 (11, 49, 'False', None) , (11, 49, 'True', None) , (3, 49, '0', None) , (11, 49, 'True', None) , (11, 49, 'True', None) , 
			 (11, 49, 'False', None) , (16396, 17, None, None) , ], 1 , 1 , 4 , 1 , 1668 , (3, 0, None, None) , 0 , )),
	(( 'FreezeLayout' , ), 553, (553, (), [ ], 1 , 1 , 4 , 0 , 1672 , (3, 0, None, None) , 0 , )),
	(( 'UnfreezeLayout' , ), 554, (554, (), [ ], 1 , 1 , 4 , 0 , 1676 , (3, 0, None, None) , 64 , )),
	(( 'OMathFontName' , 'prop' , ), 555, (555, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 1680 , (3, 0, None, None) , 0 , )),
	(( 'OMathFontName' , 'prop' , ), 555, (555, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 1684 , (3, 0, None, None) , 0 , )),
	(( 'DowngradeDocument' , ), 558, (558, (), [ ], 1 , 1 , 4 , 0 , 1688 , (3, 0, None, None) , 0 , )),
	(( 'EncryptionProvider' , 'prop' , ), 559, (559, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 1692 , (3, 0, None, None) , 0 , )),
	(( 'EncryptionProvider' , 'prop' , ), 559, (559, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 1696 , (3, 0, None, None) , 0 , )),
	(( 'UseMathDefaults' , 'prop' , ), 560, (560, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1700 , (3, 0, None, None) , 0 , )),
	(( 'UseMathDefaults' , 'prop' , ), 560, (560, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1704 , (3, 0, None, None) , 0 , )),
	(( 'CurrentRsid' , 'prop' , ), 563, (563, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 1708 , (3, 0, None, None) , 0 , )),
	(( 'Convert' , ), 561, (561, (), [ ], 1 , 1 , 4 , 0 , 1712 , (3, 0, None, None) , 0 , )),
	(( 'SelectContentControlsByTag' , 'Tag' , 'prop' , ), 562, (562, (), [ (8, 1, None, None) , 
			 (16393, 10, None, "IID('{804CD967-F83B-432D-9446-C61A45CFEFF0}')") , ], 1 , 1 , 4 , 0 , 1716 , (3, 0, None, None) , 0 , )),
	(( 'ConvertAutoHyphens' , ), 650, (650, (), [ ], 1 , 1 , 4 , 0 , 1720 , (3, 0, None, None) , 0 , )),
	(( 'DocID' , 'prop' , ), 564, (564, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 1724 , (3, 0, None, None) , 64 , )),
	(( 'ApplyQuickStyleSet2' , 'Style' , ), 566, (566, (), [ (16396, 1, None, None) , ], 1 , 1 , 4 , 0 , 1728 , (3, 0, None, None) , 0 , )),
	(( 'CompatibilityMode' , 'prop' , ), 567, (567, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 1732 , (3, 0, None, None) , 0 , )),
	(( 'SaveAs2' , 'FileName' , 'FileFormat' , 'LockComments' , 'Password' , 
			 'AddToRecentFiles' , 'WritePassword' , 'ReadOnlyRecommended' , 'EmbedTrueTypeFonts' , 'SaveNativePictureFormat' , 
			 'SaveFormsData' , 'SaveAsAOCELetter' , 'Encoding' , 'InsertLineBreaks' , 'AllowSubstitutions' , 
			 'LineEnding' , 'AddBiDiMarks' , 'CompatibilityMode' , ), 568, (568, (), [ (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , ], 1 , 1 , 4 , 17 , 1736 , (3, 0, None, None) , 0 , )),
	(( 'CoAuthoring' , 'prop' , ), 600, (600, (), [ (16393, 10, None, "IID('{65DF9F31-B1E3-4651-87E8-51D55F302161}')") , ], 1 , 2 , 4 , 0 , 1740 , (3, 0, None, None) , 0 , )),
	(( 'SetCompatibilityMode' , 'Mode' , ), 571, (571, (), [ (3, 1, None, None) , ], 1 , 1 , 4 , 0 , 1744 , (3, 0, None, None) , 0 , )),
]

win32com.client.CLSIDToClass.RegisterCLSID( "{0002096B-0000-0000-C000-000000000046}", _Document )
