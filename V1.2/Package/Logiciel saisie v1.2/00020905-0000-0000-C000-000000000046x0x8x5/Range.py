# -*- coding: mbcs -*-
# Created by makepy.py version 0.5.01
# By python version 3.3.1 (v3.3.1:d9893d13c628, Apr  6 2013, 20:25:12) [MSC v.1600 32 bit (Intel)]
# From type library '{00020905-0000-0000-C000-000000000046}'
# On Thu Mar  6 15:34:15 2014
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
class Range(DispatchBaseClass):
	CLSID = IID('{0002095E-0000-0000-C000-000000000046}')
	coclass_clsid = None

	def AutoFormat(self):
		return self._oleobj_.InvokeTypes(195, LCID, 1, (24, 0), (),)

	def Calculate(self):
		return self._oleobj_.InvokeTypes(172, LCID, 1, (4, 0), (),)

	def CheckGrammar(self):
		return self._oleobj_.InvokeTypes(204, LCID, 1, (24, 0), (),)

	def CheckSpelling(self, CustomDictionary=defaultNamedOptArg, IgnoreUppercase=defaultNamedOptArg, AlwaysSuggest=defaultNamedOptArg, CustomDictionary2=defaultNamedOptArg
			, CustomDictionary3=defaultNamedOptArg, CustomDictionary4=defaultNamedOptArg, CustomDictionary5=defaultNamedOptArg, CustomDictionary6=defaultNamedOptArg, CustomDictionary7=defaultNamedOptArg
			, CustomDictionary8=defaultNamedOptArg, CustomDictionary9=defaultNamedOptArg, CustomDictionary10=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(205, LCID, 1, (24, 0), ((16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),CustomDictionary
			, IgnoreUppercase, AlwaysSuggest, CustomDictionary2, CustomDictionary3, CustomDictionary4
			, CustomDictionary5, CustomDictionary6, CustomDictionary7, CustomDictionary8, CustomDictionary9
			, CustomDictionary10)

	def CheckSynonyms(self):
		return self._oleobj_.InvokeTypes(180, LCID, 1, (24, 0), (),)

	def Collapse(self, Direction=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(101, LCID, 1, (24, 0), ((16396, 17),),Direction
			)

	def ComputeStatistics(self, Statistic=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(178, LCID, 1, (3, 0), ((3, 1),),Statistic
			)

	def ConvertHangulAndHanja(self, ConversionsMode=defaultNamedOptArg, FastConversion=defaultNamedOptArg, CheckHangulEnding=defaultNamedOptArg, EnableRecentOrdering=defaultNamedOptArg
			, CustomDictionary=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(221, LCID, 1, (24, 0), ((16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),ConversionsMode
			, FastConversion, CheckHangulEnding, EnableRecentOrdering, CustomDictionary)

	# Result is of type Table
	def ConvertToTable(self, Separator=defaultNamedOptArg, NumRows=defaultNamedOptArg, NumColumns=defaultNamedOptArg, InitialColumnWidth=defaultNamedOptArg
			, Format=defaultNamedOptArg, ApplyBorders=defaultNamedOptArg, ApplyShading=defaultNamedOptArg, ApplyFont=defaultNamedOptArg, ApplyColor=defaultNamedOptArg
			, ApplyHeadingRows=defaultNamedOptArg, ApplyLastRow=defaultNamedOptArg, ApplyFirstColumn=defaultNamedOptArg, ApplyLastColumn=defaultNamedOptArg, AutoFit=defaultNamedOptArg
			, AutoFitBehavior=defaultNamedOptArg, DefaultTableBehavior=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(498, LCID, 1, (9, 0), ((16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),Separator
			, NumRows, NumColumns, InitialColumnWidth, Format, ApplyBorders
			, ApplyShading, ApplyFont, ApplyColor, ApplyHeadingRows, ApplyLastRow
			, ApplyFirstColumn, ApplyLastColumn, AutoFit, AutoFitBehavior, DefaultTableBehavior
			)
		if ret is not None:
			ret = Dispatch(ret, 'ConvertToTable', '{00020951-0000-0000-C000-000000000046}')
		return ret

	# Result is of type Table
	def ConvertToTableOld(self, Separator=defaultNamedOptArg, NumRows=defaultNamedOptArg, NumColumns=defaultNamedOptArg, InitialColumnWidth=defaultNamedOptArg
			, Format=defaultNamedOptArg, ApplyBorders=defaultNamedOptArg, ApplyShading=defaultNamedOptArg, ApplyFont=defaultNamedOptArg, ApplyColor=defaultNamedOptArg
			, ApplyHeadingRows=defaultNamedOptArg, ApplyLastRow=defaultNamedOptArg, ApplyFirstColumn=defaultNamedOptArg, ApplyLastColumn=defaultNamedOptArg, AutoFit=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(162, LCID, 1, (9, 0), ((16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),Separator
			, NumRows, NumColumns, InitialColumnWidth, Format, ApplyBorders
			, ApplyShading, ApplyFont, ApplyColor, ApplyHeadingRows, ApplyLastRow
			, ApplyFirstColumn, ApplyLastColumn, AutoFit)
		if ret is not None:
			ret = Dispatch(ret, 'ConvertToTableOld', '{00020951-0000-0000-C000-000000000046}')
		return ret

	def Copy(self):
		return self._oleobj_.InvokeTypes(120, LCID, 1, (24, 0), (),)

	def CopyAsPicture(self):
		return self._oleobj_.InvokeTypes(167, LCID, 1, (24, 0), (),)

	def CreatePublisher(self, Edition=defaultNamedOptArg, ContainsPICT=defaultNamedOptArg, ContainsRTF=defaultNamedOptArg, ContainsText=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(182, LCID, 1, (24, 0), ((16396, 17), (16396, 17), (16396, 17), (16396, 17)),Edition
			, ContainsPICT, ContainsRTF, ContainsText)

	def Cut(self):
		return self._oleobj_.InvokeTypes(119, LCID, 1, (24, 0), (),)

	def Delete(self, Unit=defaultNamedOptArg, Count=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(127, LCID, 1, (3, 0), ((16396, 17), (16396, 17)),Unit
			, Count)

	def DetectLanguage(self):
		return self._oleobj_.InvokeTypes(203, LCID, 1, (24, 0), (),)

	def EndOf(self, Unit=defaultNamedOptArg, Extend=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(108, LCID, 1, (3, 0), ((16396, 17), (16396, 17)),Unit
			, Extend)

	def Expand(self, Unit=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(129, LCID, 1, (3, 0), ((16396, 17),),Unit
			)

	def ExportAsFixedFormat(self, OutputFileName=defaultNamedNotOptArg, ExportFormat=defaultNamedNotOptArg, OpenAfterExport=False, OptimizeFor=0
			, ExportCurrentPage=False, Item=0, IncludeDocProps=False, KeepIRM=True, CreateBookmarks=0
			, DocStructureTags=True, BitmapMissingFonts=True, UseISO19005_1=False, FixedFormatExtClassPtr=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(503, LCID, 1, (24, 0), ((8, 1), (3, 1), (11, 49), (3, 49), (11, 49), (3, 49), (11, 49), (11, 49), (3, 49), (11, 49), (11, 49), (11, 49), (16396, 17)),OutputFileName
			, ExportFormat, OpenAfterExport, OptimizeFor, ExportCurrentPage, Item
			, IncludeDocProps, KeepIRM, CreateBookmarks, DocStructureTags, BitmapMissingFonts
			, UseISO19005_1, FixedFormatExtClassPtr)

	def ExportFragment(self, FileName=defaultNamedNotOptArg, Format=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(425, LCID, 1, (24, 0), ((8, 1), (3, 1)),FileName
			, Format)

	# Result is of type SpellingSuggestions
	def GetSpellingSuggestions(self, CustomDictionary=defaultNamedOptArg, IgnoreUppercase=defaultNamedOptArg, MainDictionary=defaultNamedOptArg, SuggestionMode=defaultNamedOptArg
			, CustomDictionary2=defaultNamedOptArg, CustomDictionary3=defaultNamedOptArg, CustomDictionary4=defaultNamedOptArg, CustomDictionary5=defaultNamedOptArg, CustomDictionary6=defaultNamedOptArg
			, CustomDictionary7=defaultNamedOptArg, CustomDictionary8=defaultNamedOptArg, CustomDictionary9=defaultNamedOptArg, CustomDictionary10=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(209, LCID, 1, (9, 0), ((16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),CustomDictionary
			, IgnoreUppercase, MainDictionary, SuggestionMode, CustomDictionary2, CustomDictionary3
			, CustomDictionary4, CustomDictionary5, CustomDictionary6, CustomDictionary7, CustomDictionary8
			, CustomDictionary9, CustomDictionary10)
		if ret is not None:
			ret = Dispatch(ret, 'GetSpellingSuggestions', '{000209AA-0000-0000-C000-000000000046}')
		return ret

	# The method GetXML is actually a property, but must be used as a method to correctly pass the arguments
	def GetXML(self, DataOnly=False):
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(344, LCID, 2, (8, 0), ((11, 49),),DataOnly
			)

	# Result is of type Range
	def GoTo(self, What=defaultNamedOptArg, Which=defaultNamedOptArg, Count=defaultNamedOptArg, Name=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(173, LCID, 1, (9, 0), ((16396, 17), (16396, 17), (16396, 17), (16396, 17)),What
			, Which, Count, Name)
		if ret is not None:
			ret = Dispatch(ret, 'GoTo', '{0002095E-0000-0000-C000-000000000046}')
		return ret

	# Result is of type Range
	def GoToEditableRange(self, EditorID=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(415, LCID, 1, (9, 0), ((16396, 17),),EditorID
			)
		if ret is not None:
			ret = Dispatch(ret, 'GoToEditableRange', '{0002095E-0000-0000-C000-000000000046}')
		return ret

	# Result is of type Range
	def GoToNext(self, What=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(174, LCID, 1, (9, 0), ((3, 1),),What
			)
		if ret is not None:
			ret = Dispatch(ret, 'GoToNext', '{0002095E-0000-0000-C000-000000000046}')
		return ret

	# Result is of type Range
	def GoToPrevious(self, What=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(175, LCID, 1, (9, 0), ((3, 1),),What
			)
		if ret is not None:
			ret = Dispatch(ret, 'GoToPrevious', '{0002095E-0000-0000-C000-000000000046}')
		return ret

	def ImportFragment(self, FileName=defaultNamedNotOptArg, MatchDestination=False):
		return self._oleobj_.InvokeTypes(502, LCID, 1, (24, 0), ((8, 1), (11, 49)),FileName
			, MatchDestination)

	def InRange(self, Range=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(126, LCID, 1, (11, 0), ((9, 1),),Range
			)

	def InStory(self, Range=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(125, LCID, 1, (11, 0), ((9, 1),),Range
			)

	# The method Information is actually a property, but must be used as a method to correctly pass the arguments
	def Information(self, Type=defaultNamedNotOptArg):
		return self._ApplyTypes_(313, 2, (12, 0), ((3, 1),), 'Information', None,Type
			)

	def InsertAfter(self, Text=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(104, LCID, 1, (24, 0), ((8, 1),),Text
			)

	def InsertAlignmentTab(self, Alignment=defaultNamedNotOptArg, RelativeTo=0):
		return self._oleobj_.InvokeTypes(500, LCID, 1, (24, 0), ((3, 1), (3, 49)),Alignment
			, RelativeTo)

	def InsertAutoText(self):
		return self._oleobj_.InvokeTypes(183, LCID, 1, (24, 0), (),)

	def InsertBefore(self, Text=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(102, LCID, 1, (24, 0), ((8, 1),),Text
			)

	def InsertBreak(self, Type=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(122, LCID, 1, (24, 0), ((16396, 17),),Type
			)

	def InsertCaption(self, Label=defaultNamedNotOptArg, Title=defaultNamedOptArg, TitleAutoText=defaultNamedOptArg, Position=defaultNamedOptArg
			, ExcludeLabel=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(417, LCID, 1, (24, 0), ((16396, 1), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),Label
			, Title, TitleAutoText, Position, ExcludeLabel)

	def InsertCaptionXP(self, Label=defaultNamedNotOptArg, Title=defaultNamedOptArg, TitleAutoText=defaultNamedOptArg, Position=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(166, LCID, 1, (24, 0), ((16396, 1), (16396, 17), (16396, 17), (16396, 17)),Label
			, Title, TitleAutoText, Position)

	def InsertCrossReference(self, ReferenceType=defaultNamedNotOptArg, ReferenceKind=defaultNamedNotOptArg, ReferenceItem=defaultNamedNotOptArg, InsertAsHyperlink=defaultNamedOptArg
			, IncludePosition=defaultNamedOptArg, SeparateNumbers=defaultNamedOptArg, SeparatorString=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(418, LCID, 1, (24, 0), ((16396, 1), (3, 1), (16396, 1), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),ReferenceType
			, ReferenceKind, ReferenceItem, InsertAsHyperlink, IncludePosition, SeparateNumbers
			, SeparatorString)

	def InsertCrossReference_2002(self, ReferenceType=defaultNamedNotOptArg, ReferenceKind=defaultNamedNotOptArg, ReferenceItem=defaultNamedNotOptArg, InsertAsHyperlink=defaultNamedOptArg
			, IncludePosition=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(165, LCID, 1, (24, 0), ((16396, 1), (3, 1), (16396, 1), (16396, 17), (16396, 17)),ReferenceType
			, ReferenceKind, ReferenceItem, InsertAsHyperlink, IncludePosition)

	def InsertDatabase(self, Format=defaultNamedOptArg, Style=defaultNamedOptArg, LinkToSource=defaultNamedOptArg, Connection=defaultNamedOptArg
			, SQLStatement=defaultNamedOptArg, SQLStatement1=defaultNamedOptArg, PasswordDocument=defaultNamedOptArg, PasswordTemplate=defaultNamedOptArg, WritePasswordDocument=defaultNamedOptArg
			, WritePasswordTemplate=defaultNamedOptArg, DataSource=defaultNamedOptArg, From=defaultNamedOptArg, To=defaultNamedOptArg, IncludeFields=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(194, LCID, 1, (24, 0), ((16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),Format
			, Style, LinkToSource, Connection, SQLStatement, SQLStatement1
			, PasswordDocument, PasswordTemplate, WritePasswordDocument, WritePasswordTemplate, DataSource
			, From, To, IncludeFields)

	def InsertDateTime(self, DateTimeFormat=defaultNamedOptArg, InsertAsField=defaultNamedOptArg, InsertAsFullWidth=defaultNamedOptArg, DateLanguage=defaultNamedOptArg
			, CalendarType=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(444, LCID, 1, (24, 0), ((16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),DateTimeFormat
			, InsertAsField, InsertAsFullWidth, DateLanguage, CalendarType)

	def InsertDateTimeOld(self, DateTimeFormat=defaultNamedOptArg, InsertAsField=defaultNamedOptArg, InsertAsFullWidth=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(163, LCID, 1, (24, 0), ((16396, 17), (16396, 17), (16396, 17)),DateTimeFormat
			, InsertAsField, InsertAsFullWidth)

	def InsertFile(self, FileName=defaultNamedNotOptArg, Range=defaultNamedOptArg, ConfirmConversions=defaultNamedOptArg, Link=defaultNamedOptArg
			, Attachment=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(123, LCID, 1, (24, 0), ((8, 1), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),FileName
			, Range, ConfirmConversions, Link, Attachment)

	def InsertParagraph(self):
		return self._oleobj_.InvokeTypes(160, LCID, 1, (24, 0), (),)

	def InsertParagraphAfter(self):
		return self._oleobj_.InvokeTypes(161, LCID, 1, (24, 0), (),)

	def InsertParagraphBefore(self):
		return self._oleobj_.InvokeTypes(212, LCID, 1, (24, 0), (),)

	def InsertSymbol(self, CharacterNumber=defaultNamedNotOptArg, Font=defaultNamedOptArg, Unicode=defaultNamedOptArg, Bias=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(164, LCID, 1, (24, 0), ((3, 1), (16396, 17), (16396, 17), (16396, 17)),CharacterNumber
			, Font, Unicode, Bias)

	def InsertXML(self, XML=defaultNamedNotOptArg, Transform=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(416, LCID, 1, (24, 0), ((8, 1), (16396, 17)),XML
			, Transform)

	def IsEqual(self, Range=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(171, LCID, 1, (11, 0), ((9, 1),),Range
			)

	def LookupNameProperties(self):
		return self._oleobj_.InvokeTypes(177, LCID, 1, (24, 0), (),)

	def ModifyEnclosure(self, Style=defaultNamedNotOptArg, Symbol=defaultNamedOptArg, EnclosedText=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(223, LCID, 1, (24, 0), ((16396, 1), (16396, 17), (16396, 17)),Style
			, Symbol, EnclosedText)

	def Move(self, Unit=defaultNamedOptArg, Count=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(109, LCID, 1, (3, 0), ((16396, 17), (16396, 17)),Unit
			, Count)

	def MoveEnd(self, Unit=defaultNamedOptArg, Count=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(111, LCID, 1, (3, 0), ((16396, 17), (16396, 17)),Unit
			, Count)

	def MoveEndUntil(self, Cset=defaultNamedNotOptArg, Count=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(117, LCID, 1, (3, 0), ((16396, 1), (16396, 17)),Cset
			, Count)

	def MoveEndWhile(self, Cset=defaultNamedNotOptArg, Count=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(114, LCID, 1, (3, 0), ((16396, 1), (16396, 17)),Cset
			, Count)

	def MoveStart(self, Unit=defaultNamedOptArg, Count=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(110, LCID, 1, (3, 0), ((16396, 17), (16396, 17)),Unit
			, Count)

	def MoveStartUntil(self, Cset=defaultNamedNotOptArg, Count=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(116, LCID, 1, (3, 0), ((16396, 1), (16396, 17)),Cset
			, Count)

	def MoveStartWhile(self, Cset=defaultNamedNotOptArg, Count=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(113, LCID, 1, (3, 0), ((16396, 1), (16396, 17)),Cset
			, Count)

	def MoveUntil(self, Cset=defaultNamedNotOptArg, Count=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(115, LCID, 1, (3, 0), ((16396, 1), (16396, 17)),Cset
			, Count)

	def MoveWhile(self, Cset=defaultNamedNotOptArg, Count=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(112, LCID, 1, (3, 0), ((16396, 1), (16396, 17)),Cset
			, Count)

	# Result is of type Range
	def Next(self, Unit=defaultNamedOptArg, Count=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(105, LCID, 1, (9, 0), ((16396, 17), (16396, 17)),Unit
			, Count)
		if ret is not None:
			ret = Dispatch(ret, 'Next', '{0002095E-0000-0000-C000-000000000046}')
		return ret

	def NextSubdocument(self):
		return self._oleobj_.InvokeTypes(219, LCID, 1, (24, 0), (),)

	def Paste(self):
		return self._oleobj_.InvokeTypes(121, LCID, 1, (24, 0), (),)

	def PasteAndFormat(self, Type=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(412, LCID, 1, (24, 0), ((3, 1),),Type
			)

	def PasteAppendTable(self):
		return self._oleobj_.InvokeTypes(414, LCID, 1, (24, 0), (),)

	def PasteAsNestedTable(self):
		return self._oleobj_.InvokeTypes(222, LCID, 1, (24, 0), (),)

	def PasteExcelTable(self, LinkedToExcel=defaultNamedNotOptArg, WordFormatting=defaultNamedNotOptArg, RTF=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(413, LCID, 1, (24, 0), ((11, 1), (11, 1), (11, 1)),LinkedToExcel
			, WordFormatting, RTF)

	def PasteSpecial(self, IconIndex=defaultNamedOptArg, Link=defaultNamedOptArg, Placement=defaultNamedOptArg, DisplayAsIcon=defaultNamedOptArg
			, DataType=defaultNamedOptArg, IconFileName=defaultNamedOptArg, IconLabel=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(176, LCID, 1, (24, 0), ((16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),IconIndex
			, Link, Placement, DisplayAsIcon, DataType, IconFileName
			, IconLabel)

	def PhoneticGuide(self, Text=defaultNamedNotOptArg, Alignment=-1, Raise=0, FontSize=0
			, FontName=''):
		return self._ApplyTypes_(224, 1, (24, 32), ((8, 1), (3, 49), (3, 49), (3, 49), (8, 49)), 'PhoneticGuide', None,Text
			, Alignment, Raise, FontSize, FontName)

	# Result is of type Range
	def Previous(self, Unit=defaultNamedOptArg, Count=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(106, LCID, 1, (9, 0), ((16396, 17), (16396, 17)),Unit
			, Count)
		if ret is not None:
			ret = Dispatch(ret, 'Previous', '{0002095E-0000-0000-C000-000000000046}')
		return ret

	def PreviousSubdocument(self):
		return self._oleobj_.InvokeTypes(220, LCID, 1, (24, 0), (),)

	def Relocate(self, Direction=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(179, LCID, 1, (24, 0), ((3, 1),),Direction
			)

	def Select(self):
		return self._oleobj_.InvokeTypes(65535, LCID, 1, (24, 0), (),)

	def SetListLevel(self, Level=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(427, LCID, 1, (24, 0), ((2, 1),),Level
			)

	def SetRange(self, Start=defaultNamedNotOptArg, End=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(100, LCID, 1, (24, 0), ((3, 1), (3, 1)),Start
			, End)

	def Sort(self, ExcludeHeader=defaultNamedOptArg, FieldNumber=defaultNamedOptArg, SortFieldType=defaultNamedOptArg, SortOrder=defaultNamedOptArg
			, FieldNumber2=defaultNamedOptArg, SortFieldType2=defaultNamedOptArg, SortOrder2=defaultNamedOptArg, FieldNumber3=defaultNamedOptArg, SortFieldType3=defaultNamedOptArg
			, SortOrder3=defaultNamedOptArg, SortColumn=defaultNamedOptArg, Separator=defaultNamedOptArg, CaseSensitive=defaultNamedOptArg, BidiSort=defaultNamedOptArg
			, IgnoreThe=defaultNamedOptArg, IgnoreKashida=defaultNamedOptArg, IgnoreDiacritics=defaultNamedOptArg, IgnoreHe=defaultNamedOptArg, LanguageID=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(484, LCID, 1, (24, 0), ((16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),ExcludeHeader
			, FieldNumber, SortFieldType, SortOrder, FieldNumber2, SortFieldType2
			, SortOrder2, FieldNumber3, SortFieldType3, SortOrder3, SortColumn
			, Separator, CaseSensitive, BidiSort, IgnoreThe, IgnoreKashida
			, IgnoreDiacritics, IgnoreHe, LanguageID)

	def SortAscending(self):
		return self._oleobj_.InvokeTypes(169, LCID, 1, (24, 0), (),)

	def SortDescending(self):
		return self._oleobj_.InvokeTypes(170, LCID, 1, (24, 0), (),)

	def SortOld(self, ExcludeHeader=defaultNamedOptArg, FieldNumber=defaultNamedOptArg, SortFieldType=defaultNamedOptArg, SortOrder=defaultNamedOptArg
			, FieldNumber2=defaultNamedOptArg, SortFieldType2=defaultNamedOptArg, SortOrder2=defaultNamedOptArg, FieldNumber3=defaultNamedOptArg, SortFieldType3=defaultNamedOptArg
			, SortOrder3=defaultNamedOptArg, SortColumn=defaultNamedOptArg, Separator=defaultNamedOptArg, CaseSensitive=defaultNamedOptArg, LanguageID=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(168, LCID, 1, (24, 0), ((16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),ExcludeHeader
			, FieldNumber, SortFieldType, SortOrder, FieldNumber2, SortFieldType2
			, SortOrder2, FieldNumber3, SortFieldType3, SortOrder3, SortColumn
			, Separator, CaseSensitive, LanguageID)

	def StartOf(self, Unit=defaultNamedOptArg, Extend=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(107, LCID, 1, (3, 0), ((16396, 17), (16396, 17)),Unit
			, Extend)

	def SubscribeTo(self, Edition=defaultNamedNotOptArg, Format=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(181, LCID, 1, (24, 0), ((8, 1), (16396, 17)),Edition
			, Format)

	def TCSCConverter(self, WdTCSCConverterDirection=2, CommonTerms=False, UseVariants=False):
		return self._oleobj_.InvokeTypes(499, LCID, 1, (24, 0), ((3, 49), (11, 49), (11, 49)),WdTCSCConverterDirection
			, CommonTerms, UseVariants)

	def WholeStory(self):
		return self._oleobj_.InvokeTypes(128, LCID, 1, (24, 0), (),)

	_prop_map_get_ = {
		# Method 'Application' returns object of type 'Application'
		"Application": (1000, 2, (13, 0), (), "Application", '{000209FF-0000-0000-C000-000000000046}'),
		"Bold": (130, 2, (3, 0), (), "Bold", None),
		"BoldBi": (400, 2, (3, 0), (), "BoldBi", None),
		"BookmarkID": (308, 2, (3, 0), (), "BookmarkID", None),
		# Method 'Bookmarks' returns object of type 'Bookmarks'
		"Bookmarks": (75, 2, (9, 0), (), "Bookmarks", '{00020967-0000-0000-C000-000000000046}'),
		# Method 'Borders' returns object of type 'Borders'
		"Borders": (1100, 2, (9, 0), (), "Borders", '{0002093C-0000-0000-C000-000000000046}'),
		"CanEdit": (304, 2, (3, 0), (), "CanEdit", None),
		"CanPaste": (305, 2, (3, 0), (), "CanPaste", None),
		"Case": (312, 2, (3, 0), (), "Case", None),
		# Method 'Cells' returns object of type 'Cells'
		"Cells": (57, 2, (9, 0), (), "Cells", '{0002094A-0000-0000-C000-000000000046}'),
		"CharacterStyle": (420, 2, (12, 0), (), "CharacterStyle", None),
		"CharacterWidth": (326, 2, (3, 0), (), "CharacterWidth", None),
		# Method 'Characters' returns object of type 'Characters'
		"Characters": (53, 2, (9, 0), (), "Characters", '{0002095D-0000-0000-C000-000000000046}'),
		# Method 'Columns' returns object of type 'Columns'
		"Columns": (302, 2, (9, 0), (), "Columns", '{0002094B-0000-0000-C000-000000000046}'),
		"CombineCharacters": (267, 2, (11, 0), (), "CombineCharacters", None),
		# Method 'Comments' returns object of type 'Comments'
		"Comments": (56, 2, (9, 0), (), "Comments", '{00020940-0000-0000-C000-000000000046}'),
		# Method 'Conflicts' returns object of type 'Conflicts'
		"Conflicts": (506, 2, (9, 0), (), "Conflicts", '{C2B83A65-B061-4469-83B6-8877437CB8A0}'),
		# Method 'ContentControls' returns object of type 'ContentControls'
		"ContentControls": (424, 2, (9, 0), (), "ContentControls", '{804CD967-F83B-432D-9446-C61A45CFEFF0}'),
		"Creator": (1001, 2, (3, 0), (), "Creator", None),
		"DisableCharacterSpaceGrid": (141, 2, (11, 0), (), "DisableCharacterSpaceGrid", None),
		# Method 'Document' returns object of type 'Document'
		"Document": (409, 2, (13, 0), (), "Document", '{00020906-0000-0000-C000-000000000046}'),
		# Method 'Duplicate' returns object of type 'Range'
		"Duplicate": (6, 2, (9, 0), (), "Duplicate", '{0002095E-0000-0000-C000-000000000046}'),
		# Method 'Editors' returns object of type 'Editors'
		"Editors": (343, 2, (9, 0), (), "Editors", '{AED7E08C-14F0-4F33-921D-4C5353137BF6}'),
		"EmphasisMark": (140, 2, (3, 0), (), "EmphasisMark", None),
		"End": (4, 2, (3, 0), (), "End", None),
		# Method 'EndnoteOptions' returns object of type 'EndnoteOptions'
		"EndnoteOptions": (411, 2, (9, 0), (), "EndnoteOptions", '{BF043168-F4DE-4E7C-B206-741A8B3EF71A}'),
		# Method 'Endnotes' returns object of type 'Endnotes'
		"Endnotes": (55, 2, (9, 0), (), "Endnotes", '{00020941-0000-0000-C000-000000000046}'),
		"EnhMetaFileBits": (345, 2, (12, 0), (), "EnhMetaFileBits", None),
		# Method 'Fields' returns object of type 'Fields'
		"Fields": (64, 2, (9, 0), (), "Fields", '{00020930-0000-0000-C000-000000000046}'),
		# Method 'Find' returns object of type 'Find'
		"Find": (262, 2, (9, 0), (), "Find", '{000209B0-0000-0000-C000-000000000046}'),
		"FitTextWidth": (264, 2, (4, 0), (), "FitTextWidth", None),
		# Method 'Font' returns object of type 'Font'
		"Font": (5, 2, (13, 0), (), "Font", '{000209F5-0000-0000-C000-000000000046}'),
		# Method 'FootnoteOptions' returns object of type 'FootnoteOptions'
		"FootnoteOptions": (410, 2, (9, 0), (), "FootnoteOptions", '{BEA85A24-D7DA-4F3D-B58C-ED90FB01D615}'),
		# Method 'Footnotes' returns object of type 'Footnotes'
		"Footnotes": (54, 2, (9, 0), (), "Footnotes", '{00020942-0000-0000-C000-000000000046}'),
		# Method 'FormFields' returns object of type 'FormFields'
		"FormFields": (65, 2, (9, 0), (), "FormFields", '{00020929-0000-0000-C000-000000000046}'),
		# Method 'FormattedText' returns object of type 'Range'
		"FormattedText": (2, 2, (9, 0), (), "FormattedText", '{0002095E-0000-0000-C000-000000000046}'),
		# Method 'Frames' returns object of type 'Frames'
		"Frames": (66, 2, (9, 0), (), "Frames", '{0002092B-0000-0000-C000-000000000046}'),
		"GrammarChecked": (260, 2, (11, 0), (), "GrammarChecked", None),
		# Method 'GrammaticalErrors' returns object of type 'ProofreadingErrors'
		"GrammaticalErrors": (315, 2, (9, 0), (), "GrammaticalErrors", '{000209BB-0000-0000-C000-000000000046}'),
		# Method 'HTMLDivisions' returns object of type 'HTMLDivisions'
		"HTMLDivisions": (406, 2, (9, 0), (), "HTMLDivisions", '{000209E8-0000-0000-C000-000000000046}'),
		"HighlightColorIndex": (301, 2, (3, 0), (), "HighlightColorIndex", None),
		"HorizontalInVertical": (265, 2, (3, 0), (), "HorizontalInVertical", None),
		# Method 'Hyperlinks' returns object of type 'Hyperlinks'
		"Hyperlinks": (156, 2, (9, 0), (), "Hyperlinks", '{0002099C-0000-0000-C000-000000000046}'),
		"ID": (405, 2, (8, 0), (), "ID", None),
		# Method 'InlineShapes' returns object of type 'InlineShapes'
		"InlineShapes": (319, 2, (9, 0), (), "InlineShapes", '{000209A9-0000-0000-C000-000000000046}'),
		"IsEndOfRowMark": (307, 2, (11, 0), (), "IsEndOfRowMark", None),
		"Italic": (131, 2, (3, 0), (), "Italic", None),
		"ItalicBi": (401, 2, (3, 0), (), "ItalicBi", None),
		"Kana": (327, 2, (3, 0), (), "Kana", None),
		"LanguageDetected": (263, 2, (11, 0), (), "LanguageDetected", None),
		"LanguageID": (153, 2, (3, 0), (), "LanguageID", None),
		"LanguageIDFarEast": (321, 2, (3, 0), (), "LanguageIDFarEast", None),
		"LanguageIDOther": (322, 2, (3, 0), (), "LanguageIDOther", None),
		# Method 'ListFormat' returns object of type 'ListFormat'
		"ListFormat": (68, 2, (9, 0), (), "ListFormat", '{000209C0-0000-0000-C000-000000000046}'),
		# Method 'ListParagraphs' returns object of type 'ListParagraphs'
		"ListParagraphs": (157, 2, (9, 0), (), "ListParagraphs", '{00020991-0000-0000-C000-000000000046}'),
		"ListStyle": (422, 2, (12, 0), (), "ListStyle", None),
		# Method 'Locks' returns object of type 'CoAuthLocks'
		"Locks": (504, 2, (9, 0), (), "Locks", '{DFF99AC2-CD2A-43AD-91B1-A2BE40BC7146}'),
		# Method 'NextStoryRange' returns object of type 'Range'
		"NextStoryRange": (320, 2, (9, 0), (), "NextStoryRange", '{0002095E-0000-0000-C000-000000000046}'),
		"NoProofing": (323, 2, (3, 0), (), "NoProofing", None),
		# Method 'OMaths' returns object of type 'OMaths'
		"OMaths": (346, 2, (9, 0), (), "OMaths", '{873E774B-926A-4CB1-878D-635A45187595}'),
		"Orientation": (317, 2, (3, 0), (), "Orientation", None),
		# Method 'PageSetup' returns object of type 'PageSetup'
		"PageSetup": (1101, 2, (9, 0), (), "PageSetup", '{00020971-0000-0000-C000-000000000046}'),
		# Method 'ParagraphFormat' returns object of type 'ParagraphFormat'
		"ParagraphFormat": (1102, 2, (13, 0), (), "ParagraphFormat", '{000209F4-0000-0000-C000-000000000046}'),
		"ParagraphStyle": (421, 2, (12, 0), (), "ParagraphStyle", None),
		# Method 'Paragraphs' returns object of type 'Paragraphs'
		"Paragraphs": (59, 2, (9, 0), (), "Paragraphs", '{00020958-0000-0000-C000-000000000046}'),
		"Parent": (1002, 2, (9, 0), (), "Parent", None),
		# Method 'ParentContentControl' returns object of type 'ContentControl'
		"ParentContentControl": (501, 2, (9, 0), (), "ParentContentControl", '{EE95AFE3-3026-4172-B078-0E79DAB5CC3D}'),
		"PreviousBookmarkID": (309, 2, (3, 0), (), "PreviousBookmarkID", None),
		# Method 'ReadabilityStatistics' returns object of type 'ReadabilityStatistics'
		"ReadabilityStatistics": (314, 2, (9, 0), (), "ReadabilityStatistics", '{000209AE-0000-0000-C000-000000000046}'),
		# Method 'Revisions' returns object of type 'Revisions'
		"Revisions": (150, 2, (9, 0), (), "Revisions", '{00020980-0000-0000-C000-000000000046}'),
		# Method 'Rows' returns object of type 'Rows'
		"Rows": (303, 2, (9, 0), (), "Rows", '{0002094C-0000-0000-C000-000000000046}'),
		# Method 'Scripts' returns object of type 'Scripts'
		"Scripts": (325, 2, (9, 0), (), "Scripts", '{000C0340-0000-0000-C000-000000000046}'),
		# Method 'Sections' returns object of type 'Sections'
		"Sections": (58, 2, (9, 0), (), "Sections", '{0002095A-0000-0000-C000-000000000046}'),
		# Method 'Sentences' returns object of type 'Sentences'
		"Sentences": (52, 2, (9, 0), (), "Sentences", '{0002095B-0000-0000-C000-000000000046}'),
		# Method 'Shading' returns object of type 'Shading'
		"Shading": (61, 2, (9, 0), (), "Shading", '{0002093A-0000-0000-C000-000000000046}'),
		# Method 'ShapeRange' returns object of type 'ShapeRange'
		"ShapeRange": (311, 2, (9, 0), (), "ShapeRange", '{000209B5-0000-0000-C000-000000000046}'),
		"ShowAll": (408, 2, (11, 0), (), "ShowAll", None),
		# Method 'SmartTags' returns object of type 'SmartTags'
		"SmartTags": (407, 2, (9, 0), (), "SmartTags", '{000209EE-0000-0000-C000-000000000046}'),
		"SpellingChecked": (261, 2, (11, 0), (), "SpellingChecked", None),
		# Method 'SpellingErrors' returns object of type 'ProofreadingErrors'
		"SpellingErrors": (316, 2, (9, 0), (), "SpellingErrors", '{000209BB-0000-0000-C000-000000000046}'),
		"Start": (3, 2, (3, 0), (), "Start", None),
		"StoryLength": (152, 2, (3, 0), (), "StoryLength", None),
		"StoryType": (7, 2, (3, 0), (), "StoryType", None),
		"Style": (151, 2, (12, 0), (), "Style", None),
		# Method 'Subdocuments' returns object of type 'Subdocuments'
		"Subdocuments": (159, 2, (9, 0), (), "Subdocuments", '{00020988-0000-0000-C000-000000000046}'),
		# Method 'SynonymInfo' returns object of type 'SynonymInfo'
		"SynonymInfo": (155, 2, (9, 0), (), "SynonymInfo", '{0002099B-0000-0000-C000-000000000046}'),
		"TableStyle": (423, 2, (12, 0), (), "TableStyle", None),
		# Method 'Tables' returns object of type 'Tables'
		"Tables": (50, 2, (9, 0), (), "Tables", '{0002094D-0000-0000-C000-000000000046}'),
		"Text": (0, 2, (8, 0), (), "Text", None),
		# Method 'TextRetrievalMode' returns object of type 'TextRetrievalMode'
		"TextRetrievalMode": (62, 2, (9, 0), (), "TextRetrievalMode", '{00020939-0000-0000-C000-000000000046}'),
		# Method 'TopLevelTables' returns object of type 'Tables'
		"TopLevelTables": (324, 2, (9, 0), (), "TopLevelTables", '{0002094D-0000-0000-C000-000000000046}'),
		"TwoLinesInOne": (266, 2, (3, 0), (), "TwoLinesInOne", None),
		"Underline": (139, 2, (3, 0), (), "Underline", None),
		# Method 'Updates' returns object of type 'CoAuthUpdates'
		"Updates": (505, 2, (9, 0), (), "Updates", '{30225CFC-5A71-4FE6-B527-90A52C54AE77}'),
		"WordOpenXML": (426, 2, (8, 0), (), "WordOpenXML", None),
		# Method 'Words' returns object of type 'Words'
		"Words": (51, 2, (9, 0), (), "Words", '{0002095C-0000-0000-C000-000000000046}'),
		"XML": (344, 2, (8, 0), ((11, 49),), "XML", None),
		# Method 'XMLNodes' returns object of type 'XMLNodes'
		"XMLNodes": (340, 2, (9, 0), (), "XMLNodes", '{D36C1F42-7044-4B9E-9CA3-85919454DB04}'),
		# Method 'XMLParentNode' returns object of type 'XMLNode'
		"XMLParentNode": (341, 2, (9, 0), (), "XMLParentNode", '{09760240-0B89-49F7-A79D-479F24723F56}'),
	}
	_prop_map_put_ = {
		"Bold": ((130, LCID, 4, 0),()),
		"BoldBi": ((400, LCID, 4, 0),()),
		"Borders": ((1100, LCID, 4, 0),()),
		"Case": ((312, LCID, 4, 0),()),
		"CharacterWidth": ((326, LCID, 4, 0),()),
		"CombineCharacters": ((267, LCID, 4, 0),()),
		"DisableCharacterSpaceGrid": ((141, LCID, 4, 0),()),
		"EmphasisMark": ((140, LCID, 4, 0),()),
		"End": ((4, LCID, 4, 0),()),
		"FitTextWidth": ((264, LCID, 4, 0),()),
		"Font": ((5, LCID, 4, 0),()),
		"FormattedText": ((2, LCID, 4, 0),()),
		"GrammarChecked": ((260, LCID, 4, 0),()),
		"HighlightColorIndex": ((301, LCID, 4, 0),()),
		"HorizontalInVertical": ((265, LCID, 4, 0),()),
		"ID": ((405, LCID, 4, 0),()),
		"Italic": ((131, LCID, 4, 0),()),
		"ItalicBi": ((401, LCID, 4, 0),()),
		"Kana": ((327, LCID, 4, 0),()),
		"LanguageDetected": ((263, LCID, 4, 0),()),
		"LanguageID": ((153, LCID, 4, 0),()),
		"LanguageIDFarEast": ((321, LCID, 4, 0),()),
		"LanguageIDOther": ((322, LCID, 4, 0),()),
		"NoProofing": ((323, LCID, 4, 0),()),
		"Orientation": ((317, LCID, 4, 0),()),
		"PageSetup": ((1101, LCID, 4, 0),()),
		"ParagraphFormat": ((1102, LCID, 4, 0),()),
		"ShowAll": ((408, LCID, 4, 0),()),
		"SpellingChecked": ((261, LCID, 4, 0),()),
		"Start": ((3, LCID, 4, 0),()),
		"Style": ((151, LCID, 4, 0),()),
		"Text": ((0, LCID, 4, 0),()),
		"TextRetrievalMode": ((62, LCID, 4, 0),()),
		"TwoLinesInOne": ((266, LCID, 4, 0),()),
		"Underline": ((139, LCID, 4, 0),()),
	}
	# Default property for this class is 'Text'
	def __call__(self):
		return self._ApplyTypes_(*(0, 2, (8, 0), (), "Text", None))
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

win32com.client.CLSIDToClass.RegisterCLSID( "{0002095E-0000-0000-C000-000000000046}", Range )
# -*- coding: mbcs -*-
# Created by makepy.py version 0.5.01
# By python version 3.3.1 (v3.3.1:d9893d13c628, Apr  6 2013, 20:25:12) [MSC v.1600 32 bit (Intel)]
# From type library '{00020905-0000-0000-C000-000000000046}'
# On Thu Mar  6 15:34:15 2014
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

Range_vtables_dispatch_ = 1
Range_vtables_ = [
	(( 'Text' , 'prop' , ), 0, (0, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 28 , (3, 0, None, None) , 0 , )),
	(( 'Text' , 'prop' , ), 0, (0, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 32 , (3, 0, None, None) , 0 , )),
	(( 'FormattedText' , 'prop' , ), 2, (2, (), [ (16393, 10, None, "IID('{0002095E-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 36 , (3, 0, None, None) , 0 , )),
	(( 'FormattedText' , 'prop' , ), 2, (2, (), [ (9, 1, None, "IID('{0002095E-0000-0000-C000-000000000046}')") , ], 1 , 4 , 4 , 0 , 40 , (3, 0, None, None) , 0 , )),
	(( 'Start' , 'prop' , ), 3, (3, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 44 , (3, 0, None, None) , 0 , )),
	(( 'Start' , 'prop' , ), 3, (3, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 48 , (3, 0, None, None) , 0 , )),
	(( 'End' , 'prop' , ), 4, (4, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 52 , (3, 0, None, None) , 0 , )),
	(( 'End' , 'prop' , ), 4, (4, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 56 , (3, 0, None, None) , 0 , )),
	(( 'Font' , 'prop' , ), 5, (5, (), [ (16397, 10, None, "IID('{000209F5-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 60 , (3, 0, None, None) , 0 , )),
	(( 'Font' , 'prop' , ), 5, (5, (), [ (13, 1, None, "IID('{000209F5-0000-0000-C000-000000000046}')") , ], 1 , 4 , 4 , 0 , 64 , (3, 0, None, None) , 0 , )),
	(( 'Duplicate' , 'prop' , ), 6, (6, (), [ (16393, 10, None, "IID('{0002095E-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 68 , (3, 0, None, None) , 0 , )),
	(( 'StoryType' , 'prop' , ), 7, (7, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 72 , (3, 0, None, None) , 0 , )),
	(( 'Tables' , 'prop' , ), 50, (50, (), [ (16393, 10, None, "IID('{0002094D-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 76 , (3, 0, None, None) , 0 , )),
	(( 'Words' , 'prop' , ), 51, (51, (), [ (16393, 10, None, "IID('{0002095C-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 80 , (3, 0, None, None) , 0 , )),
	(( 'Sentences' , 'prop' , ), 52, (52, (), [ (16393, 10, None, "IID('{0002095B-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 84 , (3, 0, None, None) , 0 , )),
	(( 'Characters' , 'prop' , ), 53, (53, (), [ (16393, 10, None, "IID('{0002095D-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 88 , (3, 0, None, None) , 0 , )),
	(( 'Footnotes' , 'prop' , ), 54, (54, (), [ (16393, 10, None, "IID('{00020942-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 92 , (3, 0, None, None) , 0 , )),
	(( 'Endnotes' , 'prop' , ), 55, (55, (), [ (16393, 10, None, "IID('{00020941-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 96 , (3, 0, None, None) , 0 , )),
	(( 'Comments' , 'prop' , ), 56, (56, (), [ (16393, 10, None, "IID('{00020940-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 100 , (3, 0, None, None) , 0 , )),
	(( 'Cells' , 'prop' , ), 57, (57, (), [ (16393, 10, None, "IID('{0002094A-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 104 , (3, 0, None, None) , 0 , )),
	(( 'Sections' , 'prop' , ), 58, (58, (), [ (16393, 10, None, "IID('{0002095A-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 108 , (3, 0, None, None) , 0 , )),
	(( 'Paragraphs' , 'prop' , ), 59, (59, (), [ (16393, 10, None, "IID('{00020958-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 112 , (3, 0, None, None) , 0 , )),
	(( 'Borders' , 'prop' , ), 1100, (1100, (), [ (16393, 10, None, "IID('{0002093C-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 116 , (3, 0, None, None) , 0 , )),
	(( 'Borders' , 'prop' , ), 1100, (1100, (), [ (9, 1, None, "IID('{0002093C-0000-0000-C000-000000000046}')") , ], 1 , 4 , 4 , 0 , 120 , (3, 0, None, None) , 0 , )),
	(( 'Shading' , 'prop' , ), 61, (61, (), [ (16393, 10, None, "IID('{0002093A-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 124 , (3, 0, None, None) , 0 , )),
	(( 'TextRetrievalMode' , 'prop' , ), 62, (62, (), [ (16393, 10, None, "IID('{00020939-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 128 , (3, 0, None, None) , 0 , )),
	(( 'TextRetrievalMode' , 'prop' , ), 62, (62, (), [ (9, 1, None, "IID('{00020939-0000-0000-C000-000000000046}')") , ], 1 , 4 , 4 , 0 , 132 , (3, 0, None, None) , 0 , )),
	(( 'Fields' , 'prop' , ), 64, (64, (), [ (16393, 10, None, "IID('{00020930-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 136 , (3, 0, None, None) , 0 , )),
	(( 'FormFields' , 'prop' , ), 65, (65, (), [ (16393, 10, None, "IID('{00020929-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 140 , (3, 0, None, None) , 0 , )),
	(( 'Frames' , 'prop' , ), 66, (66, (), [ (16393, 10, None, "IID('{0002092B-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 144 , (3, 0, None, None) , 0 , )),
	(( 'ParagraphFormat' , 'prop' , ), 1102, (1102, (), [ (16397, 10, None, "IID('{000209F4-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 148 , (3, 0, None, None) , 0 , )),
	(( 'ParagraphFormat' , 'prop' , ), 1102, (1102, (), [ (13, 1, None, "IID('{000209F4-0000-0000-C000-000000000046}')") , ], 1 , 4 , 4 , 0 , 152 , (3, 0, None, None) , 0 , )),
	(( 'ListFormat' , 'prop' , ), 68, (68, (), [ (16393, 10, None, "IID('{000209C0-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 156 , (3, 0, None, None) , 0 , )),
	(( 'Bookmarks' , 'prop' , ), 75, (75, (), [ (16393, 10, None, "IID('{00020967-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 160 , (3, 0, None, None) , 0 , )),
	(( 'Application' , 'prop' , ), 1000, (1000, (), [ (16397, 10, None, "IID('{000209FF-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 164 , (3, 0, None, None) , 0 , )),
	(( 'Creator' , 'prop' , ), 1001, (1001, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 168 , (3, 0, None, None) , 0 , )),
	(( 'Parent' , 'prop' , ), 1002, (1002, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 172 , (3, 0, None, None) , 0 , )),
	(( 'Bold' , 'prop' , ), 130, (130, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 176 , (3, 0, None, None) , 0 , )),
	(( 'Bold' , 'prop' , ), 130, (130, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 180 , (3, 0, None, None) , 0 , )),
	(( 'Italic' , 'prop' , ), 131, (131, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 184 , (3, 0, None, None) , 0 , )),
	(( 'Italic' , 'prop' , ), 131, (131, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 188 , (3, 0, None, None) , 0 , )),
	(( 'Underline' , 'prop' , ), 139, (139, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 192 , (3, 0, None, None) , 0 , )),
	(( 'Underline' , 'prop' , ), 139, (139, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 196 , (3, 0, None, None) , 0 , )),
	(( 'EmphasisMark' , 'prop' , ), 140, (140, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 200 , (3, 0, None, None) , 0 , )),
	(( 'EmphasisMark' , 'prop' , ), 140, (140, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 204 , (3, 0, None, None) , 0 , )),
	(( 'DisableCharacterSpaceGrid' , 'prop' , ), 141, (141, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 208 , (3, 0, None, None) , 0 , )),
	(( 'DisableCharacterSpaceGrid' , 'prop' , ), 141, (141, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 212 , (3, 0, None, None) , 0 , )),
	(( 'Revisions' , 'prop' , ), 150, (150, (), [ (16393, 10, None, "IID('{00020980-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 216 , (3, 0, None, None) , 0 , )),
	(( 'Style' , 'prop' , ), 151, (151, (), [ (16396, 10, None, None) , ], 1 , 2 , 4 , 0 , 220 , (3, 0, None, None) , 0 , )),
	(( 'Style' , 'prop' , ), 151, (151, (), [ (16396, 1, None, None) , ], 1 , 4 , 4 , 0 , 224 , (3, 0, None, None) , 0 , )),
	(( 'StoryLength' , 'prop' , ), 152, (152, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 228 , (3, 0, None, None) , 0 , )),
	(( 'LanguageID' , 'prop' , ), 153, (153, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 232 , (3, 0, None, None) , 0 , )),
	(( 'LanguageID' , 'prop' , ), 153, (153, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 236 , (3, 0, None, None) , 0 , )),
	(( 'SynonymInfo' , 'prop' , ), 155, (155, (), [ (16393, 10, None, "IID('{0002099B-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 240 , (3, 0, None, None) , 0 , )),
	(( 'Hyperlinks' , 'prop' , ), 156, (156, (), [ (16393, 10, None, "IID('{0002099C-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 244 , (3, 0, None, None) , 0 , )),
	(( 'ListParagraphs' , 'prop' , ), 157, (157, (), [ (16393, 10, None, "IID('{00020991-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 248 , (3, 0, None, None) , 0 , )),
	(( 'Subdocuments' , 'prop' , ), 159, (159, (), [ (16393, 10, None, "IID('{00020988-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 252 , (3, 0, None, None) , 0 , )),
	(( 'GrammarChecked' , 'prop' , ), 260, (260, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 256 , (3, 0, None, None) , 0 , )),
	(( 'GrammarChecked' , 'prop' , ), 260, (260, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 260 , (3, 0, None, None) , 0 , )),
	(( 'SpellingChecked' , 'prop' , ), 261, (261, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 264 , (3, 0, None, None) , 0 , )),
	(( 'SpellingChecked' , 'prop' , ), 261, (261, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 268 , (3, 0, None, None) , 0 , )),
	(( 'HighlightColorIndex' , 'prop' , ), 301, (301, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 272 , (3, 0, None, None) , 0 , )),
	(( 'HighlightColorIndex' , 'prop' , ), 301, (301, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 276 , (3, 0, None, None) , 0 , )),
	(( 'Columns' , 'prop' , ), 302, (302, (), [ (16393, 10, None, "IID('{0002094B-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 280 , (3, 0, None, None) , 0 , )),
	(( 'Rows' , 'prop' , ), 303, (303, (), [ (16393, 10, None, "IID('{0002094C-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 284 , (3, 0, None, None) , 0 , )),
	(( 'CanEdit' , 'prop' , ), 304, (304, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 288 , (3, 0, None, None) , 64 , )),
	(( 'CanPaste' , 'prop' , ), 305, (305, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 292 , (3, 0, None, None) , 64 , )),
	(( 'IsEndOfRowMark' , 'prop' , ), 307, (307, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 296 , (3, 0, None, None) , 0 , )),
	(( 'BookmarkID' , 'prop' , ), 308, (308, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 300 , (3, 0, None, None) , 0 , )),
	(( 'PreviousBookmarkID' , 'prop' , ), 309, (309, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 304 , (3, 0, None, None) , 0 , )),
	(( 'Find' , 'prop' , ), 262, (262, (), [ (16393, 10, None, "IID('{000209B0-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 308 , (3, 0, None, None) , 0 , )),
	(( 'PageSetup' , 'prop' , ), 1101, (1101, (), [ (16393, 10, None, "IID('{00020971-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 312 , (3, 0, None, None) , 0 , )),
	(( 'PageSetup' , 'prop' , ), 1101, (1101, (), [ (9, 1, None, "IID('{00020971-0000-0000-C000-000000000046}')") , ], 1 , 4 , 4 , 0 , 316 , (3, 0, None, None) , 0 , )),
	(( 'ShapeRange' , 'prop' , ), 311, (311, (), [ (16393, 10, None, "IID('{000209B5-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 320 , (3, 0, None, None) , 0 , )),
	(( 'Case' , 'prop' , ), 312, (312, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 324 , (3, 0, None, None) , 0 , )),
	(( 'Case' , 'prop' , ), 312, (312, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 328 , (3, 0, None, None) , 0 , )),
	(( 'Information' , 'Type' , 'prop' , ), 313, (313, (), [ (3, 1, None, None) , 
			 (16396, 10, None, None) , ], 1 , 2 , 4 , 0 , 332 , (3, 0, None, None) , 0 , )),
	(( 'ReadabilityStatistics' , 'prop' , ), 314, (314, (), [ (16393, 10, None, "IID('{000209AE-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 336 , (3, 0, None, None) , 0 , )),
	(( 'GrammaticalErrors' , 'prop' , ), 315, (315, (), [ (16393, 10, None, "IID('{000209BB-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 340 , (3, 0, None, None) , 0 , )),
	(( 'SpellingErrors' , 'prop' , ), 316, (316, (), [ (16393, 10, None, "IID('{000209BB-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 344 , (3, 0, None, None) , 0 , )),
	(( 'Orientation' , 'prop' , ), 317, (317, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 348 , (3, 0, None, None) , 0 , )),
	(( 'Orientation' , 'prop' , ), 317, (317, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 352 , (3, 0, None, None) , 0 , )),
	(( 'InlineShapes' , 'prop' , ), 319, (319, (), [ (16393, 10, None, "IID('{000209A9-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 356 , (3, 0, None, None) , 0 , )),
	(( 'NextStoryRange' , 'prop' , ), 320, (320, (), [ (16393, 10, None, "IID('{0002095E-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 360 , (3, 0, None, None) , 0 , )),
	(( 'LanguageIDFarEast' , 'prop' , ), 321, (321, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 364 , (3, 0, None, None) , 0 , )),
	(( 'LanguageIDFarEast' , 'prop' , ), 321, (321, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 368 , (3, 0, None, None) , 0 , )),
	(( 'LanguageIDOther' , 'prop' , ), 322, (322, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 372 , (3, 0, None, None) , 0 , )),
	(( 'LanguageIDOther' , 'prop' , ), 322, (322, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 376 , (3, 0, None, None) , 0 , )),
	(( 'Select' , ), 65535, (65535, (), [ ], 1 , 1 , 4 , 0 , 380 , (3, 0, None, None) , 0 , )),
	(( 'SetRange' , 'Start' , 'End' , ), 100, (100, (), [ (3, 1, None, None) , 
			 (3, 1, None, None) , ], 1 , 1 , 4 , 0 , 384 , (3, 0, None, None) , 0 , )),
	(( 'Collapse' , 'Direction' , ), 101, (101, (), [ (16396, 17, None, None) , ], 1 , 1 , 4 , 1 , 388 , (3, 0, None, None) , 0 , )),
	(( 'InsertBefore' , 'Text' , ), 102, (102, (), [ (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 392 , (3, 0, None, None) , 0 , )),
	(( 'InsertAfter' , 'Text' , ), 104, (104, (), [ (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 396 , (3, 0, None, None) , 0 , )),
	(( 'Next' , 'Unit' , 'Count' , 'prop' , ), 105, (105, (), [ 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16393, 10, None, "IID('{0002095E-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 2 , 400 , (3, 0, None, None) , 0 , )),
	(( 'Previous' , 'Unit' , 'Count' , 'prop' , ), 106, (106, (), [ 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16393, 10, None, "IID('{0002095E-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 2 , 404 , (3, 0, None, None) , 0 , )),
	(( 'StartOf' , 'Unit' , 'Extend' , 'prop' , ), 107, (107, (), [ 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16387, 10, None, None) , ], 1 , 1 , 4 , 2 , 408 , (3, 0, None, None) , 0 , )),
	(( 'EndOf' , 'Unit' , 'Extend' , 'prop' , ), 108, (108, (), [ 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16387, 10, None, None) , ], 1 , 1 , 4 , 2 , 412 , (3, 0, None, None) , 0 , )),
	(( 'Move' , 'Unit' , 'Count' , 'prop' , ), 109, (109, (), [ 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16387, 10, None, None) , ], 1 , 1 , 4 , 2 , 416 , (3, 0, None, None) , 0 , )),
	(( 'MoveStart' , 'Unit' , 'Count' , 'prop' , ), 110, (110, (), [ 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16387, 10, None, None) , ], 1 , 1 , 4 , 2 , 420 , (3, 0, None, None) , 0 , )),
	(( 'MoveEnd' , 'Unit' , 'Count' , 'prop' , ), 111, (111, (), [ 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16387, 10, None, None) , ], 1 , 1 , 4 , 2 , 424 , (3, 0, None, None) , 0 , )),
	(( 'MoveWhile' , 'Cset' , 'Count' , 'prop' , ), 112, (112, (), [ 
			 (16396, 1, None, None) , (16396, 17, None, None) , (16387, 10, None, None) , ], 1 , 1 , 4 , 1 , 428 , (3, 0, None, None) , 0 , )),
	(( 'MoveStartWhile' , 'Cset' , 'Count' , 'prop' , ), 113, (113, (), [ 
			 (16396, 1, None, None) , (16396, 17, None, None) , (16387, 10, None, None) , ], 1 , 1 , 4 , 1 , 432 , (3, 0, None, None) , 0 , )),
	(( 'MoveEndWhile' , 'Cset' , 'Count' , 'prop' , ), 114, (114, (), [ 
			 (16396, 1, None, None) , (16396, 17, None, None) , (16387, 10, None, None) , ], 1 , 1 , 4 , 1 , 436 , (3, 0, None, None) , 0 , )),
	(( 'MoveUntil' , 'Cset' , 'Count' , 'prop' , ), 115, (115, (), [ 
			 (16396, 1, None, None) , (16396, 17, None, None) , (16387, 10, None, None) , ], 1 , 1 , 4 , 1 , 440 , (3, 0, None, None) , 0 , )),
	(( 'MoveStartUntil' , 'Cset' , 'Count' , 'prop' , ), 116, (116, (), [ 
			 (16396, 1, None, None) , (16396, 17, None, None) , (16387, 10, None, None) , ], 1 , 1 , 4 , 1 , 444 , (3, 0, None, None) , 0 , )),
	(( 'MoveEndUntil' , 'Cset' , 'Count' , 'prop' , ), 117, (117, (), [ 
			 (16396, 1, None, None) , (16396, 17, None, None) , (16387, 10, None, None) , ], 1 , 1 , 4 , 1 , 448 , (3, 0, None, None) , 0 , )),
	(( 'Cut' , ), 119, (119, (), [ ], 1 , 1 , 4 , 0 , 452 , (3, 0, None, None) , 0 , )),
	(( 'Copy' , ), 120, (120, (), [ ], 1 , 1 , 4 , 0 , 456 , (3, 0, None, None) , 0 , )),
	(( 'Paste' , ), 121, (121, (), [ ], 1 , 1 , 4 , 0 , 460 , (3, 0, None, None) , 0 , )),
	(( 'InsertBreak' , 'Type' , ), 122, (122, (), [ (16396, 17, None, None) , ], 1 , 1 , 4 , 1 , 464 , (3, 0, None, None) , 0 , )),
	(( 'InsertFile' , 'FileName' , 'Range' , 'ConfirmConversions' , 'Link' , 
			 'Attachment' , ), 123, (123, (), [ (8, 1, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , ], 1 , 1 , 4 , 4 , 468 , (3, 0, None, None) , 0 , )),
	(( 'InStory' , 'Range' , 'prop' , ), 125, (125, (), [ (9, 1, None, "IID('{0002095E-0000-0000-C000-000000000046}')") , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 472 , (3, 0, None, None) , 0 , )),
	(( 'InRange' , 'Range' , 'prop' , ), 126, (126, (), [ (9, 1, None, "IID('{0002095E-0000-0000-C000-000000000046}')") , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 476 , (3, 0, None, None) , 0 , )),
	(( 'Delete' , 'Unit' , 'Count' , 'prop' , ), 127, (127, (), [ 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16387, 10, None, None) , ], 1 , 1 , 4 , 2 , 480 , (3, 0, None, None) , 0 , )),
	(( 'WholeStory' , ), 128, (128, (), [ ], 1 , 1 , 4 , 0 , 484 , (3, 0, None, None) , 0 , )),
	(( 'Expand' , 'Unit' , 'prop' , ), 129, (129, (), [ (16396, 17, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 1 , 488 , (3, 0, None, None) , 0 , )),
	(( 'InsertParagraph' , ), 160, (160, (), [ ], 1 , 1 , 4 , 0 , 492 , (3, 0, None, None) , 0 , )),
	(( 'InsertParagraphAfter' , ), 161, (161, (), [ ], 1 , 1 , 4 , 0 , 496 , (3, 0, None, None) , 0 , )),
	(( 'ConvertToTableOld' , 'Separator' , 'NumRows' , 'NumColumns' , 'InitialColumnWidth' , 
			 'Format' , 'ApplyBorders' , 'ApplyShading' , 'ApplyFont' , 'ApplyColor' , 
			 'ApplyHeadingRows' , 'ApplyLastRow' , 'ApplyFirstColumn' , 'ApplyLastColumn' , 'AutoFit' , 
			 'prop' , ), 162, (162, (), [ (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16393, 10, None, "IID('{00020951-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 14 , 500 , (3, 0, None, None) , 64 , )),
	(( 'InsertDateTimeOld' , 'DateTimeFormat' , 'InsertAsField' , 'InsertAsFullWidth' , ), 163, (163, (), [ 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , ], 1 , 1 , 4 , 3 , 504 , (3, 0, None, None) , 64 , )),
	(( 'InsertSymbol' , 'CharacterNumber' , 'Font' , 'Unicode' , 'Bias' , 
			 ), 164, (164, (), [ (3, 1, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , ], 1 , 1 , 4 , 3 , 508 , (3, 0, None, None) , 0 , )),
	(( 'InsertCrossReference_2002' , 'ReferenceType' , 'ReferenceKind' , 'ReferenceItem' , 'InsertAsHyperlink' , 
			 'IncludePosition' , ), 165, (165, (), [ (16396, 1, None, None) , (3, 1, None, None) , (16396, 1, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , ], 1 , 1 , 4 , 2 , 512 , (3, 0, None, None) , 64 , )),
	(( 'InsertCaptionXP' , 'Label' , 'Title' , 'TitleAutoText' , 'Position' , 
			 ), 166, (166, (), [ (16396, 1, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , ], 1 , 1 , 4 , 3 , 516 , (3, 0, None, None) , 64 , )),
	(( 'CopyAsPicture' , ), 167, (167, (), [ ], 1 , 1 , 4 , 0 , 520 , (3, 0, None, None) , 0 , )),
	(( 'SortOld' , 'ExcludeHeader' , 'FieldNumber' , 'SortFieldType' , 'SortOrder' , 
			 'FieldNumber2' , 'SortFieldType2' , 'SortOrder2' , 'FieldNumber3' , 'SortFieldType3' , 
			 'SortOrder3' , 'SortColumn' , 'Separator' , 'CaseSensitive' , 'LanguageID' , 
			 ), 168, (168, (), [ (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , ], 1 , 1 , 4 , 14 , 524 , (3, 0, None, None) , 64 , )),
	(( 'SortAscending' , ), 169, (169, (), [ ], 1 , 1 , 4 , 0 , 528 , (3, 0, None, None) , 0 , )),
	(( 'SortDescending' , ), 170, (170, (), [ ], 1 , 1 , 4 , 0 , 532 , (3, 0, None, None) , 0 , )),
	(( 'IsEqual' , 'Range' , 'prop' , ), 171, (171, (), [ (9, 1, None, "IID('{0002095E-0000-0000-C000-000000000046}')") , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 536 , (3, 0, None, None) , 0 , )),
	(( 'Calculate' , 'prop' , ), 172, (172, (), [ (16388, 10, None, None) , ], 1 , 1 , 4 , 0 , 540 , (3, 0, None, None) , 0 , )),
	(( 'GoTo' , 'What' , 'Which' , 'Count' , 'Name' , 
			 'prop' , ), 173, (173, (), [ (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16393, 10, None, "IID('{0002095E-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 4 , 544 , (3, 0, None, None) , 0 , )),
	(( 'GoToNext' , 'What' , 'prop' , ), 174, (174, (), [ (3, 1, None, None) , 
			 (16393, 10, None, "IID('{0002095E-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 0 , 548 , (3, 0, None, None) , 0 , )),
	(( 'GoToPrevious' , 'What' , 'prop' , ), 175, (175, (), [ (3, 1, None, None) , 
			 (16393, 10, None, "IID('{0002095E-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 0 , 552 , (3, 0, None, None) , 0 , )),
	(( 'PasteSpecial' , 'IconIndex' , 'Link' , 'Placement' , 'DisplayAsIcon' , 
			 'DataType' , 'IconFileName' , 'IconLabel' , ), 176, (176, (), [ (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , ], 1 , 1 , 4 , 7 , 556 , (3, 0, None, None) , 0 , )),
	(( 'LookupNameProperties' , ), 177, (177, (), [ ], 1 , 1 , 4 , 0 , 560 , (3, 0, None, None) , 0 , )),
	(( 'ComputeStatistics' , 'Statistic' , 'prop' , ), 178, (178, (), [ (3, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 564 , (3, 0, None, None) , 0 , )),
	(( 'Relocate' , 'Direction' , ), 179, (179, (), [ (3, 1, None, None) , ], 1 , 1 , 4 , 0 , 568 , (3, 0, None, None) , 0 , )),
	(( 'CheckSynonyms' , ), 180, (180, (), [ ], 1 , 1 , 4 , 0 , 572 , (3, 0, None, None) , 0 , )),
	(( 'SubscribeTo' , 'Edition' , 'Format' , ), 181, (181, (), [ (8, 1, None, None) , 
			 (16396, 17, None, None) , ], 1 , 1 , 4 , 1 , 576 , (3, 0, None, None) , 64 , )),
	(( 'CreatePublisher' , 'Edition' , 'ContainsPICT' , 'ContainsRTF' , 'ContainsText' , 
			 ), 182, (182, (), [ (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , ], 1 , 1 , 4 , 4 , 580 , (3, 0, None, None) , 64 , )),
	(( 'InsertAutoText' , ), 183, (183, (), [ ], 1 , 1 , 4 , 0 , 584 , (3, 0, None, None) , 0 , )),
	(( 'InsertDatabase' , 'Format' , 'Style' , 'LinkToSource' , 'Connection' , 
			 'SQLStatement' , 'SQLStatement1' , 'PasswordDocument' , 'PasswordTemplate' , 'WritePasswordDocument' , 
			 'WritePasswordTemplate' , 'DataSource' , 'From' , 'To' , 'IncludeFields' , 
			 ), 194, (194, (), [ (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , ], 1 , 1 , 4 , 14 , 588 , (3, 0, None, None) , 0 , )),
	(( 'AutoFormat' , ), 195, (195, (), [ ], 1 , 1 , 4 , 0 , 592 , (3, 0, None, None) , 0 , )),
	(( 'CheckGrammar' , ), 204, (204, (), [ ], 1 , 1 , 4 , 0 , 596 , (3, 0, None, None) , 0 , )),
	(( 'CheckSpelling' , 'CustomDictionary' , 'IgnoreUppercase' , 'AlwaysSuggest' , 'CustomDictionary2' , 
			 'CustomDictionary3' , 'CustomDictionary4' , 'CustomDictionary5' , 'CustomDictionary6' , 'CustomDictionary7' , 
			 'CustomDictionary8' , 'CustomDictionary9' , 'CustomDictionary10' , ), 205, (205, (), [ (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , ], 1 , 1 , 4 , 12 , 600 , (3, 0, None, None) , 0 , )),
	(( 'GetSpellingSuggestions' , 'CustomDictionary' , 'IgnoreUppercase' , 'MainDictionary' , 'SuggestionMode' , 
			 'CustomDictionary2' , 'CustomDictionary3' , 'CustomDictionary4' , 'CustomDictionary5' , 'CustomDictionary6' , 
			 'CustomDictionary7' , 'CustomDictionary8' , 'CustomDictionary9' , 'CustomDictionary10' , 'prop' , 
			 ), 209, (209, (), [ (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16393, 10, None, "IID('{000209AA-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 13 , 604 , (3, 0, None, None) , 0 , )),
	(( 'InsertParagraphBefore' , ), 212, (212, (), [ ], 1 , 1 , 4 , 0 , 608 , (3, 0, None, None) , 0 , )),
	(( 'NextSubdocument' , ), 219, (219, (), [ ], 1 , 1 , 4 , 0 , 612 , (3, 0, None, None) , 0 , )),
	(( 'PreviousSubdocument' , ), 220, (220, (), [ ], 1 , 1 , 4 , 0 , 616 , (3, 0, None, None) , 0 , )),
	(( 'ConvertHangulAndHanja' , 'ConversionsMode' , 'FastConversion' , 'CheckHangulEnding' , 'EnableRecentOrdering' , 
			 'CustomDictionary' , ), 221, (221, (), [ (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , ], 1 , 1 , 4 , 5 , 620 , (3, 0, None, None) , 0 , )),
	(( 'PasteAsNestedTable' , ), 222, (222, (), [ ], 1 , 1 , 4 , 0 , 624 , (3, 0, None, None) , 0 , )),
	(( 'ModifyEnclosure' , 'Style' , 'Symbol' , 'EnclosedText' , ), 223, (223, (), [ 
			 (16396, 1, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , ], 1 , 1 , 4 , 2 , 628 , (3, 0, None, None) , 0 , )),
	(( 'PhoneticGuide' , 'Text' , 'Alignment' , 'Raise' , 'FontSize' , 
			 'FontName' , ), 224, (224, (), [ (8, 1, None, None) , (3, 49, '-1', None) , (3, 49, '0', None) , 
			 (3, 49, '0', None) , (8, 49, "''", None) , ], 1 , 1 , 4 , 0 , 632 , (3, 32, None, None) , 0 , )),
	(( 'InsertDateTime' , 'DateTimeFormat' , 'InsertAsField' , 'InsertAsFullWidth' , 'DateLanguage' , 
			 'CalendarType' , ), 444, (444, (), [ (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , ], 1 , 1 , 4 , 5 , 636 , (3, 0, None, None) , 0 , )),
	(( 'Sort' , 'ExcludeHeader' , 'FieldNumber' , 'SortFieldType' , 'SortOrder' , 
			 'FieldNumber2' , 'SortFieldType2' , 'SortOrder2' , 'FieldNumber3' , 'SortFieldType3' , 
			 'SortOrder3' , 'SortColumn' , 'Separator' , 'CaseSensitive' , 'BidiSort' , 
			 'IgnoreThe' , 'IgnoreKashida' , 'IgnoreDiacritics' , 'IgnoreHe' , 'LanguageID' , 
			 ), 484, (484, (), [ (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , ], 1 , 1 , 4 , 19 , 640 , (3, 0, None, None) , 0 , )),
	(( 'DetectLanguage' , ), 203, (203, (), [ ], 1 , 1 , 4 , 0 , 644 , (3, 0, None, None) , 0 , )),
	(( 'ConvertToTable' , 'Separator' , 'NumRows' , 'NumColumns' , 'InitialColumnWidth' , 
			 'Format' , 'ApplyBorders' , 'ApplyShading' , 'ApplyFont' , 'ApplyColor' , 
			 'ApplyHeadingRows' , 'ApplyLastRow' , 'ApplyFirstColumn' , 'ApplyLastColumn' , 'AutoFit' , 
			 'AutoFitBehavior' , 'DefaultTableBehavior' , 'prop' , ), 498, (498, (), [ (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16393, 10, None, "IID('{00020951-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 16 , 648 , (3, 0, None, None) , 0 , )),
	(( 'TCSCConverter' , 'WdTCSCConverterDirection' , 'CommonTerms' , 'UseVariants' , ), 499, (499, (), [ 
			 (3, 49, '2', None) , (11, 49, 'False', None) , (11, 49, 'False', None) , ], 1 , 1 , 4 , 0 , 652 , (3, 0, None, None) , 0 , )),
	(( 'LanguageDetected' , 'prop' , ), 263, (263, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 656 , (3, 0, None, None) , 0 , )),
	(( 'LanguageDetected' , 'prop' , ), 263, (263, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 660 , (3, 0, None, None) , 0 , )),
	(( 'FitTextWidth' , 'prop' , ), 264, (264, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 664 , (3, 0, None, None) , 0 , )),
	(( 'FitTextWidth' , 'prop' , ), 264, (264, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 668 , (3, 0, None, None) , 0 , )),
	(( 'HorizontalInVertical' , 'prop' , ), 265, (265, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 672 , (3, 0, None, None) , 0 , )),
	(( 'HorizontalInVertical' , 'prop' , ), 265, (265, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 676 , (3, 0, None, None) , 0 , )),
	(( 'TwoLinesInOne' , 'prop' , ), 266, (266, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 680 , (3, 0, None, None) , 0 , )),
	(( 'TwoLinesInOne' , 'prop' , ), 266, (266, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 684 , (3, 0, None, None) , 0 , )),
	(( 'CombineCharacters' , 'prop' , ), 267, (267, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 688 , (3, 0, None, None) , 0 , )),
	(( 'CombineCharacters' , 'prop' , ), 267, (267, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 692 , (3, 0, None, None) , 0 , )),
	(( 'NoProofing' , 'prop' , ), 323, (323, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 696 , (3, 0, None, None) , 0 , )),
	(( 'NoProofing' , 'prop' , ), 323, (323, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 700 , (3, 0, None, None) , 0 , )),
	(( 'TopLevelTables' , 'prop' , ), 324, (324, (), [ (16393, 10, None, "IID('{0002094D-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 704 , (3, 0, None, None) , 0 , )),
	(( 'Scripts' , 'prop' , ), 325, (325, (), [ (16393, 10, None, "IID('{000C0340-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 708 , (3, 0, None, None) , 0 , )),
	(( 'CharacterWidth' , 'prop' , ), 326, (326, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 712 , (3, 0, None, None) , 0 , )),
	(( 'CharacterWidth' , 'prop' , ), 326, (326, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 716 , (3, 0, None, None) , 0 , )),
	(( 'Kana' , 'prop' , ), 327, (327, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 720 , (3, 0, None, None) , 0 , )),
	(( 'Kana' , 'prop' , ), 327, (327, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 724 , (3, 0, None, None) , 0 , )),
	(( 'BoldBi' , 'prop' , ), 400, (400, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 728 , (3, 0, None, None) , 0 , )),
	(( 'BoldBi' , 'prop' , ), 400, (400, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 732 , (3, 0, None, None) , 0 , )),
	(( 'ItalicBi' , 'prop' , ), 401, (401, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 736 , (3, 0, None, None) , 0 , )),
	(( 'ItalicBi' , 'prop' , ), 401, (401, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 740 , (3, 0, None, None) , 0 , )),
	(( 'ID' , 'prop' , ), 405, (405, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 744 , (3, 0, None, None) , 0 , )),
	(( 'ID' , 'prop' , ), 405, (405, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 748 , (3, 0, None, None) , 0 , )),
	(( 'HTMLDivisions' , 'prop' , ), 406, (406, (), [ (16393, 10, None, "IID('{000209E8-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 752 , (3, 0, None, None) , 0 , )),
	(( 'SmartTags' , 'prop' , ), 407, (407, (), [ (16393, 10, None, "IID('{000209EE-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 756 , (3, 0, None, None) , 64 , )),
	(( 'ShowAll' , 'prop' , ), 408, (408, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 760 , (3, 0, None, None) , 0 , )),
	(( 'ShowAll' , 'prop' , ), 408, (408, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 764 , (3, 0, None, None) , 0 , )),
	(( 'Document' , 'prop' , ), 409, (409, (), [ (16397, 10, None, "IID('{00020906-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 768 , (3, 0, None, None) , 0 , )),
	(( 'FootnoteOptions' , 'prop' , ), 410, (410, (), [ (16393, 10, None, "IID('{BEA85A24-D7DA-4F3D-B58C-ED90FB01D615}')") , ], 1 , 2 , 4 , 0 , 772 , (3, 0, None, None) , 0 , )),
	(( 'EndnoteOptions' , 'prop' , ), 411, (411, (), [ (16393, 10, None, "IID('{BF043168-F4DE-4E7C-B206-741A8B3EF71A}')") , ], 1 , 2 , 4 , 0 , 776 , (3, 0, None, None) , 0 , )),
	(( 'PasteAndFormat' , 'Type' , ), 412, (412, (), [ (3, 1, None, None) , ], 1 , 1 , 4 , 0 , 780 , (3, 0, None, None) , 0 , )),
	(( 'PasteExcelTable' , 'LinkedToExcel' , 'WordFormatting' , 'RTF' , ), 413, (413, (), [ 
			 (11, 1, None, None) , (11, 1, None, None) , (11, 1, None, None) , ], 1 , 1 , 4 , 0 , 784 , (3, 0, None, None) , 0 , )),
	(( 'PasteAppendTable' , ), 414, (414, (), [ ], 1 , 1 , 4 , 0 , 788 , (3, 0, None, None) , 0 , )),
	(( 'XMLNodes' , 'prop' , ), 340, (340, (), [ (16393, 10, None, "IID('{D36C1F42-7044-4B9E-9CA3-85919454DB04}')") , ], 1 , 2 , 4 , 0 , 792 , (3, 0, None, None) , 0 , )),
	(( 'XMLParentNode' , 'prop' , ), 341, (341, (), [ (16393, 10, None, "IID('{09760240-0B89-49F7-A79D-479F24723F56}')") , ], 1 , 2 , 4 , 0 , 796 , (3, 0, None, None) , 0 , )),
	(( 'Editors' , 'prop' , ), 343, (343, (), [ (16393, 10, None, "IID('{AED7E08C-14F0-4F33-921D-4C5353137BF6}')") , ], 1 , 2 , 4 , 0 , 800 , (3, 0, None, None) , 0 , )),
	(( 'XML' , 'DataOnly' , 'prop' , ), 344, (344, (), [ (11, 49, 'False', None) , 
			 (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 804 , (3, 0, None, None) , 0 , )),
	(( 'XML' , 'DataOnly' , 'prop' , ), 344, (344, (), [ (11, 49, 'False', None) , 
			 (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 804 , (3, 0, None, None) , 0 , )),
	(( 'EnhMetaFileBits' , 'prop' , ), 345, (345, (), [ (16396, 10, None, None) , ], 1 , 2 , 4 , 0 , 808 , (3, 0, None, None) , 0 , )),
	(( 'GoToEditableRange' , 'EditorID' , 'prop' , ), 415, (415, (), [ (16396, 17, None, None) , 
			 (16393, 10, None, "IID('{0002095E-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 1 , 812 , (3, 0, None, None) , 0 , )),
	(( 'InsertXML' , 'XML' , 'Transform' , ), 416, (416, (), [ (8, 1, None, None) , 
			 (16396, 17, None, None) , ], 1 , 1 , 4 , 1 , 816 , (3, 0, None, None) , 0 , )),
	(( 'InsertCaption' , 'Label' , 'Title' , 'TitleAutoText' , 'Position' , 
			 'ExcludeLabel' , ), 417, (417, (), [ (16396, 1, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , ], 1 , 1 , 4 , 4 , 820 , (3, 0, None, None) , 0 , )),
	(( 'InsertCrossReference' , 'ReferenceType' , 'ReferenceKind' , 'ReferenceItem' , 'InsertAsHyperlink' , 
			 'IncludePosition' , 'SeparateNumbers' , 'SeparatorString' , ), 418, (418, (), [ (16396, 1, None, None) , 
			 (3, 1, None, None) , (16396, 1, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , ], 1 , 1 , 4 , 4 , 824 , (3, 0, None, None) , 0 , )),
	(( 'OMaths' , 'prop' , ), 346, (346, (), [ (16393, 10, None, "IID('{873E774B-926A-4CB1-878D-635A45187595}')") , ], 1 , 2 , 4 , 0 , 828 , (3, 0, None, None) , 0 , )),
	(( 'CharacterStyle' , 'prop' , ), 420, (420, (), [ (16396, 10, None, None) , ], 1 , 2 , 4 , 0 , 832 , (3, 0, None, None) , 0 , )),
	(( 'ParagraphStyle' , 'prop' , ), 421, (421, (), [ (16396, 10, None, None) , ], 1 , 2 , 4 , 0 , 836 , (3, 0, None, None) , 0 , )),
	(( 'ListStyle' , 'prop' , ), 422, (422, (), [ (16396, 10, None, None) , ], 1 , 2 , 4 , 0 , 840 , (3, 0, None, None) , 0 , )),
	(( 'TableStyle' , 'prop' , ), 423, (423, (), [ (16396, 10, None, None) , ], 1 , 2 , 4 , 0 , 844 , (3, 0, None, None) , 0 , )),
	(( 'ContentControls' , 'prop' , ), 424, (424, (), [ (16393, 10, None, "IID('{804CD967-F83B-432D-9446-C61A45CFEFF0}')") , ], 1 , 2 , 4 , 0 , 848 , (3, 0, None, None) , 0 , )),
	(( 'ExportFragment' , 'FileName' , 'Format' , ), 425, (425, (), [ (8, 1, None, None) , 
			 (3, 1, None, None) , ], 1 , 1 , 4 , 0 , 852 , (3, 0, None, None) , 0 , )),
	(( 'WordOpenXML' , 'prop' , ), 426, (426, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 856 , (3, 0, None, None) , 0 , )),
	(( 'SetListLevel' , 'Level' , ), 427, (427, (), [ (2, 1, None, None) , ], 1 , 1 , 4 , 0 , 860 , (3, 0, None, None) , 0 , )),
	(( 'InsertAlignmentTab' , 'Alignment' , 'RelativeTo' , ), 500, (500, (), [ (3, 1, None, None) , 
			 (3, 49, '0', None) , ], 1 , 1 , 4 , 0 , 864 , (3, 0, None, None) , 0 , )),
	(( 'ParentContentControl' , 'prop' , ), 501, (501, (), [ (16393, 10, None, "IID('{EE95AFE3-3026-4172-B078-0E79DAB5CC3D}')") , ], 1 , 2 , 4 , 0 , 868 , (3, 0, None, None) , 0 , )),
	(( 'ImportFragment' , 'FileName' , 'MatchDestination' , ), 502, (502, (), [ (8, 1, None, None) , 
			 (11, 49, 'False', None) , ], 1 , 1 , 4 , 0 , 872 , (3, 0, None, None) , 0 , )),
	(( 'ExportAsFixedFormat' , 'OutputFileName' , 'ExportFormat' , 'OpenAfterExport' , 'OptimizeFor' , 
			 'ExportCurrentPage' , 'Item' , 'IncludeDocProps' , 'KeepIRM' , 'CreateBookmarks' , 
			 'DocStructureTags' , 'BitmapMissingFonts' , 'UseISO19005_1' , 'FixedFormatExtClassPtr' , ), 503, (503, (), [ 
			 (8, 1, None, None) , (3, 1, None, None) , (11, 49, 'False', None) , (3, 49, '0', None) , (11, 49, 'False', None) , 
			 (3, 49, '0', None) , (11, 49, 'False', None) , (11, 49, 'True', None) , (3, 49, '0', None) , (11, 49, 'True', None) , 
			 (11, 49, 'True', None) , (11, 49, 'False', None) , (16396, 17, None, None) , ], 1 , 1 , 4 , 1 , 876 , (3, 0, None, None) , 0 , )),
	(( 'Locks' , 'prop' , ), 504, (504, (), [ (16393, 10, None, "IID('{DFF99AC2-CD2A-43AD-91B1-A2BE40BC7146}')") , ], 1 , 2 , 4 , 0 , 880 , (3, 0, None, None) , 0 , )),
	(( 'Updates' , 'prop' , ), 505, (505, (), [ (16393, 10, None, "IID('{30225CFC-5A71-4FE6-B527-90A52C54AE77}')") , ], 1 , 2 , 4 , 0 , 884 , (3, 0, None, None) , 0 , )),
	(( 'Conflicts' , 'prop' , ), 506, (506, (), [ (16393, 10, None, "IID('{C2B83A65-B061-4469-83B6-8877437CB8A0}')") , ], 1 , 2 , 4 , 0 , 888 , (3, 0, None, None) , 0 , )),
]

win32com.client.CLSIDToClass.RegisterCLSID( "{0002095E-0000-0000-C000-000000000046}", Range )
