# -*- coding: mbcs -*-
# Created by makepy.py version 0.5.01
# By python version 3.3.1 (v3.3.1:d9893d13c628, Apr  6 2013, 20:25:12) [MSC v.1600 32 bit (Intel)]
# From type library '{00020813-0000-0000-C000-000000000046}'
# On Sat Apr  5 18:14:59 2014
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
class _Application(DispatchBaseClass):
	CLSID = IID('{000208D5-0000-0000-C000-000000000046}')
	coclass_clsid = IID('{00024500-0000-0000-C000-000000000046}')

	def ActivateMicrosoftApp(self, Index=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(1095, LCID, 1, (24, 0), ((3, 1),),Index
			)

	def AddChartAutoFormat(self, Chart=defaultNamedNotOptArg, Name=defaultNamedNotOptArg, Description=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(216, LCID, 1, (24, 0), ((12, 1), (8, 1), (12, 17)),Chart
			, Name, Description)

	def AddCustomList(self, ListArray=defaultNamedNotOptArg, ByRow=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(780, LCID, 1, (24, 0), ((12, 1), (12, 17)),ListArray
			, ByRow)

	def Calculate(self):
		return self._oleobj_.InvokeTypes(279, LCID, 1, (24, 0), (),)

	def CalculateFull(self):
		return self._oleobj_.InvokeTypes(1805, LCID, 1, (24, 0), (),)

	def CalculateFullRebuild(self):
		return self._oleobj_.InvokeTypes(1945, LCID, 1, (24, 0), (),)

	def CalculateUntilAsyncQueriesDone(self):
		return self._oleobj_.InvokeTypes(2387, LCID, 1, (24, 0), (),)

	def CentimetersToPoints(self, Centimeters=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(1086, LCID, 1, (5, 0), ((5, 1),),Centimeters
			)

	def CheckAbort(self, KeepAbort=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1952, LCID, 1, (24, 0), ((12, 17),),KeepAbort
			)

	def CheckSpelling(self, Word=defaultNamedNotOptArg, CustomDictionary=defaultNamedOptArg, IgnoreUppercase=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(505, LCID, 1, (11, 0), ((8, 1), (12, 17), (12, 17)),Word
			, CustomDictionary, IgnoreUppercase)

	def ConvertFormula(self, Formula=defaultNamedNotOptArg, FromReferenceStyle=defaultNamedNotOptArg, ToReferenceStyle=defaultNamedOptArg, ToAbsolute=defaultNamedOptArg
			, RelativeTo=defaultNamedOptArg):
		return self._ApplyTypes_(325, 1, (12, 0), ((12, 1), (3, 1), (12, 17), (12, 17), (12, 17)), 'ConvertFormula', None,Formula
			, FromReferenceStyle, ToReferenceStyle, ToAbsolute, RelativeTo)

	def DDEExecute(self, Channel=defaultNamedNotOptArg, String=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(333, LCID, 1, (24, 0), ((3, 1), (8, 1)),Channel
			, String)

	def DDEInitiate(self, App=defaultNamedNotOptArg, Topic=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(334, LCID, 1, (3, 0), ((8, 1), (8, 1)),App
			, Topic)

	def DDEPoke(self, Channel=defaultNamedNotOptArg, Item=defaultNamedNotOptArg, Data=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(335, LCID, 1, (24, 0), ((3, 1), (12, 1), (12, 1)),Channel
			, Item, Data)

	def DDERequest(self, Channel=defaultNamedNotOptArg, Item=defaultNamedNotOptArg):
		return self._ApplyTypes_(336, 1, (12, 0), ((3, 1), (8, 1)), 'DDERequest', None,Channel
			, Item)

	def DDETerminate(self, Channel=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(337, LCID, 1, (24, 0), ((3, 1),),Channel
			)

	def DeleteChartAutoFormat(self, Name=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(217, LCID, 1, (24, 0), ((8, 1),),Name
			)

	def DeleteCustomList(self, ListNum=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(783, LCID, 1, (24, 0), ((3, 1),),ListNum
			)

	def DisplayXMLSourcePane(self, XmlMap=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2252, LCID, 1, (24, 0), ((12, 17),),XmlMap
			)

	def DoubleClick(self):
		return self._oleobj_.InvokeTypes(349, LCID, 1, (24, 0), (),)

	def Dummy1(self, Arg1=defaultNamedOptArg, Arg2=defaultNamedOptArg, Arg3=defaultNamedOptArg, Arg4=defaultNamedOptArg):
		return self._ApplyTypes_(1782, 1, (12, 0), ((12, 17), (12, 17), (12, 17), (12, 17)), 'Dummy1', None,Arg1
			, Arg2, Arg3, Arg4)

	def Dummy10(self, arg=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1791, LCID, 1, (11, 0), ((12, 17),),arg
			)

	def Dummy11(self):
		return self._oleobj_.InvokeTypes(1792, LCID, 1, (24, 0), (),)

	def Dummy12(self, p1=defaultNamedNotOptArg, p2=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(1803, LCID, 1, (24, 0), ((9, 1), (9, 1)),p1
			, p2)

	def Dummy13(self, Arg1=defaultNamedNotOptArg, Arg2=defaultNamedOptArg, Arg3=defaultNamedOptArg, Arg4=defaultNamedOptArg
			, Arg5=defaultNamedOptArg, Arg6=defaultNamedOptArg, Arg7=defaultNamedOptArg, Arg8=defaultNamedOptArg, Arg9=defaultNamedOptArg
			, Arg10=defaultNamedOptArg, Arg11=defaultNamedOptArg, Arg12=defaultNamedOptArg, Arg13=defaultNamedOptArg, Arg14=defaultNamedOptArg
			, Arg15=defaultNamedOptArg, Arg16=defaultNamedOptArg, Arg17=defaultNamedOptArg, Arg18=defaultNamedOptArg, Arg19=defaultNamedOptArg
			, Arg20=defaultNamedOptArg, Arg21=defaultNamedOptArg, Arg22=defaultNamedOptArg, Arg23=defaultNamedOptArg, Arg24=defaultNamedOptArg
			, Arg25=defaultNamedOptArg, Arg26=defaultNamedOptArg, Arg27=defaultNamedOptArg, Arg28=defaultNamedOptArg, Arg29=defaultNamedOptArg
			, Arg30=defaultNamedOptArg):
		return self._ApplyTypes_(1933, 1, (12, 0), ((12, 1), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)), 'Dummy13', None,Arg1
			, Arg2, Arg3, Arg4, Arg5, Arg6
			, Arg7, Arg8, Arg9, Arg10, Arg11
			, Arg12, Arg13, Arg14, Arg15, Arg16
			, Arg17, Arg18, Arg19, Arg20, Arg21
			, Arg22, Arg23, Arg24, Arg25, Arg26
			, Arg27, Arg28, Arg29, Arg30)

	def Dummy14(self):
		return self._oleobj_.InvokeTypes(1944, LCID, 1, (24, 0), (),)

	def Dummy2(self, Arg1=defaultNamedOptArg, Arg2=defaultNamedOptArg, Arg3=defaultNamedOptArg, Arg4=defaultNamedOptArg
			, Arg5=defaultNamedOptArg, Arg6=defaultNamedOptArg, Arg7=defaultNamedOptArg, Arg8=defaultNamedOptArg):
		return self._ApplyTypes_(1783, 1, (12, 0), ((12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)), 'Dummy2', None,Arg1
			, Arg2, Arg3, Arg4, Arg5, Arg6
			, Arg7, Arg8)

	def Dummy20(self, grfCompareFunctions=defaultNamedNotOptArg):
		return self._ApplyTypes_(2373, 1, (12, 0), ((3, 1),), 'Dummy20', None,grfCompareFunctions
			)

	def Dummy3(self):
		return self._ApplyTypes_(1784, 1, (12, 0), (), 'Dummy3', None,)

	def Dummy4(self, Arg1=defaultNamedOptArg, Arg2=defaultNamedOptArg, Arg3=defaultNamedOptArg, Arg4=defaultNamedOptArg
			, Arg5=defaultNamedOptArg, Arg6=defaultNamedOptArg, Arg7=defaultNamedOptArg, Arg8=defaultNamedOptArg, Arg9=defaultNamedOptArg
			, Arg10=defaultNamedOptArg, Arg11=defaultNamedOptArg, Arg12=defaultNamedOptArg, Arg13=defaultNamedOptArg, Arg14=defaultNamedOptArg
			, Arg15=defaultNamedOptArg):
		return self._ApplyTypes_(1785, 1, (12, 0), ((12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)), 'Dummy4', None,Arg1
			, Arg2, Arg3, Arg4, Arg5, Arg6
			, Arg7, Arg8, Arg9, Arg10, Arg11
			, Arg12, Arg13, Arg14, Arg15)

	def Dummy5(self, Arg1=defaultNamedOptArg, Arg2=defaultNamedOptArg, Arg3=defaultNamedOptArg, Arg4=defaultNamedOptArg
			, Arg5=defaultNamedOptArg, Arg6=defaultNamedOptArg, Arg7=defaultNamedOptArg, Arg8=defaultNamedOptArg, Arg9=defaultNamedOptArg
			, Arg10=defaultNamedOptArg, Arg11=defaultNamedOptArg, Arg12=defaultNamedOptArg, Arg13=defaultNamedOptArg):
		return self._ApplyTypes_(1786, 1, (12, 0), ((12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)), 'Dummy5', None,Arg1
			, Arg2, Arg3, Arg4, Arg5, Arg6
			, Arg7, Arg8, Arg9, Arg10, Arg11
			, Arg12, Arg13)

	def Dummy6(self):
		return self._ApplyTypes_(1787, 1, (12, 0), (), 'Dummy6', None,)

	def Dummy7(self):
		return self._ApplyTypes_(1788, 1, (12, 0), (), 'Dummy7', None,)

	def Dummy8(self, Arg1=defaultNamedOptArg):
		return self._ApplyTypes_(1789, 1, (12, 0), ((12, 17),), 'Dummy8', None,Arg1
			)

	def Dummy9(self):
		return self._ApplyTypes_(1790, 1, (12, 0), (), 'Dummy9', None,)

	def Evaluate(self, Name=defaultNamedNotOptArg):
		return self._ApplyTypes_(1, 1, (12, 0), ((12, 1),), 'Evaluate', None,Name
			)

	def ExecuteExcel4Macro(self, String=defaultNamedNotOptArg):
		return self._ApplyTypes_(350, 1, (12, 0), ((8, 1),), 'ExecuteExcel4Macro', None,String
			)

	# Result is of type FileDialog
	# The method FileDialog is actually a property, but must be used as a method to correctly pass the arguments
	def FileDialog(self, fileDialogType=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(1942, LCID, 2, (9, 0), ((3, 1),),fileDialogType
			)
		if ret is not None:
			ret = Dispatch(ret, 'FileDialog', '{000C0362-0000-0000-C000-000000000046}')
		return ret

	def FindFile(self):
		return self._oleobj_.InvokeTypes(1771, LCID, 1, (11, 0), (),)

	# The method GetCaller is actually a property, but must be used as a method to correctly pass the arguments
	def GetCaller(self, Index=defaultNamedOptArg):
		return self._ApplyTypes_(317, 2, (12, 0), ((12, 17),), 'GetCaller', None,Index
			)

	# The method GetClipboardFormats is actually a property, but must be used as a method to correctly pass the arguments
	def GetClipboardFormats(self, Index=defaultNamedOptArg):
		return self._ApplyTypes_(321, 2, (12, 0), ((12, 17),), 'GetClipboardFormats', None,Index
			)

	def GetCustomListContents(self, ListNum=defaultNamedNotOptArg):
		return self._ApplyTypes_(786, 1, (12, 0), ((3, 1),), 'GetCustomListContents', None,ListNum
			)

	def GetCustomListNum(self, ListArray=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(785, LCID, 1, (3, 0), ((12, 1),),ListArray
			)

	# The method GetFileConverters is actually a property, but must be used as a method to correctly pass the arguments
	def GetFileConverters(self, Index1=defaultNamedOptArg, Index2=defaultNamedOptArg):
		return self._ApplyTypes_(931, 2, (12, 0), ((12, 17), (12, 17)), 'GetFileConverters', None,Index1
			, Index2)

	# The method GetInternational is actually a property, but must be used as a method to correctly pass the arguments
	def GetInternational(self, Index=defaultNamedOptArg):
		return self._ApplyTypes_(362, 2, (12, 0), ((12, 17),), 'GetInternational', None,Index
			)

	def GetOpenFilename(self, FileFilter=defaultNamedOptArg, FilterIndex=defaultNamedOptArg, Title=defaultNamedOptArg, ButtonText=defaultNamedOptArg
			, MultiSelect=defaultNamedOptArg):
		return self._ApplyTypes_(1075, 1, (12, 0), ((12, 17), (12, 17), (12, 17), (12, 17), (12, 17)), 'GetOpenFilename', None,FileFilter
			, FilterIndex, Title, ButtonText, MultiSelect)

	def GetPhonetic(self, Text=defaultNamedOptArg):
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(1795, LCID, 1, (8, 0), ((12, 17),),Text
			)

	# The method GetPreviousSelections is actually a property, but must be used as a method to correctly pass the arguments
	def GetPreviousSelections(self, Index=defaultNamedOptArg):
		return self._ApplyTypes_(378, 2, (12, 0), ((12, 17),), 'GetPreviousSelections', None,Index
			)

	# The method GetRegisteredFunctions is actually a property, but must be used as a method to correctly pass the arguments
	def GetRegisteredFunctions(self, Index1=defaultNamedOptArg, Index2=defaultNamedOptArg):
		return self._ApplyTypes_(775, 2, (12, 0), ((12, 17), (12, 17)), 'GetRegisteredFunctions', None,Index1
			, Index2)

	def GetSaveAsFilename(self, InitialFilename=defaultNamedOptArg, FileFilter=defaultNamedOptArg, FilterIndex=defaultNamedOptArg, Title=defaultNamedOptArg
			, ButtonText=defaultNamedOptArg):
		return self._ApplyTypes_(1076, 1, (12, 0), ((12, 17), (12, 17), (12, 17), (12, 17), (12, 17)), 'GetSaveAsFilename', None,InitialFilename
			, FileFilter, FilterIndex, Title, ButtonText)

	def Goto(self, Reference=defaultNamedOptArg, Scroll=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(475, LCID, 1, (24, 0), ((12, 17), (12, 17)),Reference
			, Scroll)

	def Help(self, HelpFile=defaultNamedOptArg, HelpContextID=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(354, LCID, 1, (24, 0), ((12, 17), (12, 17)),HelpFile
			, HelpContextID)

	def InchesToPoints(self, Inches=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(1087, LCID, 1, (5, 0), ((5, 1),),Inches
			)

	def InputBox(self, Prompt=defaultNamedNotOptArg, Title=defaultNamedOptArg, Default=defaultNamedOptArg, Left=defaultNamedOptArg
			, Top=defaultNamedOptArg, HelpFile=defaultNamedOptArg, HelpContextID=defaultNamedOptArg, Type=defaultNamedOptArg):
		return self._ApplyTypes_(357, 1, (12, 0), ((8, 1), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)), 'InputBox', None,Prompt
			, Title, Default, Left, Top, HelpFile
			, HelpContextID, Type)

	# Result is of type Range
	def Intersect(self, Arg1=defaultNamedNotOptArg, Arg2=defaultNamedNotOptArg, Arg3=defaultNamedOptArg, Arg4=defaultNamedOptArg
			, Arg5=defaultNamedOptArg, Arg6=defaultNamedOptArg, Arg7=defaultNamedOptArg, Arg8=defaultNamedOptArg, Arg9=defaultNamedOptArg
			, Arg10=defaultNamedOptArg, Arg11=defaultNamedOptArg, Arg12=defaultNamedOptArg, Arg13=defaultNamedOptArg, Arg14=defaultNamedOptArg
			, Arg15=defaultNamedOptArg, Arg16=defaultNamedOptArg, Arg17=defaultNamedOptArg, Arg18=defaultNamedOptArg, Arg19=defaultNamedOptArg
			, Arg20=defaultNamedOptArg, Arg21=defaultNamedOptArg, Arg22=defaultNamedOptArg, Arg23=defaultNamedOptArg, Arg24=defaultNamedOptArg
			, Arg25=defaultNamedOptArg, Arg26=defaultNamedOptArg, Arg27=defaultNamedOptArg, Arg28=defaultNamedOptArg, Arg29=defaultNamedOptArg
			, Arg30=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(766, LCID, 1, (9, 0), ((9, 1), (9, 1), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)),Arg1
			, Arg2, Arg3, Arg4, Arg5, Arg6
			, Arg7, Arg8, Arg9, Arg10, Arg11
			, Arg12, Arg13, Arg14, Arg15, Arg16
			, Arg17, Arg18, Arg19, Arg20, Arg21
			, Arg22, Arg23, Arg24, Arg25, Arg26
			, Arg27, Arg28, Arg29, Arg30)
		if ret is not None:
			ret = Dispatch(ret, 'Intersect', '{00020846-0000-0000-C000-000000000046}')
		return ret

	def MacroOptions(self, Macro=defaultNamedOptArg, Description=defaultNamedOptArg, HasMenu=defaultNamedOptArg, MenuText=defaultNamedOptArg
			, HasShortcutKey=defaultNamedOptArg, ShortcutKey=defaultNamedOptArg, Category=defaultNamedOptArg, StatusBar=defaultNamedOptArg, HelpContextID=defaultNamedOptArg
			, HelpFile=defaultNamedOptArg, ArgumentDescriptions=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2770, LCID, 1, (24, 0), ((12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)),Macro
			, Description, HasMenu, MenuText, HasShortcutKey, ShortcutKey
			, Category, StatusBar, HelpContextID, HelpFile, ArgumentDescriptions
			)

	def MailLogoff(self):
		return self._oleobj_.InvokeTypes(945, LCID, 1, (24, 0), (),)

	def MailLogon(self, Name=defaultNamedOptArg, Password=defaultNamedOptArg, DownloadNewMail=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(943, LCID, 1, (24, 0), ((12, 17), (12, 17), (12, 17)),Name
			, Password, DownloadNewMail)

	# Result is of type Workbook
	def NextLetter(self):
		ret = self._oleobj_.InvokeTypes(972, LCID, 1, (13, 0), (),)
		if ret is not None:
			# See if this IUnknown is really an IDispatch
			try:
				ret = ret.QueryInterface(pythoncom.IID_IDispatch)
			except pythoncom.error:
				return ret
			ret = Dispatch(ret, 'NextLetter', '{00020819-0000-0000-C000-000000000046}')
		return ret

	def OnKey(self, Key=defaultNamedNotOptArg, Procedure=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(626, LCID, 1, (24, 0), ((8, 1), (12, 17)),Key
			, Procedure)

	def OnRepeat(self, Text=defaultNamedNotOptArg, Procedure=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(769, LCID, 1, (24, 0), ((8, 1), (8, 1)),Text
			, Procedure)

	def OnTime(self, EarliestTime=defaultNamedNotOptArg, Procedure=defaultNamedNotOptArg, LatestTime=defaultNamedOptArg, Schedule=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(624, LCID, 1, (24, 0), ((12, 1), (8, 1), (12, 17), (12, 17)),EarliestTime
			, Procedure, LatestTime, Schedule)

	def OnUndo(self, Text=defaultNamedNotOptArg, Procedure=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(770, LCID, 1, (24, 0), ((8, 1), (8, 1)),Text
			, Procedure)

	def Quit(self):
		return self._oleobj_.InvokeTypes(302, LCID, 1, (24, 0), (),)

	# Result is of type Range
	# The method Range is actually a property, but must be used as a method to correctly pass the arguments
	def Range(self, Cell1=defaultNamedNotOptArg, Cell2=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(197, LCID, 2, (9, 0), ((12, 1), (12, 17)),Cell1
			, Cell2)
		if ret is not None:
			ret = Dispatch(ret, 'Range', '{00020846-0000-0000-C000-000000000046}')
		return ret

	def RecordMacro(self, BasicCode=defaultNamedOptArg, XlmCode=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(773, LCID, 1, (24, 0), ((12, 17), (12, 17)),BasicCode
			, XlmCode)

	def RegisterXLL(self, Filename=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(30, LCID, 1, (11, 0), ((8, 1),),Filename
			)

	def Repeat(self):
		return self._oleobj_.InvokeTypes(301, LCID, 1, (24, 0), (),)

	def ResetTipWizard(self):
		return self._oleobj_.InvokeTypes(928, LCID, 1, (24, 0), (),)

	def Run(self, Macro=defaultNamedOptArg, Arg1=defaultNamedOptArg, Arg2=defaultNamedOptArg, Arg3=defaultNamedOptArg
			, Arg4=defaultNamedOptArg, Arg5=defaultNamedOptArg, Arg6=defaultNamedOptArg, Arg7=defaultNamedOptArg, Arg8=defaultNamedOptArg
			, Arg9=defaultNamedOptArg, Arg10=defaultNamedOptArg, Arg11=defaultNamedOptArg, Arg12=defaultNamedOptArg, Arg13=defaultNamedOptArg
			, Arg14=defaultNamedOptArg, Arg15=defaultNamedOptArg, Arg16=defaultNamedOptArg, Arg17=defaultNamedOptArg, Arg18=defaultNamedOptArg
			, Arg19=defaultNamedOptArg, Arg20=defaultNamedOptArg, Arg21=defaultNamedOptArg, Arg22=defaultNamedOptArg, Arg23=defaultNamedOptArg
			, Arg24=defaultNamedOptArg, Arg25=defaultNamedOptArg, Arg26=defaultNamedOptArg, Arg27=defaultNamedOptArg, Arg28=defaultNamedOptArg
			, Arg29=defaultNamedOptArg, Arg30=defaultNamedOptArg):
		return self._ApplyTypes_(259, 1, (12, 0), ((12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)), 'Run', None,Macro
			, Arg1, Arg2, Arg3, Arg4, Arg5
			, Arg6, Arg7, Arg8, Arg9, Arg10
			, Arg11, Arg12, Arg13, Arg14, Arg15
			, Arg16, Arg17, Arg18, Arg19, Arg20
			, Arg21, Arg22, Arg23, Arg24, Arg25
			, Arg26, Arg27, Arg28, Arg29, Arg30
			)

	def Save(self, Filename=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(283, LCID, 1, (24, 0), ((12, 17),),Filename
			)

	def SaveWorkspace(self, Filename=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(212, LCID, 1, (24, 0), ((12, 17),),Filename
			)

	def SendKeys(self, Keys=defaultNamedNotOptArg, Wait=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(383, LCID, 1, (24, 0), ((12, 1), (12, 17)),Keys
			, Wait)

	def SetDefaultChart(self, FormatName=defaultNamedOptArg, Gallery=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(219, LCID, 1, (24, 0), ((12, 17), (12, 17)),FormatName
			, Gallery)

	def SharePointVersion(self, bstrUrl=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(2392, LCID, 1, (3, 0), ((8, 1),),bstrUrl
			)

	# Result is of type Menu
	# The method ShortcutMenus is actually a property, but must be used as a method to correctly pass the arguments
	def ShortcutMenus(self, Index=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(776, LCID, 2, (9, 0), ((3, 1),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'ShortcutMenus', '{00020866-0000-0000-C000-000000000046}')
		return ret

	def Support(self, Object=defaultNamedNotOptArg, ID=defaultNamedNotOptArg, arg=defaultNamedOptArg):
		return self._ApplyTypes_(2255, 1, (12, 0), ((9, 1), (3, 1), (12, 17)), 'Support', None,Object
			, ID, arg)

	def Undo(self):
		return self._oleobj_.InvokeTypes(303, LCID, 1, (24, 0), (),)

	# Result is of type Range
	def Union(self, Arg1=defaultNamedNotOptArg, Arg2=defaultNamedNotOptArg, Arg3=defaultNamedOptArg, Arg4=defaultNamedOptArg
			, Arg5=defaultNamedOptArg, Arg6=defaultNamedOptArg, Arg7=defaultNamedOptArg, Arg8=defaultNamedOptArg, Arg9=defaultNamedOptArg
			, Arg10=defaultNamedOptArg, Arg11=defaultNamedOptArg, Arg12=defaultNamedOptArg, Arg13=defaultNamedOptArg, Arg14=defaultNamedOptArg
			, Arg15=defaultNamedOptArg, Arg16=defaultNamedOptArg, Arg17=defaultNamedOptArg, Arg18=defaultNamedOptArg, Arg19=defaultNamedOptArg
			, Arg20=defaultNamedOptArg, Arg21=defaultNamedOptArg, Arg22=defaultNamedOptArg, Arg23=defaultNamedOptArg, Arg24=defaultNamedOptArg
			, Arg25=defaultNamedOptArg, Arg26=defaultNamedOptArg, Arg27=defaultNamedOptArg, Arg28=defaultNamedOptArg, Arg29=defaultNamedOptArg
			, Arg30=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(779, LCID, 1, (9, 0), ((9, 1), (9, 1), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)),Arg1
			, Arg2, Arg3, Arg4, Arg5, Arg6
			, Arg7, Arg8, Arg9, Arg10, Arg11
			, Arg12, Arg13, Arg14, Arg15, Arg16
			, Arg17, Arg18, Arg19, Arg20, Arg21
			, Arg22, Arg23, Arg24, Arg25, Arg26
			, Arg27, Arg28, Arg29, Arg30)
		if ret is not None:
			ret = Dispatch(ret, 'Union', '{00020846-0000-0000-C000-000000000046}')
		return ret

	def Volatile(self, Volatile=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(788, LCID, 1, (24, 0), ((12, 17),),Volatile
			)

	def Wait(self, Time=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(1770, LCID, 1, (11, 0), ((12, 1),),Time
			)

	def _Evaluate(self, Name=defaultNamedNotOptArg):
		return self._ApplyTypes_(-5, 1, (12, 0), ((12, 1),), '_Evaluate', None,Name
			)

	def _FindFile(self):
		return self._oleobj_.InvokeTypes(1068, LCID, 1, (24, 0), (),)

	def _MacroOptions(self, Macro=defaultNamedOptArg, Description=defaultNamedOptArg, HasMenu=defaultNamedOptArg, MenuText=defaultNamedOptArg
			, HasShortcutKey=defaultNamedOptArg, ShortcutKey=defaultNamedOptArg, Category=defaultNamedOptArg, StatusBar=defaultNamedOptArg, HelpContextID=defaultNamedOptArg
			, HelpFile=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1135, LCID, 1, (24, 0), ((12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)),Macro
			, Description, HasMenu, MenuText, HasShortcutKey, ShortcutKey
			, Category, StatusBar, HelpContextID, HelpFile)

	def _Run2(self, Macro=defaultNamedOptArg, Arg1=defaultNamedOptArg, Arg2=defaultNamedOptArg, Arg3=defaultNamedOptArg
			, Arg4=defaultNamedOptArg, Arg5=defaultNamedOptArg, Arg6=defaultNamedOptArg, Arg7=defaultNamedOptArg, Arg8=defaultNamedOptArg
			, Arg9=defaultNamedOptArg, Arg10=defaultNamedOptArg, Arg11=defaultNamedOptArg, Arg12=defaultNamedOptArg, Arg13=defaultNamedOptArg
			, Arg14=defaultNamedOptArg, Arg15=defaultNamedOptArg, Arg16=defaultNamedOptArg, Arg17=defaultNamedOptArg, Arg18=defaultNamedOptArg
			, Arg19=defaultNamedOptArg, Arg20=defaultNamedOptArg, Arg21=defaultNamedOptArg, Arg22=defaultNamedOptArg, Arg23=defaultNamedOptArg
			, Arg24=defaultNamedOptArg, Arg25=defaultNamedOptArg, Arg26=defaultNamedOptArg, Arg27=defaultNamedOptArg, Arg28=defaultNamedOptArg
			, Arg29=defaultNamedOptArg, Arg30=defaultNamedOptArg):
		return self._ApplyTypes_(806, 1, (12, 0), ((12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)), '_Run2', None,Macro
			, Arg1, Arg2, Arg3, Arg4, Arg5
			, Arg6, Arg7, Arg8, Arg9, Arg10
			, Arg11, Arg12, Arg13, Arg14, Arg15
			, Arg16, Arg17, Arg18, Arg19, Arg20
			, Arg21, Arg22, Arg23, Arg24, Arg25
			, Arg26, Arg27, Arg28, Arg29, Arg30
			)

	def _WSFunction(self, Arg1=defaultNamedOptArg, Arg2=defaultNamedOptArg, Arg3=defaultNamedOptArg, Arg4=defaultNamedOptArg
			, Arg5=defaultNamedOptArg, Arg6=defaultNamedOptArg, Arg7=defaultNamedOptArg, Arg8=defaultNamedOptArg, Arg9=defaultNamedOptArg
			, Arg10=defaultNamedOptArg, Arg11=defaultNamedOptArg, Arg12=defaultNamedOptArg, Arg13=defaultNamedOptArg, Arg14=defaultNamedOptArg
			, Arg15=defaultNamedOptArg, Arg16=defaultNamedOptArg, Arg17=defaultNamedOptArg, Arg18=defaultNamedOptArg, Arg19=defaultNamedOptArg
			, Arg20=defaultNamedOptArg, Arg21=defaultNamedOptArg, Arg22=defaultNamedOptArg, Arg23=defaultNamedOptArg, Arg24=defaultNamedOptArg
			, Arg25=defaultNamedOptArg, Arg26=defaultNamedOptArg, Arg27=defaultNamedOptArg, Arg28=defaultNamedOptArg, Arg29=defaultNamedOptArg
			, Arg30=defaultNamedOptArg):
		return self._ApplyTypes_(169, 1, (12, 0), ((12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)), '_WSFunction', None,Arg1
			, Arg2, Arg3, Arg4, Arg5, Arg6
			, Arg7, Arg8, Arg9, Arg10, Arg11
			, Arg12, Arg13, Arg14, Arg15, Arg16
			, Arg17, Arg18, Arg19, Arg20, Arg21
			, Arg22, Arg23, Arg24, Arg25, Arg26
			, Arg27, Arg28, Arg29, Arg30)

	def _Wait(self, Time=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(393, LCID, 1, (24, 0), ((12, 1),),Time
			)

	_prop_map_get_ = {
		# Method 'ActiveCell' returns object of type 'Range'
		"ActiveCell": (305, 2, (9, 0), (), "ActiveCell", '{00020846-0000-0000-C000-000000000046}'),
		# Method 'ActiveChart' returns object of type 'Chart'
		"ActiveChart": (183, 2, (13, 0), (), "ActiveChart", '{00020821-0000-0000-C000-000000000046}'),
		# Method 'ActiveDialog' returns object of type 'DialogSheet'
		"ActiveDialog": (815, 2, (9, 0), (), "ActiveDialog", '{000208AF-0000-0000-C000-000000000046}'),
		"ActiveEncryptionSession": (2394, 2, (3, 0), (), "ActiveEncryptionSession", None),
		# Method 'ActiveMenuBar' returns object of type 'MenuBar'
		"ActiveMenuBar": (758, 2, (9, 0), (), "ActiveMenuBar", '{00020864-0000-0000-C000-000000000046}'),
		"ActivePrinter": (306, 2, (8, 0), (), "ActivePrinter", None),
		# Method 'ActiveProtectedViewWindow' returns object of type 'ProtectedViewWindow'
		"ActiveProtectedViewWindow": (2784, 2, (9, 0), (), "ActiveProtectedViewWindow", '{000244CD-0000-0000-C000-000000000046}'),
		"ActiveSheet": (307, 2, (9, 0), (), "ActiveSheet", None),
		# Method 'ActiveWindow' returns object of type 'Window'
		"ActiveWindow": (759, 2, (9, 0), (), "ActiveWindow", '{00020893-0000-0000-C000-000000000046}'),
		# Method 'ActiveWorkbook' returns object of type 'Workbook'
		"ActiveWorkbook": (308, 2, (13, 0), (), "ActiveWorkbook", '{00020819-0000-0000-C000-000000000046}'),
		# Method 'AddIns' returns object of type 'AddIns'
		"AddIns": (549, 2, (9, 0), (), "AddIns", '{00020858-0000-0000-C000-000000000046}'),
		# Method 'AddIns2' returns object of type 'AddIns2'
		"AddIns2": (2775, 2, (9, 0), (), "AddIns2", '{000244B5-0000-0000-C000-000000000046}'),
		"AlertBeforeOverwriting": (930, 2, (11, 0), (), "AlertBeforeOverwriting", None),
		"AltStartupPath": (313, 2, (8, 0), (), "AltStartupPath", None),
		"AlwaysUseClearType": (2381, 2, (11, 0), (), "AlwaysUseClearType", None),
		# Method 'AnswerWizard' returns object of type 'AnswerWizard'
		"AnswerWizard": (1804, 2, (9, 0), (), "AnswerWizard", '{000C0360-0000-0000-C000-000000000046}'),
		# Method 'Application' returns object of type 'Application'
		"Application": (148, 2, (13, 0), (), "Application", '{00024500-0000-0000-C000-000000000046}'),
		"ArbitraryXMLSupportAvailable": (2254, 2, (11, 0), (), "ArbitraryXMLSupportAvailable", None),
		"AskToUpdateLinks": (992, 2, (11, 0), (), "AskToUpdateLinks", None),
		# Method 'Assistance' returns object of type 'IAssistance'
		"Assistance": (2386, 2, (9, 0), (), "Assistance", '{4291224C-DEFE-485B-8E69-6CF8AA85CB76}'),
		# Method 'Assistant' returns object of type 'Assistant'
		"Assistant": (1438, 2, (9, 0), (), "Assistant", '{000C0322-0000-0000-C000-000000000046}'),
		# Method 'AutoCorrect' returns object of type 'AutoCorrect'
		"AutoCorrect": (1145, 2, (9, 0), (), "AutoCorrect", '{000208D4-0000-0000-C000-000000000046}'),
		"AutoFormatAsYouTypeReplaceHyperlinks": (1955, 2, (11, 0), (), "AutoFormatAsYouTypeReplaceHyperlinks", None),
		"AutoPercentEntry": (1800, 2, (11, 0), (), "AutoPercentEntry", None),
		# Method 'AutoRecover' returns object of type 'AutoRecover'
		"AutoRecover": (1949, 2, (9, 0), (), "AutoRecover", '{0002445A-0000-0000-C000-000000000046}'),
		"AutomationSecurity": (1941, 2, (3, 0), (), "AutomationSecurity", None),
		"Build": (314, 2, (3, 0), (), "Build", None),
		# Method 'COMAddIns' returns object of type 'COMAddIns'
		"COMAddIns": (1796, 2, (9, 0), (), "COMAddIns", '{000C0339-0000-0000-C000-000000000046}'),
		"CalculateBeforeSave": (315, 2, (11, 0), (), "CalculateBeforeSave", None),
		"Calculation": (316, 2, (3, 0), (), "Calculation", None),
		"CalculationInterruptKey": (1938, 2, (3, 0), (), "CalculationInterruptKey", None),
		"CalculationState": (1937, 2, (3, 0), (), "CalculationState", None),
		"CalculationVersion": (1806, 2, (3, 0), (), "CalculationVersion", None),
		"Caller": (317, 2, (12, 0), ((12, 17),), "Caller", None),
		"CanPlaySounds": (318, 2, (11, 0), (), "CanPlaySounds", None),
		"CanRecordSounds": (319, 2, (11, 0), (), "CanRecordSounds", None),
		"Caption": (139, 2, (8, 0), (), "Caption", None),
		"CellDragAndDrop": (320, 2, (11, 0), (), "CellDragAndDrop", None),
		# Method 'Cells' returns object of type 'Range'
		"Cells": (238, 2, (9, 0), (), "Cells", '{00020846-0000-0000-C000-000000000046}'),
		# Method 'Charts' returns object of type 'Sheets'
		"Charts": (121, 2, (9, 0), (), "Charts", '{000208D7-0000-0000-C000-000000000046}'),
		"ClipboardFormats": (321, 2, (12, 0), ((12, 17),), "ClipboardFormats", None),
		"ClusterConnector": (2779, 2, (8, 0), (), "ClusterConnector", None),
		"ColorButtons": (365, 2, (11, 0), (), "ColorButtons", None),
		# Method 'Columns' returns object of type 'Range'
		"Columns": (241, 2, (9, 0), (), "Columns", '{00020846-0000-0000-C000-000000000046}'),
		# Method 'CommandBars' returns object of type 'CommandBars'
		"CommandBars": (1439, 2, (13, 0), (), "CommandBars", '{55F88893-7708-11D1-ACEB-006008961DA5}'),
		"CommandUnderlines": (323, 2, (3, 0), (), "CommandUnderlines", None),
		"ConstrainNumeric": (324, 2, (11, 0), (), "ConstrainNumeric", None),
		"ControlCharacters": (233, 2, (11, 0), (), "ControlCharacters", None),
		"CopyObjectsWithCells": (991, 2, (11, 0), (), "CopyObjectsWithCells", None),
		"Creator": (149, 2, (3, 0), (), "Creator", None),
		"Cursor": (1161, 2, (3, 0), (), "Cursor", None),
		"CursorMovement": (232, 2, (3, 0), (), "CursorMovement", None),
		"CustomListCount": (787, 2, (3, 0), (), "CustomListCount", None),
		"CutCopyMode": (330, 2, (3, 0), (), "CutCopyMode", None),
		"DDEAppReturnCode": (332, 2, (3, 0), (), "DDEAppReturnCode", None),
		"DataEntryMode": (331, 2, (3, 0), (), "DataEntryMode", None),
		"DecimalSeparator": (1809, 2, (8, 0), (), "DecimalSeparator", None),
		"DefaultFilePath": (1038, 2, (8, 0), (), "DefaultFilePath", None),
		"DefaultSaveFormat": (1209, 2, (3, 0), (), "DefaultSaveFormat", None),
		"DefaultSheetDirection": (229, 2, (3, 0), (), "DefaultSheetDirection", None),
		# Method 'DefaultWebOptions' returns object of type 'DefaultWebOptions'
		"DefaultWebOptions": (1797, 2, (9, 0), (), "DefaultWebOptions", '{00024448-0000-0000-C000-000000000046}'),
		"DeferAsyncQueries": (2390, 2, (11, 0), (), "DeferAsyncQueries", None),
		# Method 'DialogSheets' returns object of type 'Sheets'
		"DialogSheets": (764, 2, (9, 0), (), "DialogSheets", '{000208D7-0000-0000-C000-000000000046}'),
		# Method 'Dialogs' returns object of type 'Dialogs'
		"Dialogs": (761, 2, (9, 0), (), "Dialogs", '{00020879-0000-0000-C000-000000000046}'),
		"DisplayAlerts": (343, 2, (11, 0), (), "DisplayAlerts", None),
		"DisplayClipboardWindow": (322, 2, (11, 0), (), "DisplayClipboardWindow", None),
		"DisplayCommentIndicator": (1196, 2, (3, 0), (), "DisplayCommentIndicator", None),
		"DisplayDocumentActionTaskPane": (2251, 2, (11, 0), (), "DisplayDocumentActionTaskPane", None),
		"DisplayDocumentInformationPanel": (2380, 2, (11, 0), (), "DisplayDocumentInformationPanel", None),
		"DisplayExcel4Menus": (927, 2, (11, 0), (), "DisplayExcel4Menus", None),
		"DisplayFormulaAutoComplete": (2384, 2, (11, 0), (), "DisplayFormulaAutoComplete", None),
		"DisplayFormulaBar": (344, 2, (11, 0), (), "DisplayFormulaBar", None),
		"DisplayFullScreen": (1061, 2, (11, 0), (), "DisplayFullScreen", None),
		"DisplayFunctionToolTips": (1940, 2, (11, 0), (), "DisplayFunctionToolTips", None),
		"DisplayInfoWindow": (1213, 2, (11, 0), (), "DisplayInfoWindow", None),
		"DisplayInsertOptions": (1947, 2, (11, 0), (), "DisplayInsertOptions", None),
		"DisplayNoteIndicator": (345, 2, (11, 0), (), "DisplayNoteIndicator", None),
		"DisplayPasteOptions": (1946, 2, (11, 0), (), "DisplayPasteOptions", None),
		"DisplayRecentFiles": (926, 2, (11, 0), (), "DisplayRecentFiles", None),
		"DisplayScrollBars": (346, 2, (11, 0), (), "DisplayScrollBars", None),
		"DisplayStatusBar": (347, 2, (11, 0), (), "DisplayStatusBar", None),
		"Dummy101": (1802, 2, (9, 0), (), "Dummy101", None),
		"Dummy22": (2781, 2, (11, 0), (), "Dummy22", None),
		"Dummy23": (2782, 2, (11, 0), (), "Dummy23", None),
		"EditDirectlyInCell": (929, 2, (11, 0), (), "EditDirectlyInCell", None),
		"EnableAnimations": (1180, 2, (11, 0), (), "EnableAnimations", None),
		"EnableAutoComplete": (1179, 2, (11, 0), (), "EnableAutoComplete", None),
		"EnableCancelKey": (1096, 2, (3, 0), (), "EnableCancelKey", None),
		"EnableEvents": (1212, 2, (11, 0), (), "EnableEvents", None),
		"EnableLargeOperationAlert": (2388, 2, (11, 0), (), "EnableLargeOperationAlert", None),
		"EnableLivePreview": (2379, 2, (11, 0), (), "EnableLivePreview", None),
		"EnableSound": (1197, 2, (11, 0), (), "EnableSound", None),
		"EnableTipWizard": (1064, 2, (11, 0), (), "EnableTipWizard", None),
		# Method 'ErrorCheckingOptions' returns object of type 'ErrorCheckingOptions'
		"ErrorCheckingOptions": (1954, 2, (9, 0), (), "ErrorCheckingOptions", '{0002445B-0000-0000-C000-000000000046}'),
		# Method 'Excel4IntlMacroSheets' returns object of type 'Sheets'
		"Excel4IntlMacroSheets": (581, 2, (9, 0), (), "Excel4IntlMacroSheets", '{000208D7-0000-0000-C000-000000000046}'),
		# Method 'Excel4MacroSheets' returns object of type 'Sheets'
		"Excel4MacroSheets": (579, 2, (9, 0), (), "Excel4MacroSheets", '{000208D7-0000-0000-C000-000000000046}'),
		"ExtendList": (1793, 2, (11, 0), (), "ExtendList", None),
		"FeatureInstall": (1808, 2, (3, 0), (), "FeatureInstall", None),
		"FileConverters": (931, 2, (12, 0), ((12, 17), (12, 17)), "FileConverters", None),
		# Method 'FileExportConverters' returns object of type 'FileExportConverters'
		"FileExportConverters": (2768, 2, (9, 0), (), "FileExportConverters", '{000244B4-0000-0000-C000-000000000046}'),
		# Method 'FileFind' returns object of type 'IFind'
		"FileFind": (1201, 2, (9, 0), (), "FileFind", '{000C0337-0000-0000-C000-000000000046}'),
		# Method 'FileSearch' returns object of type 'FileSearch'
		"FileSearch": (1200, 2, (9, 0), (), "FileSearch", '{000C0332-0000-0000-C000-000000000046}'),
		"FileValidation": (2788, 2, (3, 0), (), "FileValidation", None),
		"FileValidationPivot": (2789, 2, (3, 0), (), "FileValidationPivot", None),
		# Method 'FindFormat' returns object of type 'CellFormat'
		"FindFormat": (1934, 2, (9, 0), (), "FindFormat", '{00024450-0000-0000-C000-000000000046}'),
		"FixedDecimal": (351, 2, (11, 0), (), "FixedDecimal", None),
		"FixedDecimalPlaces": (352, 2, (3, 0), (), "FixedDecimalPlaces", None),
		"FormulaBarHeight": (2383, 2, (3, 0), (), "FormulaBarHeight", None),
		"GenerateGetPivotData": (1948, 2, (11, 0), (), "GenerateGetPivotData", None),
		"GenerateTableRefs": (2385, 2, (3, 0), (), "GenerateTableRefs", None),
		"Height": (123, 2, (5, 0), (), "Height", None),
		"HighQualityModeForGraphics": (2395, 2, (11, 0), (), "HighQualityModeForGraphics", None),
		"Hinstance": (1951, 2, (3, 0), (), "Hinstance", None),
		"HinstancePtr": (2787, 2, (12, 0), (), "HinstancePtr", None),
		"Hwnd": (1950, 2, (3, 0), (), "Hwnd", None),
		"IgnoreRemoteRequests": (356, 2, (11, 0), (), "IgnoreRemoteRequests", None),
		"Interactive": (361, 2, (11, 0), (), "Interactive", None),
		"International": (362, 2, (12, 0), ((12, 17),), "International", None),
		"IsSandboxed": (2785, 2, (11, 0), (), "IsSandboxed", None),
		"Iteration": (363, 2, (11, 0), (), "Iteration", None),
		# Method 'LanguageSettings' returns object of type 'LanguageSettings'
		"LanguageSettings": (1801, 2, (9, 0), (), "LanguageSettings", '{000C0353-0000-0000-C000-000000000046}'),
		"LargeButtons": (364, 2, (11, 0), (), "LargeButtons", None),
		"LargeOperationCellThousandCount": (2389, 2, (3, 0), (), "LargeOperationCellThousandCount", None),
		"Left": (127, 2, (5, 0), (), "Left", None),
		"LibraryPath": (366, 2, (8, 0), (), "LibraryPath", None),
		"MailSession": (942, 2, (12, 0), (), "MailSession", None),
		"MailSystem": (971, 2, (3, 0), (), "MailSystem", None),
		"MapPaperSize": (1959, 2, (11, 0), (), "MapPaperSize", None),
		"MathCoprocessorAvailable": (367, 2, (11, 0), (), "MathCoprocessorAvailable", None),
		"MaxChange": (368, 2, (5, 0), (), "MaxChange", None),
		"MaxIterations": (369, 2, (3, 0), (), "MaxIterations", None),
		"MeasurementUnit": (2375, 2, (3, 0), (), "MeasurementUnit", None),
		"MemoryFree": (370, 2, (3, 0), (), "MemoryFree", None),
		"MemoryTotal": (371, 2, (3, 0), (), "MemoryTotal", None),
		"MemoryUsed": (372, 2, (3, 0), (), "MemoryUsed", None),
		# Method 'MenuBars' returns object of type 'MenuBars'
		"MenuBars": (589, 2, (9, 0), (), "MenuBars", '{00020863-0000-0000-C000-000000000046}'),
		# Method 'Modules' returns object of type 'Modules'
		"Modules": (582, 2, (9, 0), (), "Modules", '{000208AE-0000-0000-C000-000000000046}'),
		"MouseAvailable": (373, 2, (11, 0), (), "MouseAvailable", None),
		"MoveAfterReturn": (374, 2, (11, 0), (), "MoveAfterReturn", None),
		"MoveAfterReturnDirection": (1144, 2, (3, 0), (), "MoveAfterReturnDirection", None),
		# Method 'MultiThreadedCalculation' returns object of type 'MultiThreadedCalculation'
		"MultiThreadedCalculation": (2391, 2, (9, 0), (), "MultiThreadedCalculation", '{000244B1-0000-0000-C000-000000000046}'),
		"Name": (110, 2, (8, 0), (), "Name", None),
		# Method 'Names' returns object of type 'Names'
		"Names": (442, 2, (9, 0), (), "Names", '{000208B8-0000-0000-C000-000000000046}'),
		"NetworkTemplatesPath": (388, 2, (8, 0), (), "NetworkTemplatesPath", None),
		# Method 'NewWorkbook' returns object of type 'NewFile'
		"NewWorkbook": (1565, 2, (9, 0), (), "NewWorkbook", '{000C0936-0000-0000-C000-000000000046}'),
		# Method 'ODBCErrors' returns object of type 'ODBCErrors'
		"ODBCErrors": (1203, 2, (9, 0), (), "ODBCErrors", '{0002442D-0000-0000-C000-000000000046}'),
		"ODBCTimeout": (1204, 2, (3, 0), (), "ODBCTimeout", None),
		# Method 'OLEDBErrors' returns object of type 'OLEDBErrors'
		"OLEDBErrors": (1794, 2, (9, 0), (), "OLEDBErrors", '{00024446-0000-0000-C000-000000000046}'),
		"OnCalculate": (625, 2, (8, 0), (), "OnCalculate", None),
		"OnData": (629, 2, (8, 0), (), "OnData", None),
		"OnDoubleClick": (628, 2, (8, 0), (), "OnDoubleClick", None),
		"OnEntry": (627, 2, (8, 0), (), "OnEntry", None),
		"OnSheetActivate": (1031, 2, (8, 0), (), "OnSheetActivate", None),
		"OnSheetDeactivate": (1081, 2, (8, 0), (), "OnSheetDeactivate", None),
		"OnWindow": (623, 2, (8, 0), (), "OnWindow", None),
		"OperatingSystem": (375, 2, (8, 0), (), "OperatingSystem", None),
		"OrganizationName": (376, 2, (8, 0), (), "OrganizationName", None),
		# Method 'Parent' returns object of type 'Application'
		"Parent": (150, 2, (13, 0), (), "Parent", '{00024500-0000-0000-C000-000000000046}'),
		"Path": (291, 2, (8, 0), (), "Path", None),
		"PathSeparator": (377, 2, (8, 0), (), "PathSeparator", None),
		"PivotTableSelection": (1205, 2, (11, 0), (), "PivotTableSelection", None),
		"PreviousSelections": (378, 2, (12, 0), ((12, 17),), "PreviousSelections", None),
		"PrintCommunication": (2776, 2, (11, 0), (), "PrintCommunication", None),
		"ProductCode": (1798, 2, (8, 0), (), "ProductCode", None),
		"PromptForSummaryInfo": (1062, 2, (11, 0), (), "PromptForSummaryInfo", None),
		# Method 'ProtectedViewWindows' returns object of type 'ProtectedViewWindows'
		"ProtectedViewWindows": (2783, 2, (9, 0), (), "ProtectedViewWindows", '{000244CC-0000-0000-C000-000000000046}'),
		"Quitting": (2780, 2, (11, 0), (), "Quitting", None),
		# Method 'RTD' returns object of type 'RTD'
		"RTD": (1963, 2, (9, 0), (), "RTD", '{0002446E-0000-0000-C000-000000000046}'),
		"Ready": (1932, 2, (11, 0), (), "Ready", None),
		# Method 'RecentFiles' returns object of type 'RecentFiles'
		"RecentFiles": (1202, 2, (9, 0), (), "RecentFiles", '{00024406-0000-0000-C000-000000000046}'),
		"RecordRelative": (379, 2, (11, 0), (), "RecordRelative", None),
		"ReferenceStyle": (380, 2, (3, 0), (), "ReferenceStyle", None),
		"RegisteredFunctions": (775, 2, (12, 0), ((12, 17), (12, 17)), "RegisteredFunctions", None),
		# Method 'ReplaceFormat' returns object of type 'CellFormat'
		"ReplaceFormat": (1935, 2, (9, 0), (), "ReplaceFormat", '{00024450-0000-0000-C000-000000000046}'),
		"RollZoom": (1206, 2, (11, 0), (), "RollZoom", None),
		# Method 'Rows' returns object of type 'Range'
		"Rows": (258, 2, (9, 0), (), "Rows", '{00020846-0000-0000-C000-000000000046}'),
		"SaveISO8601Dates": (2786, 2, (11, 0), (), "SaveISO8601Dates", None),
		"ScreenUpdating": (382, 2, (11, 0), (), "ScreenUpdating", None),
		"Selection": (147, 2, (9, 0), (), "Selection", None),
		# Method 'Sheets' returns object of type 'Sheets'
		"Sheets": (485, 2, (9, 0), (), "Sheets", '{000208D7-0000-0000-C000-000000000046}'),
		"SheetsInNewWorkbook": (993, 2, (3, 0), (), "SheetsInNewWorkbook", None),
		"ShowChartTipNames": (1207, 2, (11, 0), (), "ShowChartTipNames", None),
		"ShowChartTipValues": (1208, 2, (11, 0), (), "ShowChartTipValues", None),
		"ShowDevTools": (2378, 2, (11, 0), (), "ShowDevTools", None),
		"ShowMenuFloaties": (2377, 2, (11, 0), (), "ShowMenuFloaties", None),
		"ShowSelectionFloaties": (2376, 2, (11, 0), (), "ShowSelectionFloaties", None),
		"ShowStartupDialog": (1960, 2, (11, 0), (), "ShowStartupDialog", None),
		"ShowToolTips": (387, 2, (11, 0), (), "ShowToolTips", None),
		"ShowWindowsInTaskbar": (1807, 2, (11, 0), (), "ShowWindowsInTaskbar", None),
		# Method 'SmartArtColors' returns object of type 'SmartArtColors'
		"SmartArtColors": (2774, 2, (9, 0), (), "SmartArtColors", '{000C03CD-0000-0000-C000-000000000046}'),
		# Method 'SmartArtLayouts' returns object of type 'SmartArtLayouts'
		"SmartArtLayouts": (2772, 2, (9, 0), (), "SmartArtLayouts", '{000C03C9-0000-0000-C000-000000000046}'),
		# Method 'SmartArtQuickStyles' returns object of type 'SmartArtQuickStyles'
		"SmartArtQuickStyles": (2773, 2, (9, 0), (), "SmartArtQuickStyles", '{000C03CB-0000-0000-C000-000000000046}'),
		# Method 'SmartTagRecognizers' returns object of type 'SmartTagRecognizers'
		"SmartTagRecognizers": (1956, 2, (9, 0), (), "SmartTagRecognizers", '{00024463-0000-0000-C000-000000000046}'),
		# Method 'Speech' returns object of type 'Speech'
		"Speech": (1958, 2, (9, 0), (), "Speech", '{00024466-0000-0000-C000-000000000046}'),
		# Method 'SpellingOptions' returns object of type 'SpellingOptions'
		"SpellingOptions": (1957, 2, (9, 0), (), "SpellingOptions", '{00024465-0000-0000-C000-000000000046}'),
		"StandardFont": (924, 2, (8, 0), (), "StandardFont", None),
		"StandardFontSize": (925, 2, (5, 0), (), "StandardFontSize", None),
		"StartupPath": (385, 2, (8, 0), (), "StartupPath", None),
		"StatusBar": (386, 2, (12, 0), (), "StatusBar", None),
		"TemplatesPath": (381, 2, (8, 0), (), "TemplatesPath", None),
		# Method 'ThisCell' returns object of type 'Range'
		"ThisCell": (1962, 2, (9, 0), (), "ThisCell", '{00020846-0000-0000-C000-000000000046}'),
		# Method 'ThisWorkbook' returns object of type 'Workbook'
		"ThisWorkbook": (778, 2, (13, 0), (), "ThisWorkbook", '{00020819-0000-0000-C000-000000000046}'),
		"ThousandsSeparator": (1810, 2, (8, 0), (), "ThousandsSeparator", None),
		# Method 'Toolbars' returns object of type 'Toolbars'
		"Toolbars": (552, 2, (9, 0), (), "Toolbars", '{0002085D-0000-0000-C000-000000000046}'),
		"Top": (126, 2, (5, 0), (), "Top", None),
		"TransitionMenuKey": (310, 2, (8, 0), (), "TransitionMenuKey", None),
		"TransitionMenuKeyAction": (311, 2, (3, 0), (), "TransitionMenuKeyAction", None),
		"TransitionNavigKeys": (312, 2, (11, 0), (), "TransitionNavigKeys", None),
		"UILanguage": (2, 2, (3, 0), (), "UILanguage", None),
		"UsableHeight": (389, 2, (5, 0), (), "UsableHeight", None),
		"UsableWidth": (390, 2, (5, 0), (), "UsableWidth", None),
		"UseClusterConnector": (2778, 2, (11, 0), (), "UseClusterConnector", None),
		"UseSystemSeparators": (1961, 2, (11, 0), (), "UseSystemSeparators", None),
		# Method 'UsedObjects' returns object of type 'UsedObjects'
		"UsedObjects": (1936, 2, (9, 0), (), "UsedObjects", '{00024451-0000-0000-C000-000000000046}'),
		"UserControl": (1210, 2, (11, 0), (), "UserControl", None),
		"UserLibraryPath": (1799, 2, (8, 0), (), "UserLibraryPath", None),
		"UserName": (391, 2, (8, 0), (), "UserName", None),
		# Method 'VBE' returns object of type 'VBE'
		"VBE": (1211, 2, (9, 0), (), "VBE", '{0002E166-0000-0000-C000-000000000046}'),
		"Value": (6, 2, (8, 0), (), "Value", None),
		"Version": (392, 2, (8, 0), (), "Version", None),
		"Visible": (558, 2, (11, 0), (), "Visible", None),
		"WarnOnFunctionNameConflict": (2382, 2, (11, 0), (), "WarnOnFunctionNameConflict", None),
		# Method 'Watches' returns object of type 'Watches'
		"Watches": (1939, 2, (9, 0), (), "Watches", '{00024456-0000-0000-C000-000000000046}'),
		"Width": (122, 2, (5, 0), (), "Width", None),
		"WindowState": (396, 2, (3, 0), (), "WindowState", None),
		# Method 'Windows' returns object of type 'Windows'
		"Windows": (430, 2, (9, 0), (), "Windows", '{00020892-0000-0000-C000-000000000046}'),
		"WindowsForPens": (395, 2, (11, 0), (), "WindowsForPens", None),
		# Method 'Workbooks' returns object of type 'Workbooks'
		"Workbooks": (572, 2, (9, 0), (), "Workbooks", '{000208DB-0000-0000-C000-000000000046}'),
		# Method 'WorksheetFunction' returns object of type 'WorksheetFunction'
		"WorksheetFunction": (1440, 2, (9, 0), (), "WorksheetFunction", '{00020845-0000-0000-C000-000000000046}'),
		# Method 'Worksheets' returns object of type 'Sheets'
		"Worksheets": (494, 2, (9, 0), (), "Worksheets", '{000208D7-0000-0000-C000-000000000046}'),
		"_Default": (0, 2, (8, 0), (), "_Default", None),
	}
	_prop_map_put_ = {
		"ActivePrinter": ((306, LCID, 4, 0),()),
		"AlertBeforeOverwriting": ((930, LCID, 4, 0),()),
		"AltStartupPath": ((313, LCID, 4, 0),()),
		"AlwaysUseClearType": ((2381, LCID, 4, 0),()),
		"AskToUpdateLinks": ((992, LCID, 4, 0),()),
		"AutoFormatAsYouTypeReplaceHyperlinks": ((1955, LCID, 4, 0),()),
		"AutoPercentEntry": ((1800, LCID, 4, 0),()),
		"AutomationSecurity": ((1941, LCID, 4, 0),()),
		"CalculateBeforeSave": ((315, LCID, 4, 0),()),
		"Calculation": ((316, LCID, 4, 0),()),
		"CalculationInterruptKey": ((1938, LCID, 4, 0),()),
		"Caption": ((139, LCID, 4, 0),()),
		"CellDragAndDrop": ((320, LCID, 4, 0),()),
		"ClusterConnector": ((2779, LCID, 4, 0),()),
		"ColorButtons": ((365, LCID, 4, 0),()),
		"CommandUnderlines": ((323, LCID, 4, 0),()),
		"ConstrainNumeric": ((324, LCID, 4, 0),()),
		"ControlCharacters": ((233, LCID, 4, 0),()),
		"CopyObjectsWithCells": ((991, LCID, 4, 0),()),
		"Cursor": ((1161, LCID, 4, 0),()),
		"CursorMovement": ((232, LCID, 4, 0),()),
		"CutCopyMode": ((330, LCID, 4, 0),()),
		"DataEntryMode": ((331, LCID, 4, 0),()),
		"DecimalSeparator": ((1809, LCID, 4, 0),()),
		"DefaultFilePath": ((1038, LCID, 4, 0),()),
		"DefaultSaveFormat": ((1209, LCID, 4, 0),()),
		"DefaultSheetDirection": ((229, LCID, 4, 0),()),
		"DeferAsyncQueries": ((2390, LCID, 4, 0),()),
		"DisplayAlerts": ((343, LCID, 4, 0),()),
		"DisplayClipboardWindow": ((322, LCID, 4, 0),()),
		"DisplayCommentIndicator": ((1196, LCID, 4, 0),()),
		"DisplayDocumentActionTaskPane": ((2251, LCID, 4, 0),()),
		"DisplayDocumentInformationPanel": ((2380, LCID, 4, 0),()),
		"DisplayExcel4Menus": ((927, LCID, 4, 0),()),
		"DisplayFormulaAutoComplete": ((2384, LCID, 4, 0),()),
		"DisplayFormulaBar": ((344, LCID, 4, 0),()),
		"DisplayFullScreen": ((1061, LCID, 4, 0),()),
		"DisplayFunctionToolTips": ((1940, LCID, 4, 0),()),
		"DisplayInfoWindow": ((1213, LCID, 4, 0),()),
		"DisplayInsertOptions": ((1947, LCID, 4, 0),()),
		"DisplayNoteIndicator": ((345, LCID, 4, 0),()),
		"DisplayPasteOptions": ((1946, LCID, 4, 0),()),
		"DisplayRecentFiles": ((926, LCID, 4, 0),()),
		"DisplayScrollBars": ((346, LCID, 4, 0),()),
		"DisplayStatusBar": ((347, LCID, 4, 0),()),
		"Dummy22": ((2781, LCID, 4, 0),()),
		"Dummy23": ((2782, LCID, 4, 0),()),
		"EditDirectlyInCell": ((929, LCID, 4, 0),()),
		"EnableAnimations": ((1180, LCID, 4, 0),()),
		"EnableAutoComplete": ((1179, LCID, 4, 0),()),
		"EnableCancelKey": ((1096, LCID, 4, 0),()),
		"EnableEvents": ((1212, LCID, 4, 0),()),
		"EnableLargeOperationAlert": ((2388, LCID, 4, 0),()),
		"EnableLivePreview": ((2379, LCID, 4, 0),()),
		"EnableSound": ((1197, LCID, 4, 0),()),
		"EnableTipWizard": ((1064, LCID, 4, 0),()),
		"ExtendList": ((1793, LCID, 4, 0),()),
		"FeatureInstall": ((1808, LCID, 4, 0),()),
		"FileValidation": ((2788, LCID, 4, 0),()),
		"FileValidationPivot": ((2789, LCID, 4, 0),()),
		"FindFormat": ((1934, LCID, 8, 0),()),
		"FixedDecimal": ((351, LCID, 4, 0),()),
		"FixedDecimalPlaces": ((352, LCID, 4, 0),()),
		"FormulaBarHeight": ((2383, LCID, 4, 0),()),
		"GenerateGetPivotData": ((1948, LCID, 4, 0),()),
		"GenerateTableRefs": ((2385, LCID, 4, 0),()),
		"Height": ((123, LCID, 4, 0),()),
		"HighQualityModeForGraphics": ((2395, LCID, 4, 0),()),
		"IgnoreRemoteRequests": ((356, LCID, 4, 0),()),
		"Interactive": ((361, LCID, 4, 0),()),
		"Iteration": ((363, LCID, 4, 0),()),
		"LargeButtons": ((364, LCID, 4, 0),()),
		"LargeOperationCellThousandCount": ((2389, LCID, 4, 0),()),
		"Left": ((127, LCID, 4, 0),()),
		"MapPaperSize": ((1959, LCID, 4, 0),()),
		"MaxChange": ((368, LCID, 4, 0),()),
		"MaxIterations": ((369, LCID, 4, 0),()),
		"MeasurementUnit": ((2375, LCID, 4, 0),()),
		"MoveAfterReturn": ((374, LCID, 4, 0),()),
		"MoveAfterReturnDirection": ((1144, LCID, 4, 0),()),
		"ODBCTimeout": ((1204, LCID, 4, 0),()),
		"OnCalculate": ((625, LCID, 4, 0),()),
		"OnData": ((629, LCID, 4, 0),()),
		"OnDoubleClick": ((628, LCID, 4, 0),()),
		"OnEntry": ((627, LCID, 4, 0),()),
		"OnSheetActivate": ((1031, LCID, 4, 0),()),
		"OnSheetDeactivate": ((1081, LCID, 4, 0),()),
		"OnWindow": ((623, LCID, 4, 0),()),
		"PivotTableSelection": ((1205, LCID, 4, 0),()),
		"PrintCommunication": ((2776, LCID, 4, 0),()),
		"PromptForSummaryInfo": ((1062, LCID, 4, 0),()),
		"ReferenceStyle": ((380, LCID, 4, 0),()),
		"ReplaceFormat": ((1935, LCID, 8, 0),()),
		"RollZoom": ((1206, LCID, 4, 0),()),
		"SaveISO8601Dates": ((2786, LCID, 4, 0),()),
		"ScreenUpdating": ((382, LCID, 4, 0),()),
		"SheetsInNewWorkbook": ((993, LCID, 4, 0),()),
		"ShowChartTipNames": ((1207, LCID, 4, 0),()),
		"ShowChartTipValues": ((1208, LCID, 4, 0),()),
		"ShowDevTools": ((2378, LCID, 4, 0),()),
		"ShowMenuFloaties": ((2377, LCID, 4, 0),()),
		"ShowSelectionFloaties": ((2376, LCID, 4, 0),()),
		"ShowStartupDialog": ((1960, LCID, 4, 0),()),
		"ShowToolTips": ((387, LCID, 4, 0),()),
		"ShowWindowsInTaskbar": ((1807, LCID, 4, 0),()),
		"StandardFont": ((924, LCID, 4, 0),()),
		"StandardFontSize": ((925, LCID, 4, 0),()),
		"StatusBar": ((386, LCID, 4, 0),()),
		"ThousandsSeparator": ((1810, LCID, 4, 0),()),
		"Top": ((126, LCID, 4, 0),()),
		"TransitionMenuKey": ((310, LCID, 4, 0),()),
		"TransitionMenuKeyAction": ((311, LCID, 4, 0),()),
		"TransitionNavigKeys": ((312, LCID, 4, 0),()),
		"UILanguage": ((2, LCID, 4, 0),()),
		"UseClusterConnector": ((2778, LCID, 4, 0),()),
		"UseSystemSeparators": ((1961, LCID, 4, 0),()),
		"UserControl": ((1210, LCID, 4, 0),()),
		"UserName": ((391, LCID, 4, 0),()),
		"Visible": ((558, LCID, 4, 0),()),
		"WarnOnFunctionNameConflict": ((2382, LCID, 4, 0),()),
		"Width": ((122, LCID, 4, 0),()),
		"WindowState": ((396, LCID, 4, 0),()),
	}
	# Default property for this class is 'Value'
	def __call__(self):
		return self._ApplyTypes_(*(6, 2, (8, 0), (), "Value", None))
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

win32com.client.CLSIDToClass.RegisterCLSID( "{000208D5-0000-0000-C000-000000000046}", _Application )
# -*- coding: mbcs -*-
# Created by makepy.py version 0.5.01
# By python version 3.3.1 (v3.3.1:d9893d13c628, Apr  6 2013, 20:25:12) [MSC v.1600 32 bit (Intel)]
# From type library '{00020813-0000-0000-C000-000000000046}'
# On Sat Apr  5 18:14:59 2014
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

_Application_vtables_dispatch_ = 1
_Application_vtables_ = [
	(( 'Application' , 'RHS' , ), 148, (148, (), [ (16397, 10, None, "IID('{00024500-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 28 , (3, 0, None, None) , 0 , )),
	(( 'Creator' , 'RHS' , ), 149, (149, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 32 , (3, 0, None, None) , 0 , )),
	(( 'Parent' , 'RHS' , ), 150, (150, (), [ (16397, 10, None, "IID('{00024500-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 36 , (3, 0, None, None) , 0 , )),
	(( 'ActiveCell' , 'RHS' , ), 305, (305, (), [ (16393, 10, None, "IID('{00020846-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 40 , (3, 0, None, None) , 0 , )),
	(( 'ActiveChart' , 'RHS' , ), 183, (183, (), [ (16397, 10, None, "IID('{00020821-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 44 , (3, 0, None, None) , 0 , )),
	(( 'ActiveDialog' , 'RHS' , ), 815, (815, (), [ (16393, 10, None, "IID('{000208AF-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 48 , (3, 0, None, None) , 64 , )),
	(( 'ActiveMenuBar' , 'RHS' , ), 758, (758, (), [ (16393, 10, None, "IID('{00020864-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 52 , (3, 0, None, None) , 64 , )),
	(( 'ActivePrinter' , 'lcid' , 'RHS' , ), 306, (306, (), [ (3, 5, None, None) , 
			 (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 56 , (3, 0, None, None) , 0 , )),
	(( 'ActivePrinter' , 'lcid' , 'RHS' , ), 306, (306, (), [ (3, 5, None, None) , 
			 (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 60 , (3, 0, None, None) , 0 , )),
	(( 'ActiveSheet' , 'RHS' , ), 307, (307, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 64 , (3, 0, None, None) , 0 , )),
	(( 'ActiveWindow' , 'RHS' , ), 759, (759, (), [ (16393, 10, None, "IID('{00020893-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 68 , (3, 0, None, None) , 0 , )),
	(( 'ActiveWorkbook' , 'RHS' , ), 308, (308, (), [ (16397, 10, None, "IID('{00020819-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 72 , (3, 0, None, None) , 0 , )),
	(( 'AddIns' , 'RHS' , ), 549, (549, (), [ (16393, 10, None, "IID('{00020858-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 76 , (3, 0, None, None) , 1024 , )),
	(( 'Assistant' , 'RHS' , ), 1438, (1438, (), [ (16393, 10, None, "IID('{000C0322-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 80 , (3, 0, None, None) , 64 , )),
	(( 'Calculate' , 'lcid' , ), 279, (279, (), [ (3, 5, None, None) , ], 1 , 1 , 4 , 0 , 84 , (3, 0, None, None) , 0 , )),
	(( 'Cells' , 'RHS' , ), 238, (238, (), [ (16393, 10, None, "IID('{00020846-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 88 , (3, 0, None, None) , 0 , )),
	(( 'Charts' , 'RHS' , ), 121, (121, (), [ (16393, 10, None, "IID('{000208D7-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 92 , (3, 0, None, None) , 0 , )),
	(( 'Columns' , 'RHS' , ), 241, (241, (), [ (16393, 10, None, "IID('{00020846-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 96 , (3, 0, None, None) , 1024 , )),
	(( 'CommandBars' , 'RHS' , ), 1439, (1439, (), [ (16397, 10, None, "IID('{55F88893-7708-11D1-ACEB-006008961DA5}')") , ], 1 , 2 , 4 , 0 , 100 , (3, 0, None, None) , 0 , )),
	(( 'DDEAppReturnCode' , 'lcid' , 'RHS' , ), 332, (332, (), [ (3, 5, None, None) , 
			 (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 104 , (3, 0, None, None) , 0 , )),
	(( 'DDEExecute' , 'Channel' , 'String' , 'lcid' , ), 333, (333, (), [ 
			 (3, 1, None, None) , (8, 1, None, None) , (3, 5, None, None) , ], 1 , 1 , 4 , 0 , 108 , (3, 0, None, None) , 0 , )),
	(( 'DDEInitiate' , 'App' , 'Topic' , 'lcid' , 'RHS' , 
			 ), 334, (334, (), [ (8, 1, None, None) , (8, 1, None, None) , (3, 5, None, None) , (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 112 , (3, 0, None, None) , 0 , )),
	(( 'DDEPoke' , 'Channel' , 'Item' , 'Data' , 'lcid' , 
			 ), 335, (335, (), [ (3, 1, None, None) , (12, 1, None, None) , (12, 1, None, None) , (3, 5, None, None) , ], 1 , 1 , 4 , 0 , 116 , (3, 0, None, None) , 0 , )),
	(( 'DDERequest' , 'Channel' , 'Item' , 'lcid' , 'RHS' , 
			 ), 336, (336, (), [ (3, 1, None, None) , (8, 1, None, None) , (3, 5, None, None) , (16396, 10, None, None) , ], 1 , 1 , 4 , 0 , 120 , (3, 0, None, None) , 0 , )),
	(( 'DDETerminate' , 'Channel' , 'lcid' , ), 337, (337, (), [ (3, 1, None, None) , 
			 (3, 5, None, None) , ], 1 , 1 , 4 , 0 , 124 , (3, 0, None, None) , 0 , )),
	(( 'DialogSheets' , 'RHS' , ), 764, (764, (), [ (16393, 10, None, "IID('{000208D7-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 128 , (3, 0, None, None) , 64 , )),
	(( 'Evaluate' , 'Name' , 'lcid' , 'RHS' , ), 1, (1, (), [ 
			 (12, 1, None, None) , (3, 5, None, None) , (16396, 10, None, None) , ], 1 , 1 , 4 , 0 , 132 , (3, 0, None, None) , 0 , )),
	(( '_Evaluate' , 'Name' , 'lcid' , 'RHS' , ), -5, (-5, (), [ 
			 (12, 1, None, None) , (3, 5, None, None) , (16396, 10, None, None) , ], 1 , 1 , 4 , 0 , 136 , (3, 0, None, None) , 1024 , )),
	(( 'ExecuteExcel4Macro' , 'String' , 'lcid' , 'RHS' , ), 350, (350, (), [ 
			 (8, 1, None, None) , (3, 5, None, None) , (16396, 10, None, None) , ], 1 , 1 , 4 , 0 , 140 , (3, 0, None, None) , 0 , )),
	(( 'Intersect' , 'Arg1' , 'Arg2' , 'Arg3' , 'Arg4' , 
			 'Arg5' , 'Arg6' , 'Arg7' , 'Arg8' , 'Arg9' , 
			 'Arg10' , 'Arg11' , 'Arg12' , 'Arg13' , 'Arg14' , 
			 'Arg15' , 'Arg16' , 'Arg17' , 'Arg18' , 'Arg19' , 
			 'Arg20' , 'Arg21' , 'Arg22' , 'Arg23' , 'Arg24' , 
			 'Arg25' , 'Arg26' , 'Arg27' , 'Arg28' , 'Arg29' , 
			 'Arg30' , 'lcid' , 'RHS' , ), 766, (766, (), [ (9, 1, None, "IID('{00020846-0000-0000-C000-000000000046}')") , 
			 (9, 1, None, "IID('{00020846-0000-0000-C000-000000000046}')") , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (3, 5, None, None) , 
			 (16393, 10, None, "IID('{00020846-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 28 , 144 , (3, 0, None, None) , 0 , )),
	(( 'MenuBars' , 'RHS' , ), 589, (589, (), [ (16393, 10, None, "IID('{00020863-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 148 , (3, 0, None, None) , 64 , )),
	(( 'Modules' , 'RHS' , ), 582, (582, (), [ (16393, 10, None, "IID('{000208AE-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 152 , (3, 0, None, None) , 64 , )),
	(( 'Names' , 'RHS' , ), 442, (442, (), [ (16393, 10, None, "IID('{000208B8-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 156 , (3, 0, None, None) , 0 , )),
	(( 'Range' , 'Cell1' , 'Cell2' , 'RHS' , ), 197, (197, (), [ 
			 (12, 1, None, None) , (12, 17, None, None) , (16393, 10, None, "IID('{00020846-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 1 , 160 , (3, 0, None, None) , 0 , )),
	(( 'Rows' , 'RHS' , ), 258, (258, (), [ (16393, 10, None, "IID('{00020846-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 164 , (3, 0, None, None) , 1024 , )),
	(( 'Run' , 'Macro' , 'Arg1' , 'Arg2' , 'Arg3' , 
			 'Arg4' , 'Arg5' , 'Arg6' , 'Arg7' , 'Arg8' , 
			 'Arg9' , 'Arg10' , 'Arg11' , 'Arg12' , 'Arg13' , 
			 'Arg14' , 'Arg15' , 'Arg16' , 'Arg17' , 'Arg18' , 
			 'Arg19' , 'Arg20' , 'Arg21' , 'Arg22' , 'Arg23' , 
			 'Arg24' , 'Arg25' , 'Arg26' , 'Arg27' , 'Arg28' , 
			 'Arg29' , 'Arg30' , 'RHS' , ), 259, (259, (), [ (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (16396, 10, None, None) , ], 1 , 1 , 4 , 31 , 168 , (3, 0, None, None) , 0 , )),
	(( '_Run2' , 'Macro' , 'Arg1' , 'Arg2' , 'Arg3' , 
			 'Arg4' , 'Arg5' , 'Arg6' , 'Arg7' , 'Arg8' , 
			 'Arg9' , 'Arg10' , 'Arg11' , 'Arg12' , 'Arg13' , 
			 'Arg14' , 'Arg15' , 'Arg16' , 'Arg17' , 'Arg18' , 
			 'Arg19' , 'Arg20' , 'Arg21' , 'Arg22' , 'Arg23' , 
			 'Arg24' , 'Arg25' , 'Arg26' , 'Arg27' , 'Arg28' , 
			 'Arg29' , 'Arg30' , 'lcid' , 'RHS' , ), 806, (806, (), [ 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (3, 5, None, None) , (16396, 10, None, None) , ], 1 , 1 , 4 , 31 , 172 , (3, 0, None, None) , 1024 , )),
	(( 'Selection' , 'lcid' , 'RHS' , ), 147, (147, (), [ (3, 5, None, None) , 
			 (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 176 , (3, 0, None, None) , 0 , )),
	(( 'SendKeys' , 'Keys' , 'Wait' , 'lcid' , ), 383, (383, (), [ 
			 (12, 1, None, None) , (12, 17, None, None) , (3, 5, None, None) , ], 1 , 1 , 4 , 1 , 180 , (3, 0, None, None) , 0 , )),
	(( 'Sheets' , 'RHS' , ), 485, (485, (), [ (16393, 10, None, "IID('{000208D7-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 184 , (3, 0, None, None) , 0 , )),
	(( 'ShortcutMenus' , 'Index' , 'RHS' , ), 776, (776, (), [ (3, 1, None, None) , 
			 (16393, 10, None, "IID('{00020866-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 188 , (3, 0, None, None) , 64 , )),
	(( 'ThisWorkbook' , 'lcid' , 'RHS' , ), 778, (778, (), [ (3, 5, None, None) , 
			 (16397, 10, None, "IID('{00020819-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 192 , (3, 0, None, None) , 0 , )),
	(( 'Toolbars' , 'RHS' , ), 552, (552, (), [ (16393, 10, None, "IID('{0002085D-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 196 , (3, 0, None, None) , 64 , )),
	(( 'Union' , 'Arg1' , 'Arg2' , 'Arg3' , 'Arg4' , 
			 'Arg5' , 'Arg6' , 'Arg7' , 'Arg8' , 'Arg9' , 
			 'Arg10' , 'Arg11' , 'Arg12' , 'Arg13' , 'Arg14' , 
			 'Arg15' , 'Arg16' , 'Arg17' , 'Arg18' , 'Arg19' , 
			 'Arg20' , 'Arg21' , 'Arg22' , 'Arg23' , 'Arg24' , 
			 'Arg25' , 'Arg26' , 'Arg27' , 'Arg28' , 'Arg29' , 
			 'Arg30' , 'lcid' , 'RHS' , ), 779, (779, (), [ (9, 1, None, "IID('{00020846-0000-0000-C000-000000000046}')") , 
			 (9, 1, None, "IID('{00020846-0000-0000-C000-000000000046}')") , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (3, 5, None, None) , 
			 (16393, 10, None, "IID('{00020846-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 28 , 200 , (3, 0, None, None) , 0 , )),
	(( 'Windows' , 'RHS' , ), 430, (430, (), [ (16393, 10, None, "IID('{00020892-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 204 , (3, 0, None, None) , 0 , )),
	(( 'Workbooks' , 'RHS' , ), 572, (572, (), [ (16393, 10, None, "IID('{000208DB-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 208 , (3, 0, None, None) , 0 , )),
	(( 'WorksheetFunction' , 'RHS' , ), 1440, (1440, (), [ (16393, 10, None, "IID('{00020845-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 212 , (3, 0, None, None) , 0 , )),
	(( 'Worksheets' , 'RHS' , ), 494, (494, (), [ (16393, 10, None, "IID('{000208D7-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 216 , (3, 0, None, None) , 0 , )),
	(( 'Excel4IntlMacroSheets' , 'RHS' , ), 581, (581, (), [ (16393, 10, None, "IID('{000208D7-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 220 , (3, 0, None, None) , 0 , )),
	(( 'Excel4MacroSheets' , 'RHS' , ), 579, (579, (), [ (16393, 10, None, "IID('{000208D7-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 224 , (3, 0, None, None) , 0 , )),
	(( 'ActivateMicrosoftApp' , 'Index' , 'lcid' , ), 1095, (1095, (), [ (3, 1, None, None) , 
			 (3, 5, None, None) , ], 1 , 1 , 4 , 0 , 228 , (3, 0, None, None) , 0 , )),
	(( 'AddChartAutoFormat' , 'Chart' , 'Name' , 'Description' , 'lcid' , 
			 ), 216, (216, (), [ (12, 1, None, None) , (8, 1, None, None) , (12, 17, None, None) , (3, 5, None, None) , ], 1 , 1 , 4 , 1 , 232 , (3, 0, None, None) , 64 , )),
	(( 'AddCustomList' , 'ListArray' , 'ByRow' , 'lcid' , ), 780, (780, (), [ 
			 (12, 1, None, None) , (12, 17, None, None) , (3, 5, None, None) , ], 1 , 1 , 4 , 1 , 236 , (3, 0, None, None) , 0 , )),
	(( 'AlertBeforeOverwriting' , 'lcid' , 'RHS' , ), 930, (930, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 240 , (3, 0, None, None) , 0 , )),
	(( 'AlertBeforeOverwriting' , 'lcid' , 'RHS' , ), 930, (930, (), [ (3, 5, None, None) , 
			 (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 244 , (3, 0, None, None) , 0 , )),
	(( 'AltStartupPath' , 'lcid' , 'RHS' , ), 313, (313, (), [ (3, 5, None, None) , 
			 (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 248 , (3, 0, None, None) , 0 , )),
	(( 'AltStartupPath' , 'lcid' , 'RHS' , ), 313, (313, (), [ (3, 5, None, None) , 
			 (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 252 , (3, 0, None, None) , 0 , )),
	(( 'AskToUpdateLinks' , 'lcid' , 'RHS' , ), 992, (992, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 256 , (3, 0, None, None) , 0 , )),
	(( 'AskToUpdateLinks' , 'lcid' , 'RHS' , ), 992, (992, (), [ (3, 5, None, None) , 
			 (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 260 , (3, 0, None, None) , 0 , )),
	(( 'EnableAnimations' , 'lcid' , 'RHS' , ), 1180, (1180, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 264 , (3, 0, None, None) , 0 , )),
	(( 'EnableAnimations' , 'lcid' , 'RHS' , ), 1180, (1180, (), [ (3, 5, None, None) , 
			 (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 268 , (3, 0, None, None) , 0 , )),
	(( 'AutoCorrect' , 'RHS' , ), 1145, (1145, (), [ (16393, 10, None, "IID('{000208D4-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 272 , (3, 0, None, None) , 0 , )),
	(( 'Build' , 'lcid' , 'RHS' , ), 314, (314, (), [ (3, 5, None, None) , 
			 (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 276 , (3, 0, None, None) , 0 , )),
	(( 'CalculateBeforeSave' , 'lcid' , 'RHS' , ), 315, (315, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 280 , (3, 0, None, None) , 0 , )),
	(( 'CalculateBeforeSave' , 'lcid' , 'RHS' , ), 315, (315, (), [ (3, 5, None, None) , 
			 (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 284 , (3, 0, None, None) , 0 , )),
	(( 'Calculation' , 'lcid' , 'RHS' , ), 316, (316, (), [ (3, 5, None, None) , 
			 (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 288 , (3, 0, None, None) , 0 , )),
	(( 'Calculation' , 'lcid' , 'RHS' , ), 316, (316, (), [ (3, 5, None, None) , 
			 (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 292 , (3, 0, None, None) , 0 , )),
	(( 'Caller' , 'Index' , 'lcid' , 'RHS' , ), 317, (317, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16396, 10, None, None) , ], 1 , 2 , 4 , 1 , 296 , (3, 0, None, None) , 0 , )),
	(( 'CanPlaySounds' , 'lcid' , 'RHS' , ), 318, (318, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 300 , (3, 0, None, None) , 0 , )),
	(( 'CanRecordSounds' , 'lcid' , 'RHS' , ), 319, (319, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 304 , (3, 0, None, None) , 0 , )),
	(( 'Caption' , 'RHS' , ), 139, (139, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 308 , (3, 0, None, None) , 0 , )),
	(( 'Caption' , 'RHS' , ), 139, (139, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 312 , (3, 0, None, None) , 0 , )),
	(( 'CellDragAndDrop' , 'lcid' , 'RHS' , ), 320, (320, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 316 , (3, 0, None, None) , 0 , )),
	(( 'CellDragAndDrop' , 'lcid' , 'RHS' , ), 320, (320, (), [ (3, 5, None, None) , 
			 (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 320 , (3, 0, None, None) , 0 , )),
	(( 'CentimetersToPoints' , 'Centimeters' , 'lcid' , 'RHS' , ), 1086, (1086, (), [ 
			 (5, 1, None, None) , (3, 5, None, None) , (16389, 10, None, None) , ], 1 , 1 , 4 , 0 , 324 , (3, 0, None, None) , 0 , )),
	(( 'CheckSpelling' , 'Word' , 'CustomDictionary' , 'IgnoreUppercase' , 'lcid' , 
			 'RHS' , ), 505, (505, (), [ (8, 1, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (3, 5, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 2 , 328 , (3, 0, None, None) , 0 , )),
	(( 'ClipboardFormats' , 'Index' , 'lcid' , 'RHS' , ), 321, (321, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16396, 10, None, None) , ], 1 , 2 , 4 , 1 , 332 , (3, 0, None, None) , 0 , )),
	(( 'DisplayClipboardWindow' , 'lcid' , 'RHS' , ), 322, (322, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 336 , (3, 0, None, None) , 0 , )),
	(( 'DisplayClipboardWindow' , 'lcid' , 'RHS' , ), 322, (322, (), [ (3, 5, None, None) , 
			 (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 340 , (3, 0, None, None) , 0 , )),
	(( 'ColorButtons' , 'RHS' , ), 365, (365, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 344 , (3, 0, None, None) , 64 , )),
	(( 'ColorButtons' , 'RHS' , ), 365, (365, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 348 , (3, 0, None, None) , 64 , )),
	(( 'CommandUnderlines' , 'lcid' , 'RHS' , ), 323, (323, (), [ (3, 5, None, None) , 
			 (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 352 , (3, 0, None, None) , 0 , )),
	(( 'CommandUnderlines' , 'lcid' , 'RHS' , ), 323, (323, (), [ (3, 5, None, None) , 
			 (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 356 , (3, 0, None, None) , 0 , )),
	(( 'ConstrainNumeric' , 'lcid' , 'RHS' , ), 324, (324, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 360 , (3, 0, None, None) , 0 , )),
	(( 'ConstrainNumeric' , 'lcid' , 'RHS' , ), 324, (324, (), [ (3, 5, None, None) , 
			 (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 364 , (3, 0, None, None) , 0 , )),
	(( 'ConvertFormula' , 'Formula' , 'FromReferenceStyle' , 'ToReferenceStyle' , 'ToAbsolute' , 
			 'RelativeTo' , 'lcid' , 'RHS' , ), 325, (325, (), [ (12, 1, None, None) , 
			 (3, 1, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (3, 5, None, None) , 
			 (16396, 10, None, None) , ], 1 , 1 , 4 , 3 , 368 , (3, 0, None, None) , 0 , )),
	(( 'CopyObjectsWithCells' , 'lcid' , 'RHS' , ), 991, (991, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 372 , (3, 0, None, None) , 0 , )),
	(( 'CopyObjectsWithCells' , 'lcid' , 'RHS' , ), 991, (991, (), [ (3, 5, None, None) , 
			 (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 376 , (3, 0, None, None) , 0 , )),
	(( 'Cursor' , 'lcid' , 'RHS' , ), 1161, (1161, (), [ (3, 5, None, None) , 
			 (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 380 , (3, 0, None, None) , 0 , )),
	(( 'Cursor' , 'lcid' , 'RHS' , ), 1161, (1161, (), [ (3, 5, None, None) , 
			 (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 384 , (3, 0, None, None) , 0 , )),
	(( 'CustomListCount' , 'lcid' , 'RHS' , ), 787, (787, (), [ (3, 5, None, None) , 
			 (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 388 , (3, 0, None, None) , 0 , )),
	(( 'CutCopyMode' , 'lcid' , 'RHS' , ), 330, (330, (), [ (3, 5, None, None) , 
			 (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 392 , (3, 0, None, None) , 0 , )),
	(( 'CutCopyMode' , 'lcid' , 'RHS' , ), 330, (330, (), [ (3, 5, None, None) , 
			 (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 396 , (3, 0, None, None) , 0 , )),
	(( 'DataEntryMode' , 'lcid' , 'RHS' , ), 331, (331, (), [ (3, 5, None, None) , 
			 (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 400 , (3, 0, None, None) , 0 , )),
	(( 'DataEntryMode' , 'lcid' , 'RHS' , ), 331, (331, (), [ (3, 5, None, None) , 
			 (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 404 , (3, 0, None, None) , 0 , )),
	(( 'Dummy1' , 'Arg1' , 'Arg2' , 'Arg3' , 'Arg4' , 
			 'RHS' , ), 1782, (1782, (), [ (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (16396, 10, None, None) , ], 1 , 1 , 4 , 4 , 408 , (3, 0, None, None) , 64 , )),
	(( 'Dummy2' , 'Arg1' , 'Arg2' , 'Arg3' , 'Arg4' , 
			 'Arg5' , 'Arg6' , 'Arg7' , 'Arg8' , 'RHS' , 
			 ), 1783, (1783, (), [ (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (16396, 10, None, None) , ], 1 , 1 , 4 , 8 , 412 , (3, 0, None, None) , 64 , )),
	(( 'Dummy3' , 'RHS' , ), 1784, (1784, (), [ (16396, 10, None, None) , ], 1 , 1 , 4 , 0 , 416 , (3, 0, None, None) , 64 , )),
	(( 'Dummy4' , 'Arg1' , 'Arg2' , 'Arg3' , 'Arg4' , 
			 'Arg5' , 'Arg6' , 'Arg7' , 'Arg8' , 'Arg9' , 
			 'Arg10' , 'Arg11' , 'Arg12' , 'Arg13' , 'Arg14' , 
			 'Arg15' , 'RHS' , ), 1785, (1785, (), [ (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (16396, 10, None, None) , ], 1 , 1 , 4 , 15 , 420 , (3, 0, None, None) , 64 , )),
	(( 'Dummy5' , 'Arg1' , 'Arg2' , 'Arg3' , 'Arg4' , 
			 'Arg5' , 'Arg6' , 'Arg7' , 'Arg8' , 'Arg9' , 
			 'Arg10' , 'Arg11' , 'Arg12' , 'Arg13' , 'RHS' , 
			 ), 1786, (1786, (), [ (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (16396, 10, None, None) , ], 1 , 1 , 4 , 13 , 424 , (3, 0, None, None) , 64 , )),
	(( 'Dummy6' , 'RHS' , ), 1787, (1787, (), [ (16396, 10, None, None) , ], 1 , 1 , 4 , 0 , 428 , (3, 0, None, None) , 64 , )),
	(( 'Dummy7' , 'RHS' , ), 1788, (1788, (), [ (16396, 10, None, None) , ], 1 , 1 , 4 , 0 , 432 , (3, 0, None, None) , 64 , )),
	(( 'Dummy8' , 'Arg1' , 'RHS' , ), 1789, (1789, (), [ (12, 17, None, None) , 
			 (16396, 10, None, None) , ], 1 , 1 , 4 , 1 , 436 , (3, 0, None, None) , 64 , )),
	(( 'Dummy9' , 'RHS' , ), 1790, (1790, (), [ (16396, 10, None, None) , ], 1 , 1 , 4 , 0 , 440 , (3, 0, None, None) , 64 , )),
	(( 'Dummy10' , 'arg' , 'RHS' , ), 1791, (1791, (), [ (12, 17, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 1 , 444 , (3, 0, None, None) , 64 , )),
	(( 'Dummy11' , ), 1792, (1792, (), [ ], 1 , 1 , 4 , 0 , 448 , (3, 0, None, None) , 64 , )),
	(( '_Default' , 'RHS' , ), 0, (0, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 452 , (3, 0, None, None) , 1024 , )),
	(( 'DefaultFilePath' , 'lcid' , 'RHS' , ), 1038, (1038, (), [ (3, 5, None, None) , 
			 (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 456 , (3, 0, None, None) , 0 , )),
	(( 'DefaultFilePath' , 'lcid' , 'RHS' , ), 1038, (1038, (), [ (3, 5, None, None) , 
			 (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 460 , (3, 0, None, None) , 0 , )),
	(( 'DeleteChartAutoFormat' , 'Name' , 'lcid' , ), 217, (217, (), [ (8, 1, None, None) , 
			 (3, 5, None, None) , ], 1 , 1 , 4 , 0 , 464 , (3, 0, None, None) , 64 , )),
	(( 'DeleteCustomList' , 'ListNum' , 'lcid' , ), 783, (783, (), [ (3, 1, None, None) , 
			 (3, 5, None, None) , ], 1 , 1 , 4 , 0 , 468 , (3, 0, None, None) , 0 , )),
	(( 'Dialogs' , 'RHS' , ), 761, (761, (), [ (16393, 10, None, "IID('{00020879-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 472 , (3, 0, None, None) , 0 , )),
	(( 'DisplayAlerts' , 'lcid' , 'RHS' , ), 343, (343, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 476 , (3, 0, None, None) , 0 , )),
	(( 'DisplayAlerts' , 'lcid' , 'RHS' , ), 343, (343, (), [ (3, 5, None, None) , 
			 (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 480 , (3, 0, None, None) , 0 , )),
	(( 'DisplayFormulaBar' , 'lcid' , 'RHS' , ), 344, (344, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 484 , (3, 0, None, None) , 0 , )),
	(( 'DisplayFormulaBar' , 'lcid' , 'RHS' , ), 344, (344, (), [ (3, 5, None, None) , 
			 (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 488 , (3, 0, None, None) , 0 , )),
	(( 'DisplayFullScreen' , 'lcid' , 'RHS' , ), 1061, (1061, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 492 , (3, 0, None, None) , 0 , )),
	(( 'DisplayFullScreen' , 'lcid' , 'RHS' , ), 1061, (1061, (), [ (3, 5, None, None) , 
			 (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 496 , (3, 0, None, None) , 0 , )),
	(( 'DisplayNoteIndicator' , 'RHS' , ), 345, (345, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 500 , (3, 0, None, None) , 0 , )),
	(( 'DisplayNoteIndicator' , 'RHS' , ), 345, (345, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 504 , (3, 0, None, None) , 0 , )),
	(( 'DisplayCommentIndicator' , 'RHS' , ), 1196, (1196, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 508 , (3, 0, None, None) , 0 , )),
	(( 'DisplayCommentIndicator' , 'RHS' , ), 1196, (1196, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 512 , (3, 0, None, None) , 0 , )),
	(( 'DisplayExcel4Menus' , 'lcid' , 'RHS' , ), 927, (927, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 516 , (3, 0, None, None) , 0 , )),
	(( 'DisplayExcel4Menus' , 'lcid' , 'RHS' , ), 927, (927, (), [ (3, 5, None, None) , 
			 (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 520 , (3, 0, None, None) , 0 , )),
	(( 'DisplayRecentFiles' , 'RHS' , ), 926, (926, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 524 , (3, 0, None, None) , 0 , )),
	(( 'DisplayRecentFiles' , 'RHS' , ), 926, (926, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 528 , (3, 0, None, None) , 0 , )),
	(( 'DisplayScrollBars' , 'lcid' , 'RHS' , ), 346, (346, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 532 , (3, 0, None, None) , 0 , )),
	(( 'DisplayScrollBars' , 'lcid' , 'RHS' , ), 346, (346, (), [ (3, 5, None, None) , 
			 (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 536 , (3, 0, None, None) , 0 , )),
	(( 'DisplayStatusBar' , 'lcid' , 'RHS' , ), 347, (347, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 540 , (3, 0, None, None) , 0 , )),
	(( 'DisplayStatusBar' , 'lcid' , 'RHS' , ), 347, (347, (), [ (3, 5, None, None) , 
			 (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 544 , (3, 0, None, None) , 0 , )),
	(( 'DoubleClick' , 'lcid' , ), 349, (349, (), [ (3, 5, None, None) , ], 1 , 1 , 4 , 0 , 548 , (3, 0, None, None) , 0 , )),
	(( 'EditDirectlyInCell' , 'lcid' , 'RHS' , ), 929, (929, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 552 , (3, 0, None, None) , 0 , )),
	(( 'EditDirectlyInCell' , 'lcid' , 'RHS' , ), 929, (929, (), [ (3, 5, None, None) , 
			 (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 556 , (3, 0, None, None) , 0 , )),
	(( 'EnableAutoComplete' , 'RHS' , ), 1179, (1179, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 560 , (3, 0, None, None) , 0 , )),
	(( 'EnableAutoComplete' , 'RHS' , ), 1179, (1179, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 564 , (3, 0, None, None) , 0 , )),
	(( 'EnableCancelKey' , 'lcid' , 'RHS' , ), 1096, (1096, (), [ (3, 5, None, None) , 
			 (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 568 , (3, 0, None, None) , 0 , )),
	(( 'EnableCancelKey' , 'lcid' , 'RHS' , ), 1096, (1096, (), [ (3, 5, None, None) , 
			 (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 572 , (3, 0, None, None) , 0 , )),
	(( 'EnableSound' , 'RHS' , ), 1197, (1197, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 576 , (3, 0, None, None) , 0 , )),
	(( 'EnableSound' , 'RHS' , ), 1197, (1197, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 580 , (3, 0, None, None) , 0 , )),
	(( 'EnableTipWizard' , 'lcid' , 'RHS' , ), 1064, (1064, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 584 , (3, 0, None, None) , 64 , )),
	(( 'EnableTipWizard' , 'lcid' , 'RHS' , ), 1064, (1064, (), [ (3, 5, None, None) , 
			 (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 588 , (3, 0, None, None) , 64 , )),
	(( 'FileConverters' , 'Index1' , 'Index2' , 'lcid' , 'RHS' , 
			 ), 931, (931, (), [ (12, 17, None, None) , (12, 17, None, None) , (3, 5, None, None) , (16396, 10, None, None) , ], 1 , 2 , 4 , 2 , 592 , (3, 0, None, None) , 0 , )),
	(( 'FileSearch' , 'RHS' , ), 1200, (1200, (), [ (16393, 10, None, "IID('{000C0332-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 596 , (3, 0, None, None) , 64 , )),
	(( 'FileFind' , 'RHS' , ), 1201, (1201, (), [ (16393, 10, None, "IID('{000C0337-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 600 , (3, 0, None, None) , 64 , )),
	(( '_FindFile' , 'lcid' , ), 1068, (1068, (), [ (3, 5, None, None) , ], 1 , 1 , 4 , 0 , 604 , (3, 0, None, None) , 1088 , )),
	(( 'FixedDecimal' , 'lcid' , 'RHS' , ), 351, (351, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 608 , (3, 0, None, None) , 0 , )),
	(( 'FixedDecimal' , 'lcid' , 'RHS' , ), 351, (351, (), [ (3, 5, None, None) , 
			 (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 612 , (3, 0, None, None) , 0 , )),
	(( 'FixedDecimalPlaces' , 'lcid' , 'RHS' , ), 352, (352, (), [ (3, 5, None, None) , 
			 (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 616 , (3, 0, None, None) , 0 , )),
	(( 'FixedDecimalPlaces' , 'lcid' , 'RHS' , ), 352, (352, (), [ (3, 5, None, None) , 
			 (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 620 , (3, 0, None, None) , 0 , )),
	(( 'GetCustomListContents' , 'ListNum' , 'lcid' , 'RHS' , ), 786, (786, (), [ 
			 (3, 1, None, None) , (3, 5, None, None) , (16396, 10, None, None) , ], 1 , 1 , 4 , 0 , 624 , (3, 0, None, None) , 0 , )),
	(( 'GetCustomListNum' , 'ListArray' , 'lcid' , 'RHS' , ), 785, (785, (), [ 
			 (12, 1, None, None) , (3, 5, None, None) , (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 628 , (3, 0, None, None) , 0 , )),
	(( 'GetOpenFilename' , 'FileFilter' , 'FilterIndex' , 'Title' , 'ButtonText' , 
			 'MultiSelect' , 'lcid' , 'RHS' , ), 1075, (1075, (), [ (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (3, 5, None, None) , 
			 (16396, 10, None, None) , ], 1 , 1 , 4 , 5 , 632 , (3, 0, None, None) , 0 , )),
	(( 'GetSaveAsFilename' , 'InitialFilename' , 'FileFilter' , 'FilterIndex' , 'Title' , 
			 'ButtonText' , 'lcid' , 'RHS' , ), 1076, (1076, (), [ (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (3, 5, None, None) , 
			 (16396, 10, None, None) , ], 1 , 1 , 4 , 5 , 636 , (3, 0, None, None) , 0 , )),
	(( 'Goto' , 'Reference' , 'Scroll' , 'lcid' , ), 475, (475, (), [ 
			 (12, 17, None, None) , (12, 17, None, None) , (3, 5, None, None) , ], 1 , 1 , 4 , 2 , 640 , (3, 0, None, None) , 0 , )),
	(( 'Height' , 'lcid' , 'RHS' , ), 123, (123, (), [ (3, 5, None, None) , 
			 (16389, 10, None, None) , ], 1 , 2 , 4 , 0 , 644 , (3, 0, None, None) , 0 , )),
	(( 'Height' , 'lcid' , 'RHS' , ), 123, (123, (), [ (3, 5, None, None) , 
			 (5, 1, None, None) , ], 1 , 4 , 4 , 0 , 648 , (3, 0, None, None) , 0 , )),
	(( 'Help' , 'HelpFile' , 'HelpContextID' , 'lcid' , ), 354, (354, (), [ 
			 (12, 17, None, None) , (12, 17, None, None) , (3, 5, None, None) , ], 1 , 1 , 4 , 2 , 652 , (3, 0, None, None) , 0 , )),
	(( 'IgnoreRemoteRequests' , 'lcid' , 'RHS' , ), 356, (356, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 656 , (3, 0, None, None) , 0 , )),
	(( 'IgnoreRemoteRequests' , 'lcid' , 'RHS' , ), 356, (356, (), [ (3, 5, None, None) , 
			 (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 660 , (3, 0, None, None) , 0 , )),
	(( 'InchesToPoints' , 'Inches' , 'lcid' , 'RHS' , ), 1087, (1087, (), [ 
			 (5, 1, None, None) , (3, 5, None, None) , (16389, 10, None, None) , ], 1 , 1 , 4 , 0 , 664 , (3, 0, None, None) , 0 , )),
	(( 'InputBox' , 'Prompt' , 'Title' , 'Default' , 'Left' , 
			 'Top' , 'HelpFile' , 'HelpContextID' , 'Type' , 'lcid' , 
			 'RHS' , ), 357, (357, (), [ (8, 1, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (3, 5, None, None) , (16396, 10, None, None) , ], 1 , 1 , 4 , 7 , 668 , (3, 0, None, None) , 0 , )),
	(( 'Interactive' , 'lcid' , 'RHS' , ), 361, (361, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 672 , (3, 0, None, None) , 0 , )),
	(( 'Interactive' , 'lcid' , 'RHS' , ), 361, (361, (), [ (3, 5, None, None) , 
			 (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 676 , (3, 0, None, None) , 0 , )),
	(( 'International' , 'Index' , 'lcid' , 'RHS' , ), 362, (362, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16396, 10, None, None) , ], 1 , 2 , 4 , 1 , 680 , (3, 0, None, None) , 0 , )),
	(( 'Iteration' , 'lcid' , 'RHS' , ), 363, (363, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 684 , (3, 0, None, None) , 0 , )),
	(( 'Iteration' , 'lcid' , 'RHS' , ), 363, (363, (), [ (3, 5, None, None) , 
			 (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 688 , (3, 0, None, None) , 0 , )),
	(( 'LargeButtons' , 'RHS' , ), 364, (364, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 692 , (3, 0, None, None) , 64 , )),
	(( 'LargeButtons' , 'RHS' , ), 364, (364, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 696 , (3, 0, None, None) , 64 , )),
	(( 'Left' , 'lcid' , 'RHS' , ), 127, (127, (), [ (3, 5, None, None) , 
			 (16389, 10, None, None) , ], 1 , 2 , 4 , 0 , 700 , (3, 0, None, None) , 0 , )),
	(( 'Left' , 'lcid' , 'RHS' , ), 127, (127, (), [ (3, 5, None, None) , 
			 (5, 1, None, None) , ], 1 , 4 , 4 , 0 , 704 , (3, 0, None, None) , 0 , )),
	(( 'LibraryPath' , 'lcid' , 'RHS' , ), 366, (366, (), [ (3, 5, None, None) , 
			 (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 708 , (3, 0, None, None) , 0 , )),
	(( '_MacroOptions' , 'Macro' , 'Description' , 'HasMenu' , 'MenuText' , 
			 'HasShortcutKey' , 'ShortcutKey' , 'Category' , 'StatusBar' , 'HelpContextID' , 
			 'HelpFile' , 'lcid' , ), 1135, (1135, (), [ (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (3, 5, None, None) , ], 1 , 1 , 4 , 10 , 712 , (3, 0, None, None) , 1088 , )),
	(( 'MailLogoff' , 'lcid' , ), 945, (945, (), [ (3, 5, None, None) , ], 1 , 1 , 4 , 0 , 716 , (3, 0, None, None) , 0 , )),
	(( 'MailLogon' , 'Name' , 'Password' , 'DownloadNewMail' , 'lcid' , 
			 ), 943, (943, (), [ (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (3, 5, None, None) , ], 1 , 1 , 4 , 3 , 720 , (3, 0, None, None) , 0 , )),
	(( 'MailSession' , 'lcid' , 'RHS' , ), 942, (942, (), [ (3, 5, None, None) , 
			 (16396, 10, None, None) , ], 1 , 2 , 4 , 0 , 724 , (3, 0, None, None) , 0 , )),
	(( 'MailSystem' , 'lcid' , 'RHS' , ), 971, (971, (), [ (3, 5, None, None) , 
			 (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 728 , (3, 0, None, None) , 0 , )),
	(( 'MathCoprocessorAvailable' , 'lcid' , 'RHS' , ), 367, (367, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 732 , (3, 0, None, None) , 0 , )),
	(( 'MaxChange' , 'lcid' , 'RHS' , ), 368, (368, (), [ (3, 5, None, None) , 
			 (16389, 10, None, None) , ], 1 , 2 , 4 , 0 , 736 , (3, 0, None, None) , 0 , )),
	(( 'MaxChange' , 'lcid' , 'RHS' , ), 368, (368, (), [ (3, 5, None, None) , 
			 (5, 1, None, None) , ], 1 , 4 , 4 , 0 , 740 , (3, 0, None, None) , 0 , )),
	(( 'MaxIterations' , 'lcid' , 'RHS' , ), 369, (369, (), [ (3, 5, None, None) , 
			 (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 744 , (3, 0, None, None) , 0 , )),
	(( 'MaxIterations' , 'lcid' , 'RHS' , ), 369, (369, (), [ (3, 5, None, None) , 
			 (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 748 , (3, 0, None, None) , 0 , )),
	(( 'MemoryFree' , 'lcid' , 'RHS' , ), 370, (370, (), [ (3, 5, None, None) , 
			 (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 752 , (3, 0, None, None) , 64 , )),
	(( 'MemoryTotal' , 'lcid' , 'RHS' , ), 371, (371, (), [ (3, 5, None, None) , 
			 (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 756 , (3, 0, None, None) , 64 , )),
	(( 'MemoryUsed' , 'lcid' , 'RHS' , ), 372, (372, (), [ (3, 5, None, None) , 
			 (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 760 , (3, 0, None, None) , 64 , )),
	(( 'MouseAvailable' , 'lcid' , 'RHS' , ), 373, (373, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 764 , (3, 0, None, None) , 0 , )),
	(( 'MoveAfterReturn' , 'lcid' , 'RHS' , ), 374, (374, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 768 , (3, 0, None, None) , 0 , )),
	(( 'MoveAfterReturn' , 'lcid' , 'RHS' , ), 374, (374, (), [ (3, 5, None, None) , 
			 (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 772 , (3, 0, None, None) , 0 , )),
	(( 'MoveAfterReturnDirection' , 'lcid' , 'RHS' , ), 1144, (1144, (), [ (3, 5, None, None) , 
			 (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 776 , (3, 0, None, None) , 0 , )),
	(( 'MoveAfterReturnDirection' , 'lcid' , 'RHS' , ), 1144, (1144, (), [ (3, 5, None, None) , 
			 (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 780 , (3, 0, None, None) , 0 , )),
	(( 'RecentFiles' , 'RHS' , ), 1202, (1202, (), [ (16393, 10, None, "IID('{00024406-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 784 , (3, 0, None, None) , 0 , )),
	(( 'Name' , 'RHS' , ), 110, (110, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 788 , (3, 0, None, None) , 0 , )),
	(( 'NextLetter' , 'lcid' , 'RHS' , ), 972, (972, (), [ (3, 5, None, None) , 
			 (16397, 10, None, "IID('{00020819-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 0 , 792 , (3, 0, None, None) , 0 , )),
	(( 'NetworkTemplatesPath' , 'lcid' , 'RHS' , ), 388, (388, (), [ (3, 5, None, None) , 
			 (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 796 , (3, 0, None, None) , 0 , )),
	(( 'ODBCErrors' , 'RHS' , ), 1203, (1203, (), [ (16393, 10, None, "IID('{0002442D-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 800 , (3, 0, None, None) , 0 , )),
	(( 'ODBCTimeout' , 'RHS' , ), 1204, (1204, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 804 , (3, 0, None, None) , 0 , )),
	(( 'ODBCTimeout' , 'RHS' , ), 1204, (1204, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 808 , (3, 0, None, None) , 0 , )),
	(( 'OnCalculate' , 'lcid' , 'RHS' , ), 625, (625, (), [ (3, 5, None, None) , 
			 (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 812 , (3, 0, None, None) , 64 , )),
	(( 'OnCalculate' , 'lcid' , 'RHS' , ), 625, (625, (), [ (3, 5, None, None) , 
			 (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 816 , (3, 0, None, None) , 64 , )),
	(( 'OnData' , 'lcid' , 'RHS' , ), 629, (629, (), [ (3, 5, None, None) , 
			 (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 820 , (3, 0, None, None) , 64 , )),
	(( 'OnData' , 'lcid' , 'RHS' , ), 629, (629, (), [ (3, 5, None, None) , 
			 (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 824 , (3, 0, None, None) , 64 , )),
	(( 'OnDoubleClick' , 'lcid' , 'RHS' , ), 628, (628, (), [ (3, 5, None, None) , 
			 (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 828 , (3, 0, None, None) , 64 , )),
	(( 'OnDoubleClick' , 'lcid' , 'RHS' , ), 628, (628, (), [ (3, 5, None, None) , 
			 (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 832 , (3, 0, None, None) , 64 , )),
	(( 'OnEntry' , 'lcid' , 'RHS' , ), 627, (627, (), [ (3, 5, None, None) , 
			 (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 836 , (3, 0, None, None) , 64 , )),
	(( 'OnEntry' , 'lcid' , 'RHS' , ), 627, (627, (), [ (3, 5, None, None) , 
			 (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 840 , (3, 0, None, None) , 64 , )),
	(( 'OnKey' , 'Key' , 'Procedure' , 'lcid' , ), 626, (626, (), [ 
			 (8, 1, None, None) , (12, 17, None, None) , (3, 5, None, None) , ], 1 , 1 , 4 , 1 , 844 , (3, 0, None, None) , 0 , )),
	(( 'OnRepeat' , 'Text' , 'Procedure' , 'lcid' , ), 769, (769, (), [ 
			 (8, 1, None, None) , (8, 1, None, None) , (3, 5, None, None) , ], 1 , 1 , 4 , 0 , 848 , (3, 0, None, None) , 0 , )),
	(( 'OnSheetActivate' , 'lcid' , 'RHS' , ), 1031, (1031, (), [ (3, 5, None, None) , 
			 (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 852 , (3, 0, None, None) , 64 , )),
	(( 'OnSheetActivate' , 'lcid' , 'RHS' , ), 1031, (1031, (), [ (3, 5, None, None) , 
			 (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 856 , (3, 0, None, None) , 64 , )),
	(( 'OnSheetDeactivate' , 'lcid' , 'RHS' , ), 1081, (1081, (), [ (3, 5, None, None) , 
			 (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 860 , (3, 0, None, None) , 64 , )),
	(( 'OnSheetDeactivate' , 'lcid' , 'RHS' , ), 1081, (1081, (), [ (3, 5, None, None) , 
			 (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 864 , (3, 0, None, None) , 64 , )),
	(( 'OnTime' , 'EarliestTime' , 'Procedure' , 'LatestTime' , 'Schedule' , 
			 'lcid' , ), 624, (624, (), [ (12, 1, None, None) , (8, 1, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (3, 5, None, None) , ], 1 , 1 , 4 , 2 , 868 , (3, 0, None, None) , 0 , )),
	(( 'OnUndo' , 'Text' , 'Procedure' , 'lcid' , ), 770, (770, (), [ 
			 (8, 1, None, None) , (8, 1, None, None) , (3, 5, None, None) , ], 1 , 1 , 4 , 0 , 872 , (3, 0, None, None) , 0 , )),
	(( 'OnWindow' , 'lcid' , 'RHS' , ), 623, (623, (), [ (3, 5, None, None) , 
			 (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 876 , (3, 0, None, None) , 0 , )),
	(( 'OnWindow' , 'lcid' , 'RHS' , ), 623, (623, (), [ (3, 5, None, None) , 
			 (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 880 , (3, 0, None, None) , 0 , )),
	(( 'OperatingSystem' , 'lcid' , 'RHS' , ), 375, (375, (), [ (3, 5, None, None) , 
			 (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 884 , (3, 0, None, None) , 0 , )),
	(( 'OrganizationName' , 'lcid' , 'RHS' , ), 376, (376, (), [ (3, 5, None, None) , 
			 (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 888 , (3, 0, None, None) , 0 , )),
	(( 'Path' , 'lcid' , 'RHS' , ), 291, (291, (), [ (3, 5, None, None) , 
			 (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 892 , (3, 0, None, None) , 0 , )),
	(( 'PathSeparator' , 'lcid' , 'RHS' , ), 377, (377, (), [ (3, 5, None, None) , 
			 (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 896 , (3, 0, None, None) , 0 , )),
	(( 'PreviousSelections' , 'Index' , 'lcid' , 'RHS' , ), 378, (378, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16396, 10, None, None) , ], 1 , 2 , 4 , 1 , 900 , (3, 0, None, None) , 0 , )),
	(( 'PivotTableSelection' , 'RHS' , ), 1205, (1205, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 904 , (3, 0, None, None) , 0 , )),
	(( 'PivotTableSelection' , 'RHS' , ), 1205, (1205, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 908 , (3, 0, None, None) , 0 , )),
	(( 'PromptForSummaryInfo' , 'lcid' , 'RHS' , ), 1062, (1062, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 912 , (3, 0, None, None) , 0 , )),
	(( 'PromptForSummaryInfo' , 'lcid' , 'RHS' , ), 1062, (1062, (), [ (3, 5, None, None) , 
			 (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 916 , (3, 0, None, None) , 0 , )),
	(( 'Quit' , ), 302, (302, (), [ ], 1 , 1 , 4 , 0 , 920 , (3, 0, None, None) , 0 , )),
	(( 'RecordMacro' , 'BasicCode' , 'XlmCode' , 'lcid' , ), 773, (773, (), [ 
			 (12, 17, None, None) , (12, 17, None, None) , (3, 5, None, None) , ], 1 , 1 , 4 , 2 , 924 , (3, 0, None, None) , 0 , )),
	(( 'RecordRelative' , 'lcid' , 'RHS' , ), 379, (379, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 928 , (3, 0, None, None) , 0 , )),
	(( 'ReferenceStyle' , 'lcid' , 'RHS' , ), 380, (380, (), [ (3, 5, None, None) , 
			 (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 932 , (3, 0, None, None) , 0 , )),
	(( 'ReferenceStyle' , 'lcid' , 'RHS' , ), 380, (380, (), [ (3, 5, None, None) , 
			 (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 936 , (3, 0, None, None) , 0 , )),
	(( 'RegisteredFunctions' , 'Index1' , 'Index2' , 'lcid' , 'RHS' , 
			 ), 775, (775, (), [ (12, 17, None, None) , (12, 17, None, None) , (3, 5, None, None) , (16396, 10, None, None) , ], 1 , 2 , 4 , 2 , 940 , (3, 0, None, None) , 0 , )),
	(( 'RegisterXLL' , 'Filename' , 'lcid' , 'RHS' , ), 30, (30, (), [ 
			 (8, 1, None, None) , (3, 5, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 944 , (3, 0, None, None) , 0 , )),
	(( 'Repeat' , 'lcid' , ), 301, (301, (), [ (3, 5, None, None) , ], 1 , 1 , 4 , 0 , 948 , (3, 0, None, None) , 0 , )),
	(( 'ResetTipWizard' , 'lcid' , ), 928, (928, (), [ (3, 5, None, None) , ], 1 , 1 , 4 , 0 , 952 , (3, 0, None, None) , 64 , )),
	(( 'RollZoom' , 'RHS' , ), 1206, (1206, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 956 , (3, 0, None, None) , 0 , )),
	(( 'RollZoom' , 'RHS' , ), 1206, (1206, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 960 , (3, 0, None, None) , 0 , )),
	(( 'Save' , 'Filename' , 'lcid' , ), 283, (283, (), [ (12, 17, None, None) , 
			 (3, 5, None, None) , ], 1 , 1 , 4 , 1 , 964 , (3, 0, None, None) , 64 , )),
	(( 'SaveWorkspace' , 'Filename' , 'lcid' , ), 212, (212, (), [ (12, 17, None, None) , 
			 (3, 5, None, None) , ], 1 , 1 , 4 , 1 , 968 , (3, 0, None, None) , 0 , )),
	(( 'ScreenUpdating' , 'lcid' , 'RHS' , ), 382, (382, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 972 , (3, 0, None, None) , 0 , )),
	(( 'ScreenUpdating' , 'lcid' , 'RHS' , ), 382, (382, (), [ (3, 5, None, None) , 
			 (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 976 , (3, 0, None, None) , 0 , )),
	(( 'SetDefaultChart' , 'FormatName' , 'Gallery' , ), 219, (219, (), [ (12, 17, None, None) , 
			 (12, 17, None, None) , ], 1 , 1 , 4 , 2 , 980 , (3, 0, None, None) , 64 , )),
	(( 'SheetsInNewWorkbook' , 'lcid' , 'RHS' , ), 993, (993, (), [ (3, 5, None, None) , 
			 (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 984 , (3, 0, None, None) , 0 , )),
	(( 'SheetsInNewWorkbook' , 'lcid' , 'RHS' , ), 993, (993, (), [ (3, 5, None, None) , 
			 (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 988 , (3, 0, None, None) , 0 , )),
	(( 'ShowChartTipNames' , 'RHS' , ), 1207, (1207, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 992 , (3, 0, None, None) , 0 , )),
	(( 'ShowChartTipNames' , 'RHS' , ), 1207, (1207, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 996 , (3, 0, None, None) , 0 , )),
	(( 'ShowChartTipValues' , 'RHS' , ), 1208, (1208, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1000 , (3, 0, None, None) , 0 , )),
	(( 'ShowChartTipValues' , 'RHS' , ), 1208, (1208, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1004 , (3, 0, None, None) , 0 , )),
	(( 'StandardFont' , 'lcid' , 'RHS' , ), 924, (924, (), [ (3, 5, None, None) , 
			 (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 1008 , (3, 0, None, None) , 0 , )),
	(( 'StandardFont' , 'lcid' , 'RHS' , ), 924, (924, (), [ (3, 5, None, None) , 
			 (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 1012 , (3, 0, None, None) , 0 , )),
	(( 'StandardFontSize' , 'lcid' , 'RHS' , ), 925, (925, (), [ (3, 5, None, None) , 
			 (16389, 10, None, None) , ], 1 , 2 , 4 , 0 , 1016 , (3, 0, None, None) , 0 , )),
	(( 'StandardFontSize' , 'lcid' , 'RHS' , ), 925, (925, (), [ (3, 5, None, None) , 
			 (5, 1, None, None) , ], 1 , 4 , 4 , 0 , 1020 , (3, 0, None, None) , 0 , )),
	(( 'StartupPath' , 'lcid' , 'RHS' , ), 385, (385, (), [ (3, 5, None, None) , 
			 (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 1024 , (3, 0, None, None) , 0 , )),
	(( 'StatusBar' , 'lcid' , 'RHS' , ), 386, (386, (), [ (3, 5, None, None) , 
			 (16396, 10, None, None) , ], 1 , 2 , 4 , 0 , 1028 , (3, 0, None, None) , 0 , )),
	(( 'StatusBar' , 'lcid' , 'RHS' , ), 386, (386, (), [ (3, 5, None, None) , 
			 (12, 1, None, None) , ], 1 , 4 , 4 , 0 , 1032 , (3, 0, None, None) , 0 , )),
	(( 'TemplatesPath' , 'lcid' , 'RHS' , ), 381, (381, (), [ (3, 5, None, None) , 
			 (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 1036 , (3, 0, None, None) , 0 , )),
	(( 'ShowToolTips' , 'RHS' , ), 387, (387, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1040 , (3, 0, None, None) , 0 , )),
	(( 'ShowToolTips' , 'RHS' , ), 387, (387, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1044 , (3, 0, None, None) , 0 , )),
	(( 'Top' , 'lcid' , 'RHS' , ), 126, (126, (), [ (3, 5, None, None) , 
			 (16389, 10, None, None) , ], 1 , 2 , 4 , 0 , 1048 , (3, 0, None, None) , 0 , )),
	(( 'Top' , 'lcid' , 'RHS' , ), 126, (126, (), [ (3, 5, None, None) , 
			 (5, 1, None, None) , ], 1 , 4 , 4 , 0 , 1052 , (3, 0, None, None) , 0 , )),
	(( 'DefaultSaveFormat' , 'RHS' , ), 1209, (1209, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 1056 , (3, 0, None, None) , 0 , )),
	(( 'DefaultSaveFormat' , 'RHS' , ), 1209, (1209, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 1060 , (3, 0, None, None) , 0 , )),
	(( 'TransitionMenuKey' , 'lcid' , 'RHS' , ), 310, (310, (), [ (3, 5, None, None) , 
			 (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 1064 , (3, 0, None, None) , 0 , )),
	(( 'TransitionMenuKey' , 'lcid' , 'RHS' , ), 310, (310, (), [ (3, 5, None, None) , 
			 (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 1068 , (3, 0, None, None) , 0 , )),
	(( 'TransitionMenuKeyAction' , 'lcid' , 'RHS' , ), 311, (311, (), [ (3, 5, None, None) , 
			 (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 1072 , (3, 0, None, None) , 0 , )),
	(( 'TransitionMenuKeyAction' , 'lcid' , 'RHS' , ), 311, (311, (), [ (3, 5, None, None) , 
			 (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 1076 , (3, 0, None, None) , 0 , )),
	(( 'TransitionNavigKeys' , 'lcid' , 'RHS' , ), 312, (312, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1080 , (3, 0, None, None) , 0 , )),
	(( 'TransitionNavigKeys' , 'lcid' , 'RHS' , ), 312, (312, (), [ (3, 5, None, None) , 
			 (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1084 , (3, 0, None, None) , 0 , )),
	(( 'Undo' , 'lcid' , ), 303, (303, (), [ (3, 5, None, None) , ], 1 , 1 , 4 , 0 , 1088 , (3, 0, None, None) , 0 , )),
	(( 'UsableHeight' , 'lcid' , 'RHS' , ), 389, (389, (), [ (3, 5, None, None) , 
			 (16389, 10, None, None) , ], 1 , 2 , 4 , 0 , 1092 , (3, 0, None, None) , 0 , )),
	(( 'UsableWidth' , 'lcid' , 'RHS' , ), 390, (390, (), [ (3, 5, None, None) , 
			 (16389, 10, None, None) , ], 1 , 2 , 4 , 0 , 1096 , (3, 0, None, None) , 0 , )),
	(( 'UserControl' , 'RHS' , ), 1210, (1210, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1100 , (3, 0, None, None) , 0 , )),
	(( 'UserControl' , 'RHS' , ), 1210, (1210, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1104 , (3, 0, None, None) , 0 , )),
	(( 'UserName' , 'lcid' , 'RHS' , ), 391, (391, (), [ (3, 5, None, None) , 
			 (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 1108 , (3, 0, None, None) , 0 , )),
	(( 'UserName' , 'lcid' , 'RHS' , ), 391, (391, (), [ (3, 5, None, None) , 
			 (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 1112 , (3, 0, None, None) , 0 , )),
	(( 'Value' , 'RHS' , ), 6, (6, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 1116 , (3, 0, None, None) , 0 , )),
	(( 'VBE' , 'RHS' , ), 1211, (1211, (), [ (16393, 10, None, "IID('{0002E166-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 1120 , (3, 0, None, None) , 0 , )),
	(( 'Version' , 'lcid' , 'RHS' , ), 392, (392, (), [ (3, 5, None, None) , 
			 (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 1124 , (3, 0, None, None) , 0 , )),
	(( 'Visible' , 'lcid' , 'RHS' , ), 558, (558, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1128 , (3, 0, None, None) , 0 , )),
	(( 'Visible' , 'lcid' , 'RHS' , ), 558, (558, (), [ (3, 5, None, None) , 
			 (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1132 , (3, 0, None, None) , 0 , )),
	(( 'Volatile' , 'Volatile' , 'lcid' , ), 788, (788, (), [ (12, 17, None, None) , 
			 (3, 5, None, None) , ], 1 , 1 , 4 , 1 , 1136 , (3, 0, None, None) , 0 , )),
	(( '_Wait' , 'Time' , 'lcid' , ), 393, (393, (), [ (12, 1, None, None) , 
			 (3, 5, None, None) , ], 1 , 1 , 4 , 0 , 1140 , (3, 0, None, None) , 1088 , )),
	(( 'Width' , 'lcid' , 'RHS' , ), 122, (122, (), [ (3, 5, None, None) , 
			 (16389, 10, None, None) , ], 1 , 2 , 4 , 0 , 1144 , (3, 0, None, None) , 0 , )),
	(( 'Width' , 'lcid' , 'RHS' , ), 122, (122, (), [ (3, 5, None, None) , 
			 (5, 1, None, None) , ], 1 , 4 , 4 , 0 , 1148 , (3, 0, None, None) , 0 , )),
	(( 'WindowsForPens' , 'lcid' , 'RHS' , ), 395, (395, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1152 , (3, 0, None, None) , 0 , )),
	(( 'WindowState' , 'lcid' , 'RHS' , ), 396, (396, (), [ (3, 5, None, None) , 
			 (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 1156 , (3, 0, None, None) , 0 , )),
	(( 'WindowState' , 'lcid' , 'RHS' , ), 396, (396, (), [ (3, 5, None, None) , 
			 (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 1160 , (3, 0, None, None) , 0 , )),
	(( 'UILanguage' , 'lcid' , 'RHS' , ), 2, (2, (), [ (3, 5, None, None) , 
			 (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 1164 , (3, 0, None, None) , 64 , )),
	(( 'UILanguage' , 'lcid' , 'RHS' , ), 2, (2, (), [ (3, 5, None, None) , 
			 (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 1168 , (3, 0, None, None) , 64 , )),
	(( 'DefaultSheetDirection' , 'lcid' , 'RHS' , ), 229, (229, (), [ (3, 5, None, None) , 
			 (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 1172 , (3, 0, None, None) , 0 , )),
	(( 'DefaultSheetDirection' , 'lcid' , 'RHS' , ), 229, (229, (), [ (3, 5, None, None) , 
			 (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 1176 , (3, 0, None, None) , 0 , )),
	(( 'CursorMovement' , 'lcid' , 'RHS' , ), 232, (232, (), [ (3, 5, None, None) , 
			 (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 1180 , (3, 0, None, None) , 0 , )),
	(( 'CursorMovement' , 'lcid' , 'RHS' , ), 232, (232, (), [ (3, 5, None, None) , 
			 (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 1184 , (3, 0, None, None) , 0 , )),
	(( 'ControlCharacters' , 'lcid' , 'RHS' , ), 233, (233, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1188 , (3, 0, None, None) , 0 , )),
	(( 'ControlCharacters' , 'lcid' , 'RHS' , ), 233, (233, (), [ (3, 5, None, None) , 
			 (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1192 , (3, 0, None, None) , 0 , )),
	(( '_WSFunction' , 'Arg1' , 'Arg2' , 'Arg3' , 'Arg4' , 
			 'Arg5' , 'Arg6' , 'Arg7' , 'Arg8' , 'Arg9' , 
			 'Arg10' , 'Arg11' , 'Arg12' , 'Arg13' , 'Arg14' , 
			 'Arg15' , 'Arg16' , 'Arg17' , 'Arg18' , 'Arg19' , 
			 'Arg20' , 'Arg21' , 'Arg22' , 'Arg23' , 'Arg24' , 
			 'Arg25' , 'Arg26' , 'Arg27' , 'Arg28' , 'Arg29' , 
			 'Arg30' , 'lcid' , 'RHS' , ), 169, (169, (), [ (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (3, 5, None, None) , 
			 (16396, 10, None, None) , ], 1 , 1 , 4 , 30 , 1196 , (3, 0, None, None) , 1088 , )),
	(( 'EnableEvents' , 'RHS' , ), 1212, (1212, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1200 , (3, 0, None, None) , 0 , )),
	(( 'EnableEvents' , 'RHS' , ), 1212, (1212, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1204 , (3, 0, None, None) , 0 , )),
	(( 'DisplayInfoWindow' , 'RHS' , ), 1213, (1213, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1208 , (3, 0, None, None) , 64 , )),
	(( 'DisplayInfoWindow' , 'RHS' , ), 1213, (1213, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1212 , (3, 0, None, None) , 64 , )),
	(( 'Wait' , 'Time' , 'lcid' , 'RHS' , ), 1770, (1770, (), [ 
			 (12, 1, None, None) , (3, 5, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 1216 , (3, 0, None, None) , 0 , )),
	(( 'ExtendList' , 'RHS' , ), 1793, (1793, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1220 , (3, 0, None, None) , 0 , )),
	(( 'ExtendList' , 'RHS' , ), 1793, (1793, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1224 , (3, 0, None, None) , 0 , )),
	(( 'OLEDBErrors' , 'RHS' , ), 1794, (1794, (), [ (16393, 10, None, "IID('{00024446-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 1228 , (3, 0, None, None) , 0 , )),
	(( 'GetPhonetic' , 'Text' , 'RHS' , ), 1795, (1795, (), [ (12, 17, None, None) , 
			 (16392, 10, None, None) , ], 1 , 1 , 4 , 1 , 1232 , (3, 0, None, None) , 0 , )),
	(( 'COMAddIns' , 'RHS' , ), 1796, (1796, (), [ (16393, 10, None, "IID('{000C0339-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 1236 , (3, 0, None, None) , 0 , )),
	(( 'DefaultWebOptions' , 'RHS' , ), 1797, (1797, (), [ (16393, 10, None, "IID('{00024448-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 1240 , (3, 0, None, None) , 0 , )),
	(( 'ProductCode' , 'RHS' , ), 1798, (1798, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 1244 , (3, 0, None, None) , 0 , )),
	(( 'UserLibraryPath' , 'RHS' , ), 1799, (1799, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 1248 , (3, 0, None, None) , 0 , )),
	(( 'AutoPercentEntry' , 'RHS' , ), 1800, (1800, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1252 , (3, 0, None, None) , 0 , )),
	(( 'AutoPercentEntry' , 'RHS' , ), 1800, (1800, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1256 , (3, 0, None, None) , 0 , )),
	(( 'LanguageSettings' , 'RHS' , ), 1801, (1801, (), [ (16393, 10, None, "IID('{000C0353-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 1260 , (3, 0, None, None) , 0 , )),
	(( 'Dummy101' , 'RHS' , ), 1802, (1802, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 1264 , (3, 0, None, None) , 64 , )),
	(( 'Dummy12' , 'p1' , 'p2' , ), 1803, (1803, (), [ (9, 1, None, "IID('{00020872-0000-0000-C000-000000000046}')") , 
			 (9, 1, None, "IID('{00020872-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 0 , 1268 , (3, 0, None, None) , 64 , )),
	(( 'AnswerWizard' , 'RHS' , ), 1804, (1804, (), [ (16393, 10, None, "IID('{000C0360-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 1272 , (3, 0, None, None) , 64 , )),
	(( 'CalculateFull' , ), 1805, (1805, (), [ ], 1 , 1 , 4 , 0 , 1276 , (3, 0, None, None) , 0 , )),
	(( 'FindFile' , 'lcid' , 'RHS' , ), 1771, (1771, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 1280 , (3, 0, None, None) , 0 , )),
	(( 'CalculationVersion' , 'RHS' , ), 1806, (1806, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 1284 , (3, 0, None, None) , 0 , )),
	(( 'ShowWindowsInTaskbar' , 'RHS' , ), 1807, (1807, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1288 , (3, 0, None, None) , 0 , )),
	(( 'ShowWindowsInTaskbar' , 'RHS' , ), 1807, (1807, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1292 , (3, 0, None, None) , 0 , )),
	(( 'FeatureInstall' , 'RHS' , ), 1808, (1808, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 1296 , (3, 0, None, None) , 0 , )),
	(( 'FeatureInstall' , 'RHS' , ), 1808, (1808, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 1300 , (3, 0, None, None) , 0 , )),
	(( 'Ready' , 'RHS' , ), 1932, (1932, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1304 , (3, 0, None, None) , 0 , )),
	(( 'Dummy13' , 'Arg1' , 'Arg2' , 'Arg3' , 'Arg4' , 
			 'Arg5' , 'Arg6' , 'Arg7' , 'Arg8' , 'Arg9' , 
			 'Arg10' , 'Arg11' , 'Arg12' , 'Arg13' , 'Arg14' , 
			 'Arg15' , 'Arg16' , 'Arg17' , 'Arg18' , 'Arg19' , 
			 'Arg20' , 'Arg21' , 'Arg22' , 'Arg23' , 'Arg24' , 
			 'Arg25' , 'Arg26' , 'Arg27' , 'Arg28' , 'Arg29' , 
			 'Arg30' , 'RHS' , ), 1933, (1933, (), [ (12, 1, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (16396, 10, None, None) , ], 1 , 1 , 4 , 29 , 1308 , (3, 0, None, None) , 64 , )),
	(( 'FindFormat' , 'RHS' , ), 1934, (1934, (), [ (16393, 10, None, "IID('{00024450-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 1312 , (3, 0, None, None) , 0 , )),
	(( 'FindFormat' , 'RHS' , ), 1934, (1934, (), [ (9, 1, None, "IID('{00024450-0000-0000-C000-000000000046}')") , ], 1 , 8 , 4 , 0 , 1316 , (3, 0, None, None) , 0 , )),
	(( 'ReplaceFormat' , 'RHS' , ), 1935, (1935, (), [ (16393, 10, None, "IID('{00024450-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 1320 , (3, 0, None, None) , 0 , )),
	(( 'ReplaceFormat' , 'RHS' , ), 1935, (1935, (), [ (9, 1, None, "IID('{00024450-0000-0000-C000-000000000046}')") , ], 1 , 8 , 4 , 0 , 1324 , (3, 0, None, None) , 0 , )),
	(( 'UsedObjects' , 'RHS' , ), 1936, (1936, (), [ (16393, 10, None, "IID('{00024451-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 1328 , (3, 0, None, None) , 0 , )),
	(( 'CalculationState' , 'RHS' , ), 1937, (1937, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 1332 , (3, 0, None, None) , 0 , )),
	(( 'CalculationInterruptKey' , 'RHS' , ), 1938, (1938, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 1336 , (3, 0, None, None) , 0 , )),
	(( 'CalculationInterruptKey' , 'RHS' , ), 1938, (1938, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 1340 , (3, 0, None, None) , 0 , )),
	(( 'Watches' , 'RHS' , ), 1939, (1939, (), [ (16393, 10, None, "IID('{00024456-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 1344 , (3, 0, None, None) , 0 , )),
	(( 'DisplayFunctionToolTips' , 'RHS' , ), 1940, (1940, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1348 , (3, 0, None, None) , 0 , )),
	(( 'DisplayFunctionToolTips' , 'RHS' , ), 1940, (1940, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1352 , (3, 0, None, None) , 0 , )),
	(( 'AutomationSecurity' , 'RHS' , ), 1941, (1941, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 1356 , (3, 0, None, None) , 0 , )),
	(( 'AutomationSecurity' , 'RHS' , ), 1941, (1941, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 1360 , (3, 0, None, None) , 0 , )),
	(( 'FileDialog' , 'fileDialogType' , 'RHS' , ), 1942, (1942, (), [ (3, 1, None, None) , 
			 (16393, 10, None, "IID('{000C0362-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 1364 , (3, 0, None, None) , 0 , )),
	(( 'Dummy14' , ), 1944, (1944, (), [ ], 1 , 1 , 4 , 0 , 1368 , (3, 0, None, None) , 64 , )),
	(( 'CalculateFullRebuild' , ), 1945, (1945, (), [ ], 1 , 1 , 4 , 0 , 1372 , (3, 0, None, None) , 0 , )),
	(( 'DisplayPasteOptions' , 'RHS' , ), 1946, (1946, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1376 , (3, 0, None, None) , 0 , )),
	(( 'DisplayPasteOptions' , 'RHS' , ), 1946, (1946, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1380 , (3, 0, None, None) , 0 , )),
	(( 'DisplayInsertOptions' , 'RHS' , ), 1947, (1947, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1384 , (3, 0, None, None) , 0 , )),
	(( 'DisplayInsertOptions' , 'RHS' , ), 1947, (1947, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1388 , (3, 0, None, None) , 0 , )),
	(( 'GenerateGetPivotData' , 'RHS' , ), 1948, (1948, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1392 , (3, 0, None, None) , 0 , )),
	(( 'GenerateGetPivotData' , 'RHS' , ), 1948, (1948, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1396 , (3, 0, None, None) , 0 , )),
	(( 'AutoRecover' , 'RHS' , ), 1949, (1949, (), [ (16393, 10, None, "IID('{0002445A-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 1400 , (3, 0, None, None) , 0 , )),
	(( 'Hwnd' , 'RHS' , ), 1950, (1950, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 1404 , (3, 0, None, None) , 0 , )),
	(( 'Hinstance' , 'RHS' , ), 1951, (1951, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 1408 , (3, 0, None, None) , 0 , )),
	(( 'CheckAbort' , 'KeepAbort' , ), 1952, (1952, (), [ (12, 17, None, None) , ], 1 , 1 , 4 , 1 , 1412 , (3, 0, None, None) , 0 , )),
	(( 'ErrorCheckingOptions' , 'RHS' , ), 1954, (1954, (), [ (16393, 10, None, "IID('{0002445B-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 1416 , (3, 0, None, None) , 0 , )),
	(( 'AutoFormatAsYouTypeReplaceHyperlinks' , 'RHS' , ), 1955, (1955, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1420 , (3, 0, None, None) , 0 , )),
	(( 'AutoFormatAsYouTypeReplaceHyperlinks' , 'RHS' , ), 1955, (1955, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1424 , (3, 0, None, None) , 0 , )),
	(( 'SmartTagRecognizers' , 'RHS' , ), 1956, (1956, (), [ (16393, 10, None, "IID('{00024463-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 1428 , (3, 0, None, None) , 64 , )),
	(( 'NewWorkbook' , 'RHS' , ), 1565, (1565, (), [ (16393, 10, None, "IID('{000C0936-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 1432 , (3, 0, None, None) , 0 , )),
	(( 'SpellingOptions' , 'RHS' , ), 1957, (1957, (), [ (16393, 10, None, "IID('{00024465-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 1436 , (3, 0, None, None) , 0 , )),
	(( 'Speech' , 'RHS' , ), 1958, (1958, (), [ (16393, 10, None, "IID('{00024466-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 1440 , (3, 0, None, None) , 0 , )),
	(( 'MapPaperSize' , 'RHS' , ), 1959, (1959, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1444 , (3, 0, None, None) , 0 , )),
	(( 'MapPaperSize' , 'RHS' , ), 1959, (1959, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1448 , (3, 0, None, None) , 0 , )),
	(( 'ShowStartupDialog' , 'RHS' , ), 1960, (1960, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1452 , (3, 0, None, None) , 0 , )),
	(( 'ShowStartupDialog' , 'RHS' , ), 1960, (1960, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1456 , (3, 0, None, None) , 0 , )),
	(( 'DecimalSeparator' , 'RHS' , ), 1809, (1809, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 1460 , (3, 0, None, None) , 0 , )),
	(( 'DecimalSeparator' , 'RHS' , ), 1809, (1809, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 1464 , (3, 0, None, None) , 0 , )),
	(( 'ThousandsSeparator' , 'RHS' , ), 1810, (1810, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 1468 , (3, 0, None, None) , 0 , )),
	(( 'ThousandsSeparator' , 'RHS' , ), 1810, (1810, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 1472 , (3, 0, None, None) , 0 , )),
	(( 'UseSystemSeparators' , 'RHS' , ), 1961, (1961, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1476 , (3, 0, None, None) , 0 , )),
	(( 'UseSystemSeparators' , 'RHS' , ), 1961, (1961, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1480 , (3, 0, None, None) , 0 , )),
	(( 'ThisCell' , 'RHS' , ), 1962, (1962, (), [ (16393, 10, None, "IID('{00020846-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 1484 , (3, 0, None, None) , 0 , )),
	(( 'RTD' , 'RHS' , ), 1963, (1963, (), [ (16393, 10, None, "IID('{0002446E-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 1488 , (3, 0, None, None) , 0 , )),
	(( 'DisplayDocumentActionTaskPane' , 'RHS' , ), 2251, (2251, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1492 , (3, 0, None, None) , 0 , )),
	(( 'DisplayDocumentActionTaskPane' , 'RHS' , ), 2251, (2251, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1496 , (3, 0, None, None) , 0 , )),
	(( 'DisplayXMLSourcePane' , 'XmlMap' , ), 2252, (2252, (), [ (12, 17, None, None) , ], 1 , 1 , 4 , 1 , 1500 , (3, 0, None, None) , 0 , )),
	(( 'ArbitraryXMLSupportAvailable' , 'RHS' , ), 2254, (2254, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1504 , (3, 0, None, None) , 0 , )),
	(( 'Support' , 'Object' , 'ID' , 'arg' , 'RHS' , 
			 ), 2255, (2255, (), [ (9, 1, None, None) , (3, 1, None, None) , (12, 17, None, None) , (16396, 10, None, None) , ], 1 , 1 , 4 , 1 , 1508 , (3, 0, None, None) , 64 , )),
	(( 'Dummy20' , 'grfCompareFunctions' , 'RHS' , ), 2373, (2373, (), [ (3, 1, None, None) , 
			 (16396, 10, None, None) , ], 1 , 1 , 4 , 0 , 1512 , (3, 0, None, None) , 64 , )),
	(( 'MeasurementUnit' , 'RHS' , ), 2375, (2375, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 1516 , (3, 0, None, None) , 0 , )),
	(( 'MeasurementUnit' , 'RHS' , ), 2375, (2375, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 1520 , (3, 0, None, None) , 0 , )),
	(( 'ShowSelectionFloaties' , 'RHS' , ), 2376, (2376, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1524 , (3, 0, None, None) , 0 , )),
	(( 'ShowSelectionFloaties' , 'RHS' , ), 2376, (2376, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1528 , (3, 0, None, None) , 0 , )),
	(( 'ShowMenuFloaties' , 'RHS' , ), 2377, (2377, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1532 , (3, 0, None, None) , 0 , )),
	(( 'ShowMenuFloaties' , 'RHS' , ), 2377, (2377, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1536 , (3, 0, None, None) , 0 , )),
	(( 'ShowDevTools' , 'RHS' , ), 2378, (2378, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1540 , (3, 0, None, None) , 0 , )),
	(( 'ShowDevTools' , 'RHS' , ), 2378, (2378, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1544 , (3, 0, None, None) , 0 , )),
	(( 'EnableLivePreview' , 'RHS' , ), 2379, (2379, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1548 , (3, 0, None, None) , 0 , )),
	(( 'EnableLivePreview' , 'RHS' , ), 2379, (2379, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1552 , (3, 0, None, None) , 0 , )),
	(( 'DisplayDocumentInformationPanel' , 'RHS' , ), 2380, (2380, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1556 , (3, 0, None, None) , 0 , )),
	(( 'DisplayDocumentInformationPanel' , 'RHS' , ), 2380, (2380, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1560 , (3, 0, None, None) , 0 , )),
	(( 'AlwaysUseClearType' , 'RHS' , ), 2381, (2381, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1564 , (3, 0, None, None) , 0 , )),
	(( 'AlwaysUseClearType' , 'RHS' , ), 2381, (2381, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1568 , (3, 0, None, None) , 0 , )),
	(( 'WarnOnFunctionNameConflict' , 'RHS' , ), 2382, (2382, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1572 , (3, 0, None, None) , 0 , )),
	(( 'WarnOnFunctionNameConflict' , 'RHS' , ), 2382, (2382, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1576 , (3, 0, None, None) , 0 , )),
	(( 'FormulaBarHeight' , 'RHS' , ), 2383, (2383, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 1580 , (3, 0, None, None) , 0 , )),
	(( 'FormulaBarHeight' , 'RHS' , ), 2383, (2383, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 1584 , (3, 0, None, None) , 0 , )),
	(( 'DisplayFormulaAutoComplete' , 'RHS' , ), 2384, (2384, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1588 , (3, 0, None, None) , 0 , )),
	(( 'DisplayFormulaAutoComplete' , 'RHS' , ), 2384, (2384, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1592 , (3, 0, None, None) , 0 , )),
	(( 'GenerateTableRefs' , 'lcid' , 'RHS' , ), 2385, (2385, (), [ (3, 5, None, None) , 
			 (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 1596 , (3, 0, None, None) , 0 , )),
	(( 'GenerateTableRefs' , 'lcid' , 'RHS' , ), 2385, (2385, (), [ (3, 5, None, None) , 
			 (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 1600 , (3, 0, None, None) , 0 , )),
	(( 'Assistance' , 'RHS' , ), 2386, (2386, (), [ (16393, 10, None, "IID('{4291224C-DEFE-485B-8E69-6CF8AA85CB76}')") , ], 1 , 2 , 4 , 0 , 1604 , (3, 0, None, None) , 0 , )),
	(( 'CalculateUntilAsyncQueriesDone' , ), 2387, (2387, (), [ ], 1 , 1 , 4 , 0 , 1608 , (3, 0, None, None) , 0 , )),
	(( 'EnableLargeOperationAlert' , 'RHS' , ), 2388, (2388, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1612 , (3, 0, None, None) , 0 , )),
	(( 'EnableLargeOperationAlert' , 'RHS' , ), 2388, (2388, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1616 , (3, 0, None, None) , 0 , )),
	(( 'LargeOperationCellThousandCount' , 'RHS' , ), 2389, (2389, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 1620 , (3, 0, None, None) , 0 , )),
	(( 'LargeOperationCellThousandCount' , 'RHS' , ), 2389, (2389, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 1624 , (3, 0, None, None) , 0 , )),
	(( 'DeferAsyncQueries' , 'RHS' , ), 2390, (2390, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1628 , (3, 0, None, None) , 0 , )),
	(( 'DeferAsyncQueries' , 'RHS' , ), 2390, (2390, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1632 , (3, 0, None, None) , 0 , )),
	(( 'MultiThreadedCalculation' , 'RHS' , ), 2391, (2391, (), [ (16393, 10, None, "IID('{000244B1-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 1636 , (3, 0, None, None) , 0 , )),
	(( 'SharePointVersion' , 'bstrUrl' , 'RHS' , ), 2392, (2392, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 1640 , (3, 0, None, None) , 0 , )),
	(( 'ActiveEncryptionSession' , 'RHS' , ), 2394, (2394, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 1644 , (3, 0, None, None) , 0 , )),
	(( 'HighQualityModeForGraphics' , 'RHS' , ), 2395, (2395, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1648 , (3, 0, None, None) , 0 , )),
	(( 'HighQualityModeForGraphics' , 'RHS' , ), 2395, (2395, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1652 , (3, 0, None, None) , 0 , )),
	(( 'FileExportConverters' , 'RHS' , ), 2768, (2768, (), [ (16393, 10, None, "IID('{000244B4-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 1656 , (3, 0, None, None) , 0 , )),
	(( 'SmartArtLayouts' , 'RHS' , ), 2772, (2772, (), [ (16393, 10, None, "IID('{000C03C9-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 1660 , (3, 0, None, None) , 0 , )),
	(( 'SmartArtQuickStyles' , 'RHS' , ), 2773, (2773, (), [ (16393, 10, None, "IID('{000C03CB-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 1664 , (3, 0, None, None) , 0 , )),
	(( 'SmartArtColors' , 'RHS' , ), 2774, (2774, (), [ (16393, 10, None, "IID('{000C03CD-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 1668 , (3, 0, None, None) , 0 , )),
	(( 'AddIns2' , 'RHS' , ), 2775, (2775, (), [ (16393, 10, None, "IID('{000244B5-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 1672 , (3, 0, None, None) , 0 , )),
	(( 'PrintCommunication' , 'RHS' , ), 2776, (2776, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1676 , (3, 0, None, None) , 0 , )),
	(( 'PrintCommunication' , 'RHS' , ), 2776, (2776, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1680 , (3, 0, None, None) , 0 , )),
	(( 'MacroOptions' , 'Macro' , 'Description' , 'HasMenu' , 'MenuText' , 
			 'HasShortcutKey' , 'ShortcutKey' , 'Category' , 'StatusBar' , 'HelpContextID' , 
			 'HelpFile' , 'ArgumentDescriptions' , 'lcid' , ), 2770, (2770, (), [ (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (3, 5, None, None) , ], 1 , 1 , 4 , 11 , 1684 , (3, 0, None, None) , 0 , )),
	(( 'UseClusterConnector' , 'RHS' , ), 2778, (2778, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1688 , (3, 0, None, None) , 0 , )),
	(( 'UseClusterConnector' , 'RHS' , ), 2778, (2778, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1692 , (3, 0, None, None) , 0 , )),
	(( 'ClusterConnector' , 'RHS' , ), 2779, (2779, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 1696 , (3, 0, None, None) , 0 , )),
	(( 'ClusterConnector' , 'RHS' , ), 2779, (2779, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 1700 , (3, 0, None, None) , 0 , )),
	(( 'Quitting' , 'RHS' , ), 2780, (2780, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1704 , (3, 0, None, None) , 64 , )),
	(( 'Dummy22' , 'RHS' , ), 2781, (2781, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1708 , (3, 0, None, None) , 64 , )),
	(( 'Dummy22' , 'RHS' , ), 2781, (2781, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1712 , (3, 0, None, None) , 64 , )),
	(( 'Dummy23' , 'RHS' , ), 2782, (2782, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1716 , (3, 0, None, None) , 64 , )),
	(( 'Dummy23' , 'RHS' , ), 2782, (2782, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1720 , (3, 0, None, None) , 64 , )),
	(( 'ProtectedViewWindows' , 'RHS' , ), 2783, (2783, (), [ (16393, 10, None, "IID('{000244CC-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 1724 , (3, 0, None, None) , 0 , )),
	(( 'ActiveProtectedViewWindow' , 'RHS' , ), 2784, (2784, (), [ (16393, 10, None, "IID('{000244CD-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 1728 , (3, 0, None, None) , 0 , )),
	(( 'IsSandboxed' , 'RHS' , ), 2785, (2785, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1732 , (3, 0, None, None) , 0 , )),
	(( 'SaveISO8601Dates' , 'RHS' , ), 2786, (2786, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1736 , (3, 0, None, None) , 0 , )),
	(( 'SaveISO8601Dates' , 'RHS' , ), 2786, (2786, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1740 , (3, 0, None, None) , 0 , )),
	(( 'HinstancePtr' , 'RHS' , ), 2787, (2787, (), [ (16396, 10, None, None) , ], 1 , 2 , 4 , 0 , 1744 , (3, 0, None, None) , 0 , )),
	(( 'FileValidation' , 'RHS' , ), 2788, (2788, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 1748 , (3, 0, None, None) , 0 , )),
	(( 'FileValidation' , 'RHS' , ), 2788, (2788, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 1752 , (3, 0, None, None) , 0 , )),
	(( 'FileValidationPivot' , 'RHS' , ), 2789, (2789, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 1756 , (3, 0, None, None) , 0 , )),
	(( 'FileValidationPivot' , 'RHS' , ), 2789, (2789, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 1760 , (3, 0, None, None) , 0 , )),
]

win32com.client.CLSIDToClass.RegisterCLSID( "{000208D5-0000-0000-C000-000000000046}", _Application )
