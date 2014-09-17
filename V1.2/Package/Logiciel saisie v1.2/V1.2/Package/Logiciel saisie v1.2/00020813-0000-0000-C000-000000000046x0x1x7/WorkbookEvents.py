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

class WorkbookEvents:
	CLSID = CLSID_Sink = IID('{00024412-0000-0000-C000-000000000046}')
	coclass_clsid = IID('{00020819-0000-0000-C000-000000000046}')
	_public_methods_ = [] # For COM Server support
	_dispid_to_func_ = {
		     1563 : "OnSheetCalculate",
		     2158 : "OnPivotTableCloseConnection",
		1610678273 : "OnGetTypeInfo",
		1610612738 : "OnRelease",
		     1553 : "OnAddinUninstall",
		     1549 : "OnBeforePrint",
		1610678272 : "OnGetTypeInfoCount",
		     2900 : "OnAfterSave",
		     2901 : "OnNewChart",
		     1552 : "OnAddinInstall",
		      304 : "OnActivate",
		     2895 : "OnSheetPivotTableAfterValueChange",
		     1562 : "OnSheetDeactivate",
		     2288 : "OnAfterXmlExport",
		     1554 : "OnWindowResize",
		1610612736 : "OnQueryInterface",
		1610678275 : "OnInvoke",
		     2159 : "OnPivotTableOpenConnection",
		     1530 : "OnDeactivate",
		     2285 : "OnAfterXmlImport",
		     2287 : "OnBeforeXmlExport",
		1610678274 : "OnGetIDsOfNames",
		     2897 : "OnSheetPivotTableBeforeCommitChanges",
		     1923 : "OnOpen",
		     2283 : "OnBeforeXmlImport",
		     1854 : "OnSheetFollowHyperlink",
		     2899 : "OnSheetPivotTableChangeSync",
		1610612737 : "OnAddRef",
		     1546 : "OnBeforeClose",
		     1560 : "OnSheetBeforeRightClick",
		     2898 : "OnSheetPivotTableBeforeDiscardChanges",
		     1559 : "OnSheetBeforeDoubleClick",
		     1550 : "OnNewSheet",
		     1557 : "OnWindowDeactivate",
		     1558 : "OnSheetSelectionChange",
		     1547 : "OnBeforeSave",
		     2610 : "OnRowsetComplete",
		     2896 : "OnSheetPivotTableBeforeAllocateChanges",
		     2266 : "OnSync",
		     2157 : "OnSheetPivotTableUpdate",
		     1564 : "OnSheetChange",
		     1561 : "OnSheetActivate",
		     1556 : "OnWindowActivate",
		}

	def __init__(self, oobj = None):
		if oobj is None:
			self._olecp = None
		else:
			import win32com.server.util
			from win32com.server.policy import EventHandlerPolicy
			cpc=oobj._oleobj_.QueryInterface(pythoncom.IID_IConnectionPointContainer)
			cp=cpc.FindConnectionPoint(self.CLSID_Sink)
			cookie=cp.Advise(win32com.server.util.wrap(self, usePolicy=EventHandlerPolicy))
			self._olecp,self._olecp_cookie = cp,cookie
	def __del__(self):
		try:
			self.close()
		except pythoncom.com_error:
			pass
	def close(self):
		if self._olecp is not None:
			cp,cookie,self._olecp,self._olecp_cookie = self._olecp,self._olecp_cookie,None,None
			cp.Unadvise(cookie)
	def _query_interface_(self, iid):
		import win32com.server.util
		if iid==self.CLSID_Sink: return win32com.server.util.wrap(self)

	# Event Handlers
	# If you create handlers, they should have the following prototypes:
#	def OnSheetCalculate(self, Sh=defaultNamedNotOptArg):
#	def OnPivotTableCloseConnection(self, Target=defaultNamedNotOptArg):
#	def OnGetTypeInfo(self, itinfo=defaultNamedNotOptArg, lcid=defaultNamedNotOptArg, pptinfo=pythoncom.Missing):
#	def OnRelease(self):
#	def OnAddinUninstall(self):
#	def OnBeforePrint(self, Cancel=defaultNamedNotOptArg):
#	def OnGetTypeInfoCount(self, pctinfo=pythoncom.Missing):
#	def OnAfterSave(self, Success=defaultNamedNotOptArg):
#	def OnNewChart(self, Ch=defaultNamedNotOptArg):
#	def OnAddinInstall(self):
#	def OnActivate(self):
#	def OnSheetPivotTableAfterValueChange(self, Sh=defaultNamedNotOptArg, TargetPivotTable=defaultNamedNotOptArg, TargetRange=defaultNamedNotOptArg):
#	def OnSheetDeactivate(self, Sh=defaultNamedNotOptArg):
#	def OnAfterXmlExport(self, Map=defaultNamedNotOptArg, Url=defaultNamedNotOptArg, Result=defaultNamedNotOptArg):
#	def OnWindowResize(self, Wn=defaultNamedNotOptArg):
#	def OnQueryInterface(self, riid=defaultNamedNotOptArg, ppvObj=pythoncom.Missing):
#	def OnInvoke(self, dispidMember=defaultNamedNotOptArg, riid=defaultNamedNotOptArg, lcid=defaultNamedNotOptArg, wFlags=defaultNamedNotOptArg
#			, pdispparams=defaultNamedNotOptArg, pvarResult=pythoncom.Missing, pexcepinfo=pythoncom.Missing, puArgErr=pythoncom.Missing):
#	def OnPivotTableOpenConnection(self, Target=defaultNamedNotOptArg):
#	def OnDeactivate(self):
#	def OnAfterXmlImport(self, Map=defaultNamedNotOptArg, IsRefresh=defaultNamedNotOptArg, Result=defaultNamedNotOptArg):
#	def OnBeforeXmlExport(self, Map=defaultNamedNotOptArg, Url=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnGetIDsOfNames(self, riid=defaultNamedNotOptArg, rgszNames=defaultNamedNotOptArg, cNames=defaultNamedNotOptArg, lcid=defaultNamedNotOptArg
#			, rgdispid=pythoncom.Missing):
#	def OnSheetPivotTableBeforeCommitChanges(self, Sh=defaultNamedNotOptArg, TargetPivotTable=defaultNamedNotOptArg, ValueChangeStart=defaultNamedNotOptArg, ValueChangeEnd=defaultNamedNotOptArg
#			, Cancel=defaultNamedNotOptArg):
#	def OnOpen(self):
#	def OnBeforeXmlImport(self, Map=defaultNamedNotOptArg, Url=defaultNamedNotOptArg, IsRefresh=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnSheetFollowHyperlink(self, Sh=defaultNamedNotOptArg, Target=defaultNamedNotOptArg):
#	def OnSheetPivotTableChangeSync(self, Sh=defaultNamedNotOptArg, Target=defaultNamedNotOptArg):
#	def OnAddRef(self):
#	def OnBeforeClose(self, Cancel=defaultNamedNotOptArg):
#	def OnSheetBeforeRightClick(self, Sh=defaultNamedNotOptArg, Target=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnSheetPivotTableBeforeDiscardChanges(self, Sh=defaultNamedNotOptArg, TargetPivotTable=defaultNamedNotOptArg, ValueChangeStart=defaultNamedNotOptArg, ValueChangeEnd=defaultNamedNotOptArg):
#	def OnSheetBeforeDoubleClick(self, Sh=defaultNamedNotOptArg, Target=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnNewSheet(self, Sh=defaultNamedNotOptArg):
#	def OnWindowDeactivate(self, Wn=defaultNamedNotOptArg):
#	def OnSheetSelectionChange(self, Sh=defaultNamedNotOptArg, Target=defaultNamedNotOptArg):
#	def OnBeforeSave(self, SaveAsUI=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnRowsetComplete(self, Description=defaultNamedNotOptArg, Sheet=defaultNamedNotOptArg, Success=defaultNamedNotOptArg):
#	def OnSheetPivotTableBeforeAllocateChanges(self, Sh=defaultNamedNotOptArg, TargetPivotTable=defaultNamedNotOptArg, ValueChangeStart=defaultNamedNotOptArg, ValueChangeEnd=defaultNamedNotOptArg
#			, Cancel=defaultNamedNotOptArg):
#	def OnSync(self, SyncEventType=defaultNamedNotOptArg):
#	def OnSheetPivotTableUpdate(self, Sh=defaultNamedNotOptArg, Target=defaultNamedNotOptArg):
#	def OnSheetChange(self, Sh=defaultNamedNotOptArg, Target=defaultNamedNotOptArg):
#	def OnSheetActivate(self, Sh=defaultNamedNotOptArg):
#	def OnWindowActivate(self, Wn=defaultNamedNotOptArg):


win32com.client.CLSIDToClass.RegisterCLSID( "{00024412-0000-0000-C000-000000000046}", WorkbookEvents )
