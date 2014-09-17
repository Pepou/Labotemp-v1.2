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

class AppEvents:
	CLSID = CLSID_Sink = IID('{00024413-0000-0000-C000-000000000046}')
	coclass_clsid = IID('{00024500-0000-0000-C000-000000000046}')
	_public_methods_ = [] # For COM Server support
	_dispid_to_func_ = {
		     1573 : "OnWorkbookNewSheet",
		     2906 : "OnProtectedViewWindowBeforeClose",
		     2897 : "OnSheetPivotTableBeforeCommitChanges",
		     1565 : "OnNewWorkbook",
		     1554 : "OnWindowResize",
		     2611 : "OnWorkbookRowsetComplete",
		     2291 : "OnWorkbookAfterXmlImport",
		     2910 : "OnProtectedViewWindowDeactivate",
		     1558 : "OnSheetSelectionChange",
		     1562 : "OnSheetDeactivate",
		     1575 : "OnWorkbookAddinUninstall",
		     1574 : "OnWorkbookAddinInstall",
		     2612 : "OnAfterCalculate",
		1610612737 : "OnAddRef",
		     2896 : "OnSheetPivotTableBeforeAllocateChanges",
		     2895 : "OnSheetPivotTableAfterValueChange",
		     2290 : "OnWorkbookBeforeXmlImport",
		     2289 : "OnWorkbookSync",
		     1560 : "OnSheetBeforeRightClick",
		     1559 : "OnSheetBeforeDoubleClick",
		     1569 : "OnWorkbookDeactivate",
		     2157 : "OnSheetPivotTableUpdate",
		     2905 : "OnProtectedViewWindowBeforeEdit",
		     1556 : "OnWindowActivate",
		     1570 : "OnWorkbookBeforeClose",
		1610678275 : "OnInvoke",
		     1854 : "OnSheetFollowHyperlink",
		     2908 : "OnProtectedViewWindowResize",
		     2292 : "OnWorkbookBeforeXmlExport",
		1610678272 : "OnGetTypeInfoCount",
		     1557 : "OnWindowDeactivate",
		1610678274 : "OnGetIDsOfNames",
		     1561 : "OnSheetActivate",
		1610612738 : "OnRelease",
		     2909 : "OnProtectedViewWindowActivate",
		1610612736 : "OnQueryInterface",
		     1567 : "OnWorkbookOpen",
		1610678273 : "OnGetTypeInfo",
		     2912 : "OnWorkbookNewChart",
		     1564 : "OnSheetChange",
		     2160 : "OnWorkbookPivotTableCloseConnection",
		     1563 : "OnSheetCalculate",
		     1568 : "OnWorkbookActivate",
		     2903 : "OnProtectedViewWindowOpen",
		     1571 : "OnWorkbookBeforeSave",
		     2161 : "OnWorkbookPivotTableOpenConnection",
		     2293 : "OnWorkbookAfterXmlExport",
		     1572 : "OnWorkbookBeforePrint",
		     2898 : "OnSheetPivotTableBeforeDiscardChanges",
		     2911 : "OnWorkbookAfterSave",
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
#	def OnWorkbookNewSheet(self, Wb=defaultNamedNotOptArg, Sh=defaultNamedNotOptArg):
#	def OnProtectedViewWindowBeforeClose(self, Pvw=defaultNamedNotOptArg, Reason=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnSheetPivotTableBeforeCommitChanges(self, Sh=defaultNamedNotOptArg, TargetPivotTable=defaultNamedNotOptArg, ValueChangeStart=defaultNamedNotOptArg, ValueChangeEnd=defaultNamedNotOptArg
#			, Cancel=defaultNamedNotOptArg):
#	def OnNewWorkbook(self, Wb=defaultNamedNotOptArg):
#	def OnWindowResize(self, Wb=defaultNamedNotOptArg, Wn=defaultNamedNotOptArg):
#	def OnWorkbookRowsetComplete(self, Wb=defaultNamedNotOptArg, Description=defaultNamedNotOptArg, Sheet=defaultNamedNotOptArg, Success=defaultNamedNotOptArg):
#	def OnWorkbookAfterXmlImport(self, Wb=defaultNamedNotOptArg, Map=defaultNamedNotOptArg, IsRefresh=defaultNamedNotOptArg, Result=defaultNamedNotOptArg):
#	def OnProtectedViewWindowDeactivate(self, Pvw=defaultNamedNotOptArg):
#	def OnSheetSelectionChange(self, Sh=defaultNamedNotOptArg, Target=defaultNamedNotOptArg):
#	def OnSheetDeactivate(self, Sh=defaultNamedNotOptArg):
#	def OnWorkbookAddinUninstall(self, Wb=defaultNamedNotOptArg):
#	def OnWorkbookAddinInstall(self, Wb=defaultNamedNotOptArg):
#	def OnAfterCalculate(self):
#	def OnAddRef(self):
#	def OnSheetPivotTableBeforeAllocateChanges(self, Sh=defaultNamedNotOptArg, TargetPivotTable=defaultNamedNotOptArg, ValueChangeStart=defaultNamedNotOptArg, ValueChangeEnd=defaultNamedNotOptArg
#			, Cancel=defaultNamedNotOptArg):
#	def OnSheetPivotTableAfterValueChange(self, Sh=defaultNamedNotOptArg, TargetPivotTable=defaultNamedNotOptArg, TargetRange=defaultNamedNotOptArg):
#	def OnWorkbookBeforeXmlImport(self, Wb=defaultNamedNotOptArg, Map=defaultNamedNotOptArg, Url=defaultNamedNotOptArg, IsRefresh=defaultNamedNotOptArg
#			, Cancel=defaultNamedNotOptArg):
#	def OnWorkbookSync(self, Wb=defaultNamedNotOptArg, SyncEventType=defaultNamedNotOptArg):
#	def OnSheetBeforeRightClick(self, Sh=defaultNamedNotOptArg, Target=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnSheetBeforeDoubleClick(self, Sh=defaultNamedNotOptArg, Target=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnWorkbookDeactivate(self, Wb=defaultNamedNotOptArg):
#	def OnSheetPivotTableUpdate(self, Sh=defaultNamedNotOptArg, Target=defaultNamedNotOptArg):
#	def OnProtectedViewWindowBeforeEdit(self, Pvw=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnWindowActivate(self, Wb=defaultNamedNotOptArg, Wn=defaultNamedNotOptArg):
#	def OnWorkbookBeforeClose(self, Wb=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnInvoke(self, dispidMember=defaultNamedNotOptArg, riid=defaultNamedNotOptArg, lcid=defaultNamedNotOptArg, wFlags=defaultNamedNotOptArg
#			, pdispparams=defaultNamedNotOptArg, pvarResult=pythoncom.Missing, pexcepinfo=pythoncom.Missing, puArgErr=pythoncom.Missing):
#	def OnSheetFollowHyperlink(self, Sh=defaultNamedNotOptArg, Target=defaultNamedNotOptArg):
#	def OnProtectedViewWindowResize(self, Pvw=defaultNamedNotOptArg):
#	def OnWorkbookBeforeXmlExport(self, Wb=defaultNamedNotOptArg, Map=defaultNamedNotOptArg, Url=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnGetTypeInfoCount(self, pctinfo=pythoncom.Missing):
#	def OnWindowDeactivate(self, Wb=defaultNamedNotOptArg, Wn=defaultNamedNotOptArg):
#	def OnGetIDsOfNames(self, riid=defaultNamedNotOptArg, rgszNames=defaultNamedNotOptArg, cNames=defaultNamedNotOptArg, lcid=defaultNamedNotOptArg
#			, rgdispid=pythoncom.Missing):
#	def OnSheetActivate(self, Sh=defaultNamedNotOptArg):
#	def OnRelease(self):
#	def OnProtectedViewWindowActivate(self, Pvw=defaultNamedNotOptArg):
#	def OnQueryInterface(self, riid=defaultNamedNotOptArg, ppvObj=pythoncom.Missing):
#	def OnWorkbookOpen(self, Wb=defaultNamedNotOptArg):
#	def OnGetTypeInfo(self, itinfo=defaultNamedNotOptArg, lcid=defaultNamedNotOptArg, pptinfo=pythoncom.Missing):
#	def OnWorkbookNewChart(self, Wb=defaultNamedNotOptArg, Ch=defaultNamedNotOptArg):
#	def OnSheetChange(self, Sh=defaultNamedNotOptArg, Target=defaultNamedNotOptArg):
#	def OnWorkbookPivotTableCloseConnection(self, Wb=defaultNamedNotOptArg, Target=defaultNamedNotOptArg):
#	def OnSheetCalculate(self, Sh=defaultNamedNotOptArg):
#	def OnWorkbookActivate(self, Wb=defaultNamedNotOptArg):
#	def OnProtectedViewWindowOpen(self, Pvw=defaultNamedNotOptArg):
#	def OnWorkbookBeforeSave(self, Wb=defaultNamedNotOptArg, SaveAsUI=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnWorkbookPivotTableOpenConnection(self, Wb=defaultNamedNotOptArg, Target=defaultNamedNotOptArg):
#	def OnWorkbookAfterXmlExport(self, Wb=defaultNamedNotOptArg, Map=defaultNamedNotOptArg, Url=defaultNamedNotOptArg, Result=defaultNamedNotOptArg):
#	def OnWorkbookBeforePrint(self, Wb=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnSheetPivotTableBeforeDiscardChanges(self, Sh=defaultNamedNotOptArg, TargetPivotTable=defaultNamedNotOptArg, ValueChangeStart=defaultNamedNotOptArg, ValueChangeEnd=defaultNamedNotOptArg):
#	def OnWorkbookAfterSave(self, Wb=defaultNamedNotOptArg, Success=defaultNamedNotOptArg):


win32com.client.CLSIDToClass.RegisterCLSID( "{00024413-0000-0000-C000-000000000046}", AppEvents )
