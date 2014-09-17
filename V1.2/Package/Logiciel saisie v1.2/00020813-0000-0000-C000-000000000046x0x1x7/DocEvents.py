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

class DocEvents:
	CLSID = CLSID_Sink = IID('{00024411-0000-0000-C000-000000000046}')
	coclass_clsid = IID('{00020820-0000-0000-C000-000000000046}')
	_public_methods_ = [] # For COM Server support
	_dispid_to_func_ = {
		     1543 : "OnSelectionChange",
		1610678273 : "OnGetTypeInfo",
		     2893 : "OnPivotTableBeforeDiscardChanges",
		      279 : "OnCalculate",
		1610612738 : "OnRelease",
		1610612737 : "OnAddRef",
		1610678272 : "OnGetTypeInfoCount",
		      304 : "OnActivate",
		     2886 : "OnPivotTableAfterValueChange",
		     1530 : "OnDeactivate",
		1610612736 : "OnQueryInterface",
		1610678275 : "OnInvoke",
		     2894 : "OnPivotTableChangeSync",
		     2889 : "OnPivotTableBeforeAllocateChanges",
		     1534 : "OnBeforeRightClick",
		     1545 : "OnChange",
		     1470 : "OnFollowHyperlink",
		     1537 : "OnBeforeDoubleClick",
		     2892 : "OnPivotTableBeforeCommitChanges",
		     2156 : "OnPivotTableUpdate",
		1610678274 : "OnGetIDsOfNames",
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
#	def OnSelectionChange(self, Target=defaultNamedNotOptArg):
#	def OnGetTypeInfo(self, itinfo=defaultNamedNotOptArg, lcid=defaultNamedNotOptArg, pptinfo=pythoncom.Missing):
#	def OnPivotTableBeforeDiscardChanges(self, TargetPivotTable=defaultNamedNotOptArg, ValueChangeStart=defaultNamedNotOptArg, ValueChangeEnd=defaultNamedNotOptArg):
#	def OnCalculate(self):
#	def OnRelease(self):
#	def OnAddRef(self):
#	def OnGetTypeInfoCount(self, pctinfo=pythoncom.Missing):
#	def OnActivate(self):
#	def OnPivotTableAfterValueChange(self, TargetPivotTable=defaultNamedNotOptArg, TargetRange=defaultNamedNotOptArg):
#	def OnDeactivate(self):
#	def OnQueryInterface(self, riid=defaultNamedNotOptArg, ppvObj=pythoncom.Missing):
#	def OnInvoke(self, dispidMember=defaultNamedNotOptArg, riid=defaultNamedNotOptArg, lcid=defaultNamedNotOptArg, wFlags=defaultNamedNotOptArg
#			, pdispparams=defaultNamedNotOptArg, pvarResult=pythoncom.Missing, pexcepinfo=pythoncom.Missing, puArgErr=pythoncom.Missing):
#	def OnPivotTableChangeSync(self, Target=defaultNamedNotOptArg):
#	def OnPivotTableBeforeAllocateChanges(self, TargetPivotTable=defaultNamedNotOptArg, ValueChangeStart=defaultNamedNotOptArg, ValueChangeEnd=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnBeforeRightClick(self, Target=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnChange(self, Target=defaultNamedNotOptArg):
#	def OnFollowHyperlink(self, Target=defaultNamedNotOptArg):
#	def OnBeforeDoubleClick(self, Target=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnPivotTableBeforeCommitChanges(self, TargetPivotTable=defaultNamedNotOptArg, ValueChangeStart=defaultNamedNotOptArg, ValueChangeEnd=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnPivotTableUpdate(self, Target=defaultNamedNotOptArg):
#	def OnGetIDsOfNames(self, riid=defaultNamedNotOptArg, rgszNames=defaultNamedNotOptArg, cNames=defaultNamedNotOptArg, lcid=defaultNamedNotOptArg
#			, rgdispid=pythoncom.Missing):


win32com.client.CLSIDToClass.RegisterCLSID( "{00024411-0000-0000-C000-000000000046}", DocEvents )
