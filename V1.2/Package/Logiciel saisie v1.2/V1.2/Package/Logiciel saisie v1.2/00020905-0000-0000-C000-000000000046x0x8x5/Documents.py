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
class Documents(DispatchBaseClass):
	CLSID = IID('{0002096C-0000-0000-C000-000000000046}')
	coclass_clsid = None

	# Result is of type Document
	def Add(self, Template=defaultNamedOptArg, NewTemplate=defaultNamedOptArg, DocumentType=defaultNamedOptArg, Visible=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(14, LCID, 1, (13, 0), ((16396, 17), (16396, 17), (16396, 17), (16396, 17)),Template
			, NewTemplate, DocumentType, Visible)
		if ret is not None:
			# See if this IUnknown is really an IDispatch
			try:
				ret = ret.QueryInterface(pythoncom.IID_IDispatch)
			except pythoncom.error:
				return ret
			ret = Dispatch(ret, 'Add', '{00020906-0000-0000-C000-000000000046}')
		return ret

	# Result is of type Document
	def AddBlogDocument(self, ProviderID=defaultNamedNotOptArg, PostURL=defaultNamedNotOptArg, BlogName=defaultNamedNotOptArg, PostID=''):
		return self._ApplyTypes_(21, 1, (13, 32), ((8, 1), (8, 1), (8, 1), (8, 49)), 'AddBlogDocument', '{00020906-0000-0000-C000-000000000046}',ProviderID
			, PostURL, BlogName, PostID)

	# Result is of type Document
	def AddOld(self, Template=defaultNamedOptArg, NewTemplate=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(11, LCID, 1, (13, 0), ((16396, 17), (16396, 17)),Template
			, NewTemplate)
		if ret is not None:
			# See if this IUnknown is really an IDispatch
			try:
				ret = ret.QueryInterface(pythoncom.IID_IDispatch)
			except pythoncom.error:
				return ret
			ret = Dispatch(ret, 'AddOld', '{00020906-0000-0000-C000-000000000046}')
		return ret

	def CanCheckOut(self, FileName=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(17, LCID, 1, (11, 0), ((8, 1),),FileName
			)

	def CheckOut(self, FileName=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(16, LCID, 1, (24, 0), ((8, 1),),FileName
			)

	def Close(self, SaveChanges=defaultNamedOptArg, OriginalFormat=defaultNamedOptArg, RouteDocument=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1105, LCID, 1, (24, 0), ((16396, 17), (16396, 17), (16396, 17)),SaveChanges
			, OriginalFormat, RouteDocument)

	# Result is of type Document
	def Item(self, Index=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(0, LCID, 1, (13, 0), ((16396, 1),),Index
			)
		if ret is not None:
			# See if this IUnknown is really an IDispatch
			try:
				ret = ret.QueryInterface(pythoncom.IID_IDispatch)
			except pythoncom.error:
				return ret
			ret = Dispatch(ret, 'Item', '{00020906-0000-0000-C000-000000000046}')
		return ret

	# Result is of type Document
	def Open(self, FileName=defaultNamedNotOptArg, ConfirmConversions=defaultNamedOptArg, ReadOnly=defaultNamedOptArg, AddToRecentFiles=defaultNamedOptArg
			, PasswordDocument=defaultNamedOptArg, PasswordTemplate=defaultNamedOptArg, Revert=defaultNamedOptArg, WritePasswordDocument=defaultNamedOptArg, WritePasswordTemplate=defaultNamedOptArg
			, Format=defaultNamedOptArg, Encoding=defaultNamedOptArg, Visible=defaultNamedOptArg, OpenAndRepair=defaultNamedOptArg, DocumentDirection=defaultNamedOptArg
			, NoEncodingDialog=defaultNamedOptArg, XMLTransform=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(19, LCID, 1, (13, 0), ((16396, 1), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),FileName
			, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, PasswordTemplate
			, Revert, WritePasswordDocument, WritePasswordTemplate, Format, Encoding
			, Visible, OpenAndRepair, DocumentDirection, NoEncodingDialog, XMLTransform
			)
		if ret is not None:
			# See if this IUnknown is really an IDispatch
			try:
				ret = ret.QueryInterface(pythoncom.IID_IDispatch)
			except pythoncom.error:
				return ret
			ret = Dispatch(ret, 'Open', '{00020906-0000-0000-C000-000000000046}')
		return ret

	# Result is of type Document
	def Open2000(self, FileName=defaultNamedNotOptArg, ConfirmConversions=defaultNamedOptArg, ReadOnly=defaultNamedOptArg, AddToRecentFiles=defaultNamedOptArg
			, PasswordDocument=defaultNamedOptArg, PasswordTemplate=defaultNamedOptArg, Revert=defaultNamedOptArg, WritePasswordDocument=defaultNamedOptArg, WritePasswordTemplate=defaultNamedOptArg
			, Format=defaultNamedOptArg, Encoding=defaultNamedOptArg, Visible=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(15, LCID, 1, (13, 0), ((16396, 1), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),FileName
			, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, PasswordTemplate
			, Revert, WritePasswordDocument, WritePasswordTemplate, Format, Encoding
			, Visible)
		if ret is not None:
			# See if this IUnknown is really an IDispatch
			try:
				ret = ret.QueryInterface(pythoncom.IID_IDispatch)
			except pythoncom.error:
				return ret
			ret = Dispatch(ret, 'Open2000', '{00020906-0000-0000-C000-000000000046}')
		return ret

	# Result is of type Document
	def Open2002(self, FileName=defaultNamedNotOptArg, ConfirmConversions=defaultNamedOptArg, ReadOnly=defaultNamedOptArg, AddToRecentFiles=defaultNamedOptArg
			, PasswordDocument=defaultNamedOptArg, PasswordTemplate=defaultNamedOptArg, Revert=defaultNamedOptArg, WritePasswordDocument=defaultNamedOptArg, WritePasswordTemplate=defaultNamedOptArg
			, Format=defaultNamedOptArg, Encoding=defaultNamedOptArg, Visible=defaultNamedOptArg, OpenAndRepair=defaultNamedOptArg, DocumentDirection=defaultNamedOptArg
			, NoEncodingDialog=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(18, LCID, 1, (13, 0), ((16396, 1), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),FileName
			, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, PasswordTemplate
			, Revert, WritePasswordDocument, WritePasswordTemplate, Format, Encoding
			, Visible, OpenAndRepair, DocumentDirection, NoEncodingDialog)
		if ret is not None:
			# See if this IUnknown is really an IDispatch
			try:
				ret = ret.QueryInterface(pythoncom.IID_IDispatch)
			except pythoncom.error:
				return ret
			ret = Dispatch(ret, 'Open2002', '{00020906-0000-0000-C000-000000000046}')
		return ret

	# Result is of type Document
	def OpenNoRepairDialog(self, FileName=defaultNamedNotOptArg, ConfirmConversions=defaultNamedOptArg, ReadOnly=defaultNamedOptArg, AddToRecentFiles=defaultNamedOptArg
			, PasswordDocument=defaultNamedOptArg, PasswordTemplate=defaultNamedOptArg, Revert=defaultNamedOptArg, WritePasswordDocument=defaultNamedOptArg, WritePasswordTemplate=defaultNamedOptArg
			, Format=defaultNamedOptArg, Encoding=defaultNamedOptArg, Visible=defaultNamedOptArg, OpenAndRepair=defaultNamedOptArg, DocumentDirection=defaultNamedOptArg
			, NoEncodingDialog=defaultNamedOptArg, XMLTransform=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(20, LCID, 1, (13, 0), ((16396, 1), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),FileName
			, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, PasswordTemplate
			, Revert, WritePasswordDocument, WritePasswordTemplate, Format, Encoding
			, Visible, OpenAndRepair, DocumentDirection, NoEncodingDialog, XMLTransform
			)
		if ret is not None:
			# See if this IUnknown is really an IDispatch
			try:
				ret = ret.QueryInterface(pythoncom.IID_IDispatch)
			except pythoncom.error:
				return ret
			ret = Dispatch(ret, 'OpenNoRepairDialog', '{00020906-0000-0000-C000-000000000046}')
		return ret

	# Result is of type Document
	def OpenOld(self, FileName=defaultNamedNotOptArg, ConfirmConversions=defaultNamedOptArg, ReadOnly=defaultNamedOptArg, AddToRecentFiles=defaultNamedOptArg
			, PasswordDocument=defaultNamedOptArg, PasswordTemplate=defaultNamedOptArg, Revert=defaultNamedOptArg, WritePasswordDocument=defaultNamedOptArg, WritePasswordTemplate=defaultNamedOptArg
			, Format=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(12, LCID, 1, (13, 0), ((16396, 1), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),FileName
			, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, PasswordTemplate
			, Revert, WritePasswordDocument, WritePasswordTemplate, Format)
		if ret is not None:
			# See if this IUnknown is really an IDispatch
			try:
				ret = ret.QueryInterface(pythoncom.IID_IDispatch)
			except pythoncom.error:
				return ret
			ret = Dispatch(ret, 'OpenOld', '{00020906-0000-0000-C000-000000000046}')
		return ret

	def Save(self, NoPrompt=defaultNamedOptArg, OriginalFormat=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(13, LCID, 1, (24, 0), ((16396, 17), (16396, 17)),NoPrompt
			, OriginalFormat)

	_prop_map_get_ = {
		# Method 'Application' returns object of type 'Application'
		"Application": (1000, 2, (13, 0), (), "Application", '{000209FF-0000-0000-C000-000000000046}'),
		"Count": (2, 2, (3, 0), (), "Count", None),
		"Creator": (1001, 2, (3, 0), (), "Creator", None),
		"Parent": (1002, 2, (9, 0), (), "Parent", None),
	}
	_prop_map_put_ = {
	}
	# Default method for this class is 'Item'
	def __call__(self, Index=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(0, LCID, 1, (13, 0), ((16396, 1),),Index
			)
		if ret is not None:
			# See if this IUnknown is really an IDispatch
			try:
				ret = ret.QueryInterface(pythoncom.IID_IDispatch)
			except pythoncom.error:
				return ret
			ret = Dispatch(ret, '__call__', '{00020906-0000-0000-C000-000000000046}')
		return ret

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
		return win32com.client.util.Iterator(ob, '{00020906-0000-0000-C000-000000000046}')
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		return self._ApplyTypes_(*(2, 2, (3, 0), (), "Count", None))
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

win32com.client.CLSIDToClass.RegisterCLSID( "{0002096C-0000-0000-C000-000000000046}", Documents )
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

Documents_vtables_dispatch_ = 1
Documents_vtables_ = [
	(( '_NewEnum' , 'prop' , ), -4, (-4, (), [ (16397, 10, None, None) , ], 1 , 2 , 4 , 0 , 28 , (3, 0, None, None) , 1024 , )),
	(( 'Count' , 'prop' , ), 2, (2, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 32 , (3, 0, None, None) , 0 , )),
	(( 'Application' , 'prop' , ), 1000, (1000, (), [ (16397, 10, None, "IID('{000209FF-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 36 , (3, 0, None, None) , 0 , )),
	(( 'Creator' , 'prop' , ), 1001, (1001, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 40 , (3, 0, None, None) , 0 , )),
	(( 'Parent' , 'prop' , ), 1002, (1002, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 44 , (3, 0, None, None) , 0 , )),
	(( 'Item' , 'Index' , 'prop' , ), 0, (0, (), [ (16396, 1, None, None) , 
			 (16397, 10, None, "IID('{00020906-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 0 , 48 , (3, 0, None, None) , 0 , )),
	(( 'Close' , 'SaveChanges' , 'OriginalFormat' , 'RouteDocument' , ), 1105, (1105, (), [ 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , ], 1 , 1 , 4 , 3 , 52 , (3, 0, None, None) , 0 , )),
	(( 'AddOld' , 'Template' , 'NewTemplate' , 'prop' , ), 11, (11, (), [ 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16397, 10, None, "IID('{00020906-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 2 , 56 , (3, 0, None, None) , 64 , )),
	(( 'OpenOld' , 'FileName' , 'ConfirmConversions' , 'ReadOnly' , 'AddToRecentFiles' , 
			 'PasswordDocument' , 'PasswordTemplate' , 'Revert' , 'WritePasswordDocument' , 'WritePasswordTemplate' , 
			 'Format' , 'prop' , ), 12, (12, (), [ (16396, 1, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16397, 10, None, "IID('{00020906-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 9 , 60 , (3, 0, None, None) , 64 , )),
	(( 'Save' , 'NoPrompt' , 'OriginalFormat' , ), 13, (13, (), [ (16396, 17, None, None) , 
			 (16396, 17, None, None) , ], 1 , 1 , 4 , 2 , 64 , (3, 0, None, None) , 0 , )),
	(( 'Add' , 'Template' , 'NewTemplate' , 'DocumentType' , 'Visible' , 
			 'prop' , ), 14, (14, (), [ (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16397, 10, None, "IID('{00020906-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 4 , 68 , (3, 0, None, None) , 0 , )),
	(( 'Open2000' , 'FileName' , 'ConfirmConversions' , 'ReadOnly' , 'AddToRecentFiles' , 
			 'PasswordDocument' , 'PasswordTemplate' , 'Revert' , 'WritePasswordDocument' , 'WritePasswordTemplate' , 
			 'Format' , 'Encoding' , 'Visible' , 'prop' , ), 15, (15, (), [ 
			 (16396, 1, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16397, 10, None, "IID('{00020906-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 11 , 72 , (3, 0, None, None) , 64 , )),
	(( 'CheckOut' , 'FileName' , ), 16, (16, (), [ (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 76 , (3, 0, None, None) , 0 , )),
	(( 'CanCheckOut' , 'FileName' , 'prop' , ), 17, (17, (), [ (8, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 80 , (3, 0, None, None) , 0 , )),
	(( 'Open2002' , 'FileName' , 'ConfirmConversions' , 'ReadOnly' , 'AddToRecentFiles' , 
			 'PasswordDocument' , 'PasswordTemplate' , 'Revert' , 'WritePasswordDocument' , 'WritePasswordTemplate' , 
			 'Format' , 'Encoding' , 'Visible' , 'OpenAndRepair' , 'DocumentDirection' , 
			 'NoEncodingDialog' , 'prop' , ), 18, (18, (), [ (16396, 1, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16397, 10, None, "IID('{00020906-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 14 , 84 , (3, 0, None, None) , 64 , )),
	(( 'Open' , 'FileName' , 'ConfirmConversions' , 'ReadOnly' , 'AddToRecentFiles' , 
			 'PasswordDocument' , 'PasswordTemplate' , 'Revert' , 'WritePasswordDocument' , 'WritePasswordTemplate' , 
			 'Format' , 'Encoding' , 'Visible' , 'OpenAndRepair' , 'DocumentDirection' , 
			 'NoEncodingDialog' , 'XMLTransform' , 'prop' , ), 19, (19, (), [ (16396, 1, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16397, 10, None, "IID('{00020906-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 15 , 88 , (3, 0, None, None) , 0 , )),
	(( 'OpenNoRepairDialog' , 'FileName' , 'ConfirmConversions' , 'ReadOnly' , 'AddToRecentFiles' , 
			 'PasswordDocument' , 'PasswordTemplate' , 'Revert' , 'WritePasswordDocument' , 'WritePasswordTemplate' , 
			 'Format' , 'Encoding' , 'Visible' , 'OpenAndRepair' , 'DocumentDirection' , 
			 'NoEncodingDialog' , 'XMLTransform' , 'prop' , ), 20, (20, (), [ (16396, 1, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16397, 10, None, "IID('{00020906-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 15 , 92 , (3, 0, None, None) , 0 , )),
	(( 'AddBlogDocument' , 'ProviderID' , 'PostURL' , 'BlogName' , 'PostID' , 
			 'prop' , ), 21, (21, (), [ (8, 1, None, None) , (8, 1, None, None) , (8, 1, None, None) , 
			 (8, 49, "''", None) , (16397, 10, None, "IID('{00020906-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 0 , 96 , (3, 32, None, None) , 0 , )),
]

win32com.client.CLSIDToClass.RegisterCLSID( "{0002096C-0000-0000-C000-000000000046}", Documents )
