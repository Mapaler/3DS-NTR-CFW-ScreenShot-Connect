VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFileDialogEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'cFileDialogEvents
'Handles events in an IFileDialog dialog
'This class is used by creating a new instance and then
' using the IFileDialog.Advise method
'Default controls (IFileDialogEvents) and custom controls
' (IFileDialogControlEvents) must be handled here.

Implements IFileDialogEvents
Implements IFileDialogControlEvents
Private m_pOldOnFolderChanging As Long
Private m_pOldOnFileOK As Long


Private Sub Class_Initialize()
Dim pVTable As IFileDialogEvents
Set pVTable = Me
m_pOldOnFileOK = SwapVtableEntry(ObjPtr(pVTable), 4, AddressOf OnFileOkVB)
m_pOldOnFolderChanging = SwapVtableEntry(ObjPtr(pVTable), 5, AddressOf OnFolderChangingVB)


End Sub

Private Sub Class_Terminate()
Dim pVTable As IFileDialogEvents
Set pVTable = Me
m_pOldOnFileOK = SwapVtableEntry(ObjPtr(pVTable), 4, m_pOldOnFileOK)
m_pOldOnFolderChanging = SwapVtableEntry(ObjPtr(pVTable), 5, m_pOldOnFolderChanging)

End Sub

Private Sub IFileDialogEvents_OnFileOk(ByVal pfd As IFileDialog)
'Now implemented in module
'Form1.List1.AddItem "OnFileOk"
'Form1.FD_OnFileOk
End Sub
Public Sub IFileDialogEvents_OnFolderChange(ByVal pdf As IFileDialog)
'Form1.List1.AddItem "OnFolderChange"
End Sub
Public Sub IFileDialogEvents_OnFolderChanging(ByVal pdf As IFileDialog, ByVal psiFolder As IShellItem)
'Now implemented in module
'Form1.List1.AddItem "OnFolderChanging"

End Sub

Public Sub IFileDialogEvents_OnOverwrite(ByVal pdf As IFileDialog, ByVal psi As IShellItem, pResponse As FDE_OVERWRITE_RESPONSE)
'Form1.List1.AddItem "OnOverwrite"
End Sub

Public Sub IFileDialogEvents_OnSelectionChange(ByVal pdf As IFileDialog)
'Form1.List1.AddItem "OnSelectionChange"
End Sub

Public Sub IFileDialogEvents_OnShareViolation(ByVal pdf As IFileDialog, ByVal psi As IShellItem, pResponse As FDE_SHAREVIOLATION_RESPONSE)
'Form1.List1.AddItem "OnShareViolation"
End Sub
Public Sub IFileDialogEvents_OnTypeChange(ByVal pdf As IFileDialog)
'Form1.List1.AddItem "OnTypeChange"
End Sub

Public Sub IFileDialogControlEvents_OnButtonClicked(ByVal pfdc As IFileDialogCustomize, ByVal dwIDCtl As Long)
'Form1.List1.AddItem "OnButtonClicked, id=" & dwIDCtl
'Form1.FD_ButtonClick dwIDCtl
End Sub
Public Sub IFileDialogControlEvents_OnCheckButtonToggled(ByVal pfdc As IFileDialogCustomize, ByVal dwIDCtl As Long, ByVal bChecked As BOOL)
'Form1.List1.AddItem "OnCheckButtonToggled, id=" & dwIDCtl & ",checked=" & bChecked

End Sub
Public Sub IFileDialogControlEvents_OnControlActivating(ByVal pfdc As IFileDialogCustomize, ByVal dwIDCtl As Long)
'Form1.List1.AddItem "OnControlActivating, id=" & dwIDCtl

End Sub
Public Sub IFileDialogControlEvents_OnItemSelected(ByVal pfdc As IFileDialogCustomize, ByVal dwIDCtl As Long, ByVal dwIDItem As Long)
'Form1.List1.AddItem "OnItemSelected, id=" & dwIDCtl & "item=" & dwIDItem
'Form1.FD_ItemSelected dwIDCtl, dwIDItem
End Sub

