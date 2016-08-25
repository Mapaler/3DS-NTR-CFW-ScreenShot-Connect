Attribute VB_Name = "mTest"
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Public Declare Function VirtualProtect Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, ByRef lpflOldProtect As Long) As Long
Public Const PAGE_EXECUTE_READWRITE As Long = &H40&
Public lct2 As Long

Public Function SwapVtableEntry(pObj As Long, EntryNumber As Integer, ByVal lpfn As Long) As Long

    Dim lOldAddr As Long
    Dim lpVtableHead As Long
    Dim lpfnAddr As Long
    Dim lOldProtect As Long

    CopyMemory lpVtableHead, ByVal pObj, 4
    lpfnAddr = lpVtableHead + (EntryNumber - 1) * 4
    CopyMemory lOldAddr, ByVal lpfnAddr, 4

    Call VirtualProtect(lpfnAddr, 4, PAGE_EXECUTE_READWRITE, lOldProtect)
    CopyMemory ByVal lpfnAddr, lpfn, 4
    Call VirtualProtect(lpfnAddr, 4, lOldProtect, lOldProtect)

    SwapVtableEntry = lOldAddr

End Function
Public Function OnFolderChangingVB(ByVal this As IFileDialogEvents, ByVal pdf As IFileDialog, ByVal psiFolder As IShellItem) As Long
'OnFolderChangingVB = 0 (S_OK) or E_NOTIMPL - ok to change folder
'anything else, does not allow folder to change
'Form1.List1.AddItem "OnFolderChanging"
End Function
Public Function OnFileOkVB(ByVal this As IFileDialogEvents, ByVal pfd As IFileDialog) As Long
Dim bAllowContinue As Long

'Form1.List1.AddItem "OnFileOk"
'Form1.FD_OnFileOk bAllowContinue
OnFileOkVB = bAllowContinue
End Function
