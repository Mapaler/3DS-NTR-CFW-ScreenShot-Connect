VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "3DS NTR截图自动拼接"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4005
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   4005
   StartUpPosition =   1  '所有者中心
   Begin VB.ListBox lstLog 
      Height          =   2040
      Left            =   120
      TabIndex        =   8
      Top             =   2880
      Width           =   3735
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "拼接"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   3735
   End
   Begin VB.CheckBox chkSubfolder 
      Caption         =   "输入目录内“connect”子文件夹"
      Height          =   180
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   3015
   End
   Begin VB.CommandButton cmdOutput 
      Caption         =   "浏览"
      Height          =   375
      Left            =   2880
      OLEDropMode     =   1  'Manual
      TabIndex        =   5
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox txtOutput 
      Height          =   375
      Left            =   120
      OLEDropMode     =   1  'Manual
      TabIndex        =   4
      Top             =   1200
      Width           =   2655
   End
   Begin VB.CommandButton cmdInput 
      Caption         =   "浏览"
      Height          =   375
      Left            =   2880
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox txtInput 
      Height          =   375
      Left            =   120
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label lblProgress 
      BackStyle       =   0  'Transparent
      Caption         =   "进度"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "输出目录"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "输入目录"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'IFileDialog Sample Project
'
'This project shows how to use the IFileOpenDialog and IFileSaveDialog interfaces introduced
'in Vista that are meant to replace the older GetOpenFileName API
'
'There's pros and cons, but the new interfaces are easier to use once they've been defined
'It's so straightforward that I didn't think a class module wrapper was worthwhile, since
'it is just like using the old common dialog ocx.
'
'REQUIREMENTS
'olelib.tlb - If you are already using this, you must replace it with the upgraded version
'             included with this project. I've maintained very tight compatibility, with
'             only minor changes where they were absolutely needed. See the changes.txt
'             file for more information, but the changes only effect shell32.dll API declares,
'             and the IShellFolder/IShellFolder2/IEnumIDList interfaces
'
'oleexp.tlb - This contains all the modern interfaces used by this project and my other projects,
'             it is an expansion of olelib and depends on that.
'
'cFileDialogEvents - This class module handles events from the dialog and from custom controls.
'                    It's an optional component that isn't needed for bare-bones functionality.
'
'(c)2015 by fafalone
'Feel free to re-use any of this code in any way you see fit as long as you credit

Private Type COMDLG_FILTERSPEC
    pszName As String
    pszSpec As String
End Type

Private Declare Function SHCreateShellItem Lib "shell32" (ByVal pidlParent As Long, ByVal psfParent As Long, ByVal pidl As Long, ppsi As IShellItem) As Long
Private Declare Function SysReAllocString Lib "oleaut32" (ByVal pBSTR As Long, ByVal lpWStr As Long) As Long
Private Declare Function ILCreateFromPathW Lib "shell32" (ByVal pwszPath As Long) As Long
Private Declare Function SHGetKnownFolderIDList Lib "shell32" (rfid As UUID, ByVal dwFlags As Long, ByVal hToken As Long, ppidl As Long) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpszGuid As Long, pGuid As Any) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)

Private Const fidPictures As String = "{33E28130-4E1E-4676-835A-98395C3BC3BB}"

Private fod As FileOpenDialog
Private fodSimple As FileOpenDialog
Private fodMulti As FileOpenDialog
Private fsd As FileSaveDialog
Private cEvents As cFileDialogEvents
Private fdc As IFileDialogCustomize


Private Const scrShotPtn As String = "top_(\d{4})\.bmp" '图片名称的匹配式
Private fso As Object


Private Sub chkSubfolder_Click()
    If chkSubfolder.value = Checked Then
        txtOutput.Enabled = False
        cmdOutput.Enabled = False
        fillSubFolderOutTxt
    Else
        txtOutput.Enabled = True
        cmdOutput.Enabled = True
    End If
End Sub

Private Sub cmdInput_Click()
'Shows the simplest Open File Dialog
On Error Resume Next 'A major error is thrown when the user cancels the dialog box

Dim isiRes As IShellItem
Dim lPtr As Long
Dim lOptions As FILEOPENDIALOGOPTIONS
Dim StrTmp As String

Set fodSimple = New FileOpenDialog

With fodSimple
    .SetTitle "选择输入文件夹"
    
    'When setting options, you should first get them
    .GetOptions lOptions
    lOptions = lOptions Or FOS_FILEMUSTEXIST Or FOS_PICKFOLDERS 'just an example of options... shows hidden files even if they're normally not shown
    .SetOptions lOptions
        
    .Show Me.hWnd
    
    .GetResult isiRes
    isiRes.GetDisplayName SIGDN_FILESYSPATH, lPtr
    StrTmp = BStrFromLPWStr(lPtr, True)
    If Len(StrTmp) Then txtInput.Text = StrTmp
    
End With
Set isiRes = Nothing
Set fodSimple = Nothing
End Sub


Private Sub cmdOutput_Click()
'Shows the simplest Open File Dialog
On Error Resume Next 'A major error is thrown when the user cancels the dialog box

Dim isiRes As IShellItem
Dim lPtr As Long
Dim lOptions As FILEOPENDIALOGOPTIONS
Dim StrTmp As String

Set fodSimple = New FileOpenDialog

With fodSimple
    .SetTitle "选择输出文件夹"
    
    'When setting options, you should first get them
    .GetOptions lOptions
    lOptions = lOptions Or FOS_FILEMUSTEXIST Or FOS_PICKFOLDERS 'just an example of options... shows hidden files even if they're normally not shown
    .SetOptions lOptions
    
    .Show Me.hWnd
    
    .GetResult isiRes
    isiRes.GetDisplayName SIGDN_FILESYSPATH, lPtr
    StrTmp = BStrFromLPWStr(lPtr, True)
    If Len(StrTmp) Then txtInput.Text = StrTmp
    
End With
Set isiRes = Nothing
Set fodSimple = Nothing
End Sub

Private Sub fillSubFolderOutTxt()
'添加 connect 子文件夹
    txtOutput.Text = txtInput & "\connect"
End Sub


Private Sub Form_Load()
    Set fso = CreateObject("Scripting.FileSystemObject") '文件操作系统对象
    chkSubfolder_Click
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub txtInput_Change()
    If chkSubfolder.value = Checked Then
        fillSubFolderOutTxt
    End If
End Sub

Private Sub txtInput_OLEDragDrop(data As DataObject, effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    If data.Files.Count > 0 Then '如果传入文件
        sendPathToTextBox data.Files(1), txtInput
    End If
End Sub

Private Sub cmdInput_OLEDragDrop(data As DataObject, effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    If data.Files.Count > 0 Then '如果传入文件
        sendPathToTextBox data.Files(1), txtInput
    End If
End Sub

Private Sub txtOutput_OLEDragDrop(data As DataObject, effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    If data.Files.Count > 0 Then '如果传入文件
        sendPathToTextBox data.Files(1), txtOutput
    End If
End Sub

Private Sub cmdOutput_OLEDragDrop(data As DataObject, effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    If data.Files.Count > 0 Then '如果传入文件
        sendPathToTextBox data.Files(1), txtOutput
    End If
End Sub

'把文件夹路径发送到文本框
Private Sub sendPathToTextBox(file As String, txtBox As TextBox)

    Dim ipt As String
    ipt = file
    
    If fso.FileExists(ipt) Then
        ipt = fso.GetParentFolderName(ipt)
    End If
    
    If fso.FolderExists(ipt) Then
        txtBox.Text = ipt
    End If
    
End Sub

Private Sub cmdStart_Click()
    Dim ipt As String, opt As String
    
    ipt = txtInput.Text
    opt = txtOutput.Text
    If Not fso.FolderExists(ipt) Then
        MsgBox "输入目录不存在", vbCritical
        Exit Sub
    End If
    
    If Not fso.FolderExists(opt) And fso.DriveExists(fso.GetDriveName(opt)) Then
        fso.CreateFolder opt '创建不存在的输出目录
    End If
    
    If Not fso.FolderExists(opt) Then
        MsgBox "输出目录不存在", vbCritical
        Exit Sub
    End If

    dealPictures ipt, opt
End Sub

Private Sub dealPictures(ipt As String, opt As String)
    Dim Ipto As Object, Opto As Object, Flo As Object
    Dim regRes As MatchCollection
    Dim picIndex As String
    
    Set Ipto = fso.GetFolder(ipt)
    
    lstLog.Clear
    
    For Each Flo In Ipto.Files
        If RegExpTest(Flo.name, scrShotPtn) Then
            Set regRes = RegExpSearch(Flo.name, scrShotPtn)
            'regRes(0).SubMatches(0) 获得的是 文件名里的4位数字编号
            picIndex = regRes(0).SubMatches(0)
            connectPictureFronIndex picIndex, ipt, opt
        
            lstLog.AddItem "已连接 " & getPicName("connect", picIndex, "png")
            lstLog.Selected(lstLog.ListCount - 1) = True
        End If
    Next

    lstLog.AddItem "文件夹扫描完毕"
    lstLog.Selected(lstLog.ListCount - 1) = True
End Sub

Private Function getPicName(prefix As String, index As String, Optional extension As String = "bmp") As String
'从数字获取图片名
    getPicName = prefix & "_" & index & "." & extension
End Function

Private Sub connectPictureFronIndex(index As String, ipt As String, opt As String)
'从数字来连接图片
    connectPicture ipt & "\" & getPicName("top", index), ipt & "\" & getPicName("bot", index), opt & "\" & getPicName("connect", index, "png")
End Sub

Private Sub connectPicture(Top As String, Bottom As String, output As String)
'连接图片的函数
    Dim Bitmap_top As Long, Bitmap_bottom As Long, Bitmap_output As Long, Graphics As Long
    'Dim Bitmap_BGout As Long, Bitmap_BGt As Long, Bitmap_Fout As Long, Bitmap_Ft As Long, Graphics As Long
    'Dim bmW_BG As Long, bmH_BG As Long, bmW_F As Long, bmH_F As Long
    InitGDIPlus
    
    '从文件载入Bitmap
    GdipCreateBitmapFromFile StrPtr(Top), Bitmap_top
    GdipCreateBitmapFromFile StrPtr(Bottom), Bitmap_bottom
    
    '获取图像尺寸
    'GdipGetImageWidth Bitmap_Ft, bmW_F
    'GdipGetImageHeight Bitmap_Ft, bmH_F

    CreateBitmapWithGraphics Bitmap_output, Graphics, 400, 480 '创建400x480的画布，将一个Image和Graphics关联

    GdipDrawImageRectI Graphics, Bitmap_top, 0, 0, 400, 240 '画上屏
    GdipDrawImageRectI Graphics, Bitmap_bottom, 40, 240, 320, 240 '画下屏
    
    SaveImageToPNG Bitmap_output, output

    '扫地工作
    GdipDeleteGraphics Graphics
    
    GdipDisposeImage Bitmap_top
    GdipDisposeImage Bitmap_bottom
    GdipDisposeImage Bitmap_output
    
    TerminateGDIPlus
End Sub


'是否符合正则表达式
Private Function RegExpTest(strng, patrn) As Boolean
    Dim regEx As RegExp
    
    Set regEx = New RegExp         ' 创建正则表达式。
    regEx.Pattern = patrn         ' 设置模式。
    regEx.IgnoreCase = True         ' 设置是否区分大小写，True为不区分。
    regEx.Global = True         ' 设置全程匹配。
    RegExpTest = regEx.Test(strng)   ' 执行搜索。
    Set regEx = Nothing
End Function
'正则表达式搜索
Private Function RegExpSearch(strng, patrn) As MatchCollection
    Dim regEx As RegExp
    
    Set regEx = New RegExp         ' 创建正则表达式。
    regEx.Pattern = patrn         ' 设置模式。
    regEx.IgnoreCase = True         ' 设置是否区分大小写，True为不区分。
    regEx.Global = True         ' 设置全程匹配。
    regEx.MultiLine = True
    Set RegExpSearch = regEx.Execute(strng)
'   If RegExpSearch.Count > 0 Then
'       MsgBox RegExpSearch.Item(0)
'       If RegExpSearch.Item(0).Submatches.Count > 0 Then
'           Set SubMatches = RegExpSearch.Item(0).Submatches
'           MsgBox SubMatches.Item(0)
'       End If
'   End If
    Set regEx = Nothing
End Function

'Helper Functions
Private Function GetPIDLFromPathW(sPath As String) As Long
   GetPIDLFromPathW = ILCreateFromPathW(StrPtr(sPath))
End Function
Private Function BStrFromLPWStr(lpWStr As Long, Optional ByVal CleanupLPWStr As Boolean = True) As String
SysReAllocString VarPtr(BStrFromLPWStr), lpWStr
If CleanupLPWStr Then CoTaskMemFree lpWStr
End Function

Private Sub Form_Unload(Cancel As Integer)
Set cEvents = Nothing
Set fod = Nothing
Set fsd = Nothing
Set fodSimple = Nothing
Set fdc = Nothing
End Sub
