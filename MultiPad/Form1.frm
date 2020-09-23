VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Multi-Pad"
   ClientHeight    =   5430
   ClientLeft      =   1530
   ClientTop       =   1890
   ClientWidth     =   8550
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   8550
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   0
      MaxLength       =   23000
      MultiLine       =   -1  'True
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   2  'Automatic
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   8535
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.txt"
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save &As"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
         Begin VB.Menu mnuAll 
            Caption         =   "&All"
            Shortcut        =   ^P
         End
         Begin VB.Menu mnuSelected 
            Caption         =   "&Selected"
         End
      End
      Begin VB.Menu mnuPause 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuAlign 
         Caption         =   "&Align Text"
         Begin VB.Menu mnuLeft 
            Caption         =   "&Left"
         End
         Begin VB.Menu mnuCenter 
            Caption         =   "&Center"
         End
         Begin VB.Menu mnuRight 
            Caption         =   "&Right"
         End
      End
      Begin VB.Menu mnuType 
         Caption         =   "&Type"
         Begin VB.Menu mnuBold 
            Caption         =   "&Bold"
         End
         Begin VB.Menu mnuItalic 
            Caption         =   "&Italic"
         End
         Begin VB.Menu mnuUnderlined 
            Caption         =   "&Underlined"
         End
      End
      Begin VB.Menu mnuPuase 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUndo 
         Caption         =   "&Undo"
      End
      Begin VB.Menu mnuRedo 
         Caption         =   "&Redo"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuPause2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuClear 
         Caption         =   "Cl&ear"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuNormal 
         Caption         =   "&Normal"
      End
      Begin VB.Menu mnuFullScreen 
         Caption         =   "&FullScreen"
      End
      Begin VB.Menu mnuMinimize 
         Caption         =   "&Minimize"
      End
      Begin VB.Menu mnuText 
         Caption         =   "&Text"
         Begin VB.Menu mnuAl 
            Caption         =   "&All"
         End
         Begin VB.Menu mnuSelec 
            Caption         =   "&Selected only"
         End
      End
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "F&ormat"
      Begin VB.Menu mnuFont 
         Caption         =   "&Font..."
      End
      Begin VB.Menu mnuColors 
         Caption         =   "Colors"
         Begin VB.Menu mnuBackground 
            Caption         =   "&Background"
         End
         Begin VB.Menu mnuFontColor 
            Caption         =   "Font &Color"
         End
      End
      Begin VB.Menu mnuPause3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLock 
         Caption         =   "&Lock Text"
      End
      Begin VB.Menu mnuhide 
         Caption         =   "&Hide Text"
      End
      Begin VB.Menu mnuCase 
         Caption         =   "Change Cas&e"
         Begin VB.Menu mnuUpper 
            Caption         =   "UPPER CASE"
         End
         Begin VB.Menu mnulower 
            Caption         =   "lower case"
         End
         Begin VB.Menu mnuReverse 
            Caption         =   "Reverse esreveR"
         End
         Begin VB.Menu mnupuse 
            Caption         =   "-"
         End
         Begin VB.Menu mnuOld 
            Caption         =   "Undo"
            Enabled         =   0   'False
         End
      End
   End
   Begin VB.Menu mnuDocType 
      Caption         =   "&Document Type"
      Begin VB.Menu mnuLetter 
         Caption         =   "&Letter"
      End
      Begin VB.Menu mnuWebPage 
         Caption         =   "&Web  Page"
         Begin VB.Menu mnuBlank 
            Caption         =   "&Blank (new)"
         End
         Begin VB.Menu mnuPreview 
            Caption         =   "&Store And Preview..."
         End
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "&Search"
      Begin VB.Menu mnuFind 
         Caption         =   "&Find..."
      End
      Begin VB.Menu mnuRep 
         Caption         =   "&Replace..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpTopics 
         Caption         =   "&Help Topics"
      End
      Begin VB.Menu mnuPause4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "A&bout"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim changed As Boolean, saved, filen As String, webp, wordd, p As Variant, firstime As Boolean, old As String, fullscreen As Boolean, bcase As String, col As String

Sub LoadText(Lst As TextBox, file As String)
'Call LoadText (Text1,"C:\Windows\System\Saved.txt")
On Error GoTo error
Dim mystr As String
Open file For Input As #1
Do While Not EOF(1)
            Line Input #1, a$
            texto$ = texto$ + a$ + Chr$(13) + Chr$(10)
        Loop
        Lst = texto$
Close #1
Exit Sub
error:
x = MsgBox("Please enter a valid filename!", vbOKOnly, "Error")
End Sub

Sub SaveText(Lst As TextBox, file As String)
'Call SaveText (Text1,"C:\Windows\System\Saved.txt")
On Error GoTo error
Dim mystr As String
Open file For Output As #1
Print #1, Lst
Close 1
Exit Sub
error:
x = MsgBox("Action canceled due to bad filename given. Text has been saved to clipboard for safety reasons", vbOKOnly, "Error")
Clipboard.SetText Text1.Text
End Sub

Private Sub Form_Load()
webp = 1
firstime = True
gHTxtWnd = Text1.hwnd
col = ""
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.WindowState = 0 Then fullscreen = False
If fullscreen = False Then
Text1.Width = Form1.Width - 105
Text1.Height = Form1.Height - 680
Else
Text1.Width = Form1.Width - 105
Text1.Height = Form1.Height - 400
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim a
If changed = False Then
    End
 Else
    a = MsgBox("Do you want to save the changes to your text?", vbYesNo, "MultiPad New")
     If a = vbYes Then
        CommonDialog1.Filter = "Text Files|*.doc;*.txt|Webpages|*.html|Excel|*.xcls|All files|*.*"
        CommonDialog1.ShowSave
        Call SaveText(Text1, CommonDialog1.FileName)
     ElseIf a = vbNo Then
        End
     End If
End If
End Sub

Private Sub mnuAbout_Click()
Form3.Left = Me.Left + 1200
Form3.Top = Me.Top + 300
Form3.Show
Me.Enabled = False
End Sub

Private Sub mnuAl_Click()
Text1.Text = old
End Sub

Private Sub mnuAll_Click()
Printer.Print Text1.Text
Printer.EndDoc
End Sub

Private Sub mnuBackground_Click()
CommonDialog1.ShowColor
Text1.BackColor = CommonDialog1.Color
End Sub

Private Sub mnuBlank_Click()
Dim a As String
a = "<html>" + vbCrLf + vbCrLf + "<head>" + vbCrLf + "<title>New Page 1</title>" + vbCrLf + "</head>" + vbCrLf + vbCrLf + "<body>" + vbCrLf + vbCrLf + "</body>" + vbCrLf + vbCrLf + "</html>"
Text1.Text = a
End Sub

Private Sub mnuBold_Click()
If mnuBold.Checked = True Then
  Text1.FontBold = False
  mnuBold.Checked = False
Else
  Text1.FontBold = True
  mnuBold.Checked = True
End If
End Sub

Private Sub mnuCenter_Click()
Text1.Alignment = 2
End Sub

Private Sub mnuClear_Click()
Text1.Text = ""
End Sub

Private Sub mnuCopy_Click()
Clipboard.SetText Text1.SelText
End Sub

Private Sub mnuCut_Click()
     Clipboard.SetText Text1.SelText
     Text1.SelText = ""
End Sub

Private Sub mnuExit_Click()
If changed = False Then
    End
 Else
    MsgBox "Do you want to save the changes to your text?", vbYesNoCancel, "MultiPad Exit"
End If
End Sub

Private Sub mnuFind_Click()
Dim szFindString As String  ' initial string to find
Dim hCmdBtn As Long         ' handle of 'Find Next' command button
Dim strArr() As Byte        ' for API use
Dim i As Integer            ' position indicator in the loop
' Fill in the structure.
szFindString = "Find Me"

ReDim strArr(0 To Len(szFindString) - 1)

For i = 1 To Len(szFindString)
    strArr(i - 1) = Asc(Mid(szFindString, i, 1))
Next i

frText.flags = FR_MATCHCASE Or FR_NOUPDOWN Or FR_NOWHOLEWORD
frText.lpfnHook = 0&
frText.lpTemplateName = 0&
frText.lStructSize = Len(frText)
frText.hwndOwner = Me.hwnd
frText.hInstance = App.hInstance
frText.lpstrFindWhat = VarPtr(strArr(0))
frText.lpstrReplaceWith = 0&
frText.wFindWhatLen = Len(szFindString)
frText.wReplaceWithLen = 0
frText.lCustData = 0

' Show the dialog box.
gHDlg = FindText(frText)

' Get the handle of the dialog box
hCmdBtn = GetDlgItem(gHDlg, 1)

' Get necessary value for calling default window procedure.
gOldDlgWndHandle = GetWindowLong(hCmdBtn, GWL_WNDPROC)

If SetWindowLong(hCmdBtn, GWL_WNDPROC, AddressOf FindTextHookProc) = 0 Then
    gOldDlgWndHandle = 0
End If
End Sub

Private Sub mnuFont_Click()
Form2.Show
End Sub

Private Sub mnuFontColor_Click()
CommonDialog1.ShowColor
Text1.ForeColor = CommonDialog1.Color
End Sub

Private Sub mnuFullScreen_Click()
Me.Caption = Me.Caption & "- Fullscreen - Press ESC to view normal mode"
Me.WindowState = 2
mnuFile.Visible = False
mnuEdit.Visible = False
mnuDocType.Visible = False
mnuHelp.Visible = False
mnuSearch.Visible = False
mnuView.Visible = False
fullscreen = True
End Sub

Private Sub mnuHelpTopics_Click()
Form4.Show
End Sub

Private Sub mnuhide_Click()
If mnuhide.Checked = False Then
mnuhide.Checked = True
col = Text1.ForeColor
Text1.ForeColor = Text1.BackColor
Else
Text1.ForeColor = col
mnuhide.Checked = False
End If
End Sub

Private Sub mnuItalic_Click()
If mnuItalic.Checked = True Then
  Text1.FontItalic = False
  mnuItalic.Checked = False
Else
  Text1.FontItalic = True
  mnuItalic.Checked = True
End If
End Sub

Private Sub mnuLeft_Click()
Text1.Alignment = 0
End Sub

Private Sub mnuLetter_Click()
Text1.Text = "Dear " & InputBox("Refering to...:", "Info") & "," + vbCrLf + vbCrLf + vbCrLf + vbCrLf + vbCrLf + vbCrLf + vbCrLf + vbCrLf + vbCrLf + vbCrLf + vbCrLf + vbCrLf + vbCrLf + vbCrLf + vbCrLf + vbCrLf & InputBox("How to close letter? e.g King Regards", "Info") & "," + vbCrLf + InputBox("Your name?", "Info")
End Sub

Private Sub mnuLock_Click()
If mnuLock.Checked = False Then
mnuLock.Checked = True
Text1.Locked = True
Else
mnuLock.Checked = False
Text1.Locked = False
End If
End Sub

Private Sub mnulower_Click()
bcase = Text1.Text
mnuOld.Enabled = True
Text1.Text = LCase(Text1.Text)
End Sub

Private Sub mnuMinimize_Click()
Me.WindowState = 1
End Sub

Private Sub mnuNew_Click()
If changed = False Then
    Text1.Text = ""
 Else
    MsgBox "Do you want to save the changes to your text?", vbYesNoCancel, "MultiPad New"
End If
End Sub

Private Sub mnuNormal_Click()
Me.WindowState = 0
End Sub

Private Sub mnuOld_Click()
Text1.Text = bcase
mnuOld.Enabled = False
End Sub

Private Sub mnuOpen_Click()
CommonDialog1.Filter = "Text Files|*.doc;*.txt|Webpages|*.html|Excel|*.xcls|All files|*.*"
CommonDialog1.ShowOpen
Form1.Caption = "Loading..."
Call LoadText(Text1, CommonDialog1.FileName)
Form1.Caption = "Multi-Pad"
End Sub

Private Sub mnuPaste_Click()
Text1.SelText = Clipboard.GetText()
End Sub

Private Sub mnuPreview_Click()
If webp < 2 Then
MsgBox "Web page will be saved in C:\WebPage.htm", vbInformation, ""
webp = webp + 1
End If
Call SaveText(Text1, "C:\WebPage.htm")
retval = ShellExecute(0&, vbNullString, "c:\WebPage.htm", vbNullString, vbNullString, vbNormalFocus)
End Sub

Private Sub mnuRedo_Click()
SendKeys "^{z}"
mnuRedo.Enabled = False
mnuUndo.Enabled = False
End Sub

Private Sub mnuRep_Click()
Dim a
a = InputBox("Replace with", "Replace", Text1.SelText)
Text1.SelText = a
End Sub

Private Sub mnuReverse_Click()
bcase = Text1.Text
mnuOld.Enabled = True
Text1.Text = StrReverse(Text1.Text)
End Sub

Private Sub mnuRight_Click()
Text1.Alignment = 1
End Sub

Private Sub mnuSave_Click()
If saved = True Then
   Call SaveText(Text1, filen)
End If
End Sub

Private Sub mnuSaveAs_Click()
Dim a
If saved = False Then
   CommonDialog1.Filter = "Text Files|*.doc;*.txt|Webpages|*.html|Excel|*.xcls|All files|*.*"
   CommonDialog1.ShowSave
   filen = CommonDialog1.FileName
   Call SaveText(Text1, CommonDialog1.FileName)
   mnuSaveAs.Enabled = False
   mnuSave.Enabled = True
   saved = True
  Else
   Exit Sub
End If
End Sub

Private Sub mnuSelec_Click()
old = Text1.Text
Text1.Text = Text1.SelText
End Sub

Private Sub mnuSelectAll_Click()
Text1.SelStart = 0
Text1.SelLength = Len(Text1)
End Sub

Private Sub mnuSelected_Click()
Printer.Print Text1.SelText
Printer.EndDoc
End Sub

Private Sub mnuUnderlined_Click()
If mnuUnderlined.Checked = True Then
  Text1.FontUnderline = False
  mnuUnderlined.Checked = False
Else
  Text1.FontUnderline = True
  mnuUnderlined.Checked = True
End If
End Sub

Private Sub mnuUndo_Click()
SendKeys "^{z}"
mnuRedo.Enabled = True
mnuUndo.Enabled = False
End Sub

Private Sub mnuUpper_Click()
bcase = Text1.Text
mnuOld.Enabled = True
Text1.Text = UCase(Text1.Text)
End Sub

Private Sub Text1_Change()
gTxtSrc = Text1.Text
If Text1.Text = "" Then
changed = False
Else
changed = True
End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If fullscreen = True Then
  If KeyCode = vbKeyEscape Then
   Me.WindowState = 0
   mnuFile.Visible = True
   mnuEdit.Visible = True
   mnuDocType.Visible = True
   mnuHelp.Visible = True
   mnuSearch.Visible = True
   mnuView.Visible = True
   fullscreen = False
   Me.Caption = "Multi-Pad"
  End If
End If
End Sub
