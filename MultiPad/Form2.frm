VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   Caption         =   "Select Font:"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4590
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   4590
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1920
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Standard Font Dialog"
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   2840
      Width           =   1695
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Underlined"
      Height          =   195
      Left            =   3360
      TabIndex        =   8
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Italic"
      Height          =   195
      Left            =   2160
      TabIndex        =   7
      Top             =   2280
      Width           =   735
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Bold"
      Height          =   195
      Left            =   1080
      TabIndex        =   6
      Top             =   2280
      Width           =   615
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      Max             =   72
      Min             =   8
      TabIndex        =   3
      Top             =   2520
      Value           =   8
      Width           =   4575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save && Exit"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   2840
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Undo Changes"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   2840
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Double click to apply"
      Top             =   0
      Width           =   4575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Example"
      Height          =   735
      Left            =   0
      TabIndex        =   5
      Top             =   3240
      Width           =   4575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Size:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   0
      TabIndex        =   4
      Top             =   2205
      Width           =   525
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim old

Private Sub Check1_Click()
If Check1.Value = 1 Then
  Label2.FontBold = True
  Check1.Value = 1
Else
  Label2.FontBold = False
  Check1.Value = 0
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
  Label2.FontItalic = True
  Check2.Value = 1
Else
  Label2.FontItalic = False
  Check2.Value = 0
End If
End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
  Label2.FontUnderline = True
  Check3.Value = 1
Else
  Label2.FontUnderline = False
  Check3.Value = 0
End If
End Sub

Private Sub Command1_Click()
    CommonDialog1.Flags = cdlCFBoth
    CommonDialog1.ShowFont
    Form1.Text1.FontName = CommonDialog1.FontName
    Form1.Text1.FontSize = CommonDialog1.FontSize
    Form1.Text1.FontBold = CommonDialog1.FontBold
    Form1.Text1.FontItalic = CommonDialog1.FontItalic
    Form2.Hide
    Form1.Enabled = True
End Sub

Private Sub Command2_Click()
Form1.Text1.Font = old
End Sub

Private Sub Command3_Click()
If List1.Text = "" Then
MsgBox "Please select a font", vbCritical, "Font"
Exit Sub
End If
Form1.Enabled = True
Form1.Text1.Font = Form2.List1.Text
Form1.Text1.FontSize = HScroll1.Value
Command2.Enabled = True
If Check1.Value = 1 Then Form1.Text1.FontBold = True
If Check2.Value = 1 Then Form1.Text1.FontItalic = True
If Check3.Value = 1 Then Form1.Text1.FontUnderline = True
If Check1.Value = 0 Then Form1.Text1.FontBold = False
If Check2.Value = 0 Then Form1.Text1.FontItalic = False
If Check3.Value = 0 Then Form1.Text1.FontUnderline = False
Unload Form2
End Sub

Private Sub Form_Load()

    Dim NUM As Single
    Dim x As Single
    '- gets the numbers of fonts you have
    If Form1.Text1.FontBold = True Then Check1.Value = 1
    If Form1.Text1.FontItalic = True Then Check2.Value = 1
    If Form1.Text1.FontUnderline = True Then Check3.Value = 1
    If Form1.Text1.FontBold = False Then Check1.Value = 0
    If Form1.Text1.FontItalic = False Then Check2.Value = 0
    If Form1.Text1.FontUnderline = False Then Check3.Value = 0
    NUM = Screen.FontCount
    old = Form1.Text1.Font
    '- Set the listbox properties
    '- Set List1, Sorted = True
    '- Goes from 1 to number of fonts
    Form1.Enabled = False

    For x = 1 To NUM
        List1.AddItem Screen.Fonts(x)
    Next x
    '- for some reason there will be a blank
    '     itme
    '- this removes it
    List1.RemoveItem (0)

End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.Enabled = True
End Sub

Private Sub HScroll1_Change()
Label2.FontSize = HScroll1.Value
End Sub

Private Sub HScroll1_Scroll()
Label2.FontSize = HScroll1.Value
End Sub

Private Sub List1_Click()
Label2.Caption = List1.Text
Label2.Font = List1.Text
End Sub

Private Sub List1_DblClick()
Form1.Text1.Font = List1.Text
Command2.Enabled = True
End Sub
