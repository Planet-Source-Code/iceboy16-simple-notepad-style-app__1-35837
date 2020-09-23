VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Help"
   ClientHeight    =   5190
   ClientLeft      =   4440
   ClientTop       =   2430
   ClientWidth     =   2895
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   2895
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   4800
      Width           =   2900
   End
   Begin VB.TextBox Text2 
      Height          =   2535
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Form4.frx":0000
      Top             =   2235
      Width           =   2895
   End
   Begin VB.ListBox List1 
      Height          =   2205
      ItemData        =   "Form4.frx":0018
      Left            =   0
      List            =   "Form4.frx":0040
      TabIndex        =   0
      Top             =   0
      Width           =   2895
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Hide
Form1.Show
Form1.Enabled = True
End Sub

Private Sub List1_Click()
If List1.Text = "New" Then Text2.Text = "Creates a new (blank) document. Ask's you if you want to save changes (if made) to the old document before proceeding."
If List1.Text = "Open" Then Text2.Text = "Opens an already made document stored in your hard drive. For right use please choose .txt files. You can also view other type of file's source code (e.g .exe)"
If List1.Text = "Save & Save As" Then Text2.Text = "After finishing document you might want to save it, so you an use it at another time. With this option you save your file into your hard disk, and open it later with the 'Open' option." + vbCrLf + "-Save: It gets enabled after your save your document. This way you don't have to specify file's name every time, so save changes and overwrite old instead." + vbCrLf + "-Save as: Lets you specify the file's name or even save it more than one times (diffrent name)"
If List1.Text = "Print" Then Text2.Text = "Let's you print your document to your default (installed) printer." + vbCrLf + "-All: Prints the whole document" + vbCrLf + "-Selected: Print's only the word's you have highlighted (drag mouse over)"
If List1.Text = "Undo, Redo" Then Text2.Text = "Let's you correct a mistake that you might have done during the editing period. For example, if you erase the whole document, pressing 'undo' will restore it back." + vbCrLf + "-Undo: One action back" + vbCrLf + "-ReDo: One action forward. This command is used if you have selected 'undo' and you want to restore back the original changes"
If List1.Text = "Copy, Paste, Cut, Select all, Delete" Then Text2.Text = "-Copy: Copies the selected (highlighted) text to your clipboard." + vbCrLf + "-Paste: Writes the copied (or cut) text to the current cursor area" + vbCrLf + "-Cut: Removes the selected text, without losing it. The text gets written to the clipboard and you can paste it anywhere you want with 'paste' command" + vbCrLf + "-Select All: You can select all the document(words), and then cut the or copy them easily, without having to drag all text"
If List1.Text = "Find & find next" Then Text2.Text = "Let's you find a word or a phrase from the document. You simply specify the text and it automaticly highlights it for you"
If List1.Text = "Selecting Font" Then Text2.Text = "Font's are letters, only with diffrent character. All fonts are installed into your windows root\fonts\ directory and can be viewed and selected with this option." + vbCrLf + "How to use:" + vbCrLf + "Click on the font once to see its preview, or double click to auto apply. You can undo your apply by clicking 'Undo Changes'. You can also select if you want your text to be bold, underlined or italic by checking the boxes. Finally select the size of your font and then click 'Save & Exit'. You can view how the text will look like in the box below. Changes will not be applied untill you press 'Save & Quit'"
If List1.Text = "Selecting Color " Then Text2.Text = "You can change the background color (white), to another color you want. Simply choose the color from the dialog that will popup. You can also change your font's color by doing the same actions"
If List1.Text = "Document Type" Then Text2.Text = "You can select from an already made options of documentary like:" + vbCrLf + "-Letter: Will use the expression Dear... e.t.c" + vbCrLf + "-Web page: Start building the source code of an html web. After finishing click 'Save & Preview' to view it in your default explrorer" + vbCrLf + "Note: When saving and previewing the file will be saved in c:\mywebpage.htm"
If List1.Text = "Key Shortcuts" Then Text2.Text = "Ctrl + C - Copy" + vbCrLf + "Ctrl + X - Cut" + vbCrLf + "Ctrl + V - Paste" + vbCrLf + "Ctrl + A - Select All" + vbCrLf + "Ctrl + Z - Undo" + vbCrLf + "Ctrl + F - Find" + vbCrLf + "Ctrl + P - Pring (ALL)" + vbCrLf + "Alt + F4 - Exit"
If List1.Text = "Open With..." Then Text2.Text = "If you know how to handle another program better than Multi-Pad you can open your document with it. Also you can select another program (other than notepad or wordpad)" + vbCrLf + "Warning: Don't open text with MS Word or other program that require the selection of type of document to use. Notepad and Wordpad directly open new document when ran"
End Sub


