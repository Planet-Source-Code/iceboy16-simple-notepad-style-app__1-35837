Attribute VB_Name = "Module1"
Public Type FINDREPLACE
        lStructSize As Long         '   size of this struct 0x20
        hwndOwner As Long           '   handle to owner's window
        hInstance As Long           '   instance handle of.EXE that
                                    '   contains cust. dlg. template
        flags As Long               '   one or more of the FR_??
        lpstrFindWhat As Long       '   ptr. to search string
        lpstrReplaceWith As Long    '   ptr. to replace string
        wFindWhatLen As Integer     '   size of find buffer
        wReplaceWithLen As Integer  '   size of replace buffer
        lCustData As Long           '   data passed to hook fn.
        lpfnHook As Long            '   ptr. to hook fn. or NULL
        lpTemplateName As Long      '   custom template name
End Type

Public Declare Function FindText Lib "comdlg32.dll" Alias "FindTextA" (pFindreplace As FINDREPLACE) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetDlgItem Lib "user32" (ByVal hDlg As Long, ByVal nIDDlgItem As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function IsDlgButtonChecked Lib "user32" (ByVal hDlg As Long, ByVal nIDButton As Long) As Long
Public Declare Function GetDlgItemText Lib "user32" Alias "GetDlgItemTextA" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long

Public Const GWL_WNDPROC = (-4)
Public Const WM_LBUTTONDOWN = &H201

Public Const FR_NOMATCHCASE = &H800
Public Const FR_MATCHCASE = &H4
Public Const FR_NOUPDOWN = &H400
Public Const FR_UPDOWN = &H1
Public Const FR_NOWHOLEWORD = &H1000
Public Const FR_WHOLEWORD = &H2
Public Const EM_SETSEL = &HB1

Public Const MaxPatternLen = 50   ' Maximum Pattern Length

Global gOldDlgWndHandle As Long
Global frText As FINDREPLACE
Global gTxtSrc As String
Global gHDlg As Long
Global gHTxtWnd As Long

Function FindTextHookProc(ByVal hDlg As Long, ByVal uMsg As Long, _
   ByVal wParam As Long, ByVal lParam As Long) As Long

Dim strPtn As String    ' pattern string
Dim hTxtBox As Long     ' handle of the text box in dialog box
Dim ptnLen As Integer   ' actual length read by GetWindowString
Dim sp As Integer       ' start point of matching string
Dim ep As Integer       ' end point of matchiing string
Dim ret As Long         ' return value for SendMessage

strPtn = Space(MaxPatternLen)

    Select Case uMsg
        Case WM_LBUTTONDOWN
             ' Get the pattern string
             ptnLen = GetDlgItemText(gHDlg, &H480, strPtn, MaxPatternLen)
            
             ' Call default window procedure
             If gOldDlgWndHandle <> 0 Then
                 FindTextHookProc = CallWindowProc(gOldDlgWndHandle, _
                    hDlg, uMsg, wParam, lParam)
             End If
             
             ' Customize the winodw procedure
             If ptnLen <> 0 Then
                 strPtn = Left(strPtn, ptnLen)
                 SetFocus gHTxtWnd
                 
                 ' Get the MatchCase option
                 If IsDlgButtonChecked(gHDlg, &H411) = 0 Then
                     sp = InStr(LCase(gTxtSrc), LCase(strPtn))
                 Else
                     sp = InStr(gTxtSrc, strPtn)
                 End If
                 
                 sp = IIf(sp = 0, -1, sp - 1)
                 
                 If sp = -1 Then
                     Call MessageNoFound
                 End If
                 
                 ep = Len(strPtn)
                 ret = SendMessage(gHTxtWnd, EM_SETSEL, sp, sp + ep)
             End If
                
        Case Else
            ' Call the default window procedure
            If gOldDlgWndHandle <> 0 Then
               FindTextHookProc = CallWindowProc(gOldDlgWndHandle, _
                  hDlg, uMsg, wParam, lParam)
            End If
    End Select
End Function

Sub MessageNoFound()
MsgBox "No other matches found", vbInformation, "End of search"
End Sub


