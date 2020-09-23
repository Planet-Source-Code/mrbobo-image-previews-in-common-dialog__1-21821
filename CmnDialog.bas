Attribute VB_Name = "CmnDialog"
Option Explicit
'This module is all standard OpenFile stuff
'with the exception of the Hook
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lparam As Any) As Long
Private Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lparam As String) As Long
Const OFN_ALLOWMULTISELECT As Long = &H200
Const OFN_CREATEPROMPT As Long = &H2000
Const OFN_ENABLEHOOK As Long = &H20
Const OFN_ENABLETEMPLATE As Long = &H40
Const OFN_ENABLETEMPLATEHANDLE As Long = &H80
Const OFN_EXPLORER As Long = &H80000
Const OFN_EXTENSIONDIFFERENT As Long = &H400
Const OFN_FILEMUSTEXIST As Long = &H1000
Const OFN_HIDEREADONLY As Long = &H4
Const OFN_LONGNAMES As Long = &H200000
Const OFN_NOCHANGEDIR As Long = &H8
Const OFN_NODEREFERENCELINKS As Long = &H100000
Const OFN_NOLONGNAMES As Long = &H40000
Const OFN_NONETWORKBUTTON As Long = &H20000
Const OFN_NOREADONLYRETURN As Long = &H8000&
Const OFN_NOTESTFILECREATE As Long = &H10000
Const OFN_NOVALIDATE As Long = &H100
Const OFN_OVERWRITEPROMPT As Long = &H2
Const OFN_PATHMUSTEXIST As Long = &H800
Const OFN_READONLY As Long = &H1
Const OFN_SHAREAWARE As Long = &H4000
Const OFN_SHAREFALLTHROUGH As Long = 2
Const OFN_SHAREWARN As Long = 0
Const OFN_SHARENOWARN As Long = 1
Const OFN_SHOWHELP As Long = &H10
Const OFS_MAXPATHNAME As Long = 260
Const OFN_SELECTED As Long = &H78
Const WM_INITDIALOG = &H110
Const SW_SHOWNORMAL = 1
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Const GW_NEXT = 2
Const GW_CHILD = 5
Const WM_GETTEXT = &HD
Const WM_GETTEXTLENGTH = &HE
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type
Dim OFN As OPENFILENAME
Dim cdlhwnd As Long
Public Function ShowOpen(hParent As Long, Optional mFilter As String, Optional mflags As Long, Optional mInitDir As String, Optional mTitle As String, Optional Pictures As Boolean) As String
    If mInitDir = "" Then mInitDir = "c:\"
    If mFilter = "" Then mFilter = "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    If mTitle = "" Then mTitle = App.Title
    OFN.lStructSize = Len(OFN)
    OFN.hwndOwner = hParent
    OFN.hInstance = App.hInstance
    OFN.lpstrFilter = mFilter
    OFN.lpstrFile = Space$(254)
    OFN.nMaxFile = 255
    OFN.lpstrFileTitle = Space$(254)
    OFN.nMaxFileTitle = 255
    OFN.lpstrInitialDir = mInitDir
    OFN.lpstrTitle = mTitle
    OFN.flags = mflags Or OFN_ENABLEHOOK Or OFN_EXPLORER
    'If you want the pictures you have top start hooking here
    If Pictures Then OFN.lpfnHook = DummyProc(AddressOf CdlgHook)
    If GetOpenFileName(OFN) Then
        ShowOpen = Trim$(OFN.lpstrFile)
    Else
        ShowOpen = ""
    End If

End Function
Private Function CdlgHook(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lparam As Long) As Long
'Here's the trick. Find the Commondialog window, stretch it
'to accomodate our frmPic, catch clicks in the Commondialog
'so we can load the image they click on.
Dim hwnda As Long, ClWind As String * 5, ClWind2 As String * 9, lngtextlen As Long
Dim temp As String, tmpFilename As String
Dim hwndParent As Long
Dim R As RECT
Dim NewCdlL As Long
Dim NewCdlT As Long
Dim mCdlW As Long
Dim mCdlH As Long
Dim scrWidth As Long
Dim scrHeight As Long
Select Case uMsg
    Case WM_INITDIALOG
        hwndParent = GetParent(hwnd)
        cdlhwnd = hwndParent
        If hwndParent <> 0 Then
            'Get the Commondialog dimensions
            Call GetWindowRect(hwndParent, R)
            mCdlW = R.Right - R.Left
            mCdlH = R.Bottom - R.Top
            scrWidth = Screen.Width \ Screen.TwipsPerPixelX
            scrHeight = Screen.Height \ Screen.TwipsPerPixelY
            NewCdlL = (scrWidth - mCdlW) \ 2
            NewCdlT = (scrHeight - mCdlH) \ 2
            'Stretch to fit our frmPic
            Call MoveWindow(hwndParent, NewCdlL, NewCdlT, mCdlW, mCdlH + 100, True)
            CdlgHook = 1
            frmPic.Show
            frmPic.Enabled = False
            'Place our form on the Commondialog
            Call MoveWindow(frmPic.hwnd, NewCdlL + 4, NewCdlT + 265, mCdlW - 18, mCdlH - 274, True)
        End If
    Case 78 'They selected something
        hwndParent = GetParent(hwnd)
        hwnda = GetWindow(hwndParent, GW_CHILD)
        'Find the "Filename" textbox
        Do While hwnda <> 0
            GetClassName hwnda, ClWind, 5
            If Left(ClWind, 4) = "Edit" Then
                tmpFilename = gettext(hwnda)
                Exit Do
            End If
            hwnda = GetWindow(hwnda, GW_NEXT)
        Loop
        hwnda = GetWindow(hwndParent, GW_CHILD)
        'Find the top combobox - this holds the path
        Do While hwnda <> 0
            GetClassName hwnda, ClWind2, 9
            If Left(ClWind2, 8) = "ComboBox" Then
                temp = gettext(hwnda)
                If Not FileExists(".\..\" + temp) Then
                'temp must have been the root directory
                'something like "MyDrive(c:)", well that wont
                'work. Scrub the lot and call it "." - we
                'check to see if it exists later anyway
                    temp = "."
                Else
                'Cool!! It's a file lets use it
                    temp = ".\..\" + temp
                End If
                    If tmpFilename <> "" Then
                    'Combine path and filename to make sure they exist
                    'and then if they do loadem up
                        If FileExists(temp + "\" + tmpFilename) Then
                            frmPic.LoadImage temp + "\" + tmpFilename
                        End If
                    End If
                Exit Do
            End If
            hwnda = GetWindow(hwnda, GW_NEXT)
        Loop
    Case 2 'They've finished - better hide our form
        Unload frmPic
    Case Else
End Select
End Function
'Used to steal the filename text from the Commondialog
'So we can load the image.
Private Function gettext(lngwindow As Long) As String
    Dim strbuffer As String, lngtextlen As Long
    Let lngtextlen& = SendMessage(lngwindow&, WM_GETTEXTLENGTH, 0&, 0&)
    Let strbuffer$ = String(lngtextlen&, 0&)
    Call SendMessageByString(lngwindow&, WM_GETTEXT, lngtextlen& + 1&, strbuffer$)
    Let gettext$ = strbuffer$
End Function
Public Sub MovePics()
'Keeps everything positioned with the correct Z order
Dim R As RECT
Dim mCdlW As Long
Dim mCdlH As Long
Call GetWindowRect(cdlhwnd, R)
mCdlW = R.Right - R.Left
mCdlH = R.Bottom - R.Top
Call MoveWindow(frmPic.hwnd, R.Left + 4, R.Top + 265, mCdlW - 18, mCdlH - 274, True)
Call SetWindowPos(cdlhwnd, frmPic.hwnd, R.Left, R.Top, mCdlW, mCdlH, 0)
End Sub

Private Function FileExists(ByVal Filename As String) As Integer
Dim temp$, MB_OK
    FileExists = True
On Error Resume Next
    temp$ = FileDateTime(Filename)
    Select Case Err
        Case 53, 76, 68
            FileExists = False
            Err = 0
        Case Else
            If Err <> 0 Then
                MsgBox "Error Number: " & Err & Chr$(10) & Chr$(13) & " " & Error, MB_OK, "Error"
                End
            End If
    End Select
End Function
'Used for hooking
Private Function DummyProc(ByVal dProc As Long) As Long
  DummyProc = dProc
End Function
