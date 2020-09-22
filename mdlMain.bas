Attribute VB_Name = "mdlMain"
'declarations
Const SWP_NOSIZE = 1
Const SWP_NOMOVE = 2
Const SPI_SETSCREENSAVEACTIVE = 17
Public Const SPI_SCREENSAVERRUNNING = 97

'Option Explicit

Public Preview_Mode As Boolean

Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOZORDER = &H4
Public Const SWP_SHOWWINDOW = &H40

Public Const HWND_TOP = 0

Public Const WS_CHILD = &H40000000
Public Const GWL_HWNDPARENT = (-8)
Public Const GWL_STYLE = (-16)

Private TaskBarhWnd As Long
Private IsTaskBarEnabled As Integer
Private TaskBarMenuHwnd As Integer

'API declarations
Declare Function PwdChangePassword Lib "mpr" Alias "PwdChangePasswordA" (ByVal lpcRegkeyname As String, ByVal hWnd As Long, ByVal uiReserved1 As Long, ByVal uiReserved2 As Long) As Long
Declare Function VerifyScreenSavePwd Lib "password.cpl" (ByVal hWnd As Long) As Boolean
Public Declare Function ShowCursor Lib "User32" (ByVal bShow As Long) As Long
Public Declare Function SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SystemParametersInfo Lib "User32" Alias "SystemParametersInfoA" (ByVal uAction As _
      Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Declare Function GetClientRect Lib "User32" (ByVal hWnd As Long, lpRect As Rect) As Long
Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function SetParent Lib "User32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function EnableWindow Lib "User32" (ByVal hWnd As Integer, ByVal aBOOL As Integer) As Integer
Private Declare Function IsWindowEnabled Lib "User32" (ByVal hWnd As Integer) As Integer
Private Declare Function GetMenu Lib "User32" (ByVal hWnd As Integer) As Integer

Public CheckingPassword As Boolean

Sub Main()

    'see if password protected
    Dim CResult As String, CResult1 As String, CResult2, CmdLine As String
    Dim CBack As Long
    Dim args As String
    Dim preview_hwnd As Long
    Dim preview_rect As Rect
    Dim window_style As Long

    ' Get the command line arguments.
    args = UCase$(Trim$(Command$))
    CResult = Left(Command$, 2)
    
    Preview_Mode = False

    If CResult = "/a" Then
        If GetWinVersion() = 0 Then
        'win 9x only
            CResult1 = Len(Command$) - 3
            CResult2 = Right(Command$, CResult1)
            CBack = PwdChangePassword("SCRSAVE", CResult2, 0, 0)
            End
        End If
    End If
    
    If CResult = "/p" Then
        'preview mode
        Preview_Mode = True
        ' Get the preview area hWnd.
            preview_hwnd = GetHwndFromCommand(args)

            ' Get the dimensions of the preview area.
            GetClientRect preview_hwnd, preview_rect

            Load frmSaver

            ' Set the caption for Windows 95.
            frmSaver.Caption = "Preview"

            ' Get the current window style.
            window_style = GetWindowLong(frmSaver.hWnd, GWL_STYLE)

            ' Add WS_CHILD to make this a child window.
            window_style = (window_style Or WS_CHILD)

            ' Set the window's new style.
            SetWindowLong frmSaver.hWnd, _
                GWL_STYLE, window_style

            ' Set the window's parent so it appears
            ' inside the preview area.
            SetParent frmSaver.hWnd, preview_hwnd

            ' Save the preview area's hWnd in
            ' the form's window structure.
            SetWindowLong frmSaver.hWnd, _
                GWL_HWNDPARENT, preview_hwnd

            ' Show the preview.
            SetWindowPos frmSaver.hWnd, _
                HWND_TOP, 0&, 0&, _
                preview_rect.Right, _
                preview_rect.Bottom, _
                SWP_NOZORDER Or SWP_NOACTIVATE Or _
                    SWP_SHOWWINDOW
            Exit Sub
    End If
    
    'check for command line arguments
    If Command$ <> "" Then
        'check for options mode
        If InStr(1, Command$, "/c", vbTextCompare) <> 0 Or InStr(1, Command$, "/C", vbTextCompare) <> 0 Then
                frmOptions.Show
        ElseIf InStr(1, Command$, "/s", vbTextCompare) <> 0 Or InStr(1, Command$, "/S", vbTextCompare) <> 0 Then
            Res = SystemParametersInfo(SPI_SETSCREENSAVEACTIVE, 0, ByVal 0&, 0)
            Load frmSaver
            SetWindowPos frmSaver.hWnd, -1, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
            ok = DoEvents()
            frmSaver.Show
        End If
            

    End If
    
End Sub
    
    
    

Sub Exit_Saver()
    'check for password
    Dim Result As Boolean
    CheckingPassword = True
    Res = ShowCursor(True) 'Turn the cursor back on
    
    If GetWinVersion() = 0 Then
        'win9x
        Result = VerifyScreenSavePwd(frmSaver.hWnd)
    Else
        'NT or 2000
        Result = True
    End If
    
        If Preview_Mode Then
            Res = ShowCursor(True)
            End
        End If
        If Result = False Then
            'no password given continue
            Res = ShowCursor(False) 'Turn the cursor back off
            CheckingPassword = False
        Else
            
            'reset screensaver
            Res = SystemParametersInfo(SPI_SETSCREENSAVEACTIVE, 1, ByVal 0&, 0)
            'enable taskbar and alt-tab
            EnableTaskBar
            FastTaskSwitching True
            Unload frmSaver
            'exit out
            End
        End If
    
End Sub

' Get the hWnd for the preview window from the
' command line arguments.
Private Function GetHwndFromCommand(ByVal args As String) As Long
Dim argslen As Integer
Dim i As Integer
Dim ch As String

    ' Take the rightmost numeric characters.
    args = Trim$(args)
    argslen = Len(args)
    For i = argslen To 1 Step -1
        ch = Mid$(args, i, 1)
        If ch < "0" Or ch > "9" Then Exit For
    Next i

    GetHwndFromCommand = CLng(Mid$(args, i + 1))
End Function

Sub DisableCtrlAltDelete(bDisabled As Boolean)
    Dim X As Long
    X = SystemParametersInfo(97, bDisabled, CStr(1), 0)
End Sub

Sub FastTaskSwitching(bEnabled As Boolean)

Dim X As Long, bDisabled As Long

    bDisabled = Not bEnabled

    X = SystemParametersInfo(97, bDisabled, CStr(1), 0)

End Sub

Public Sub DisableTaskBar()

Dim EWindow As Integer

    TaskBarhWnd = FindWindow("Shell_traywnd", "")

    If TaskBarhWnd <> 0 Then

        'check to see if window is enabled

        EWindow = IsWindowEnabled(TaskBarhWnd)

        If EWindow = 1 Then 'need to disable it

            IsTaskBarEnabled = EnableWindow(TaskBarhWnd, 0)

        End If

    End If

End Sub

 

Public Sub EnableTaskBar()

    If IsTaskBarEnabled = 0 Then

        IsTaskBarEnabled = EnableWindow(TaskBarhWnd, 1)

    End If

End Sub
