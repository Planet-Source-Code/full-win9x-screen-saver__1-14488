VERSION 5.00
Begin VB.Form frmSaver 
   BorderStyle     =   0  'None
   Caption         =   "Unusual Cars"
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8925
   Icon            =   "frmSaver.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   8925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   120
   End
   Begin VB.Image imgPicture 
      Height          =   4005
      Index           =   7
      Left            =   1320
      Picture         =   "frmSaver.frx":030A
      Top             =   1080
      Width           =   6000
   End
   Begin VB.Image imgPicture 
      Height          =   4005
      Index           =   6
      Left            =   1800
      Picture         =   "frmSaver.frx":67C7
      Top             =   1560
      Width           =   6000
   End
   Begin VB.Image imgPicture 
      Height          =   5940
      Index           =   5
      Left            =   1200
      Picture         =   "frmSaver.frx":CA56
      Top             =   840
      Width           =   7500
   End
   Begin VB.Image imgPicture 
      Height          =   4875
      Index           =   4
      Left            =   3840
      Picture         =   "frmSaver.frx":21EE3
      Top             =   1080
      Width           =   7500
   End
   Begin VB.Image imgPicture 
      Height          =   5160
      Index           =   3
      Left            =   3840
      Picture         =   "frmSaver.frx":30F59
      Top             =   720
      Width           =   8250
   End
   Begin VB.Image imgPicture 
      Height          =   6000
      Index           =   2
      Left            =   3840
      Picture         =   "frmSaver.frx":3B140
      Top             =   720
      Width           =   7620
   End
   Begin VB.Image imgPicture 
      Height          =   6000
      Index           =   1
      Left            =   3840
      Picture         =   "frmSaver.frx":45C21
      Top             =   720
      Width           =   8250
   End
   Begin VB.Image imgPicture 
      Height          =   5985
      Index           =   0
      Left            =   3840
      Picture         =   "frmSaver.frx":59FC3
      Top             =   720
      Width           =   8250
   End
End
Attribute VB_Name = "frmSaver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private TimePassed As Integer
Private CurrentPic As Integer
Private oldMouseX, oldMouseY As Integer
Private HaveDivided As Boolean

Private Sub Form_Click()
    If Not Preview_Mode Then
        Exit_Saver
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not Preview_Mode Then
        Exit_Saver
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If Not Preview_Mode Then
        Exit_Saver
    End If
End Sub

Private Sub ShowPic(PicNumber As Integer)
    For i = 0 To NumOfPics - 1
        Me.imgPicture(i).Visible = False
    Next i
        Me.imgPicture(PicNumber).Visible = True
End Sub

Private Sub CenterPics()
    If Not Preview_Mode Then
        'center all pictures on form
        For i = 0 To NumOfPics - 1
            If Me.imgPicture(i).Width > Me.Width Then
                Me.imgPicture(i).Left = 0
            Else
                Me.imgPicture(i).Left = (Me.Width / 2) - (Me.imgPicture(i).Width / 2)
            End If
            If Me.imgPicture(i).Height > Me.Height Then
                Me.imgPicture(i).Top = 0
            Else
                Me.imgPicture(i).Top = (Me.Height / 2) - (Me.imgPicture(i).Height / 2)
            End If
        Next i
    Else
        'scale to window
        Dim scaleIt As Double

        'scale pictures
        For i = 0 To Me.imgPicture.Count - 1
            Me.imgPicture(i).Top = 0
            Me.imgPicture(i).Left = 0
            
            If Me.imgPicture(i).Height > Me.imgPicture(i).Width Then
                Me.imgPicture(i).Stretch = True
                Me.imgPicture(i).Height = Me.Height
                scaleIt = Me.imgPicture(i).Height / Me.Height
                Me.imgPicture(i).Width = Int(scaleIt * Me.Width)
            Else
                Me.imgPicture(i).Stretch = True
                Me.imgPicture(i).Width = Me.Width
                scaleIt = Me.imgPicture(i).Width / Me.Width
                Me.imgPicture(i).Height = Int(scaleIt * Me.Height)
            End If
        Next i
    End If
End Sub


Private Sub Form_Load()
    'get all settings
    TimeDelay = GetSetting(App.Title, "Settings", "Delay", 0)
    If TimeDelay = 0 Then
        SaveSetting App.Title, "Settings", "Delay", 2
        TimeDelay = 2
    End If
    If GetSetting(App.Title, "Settings", "BackColor", "") = "" Then
        SaveSetting App.Title, "Settings", "BackColor", Me.BackColor
    End If
    Me.BackColor = GetSetting(App.Title, "Settings", "BackColor", 0)
    
    If Not Preview_Mode Then
        'hide cursor
        This = ShowCursor(False)
        'disable Ctrl-Alt and taskbar
        DisableTaskBar
        FastTaskSwitching False
    End If

    oldMouseX = 0
    oldMouseY = 0
    
   ' Call DisableCtrlAltDelete(True)
    
    NumOfPics = Me.imgPicture.Count
    TimePassed = 0
    'center all pictures
    CenterPics
    'show first
    ShowPic 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Call DisableCtrlAltDelete(False)
End Sub

Private Sub Form_Resize()
    CenterPics
End Sub

Private Sub Timer1_Timer()
    TimePassed = TimePassed + 1
    If TimePassed = TimeDelay Then
        'reset time passed
        TimePassed = 0
        'go to next picture
        CurrentPic = CurrentPic + 1
        'see if we need to reset
        If CurrentPic = NumOfPics Then
            CurrentPic = 0
        End If
        'change picture
        ShowPic CurrentPic
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    Exit_Saver
    If Not Preview_Mode Then
        If (oldMouseX = 0) And (oldMouseY = 0) Then
            oldMouseX = X
            oldMouseY = Y
            Exit Sub
        End If
    
        If (oldMouseX <> X) Or (oldMouseY <> Y) Then
            Exit_Saver
            Else
                oldMouseX = 0
                oldMouseX = 0
        End If
    End If
End Sub

