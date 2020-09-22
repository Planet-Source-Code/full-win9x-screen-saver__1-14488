VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Screen Saver Options"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog cmdColor 
      Left            =   360
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin VB.VScrollBar vsDelay 
      Height          =   375
      Left            =   3360
      Max             =   1
      Min             =   99
      TabIndex        =   5
      Top             =   720
      Value           =   99
      Width           =   255
   End
   Begin VB.TextBox txtDelay 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   2640
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   2
      Text            =   "1"
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label lblBackColor 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Background Color:"
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Time Delay (s):"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOK_Click()
    'save options
    SaveSetting App.Title, "Settings", "Delay", Me.txtDelay
    SaveSetting App.Title, "Settings", "BackColor", Me.lblBackColor.BackColor
    'exit
    End
End Sub

Private Sub Form_Load()
    Dim rgbColor As Long
    Dim Rcolor, Gcolor, Bcolor As Integer
    
    Me.Caption = "Options for " + App.Title
    TimeDelay = GetSetting(App.Title, "Settings", "Delay", 0)
    If TimeDelay = 0 Then
        SaveSetting App.Title, "Settings", "Delay", 2
        TimeDelay = 2
    End If
    Me.vsDelay = TimeDelay
    If GetSetting(App.Title, "Settings", "BackColor", "") = "" Then
        SaveSetting App.Title, "Settings", "BackColor", Me.BackColor
    End If
    Me.lblBackColor.BackColor = GetSetting(App.Title, "Settings", "BackColor", 0)
    'MsgBox GetSetting(App.Title, "Settings", "Delay", 0)
End Sub

Private Sub lblBackColor_Click()
    On Error GoTo NoChange
    cmdColor.Color = lblBackColor.BackColor
    cmdColor.ShowColor
    Me.lblBackColor.BackColor = cmdColor.Color
NoChange:
End Sub

Private Sub vsDelay_Change()
    Me.txtDelay = Me.vsDelay.Value
End Sub
