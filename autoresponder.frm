VERSION 5.00
Begin VB.Form frmRespond 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Responder"
   ClientHeight    =   1455
   ClientLeft      =   2730
   ClientTop       =   3900
   ClientWidth     =   6030
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "autoresponder.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1455
   ScaleWidth      =   6030
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Enable"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   5775
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   0
      Text            =   "I'm not here at the moment. So please leave a message."
      Top             =   600
      Width           =   5775
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5520
      TabIndex        =   4
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5760
      TabIndex        =   3
      Top             =   0
      Width           =   135
   End
   Begin VB.Shape Shape1 
      Height          =   1455
      Left            =   0
      Top             =   0
      Width           =   6015
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Type in what you want as your automated message."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   3975
   End
   Begin VB.Label lblTitleBar 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Doggie's Auto-Responder"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6015
   End
End
Attribute VB_Name = "frmRespond"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Autoresponder code - By Doggie
' Freely use this code
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2 ' these are to make the form move
Dim WithEvents respond As MsgrObject
Attribute respond.VB_VarHelpID = -1
Dim MsnApp As IMessengerApp

Private Sub Form_Load()
Set respond = New MsgrObject 'declaring the messenger object
 Set MsnApp = CreateObject("messenger.messengerapp") 'declaring the messenger object
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim lngReturnValue As Long 'for the form to move by dragging
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set respond = Nothing ' after use it will delete the declared method
End Sub

Private Sub Label1_Click()
Unload Me ' unload the form
End Sub

Private Sub Label2_Click()
Me.WindowState = vbMinimized 'minimize to taskbar
End Sub

Private Sub lblInfo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  End If
End Sub

Private Sub lblTitleBar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  End If
End Sub

Private Sub respond_OnLogoff()
frmstop.Show ' detects if u log off and will shutdown the program
End Sub

Private Sub respond_OnTextReceived(ByVal pIMSession As Messenger.IMsgrIMSession, ByVal pSourceUser As Messenger.IMsgrUser, ByVal bstrMsgHeader As String, ByVal bstrMsgText As String, pfEnableDefault As Boolean)
If Check1.Value = 1 Then
MsnApp.LaunchIMUI pSourceUser ' launches a im window to the person that trying to contact u
MsnApp.IMWindows.Item(0).SendText ("Auto Away Message:") & (Text1.Text) 'away message details
End If
End Sub

