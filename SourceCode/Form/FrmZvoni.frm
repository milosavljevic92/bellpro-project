VERSION 5.00
Begin VB.Form FrmZvoni 
   BorderStyle     =   0  'None
   ClientHeight    =   1815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2760
   Icon            =   "FrmZvoni.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1815
   ScaleWidth      =   2760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TmrUnloader 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3000
      Top             =   1320
   End
   Begin VB.Timer TmrUnload 
      Enabled         =   0   'False
      Left            =   3000
      Top             =   840
   End
   Begin VB.Timer TmrMove 
      Interval        =   1
      Left            =   3000
      Top             =   360
   End
   Begin VB.Image ImgClose1 
      Height          =   225
      Left            =   3000
      Picture         =   "FrmZvoni.frx":802D
      Top             =   120
      Width           =   225
   End
   Begin VB.Image ImgClose2 
      Height          =   225
      Left            =   3240
      Picture         =   "FrmZvoni.frx":833F
      Top             =   120
      Width           =   225
   End
   Begin VB.Image ImgClose3 
      Height          =   225
      Left            =   3480
      Picture         =   "FrmZvoni.frx":8651
      Top             =   120
      Width           =   225
   End
   Begin VB.Image Image1 
      Height          =   180
      Left            =   90
      Picture         =   "FrmZvoni.frx":8963
      Stretch         =   -1  'True
      Top             =   90
      Width           =   240
   End
   Begin VB.Label LblOptions 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   1400
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   405
      Width           =   1215
   End
   Begin VB.Label LblText 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Image ImgClose 
      Height          =   225
      Left            =   2400
      Top             =   90
      Width           =   225
   End
   Begin VB.Label LblMessage 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   885
      Left            =   130
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   590
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   885
      Left            =   130
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   600
      Width           =   2535
   End
   Begin VB.Image ImgMsnBG 
      Height          =   1740
      Left            =   0
      Picture         =   "FrmZvoni.frx":91BD
      Top             =   0
      Width           =   2715
   End
End
Attribute VB_Name = "FrmZvoni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Number As Integer
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Sub Form_Load()
    Me.Width = ImgMsnBG.Width
    Me.Height = ImgMsnBG.Height
    ImgClose.Picture = ImgClose1.Picture
    Me.Top = Screen.Height
    Me.Left = Screen.Width - Me.Width - 220
    ImgClose.Picture = ImgClose1.Picture
    LblOptions.FontUnderline = False
    TmrUnload.Enabled = True
End Sub
Public Sub VremeZatvaranja(Interval As String)
    TmrUnload.Interval = Interval * 1000
End Sub
Private Sub ImgClose_Click()
Unload Me
End Sub

Private Sub ImgClose_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
ImgClose.Picture = ImgClose3.Picture
End Sub

Private Sub ImgClose_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
ImgClose.Picture = ImgClose2.Picture
End Sub



Private Sub ImgMsnBG_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
ImgClose.Picture = ImgClose1.Picture
LblOptions.FontUnderline = False
End Sub

Private Sub LblMessage_Change()
  Label1 = LblMessage
End Sub



Private Sub LblOptions_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
LblOptions.FontUnderline = True
End Sub



Private Sub TmrMove_Timer()
If Me.Top <= Screen.Height - Me.Height Then
  TmrMove.Enabled = False
Else
  Me.Top = Me.Top - 50
End If
End Sub

Private Sub TmrUnload_Timer()
Unload Me
End Sub


