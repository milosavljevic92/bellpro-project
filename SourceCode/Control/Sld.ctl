VERSION 5.00
Begin VB.UserControl Slider 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   1155
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1635
   ScaleHeight     =   1155
   ScaleWidth      =   1635
   Begin VB.Image Handle 
      Height          =   270
      Left            =   840
      Picture         =   "Sld.ctx":0000
      Top             =   480
      Width           =   150
   End
   Begin VB.Image IMBar 
      Height          =   60
      Index           =   2
      Left            =   360
      Picture         =   "Sld.ctx":016F
      Top             =   360
      Width           =   30
   End
   Begin VB.Image IMBar 
      Height          =   60
      Index           =   1
      Left            =   1320
      Picture         =   "Sld.ctx":01D1
      Stretch         =   -1  'True
      Top             =   720
      Width           =   15
   End
   Begin VB.Image IMBar 
      Height          =   60
      Index           =   0
      Left            =   360
      Picture         =   "Sld.ctx":0223
      Top             =   720
      Width           =   45
   End
End
Attribute VB_Name = "Slider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim OldX As Long, myValue As Long, myMax As Long
Dim myEnabled As Boolean

Public Event Change(MyVal As Long, myMaxVal As Long)

Private Sub Handle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not DebugMode = True Then On Error Resume Next
    OldX = X
End Sub

Private Sub Handle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not DebugMode = True Then On Error Resume Next
    Dim K As Long
    If Button = 1 And myEnabled Then
        K = Handle.Left - OldX + X
        If K >= 0 And K <= UserControl.Width - Handle.Width Then
            Handle.Left = K
            myValue = Round(Handle.Left / (UserControl.Width - Handle.Width), 2) * myMax
            RaiseEvent Change(myValue, myMax)
        End If
    End If
End Sub

Private Sub UserControl_InitProperties()
    If Not DebugMode = True Then On Error Resume Next
    myEnabled = True
    myValue = 0
    myMax = 100
End Sub

Private Sub UserControl_Resize()
    If Not DebugMode = True Then On Error Resume Next
    size UserControl.Width, 285 'limiting the height of the bar
    IMBar(0).Move 0, (UserControl.Height - IMBar(0).Height) / 2
    IMBar(2).Move UserControl.Width - IMBar(2).Width, IMBar(0).Top
    IMBar(1).Move IMBar(0).Left + IMBar(0).Width, IMBar(0).Top, UserControl.Width - IMBar(0).Width - IMBar(2).Width
    Handle.Move ((myValue + 1) / (myMax + 1)) * (UserControl.Width - Handle.Width), (UserControl.Height - Handle.Height) / 2
End Sub

Public Property Let Enabled(ByVal nwEnabled As Boolean)
    myEnabled = nwEnabled
    UserControl_Resize
    PropertyChanged "Enabled"
End Property

Public Property Get Enabled() As Boolean
    Enabled = myEnabled
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    If Not DebugMode = True Then On Error Resume Next
    With PropBag
        myEnabled = .ReadProperty("Enabled", True)
        myValue = .ReadProperty("Value", 0)
        myMax = .ReadProperty("Max", 100)
    End With 'PROPBAG
    UserControl_Resize
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    If Not DebugMode = True Then On Error Resume Next
    With PropBag
        .WriteProperty "Enabled", myEnabled, True
        .WriteProperty "Value", myValue, 0
        .WriteProperty "Max", myMax, 100
    End With
End Sub

Public Property Get value() As Long
    If Not DebugMode = True Then On Error Resume Next
    value = myValue
End Property

Public Property Let value(ByVal nwVal As Long)
    If Not DebugMode = True Then On Error Resume Next
    myValue = nwVal
    UserControl_Resize
    PropertyChanged "Value"
End Property

Public Property Get Max() As Long
    If Not DebugMode = True Then On Error Resume Next
    Max = myMax
End Property

Public Property Let Max(ByVal nwMax As Long)
    If Not DebugMode = True Then On Error Resume Next
    myMax = nwMax
    UserControl_Resize
    PropertyChanged "Max"
End Property
