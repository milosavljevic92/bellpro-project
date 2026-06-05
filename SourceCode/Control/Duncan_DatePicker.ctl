VERSION 5.00
Begin VB.UserControl Duncan_DatePicker 
   BackColor       =   &H00C0C0FF&
   ClientHeight    =   6570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7245
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   6570
   ScaleWidth      =   7245
   ToolboxBitmap   =   "Duncan_DatePicker.ctx":0000
   Begin VB.PictureBox picDrop 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   1680
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   1
      Top             =   120
      Width           =   495
      Begin VB.CommandButton btnDrop 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   8.25
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Timer TimerMonthTicker 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   5280
      Top             =   840
   End
   Begin VB.PictureBox picListBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   3840
      ScaleHeight     =   2025
      ScaleWidth      =   1710
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1740
      Begin VB.Label lblSelDate 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "January 2005"
         Height          =   240
         Index           =   6
         Left            =   0
         TabIndex        =   65
         Top             =   1440
         Width           =   1500
      End
      Begin VB.Label lblSelDate 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "January 2005"
         Height          =   240
         Index           =   5
         Left            =   0
         TabIndex        =   64
         Top             =   1200
         Width           =   1500
      End
      Begin VB.Label lblSelDate 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "January 2005"
         Height          =   240
         Index           =   4
         Left            =   0
         TabIndex        =   63
         Top             =   960
         Width           =   1500
      End
      Begin VB.Label lblSelDate 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "January 2005"
         Height          =   240
         Index           =   3
         Left            =   0
         TabIndex        =   62
         Top             =   720
         Width           =   1500
      End
      Begin VB.Label lblSelDate 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "January 2005"
         Height          =   240
         Index           =   2
         Left            =   0
         TabIndex        =   61
         Top             =   480
         Width           =   1500
      End
      Begin VB.Label lblSelDate 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "January 2005"
         Height          =   240
         Index           =   1
         Left            =   0
         TabIndex        =   60
         Top             =   240
         Width           =   1500
      End
      Begin VB.Label lblSelDate 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "January 2005"
         Height          =   240
         Index           =   0
         Left            =   0
         TabIndex        =   59
         Top             =   0
         Width           =   1500
      End
   End
   Begin VB.PictureBox picCanvas 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   600
      ScaleHeight     =   3495
      ScaleWidth      =   2775
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1080
      Visible         =   0   'False
      Width           =   2775
      Begin VB.CommandButton btnOK 
         Caption         =   "OK"
         Height          =   255
         Left            =   2040
         TabIndex        =   68
         ToolTipText     =   "Confirm your selection"
         Top             =   2880
         Width           =   600
      End
      Begin VB.CommandButton btnToday 
         Caption         =   "Today"
         Height          =   255
         Left            =   120
         TabIndex        =   67
         ToolTipText     =   "Show today"
         Top             =   2880
         Width           =   600
      End
      Begin VB.PictureBox picOK 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FF00FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   2040
         ScaleHeight     =   495
         ScaleWidth      =   600
         TabIndex        =   4
         ToolTipText     =   "Confirm your selection"
         Top             =   2280
         Width           =   600
      End
      Begin VB.PictureBox picToday 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FF00FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   600
         TabIndex        =   3
         ToolTipText     =   "Show today"
         Top             =   2280
         Width           =   600
      End
      Begin VB.PictureBox picMonthHeaderBackground 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   75
         ScaleHeight     =   285
         ScaleWidth      =   2595
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   60
         Width           =   2595
         Begin VB.Label lblDateTitle 
            Alignment       =   2  'Center
            BackColor       =   &H00FF8080&
            Caption         =   "January 2005"
            Height          =   240
            Left            =   622
            TabIndex        =   66
            Top             =   30
            Width           =   1500
         End
         Begin VB.Image imgRight 
            Height          =   195
            Left            =   2370
            Picture         =   "Duncan_DatePicker.ctx":0312
            Top             =   60
            Width           =   165
         End
         Begin VB.Image imgLeft 
            Height          =   195
            Left            =   60
            Picture         =   "Duncan_DatePicker.ctx":0361
            Top             =   60
            Width           =   165
         End
      End
      Begin VB.PictureBox picDaysHeaderBackground 
         Appearance      =   0  'Flat
         BackColor       =   &H00DC9670&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   60
         ScaleHeight     =   285
         ScaleWidth      =   2625
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   360
         Width           =   2625
         Begin VB.Label lblDOW 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            Caption         =   "Sat"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   6
            Left            =   2280
            TabIndex        =   11
            Top             =   30
            Width           =   240
         End
         Begin VB.Label lblDOW 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            Caption         =   "Fri"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   5
            Left            =   1950
            TabIndex        =   10
            Top             =   30
            Width           =   180
         End
         Begin VB.Label lblDOW 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            Caption         =   "Thu"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   4
            Left            =   1560
            TabIndex        =   9
            Top             =   30
            Width           =   270
         End
         Begin VB.Label lblDOW 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            Caption         =   "Wed"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   3
            Left            =   1170
            TabIndex        =   8
            Top             =   30
            Width           =   330
         End
         Begin VB.Label lblDOW 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            Caption         =   "Tue"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   2
            Left            =   840
            TabIndex        =   7
            Top             =   30
            Width           =   270
         End
         Begin VB.Label lblDOW 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            Caption         =   "Mon"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   1
            Left            =   450
            TabIndex        =   6
            Top             =   30
            Width           =   300
         End
         Begin VB.Label lblDOW 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            Caption         =   "Sun"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Top             =   30
            Width           =   270
         End
      End
      Begin VB.Shape ShapeSelection 
         BorderColor     =   &H000355BB&
         Height          =   255
         Left            =   0
         Top             =   0
         Width           =   345
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         Caption         =   "2"
         Height          =   495
         Index           =   9
         Left            =   840
         TabIndex        =   23
         Top             =   960
         Width           =   585
      End
      Begin VB.Label lblDateDescription 
         Alignment       =   2  'Center
         Caption         =   "30 January 2005"
         Height          =   255
         Left            =   720
         TabIndex        =   56
         ToolTipText     =   "The currently highlighted date. Press OK to select this date."
         Top             =   2280
         Width           =   1305
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         Caption         =   "33"
         Height          =   225
         Index           =   41
         Left            =   2400
         TabIndex        =   55
         Top             =   1800
         Width           =   330
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         Caption         =   "2"
         Height          =   255
         Index           =   40
         Left            =   1920
         TabIndex        =   54
         Top             =   1920
         Width           =   345
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         Caption         =   "2"
         Height          =   255
         Index           =   39
         Left            =   1560
         TabIndex        =   53
         Top             =   1920
         Width           =   345
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         Caption         =   "2"
         Height          =   255
         Index           =   38
         Left            =   1200
         TabIndex        =   52
         Top             =   1920
         Width           =   345
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         Caption         =   "2"
         Height          =   255
         Index           =   37
         Left            =   840
         TabIndex        =   51
         Top             =   1920
         Width           =   345
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         Caption         =   "2"
         Height          =   255
         Index           =   36
         Left            =   480
         TabIndex        =   50
         Top             =   1920
         Width           =   345
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         Caption         =   "2"
         Height          =   255
         Index           =   35
         Left            =   120
         TabIndex        =   49
         Top             =   1920
         Width           =   345
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FF00FF&
         Caption         =   "S"
         Height          =   230
         Index           =   34
         Left            =   2280
         TabIndex        =   48
         Top             =   1680
         Width           =   220
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         Caption         =   "2"
         Height          =   255
         Index           =   33
         Left            =   1920
         TabIndex        =   47
         Top             =   1680
         Width           =   345
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         Caption         =   "2"
         Height          =   255
         Index           =   32
         Left            =   1560
         TabIndex        =   46
         Top             =   1680
         Width           =   345
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         Caption         =   "2"
         Height          =   255
         Index           =   31
         Left            =   1200
         TabIndex        =   45
         Top             =   1680
         Width           =   345
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         Caption         =   "2"
         Height          =   255
         Index           =   30
         Left            =   840
         TabIndex        =   44
         Top             =   1680
         Width           =   345
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         Caption         =   "2"
         Height          =   255
         Index           =   29
         Left            =   480
         TabIndex        =   43
         Top             =   1680
         Width           =   345
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         Caption         =   "2"
         Height          =   255
         Index           =   28
         Left            =   120
         TabIndex        =   42
         Top             =   1680
         Width           =   345
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         Caption         =   "2"
         Height          =   255
         Index           =   27
         Left            =   2280
         TabIndex        =   41
         Top             =   1440
         Width           =   345
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         Caption         =   "2"
         Height          =   255
         Index           =   26
         Left            =   1920
         TabIndex        =   40
         Top             =   1440
         Width           =   345
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         Caption         =   "2"
         Height          =   255
         Index           =   25
         Left            =   1560
         TabIndex        =   39
         Top             =   1440
         Width           =   345
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         Caption         =   "2"
         Height          =   255
         Index           =   24
         Left            =   1200
         TabIndex        =   38
         Top             =   1440
         Width           =   345
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         Caption         =   "2"
         Height          =   255
         Index           =   23
         Left            =   840
         TabIndex        =   37
         Top             =   1440
         Width           =   345
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         Caption         =   "2"
         Height          =   255
         Index           =   22
         Left            =   480
         TabIndex        =   36
         Top             =   1440
         Width           =   345
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         Caption         =   "2"
         Height          =   255
         Index           =   21
         Left            =   120
         TabIndex        =   35
         Top             =   1440
         Width           =   345
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         Caption         =   "2"
         Height          =   255
         Index           =   20
         Left            =   2280
         TabIndex        =   34
         Top             =   1200
         Width           =   345
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         Caption         =   "2"
         Height          =   255
         Index           =   19
         Left            =   1920
         TabIndex        =   33
         Top             =   1200
         Width           =   345
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         Caption         =   "2"
         Height          =   255
         Index           =   18
         Left            =   1560
         TabIndex        =   32
         Top             =   1200
         Width           =   345
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         Caption         =   "2"
         Height          =   255
         Index           =   17
         Left            =   1200
         TabIndex        =   31
         Top             =   1200
         Width           =   345
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         Caption         =   "2"
         Height          =   255
         Index           =   16
         Left            =   840
         TabIndex        =   30
         Top             =   1200
         Width           =   345
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         Caption         =   "2"
         Height          =   255
         Index           =   15
         Left            =   480
         TabIndex        =   29
         Top             =   1200
         Width           =   345
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         Caption         =   "2"
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   28
         Top             =   1200
         Width           =   345
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         Caption         =   "2"
         Height          =   255
         Index           =   13
         Left            =   2280
         TabIndex        =   27
         Top             =   960
         Width           =   345
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         Caption         =   "2"
         Height          =   255
         Index           =   12
         Left            =   1920
         TabIndex        =   26
         Top             =   960
         Width           =   345
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         Caption         =   "2"
         Height          =   255
         Index           =   11
         Left            =   1560
         TabIndex        =   25
         Top             =   960
         Width           =   345
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         Caption         =   "2"
         Height          =   255
         Index           =   10
         Left            =   1200
         TabIndex        =   24
         Top             =   960
         Width           =   345
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         Caption         =   "2"
         Height          =   255
         Index           =   8
         Left            =   480
         TabIndex        =   22
         Top             =   960
         Width           =   345
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         Caption         =   "2"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Width           =   345
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         Caption         =   "2"
         Height          =   255
         Index           =   6
         Left            =   2280
         TabIndex        =   20
         Top             =   720
         Width           =   345
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         Caption         =   "2"
         Height          =   255
         Index           =   5
         Left            =   1920
         TabIndex        =   19
         Top             =   720
         Width           =   345
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         Caption         =   "2"
         Height          =   255
         Index           =   4
         Left            =   1560
         TabIndex        =   18
         Top             =   720
         Width           =   345
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         Caption         =   "2"
         Height          =   255
         Index           =   3
         Left            =   1320
         TabIndex        =   17
         Top             =   720
         Width           =   345
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         Caption         =   "2"
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   16
         Top             =   720
         Width           =   345
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Caption         =   "2"
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   15
         Top             =   840
         Width           =   345
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "2"
         Height          =   255
         Index           =   0
         Left            =   135
         TabIndex        =   13
         Top             =   720
         Width           =   345
      End
      Begin VB.Shape ShapeBorderLarge 
         BorderColor     =   &H00B99D7F&
         Height          =   2595
         Left            =   0
         Top             =   0
         Width           =   2745
      End
      Begin VB.Shape ShapeBorderSmall 
         BorderColor     =   &H00B99D7F&
         Height          =   1845
         Left            =   60
         Top             =   390
         Width           =   2625
      End
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Shape ShapeBorderTextbox 
      BorderColor     =   &H000000FF&
      Height          =   255
      Left            =   360
      Top             =   240
      Width           =   1275
   End
End
Attribute VB_Name = "Duncan_DatePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'What?
'A datepicker / calendar in one usercontrol
'Implements theme drawing of buttons and drop arrow.

'Why?
'Because all the datepickers are in huge OCX's and I
'like stand alone projects. I also wanted theme support and
'a consistant design to my other controls.

'How?
'The calendar is drawn on a picture box within the usercontrol.
'This is then set to have parent of desktop so that it will popup
'outside the bounds of any form it may be on.
'Similarly the month / year picker is created.
'Events are then processed and a date can be picked.

'Who?
'Thanks to Paul (programming god) Catton for his amazing subclassing work
'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=54117&lngWId=1
'Also http://www.vbaccelerator.com for insights and code samples

'When?
'Last Updated : June 2005

'todo ideas:
'dropshadow?
'min / max date range ?
'show week number?

'======================================================================================================================================================
'MY DECLARES FOR THIS CONTROL
'======================================================================================================================================================
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

'for making child window
Private Const GWL_EXSTYLE As Long = -20
Private Const WS_EX_TOOLWINDOW As Long = &H80&
Private Const WS_EX_APPWINDOW As Long = &H40000
Private Const SWP_SHOWWINDOW As Long = &H40
Private Const SWP_NOSIZE As Long = &H1
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Private Declare Function ClientToScreen Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SetFocus Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function SetCapture Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long

'My Vars
Private m_CanvasShown As Boolean        'Is the canvas showing?
Private m_ListBoxShown As Boolean       'Is the listbox showing?
Private m_DateHighlighted As Date       'the date shown as highlighted on the calendar
Private m_DateDisplaying As Date        'the date used to display the month shown in the calendar
Private m_DateSelected As Date          'the return value / starting point for when the control is shown
Private m_FirstDayOfWeek As DayOfWeek   'what it says
Private COL_SELECTEDDAYBACKGROUND As OLE_COLOR
Private m_ShortDayNames As Boolean
Private m_Moving As Boolean             'is window moving?
Private m_Active As Boolean             'is form active?
Private m_DescriptionFormat As String   'how lblDescription is to be formatted
Private m_UseThemes As Boolean          'do we draw theme buttons?
Private m_OKButtonStateId As ButtonState    'what state the button should be drawn in
Private m_TodayButtonStateId As ButtonState
Private m_DropButtonStateId As ButtonState
Private m_UseHandCursor As Boolean
Private m_ShowNonMonthDays As Boolean
Private m_Enabled As Boolean
Private Const DEF_DATEFORMAT As String = "d mmm yyyy"

Public Event DateChanged(ByVal FromDate As Date, ByVal ToDate As Date)

Private Enum ButtonState
    Normal = 1
    Hot = 2
    Pressed = 3
    disabled = 4
    Defaulted = 5
End Enum
Public Enum DayOfWeek
    Sunday = vbSunday
    Monday = vbMonday
    Tuesday = vbTuesday
    Wednesday = vbWednesday
    Thursday = vbThursday
    Friday = vbFriday
    Saturday = vbSaturday
End Enum

'Theme drawing - buttons etc
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function OpenThemeData Lib "uxtheme.dll" _
   (ByVal hwnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme.dll" _
   (ByVal hTheme As Long) As Long
Private Declare Function DrawThemeParentBackground Lib "uxtheme.dll" _
   (ByVal hwnd As Long, ByVal hdc As Long, prc As RECT) As Long
Private Declare Function DrawThemeBackground Lib "uxtheme.dll" _
   (ByVal hTheme As Long, ByVal lHDC As Long, _
    ByVal iPartId As Long, ByVal iStateId As Long, _
    pRect As RECT, pClipRect As RECT) As Long
Private Declare Function GetThemeBackgroundContentRect Lib "uxtheme.dll" _
   (ByVal hTheme As Long, ByVal hdc As Long, _
    ByVal iPartId As Long, ByVal iStateId As Long, _
    pBoundingRect As RECT, pContentRect As RECT) As Long
Private Declare Function DrawThemeText Lib "uxtheme.dll" _
   (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, _
    ByVal iStateId As Long, ByVal pszText As Long, _
    ByVal iCharCount As Long, ByVal dwTextFlag As Long, _
    ByVal dwTextFlags2 As Long, pRect As RECT) As Long
Private Declare Function GetCurrentThemeName Lib "uxtheme.dll" ( _
    ByVal pszThemeFileName As Long, _
    ByVal dwMaxNameChars As Long, _
    ByVal pszColorBuff As Long, _
    ByVal cchMaxColorChars As Long, _
    ByVal pszSizeBuff As Long, _
    ByVal cchMaxSizeChars As Long _
   ) As Long
Private Declare Function GetThemeBackgroundRegion Lib "uxtheme.dll" _
   (ByVal hTheme As Long, ByVal hdc As Long, _
   ByVal iPartId As Long, ByVal iStateId As Long, _
   pBoundingRect As RECT, pRegion As Long) As Long
 
Private Const DT_CENTER As Long = &H1
Private Const DT_VCENTER As Long = &H4
Private Const DT_SINGLELINE As Long = &H20
Private Const CENTERED As Long = DT_CENTER Or DT_VCENTER Or DT_SINGLELINE

Private Const THEME_BLUE = 1
Private Const THEME_OLIVE = 2
Private Const THEME_SILVER = 3

'hand pointer
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Const IDC_HAND = 32649&
Private Const IDC_ARROW = 32512&
Private m_hHandCursor As Long

'======================================================================================================================================================
'SUBCLASSING DECLARES
'======================================================================================================================================================
Private Enum eMsgWhen
  MSG_AFTER = 1                                                                         'Message calls back after the original (previous) WndProc
  MSG_BEFORE = 2                                                                        'Message calls back before the original (previous) WndProc
  MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE                                        'Message calls back before and after the original (previous) WndProc
End Enum
Private Type tSubData                                                                   'Subclass data type
  hwnd                               As Long                                            'Handle of the window being subclassed
  nAddrSub                           As Long                                            'The address of our new WndProc (allocated memory).
  nAddrOrig                          As Long                                            'The address of the pre-existing WndProc
  nMsgCntA                           As Long                                            'Msg after table entry count
  nMsgCntB                           As Long                                            'Msg before table entry count
  aMsgTblA()                         As Long                                            'Msg after table array
  aMsgTblB()                         As Long                                            'Msg Before table array
End Type
Private sc_aSubData()                As tSubData                                        'Subclass data array
Private Const ALL_MESSAGES           As Long = -1                                       'All messages added or deleted
Private Const GMEM_FIXED             As Long = 0                                        'Fixed memory GlobalAlloc flag
Private Const GWL_WNDPROC            As Long = -4                                       'Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04               As Long = 88                                       'Table B (before) address patch offset
Private Const PATCH_05               As Long = 93                                       'Table B (before) entry count patch offset
Private Const PATCH_08               As Long = 132                                      'Table A (after) address patch offset
Private Const PATCH_09               As Long = 137                                      'Table A (after) entry count patch offset
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Window Messages
Private Const WM_ACTIVATE = &H6
Private Const WM_NCACTIVATE = &H86
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_RBUTTONDOWN As Long = &H204
Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_ENTERSIZEMOVE As Long = &H231
Private Const WM_EXITSIZEMOVE As Long = &H232
Private Const WM_PAINT As Long = &HF&
Private Const WM_SIZING As Long = &H214
Private Const WM_MOUSELEAVE As Long = &H2A3
Private Const WM_SYSCOLORCHANGE As Long = &H15
Private Const WM_THEMECHANGED As Long = &H31A
Private Const WM_MOVE As Long = &H3
Private Const WM_SHOWWINDOW As Long = &H18
'not sure why this message is specific but it is sometimes sent
'to the control just before it is hidden. capture it so we dont leave the
'canvas exposed when control is hidden
Private Const WM_PRINT As Long = &H317

'Mouse tracking declares
Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Enum TRACKMOUSEEVENT_FLAGS
    TME_HOVER = &H1&
    TME_LEAVE = &H2&
    TME_QUERY = &H40000000
    TME_CANCEL = &H80000000
End Enum
Private Type TRACKMOUSEEVENT_STRUCT
    cbSize                              As Long
    dwFlags                             As TRACKMOUSEEVENT_FLAGS
    hwndTrack                           As Long
    dwHoverTime                         As Long
End Type



'Subclass handler
Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lng_hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)
Attribute zSubclass_Proc.VB_MemberFlags = "40"
 On Error Resume Next
'THIS MUST BE THE FIRST PUBLIC ROUTINE IN THIS FILE.
'That includes public properties also
'Parameters:
  'bBefore  - Indicates whether the the message is being processed before or after the default handler - only really needed if a message is set to callback both before & after.
  'bHandled - Set this variable to True in a 'before' callback to prevent the message being subsequently processed by the default handler... and if set, an 'after' callback
  'lReturn  - Set this variable as per your intentions and requirements, see the MSDN documentation for each individual message value.
  'hWnd     - The window handle
  'uMsg     - The message number
  'wParam   - Message related data
  'lParam   - Message related data
    Select Case lng_hWnd
    Case UserControl.Parent.hwnd
        'FORM
        'Debug.Print WMbyName(uMsg)
        'All of this is to close the canvas if the user moves away from it in anyway
        Select Case uMsg
            Case WM_THEMECHANGED, WM_SYSCOLORCHANGE
                'theme has changed
                InitialiseColours
                InitialiseThemes
                InitialiseCombo
                HighlightSelection
            Case WM_SIZING
                HideCalendar
            Case WM_ENTERSIZEMOVE
                m_Moving = True
            Case WM_EXITSIZEMOVE
                m_Moving = False
                RepositionCalendar
            Case WM_PAINT
                'If m_CanvasShown And (Not m_Moving) And m_Active Then   'make sure it stays on top
                '    RepositionCalendar
                'End If
            Case WM_MOVE
                HideListBox
                RepositionCalendar
            Case WM_LBUTTONDOWN, WM_RBUTTONDOWN
                'Debug.Print "button down - destroying calendar"
                HideCalendar
            Case WM_ACTIVATE, WM_NCACTIVATE
                If wParam Then  '----------------------------------- Activated
                    'Debug.Print "activated " & wParam & " " & lParam & " " & Now
                    m_Active = True
                Else            '----------------------------------- Deactivated
                    'Debug.Print "deactivated " & wParam & " " & lParam & " " & Now
                    m_Active = False
                End If
            Case Else
                'Debug.Print WMbyName(uMsg)
        End Select
    Case UserControl.hwnd
        Select Case uMsg
        Case WM_SHOWWINDOW
            If wParam = 0 Then
                If m_CanvasShown Then
                    'usercontrol is being hidden (made invisible)
                    'we dont want to leave the canvas exposed
                    HideCalendar
                End If
            End If
        Case WM_PRINT
            If m_CanvasShown Then
                'usercontrol is being hidden (made invisible)
                'we dont want to leave the canvas exposed
                HideCalendar
            End If
        Case Else
            'Debug.Print WMbyName(uMsg) & " " & wParam & "x" & lParam & " " & IsWindowVisible(UserControl.hwnd)
        End Select
    Case Else
        'all of the buttons
        If uMsg = WM_MOUSELEAVE Then
            If Enabled Then
                Select Case lng_hWnd
                Case UserControl.picOK.hwnd
                    m_OKButtonStateId = Normal
                    RefreshButtons
                Case UserControl.picToday.hwnd
                    m_TodayButtonStateId = Normal
                    RefreshButtons
                Case UserControl.picDrop.hwnd
                    m_DropButtonStateId = Normal
                    DrawDrop
                End Select
            End If
        End If
    End Select
End Sub
'======================================================================================================================================================
'Functions
'======================================================================================================================================================

'-----------------
'PUBLIC PROPERTIES
'-----------------
Public Property Get DateSelected() As Date
Attribute DateSelected.VB_Description = "The date currently selected by the control."
Attribute DateSelected.VB_UserMemId = 0
    DateSelected = m_DateSelected
End Property
Public Property Let DateSelected(dVal As Date)
    If dVal <> m_DateSelected Then
        If dVal = "00:00:00" Then
            dVal = Date
        End If
        If DateDiff("d", m_DateSelected, dVal) <> 0 Then
            RaiseEvent DateChanged(m_DateSelected, dVal)
            m_DateSelected = dVal
        End If
        DateHighlighted = dVal
        DateDisplaying = dVal
        SyncCalendarDisplay
        Text1.Text = DateSelectedFormatted
        PropertyChanged "DateSelected"
    End If
End Property

Public Property Get FirstDayOfWeek() As DayOfWeek
Attribute FirstDayOfWeek.VB_Description = "Which day do you want to be the first in the display?"
    FirstDayOfWeek = m_FirstDayOfWeek
End Property
Public Property Let FirstDayOfWeek(eVal As DayOfWeek)
    If eVal <> m_FirstDayOfWeek Then
        If eVal < 1 Or eVal > 7 Then
            eVal = eVal Mod 7
        End If
        m_FirstDayOfWeek = eVal
        PropertyChanged "FirstDayOfWeek"
        InitialiseCanvas
        PopulateDays
        HighlightSelection
        RefreshButtons
    End If
End Property
Private Property Get DateHighlighted() As Date
    DateHighlighted = m_DateHighlighted
End Property
Private Property Let DateHighlighted(dVal As Date)
    If dVal <> m_DateHighlighted Then
        m_DateHighlighted = dVal
        DateDisplaying = m_DateHighlighted
        SyncCalendarDisplay
    End If
End Property
Private Property Get DateDisplaying() As Date
    If m_DateDisplaying = "00:00:00" Then m_DateDisplaying = Date
    DateDisplaying = m_DateDisplaying
End Property
Private Property Let DateDisplaying(dVal As Date)
    If dVal <> m_DateDisplaying Then
        m_DateDisplaying = dVal
        SyncCalendarDisplay
    End If
End Property
Public Property Get ShortDayNames() As Boolean
Attribute ShortDayNames.VB_Description = "False = Wed\r\nTrue = W"
    ShortDayNames = m_ShortDayNames
End Property
Public Property Let ShortDayNames(bVal As Boolean)
    If bVal <> m_ShortDayNames Then
        m_ShortDayNames = bVal
        PropertyChanged "ShortDayNames"
        ShowDayNames
    End If
End Property
Public Property Get DescriptionFormat() As String
Attribute DescriptionFormat.VB_Description = "How you would like the date displayed. Uses standard VB Format function."
    If Len(m_DescriptionFormat) = 0 Then m_DescriptionFormat = DEF_DATEFORMAT
    DescriptionFormat = m_DescriptionFormat
End Property
Public Property Let DescriptionFormat(sVal As String)
    If sVal <> m_DescriptionFormat Then
        m_DescriptionFormat = sVal
        PropertyChanged "DescriptionFormat"
        ShowDateDescription
        Text1.Text = DateSelectedFormatted
    End If
End Property
Public Property Get UseHandCursor() As Boolean
Attribute UseHandCursor.VB_Description = "Change the mouse pointer to a hand when it is over something that can be clicked?"
    UseHandCursor = m_UseHandCursor
End Property
Public Property Let UseHandCursor(bVal As Boolean)
    If bVal <> m_UseHandCursor Then
        m_UseHandCursor = bVal
        PropertyChanged "UseHandCursor"
    End If
End Property
Public Property Get ShowNonMonthDays() As Boolean
Attribute ShowNonMonthDays.VB_Description = "Include in the display days that arent in the display month?"
    ShowNonMonthDays = m_ShowNonMonthDays
End Property
Public Property Let ShowNonMonthDays(bVal As Boolean)
    If bVal <> m_ShowNonMonthDays Then
        m_ShowNonMonthDays = bVal
        PropertyChanged "ShowNonMonthDays"
        PopulateDays
        HighlightSelection  'selection might have been one of the days we removed/added
    End If
End Property
Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property
Public Property Let Enabled(bVal As Boolean)
    If bVal <> m_Enabled Then
        m_Enabled = bVal
        PropertyChanged "Enabled"
        If m_Enabled Then
            m_DropButtonStateId = Normal
            ShapeBorderTextbox.BorderColor = ShapeBorderLarge.BorderColor
        Else
            'disabeling the control
            m_DropButtonStateId = disabled
            If m_CanvasShown Then
                HideCalendar
            End If
            ShapeBorderTextbox.BorderColor = vbButtonFace
        End If
        Text1.Enabled = bVal
        DrawDrop
    End If
End Property
Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property
Public Property Set Font(ByVal fVal As Font)
    Dim I As Long
    
    Set UserControl.Font = fVal
    Set Text1.Font = fVal
    Set picDrop.Font = fVal
    Set picToday.Font = fVal
    Set btnToday.Font = fVal
    Set picOK.Font = fVal
    Set btnOK.Font = fVal
    Set lblDateTitle.Font = fVal
    Set lblDateDescription.Font = fVal
    For I = 0 To 6
        Set lblDOW(I).Font = fVal
        Set lblSelDate(I).Font = fVal
    Next
    For I = 0 To 41
        Set lblDay(I).Font = fVal
    Next
    
    PropertyChanged "Font"
    RefreshButtons
End Property


'--------
'CONTROLS
'--------
Private Sub lblDay_Click(Index As Integer)
    If Len(lblDay(Index).Tag) > 0 Then
        DateHighlighted = lblDay(Index).Tag
        HighlightSelection
        ShowDateDescription
    End If
End Sub

Private Sub lblDay_DblClick(Index As Integer)
    picOK_Click
End Sub

Private Sub lblDay_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Len(lblDay(Index).Caption) > 0 Then
        ShowHandCursor
    End If
End Sub

Private Sub lblDay_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Len(lblDay(Index).Caption) > 0 Then
        ShowHandCursor
    End If
End Sub

Private Sub btnDrop_LostFocus()
    CloseIfFocusLost
End Sub

Private Sub btnOK_LostFocus()
    CloseIfFocusLost
End Sub

Private Sub btnToday_LostFocus()
    CloseIfFocusLost
End Sub

Private Sub btnOK_Click()
    picOK_Click
End Sub

Private Sub btnToday_Click()
    picToday_Click
End Sub

Private Sub picOK_Click()
    DateSelected = DateHighlighted
    HideCalendar
    SetFocus Text1.hwnd
End Sub

Private Sub picToday_Click()
    DateDisplaying = Date
End Sub

Private Sub picDrop_Click()
    If Enabled Then
        If m_CanvasShown Then
            HideCalendar
            SetFocus Text1.hwnd
        Else
            ShowCalendar
        End If
    End If
End Sub


Private Sub btnDrop_Click()
    picDrop_Click
End Sub

Private Sub picOK_GotFocus()
    RefreshButtons
End Sub

Private Sub picOK_LostFocus()
    RefreshButtons
    CloseIfFocusLost
End Sub

Private Sub picOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
    ShowHandCursor
    
    If Button = 1 Then
        If m_OKButtonStateId <> Pressed Then
            m_OKButtonStateId = Pressed
            RefreshButtons
        End If
    Else
        If m_OKButtonStateId <> Hot Then
            m_OKButtonStateId = Hot
            RefreshButtons
        End If
    End If
    TrackMouseLeave picOK.hwnd
End Sub

Private Sub picToday_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShowHandCursor
    
    If Button = 1 Then
        If m_TodayButtonStateId <> Pressed Then
            m_TodayButtonStateId = Pressed
            RefreshButtons
        End If
    Else
        If m_TodayButtonStateId <> Hot Then
            m_TodayButtonStateId = Hot
            RefreshButtons
        End If
    End If
    TrackMouseLeave picToday.hwnd
End Sub
Private Sub picDrop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Enabled Then
        If Button = 1 Then
            If m_DropButtonStateId <> Pressed Then
                m_DropButtonStateId = Pressed
                DrawDrop
            End If
        Else
            If m_DropButtonStateId <> Hot Then
                m_DropButtonStateId = Hot
                DrawDrop
            End If
        End If
        TrackMouseLeave picDrop.hwnd
    End If
End Sub

Private Sub picToday_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picToday_MouseMove Button, Shift, X, Y
End Sub
Private Sub picOK_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picOK_MouseMove Button, Shift, X, Y
End Sub

Private Sub picDrop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Enabled Then
        picDrop_MouseMove Button, Shift, X, Y
    End If
End Sub
Private Sub picDrop_GotFocus()
    If LButtonIsDown Then   'we are clicking the button
        'do nothing it will
    Else
        'we got focus somehow
        'we dont want it unless the canvas is open
        If Not m_CanvasShown Then
            'move on
            SendKeys "{TAB}"
        End If
    End If
End Sub
Private Function LButtonIsDown() As Boolean
    'lets you know if the left mouse button is down
    Dim retval As Long
    retval = GetKeyState(vbKeyLButton)  'returns a negative value while the button is being depressed
    If retval < False Then
        LButtonIsDown = True
    End If
End Function
Private Sub picDrop_KeyDown(KeyCode As Integer, Shift As Integer)
    If Enabled Then
        If m_CanvasShown Then
            Select Case KeyCode
            Case 13, 32 'enter,space
                picOK_Click
            Case 27 'esc
                HideCalendar
                SetFocus Text1.hwnd
            Case 38 'up
                'move up one row
                DateHighlighted = DateAdd("d", -7, DateHighlighted)
            Case 40 'down
                'move down one row
                DateHighlighted = DateAdd("d", 7, DateHighlighted)
            Case 37 'left
                'move left one cell
                DateHighlighted = DateAdd("d", -1, DateHighlighted)
            Case 39 'right
                'move right one cell
                DateHighlighted = DateAdd("d", 1, DateHighlighted)
            Case 34 'pg down
                'move forward one month
                DateHighlighted = DateAdd("m", 1, DateHighlighted)
            Case 33 'pg up
                'move backward one month
                DateHighlighted = DateAdd("m", -1, DateHighlighted)
            Case 36 'home
                DateHighlighted = Date
                picToday_Click
            Case Else
                'Debug.Print KeyCode
            End Select
        Else
            ShowCalendar
        End If
    End If
End Sub

Private Sub btnOK_KeyDown(KeyCode As Integer, Shift As Integer)
    picOK_KeyDown KeyCode, Shift
End Sub

Private Sub CloseIfFocusLost()
    If Not FocusIsWithThisControl Then
        HideCalendar
    End If
End Sub

Private Sub picDrop_LostFocus()
    CloseIfFocusLost
End Sub

Private Sub picListBox_LostFocus()
    CloseIfFocusLost
End Sub

Private Sub picOK_KeyDown(KeyCode As Integer, Shift As Integer)
    If Enabled Then
        If m_CanvasShown Then
            Select Case KeyCode
            Case 13, 32 'enter,space
                picOK_Click
            Case 27 'esc
                HideCalendar
            Case 37, 38 'left,up
                'move back to today button
                SendKeys "+{TAB}"
            Case 39, 40 'right,down
                'move on to next control
                SendKeys "{TAB}"
            Case Else
                'Debug.Print KeyCode
            End Select
        End If
    End If
End Sub

Private Sub picToday_GotFocus()
    RefreshButtons
End Sub

Private Sub btnToday_KeyDown(KeyCode As Integer, Shift As Integer)
    'forward info
    picToday_KeyDown KeyCode, Shift
End Sub

Private Sub picToday_KeyDown(KeyCode As Integer, Shift As Integer)
    If Enabled Then
        If m_CanvasShown Then
            Select Case KeyCode
            Case 13, 32 'enter,space
                picToday_Click
            Case 27 'esc
                HideCalendar
            Case 37, 38 'left,up
                'move back to calendar
                SendKeys "+{TAB}"
            Case 39, 40 'right,down
                'move on to ok button
                SendKeys "{TAB}"
            Case Else
                'Debug.Print KeyCode
            End Select
        End If
    End If
End Sub

Private Sub picToday_LostFocus()
    RefreshButtons
    CloseIfFocusLost
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If Enabled Then
        Select Case KeyCode
        Case 13 'enter
            If IsDate(Text1.Text) Then
                If m_CanvasShown Then
                    If DateDiff("d", Text1, DateHighlighted) = 0 Then
                        'they pressed enter on a day that they already had entered
                        'close the canvas
                        picOK_Click
                    Else
                        DateHighlighted = Text1
                    End If
                Else
                    SetTextToDateSelected
                End If
            Else
                SetTextToDateSelected
            End If
        Case 38 'up
            'hide canvas if needed
            If m_CanvasShown Then
                HideCalendar
            End If
            'eat it
            KeyCode = 0
        Case 40 'down
            'show canvas if needed
            If m_CanvasShown Then
                'shift focus to the canvas
                SetFocus picDrop.hwnd
            Else
                'open the calendar
                ShowCalendar
            End If
            'eat it
            KeyCode = 0
        Case Else
            'Debug.Print KeyCode
        End Select
    End If
End Sub
Private Sub SetTextToDateSelected()
    'validate the text
    If IsDate(Text1.Text) Then
        DateSelected = Text1.Text
    Else
        DateSelected = DateHighlighted
    End If
    'then redisplay it
    Text1 = DateSelectedFormatted
  
End Sub

Private Sub Text1_LostFocus()
    CloseIfFocusLost
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
    SetTextToDateSelected
End Sub
Private Sub TimerMonthTicker_Timer()
    'causes months to tick through the listbox
    ShuffelListItems
End Sub

Private Sub ShuffelListItems()
    If TimerMonthTicker.Tag Then
        PopulateMonths lblSelDate(2).Tag
    Else
        PopulateMonths lblSelDate(4).Tag
    End If
End Sub


Private Sub imgLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DateDisplaying = DateAdd("m", -1, DateDisplaying)
    imgLeft_MouseMove Button, Shift, X, Y
End Sub

Private Sub imgLeft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShowHandCursor
End Sub

Private Sub imgRight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DateDisplaying = DateAdd("m", 1, DateDisplaying)
    imgRight_MouseMove Button, Shift, X, Y
End Sub

Private Sub imgRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShowHandCursor
End Sub

Private Sub lblDateDescription_Click()
    'take us to the page that shows this date
    DateDisplaying = DateHighlighted
End Sub

Private Sub lblDateDescription_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblDateDescription_MouseMove Button, Shift, X, Y
End Sub

Private Sub lblDateDescription_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShowHandCursor
End Sub

Private Sub picMonthHeaderBackground_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'compensates in case they missed the arrow buttons
    picMonthHeaderBackground_MouseMove Button, Shift, X, Y
    
    If X < lblDateTitle.Left Then
        'they tried to click left
        imgLeft_MouseDown Button, Shift, X, Y
    Else
        If X > (lblDateTitle.Left + lblDateTitle.Width) Then
            'they tried to click right
            imgRight_MouseDown Button, Shift, X, Y
        End If
    End If
End Sub
Private Sub picMonthHeaderBackground_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X < (imgLeft.Left + imgLeft.Width + (10 * Screen.TwipsPerPixelX)) _
    Or X > (imgRight.Left - (10 * Screen.TwipsPerPixelX)) Then
        ShowHandCursor
    End If
End Sub

'--------
'LIST BOX
'--------
Private Sub lblDateTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShowHandCursor
    PopulateMonths DateDisplaying
    ShowListBox
End Sub

Private Sub lblDateTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'this is where we change the selection of the month picker
    Dim R As RECT
    Dim P As POINTAPI
    Dim H As Long
    Dim Index As Long
    Dim I As Long
    
    'show hand cursor
    ShowHandCursor
    
    'is the mouse over the listbox?
    GetWindowRect picListBox.hwnd, R
    GetCursorPos P
    
    If PtInRect(R, P.X, P.Y) Then
        'mouse is in the listbox
        TimerMonthTicker.Enabled = False
        'highlight whichever lable it is over
        H = lblSelDate(0).Height / Screen.TwipsPerPixelY
        Index = ((P.Y - R.Top) / H) - 0.5
        For I = 0 To 6
            If I = Index Then
                lblSelDate(I).BackColor = vbBlack
                lblSelDate(I).ForeColor = vbWhite
                picListBox.Tag = Index
            Else
                lblSelDate(I).BackColor = picListBox.BackColor
                lblSelDate(I).ForeColor = vbBlack
            End If
        Next
    Else
        'so mouse is moving outside the box somewhere
        'if it is left or right of the control then ignore it
        If P.X < R.Left Or P.X > R.Right Then
            'ignore
            TimerMonthTicker.Enabled = False
        Else
            'they have moved the mouse beyond the boundry of the
            'list box at the top or the bottom of the control
            'start the month ticker
            'which will cause months/years to cycle past
            If P.Y < R.Top Then
                TimerMonthTicker.Tag = True
            Else
                TimerMonthTicker.Tag = False
            End If
            'ShuffelListItems
            TimerMonthTicker.Enabled = True
        End If
    End If
End Sub

Private Sub lblDateTitle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Len(picListBox.Tag) > 0 Then
        DateDisplaying = lblSelDate(picListBox.Tag).Tag
    End If
    HideListBox
End Sub


'--------------
'PICTURE CANVAS
'--------------
Private Sub picCanvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X < 0 Or Y < 0 Or X > picCanvas.Width Or Y > picCanvas.Height Then
        HideCalendar
    Else
        Debug.Print "inside"
    End If

End Sub


'-----------------
'PRIVATE FUNCTIONS
'-----------------
Private Sub InitialiseCombo()
    Dim X As Long
    'make sure shape is over the textbox
    ShapeBorderTextbox.ZOrder
    
    'make sure the drop box is in position
    X = ShapeBorderTextbox.Height - (4 * Screen.TwipsPerPixelY)
    If picDrop.Height <> X Then picDrop.Height = X
    If picDrop.Width <> X Then picDrop.Width = X
    X = ShapeBorderTextbox.Width - (picDrop.Width + (2 * Screen.TwipsPerPixelY))
    If picDrop.Left <> X Then picDrop.Left = X
    X = (ShapeBorderTextbox.Height - picDrop.Height) / 2
    If picDrop.Top <> X Then picDrop.Top = X
    
    'show which box?
    If m_UseThemes Then
        btnDrop.Visible = False
        DrawDrop
    Else
        btnDrop.Move 0, 0, picDrop.Width, picDrop.Height
        btnDrop.Visible = True
    End If
    
End Sub

Private Sub InitialiseCanvas()
    'position controls properly
    DynamicallyPlaceControls
    
    'the month / year title
    lblDateTitle.BackColor = picMonthHeaderBackground.BackColor
    lblDateDescription.BackColor = picCanvas.BackColor
    
    'the weekday names
    ShowDayNames
    
    'colour the days
    HighlightSelection
    
    'hide the control until needed
    With picCanvas
         .BorderStyle = 0
         .Visible = False
         .Width = ShapeBorderLarge.Width
         .Height = ShapeBorderLarge.Height
    End With
    
    picOK.BackColor = picCanvas.BackColor
    picToday.BackColor = picCanvas.BackColor
End Sub

Private Sub ShowDayNames()
    Dim I As Long
    Dim WDN As String

    For I = 0 To 6
        WDN = WeekdayName(I + 1, True, FirstDayOfWeek)
        If ShortDayNames Then
            lblDOW(I).Caption = Mid(WDN, 1, 1)
        Else
            lblDOW(I).Caption = WDN
        End If
    Next

End Sub
Private Sub ShowDateTitle()
    'shows the date in the window
    lblDateTitle = format(DateDisplaying, "mmmm yyyy")
End Sub
Private Sub ShowDateDescription()
    'shows the date in the window
    If Len(DescriptionFormat) > 0 Then
        lblDateDescription = format(DateHighlighted, m_DescriptionFormat)
    Else
        lblDateDescription = FormatDateTime(DateHighlighted, vbGeneralDate)
    End If
End Sub

Private Function DateSelectedFormatted() As String
    'returns Date Selected in the format you supply
    If Len(DescriptionFormat) > 0 Then
        DateSelectedFormatted = format(DateSelected, m_DescriptionFormat)
    Else
        DateSelectedFormatted = FormatDateTime(DateSelected, vbGeneralDate)
    End If
End Function

Private Sub PopulateDays(Optional lMonth As Long, Optional lYear As Long)
    'pass it a month and a year and it will populate the
    'day values into the labels
    Dim DOM As Long         'Day Of Month that the 1st falls on
    Dim D As Date
    Dim I As Long
    
    If lMonth = 0 Then lMonth = Month(DateDisplaying)
    If lYear = 0 Then lYear = Year(DateDisplaying)
    
    If lMonth > 0 And lMonth < 13 Then
        'first find out what day is the first day of the month
        D = "1 " & MonthName(lMonth) & " " & lYear
        DOM = Weekday(D, FirstDayOfWeek) - 1
        
        'blank out the days at the start of the month
        For I = 0 To DOM - 1
            If ShowNonMonthDays Then
                lblDay(I).Tag = DateAdd("d", I - DOM, D)
                lblDay(I).Caption = Day(lblDay(I).Tag)
            Else
                lblDay(I).Caption = ""
                lblDay(I).Tag = ""
            End If
            lblDay(I).ToolTipText = lblDay(I).Tag
            lblDay(I).ForeColor = vbGrayText
        Next
        
        'put in the days of the month
        Do While Month(D) = lMonth
            lblDay(I).Caption = Day(D)
            lblDay(I).Tag = D
            lblDay(I).ToolTipText = D
            lblDay(I).ForeColor = vbBlack
            D = DateAdd("d", 1, D)
            I = I + 1
        Loop
        
        'blank out the days at the end of the month
        Do While I < 42
            If ShowNonMonthDays Then
                lblDay(I).Tag = D
                lblDay(I).Caption = Day(lblDay(I).Tag)
                D = DateAdd("d", 1, D)
            Else
                lblDay(I).Caption = ""
                lblDay(I).Tag = ""
            End If
            lblDay(I).ToolTipText = lblDay(I).Tag
            lblDay(I).ForeColor = vbGrayText
            I = I + 1
        Loop
    End If
    
End Sub

Private Sub RepositionCalendar()
    'positions the calendar under the usercontrol
    Dim R As RECT
    Dim RP As RECT
    Dim P As POINTAPI
    Dim retval As Long
    Dim hParent As Long
    
    If m_CanvasShown Then
        ' Determine where to show it in Screen coordinates:
        GetWindowRect UserControl.hwnd, R
        P.X = R.Left
        P.Y = R.Bottom - 1
        
        If UserControl.Parent.MDIChild Then
            'position needs to change relative to the MDI parent form
            hParent = GetParent(UserControl.Parent.hwnd)
            GetWindowRect hParent, RP
            P.X = P.X - RP.Left - 2
            P.Y = P.Y - RP.Top - 2
        End If
        
        'do the move
        retval = SetWindowPos(picCanvas.hwnd, UserControl.Parent.hwnd, _
            P.X, P.Y, _
            0, 0, _
            SWP_NOSIZE) 'or SWP_SHOWWINDOW
        
        'If retval = 0 Then
        '    Debug.Print "move failed"
        'Else
        '    Debug.Print "moved ok"
        'End If
        
        ' Tell VB it is shown:
        picCanvas.Visible = True
        picCanvas.ZOrder
    End If
End Sub

Private Sub ShowCanvas()
    Dim lStyle As Long
    ' Store a flag saying we are showing:
    m_CanvasShown = True
    
    ' Make sure the picture box won't appear in the
    ' task bar by making it into a Tool Window:
    lStyle = GetWindowLong(picCanvas.hwnd, GWL_EXSTYLE)
    lStyle = lStyle Or WS_EX_TOOLWINDOW
    lStyle = lStyle And Not (WS_EX_APPWINDOW)
    SetWindowLong picCanvas.hwnd, GWL_EXSTYLE, lStyle

    ' Make the picture box a child of the parent window (so
    ' it can be fully shown even if it extends beyond
    ' the form boundaries):
    SetParent picCanvas.hwnd, GetParent(UserControl.Parent.hwnd)

    'put calendar under the usercontrol
    RepositionCalendar

    RefreshButtons
    
    'Set focus on the calendar
    'SetFocus picCanvas.hwnd
    SetFocus picDrop.hwnd
    
End Sub
Private Sub HighlightSelection()
    'make sure all the days are in default colour
    Dim I As Long
    
    ShapeSelection.Move -1000, -1000
    For I = 0 To 41
        'show normal
        lblDay(I).BackColor = picCanvas.BackColor
        'unless its the selected date
        If Len(lblDay(I).Tag) > 0 Then
            If DateDiff("d", Date, lblDay(I).Tag) = 0 Then
                'its today - indicate it by changing backcolor
                lblDay(I).BackColor = picDaysHeaderBackground.BackColor
            End If
        
            If DateDiff("d", DateHighlighted, lblDay(I).Tag) = 0 Then
                'highlight the selection
                lblDay(I).BackColor = COL_SELECTEDDAYBACKGROUND
                ShapeSelection.Move lblDay(I).Left, lblDay(I).Top
                ShapeSelection.ZOrder
            End If
        End If
    Next
End Sub

Private Sub SyncCalendarDisplay()
    'Refreshes the data
    'called when initialising or changing months/years
    
    'put the days in the labels
    PopulateDays
    'add text to the main display
    ShowDateTitle
    ShowDateDescription
    'show day
    HighlightSelection

End Sub

Private Sub ShowCalendar()
    'puts it all together so when it pops up everything is displayed ok
    DateHighlighted = DateSelected
    'first initialise the vars
    InitialiseCanvas
    'then sync display with date
    SyncCalendarDisplay
    'then popup the window
    ShowCanvas
End Sub

Private Sub HideCalendar()
    'closes everything
    HideListBox
    HideCanvas
End Sub
Private Sub HideCanvas()
    'return ownership of the window to the control
    If m_CanvasShown Then
        SetParent picCanvas.hwnd, UserControl.hwnd
        picCanvas.Visible = False
        m_CanvasShown = False
    End If
End Sub

Private Function FocusIsWithThisControl() As Boolean
    'do we have the focus
    Dim hF As Long
    'Dim sName As String
    hF = GetFocus
    If hF = picToday.hwnd _
    Or hF = btnToday.hwnd _
    Or hF = picOK.hwnd _
    Or hF = btnOK.hwnd _
    Or hF = picDrop.hwnd _
    Or hF = btnDrop.hwnd _
    Or hF = Text1.hwnd _
    Or hF = picListBox.hwnd Then
        FocusIsWithThisControl = True
    Else
        'sName = String(100, Chr$(0))
        'GetWindowText hF, sName, 100
        'sName = Left$(sName, InStr(sName, Chr$(0)) - 1)
        'Debug.Print "focus is on " & hF & " " & sName
    End If
End Function

'----------
'MY LISTBOX
'----------
Private Sub InitialiseListBox()
    Dim I As Long
    
    'set this the same so that window will map properly when overlayed
    'For I = 0 To 6
    '    lblSelDate(I).Width = lblDateTitle.Width
    'Next
    
    'set width
    picListBox.Height = lblSelDate(0).Height * 7
    picListBox.Width = lblSelDate(0).Width
    'visibility
    picListBox.Visible = False
    picListBox.BorderStyle = 1
    
    PopulateMonths DateDisplaying
End Sub

Private Sub ShowListBox()
    Dim P As POINTAPI
    Dim lStyle As Long
    'Dim hParent As Long
    Dim R As RECT
    Dim RP As RECT
    
    'Make sure the list box won't appear in the
    'task bar by making it into a Tool Window:
    lStyle = GetWindowLong(picListBox.hwnd, GWL_EXSTYLE)
    lStyle = lStyle Or WS_EX_TOOLWINDOW
    lStyle = lStyle And Not (WS_EX_APPWINDOW)
    SetWindowLong picListBox.hwnd, GWL_EXSTYLE, lStyle

    'Make the list box is a child of the parent window
    SetParent picListBox.hwnd, GetParent(UserControl.Parent.hwnd)

    'Calc position.
    'Place it relative to lblDateTitle
    'Because lblDateTitle doesnt have a hWnd, have to work from the background pic
    GetWindowRect UserControl.picMonthHeaderBackground.hwnd, R
    P.X = R.Left + (lblDateTitle.Left / Screen.TwipsPerPixelX)
    P.Y = R.Bottom - (((picListBox.Height + lblSelDate(0).Height) / Screen.TwipsPerPixelY) / 2) - 2
    
    'adjust position if in MDI environment
    If UserControl.Parent.MDIChild Then
        'position needs to change relative to the MDI parent form
        GetWindowRect GetParent(UserControl.Parent.hwnd), RP
        P.X = P.X - RP.Left - 3
        P.Y = P.Y - RP.Top - 3
    End If
    
    'Show it
    SetWindowPos picListBox.hwnd, UserControl.Parent.hwnd, _
        P.X, P.Y, _
        0, 0, _
        SWP_SHOWWINDOW Or SWP_NOSIZE

    ' Tell VB it is shown:
    picListBox.Visible = True
    picListBox.ZOrder

    ' Try to set focus:
    SetFocus picListBox.hwnd

    ' Store a flag saying we're shown:
    m_ListBoxShown = True
End Sub

Private Sub HideListBox()
    If m_ListBoxShown Then
        TimerMonthTicker.Enabled = False
        SetParent picListBox.hwnd, UserControl.hwnd
        picListBox.Visible = False
        m_ListBoxShown = False
        SetFocus picDrop.hwnd
    End If
End Sub

Private Sub PopulateMonths(mDate As Date)
    'mDate is the middle date within the display

    Dim D As Date
    D = DateAdd("m", -3, mDate)
    lblSelDate(0) = MonthName(Month(D)) & " " & Year(D)
    lblSelDate(0).Tag = D
    
    D = DateAdd("m", -2, mDate)
    lblSelDate(1) = MonthName(Month(D)) & " " & Year(D)
    lblSelDate(1).Tag = D
    
    D = DateAdd("m", -1, mDate)
    lblSelDate(2) = MonthName(Month(D)) & " " & Year(D)
    lblSelDate(2).Tag = D
    
    D = mDate
    lblSelDate(3) = MonthName(Month(D)) & " " & Year(D)
    lblSelDate(3).Tag = D
    
    D = DateAdd("m", 1, mDate)
    lblSelDate(4) = MonthName(Month(D)) & " " & Year(D)
    lblSelDate(4).Tag = D
    
    D = DateAdd("m", 2, mDate)
    lblSelDate(5) = MonthName(Month(D)) & " " & Year(D)
    lblSelDate(5).Tag = D
    
    D = DateAdd("m", 3, mDate)
    lblSelDate(6) = MonthName(Month(D)) & " " & Year(D)
    lblSelDate(6).Tag = D
  
End Sub

Private Sub DynamicallyPlaceControls()
    Dim I As Long
    Dim R As Long
    Dim C As Long
    Dim W As Long
    Dim H As Long
    Dim P As Long   'padding
    
    'the best size is 330x225
    'have manipulated these numbers to get this result
    'but it might be better to do it another way
    'W = ((2.8 + lblDay(0).FontSize) * Screen.TwipsPerPixelX) * 2
    'H = (7 + lblDay(0).FontSize) * Screen.TwipsPerPixelY
    W = 330 '22
    H = 225 '15
    P = 2 * Screen.TwipsPerPixelX
    
    'apply result to day cells
    For I = 0 To 41
        lblDay(I).Width = W
        lblDay(I).Height = H
    Next
    'apply result to highlight shape
    ShapeSelection.Width = W
    ShapeSelection.Height = H
    'apply result to day names
    For I = 0 To 6
        lblDOW(I).Width = W
        lblDOW(I).Height = H
    Next
    
    'move day names to fit for new size
    'day names are aligned to picDaysHeaderBackground
    lblDOW(0).Left = 5 * Screen.TwipsPerPixelX
    For I = 1 To 6
        lblDOW(I).Left = lblDOW(I - 1).Left + lblDOW(I).Width + P
    Next
    
    'move days to be aligned
    'the first row
    lblDay(0).Left = picDaysHeaderBackground.Left + lblDOW(0).Left - Screen.TwipsPerPixelX
    lblDay(0).Top = picDaysHeaderBackground.Top + picDaysHeaderBackground.Height + P
    For I = 1 To 6
        lblDay(I).Left = lblDay(I - 1).Left + lblDay(I).Width + P
        lblDay(I).Top = lblDay(0).Top
    Next
    'then align the others under those
    For R = 1 To 5
        For C = 0 To 6
            I = C + (R * 7)
            If I <> 0 Then  'already set
                lblDay(I).Left = lblDay(C).Left
                lblDay(I).Top = lblDay(I - 7).Top + lblDay(I - 7).Height + P
            End If
        Next C
    Next R
    
    'center title and description
    lblDateTitle.Left = (ShapeBorderSmall.Width - lblDateTitle.Width) / 2
    
    'sort out Today button
    picToday.Width = btnToday.Width
    picToday.Height = btnToday.Height
    ShapeBorderSmall.Left = ShapeBorderSmall.Left
    picToday.Top = lblDateDescription.Top
    btnToday.Move picToday.Left, picToday.Top
    
    'sort out OK button
    picOK.Width = btnOK.Width
    picOK.Height = btnOK.Height
    picOK.Left = ShapeBorderSmall.Width - (picToday.Width - ShapeBorderSmall.Left)
    picOK.Top = lblDateDescription.Top
    btnOK.Move picOK.Left, picOK.Top
End Sub
Private Sub ShowHandCursor()
    If UseHandCursor Then
        If m_hHandCursor <> 0 Then
            SetCursor m_hHandCursor
        End If
    End If
End Sub
Private Sub InitialiseCursor()
    m_hHandCursor = LoadCursor(0, IDC_HAND)
End Sub
Private Sub InitialiseColours()
    'sets colours to windows theme
    Dim C As OLE_COLOR
    Dim I As Long
    
    UserControl.BackColor = vbWhite
    
    'selected day background colour
    COL_SELECTEDDAYBACKGROUND = RGB(251, 230, 148)
    'choose colour to apply
    Select Case CurrentTheme
    Case 0  'no theme
        C = vbButtonFace
    Case THEME_BLUE
        C = RGB(158, 190, 245)
    Case THEME_OLIVE
        C = RGB(217, 217, 167)
    Case THEME_SILVER
        C = RGB(215, 215, 229)
    End Select
    'apply colour
    picDaysHeaderBackground.BackColor = C
    ShapeBorderLarge.BorderColor = C
    ShapeBorderSmall.BorderColor = C
    ShapeBorderTextbox.BorderColor = C
    
    For I = 0 To 6
        lblDOW(I).BackColor = C
    Next
    
End Sub

'------------
'USER CONTROL
'------------
Private Sub UserControl_InitProperties()
    Text1.Text = UserControl.Extender.Name
    m_DescriptionFormat = DEF_DATEFORMAT
    m_ShowNonMonthDays = True
    FirstDayOfWeek = Sunday
    DateSelected = Date
    Enabled = True
    InitialiseThemes
    InitialiseColours
    Set Font = UserControl.Ambient.Font
End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    InitialiseThemes
    InitialiseColours
    
    If Ambient.UserMode Then
        Call Subclass_Start(UserControl.Parent.hwnd)
        'so we can hide calendar if focus shifts
        
'        Call Subclass_AddMsg(UserControl.Parent.hwnd, ALL_MESSAGES, MSG_AFTER)
        Call Subclass_AddMsg(UserControl.Parent.hwnd, WM_NCACTIVATE, MSG_AFTER)
        Call Subclass_AddMsg(UserControl.Parent.hwnd, WM_ACTIVATE, MSG_AFTER)
        Call Subclass_AddMsg(UserControl.Parent.hwnd, WM_LBUTTONDOWN, MSG_BEFORE)
        Call Subclass_AddMsg(UserControl.Parent.hwnd, WM_RBUTTONDOWN, MSG_BEFORE)
        'so we can move calendar with usercontrol
        Call Subclass_AddMsg(UserControl.Parent.hwnd, WM_SIZING, MSG_BEFORE)
        Call Subclass_AddMsg(UserControl.Parent.hwnd, WM_MOVE, MSG_AFTER)
        
        'so we can keep window on top after a form move
        Call Subclass_AddMsg(UserControl.Parent.hwnd, WM_PAINT, MSG_AFTER)
        Call Subclass_AddMsg(UserControl.Parent.hwnd, WM_ENTERSIZEMOVE, MSG_AFTER)
        Call Subclass_AddMsg(UserControl.Parent.hwnd, WM_EXITSIZEMOVE, MSG_AFTER)
        'theme support
        Call Subclass_AddMsg(UserControl.Parent.hwnd, WM_SYSCOLORCHANGE, MSG_AFTER)
        Call Subclass_AddMsg(UserControl.Parent.hwnd, WM_THEMECHANGED, MSG_AFTER)
        
        If m_UseThemes Then
            'we are drawing themed buttons so track the mouseleave events
            Call Subclass_Start(UserControl.picOK.hwnd)
            Call Subclass_AddMsg(UserControl.picOK.hwnd, WM_MOUSELEAVE, MSG_AFTER)
            Call Subclass_Start(UserControl.picToday.hwnd)
            Call Subclass_AddMsg(UserControl.picToday.hwnd, WM_MOUSELEAVE, MSG_AFTER)
            Call Subclass_Start(UserControl.picDrop.hwnd)
            Call Subclass_AddMsg(UserControl.picDrop.hwnd, WM_MOUSELEAVE, MSG_AFTER)
        End If
        
        Call Subclass_Start(UserControl.hwnd)
        'to check for visibility change
        Call Subclass_AddMsg(UserControl.hwnd, WM_SHOWWINDOW, MSG_AFTER)
        Call Subclass_AddMsg(UserControl.hwnd, WM_PRINT, MSG_AFTER)
    End If
    
    With PropBag
        FirstDayOfWeek = .ReadProperty("FirstDayOfWeek", 0)
        ShortDayNames = .ReadProperty("ShortDayNames", False)
        DescriptionFormat = .ReadProperty("DescriptionFormat", "")
        UseHandCursor = .ReadProperty("UseHandCursor", True)
        DateSelected = .ReadProperty("DateSelected", Date)
        ShowNonMonthDays = .ReadProperty("ShowNonMonthDays", False)
        Enabled = .ReadProperty("Enabled", True)
        Set Font = .ReadProperty("Font", Ambient.Font)
    End With
    
    InitialiseCanvas
    InitialiseListBox
    InitialiseCursor
    InitialiseCombo
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "FirstDayOfWeek", FirstDayOfWeek, 0
        .WriteProperty "ShortDayNames", ShortDayNames, False
        .WriteProperty "DescriptionFormat", DescriptionFormat, ""
        .WriteProperty "UseHandCursor", UseHandCursor, True
        .WriteProperty "DateSelected", DateSelected, Date
        .WriteProperty "ShowNonMonthDays", ShowNonMonthDays, False
        .WriteProperty "Enabled", Enabled, True
        .WriteProperty "Font", UserControl.Font, Ambient.Font
    End With
End Sub

Private Sub UserControl_Resize()
    Dim Padding As Long
    Dim MaxHeight As Long
    
    Padding = 3 * Screen.TwipsPerPixelX
    'fixed height
    MaxHeight = 20 * Screen.TwipsPerPixelY
    If UserControl.Height <> MaxHeight Then
        UserControl.Height = MaxHeight
    End If
    
    'outline
    ShapeBorderTextbox.Move 0, 0, UserControl.Width, UserControl.Height
    
    'set left margin
    Text1.Left = Padding
    'then center horz
    Text1.Width = (UserControl.Width - (2 * Text1.Left))
    'and center vert
    Text1.Top = (UserControl.Height - Text1.Height) / 2

    InitialiseCombo
End Sub

Private Sub UserControl_Terminate()
    On Error GoTo Errs
    HideCalendar
    
    If Ambient.UserMode Then Call Subclass_StopAll
    Debug.Print "UC terminated"
Errs:
End Sub

'---------------
'BUTTON / THEMES
'---------------
Private Sub RefreshButtons()
    If m_CanvasShown Then
        If m_UseThemes Then
            DrawButton picOK, "OK", m_OKButtonStateId
            DrawButton picToday, "Today", m_TodayButtonStateId
        End If
    End If
End Sub

Private Sub InitialiseThemes()
    m_UseThemes = CanDrawThemes
    If m_UseThemes Then
        btnOK.Visible = False
        btnToday.Visible = False
        m_OKButtonStateId = Normal
        m_TodayButtonStateId = Normal
        picToday.Visible = True
        picOK.Visible = True
        RefreshButtons
    Else
        btnOK.Visible = True
        btnToday.Visible = True
        picToday.Visible = False
        picOK.Visible = False
    End If
End Sub


Private Sub DrawButton(Surface As PictureBox, sText As String, lStateId As ButtonState)
    Dim hTheme As Long
    Dim tR As RECT
    Dim tTextR As RECT
    Dim tIconR As RECT
    Dim tImlR As RECT
    Dim retval As Long
    Dim L As Long
    Dim T As Long
    Dim hRgn As Long
    Const PARTID As Long = 1
    
    'set backup value
    If lStateId = 0 Then
        lStateId = Normal
    End If
    'get size for button
    tR.Left = 0
    tR.Top = 0
    tR.Right = Surface.Width / Screen.TwipsPerPixelX
    tR.Bottom = Surface.Height / Screen.TwipsPerPixelY
    
    If Not UserControl.Ambient.UserMode Then
        'force redraw if in development mode
        lStateId = Hot
    End If
    
    hTheme = OpenThemeData(Surface.hwnd, StrPtr("BUTTON"))
    If hTheme <> 0 Then
        Surface.cls
        'apply a region to maintain shape
        retval = GetThemeBackgroundRegion(hTheme, _
           UserControl.hdc, _
           PARTID, _
           lStateId, _
           tR, hRgn)
           
        ' free the memory.
        DeleteObject hRgn

        retval = DrawThemeBackground(hTheme, _
            Surface.hdc, _
            PARTID, _
            lStateId, _
            tR, tR)
     
     
        If Len(sText) > 0 Then
            retval = GetThemeBackgroundContentRect( _
                hTheme, _
                Surface.hdc, _
                PARTID, _
                lStateId, _
                tR, _
                tTextR)
             
            retval = DrawThemeText( _
               hTheme, _
               Surface.hdc, _
               PARTID, _
               lStateId, _
               StrPtr(sText), _
               -1, _
               CENTERED, _
               0, _
               tTextR)
        End If
        
        If GetFocus = Surface.hwnd Then
            tR.Left = tR.Left + 3
            tR.Top = tR.Top + 3
            tR.Right = tR.Right - 3
            tR.Bottom = tR.Bottom - 3
            DrawFocusRect Surface.hdc, tR
        End If
        
        CloseThemeData hTheme
        Surface.Refresh
    End If
End Sub

Private Sub DrawDrop()
    Dim hTheme As Long
    Dim tR As RECT
    Dim tTextR As RECT
    Dim tIconR As RECT
    Dim tImlR As RECT
    Dim retval As Long
    Dim L As Long
    Dim T As Long
    Dim hRgn As Long
    Const PARTID As Long = 1
    
    'only do the draw if themes are available
    If Not m_UseThemes Then Exit Sub
    
    'set backup value
    If m_DropButtonStateId = 0 Then
        m_DropButtonStateId = Normal
    End If
    
    'get size for button
    tR.Left = 0
    tR.Top = 0
    tR.Right = picDrop.Width / Screen.TwipsPerPixelX
    tR.Bottom = picDrop.Height / Screen.TwipsPerPixelY
    
    If Not UserControl.Ambient.UserMode Then
        'force redraw if in development mode
        m_DropButtonStateId = Hot
    End If
    
    hTheme = OpenThemeData(picDrop.hwnd, StrPtr("COMBOBOX"))
    If hTheme <> 0 Then
        picDrop.cls
        'trim off the background using region
        retval = GetThemeBackgroundRegion(hTheme, _
           UserControl.hdc, _
           PARTID, _
           m_DropButtonStateId, _
           tR, hRgn)
        ' free the memory.
        DeleteObject hRgn
        
        retval = DrawThemeParentBackground( _
            picDrop.hwnd, _
            picDrop.hdc, _
            tR)

        retval = DrawThemeBackground(hTheme, _
            picDrop.hdc, _
            PARTID, _
            m_DropButtonStateId, _
            tR, tR)
     
        CloseThemeData hTheme
        picDrop.Refresh
    End If
End Sub

Private Function CanDrawThemes() As Boolean
    'tests to see if the computer can do themes
    'if it can then we can use the nice XP buttons
    'if it cant we have to use the old 98 style buttons
    If Not DebugMode = True Then On Error Resume Next
    Dim hTheme As Long
    'opening and closing theme
    hTheme = OpenThemeData(UserControl.hwnd, StrPtr("BUTTON"))
    On Error GoTo 0
    If hTheme = 0 Then
        CanDrawThemes = False
    Else
        CanDrawThemes = True
        CloseThemeData hTheme
    End If
End Function

Public Function CurrentTheme() As Long
Attribute CurrentTheme.VB_Description = "Returns what theme XP is using or 0 if no themes."
    On Error GoTo whoops
    Dim hTheme As Long
    Dim bThemeFile() As Byte
    Dim sThemeFile As String
    Dim sColorName As String
    Dim lPtrThemeFile As Long
    Dim lPtrColourName As Long
    
    Dim retval As Long
    
    hTheme = OpenThemeData(UserControl.hwnd, StrPtr("BUTTON"))
   
    If hTheme <> 0 Then
        ReDim bThemeFile(0 To 260 * 2) As Byte
        lPtrThemeFile = VarPtr(bThemeFile(0))
        ReDim bColorName(0 To 260 * 2) As Byte
        lPtrColourName = VarPtr(bColorName(0))
        retval = GetCurrentThemeName(lPtrThemeFile, 260, lPtrColourName, 260, 0, 0)
    
        sThemeFile = bThemeFile
        If InStr(LCase(sThemeFile), "luna.msstyles") Then
            sColorName = bColorName
            If InStr(LCase(sColorName), "normalcolor") Then
                CurrentTheme = THEME_BLUE
            End If
            If InStr(LCase(sColorName), "homestead") Then
                CurrentTheme = THEME_OLIVE
            End If
            If InStr(LCase(sColorName), "metallic") Then
                CurrentTheme = THEME_SILVER
            End If
        End If
      
        CloseThemeData hTheme
    End If
whoops:

End Function


Private Sub TrackMouseLeave(hwnd As Long)
    'Starts tracking the mouse
    'When the mouse leaves the control the WM_MOUSELEAVE message will be sent
    'Doesnt work for transparent windows :(
    On Error GoTo Errs
    Dim tme As TRACKMOUSEEVENT_STRUCT
    With tme
        .cbSize = Len(tme)
        .dwFlags = TME_LEAVE
        .hwndTrack = hwnd
    End With
    Call TrackMouseEvent(tme) '---- Track the mouse leaving the indicated window via subclassing
Errs:
End Sub



'========================================================================================
'Subclass routines below here - The programmer may call any of the following Subclass_??? routines
'======================================================================================================================================================
'Add a message to the table of those that will invoke a callback. You should Subclass_Start first and then add the messages
Private Sub Subclass_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
On Error GoTo Errs
'Parameters:
  'lng_hWnd  - The handle of the window for which the uMsg is to be added to the callback table
  'uMsg      - The message number that will invoke a callback. NB Can also be ALL_MESSAGES, ie all messages will callback
  'When      - Whether the msg is to callback before, after or both with respect to the the default (previous) handler
  With sc_aSubData(zIdx(lng_hWnd))
    If When And eMsgWhen.MSG_BEFORE Then
      Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
    End If
    If When And eMsgWhen.MSG_AFTER Then
      Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
    End If
  End With
Errs:
End Sub

'Delete a message from the table of those that will invoke a callback.
Private Sub Subclass_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
On Error GoTo Errs

'Parameters:
  'lng_hWnd  - The handle of the window for which the uMsg is to be removed from the callback table
  'uMsg      - The message number that will be removed from the callback table. NB Can also be ALL_MESSAGES, ie all messages will callback
  'When      - Whether the msg is to be removed from the before, after or both callback tables
  With sc_aSubData(zIdx(lng_hWnd))
    If When And eMsgWhen.MSG_BEFORE Then
      Call zDelMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
    End If
    If When And eMsgWhen.MSG_AFTER Then
      Call zDelMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
    End If
  End With
Errs:
End Sub

'Return whether we're running in the IDE.
Private Function Subclass_InIDE() As Boolean
  Debug.Assert zSetTrue(Subclass_InIDE)
End Function

'Start subclassing the passed window handle
Private Function Subclass_Start(ByVal lng_hWnd As Long) As Long
On Error GoTo Errs
'Parameters:
  'lng_hWnd  - The handle of the window to be subclassed
'Returns;
  'The sc_aSubData() index
  Const CODE_LEN              As Long = 200                                             'Length of the machine code in bytes
  Const FUNC_CWP              As String = "CallWindowProcA"                             'We use CallWindowProc to call the original WndProc
  Const FUNC_EBM              As String = "EbMode"                                      'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
  Const FUNC_SWL              As String = "SetWindowLongA"                              'SetWindowLongA allows the cSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
  Const MOD_USER              As String = "user32"                                      'Location of the SetWindowLongA & CallWindowProc functions
  Const MOD_VBA5              As String = "vba5"                                        'Location of the EbMode function if running VB5
  Const MOD_VBA6              As String = "vba6"                                        'Location of the EbMode function if running VB6
  Const PATCH_01              As Long = 18                                              'Code buffer offset to the location of the relative address to EbMode
  Const PATCH_02              As Long = 68                                              'Address of the previous WndProc
  Const PATCH_03              As Long = 78                                              'Relative address of SetWindowsLong
  Const PATCH_06              As Long = 116                                             'Address of the previous WndProc
  Const PATCH_07              As Long = 121                                             'Relative address of CallWindowProc
  Const PATCH_0A              As Long = 186                                             'Address of the owner object
  Static aBuf(1 To CODE_LEN)  As Byte                                                   'Static code buffer byte array
  Static pCWP                 As Long                                                   'Address of the CallWindowsProc
  Static pEbMode              As Long                                                   'Address of the EbMode IDE break/stop/running function
  Static pSWL                 As Long                                                   'Address of the SetWindowsLong function
  Dim I                       As Long                                                   'Loop index
  Dim j                       As Long                                                   'Loop index
  Dim nSubIdx                 As Long                                                   'Subclass data index
  Dim sHex                    As String                                                 'Hex code string
  
'If it's the first time through here..
  If aBuf(1) = 0 Then
  
'The hex pair machine code representation.
    sHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & _
           "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & _
           "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & _
           "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"

'Convert the string from hex pairs to bytes and store in the static machine code buffer
    I = 1
    Do While j < CODE_LEN
      j = j + 1
      aBuf(j) = Val("&H" & Mid$(sHex, I, 2))                                            'Convert a pair of hex characters to an eight-bit value and store in the static code buffer array
      I = I + 2
    Loop                                                                                'Next pair of hex characters
    
'Get API function addresses
    If Subclass_InIDE Then                                                              'If we're running in the VB IDE
      aBuf(16) = &H90                                                                   'Patch the code buffer to enable the IDE state code
      aBuf(17) = &H90                                                                   'Patch the code buffer to enable the IDE state code
      pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)                                           'Get the address of EbMode in vba6.dll
      If pEbMode = 0 Then                                                               'Found?
        pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)                                         'VB5 perhaps
      End If
    End If
    
    pCWP = zAddrFunc(MOD_USER, FUNC_CWP)                                                'Get the address of the CallWindowsProc function
    pSWL = zAddrFunc(MOD_USER, FUNC_SWL)                                                'Get the address of the SetWindowLongA function
    ReDim sc_aSubData(0 To 0) As tSubData                                               'Create the first sc_aSubData element
  Else
    nSubIdx = zIdx(lng_hWnd, True)
    If nSubIdx = -1 Then                                                                'If an sc_aSubData element isn't being re-cycled
      nSubIdx = UBound(sc_aSubData()) + 1                                               'Calculate the next element
      ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData                              'Create a new sc_aSubData element
    End If
    
    Subclass_Start = nSubIdx
  End If

  With sc_aSubData(nSubIdx)
    .hwnd = lng_hWnd                                                                    'Store the hWnd
    .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)                                       'Allocate memory for the machine code WndProc
    .nAddrOrig = SetWindowLongA(.hwnd, GWL_WNDPROC, .nAddrSub)                          'Set our WndProc in place
    Call RtlMoveMemory(ByVal .nAddrSub, aBuf(1), CODE_LEN)                              'Copy the machine code from the static byte array to the code array in sc_aSubData
    Call zPatchRel(.nAddrSub, PATCH_01, pEbMode)                                        'Patch the relative address to the VBA EbMode api function, whether we need to not.. hardly worth testing
    Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)                                     'Original WndProc address for CallWindowProc, call the original WndProc
    Call zPatchRel(.nAddrSub, PATCH_03, pSWL)                                           'Patch the relative address of the SetWindowLongA api function
    Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)                                     'Original WndProc address for SetWindowLongA, unsubclass on IDE stop
    Call zPatchRel(.nAddrSub, PATCH_07, pCWP)                                           'Patch the relative address of the CallWindowProc api function
    Call zPatchVal(.nAddrSub, PATCH_0A, ObjPtr(Me))                                     'Patch the address of this object instance into the static machine code buffer
  End With
Errs:
End Function

'Stop all subclassing
Private Sub Subclass_StopAll()
On Error GoTo Errs
  Dim I As Long
  
  I = UBound(sc_aSubData())                                                             'Get the upper bound of the subclass data array
  Do While I >= 0                                                                       'Iterate through each element
    With sc_aSubData(I)
      If .hwnd <> 0 Then                                                                'If not previously Subclass_Stop'd
        Call Subclass_Stop(.hwnd)                                                       'Subclass_Stop
      End If
    End With
    I = I - 1                                                                           'Next element
  Loop
Errs:
End Sub

'Stop subclassing the passed window handle
Private Sub Subclass_Stop(ByVal lng_hWnd As Long)
On Error GoTo Errs
'Parameters:
  'lng_hWnd  - The handle of the window to stop being subclassed
  With sc_aSubData(zIdx(lng_hWnd))
    Call SetWindowLongA(.hwnd, GWL_WNDPROC, .nAddrOrig)                                 'Restore the original WndProc
    Call zPatchVal(.nAddrSub, PATCH_05, 0)                                              'Patch the Table B entry count to ensure no further 'before' callbacks
    Call zPatchVal(.nAddrSub, PATCH_09, 0)                                              'Patch the Table A entry count to ensure no further 'after' callbacks
    Call GlobalFree(.nAddrSub)                                                          'Release the machine code memory
    .hwnd = 0                                                                           'Mark the sc_aSubData element as available for re-use
    .nMsgCntB = 0                                                                       'Clear the before table
    .nMsgCntA = 0                                                                       'Clear the after table
    Erase .aMsgTblB                                                                     'Erase the before table
    Erase .aMsgTblA                                                                     'Erase the after table
  End With
Errs:
End Sub

'=======================================================================================================
'These z??? routines are exclusively called by the Subclass_??? routines.

'Worker sub for Subclass_AddMsg
Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
On Error GoTo Errs
  Dim nEntry  As Long                                                                   'Message table entry index
  Dim nOff1   As Long                                                                   'Machine code buffer offset 1
  Dim nOff2   As Long                                                                   'Machine code buffer offset 2
  
  If uMsg = ALL_MESSAGES Then                                                           'If all messages
    nMsgCnt = ALL_MESSAGES                                                              'Indicates that all messages will callback
  Else                                                                                  'Else a specific message number
    Do While nEntry < nMsgCnt                                                           'For each existing entry. NB will skip if nMsgCnt = 0
      nEntry = nEntry + 1
      
      If aMsgTbl(nEntry) = 0 Then                                                       'This msg table slot is a deleted entry
        aMsgTbl(nEntry) = uMsg                                                          'Re-use this entry
        Exit Sub                                                                        'Bail
      ElseIf aMsgTbl(nEntry) = uMsg Then                                                'The msg is already in the table!
        Exit Sub                                                                        'Bail
      End If
    Loop                                                                                'Next entry

    nMsgCnt = nMsgCnt + 1                                                               'New slot required, bump the table entry count
    ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long                                        'Bump the size of the table.
    aMsgTbl(nMsgCnt) = uMsg                                                             'Store the message number in the table
  End If

  If When = eMsgWhen.MSG_BEFORE Then                                                    'If before
    nOff1 = PATCH_04                                                                    'Offset to the Before table
    nOff2 = PATCH_05                                                                    'Offset to the Before table entry count
  Else                                                                                  'Else after
    nOff1 = PATCH_08                                                                    'Offset to the After table
    nOff2 = PATCH_09                                                                    'Offset to the After table entry count
  End If

  If uMsg <> ALL_MESSAGES Then
    Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))                                    'Address of the msg table, has to be re-patched because Redim Preserve will move it in memory.
  End If
  Call zPatchVal(nAddr, nOff2, nMsgCnt)                                                 'Patch the appropriate table entry count
Errs:
End Sub

'Return the memory address of the passed function in the passed dll
Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
  zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
  Debug.Assert zAddrFunc                                                                'You may wish to comment out this line if you're using vb5 else the EbMode GetProcAddress will stop here everytime because we look for vba6.dll first
End Function

'Worker sub for Subclass_DelMsg
Private Sub zDelMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
On Error GoTo Errs
  Dim nEntry As Long
  
  If uMsg = ALL_MESSAGES Then                                                           'If deleting all messages
    nMsgCnt = 0                                                                         'Message count is now zero
    If When = eMsgWhen.MSG_BEFORE Then                                                  'If before
      nEntry = PATCH_05                                                                 'Patch the before table message count location
    Else                                                                                'Else after
      nEntry = PATCH_09                                                                 'Patch the after table message count location
    End If
    Call zPatchVal(nAddr, nEntry, 0)                                                    'Patch the table message count to zero
  Else                                                                                  'Else deleteting a specific message
    Do While nEntry < nMsgCnt                                                           'For each table entry
      nEntry = nEntry + 1
      If aMsgTbl(nEntry) = uMsg Then                                                    'If this entry is the message we wish to delete
        aMsgTbl(nEntry) = 0                                                             'Mark the table slot as available
        Exit Do                                                                         'Bail
      End If
    Loop                                                                                'Next entry
  End If
Errs:
End Sub

'Get the sc_aSubData() array index of the passed hWnd
Private Function zIdx(ByVal lng_hWnd As Long, Optional ByVal bAdd As Boolean = False) As Long
On Error GoTo Errs
'Get the upper bound of sc_aSubData() - If you get an error here, you're probably Subclass_AddMsg-ing before Subclass_Start
  zIdx = UBound(sc_aSubData)
  Do While zIdx >= 0                                                                    'Iterate through the existing sc_aSubData() elements
    With sc_aSubData(zIdx)
      If .hwnd = lng_hWnd Then                                                          'If the hWnd of this element is the one we're looking for
        If Not bAdd Then                                                                'If we're searching not adding
          Exit Function                                                                 'Found
        End If
      ElseIf .hwnd = 0 Then                                                             'If this an element marked for reuse.
        If bAdd Then                                                                    'If we're adding
          Exit Function                                                                 'Re-use it
        End If
      End If
    End With
    zIdx = zIdx - 1                                                                     'Decrement the index
  Loop
  
'  If Not bAdd Then
'    Debug.Assert False                                                                  'hWnd not found, programmer error
'  End If
Errs:

'If we exit here, we're returning -1, no freed elements were found
End Function

'Patch the machine code buffer at the indicated offset with the relative address to the target address.
Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
  Call RtlMoveMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
End Sub

'Patch the machine code buffer at the indicated offset with the passed value
Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
  Call RtlMoveMemory(ByVal nAddr + nOffset, nValue, 4)
End Sub

'Worker function for Subclass_InIDE
Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
  zSetTrue = True
  bValue = True
End Function

'END Subclassing Code===================================================================================

