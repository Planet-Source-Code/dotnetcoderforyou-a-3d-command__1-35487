VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6210
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   5880
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   5880
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton Option2 
      Caption         =   "Option2"
      Height          =   375
      Left            =   3420
      TabIndex        =   13
      Top             =   4275
      Width           =   1680
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   375
      Left            =   585
      TabIndex        =   12
      Top             =   4275
      Width           =   1410
   End
   Begin VB.ComboBox Combo2 
      Height          =   420
      Left            =   3150
      TabIndex        =   11
      Text            =   "Combo2"
      Top             =   3645
      Width           =   2400
   End
   Begin VB.ComboBox Combo1 
      Height          =   420
      Left            =   45
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   3690
      Width           =   2445
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   285
      Left            =   3330
      TabIndex        =   9
      Top             =   3195
      Width           =   1995
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   240
      Left            =   675
      TabIndex        =   8
      Top             =   3330
      Width           =   1140
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   375
      Left            =   3780
      TabIndex        =   7
      Top             =   2700
      Width           =   1230
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   420
      Left            =   585
      TabIndex        =   6
      Top             =   2745
      Width           =   1320
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Normal"
      Height          =   600
      Left            =   3060
      TabIndex        =   4
      Top             =   720
      Width           =   2850
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Click ME To See 3D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   90
      TabIndex        =   2
      Top             =   675
      Width           =   2265
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   90
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   90
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3120
      TabIndex        =   0
      Text            =   "Normal"
      Top             =   120
      Width           =   2760
   End
   Begin VB.Image Image2 
      Height          =   690
      Left            =   3960
      Picture         =   "Form1.frx":0000
      Top             =   1980
      Width           =   750
   End
   Begin VB.Image Image1 
      Height          =   690
      Left            =   810
      Picture         =   "Form1.frx":396C
      Top             =   1935
      Width           =   750
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   375
      Left            =   3300
      TabIndex        =   5
      Top             =   1500
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   135
      TabIndex        =   3
      Top             =   1395
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Sub ThreeDControl(Ctrl As Control, nBevel%, nSpace%, bInset%)

PixX% = Screen.TwipsPerPixelX
PixY% = Screen.TwipsPerPixelY

CTop% = Ctrl.Top - PixX%
CLft% = Ctrl.Left - PixY%
CRgt% = Ctrl.Left + Ctrl.Width
CBtm% = Ctrl.Top + Ctrl.Height

' Color used below:
' dark gray = &H808080
' white = &HFFFFFF

If bInset% Then 'inset border
  For I% = nSpace% To (nBevel% + nSpace% - 1)
  AddX% = I% * PixX%
  AddY% = I% * PixY%
  Ctrl.Parent.Line (CLft% - AddX%, CTop% - AddY%)-(CRgt% + _
  AddX%, CTop% - AddY%), &H808080
  Ctrl.Parent.Line (CLft% - AddX%, CTop% - AddY%)-(CLft% - _
  AddX%, CBtm% + AddY%), &H808080
  Ctrl.Parent.Line (CLft% - AddX%, CBtm% + AddY%)-(CRgt% + _
  AddX% + PixX%, CBtm% + AddY%), &HFFFFFF
  Ctrl.Parent.Line (CRgt% + AddX%, CTop% - AddY%)-(CRgt% + _
  AddX%, CBtm% + AddY%), &HFFFFFF
  Next
Else 'outset border
  For I% = nSpace% To (nBevel% + nSpace% - 1)
  AddX% = I% * PixX%
  AddY% = I% * PixY%
  Ctrl.Parent.Line (CRgt% + AddX%, CBtm% + AddY%)-(CRgt% + _
  AddX%, CTop% - AddY%), &H808080
  Ctrl.Parent.Line (CRgt% + AddX%, CBtm% + AddY%)-(CLft% - _
  AddX%, CBtm% + AddY%), &H808080
  Ctrl.Parent.Line (CRgt% + AddX%, CTop% - AddY%)-(CLft% - _
  AddX% - PixX%, CTop% - AddY%), &HFFFFFF
  Ctrl.Parent.Line (CLft% - AddX%, CBtm% + AddY%)-(CLft% - _
  AddX%, CTop% - AddY%), &HFFFFFF
  Next
End If

End Sub

Private Sub Command1_Click()
 Call ThreeDControl(Command1, 1, 0, True)
 Call ThreeDControl(Text2, 1, 0, True)
 Call ThreeDControl(Label1, 1, 0, True)
 Call ThreeDControl(Check1, 1, 0, True)
  Call ThreeDControl(Image1, 1, 0, True)
 Call ThreeDControl(HScroll1, 1, 0, True)
 Call ThreeDControl(Combo2, 1, 0, True)
Call ThreeDControl(Option1, 1, 0, True)
End Sub

Private Sub Form_Load()
Dim r
End Sub



