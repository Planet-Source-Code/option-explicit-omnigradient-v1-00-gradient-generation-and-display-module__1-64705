VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "OmniGradient Demo - Matthew R. Usner"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   6705
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkMiddleOut 
      Caption         =   "Middle-Out"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3840
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin Project1.ColorPick ColorPick1 
      Height          =   3375
      Left            =   120
      TabIndex        =   2
      Top             =   4200
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   5953
   End
   Begin VB.CheckBox chkCircular 
      Caption         =   "Circular"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3480
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   120
      ScaleHeight     =   215
      ScaleMode       =   0  'User
      ScaleWidth      =   431
      TabIndex        =   0
      Top             =   120
      Width           =   6495
   End
   Begin Project1.ColorPick ColorPick2 
      Height          =   3375
      Left            =   3600
      TabIndex        =   3
      Top             =   4200
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   5953
   End
   Begin MSComctlLib.Slider sldAngle 
      Height          =   375
      Left            =   1800
      TabIndex        =   9
      Top             =   3720
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      LargeChange     =   1
      Max             =   360
      SelStart        =   116
      TickStyle       =   3
      TickFrequency   =   10
      Value           =   116
   End
   Begin VB.Label lblGradAngle 
      Alignment       =   2  'Center
      Caption         =   "Angle = 116"
      Height          =   255
      Left            =   1800
      TabIndex        =   10
      Top             =   3405
      Width           =   2535
   End
   Begin VB.Label lblDisplaySpeed 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0 ms"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5520
      TabIndex        =   7
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label lblCalcSpeed 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0 ms"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5520
      TabIndex        =   6
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Display Speed:"
      Height          =   255
      Left            =   4320
      TabIndex        =   5
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Calc Speed:"
      Height          =   255
      Left            =   4440
      TabIndex        =   4
      Top             =   3600
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As Any, ByVal wUsage As Long, ByVal dwRop As Long) As Long

' declares for sizeable picturebox
Private Const GWL_STYLE     As Long = -16
Private Const SWP_DRAWFRAME As Long = &H20
Private Const SWP_NOMOVE    As Long = &H2
Private Const SWP_NOSIZE    As Long = &H1
Private Const SWP_NOZORDER  As Long = &H4
Private Const WS_THICKFRAME As Long = &H40000
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

'  holds gradient information for displaying to sample picturebox.
Private BG_uBIH                     As BITMAPINFOHEADER    ' udt defined in module.
Private BG_lBits()                  As Long

Private GradientAngle As Single
Private MiddleOut As Boolean

Private Sub chkCircular_Click()

   If chkCircular.Value = vbChecked Then
      chkMiddleOut.Enabled = False
      sldAngle.Enabled = False
      lblGradAngle.Enabled = False
   Else
      chkMiddleOut.Enabled = True
      sldAngle.Enabled = True
      lblGradAngle.Enabled = True
   End If

   DisplayGradient

End Sub

Private Sub chkMiddleOut_Click()
   MiddleOut = (chkMiddleOut.Value = vbChecked)
   DisplayGradient
End Sub

Private Sub ColorPick1_Click()
   DisplayGradient
End Sub

Private Sub ColorPick2_Click()
   DisplayGradient
End Sub

Private Sub Form_Load()

   Dim lold As Long

'  Make picturebox canvas sizeable.
   With Picture1
      lold = GetWindowLong(.hwnd, GWL_STYLE)
      lold = SetWindowLong(.hwnd, GWL_STYLE, lold Or WS_THICKFRAME)
      SetWindowPos .hwnd, Form1.hwnd, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME
   End With

   GradientAngle = 116
   MiddleOut = True
   DisplayGradient

End Sub

Private Sub DisplayGradient()

   Dim CircularFlag As Boolean
   Dim Clr1 As Long
   Dim Clr2 As Long
   Dim SwpClr As Long

   Dim sw As New CStopWatch

   Clr1 = ColorPick1.RGBLong
   Clr2 = ColorPick2.RGBLong
   CircularFlag = (chkCircular.Value = vbChecked)

   If CircularFlag Then
      SwpClr = Clr1
      Clr1 = Clr2
      Clr2 = SwpClr
   End If

   With Picture1

      sw.Reset
      CalcGradient .ScaleWidth, .ScaleHeight, Clr1, Clr2, GradientAngle, MiddleOut, BG_uBIH, BG_lBits(), CircularFlag
      lblCalcSpeed.Caption = sw.Elapsed & " ms"

      sw.Reset
      mGradient.PaintGradient .hDC, 0, 0, .ScaleWidth, .ScaleHeight, 0, 0, .ScaleWidth, .ScaleHeight, BG_lBits(), BG_uBIH
      lblDisplaySpeed.Caption = sw.Elapsed & " ms"

      .Refresh
   End With
   
End Sub

Private Sub Picture1_Resize()
   If Picture1.ScaleHeight < 1 Or Picture1.ScaleWidth < 1 Then
      Exit Sub
   End If
   DisplayGradient
End Sub

Private Sub sldAngle_Scroll()
   GradientAngle = sldAngle.Value
   lblGradAngle.Caption = "Angle = " & CStr(GradientAngle)
   DisplayGradient
End Sub


