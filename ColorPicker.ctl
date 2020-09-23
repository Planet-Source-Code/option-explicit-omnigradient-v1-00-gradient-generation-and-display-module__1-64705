VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl ColorPick 
   ClientHeight    =   3480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2940
   ScaleHeight     =   3480
   ScaleWidth      =   2940
   Begin VB.Frame fraColorPicker 
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      Begin MSComCtl2.FlatScrollBar sbRed 
         Height          =   255
         Left            =   2400
         TabIndex        =   6
         Top             =   1320
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Arrows          =   65536
         Max             =   255
         Orientation     =   8323073
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2535
         Left            =   120
         Picture         =   "ColorPicker.ctx":0000
         ScaleHeight     =   167
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   142
         TabIndex        =   2
         Top             =   240
         Width           =   2160
         Begin VB.Label lblCrosshair 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "+"
            ForeColor       =   &H00000000&
            Height          =   150
            Left            =   360
            TabIndex        =   9
            Top             =   840
            Visible         =   0   'False
            Width           =   90
         End
      End
      Begin VB.PictureBox PicSelected 
         BackColor       =   &H00000000&
         Height          =   705
         Left            =   2400
         ScaleHeight     =   645
         ScaleWidth      =   330
         TabIndex        =   1
         Top             =   240
         Width           =   390
      End
      Begin MSComCtl2.FlatScrollBar sbBlue 
         Height          =   255
         Left            =   2400
         TabIndex        =   7
         Top             =   2520
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Arrows          =   65536
         Max             =   255
         Orientation     =   8323073
      End
      Begin MSComCtl2.FlatScrollBar sbGreen 
         Height          =   255
         Left            =   2400
         TabIndex        =   8
         Top             =   1920
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Arrows          =   65536
         Max             =   255
         Orientation     =   8323073
      End
      Begin VB.Label lblRGBLong 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0 (&&H0)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   2880
         Width           =   2655
      End
      Begin VB.Label lblBlue 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   2400
         TabIndex        =   5
         Top             =   2280
         Width           =   390
      End
      Begin VB.Label lblGreen 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   2400
         TabIndex        =   4
         Top             =   1680
         Width           =   390
      End
      Begin VB.Label lblRed 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   255
         Left            =   2400
         TabIndex        =   3
         Top             =   1080
         Width           =   390
      End
   End
   Begin VB.Menu mnuCopy 
      Caption         =   "Copy"
      Visible         =   0   'False
      Begin VB.Menu mnuCopyDecimal 
         Caption         =   "Copy Decimal Color Number to Clipboard"
      End
      Begin VB.Menu mnuCopyHex 
         Caption         =   "Copy Hex Color Number to Clipboard"
      End
   End
End
Attribute VB_Name = "ColorPick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*************************************************************************
'* cheesy color picker control                                           *
'* written 09/17/2004 by Matthew R. Usner                                *
'*************************************************************************

Option Explicit

Private Type RGBColor
   Red As Integer
   Green As Integer
   Blue As Integer
End Type

'Default Property Values:
Const m_def_Enabled = 0
Const m_def_Red = 0
Const m_def_Green = 0
Const m_def_Blue = 0
Const m_def_RGBLong = 0

'Property Variables:
Dim m_Enabled As Boolean
Dim m_Red As Long
Dim m_Green As Long
Dim m_Blue As Long
Dim m_RGBLong As Long

Public Event Click()
'Event Declarations:
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."


'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
   m_Enabled = m_def_Enabled
   m_Red = m_def_Red
   m_Green = m_def_Green
   m_Blue = m_def_Blue
   m_RGBLong = m_def_RGBLong
End Sub

Private Sub picColor_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

'***************************************************************
'* changes color on the fly if holding down left mouse button  *
'***************************************************************

   On Error Resume Next

   If Button = vbLeftButton Then
      PicSelected.BackColor = picColor.Point(x, y)
      Me.RGBLong = PicSelected.BackColor
      ConvertToRGB (PicSelected.BackColor)
      UpdateRGBLabels
      DisplayCrosshair x, y
      RaiseEvent Click
   End If

End Sub

Private Sub DisplayCrosshair(x As Single, y As Single)

'***************************************************************
'* displays the 'color selected' crosshair in palette bitmap   *
'***************************************************************

   lblCrosshair.Visible = True
   lblCrosshair.Left = x - 3
   lblCrosshair.Top = y - 7
'  make sure crosshair is visible no matter what color it's on
   lblCrosshair.ForeColor = Abs(picColor.Point(x, y) - &HFFFFFF)

End Sub

Public Sub SetColorBar(CVal As Long)

'***************************************************************
'* sets color bar to color of selected InfiniCalc attribute    *
'***************************************************************

   PicSelected.BackColor = CVal
   Me.RGBLong = PicSelected.BackColor
   ConvertToRGB (PicSelected.BackColor)
   UpdateRGBLabels

End Sub

Private Sub mnuCopyHex_click()
   Clipboard.Clear
   Clipboard.SetText "&H" & CStr(Hex(PicSelected.BackColor))
End Sub

Private Sub mnuCopyDecimal_click()
   Clipboard.Clear
   Clipboard.SetText PicSelected.BackColor
End Sub

Private Sub picColor_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

'***************************************************************
'* changes color if left mouse button is clicked               *
'***************************************************************

   On Error Resume Next

   If Button = vbLeftButton Then
      PicSelected.BackColor = picColor.Point(x, y)
      Me.RGBLong = PicSelected.BackColor
      ConvertToRGB (PicSelected.BackColor)
      UpdateRGBLabels
      DisplayCrosshair x, y
   Else
      PopupMenu mnuCopy
   End If

End Sub

Private Sub UpdateRGBLabels()

'***************************************************************
'* puts appropriate rgb values in color labels and scrollbars  *
'***************************************************************

   lblRed.Caption = CStr(Me.Red)
   lblGreen.Caption = CStr(Me.Green)
   lblBlue.Caption = CStr(Me.Blue)
   sbRed.Value = CStr(Me.Red)
   sbGreen.Value = CStr(Me.Green)
   sbBlue.Value = CStr(Me.Blue)
   DisplayColorLongs RGB(Me.Red, Me.Green, Me.Blue)

End Sub

Private Sub DisplayColorLongs(RGBNum As Long)
   lblRGBLong.Caption = CStr(RGBNum) & " (&&H" & Hex(RGBNum) & ")"
End Sub

Public Sub ConvertToRGB(ByVal ColorVal As Long)

'***************************************************************
'* converts color long to red, green, blue values              *
'***************************************************************

   Me.Blue = Int(ColorVal / 65536)
   Me.Green = Int((ColorVal - (65536 * Blue)) / 256)
   Me.Red = ColorVal - (65536 * Blue + 256 * Green)

End Sub

Private Sub sbBlue_Change()
   lblBlue.Caption = CStr(sbBlue.Value)
   Me.Blue = sbBlue.Value
   PicSelected.BackColor = RGB(Me.Red, Me.Green, Me.Blue)
   Me.RGBLong = PicSelected.BackColor
   lblRGBLong.Caption = PicSelected.BackColor
   DisplayColorLongs RGB(Me.Red, Me.Green, Me.Blue)
   RaiseEvent Click
End Sub

Private Sub sbGreen_Change()
   lblGreen.Caption = CStr(sbGreen.Value)
   Me.Green = sbGreen.Value
   PicSelected.BackColor = RGB(Me.Red, Me.Green, Me.Blue)
   Me.RGBLong = PicSelected.BackColor
   lblRGBLong.Caption = PicSelected.BackColor
   DisplayColorLongs RGB(Me.Red, Me.Green, Me.Blue)
   RaiseEvent Click
End Sub

Private Sub sbRed_Change()
   lblRed.Caption = CStr(sbRed.Value)
   Me.Red = sbRed.Value
   PicSelected.BackColor = RGB(Me.Red, Me.Green, Me.Blue)
   Me.RGBLong = PicSelected.BackColor
   lblRGBLong.Caption = PicSelected.BackColor
   DisplayColorLongs RGB(Me.Red, Me.Green, Me.Blue)
   RaiseEvent Click
End Sub

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
   Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
   m_Enabled = New_Enabled
   PropertyChanged "Enabled"
End Property

Public Property Get Red() As Long
   Red = m_Red
End Property

Public Property Let Red(ByVal New_Red As Long)
   m_Red = New_Red
   PropertyChanged "Red"
End Property

Public Property Get Green() As Long
   Green = m_Green
End Property

Public Property Let Green(ByVal New_Green As Long)
   m_Green = New_Green
   PropertyChanged "Green"
End Property

Public Property Get Blue() As Long
   Blue = m_Blue
End Property

Public Property Let Blue(ByVal New_Blue As Long)
   m_Blue = New_Blue
   PropertyChanged "Blue"
End Property

Public Property Get RGBLong() As Long
   RGBLong = m_RGBLong
End Property

Public Property Let RGBLong(ByVal New_RGBLong As Long)
   m_RGBLong = New_RGBLong
   PropertyChanged "RGBLong"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
   m_Red = PropBag.ReadProperty("Red", m_def_Red)
   m_Green = PropBag.ReadProperty("Green", m_def_Green)
   m_Blue = PropBag.ReadProperty("Blue", m_def_Blue)
   m_RGBLong = PropBag.ReadProperty("RGBLong", m_def_RGBLong)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
   Call PropBag.WriteProperty("Red", m_Red, m_def_Red)
   Call PropBag.WriteProperty("Green", m_Green, m_def_Green)
   Call PropBag.WriteProperty("Blue", m_Blue, m_def_Blue)
   Call PropBag.WriteProperty("RGBLong", m_RGBLong, m_def_RGBLong)
End Sub
