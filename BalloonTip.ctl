VERSION 5.00
Begin VB.UserControl BalloonTip 
   ClientHeight    =   240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   225
   EditAtDesignTime=   -1  'True
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   16
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   15
   ToolboxBitmap   =   "BalloonTip.ctx":0000
   Begin VB.Timer tmrCtrl 
      Interval        =   1000
      Left            =   360
      Top             =   0
   End
End
Attribute VB_Name = "BalloonTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------
' ________  Copyright EAguirre (c)1999
'(        ) eaguirre@comtrade.com.mx
'(  ______)
' \/
' BalloonToolTip
'----------------------------------------
Option Explicit
'User Defined Enumerators
Enum WordBoolValue
    No = 0
    yes = 1
End Enum

Enum TextAlignValue
    To_Left = 0
    To_Center = 1
    To_Right = 2
End Enum

Enum StyleValue
    Rectangle = 0
    Balloon = 1
    Round_Rectangle = 2
End Enum

Enum OrientationValues
    North = 0
    NE = 1
    East = 2
    SE = 3
    South = 4
    Sw = 5
    West = 6
    NW = 7
End Enum

'Default Property Values:
Const m_def_AutoSize = yes
Const m_def_TextAlign = To_Left
Const m_def_WordBreak = yes
Const m_def_Orientation = NE
Const m_def_BackColor = &HFFFF&
Const m_def_ForeColor = 0
Const m_def_Text = " "
Const m_def_Style = Balloon

'Property Variables:
Dim m_AutoSize As WordBoolValue
Dim m_TextAlign As TextAlignValue
Dim m_WordBreak As WordBoolValue
Dim m_Orientation As OrientationValues
Dim m_BackColor As OLE_COLOR
Dim m_ForeColor As OLE_COLOR
Dim m_Font As Font
Dim m_Text As String
Dim m_Style As Variant
Dim m_init As Boolean

'Properties
Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property
'
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
    UserControl.BackColor = m_BackColor
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
    UserControl.ForeColor = m_ForeColor
End Property

Public Property Get Font() As Font
Attribute Font.VB_UserMemId = -512
    Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set m_Font = New_Font
    PropertyChanged "Font"
    Set UserControl.Font = m_Font
End Property

Public Property Get Text() As String
    Text = m_Text
End Property

Public Property Let Text(ByVal New_Text As String)
    m_Text = New_Text
    PropertyChanged "Text"
   
End Property

Public Property Get Style() As StyleValue
    Style = m_Style
End Property

Public Property Let Style(ByVal new_Style As StyleValue)
    m_Style = new_Style
    PropertyChanged "Style"
End Property

Public Property Get Orientation() As OrientationValues
    Orientation = m_Orientation
End Property

Public Property Let Orientation(ByVal New_Orientation As OrientationValues)
    m_Orientation = New_Orientation
    PropertyChanged "Orientation"
End Property

Public Property Get TextAlign() As TextAlignValue
    TextAlign = m_TextAlign
End Property

Public Property Let TextAlign(ByVal New_TextAlign As TextAlignValue)
    m_TextAlign = New_TextAlign
    PropertyChanged "TextAlign"
End Property

Public Property Get WordBreak() As WordBoolValue
    WordBreak = m_WordBreak
End Property

Public Property Let WordBreak(ByVal New_WordBreak As WordBoolValue)
    m_WordBreak = New_WordBreak
    PropertyChanged "WordBreak"
End Property

Public Property Get AutoSize() As WordBoolValue
    AutoSize = m_AutoSize
End Property

Public Property Let AutoSize(ByVal New_AutoSize As WordBoolValue)
    m_AutoSize = New_AutoSize
    PropertyChanged "AutoSize"
End Property

Private Sub tmrCtrl_Timer()
'Initialize the control, because when the initialize method
'fires before the control is loaded
  If Not (UserControl.Parent Is Nothing) Then
    'Init Procedure (Hooking the window)
    InitProc UserControl.Parent
    tmrCtrl.Interval = 0
    tmrCtrl.Enabled = False
  End If
End Sub

Private Sub UserControl_Resize()
'Keep short
Width = 240
Height = 240
End Sub

Private Sub UserControl_Terminate()
    TerminateProc
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
    Set m_Font = Ambient.Font
    m_Text = m_def_Text
    m_Style = m_def_Style
    m_Orientation = m_def_Orientation
    m_TextAlign = m_def_TextAlign
    m_WordBreak = m_def_WordBreak
    m_AutoSize = m_def_AutoSize
    m_init = False
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_Text = PropBag.ReadProperty("Text", m_def_Text)
    m_Style = PropBag.ReadProperty("Style", m_def_Style)
    m_Orientation = PropBag.ReadProperty("Orientation", m_def_Orientation)
    m_TextAlign = PropBag.ReadProperty("TextAlign", m_def_TextAlign)
    m_WordBreak = PropBag.ReadProperty("WordBreak", m_def_WordBreak)
    m_AutoSize = PropBag.ReadProperty("AutoSize", m_def_AutoSize)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("Text", m_Text, m_def_Text)
    Call PropBag.WriteProperty("Style", m_Style, m_def_Style)
    Call PropBag.WriteProperty("Orientation", m_Orientation, m_def_Orientation)
    Call PropBag.WriteProperty("TextAlign", m_TextAlign, m_def_TextAlign)
    Call PropBag.WriteProperty("WordBreak", m_WordBreak, m_def_WordBreak)
    Call PropBag.WriteProperty("AutoSize", m_AutoSize, m_def_AutoSize)
End Sub

