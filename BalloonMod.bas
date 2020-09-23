Attribute VB_Name = "BalloonMod"
'------------------------------------------------
' ________  Copyright EAguirre (c)1999
'(        ) eaguirre@comtrade.com.mx
'(  ______) Be carefull with subclassing a window
' \/
' BalloonToolTip
'-------------------------------------------------
Option Explicit

Const GWL_WNDPROC = -4
Const HeightCaption = 325 'Twips

Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Dim oldAddress As Long              'Old address of the WndProc
Dim BalloonForm As Form             'Balloon Form Instance
Dim HookedForm As Form              'Hooked Form (subclassing)
Dim BalloonCtrl As Control          'Control under the mouse
Dim TipCtrl As Control              'Tip control
Dim BalloonBox As RECT              'Balloon Box coordinates

Function HiWord(dw As Long) As Long
    If dw And &H80000000 Then
        HiWord = (dw \ 65536) - 1
    Else
        HiWord = dw \ 65536
    End If
End Function

Function LoWord(dw As Long) As Long
    If dw And &H8000& Then
        LoWord = &H8000 Or (dw And &H7FFF&)
    Else
        LoWord = dw And &HFFFF&
    End If
End Function

Public Sub InitProc(ByRef frmParent As Form)
    If frmParent Is Nothing Then Exit Sub
    'Hook the window
    Set HookedForm = frmParent
    'Assign the TipControl
    For Each TipCtrl In HookedForm
      If TypeOf TipCtrl Is BalloonTip Then Exit For
    Next TipCtrl
   'Creating a balloon window
    Set BalloonForm = New frmBalloon
   'Set the new WndProc to the parent form
    oldAddress = SetWindowLong(HookedForm.hwnd, GWL_WNDPROC, AddressOf WndProc)
End Sub

Sub TerminateProc()
  Dim TProc As Long
  If HookedForm Is Nothing Then Exit Sub
  'Restore the old window procedure
  TProc = SetWindowLong(HookedForm.hwnd, GWL_WNDPROC, oldAddress)
  'Restore memory
  Unload BalloonForm
  Set BalloonForm = Nothing
  Set HookedForm = Nothing
  Set TipCtrl = Nothing
End Sub

Function WndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

On Error Resume Next
'Calling the original Window Procedure
WndProc = CallWindowProc(oldAddress, hwnd, uMsg, wParam, lParam)
'Subclassing the original form
Select Case uMsg
    Case WM_SETCURSOR
        Dim wndhWnd As Long
        Dim mouseMsg As Long
        Dim ctrl As Control
              
        ' wParam holds the handle of the window under the cursor
        wndhWnd = wParam
        ' High word of lParam = Mouse Message
        mouseMsg = HiWord(lParam)
           
        If mouseMsg = WM_MOUSEMOVE Then
          If BalloonCtrl.hwnd <> wndhWnd Then
             HideTip
             'Search the control under cursor
             For Each ctrl In HookedForm.Controls
               If ctrl.hwnd <> wndhWnd Then
               'hWnd property not supported or not found yet
               Else
                 If Len(ctrl.ToolTipText) > 0 Then
                   Set BalloonCtrl = ctrl
                   With ctrl
                     TipCtrl.Text = .ToolTipText
                     .ToolTipText = ""
                     'Turn on the timer
                     BalloonForm.Controls(0).Enabled = True
                   End With
                 End If
                 Exit For
               End If
             Next
           End If
        End If
        'Hide Tip in case of mouse click's
        If (mouseMsg = WM_LBUTTONDOWN) Or (mouseMsg = WM_MBUTTONDOWN) _
           Or (mouseMsg = WM_RBUTTONDOWN) Then HideTip
        
    Case WM_HSCROLL, WM_KEYDOWN, WM_KEYUP, WM_VSCROLL
      HideTip
End Select
End Function

Private Sub HideTip()
Dim mCount As Integer
If Not (BalloonCtrl Is Nothing) Then
     With BalloonForm
        'Turn off the timer
        .Controls(0).Enabled = False
        'Hide Balloon Form
        .Hide
     End With
    'Restore Values of the Control
     BalloonCtrl.ToolTipText = TipCtrl.Text
     TipCtrl.Text = ""
     Set BalloonCtrl = Nothing
End If
End Sub

Private Sub ChangeStyle()
Dim Reg(2) As Long
Dim P(3) As POINTAPI
Dim Box As RECT
Dim w As Single, h As Single
'Copy values to variables for optimization
w = BalloonForm.ScaleWidth: h = BalloonForm.ScaleHeight
'Establish the form of the balloon depending the Orientation
'Property.
Select Case TipCtrl.Orientation
    Case North, South
       P(0).x = (w / 2) - (w * 0.15): P(0).y = h / 2
       P(1).x = (w / 2) + (w * 0.15): P(1).y = h / 2
       P(2).x = w / 2
       Box.Left = 0: Box.Right = w
       If TipCtrl.Orientation = North Then
         Box.Top = 0:   Box.Bottom = h - (h * 0.1)
         P(2).y = h
       Else
         Box.Top = h * 0.1: Box.Bottom = h
         P(2).y = 0
       End If
    Case NE, Sw
       P(0).x = (w / 2) - (w * 0.15): P(0).y = (h / 2) - (h * 0.15)
       P(1).x = (w / 2) + (w * 0.15): P(1).y = (h / 2) + (h * 0.15)
       Box.Left = 0: Box.Right = w
       If TipCtrl.Orientation = NE Then
         Box.Top = 0: Box.Bottom = h - (h * 0.1)
         P(2).x = 0: P(2).y = h
       Else
         Box.Top = h * 0.1: Box.Bottom = h
         P(0).x = (w / 2) - (w * 0.15): P(0).y = (h / 2) - (h * 0.15)
         P(1).x = (w / 2) + (w * 0.15): P(1).y = (h / 2) + (h * 0.15)
         P(2).x = w: P(2).y = 0
       End If
    Case East, West
       P(0).x = (w / 2): P(0).y = (h / 2) + (h * 0.15)
       P(1).x = (w / 2): P(1).y = (h / 2) - (h * 0.15)
       P(2).y = h / 2
       Box.Top = 0: Box.Bottom = h
       If TipCtrl.Orientation = East Then
         Box.Left = w * 0.1: Box.Right = w
         P(2).x = 0
       Else
         Box.Left = 0: Box.Right = w - (w * 0.1)
         P(2).x = w
       End If
    Case NW, SE
       P(0).x = (w / 2) - (w * 0.15): P(0).y = (h / 2) + (h * 0.15)
       P(1).x = (w / 2) + (w * 0.15): P(1).y = (h / 2) - (h * 0.15)
       Box.Left = 0: Box.Right = w
       If TipCtrl.Orientation = NW Then
         Box.Top = 0: Box.Bottom = h - (h * 0.1)
         P(2).x = w: P(2).y = h
       Else
         Box.Top = h * 0.1: Box.Bottom = h
         P(2).x = 0: P(2).y = 0
       End If
End Select
'Create Region 1: Balloon Body
Select Case TipCtrl.Style
    Case Rectangle
      Reg(0) = CreateRectRgn(Box.Left, Box.Top, Box.Right, Box.Bottom)
    Case Balloon
      Reg(0) = CreateEllipticRgn(Box.Left, Box.Top, Box.Right, Box.Bottom)
    Case Round_Rectangle
      Reg(0) = CreateRoundRectRgn(Box.Left, Box.Top, Box.Right, Box.Bottom, w * 0.2, h * 0.2)
End Select
'Create Region 2: Tail of the balloon
Reg(1) = CreatePolygonRgn(P(0), 3, 0)
'Combine regions for balloon shape
CombineRgn Reg(1), Reg(1), Reg(0), RGN_OR
'Change the Balloonform shape
SetWindowRgn BalloonForm.hwnd, Reg(1), True
'Adjust de box for fitting the label text
'in the case of elliptic regions
If TipCtrl.Style = Balloon Then
    BalloonBox.Bottom = Box.Bottom - h * 0.15
    BalloonBox.Left = Box.Left + w * 0.15
    BalloonBox.Right = Box.Right - w * 0.15
    BalloonBox.Top = Box.Top + h * 0.15
Else
    BalloonBox.Bottom = Box.Bottom
    BalloonBox.Left = Box.Left
    BalloonBox.Right = Box.Right
    BalloonBox.Top = Box.Top
End If
End Sub

Private Sub DrawLabel()
Dim lngFormat As Long
Dim new_box As RECT
Dim sngArea As Single
Dim oldArea As Single
Dim lngHeight As Long, lngWidth As Long

'Clear control's device context and change display properties
BalloonForm.BackColor = TipCtrl.BackColor
BalloonForm.ForeColor = TipCtrl.ForeColor
Set BalloonForm.Font = TipCtrl.Font
BalloonForm.Cls
'Set the text format
If TipCtrl.WordBreak = yes Then lngFormat = DT_WORDBREAK
If TipCtrl.TextAlign = To_Left Then
    lngFormat = lngFormat Or DT_LEFT
ElseIf TipCtrl.TextAlign = To_Center Then
    lngFormat = lngFormat Or DT_CENTER
Else
    lngFormat = lngFormat Or DT_RIGHT
End If
'Calculate the rectangle
DrawText BalloonForm.hdc, TipCtrl.Text, Len(TipCtrl.Text), new_box, DT_CALCRECT
'Recalculate the balloon size for ensuring that all text will be displayed
sngArea = (new_box.Bottom - new_box.Top) * (new_box.Right - new_box.Left)
sngArea = sngArea * 1.15 'Leave extra space because the wordbreak
oldArea = (BalloonBox.Bottom - BalloonBox.Top) * (BalloonBox.Right - BalloonBox.Left)
If ((sngArea > oldArea) Or (sngArea < (oldArea * 0.65))) And (TipCtrl.AutoSize = yes) Then
   If TipCtrl.WordBreak = yes Then
    'New balloon width has to be twice the height
    lngHeight = CLng(Sqr(sngArea / 3) + 0.5) * 1.5
    lngWidth = 3.75 * CLng(Sqr(sngArea / 3) + 0.5)
  Else
    lngHeight = (new_box.Bottom - new_box.Top) * 1.2
    lngWidth = (new_box.Right - new_box.Left) * 1.2
  End If
  'Add space for the balloon tail
  Select Case TipCtrl.Orientation
    Case North, South, NE, Sw
       lngHeight = lngHeight + (lngHeight * 0.25)
    Case East, West, NW, SE
       lngWidth = lngWidth + (lngWidth * 0.25)
  End Select
  'Add more space in the case of elliptic shape
  If TipCtrl.Style = Balloon Then
   lngHeight = lngHeight + (lngHeight * 0.35)
   lngWidth = lngWidth + (lngWidth * 0.1)
  End If
  'Apply the new values to the Balloon
  'Remember: All calculations are made in pixels so
  'we have to convert it to Twips
  BalloonForm.Width = lngWidth * Screen.TwipsPerPixelX
  BalloonForm.Height = lngHeight * Screen.TwipsPerPixelY
  'Change the style of the Balloon
  ChangeStyle
End If
'Draw text
DrawText BalloonForm.hdc, TipCtrl.Text, Len(TipCtrl.Text), BalloonBox, lngFormat
End Sub

Public Sub DisplayBalloon()
Dim iL As Integer, iT As Integer, iW As Integer, iH As Integer
Dim mCount As Integer

'Avoid Errors
On Error Resume Next
'Copy data for optimization
With BalloonCtrl
    iL = .Left
    iT = .Top
    iW = .Width
    iH = .Height
End With
'Add the Caption Height if necessary
If HookedForm.BorderStyle <> 0 Then iT = iT + HeightCaption
'Calculate AutoSize
DrawLabel
With BalloonForm
  'Place the balloon tip behind the control in the position
  'indicated by the Orientation property
  Select Case TipCtrl.Orientation
    Case East, West
      .Top = HookedForm.Top + iT + (iH / 2) - (BalloonForm.Height / 2)
      If TipCtrl.Orientation = East Then
        .Left = HookedForm.Left + iL + iW
      Else
        .Left = HookedForm.Left + iL - BalloonForm.Width
      End If
    Case Else
      If (TipCtrl.Orientation = South) Or (TipCtrl.Orientation = SE) Or (TipCtrl.Orientation = Sw) Then
        .Top = HookedForm.Top + iT + iH
      Else
       .Top = HookedForm.Top + iT - BalloonForm.Height
      End If
      If (TipCtrl.Orientation = South) Or (TipCtrl.Orientation = North) Then
        .Left = HookedForm.Left + iL + (iW / 2) - (BalloonForm.Width / 2)
      ElseIf (TipCtrl.Orientation = SE) Or (TipCtrl.Orientation = NE) Then
        .Left = HookedForm.Left + iL + iW
      Else
        .Left = HookedForm.Left + iL - BalloonForm.Width
      End If
  End Select
  'Display and Draw
  .Show vbModeless, HookedForm
  'Display Text
  DrawLabel
  HookedForm.SetFocus
End With
End Sub
'----------------------------------------
' ________  Copyright EAguirre (c)1999
'(        ) eaguirre@comtrade.com.mx
'(  ______)
' \/
' BalloonToolTip
'----------------------------------------
