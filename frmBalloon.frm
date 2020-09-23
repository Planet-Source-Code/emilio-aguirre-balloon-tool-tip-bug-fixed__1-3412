VERSION 5.00
Begin VB.Form frmBalloon 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   195
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   15
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   13
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer tmrBalloon 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   360
      Top             =   240
   End
End
Attribute VB_Name = "frmBalloon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------
' ________  Copyright EAguirre (c)1999
'(        ) eaguirre@comtrade.com.mx
'(  ______)
' \/
' BalloonToolTip
'--------------------------------------
Option Explicit

Private Sub tmrBalloon_Timer()
DisplayBalloon
tmrBalloon.Enabled = False
End Sub
