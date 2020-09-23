VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Balloon Tool Tip Testing"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   8400
   StartUpPosition =   2  'CenterScreen
   Begin Project1.BalloonTip BalloonTip1 
      Left            =   3360
      Top             =   1200
      _ExtentX        =   423
      _ExtentY        =   423
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   495
      Left            =   5400
      TabIndex        =   13
      ToolTipText     =   "Exit Tool Tip Testing"
      Top             =   4080
      Width           =   2775
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   2
      Left            =   2160
      TabIndex        =   12
      ToolTipText     =   "Check Control"
      Top             =   960
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   11
      ToolTipText     =   "Check Control"
      Top             =   600
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Frame Control"
      Top             =   120
      Width           =   1935
   End
   Begin VB.DirListBox Dir1 
      Height          =   765
      Left            =   2280
      TabIndex        =   10
      ToolTipText     =   "Dir Control"
      Top             =   3240
      Width           =   3015
   End
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   2280
      ScaleHeight     =   795
      ScaleWidth      =   5955
      TabIndex        =   9
      ToolTipText     =   "Remember that Balloon Tool Tip can automatically adjust its size. Keep all at sight!."
      Top             =   1680
      Width           =   6015
   End
   Begin VB.FileListBox File1 
      Height          =   1260
      Left            =   5400
      TabIndex        =   8
      ToolTipText     =   "File Control "
      Top             =   2640
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   4800
      TabIndex        =   7
      Text            =   "Combo1"
      ToolTipText     =   "Combo Control"
      Top             =   960
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2280
      TabIndex        =   6
      Text            =   "Text1"
      ToolTipText     =   "Text Control"
      Top             =   2640
      Width           =   3015
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   255
      Left            =   3360
      TabIndex        =   5
      ToolTipText     =   "Option Control"
      Top             =   240
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   4
      ToolTipText     =   "Check Control"
      Top             =   240
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "List Control"
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   4560
      TabIndex        =   1
      ToolTipText     =   "Command Control"
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   4200
      TabIndex        =   2
      ToolTipText     =   "label Control"
      Top             =   960
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
l = Len(Command1.Caption)
Command1.Caption = Right(Command1.Caption, l - 1)
End Sub

Private Sub Command2_Click()
Unload Me
Set Form1 = Nothing
End Sub


