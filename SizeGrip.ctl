VERSION 5.00
Begin VB.UserControl SizeGrip 
   ClientHeight    =   345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   345
   ScaleHeight     =   345
   ScaleWidth      =   345
   ToolboxBitmap   =   "SizeGrip.ctx":0000
   Begin VB.Label lblGrip 
      BackStyle       =   0  'Transparent
      Caption         =   "o"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   14.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   50
      TabIndex        =   1
      Top             =   50
      Width           =   375
   End
   Begin VB.Label lblGrip 
      BackStyle       =   0  'Transparent
      Caption         =   "p"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   14.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   50
      TabIndex        =   0
      Top             =   50
      Width           =   375
   End
End
Attribute VB_Name = "SizeGrip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'+  File Description:
'       SizeGrip - Grip Control that simulates the Statusbar Control Grip
'
'   Product Name:
'       SizeGrip.ctl
'
'   Compatability:
'       Windows: 2000, XP
'
'   Software Developed by:
'       Paul R. Territo, Ph.D
'
'       Adapted from the following online article:
'       http://vb.mvps.org/articles/ap199906.pdf
'
'   Legal Copyright & Trademarks (Current Implementation):
'       Copyright © 2004, by Paul R. Territo, Ph.D, All Rights Reserved Worldwide
'       Trademark ™ 2004, by Paul R. Territo, Ph.D, All Rights Reserved Worldwide
'
'   Comments:
'       No claims or warranties are expressed or implied as to accuracy or fitness
'       for use of this software. Paul R. Territo, Ph.D shall not be liable
'       for any incidental or consequential damages suffered by any use of
'       this software.
'
'-  Modification(s) History:
'       17Mar05 - Initial build of the SizeGrip Control
'
'   Force Declarations
Option Explicit

'   API Declarations
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'   API Message Constants
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTBOTTOMRIGHT = 17

Private Sub lblGrip_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    '   Relase any events captured previously
    ReleaseCapture
    '   Send a message that we are resizing the form
    SendMessage UserControl.Parent.hwnd, WM_NCLBUTTONDOWN, HTBOTTOMRIGHT, ByVal 0&
End Sub

Private Sub UserControl_Initialize()
    Dim fnt As New StdFont
    '   Create a new font
    With fnt
        .Name = "Marlett"
        .Bold = False
        .Size = 12
    End With
    '   Set the grip labels with the respective properties
    With lblGrip(0)
        Set Font = fnt
        .AutoSize = True
        .Caption = "o"
        .ForeColor = vb3DHighlight
        .MousePointer = vbSizeNWSE
        .Left = 50
        .Top = 50
        .ZOrder
    End With
    With lblGrip(1)
        Set Font = fnt
        .AutoSize = True
        .Caption = "p"
        .ForeColor = vb3DShadow
        .MousePointer = vbSizeNWSE
        .Left = 50
        .Top = 50
        .ZOrder
    End With
End Sub

Private Sub UserControl_Resize()
    '   Prevent resizing...
    With UserControl
        .Width = 345
        .Height = 345
    End With
End Sub
