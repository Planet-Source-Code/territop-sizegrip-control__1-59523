VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "SizeGrip Control Example"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin Project1.SizeGrip SizeGrip1 
      Height          =   345
      Left            =   4320
      TabIndex        =   0
      Top             =   2760
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   609
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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
'       17Mar05 - Initial test harness for SizeGrip Control
'
'   Force Declarations
Option Explicit

Private Sub Form_Resize()
    '   Move the Control to the lower right corner of the form...
    Me.SizeGrip1.Move Me.ScaleWidth - SizeGrip1.Width + 10, _
        Me.ScaleHeight - SizeGrip1.Height + 10
End Sub

