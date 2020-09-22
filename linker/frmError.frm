VERSION 5.00
Begin VB.Form frmError 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Linker"
   ClientHeight    =   1185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdIgnore 
      Caption         =   "Ignore"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdRetry 
      Caption         =   "&Retry"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdAbort 
      Caption         =   "&Abort"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label lblError 
      Height          =   555
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4305
   End
End
Attribute VB_Name = "frmError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAbort_Click()
Me.Tag = "abort"
Me.Hide
End Sub

Private Sub cmdIgnore_Click()
Me.Tag = "ignore"
Me.Hide
End Sub

Private Sub cmdRetry_Click()
Me.Tag = "retry"
Me.Hide
End Sub
