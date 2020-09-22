VERSION 5.00
Begin VB.Form frmPleaseWait 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   495
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5430
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   495
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrFlashLabel 
      Interval        =   500
      Left            =   5040
      Top             =   0
   End
   Begin VB.Label lblWait 
      AutoSize        =   -1  'True
      Caption         =   "Please wait while your DLL file compiles..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5205
   End
End
Attribute VB_Name = "frmPleaseWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub tmrFlashLabel_Timer()
If lblWait.ForeColor = &H0 Then
    lblWait.ForeColor = &HC00000
Else
    lblWait.ForeColor = &H0
End If
End Sub
