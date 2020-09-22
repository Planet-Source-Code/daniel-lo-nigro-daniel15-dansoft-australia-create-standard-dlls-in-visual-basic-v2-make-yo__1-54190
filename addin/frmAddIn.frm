VERSION 5.00
Begin VB.Form frmAddIn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Make a DLL in Visual Basic"
   ClientHeight    =   3390
   ClientLeft      =   2175
   ClientTop       =   1935
   ClientWidth     =   6030
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkMakeDLL 
      Caption         =   "&Make a DLL file"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   4455
   End
   Begin VB.ListBox lstExport 
      Height          =   2085
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   2
      Top             =   720
      Width           =   4455
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblDesc 
      Caption         =   "NOTE: Only procedures in Modules can be exported into a DLL file! So, please put all your DLL routines in a Module."
      Height          =   435
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   4395
   End
   Begin VB.Label lblDesc 
      AutoSize        =   -1  'True
      Caption         =   "Select the functions that you want to export in your DLL file:"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   4215
   End
End
Attribute VB_Name = "frmAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Make DLL's In Visual Basic v.2
' (c) 2004 DanSoft Australia
'
' If you like this, please vote!
' (read text files included in ZIP)

Public VBInstance As VBIde.VBE
Attribute VBInstance.VB_VarHelpID = -1
Public Connect As Connect
Dim strDefPath As String
Option Explicit



Private Sub CancelButton_Click()
Unload Me
Connect.Hide
End Sub

Private Sub Form_Load()
Dim objComponent As VBComponent
Dim objMember As Member
Dim strTemp As String
Dim intTemp As Integer
Dim strCurrExports()

ReDim strCurrExports(0)

'find the path for the .def file of the current project
strDefPath = VBInstance.ActiveVBProject.FileName
If strDefPath = "" Then
    MsgBox "Please save your project before choosing what you want to export.", vbInformation, "Make DLLs"
    Unload Me
    Connect.Hide
    Exit Sub
End If
strDefPath = Left$(strDefPath, Len(strDefPath) - 3) & "def"

On Error GoTo nofile
'try to open existing definition file
Open strDefPath For Input As #1
    chkMakeDLL.Value = 1
    Do Until EOF(1)
        Line Input #1, strTemp
        Select Case Left$(Trim(strTemp), 7)
            Case "LIBRARY"
            Case "EXPORTS"
            Case Else
                ReDim Preserve strCurrExports(UBound(strCurrExports) + 1)
                strCurrExports(UBound(strCurrExports)) = Trim$(strTemp)
        End Select
    Loop
Close #1
dontread:
'enumerate the procedures in every module file within
'the current project
For Each objComponent In VBInstance.ActiveVBProject.VBComponents
    If objComponent.Type = vbext_ct_StdModule Then
        For Each objMember In objComponent.CodeModule.Members
            If objMember.Type = vbext_mt_Method Then
                lstExport.AddItem objMember.Name & " (defined in " & objComponent.Name & ")"
                'check if the procedure is mardked to be exported.
                'if so, tick the box next to it.
                For intTemp = 1 To UBound(strCurrExports)
                    If strCurrExports(intTemp) = objMember.Name Then
                        lstExport.Selected(lstExport.ListCount - 1) = True
                    End If
                Next
            End If
        Next
    End If
Next
Exit Sub

nofile:
'file didn't exist, makedll checkbox = 0
chkMakeDLL.Value = 0
Resume dontread
End Sub

Private Sub OKButton_Click()
On Error GoTo errorhandle
Dim intTemp As Integer
Dim strTemp
If chkMakeDLL.Value = 1 Then
    'open the .def file for the project - this says all
    'the exports in the end dll file.
    Open strDefPath For Output As #1
    Print #1, "LIBRARY " & VBInstance.ActiveVBProject.Name
    Print #1, "EXPORTS"
    'go throgh all procs in the list box. If it is
    'ticked, write the name of it into the file
    For intTemp = 0 To lstExport.ListCount - 1
        If lstExport.Selected(intTemp) = True Then
            Print #1, "    " & Split(lstExport.List(intTemp), " ")(0)
        End If
    Next
Else
    On Error Resume Next
    Kill strDefPath
    On Error GoTo errorhandle
End If

endit:
'close any files which are still open
Close
Unload Me
Connect.Hide
Exit Sub

errorhandle:
Select Case MsgBox("An error occured while writing the definition file: " _
            & Err.Description & " (" & Err.Number & ")", _
            vbAbortRetryIgnore + vbCritical, "Error")
    Case vbAbort
        Resume endit
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
End Select
End Sub
