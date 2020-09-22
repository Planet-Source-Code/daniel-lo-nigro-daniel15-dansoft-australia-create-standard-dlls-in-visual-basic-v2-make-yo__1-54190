Attribute VB_Name = "modLinker"
''''''''''''''''''''''''''''''''''''''''''''''''''
'' VBLinker - A replacement Visual Basic linker ''
''          ©2004 DanSoft Australia.            ''
''                                              ''
'' This is part of MakeDLL. To use this linker, ''
'' compile this into your VB folder (c:\prog... ''
'' ...ram files\microsoft visual studio\vb98)   ''
'' and call it 'MAKEDLL.EXE'                    ''
''''''''''''''''''''''''''''''''''''''''''''''''''

'!!! WARNING !!!'
'IF YOU USED VERSION 1 OF MY DLL-MAKER, YOU MUST
'UNINSTALL IT (REFER TO README.TXT)
'!!! WARNING !!!'

'If you like this, PLEASE vote on Planet Source Code
' (look at the text files included in the ZIP file)

Option Explicit

Sub Main()
On Error GoTo ErrorHandler

Dim strCmdLine As String
Dim strDefFile As String
Dim intTemp As Integer

On Error GoTo nodefs
'open makedll.txt, where the def filename is stored
Open App.Path & "\makedll.txt" For Input As #1
    Line Input #1, strDefFile
Close #1
'see if def file exists
Open strDefFile For Input As #1: Close #1
On Error GoTo ErrorHandler

'make new command line, so it compiles as a .dll file
strCmdLine = Command()
strCmdLine = Replace(strCmdLine, "/ENTRY:__vbaS", "/ENTRY:DLLMain")
strCmdLine = Replace(strCmdLine, "/BASE:0x400000", "/BASE:0x10000000")
strCmdLine = strCmdLine & " /DLL /DEF:""" & strDefFile & """"
SuperShell App.Path & "\link1.exe " & strCmdLine, App.Path, 0, SW_HIDE, HIGH_PRIORITY_CLASS

nodefs:
End

ErrorHandler:
Select Case ShowErrorMsg(Err)
    Case "abort"
        End
    Case "retry"
        Resume
    Case "ignore"
        Resume Next
End Select
End Sub

Function ShowErrorMsg(errError As ErrObject) As String
Dim frmMsg As frmError
Set frmMsg = New frmError

frmMsg.lblError.Caption = "An error occured while linking your DLL file: " & errError.Description & " (" & errError.Number & ")"
frmMsg.Show vbModal
ShowErrorMsg = frmMsg.Tag
End Function
