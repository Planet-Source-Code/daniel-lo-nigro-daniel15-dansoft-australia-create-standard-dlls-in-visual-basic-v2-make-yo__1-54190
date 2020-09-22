VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   6105
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   6495
   _ExtentX        =   11456
   _ExtentY        =   10769
   _Version        =   393216
   Description     =   "Now you can make DLL's in Visual Basic!"
   DisplayName     =   "Create DLLs in VB - Version 2"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Make DLL's In Visual Basic v.2
' (c) 2004 DanSoft Australia
'
' If you like this, please vote!
' (read text files included in ZIP)

Option Explicit

Public FormDisplayed          As Boolean
Public VBInstance             As VBIde.VBE
Dim mcbMenu         As Office.CommandBarPopup
Dim mcbMenu2        As Office.CommandBarButton
Dim mfrmAddIn                 As New frmAddIn
Public WithEvents ExportHandler As CommandBarEvents          'command bar event handler
Attribute ExportHandler.VB_VarHelpID = -1
Public WithEvents MakeHandler As CommandBarEvents          'command bar event handler
Attribute MakeHandler.VB_VarHelpID = -1

'open / save dialog stuff
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_HIDEREADONLY = &H4
Private Type OPENFILENAME
    lStructSize As Long
    hWnd As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type


Sub Hide()
    
    On Error Resume Next
    
    FormDisplayed = False
    Unload mfrmAddIn
   
End Sub

Sub Show()
   If VBInstance.ActiveVBProject Is Nothing Then Exit Sub
    On Error Resume Next
    
    If mfrmAddIn Is Nothing Then
        Set mfrmAddIn = New frmAddIn
    End If
    
    Set mfrmAddIn.VBInstance = VBInstance
    Set mfrmAddIn.Connect = Me
    FormDisplayed = True
    mfrmAddIn.Show
   
End Sub

'------------------------------------------------------
'this method adds the Add-In to VB
'------------------------------------------------------
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
Dim mcbTemp As Office.CommandBarButton
On Error GoTo error_handler

'save the vb instance
Set VBInstance = Application

If ConnectMode = ext_cm_External Then
    'Used by the wizard toolbar to start this wizard
    Me.Show
Else
    'The below section will add a 'Make DLL File' menu item
    'under the 'File menu', with 2 subitems: 'choose exports'
    'and 'make dll file'. Use this if you don't like the way
    'this addin changes the 'Make xxxx.EXE' item. To use
    'this, go into 'Project Properties', go to the 'Make' tab
    'and change 'bolSubItems = 0' to 'bolSubItems = 1'
    
    #If bolSubItems = 1 Then
        'create main Make Dll File menu item (in File menu),
        'just under 'Make xxxxxxxx.EXE'
        Set mcbMenu = VBInstance.CommandBars("File").Controls.Add(msoControlPopup, , , 14, True)
        With mcbMenu
            'set caption
            .Caption = "Make DLL File"
            'create "choose exports" subitem...
            Set mcbTemp = .Controls.Add(msoControlButton, , , , True)
            mcbTemp.Caption = "&Choose Exports..."
            '... and sink it to catch all clicks on it
            Set Me.ExportHandler = VBInstance.Events.CommandBarEvents(mcbTemp)
        
            'do same with make dll subitem
            Set mcbTemp = .Controls.Add(msoControlButton, , , , True)
            mcbTemp.Caption = "&Make DLL file..."
            Set Me.MakeHandler = VBInstance.Events.CommandBarEvents(mcbTemp)
        End With
    #Else
        'add 'choose dll exports' onto file menu...
        Set mcbMenu2 = VBInstance.CommandBars("File").Controls.Add(msoControlButton, , , 14, True)
        mcbMenu2.Caption = "&Choose DLL Exports..."
        '... and sink it to catch the clicks on it
        Set Me.ExportHandler = VBInstance.Events.CommandBarEvents(mcbMenu2)

        'sink Make xxxxx.EXE file menu item
        Set Me.MakeHandler = VBInstance.Events.CommandBarEvents(VBInstance.CommandBars("File").Controls(13))
    #End If
End If

If ConnectMode = ext_cm_AfterStartup Then
    If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
        'set this to display the form on connect
        Me.Show
    End If
End If

Exit Sub

error_handler:
Select Case MsgBox("An error occured while initializing the MakeDLL addin: " & Err.Description & " (" & Err.Number & ")", vbExclamation + vbAbortRetryIgnore, "Error")
    Case vbAbort
        Exit Sub
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
End Select
End Sub

'------------------------------------------------------
'this method removes the Add-In from VB
'------------------------------------------------------
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error Resume Next
    
    'delete the command bar entry
    mcbMenu.Delete
    mcbMenu2.Delete
    
    'shut down the Add-In
    If FormDisplayed Then
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "1"
        FormDisplayed = False
    Else
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "0"
    End If
    
    Unload mfrmAddIn
    Set mfrmAddIn = Nothing

End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
    If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
        'set this to display the form on connect
        Me.Show
    End If
End Sub

Private Sub ExportHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
Me.Show
End Sub

Private Sub MakeHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
On Error GoTo ErrorHandler

Dim strDefPath As String
Dim strWhereBuild As String
Dim strVBPath As String
'MsgBox VBInstance.ActiveVBProject.BuildFileName

'if project hasen't been saved, exit because
'we can't create a dll. let vb handle making exe file
strDefPath = VBInstance.ActiveVBProject.FileName
If strDefPath = "" Then
    Exit Sub
End If
strDefPath = Left$(strDefPath, Len(strDefPath) - 3) & "def"

'check if file exists, if not, we can't create dll.
'let vb handle creation of a exe
On Error GoTo nofile
Open strDefPath For Input As #1: Close #1
On Error GoTo 0

'we have to make a dll file, because a def file exists
'
'first, we stop vb handling this (otherwise after we make the
'dll vb will want to build an exe)
CancelDefault = True

'show please wait dialog
frmPleaseWait.Show

'now, we ask where to build the dll file using a save dialog
strWhereBuild = DialogFile(frmPleaseWait.hWnd, 0, "Make DLL File", VBInstance.ActiveVBProject.BuildFileName, "Standard (stdcall) DLL file" & Chr(0) & "*.dll", VBInstance.ActiveVBProject.BuildFileName, "dll")
If strWhereBuild = "" Then Exit Sub
VBInstance.ActiveVBProject.BuildFileName = strWhereBuild

'get the visual basic program path
strVBPath = Left$(VBInstance.FullName, Len(VBInstance.FullName) - 7)

'now, we actually make the DLL!

'rename our replacement linker so it gets run
Name strVBPath & "link.exe" As strVBPath & "link1.exe"
Name strVBPath & "makedll.exe" As strVBPath & "link.exe"

'write the .def file name to a temp file in the
'vb directory so the compiler knows where the
'def file is.
Open strVBPath & "makedll.txt" For Output As #1
    Print #1, strDefPath
Close #1

'actually compile the project, and wait for the
'compile to finish
'If SuperShell(strVBPath & "vb6.exe """ & VBInstance.ActiveVBProject.FileName & """ /make", strVBPath, 0, SW_NORMAL, HIGH_PRIORITY_CLASS) = False Then MsgBox "Error compiling project!", vbExclamation, "Error"
VBInstance.ActiveVBProject.MakeCompiledFile




'rename linker back to normal
Name strVBPath & "link.exe" As strVBPath & "makedll.exe"
Name strVBPath & "link1.exe" As strVBPath & "link.exe"
Unload frmPleaseWait
Exit Sub

'no .def file, let vb make an .exe of it!
nofile:
Exit Sub

ErrorHandler:
Select Case MsgBox("An error occured while making the DLL file: " & Err.Description & " (" & Err.Number & ")", vbExclamation + vbAbortRetryIgnore, "Error")
    Case vbAbort
        Exit Sub
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
End Select
End Sub


'open / save file wrapper
'i got this from somewhere but can't remember where...
Public Function DialogFile(hWnd As Long, wMode As Integer, szDialogTitle As String, szFilename As String, szFilter As String, szDefDir As String, szDefExt As String) As String
    Dim x As Long, OFN As OPENFILENAME, szFile As String, szFileTitle As String
    OFN.lStructSize = Len(OFN)
    OFN.hWnd = hWnd
    OFN.lpstrTitle = szDialogTitle
    OFN.lpstrFile = szFilename & String$(250 - Len(szFilename), 0)
    OFN.nMaxFile = 255
    OFN.lpstrFileTitle = String$(255, 0)
    OFN.nMaxFileTitle = 255
    OFN.lpstrFilter = szFilter
    OFN.nFilterIndex = 1
    OFN.lpstrInitialDir = szDefDir
    OFN.lpstrDefExt = szDefExt

   If wMode = 1 Then
        OFN.Flags = OFN_HIDEREADONLY Or OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST
        x = GetOpenFileName(OFN)
    Else
        OFN.Flags = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_PATHMUSTEXIST
        x = GetSaveFileName(OFN)
    End If

    If x <> 0 Then
        '// If InStr(OFN.lpstrFileTitle, Chr$(0)) > 0 Then
        '//     szFileTitle = Left$(OFN.lpstrFileTitle, InStr(OFN.lpstrFileTitle, Chr$(0)) - 1)
        '// End If
        If InStr(OFN.lpstrFile, Chr$(0)) > 0 Then
            szFile = Left$(OFN.lpstrFile, InStr(OFN.lpstrFile, Chr$(0)) - 1)
        End If
        '// OFN.nFileOffset is the number of characters from the beginning of the
        '// full path to the start of the file name
        '// OFN.nFileExtension is the number of characters from the beginning of the
        '// full path to the file's extention, including the (.)
        '// MsgBox "File Name is " & szFileTitle & Chr$(13) & Chr$(10) & "Full path and file is " & szFile, , "Open"
        '// DialogFile = szFile & "|" & szFileTitle
        DialogFile = szFile
    Else
        DialogFile = ""
    End If
End Function

