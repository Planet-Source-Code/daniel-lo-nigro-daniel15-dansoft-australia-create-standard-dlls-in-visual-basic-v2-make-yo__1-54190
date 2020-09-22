Attribute VB_Name = "Module1"
''''''''''''''''''''''''''''''''''''''''''''''''
''    DLL PROJECT ©2004 DanSoft Australia     ''
''   Your dlls MUST HAVE a DLLMain and Main   ''
'' proc, otherwise it won't compile properly! ''
''''''''''''''''''''''''''''''''''''''''''''''''

Function DLLMain(ByVal A As Long, ByVal B As Long, ByVal c As Long) As Long
    DLLMain = 1
End Function

Sub Main()
    'This is a dummy, so the IDE doesn't complain
    'there is no Sub Main.
End Sub

'add more functions here, ie.
'Function addition(ByVal A As Double, ByVal B As Double) As Double
'    addition = A + B
'End Function
