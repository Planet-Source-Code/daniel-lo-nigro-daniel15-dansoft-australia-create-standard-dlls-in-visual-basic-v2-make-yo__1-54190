    Create in Visual Basic version 2
   (c) 2004 DanSoft Australia - http://dansoftaus.r8.org/
   [MAXIMIZE NOTEPAD FOR BEST VIEWING EXPERIENCE ;)]

NOTE: If you used the first version of my 'Make DLL's in VB' utility, you need to uninstall it.
      See the instructions at the bottom.

'Visual Basic directory' refers to the folder that Visual Basic is installed in, usually
'c:\Program Files\Microsoft Visual Studio\VB98'

Here are the instructions to install this utility:

0) Unzip the zip file -- Remember to extract paths!
1) Open 'Linker.vbp' (in the 'linker' folder) and compile it.
2) Copy the 'MakeDLL.exe' file (in the 'compiled' folder) to your Visual Basic directory.
3) Open 'MakeDLLAddin.vbp' (in the 'addin' folder) and compile it
4) Go into Visual Basic, and click Add-Ins -> Add-In Manager. There should be an addin listed
   called 'Create DLLs In Visual Basic' (or similar). Make sure both 'Loaded' and 
   'Load On Startup' are ticked.
4a) If the addin wasn't listed, copy 'MakeDLL.DLL' (in the 'compiled' folder) into your 
    Visual Basic directory and restart Visual Basic.
5) Copy all the files in the 'dll project' folder to your Visual Basic Project Templates folder
   (usually C:\Program Files\Microsoft Visual Studio\VB98\template\projects)

Yay! It is now installed! 

* If you want to create a DLL yourself, go into Visual Basic and choose to create a 
  'Standard DLL' project.
* To choose what functions you want to export into your DLL file, click File -> Choose
  DLL Exports...

* To create your DLL, click File -> Make xxxxxxx.dll (this might show as 'Make xxxxxxx.exe', 
  when the save box pops up, just change the extension to .DLL)

Sample DLL: 'TestDLL.vbp' (in the 'test dll' folder)
Sample prog that uses that DLL: 'TestProg.vbp' (in the 'test program' folder)

If you want to try the Sample program, build the Sample DLL into the 'test program' folder
and then run the sample program.


Old Version Uninstall:
To uninstall version 1 of my Make DLLs in VB utility, follow these steps:

1) Go into your Visual basic directory (C:\Program Files\Microsoft Visual Studio\VB98)
2) Delete 'LINK.EXE'
3) Rename 'LINK1.EXE' to 'LINK.EXE'
4) Delete 'MakeDLL.DLL' if you know where you put it (this step can be skipped if you don't 
   know where the file is)
5) Follow installation instructions above to install this version

If you liked this, please vote! (Even if you didn't like it, post comments saying why).
(c) 2004 DanSoft Australia - http://dansoftaus.r8.org/