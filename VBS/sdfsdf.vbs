option Explicit
Dim obj, a, b, c

Set obj = CreateObject("wscript.shell")

a=MsgBox("Open a Program?",vbYesNo+vbQuestion+vbSystemModal)

If a=vbYes Then
  obj.Run "mspaint.exe"
  b=MsgBox("Open a Folder?",vbYesNo+vbQuestion+vbSystemModal)
Else
  b=MsgBox("Open a Folder?",vbYesNo+vbQuestion+vbSystemModal)
End If


If b=vbYes Then
  obj.Run "C:\Users\Jeremy\Downloads\Extracted"
  c=MsgBox("Open a File?",vbYesNo+vbQuestion+vbSystemModal)
Else
  c=MsgBox("Open a File?",vbYesNo+vbQuestion+vbSystemModal)
End If

If c=vbYes Then
  obj.Run """C:\Users\Jeremy\Desktop\other 2.txt"""
Else
  WScript.Quit
End If