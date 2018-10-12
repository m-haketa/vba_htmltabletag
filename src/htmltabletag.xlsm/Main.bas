Attribute VB_Name = "Main"
Option Explicit

Sub test()
  
  Dim ClipBoard As Object
  Set ClipBoard = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    
  Dim CT As CellsToHTMLTableTag
  Set CT = New CellsToHTMLTableTag
  
  Call ClipBoard.settext(CT.ToString(Selection))
  Call ClipBoard.putinclipboard
 

End Sub
