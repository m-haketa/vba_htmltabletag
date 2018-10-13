Attribute VB_Name = "Main"
Option Explicit

Sub ClipBoardOutput()
  '選択範囲をHTML Table化してクリップボードにコピー
  Dim ClipBoard As Object
  Set ClipBoard = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    
  Dim CT As CellsToHTMLTableTag
  Set CT = New CellsToHTMLTableTag
  
  Call ClipBoard.settext(CT.ToString(Selection))
  Call ClipBoard.putinclipboard
 

End Sub

Sub testDebugPrint()

  '動作テスト
  Dim CT As CellsToHTMLTableTag
  Set CT = New CellsToHTMLTableTag
  
  '列ヘッダあり
  Debug.Print CT.ToString(Range("A7:D11"))
  
  '列ヘッダなし
  Debug.Print CT.ToString(Range("A1:C4"), False, False)

End Sub


Sub createTestData()
 
  Range("A1").Value = 1
  Range("A1").HorizontalAlignment = xlCenter
  Range("A1:B3").Merge
  
  Range("C1").Value = "B"
  Range("C1").HorizontalAlignment = xlRight
  
  Range("C2").Value = 6
  Range("C3").Value = "A"
  Range("A4").Value = 10
  Range("B4").Value = CDate("2014/6/1")
  Range("C4").Value = 12
  
  Range("A1:C4").Borders.LineStyle = xlContinuous
  
  
  Range("B7") = "a"
  Range("C7") = "b"
  Range("D7") = "c"
  
  Range("A8") = "d"
  Range("A9") = "e"
  Range("A10") = "f"
  Range("A11") = "g"
  
  Range("A1:C4").Copy Range("B8:D11")
  Range("A7:D11").Borders.LineStyle = xlContinuous

End Sub
