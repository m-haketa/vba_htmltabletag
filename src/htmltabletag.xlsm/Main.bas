Attribute VB_Name = "Main"
Option Explicit

Sub ClipBoardOutput()
  '�I��͈͂�HTML Table�����ăN���b�v�{�[�h�ɃR�s�[
  Dim ClipBoard As Object
  Set ClipBoard = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    
  Dim CT As CellsToHTMLTableTag
  Set CT = New CellsToHTMLTableTag
  
  Call ClipBoard.settext(CT.ToString(Selection))
  Call ClipBoard.putinclipboard
 

End Sub

Sub testDebugPrint()
  '����e�X�g
  Dim CT As CellsToHTMLTableTag
  Set CT = New CellsToHTMLTableTag
  
  '��w�b�_����
  Debug.Print CT.ToString(Range("A7:D11"))
  
  '��w�b�_�Ȃ�
  Debug.Print CT.ToString(Range("A1:C4"), False, False)

End Sub


