VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CellsToHTMLTableTag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function ToString(TableRange As Range) As String
  ToString = getTableBodyHtmlData(TableRange)
End Function


Private Function getTableBodyHtmlData(TableRange As Range) As String
  Dim Ret As String
  
  Dim RowRange As Range
  For Each RowRange In TableRange.Rows
    Ret = Ret & getRowHtmlData(RowRange) & vbCrLf
  Next

  Ret = "<table>" & vbCrLf & Ret & "</table>"

  getTableBodyHtmlData = Ret
End Function

Private Function getRowHtmlData(RowRange As Range) As String
  Dim Ret As String
  
  Dim TargetRange As Range
  For Each TargetRange In RowRange.Columns
    If Not isMergedCellOtherThanTopLeft(TargetRange) Then
      Ret = Ret & "<td" & getAlignAttr(TargetRange) & getSpanAttr(TargetRange) & ">" & TargetRange.Text & "</td>"
    End If
  Next
  
  Ret = "<tr>" & Ret & "</tr>"
  
  getRowHtmlData = Ret
End Function


Private Function getAlignAttr(TargetRange As Range) As String
  Dim Ret As String
  
  With TargetRange
    If .HorizontalAlignment = xlCenter Then
      Ret = "center"
    ElseIf .HorizontalAlignment = xlRight Then
      Ret = "right"
    ElseIf .HorizontalAlignment = xlLeft Then
      Ret = ""
    ElseIf IsNumeric(.Text) Or IsDate(.Text) Then
      Ret = "right"
    End If
  End With

  If Ret <> "" Then
    Ret = " align=" & Chr(34) & Ret & Chr(34)
  End If
  
  getAlignAttr = Ret
End Function


Private Function getSpanAttr(TargetRange As Range) As String
  Dim Ret As String
  
  With TargetRange
    If .MergeCells Then
      Ret = " rowspan=" & Chr(34) & .MergeArea.Rows.Count & Chr(34) & _
            " colspan=" & Chr(34) & .MergeArea.Columns.Count & Chr(34)
    End If
  End With
  
  getSpanAttr = Ret
End Function

Private Function isMergedCellOtherThanTopLeft(TargetRange As Range) As Boolean
  Dim Ret As Boolean
  
  If TargetRange.MergeCells Then
    With TargetRange.MergeArea
      If TargetRange.Row <> .Row Or TargetRange.Column <> .Column Then
        Ret = True
        GoTo Finally
      End If
    End With
  End If
  
  Ret = False
  
Finally:
  isMergedCellOtherThanTopLeft = Ret
End Function

