Attribute VB_Name = "GroupGroups"
'===============================================================================
' Макрос           : GroupGroups
' Версия           : 2020.12.27
' Автор            : elvin-nsk (me@elvin.nsk.ru)
'===============================================================================

Option Explicit
Const RELEASE As Boolean = True

'===============================================================================

Sub Start()

  If RELEASE Then On Error GoTo ErrHandler

  Dim tRange As ShapeRange
  
  If ActiveSelectionRange.Count = 1 Then
    MsgBox "Выберите 2 или более объектов"
    Exit Sub
  ElseIf ActiveSelectionRange.Count > 1 Then
    Set tRange = ActiveSelectionRange
  Else
    Set tRange = ActiveLayer.Shapes.All
  End If
  
  lib_elvin.BoostStart "Группирование групп", RELEASE
  frm_Progress.Caption = "Группирование групп"
  
  GroupRanges OverlapsFinder.TwoPass(tRange)
  
ExitSub:
  lib_elvin.BoostFinish
  frm_Progress.Finish
  Exit Sub

ErrHandler:
  MsgBox "Ошибка: " & Err.Description, vbCritical
  Resume ExitSub

End Sub

'===============================================================================

Private Sub GroupRanges(ColRanges As Collection)
  Dim tRange As ShapeRange
  If ColRanges.Count = 0 Then Exit Sub
  For Each tRange In ColRanges
    tRange.Group
  Next tRange
End Sub

