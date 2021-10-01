Attribute VB_Name = "GroupGroups"
'===============================================================================
' ������           : GroupGroups
' ������           : 2020.12.27
' �����            : elvin-nsk (me@elvin.nsk.ru)
'===============================================================================

Option Explicit
Const RELEASE As Boolean = True

'===============================================================================

Sub Start()

  If RELEASE Then On Error GoTo ErrHandler

  Dim tRange As ShapeRange
  
  If ActiveSelectionRange.Count = 1 Then
    MsgBox "�������� 2 ��� ����� ��������"
    Exit Sub
  ElseIf ActiveSelectionRange.Count > 1 Then
    Set tRange = ActiveSelectionRange
  Else
    Set tRange = ActiveLayer.Shapes.All
  End If
  
  lib_elvin.BoostStart "������������� �����", RELEASE
  frm_Progress.Caption = "������������� �����"
  
  GroupRanges OverlapsFinder.TwoPass(tRange)
  
ExitSub:
  lib_elvin.BoostFinish
  frm_Progress.Finish
  Exit Sub

ErrHandler:
  MsgBox "������: " & Err.Description, vbCritical
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

