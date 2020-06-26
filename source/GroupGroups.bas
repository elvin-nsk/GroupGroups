Attribute VB_Name = "GroupGroups"
'=======================================================================================
' ������           : GroupGroups
' ������           : 2020.06.26
' �����            : elvin-nsk (me@elvin.nsk.ru)
'=======================================================================================

Option Explicit
Const RELEASE As Boolean = True

'=======================================================================================
' ���������� ������
'=======================================================================================

Enum Pass
  First
  Second
End Enum

'=======================================================================================
' ��������� ���������
'=======================================================================================

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
  
  BoostStart "������������� �����", RELEASE
  
  GroupRanges SecondPass(FirstPass(tRange))
  
ExitSub:
  BoostFinish
  Exit Sub

ErrHandler:
  MsgBox "������: " & Err.Description, vbCritical
  Resume ExitSub

End Sub

'=======================================================================================
' �������
'=======================================================================================

'������ ������ - ������� ����� ����������� �� bounding box
Private Function FirstPass(ShapeRange As ShapeRange) As Collection
  Set FirstPass = CollectOverlaps(ShapeRange, First)
End Function

'������ ������ - ������������� ��������� � ������ ������� �����������
Private Function SecondPass(ColRanges As Collection) As Collection
  Dim tRange As ShapeRange
  Dim tCol As Collection
  Dim i&
  Set SecondPass = New Collection
  For Each tRange In ColRanges
    Set tCol = CollectOverlaps(tRange, Second)
    For i = 1 To tCol.Count '���������� ���������
      SecondPass.Add tCol(i)
    Next i
  Next tRange
End Function

'������������� ������� ������ ����������� ��� ������� � ������� �������
Private Function CollectOverlaps(ShapeRange As ShapeRange, Pass As Pass) As Collection

  Dim tShape As Shape
  Dim tSrcRange As New ShapeRange
  Dim tRangeToCheck As ShapeRange
  Dim tRangeOverlapped As ShapeRange
  Dim tRangeToRemove As ShapeRange
  Dim tNew As Boolean
  
  Set CollectOverlaps = New Collection
  
  tSrcRange.AddRange ShapeRange

  Do
    Set tRangeToCheck = New ShapeRange
    Set tRangeOverlapped = New ShapeRange
    tRangeToCheck.AddRange tSrcRange
    tRangeOverlapped.Add tSrcRange(1)
    Do
      Set tRangeToRemove = New ShapeRange
      tNew = False
      For Each tShape In tRangeToCheck
        If tRangeOverlapped.Exists(tShape) = False Then
          If IsOverlapRange(tRangeOverlapped, tShape, Pass) Then
            tRangeOverlapped.Add tShape
            tRangeToRemove.Add tShape
            tNew = True
          End If
        End If
      Next
      If tNew = False Then Exit Do
      tRangeToCheck.RemoveRange tRangeToRemove
    Loop
    If tRangeOverlapped.Count > 1 Then
      tSrcRange.RemoveRange tRangeOverlapped
      CollectOverlaps.Add tRangeOverlapped
    Else
      tSrcRange.Remove 1
    End If
  Loop Until tSrcRange.Count = 0

End Function

Private Sub GroupRanges(ColRanges As Collection)
  Dim tRange As ShapeRange
  If ColRanges.Count = 0 Then Exit Sub
  For Each tRange In ColRanges
    tRange.Group
  Next tRange
End Sub

'���� �� � ����� ������ �� TestRange
Function IsOverlapRange(TestRange As ShapeRange, TestShape As Shape, Pass As Pass) As Boolean
  Dim tShape As Shape
  For Each tShape In TestRange
    If Pass = First Then
      If lib_elvin.IsOverlapBox(tShape, TestShape) Then
        IsOverlapRange = True
        Exit Function
      End If
    Else
      If lib_elvin.IsOverlap(tShape, TestShape) Then
        IsOverlapRange = True
        Exit Function
      End If
    End If
  Next tShape
  IsOverlapRange = False
End Function
