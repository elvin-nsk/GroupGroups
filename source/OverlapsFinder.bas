Attribute VB_Name = "OverlapsFinder"
Option Explicit

'===============================================================================

Private Enum Pass
  First
  Second
End Enum

'===============================================================================

'первый проход - быстрый поиск пересечений по bounding box
Public Function OnePass(ShapeRange As ShapeRange) As Collection
  Set OnePass = CollectOverlaps(ShapeRange, First)
End Function

Public Function TwoPass(ShapeRange As ShapeRange) As Collection
  Set TwoPass = SecondPass(OnePass(ShapeRange))
End Function

'===============================================================================

'второй проход - рассматриваем найденные в первом проходе пересечения
Private Function SecondPass(ColRanges As Collection) As Collection
  Dim tRange As ShapeRange
  Dim tCol As Collection
  Dim i&
  Set SecondPass = New Collection
  For Each tRange In ColRanges
    frm_Progress.UpdateMax ColRanges.Count
    Set tCol = CollectOverlaps(tRange, Second)
    For i = 1 To tCol.Count 'объединяем коллекции
      SecondPass.Add tCol(i)
    Next i
  Next tRange
End Function

'универсальная функция поиска пересечений для первого и второго прохода
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

'хотя бы с одним шейпом из TestRange
Private Function IsOverlapRange(TestRange As ShapeRange, TestShape As Shape, Pass As Pass) As Boolean
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

