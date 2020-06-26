Attribute VB_Name = "GroupGroups"
'=======================================================================================
' Макрос           : GroupGroups
' Версия           : 2020.06.23
' Автор            : elvin-nsk (me@elvin.nsk.ru)
'=======================================================================================

Option Explicit
Const RELEASE As Boolean = True

Sub Start()

  If RELEASE Then On Error GoTo ErrHandler

  Dim tShape As Shape
  Dim tSrcRange As New ShapeRange
  Dim tRangeToCheck As ShapeRange
  Dim tRangeOverlapped As ShapeRange
  Dim tRangeToRemove As ShapeRange
  Dim tNew As Boolean
  
  If ActivePage.Shapes.Count < 2 Then Exit Sub
  
  BoostStart "Группирование групп", RELEASE
  
  tSrcRange.AddRange ActiveLayer.Shapes.All
  
  Do
    Set tRangeToCheck = New ShapeRange
    Set tRangeOverlapped = New ShapeRange
    Set tRangeToRemove = New ShapeRange
    tRangeToCheck.AddRange tSrcRange
    tRangeOverlapped.Add tSrcRange(1)
    Do
      tNew = False
      For Each tShape In tRangeToCheck
        If tRangeOverlapped.Exists(tShape) = False Then
          If IsOverlapRange(tRangeOverlapped, tShape) Then
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
      tRangeOverlapped.Group
    Else
      tSrcRange.Remove 1
    End If
  Loop Until tSrcRange.Count = 0
  
ExitSub:
  BoostFinish
  Exit Sub

ErrHandler:
  MsgBox "Ошибка: " & Err.Description, vbCritical
  Resume ExitSub

End Sub

'хотя бы с одним шейпом из TestRange
Function IsOverlapRange(TestRange As ShapeRange, TestShape As Shape) As Boolean
  Dim s As Shape
  For Each s In TestRange
    If lib_elvin.IsOverlap(s, TestShape) Then
      IsOverlapRange = True
      Exit Function
    End If
  Next
  IsOverlapRange = False
End Function
