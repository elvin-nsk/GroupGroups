VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Progress 
   Caption         =   "ѕодождите..."
   ClientHeight    =   660
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "frm_Progress.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_Progress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'=======================================================================================
' ‘орма            : frm_Progress
' ¬ерси€           : 2020.07.14
' јвторы           : https://www.erlandsendata.no
'                    доработал elvin-nsk (me@elvin.nsk.ru)
' Ќазначение:      : отображение прогресс-бара
'=======================================================================================

Option Explicit

'=======================================================================================
' переменные
'=======================================================================================

Dim Iter#

'=======================================================================================
' публична€ часть
'=======================================================================================

'прогресс в виде дес€тичной дроби (напр. "0.3" = 30%)
Sub UpdateDec(ByVal Dec!)
  Update Dec
End Sub

'прогресс в виде текущей итерации из максимальных (напр. "2, 8" = 25%)
Sub UpdateNum(ByVal Cur#, ByVal Max#)
  If Cur > Max Then Cur = Max
  Update Cur / Max
End Sub

'прогресс в виде автоматически найденной текущей итерации из максимальных
Sub UpdateMax(ByVal Max#)
  UpdateNum Iter, Max
  Iter = Iter + 1
End Sub

'прогресс в виде процентов (напр. "15" = 15%)
Sub UpdatePct(ByVal Pct As Byte)
  If Pct > 100 Then Pct = 100
  Update Pct / 100
End Sub

Sub Finish()
  Unload Me
End Sub

'=======================================================================================
' событи€ и приватные функции
'=======================================================================================

Private Sub UserForm_Initialize()
  Me.Show vbModeless
  With Me.lblDone ' set the "progress bar" to it's initial length
    .Top = Me.lblRemain.Top + 1
    .Left = Me.lblRemain.Left + 1
    .Height = Me.lblRemain.Height - 2
    .Width = 0
  End With
  Iter = 0
End Sub

'отключаем X-unload
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If CloseMode = VbQueryClose.vbFormControlMenu Then
    Cancel = True
    Me.Hide
  End If
End Sub

Private Sub Update(Dec!)
  If Dec < 0 Then Dec = Abs(Dec)
  If Dec > 1 Then Dec = 1
  With Me
    .lblDone.Width = Dec * (.lblRemain.Width - 2)
    .lblPct.Caption = Format(Dec, "0%")
  End With
  DoEvents 'The DoEvents statement is responsible for the form updating
End Sub
