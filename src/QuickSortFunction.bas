Attribute VB_Name = "QuickSortFunction"
Option Explicit

'https://stackoverflow.com/questions/152319/vba-array-sort-function
Public Sub QuickSort(ByRef vArray As Variant, InLow As Long, InHi As Long)

  Dim Pivot   As Variant
  Dim TmpSwap As Variant
  Dim TmpLow  As Long
  Dim TmpHi   As Long
  
  TmpLow = InLow
  TmpHi = InHi
  
  Pivot = vArray((InLow + InHi) \ 2)
  
  While (TmpLow <= TmpHi)
    While (vArray(TmpLow) < Pivot And TmpLow < InHi)
      TmpLow = TmpLow + 1
    Wend
    
    While (Pivot < vArray(TmpHi) And TmpHi > InLow)
      TmpHi = TmpHi - 1
    Wend
    
    If (TmpLow <= TmpHi) Then
      TmpSwap = vArray(TmpLow)
      vArray(TmpLow) = vArray(TmpHi)
      vArray(TmpHi) = TmpSwap
      TmpLow = TmpLow + 1
      TmpHi = TmpHi - 1
    End If
  Wend
  
  If (InLow < TmpHi) Then QuickSort vArray, InLow, TmpHi
  If (TmpLow < InHi) Then QuickSort vArray, TmpLow, InHi
    
End Sub
