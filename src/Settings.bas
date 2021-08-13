Attribute VB_Name = "Settings"
Option Explicit

Const MacroName = "CollectDrawings"
Const MacroSection = "Main"

Sub SaveStrSetting(Key As String, Value As String)

  SaveSetting MacroName, MacroSection, Key, Value
    
End Sub

Sub SaveIntSetting(Key As String, Value As Integer)

  SaveStrSetting Key, Str(Value)
    
End Sub

Sub SaveBoolSetting(Key As String, Value As Boolean)

  SaveStrSetting Key, BoolToStr(Value)
    
End Sub

Function GetStrSetting(Key As String, Optional Default As String = "") As String

  GetStrSetting = GetSetting(MacroName, MacroSection, Key, Default)
    
End Function

Function GetBoolSetting(Key As String) As Boolean

  GetBoolSetting = StrToBool(GetStrSetting(Key, "0"))
    
End Function

Function GetIntSetting(Key As String) As Integer

  GetIntSetting = StrToInt(GetStrSetting(Key, "0"))
    
End Function

Function StrToInt(Value As String) As Integer

  If IsNumeric(Value) Then
    StrToInt = CInt(Value)
  Else
    StrToInt = 0
  End If
  
End Function

Function StrToBool(Value As String) As Boolean

  If IsNumeric(Value) Then
    StrToBool = CInt(Value)
  Else
    StrToBool = False
  End If
  
End Function

Function BoolToStr(Value As Boolean) As String

  BoolToStr = Str(CInt(Value))
    
End Function
