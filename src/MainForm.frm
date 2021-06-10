VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "Собрать чертежи открытой модели"
   ClientHeight    =   10785
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12330
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Sub Output(text As String)

  Me.memo.text = text
  Me.Repaint
    
End Sub

Public Sub Append(text As String)
  
  Me.memo.text = Me.memo.text + text
  Me.Repaint
  
End Sub

Private Sub cancelBtn_Click()
    ExitApp
End Sub

Private Sub excludeTxt_Change()
    SaveStrSetting "exclude", Me.excludeTxt.text
End Sub

Private Sub includeTxt_Change()
    SaveStrSetting "include", Me.includeTxt.text
End Sub

Private Sub pdfChk_Click()
    SaveBoolSetting "pdf", Me.pdfChk.value
End Sub

Private Sub xlsChk_Click()
    SaveBoolSetting "xls", Me.xlsChk.value
End Sub

Private Sub dwgChk_Click()
    SaveBoolSetting "dwg", Me.dwgChk.value
End Sub

Private Sub dxfChk_Click()
    SaveBoolSetting "dxf", Me.dxfChk.value
End Sub

Private Sub runBtn_Click()
    Dim SoughtForExtensions As Dictionary
    
    'Lower case fo extension is required!
    
    Set SoughtForExtensions = New Dictionary
    If Me.pdfChk.value Then
        SoughtForExtensions.Add "pdf", 0
    End If
    If Me.xlsChk.value Then
        SoughtForExtensions.Add "xls", 0
    End If
    If Me.dwgChk.value Then
        SoughtForExtensions.Add "dwg", 0
    End If
    If Me.dxfChk.value Then
        SoughtForExtensions.Add "dxf", 0
    End If
    If SoughtForExtensions.Count > 0 Then
        Run SoughtForExtensions, Me.targetTxt.text, Me.excludeTxt.text, Me.includeTxt.text
        Me.memo.SetFocus
    End If
End Sub

Private Sub UserForm_Initialize()
    Me.pdfChk.value = GetBoolSetting("pdf")
    Me.xlsChk.value = GetBoolSetting("xls")
    Me.dwgChk.value = GetBoolSetting("dwg")
    Me.dxfChk.value = GetBoolSetting("dxf")
    Me.excludeTxt.text = GetStrSetting("exclude")
    Me.includeTxt.text = GetStrSetting("include")
    targetTxt.text = GetDefaultTarget
End Sub
