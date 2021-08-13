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

Public Sub Output(Text As String)

  Me.Memo.Text = Text
  Me.Repaint
    
End Sub

Public Sub Append(Text As String)
  
  Me.Memo.Text = Me.Memo.Text + Text
  Me.Repaint
  
End Sub

Private Sub cancelBtn_Click()

  ExitApp
    
End Sub

Private Sub excludeTxt_Change()

  SaveStrSetting "exclude", Me.excludeTxt.Text
    
End Sub

Private Sub includeTxt_Change()

  SaveStrSetting "include", Me.includeTxt.Text
    
End Sub

Private Sub pdfChk_Click()

  SaveBoolSetting "pdf", Me.pdfChk.Value
    
End Sub

Private Sub xlsChk_Click()

  SaveBoolSetting "xls", Me.xlsChk.Value
  
End Sub

Private Sub dwgChk_Click()

  SaveBoolSetting "dwg", Me.dwgChk.Value
    
End Sub

Private Sub dxfChk_Click()

  SaveBoolSetting "dxf", Me.dxfChk.Value
    
End Sub

Private Sub runBtn_Click()

  Dim SoughtForExtensions As Dictionary
  
  'Lower case fo extension is required!
  
  Set SoughtForExtensions = New Dictionary
  If Me.pdfChk.Value Then
    SoughtForExtensions.Add "pdf", 0
  End If
  If Me.xlsChk.Value Then
    SoughtForExtensions.Add "xls", 0
  End If
  If Me.dwgChk.Value Then
    SoughtForExtensions.Add "dwg", 0
  End If
  If Me.dxfChk.Value Then
    SoughtForExtensions.Add "dxf", 0
  End If
  If SoughtForExtensions.Count > 0 Then
    Run SoughtForExtensions, Me.targetTxt.Text, Me.excludeTxt.Text, Me.includeTxt.Text
    Me.Memo.SetFocus
  End If
    
End Sub

Private Sub UserForm_Initialize()

  Me.pdfChk.Value = GetBoolSetting("pdf")
  Me.xlsChk.Value = GetBoolSetting("xls")
  Me.dwgChk.Value = GetBoolSetting("dwg")
  Me.dxfChk.Value = GetBoolSetting("dxf")
  Me.excludeTxt.Text = GetStrSetting("exclude")
  Me.includeTxt.Text = GetStrSetting("include")
  targetTxt.Text = GetDefaultTarget
    
End Sub
