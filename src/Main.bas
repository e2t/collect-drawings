Attribute VB_Name = "Main"
Option Explicit

Public swApp As Object
Public gFSO As FileSystemObject
'Public gDebugFile As TextStream

Dim gCurrentDoc As ModelDoc2

Sub Main()

  Set swApp = Application.SldWorks
  Set gFSO = New FileSystemObject
  'Set gDebugFile = gFSO.CreateTextFile("d:\debug.txt", True)
  Tools.Init
  
  Set gCurrentDoc = swApp.ActiveDoc
  If gCurrentDoc Is Nothing Then Exit Sub
  If gCurrentDoc.GetType <> swDocASSEMBLY Then
    MsgBox "��������� � �������!", vbCritical
    Exit Sub
  End If
  
  MainForm.Show
  'gDebugFile.Close
    
End Sub

Sub Run( _
  SoughtForExtensions As Dictionary, Target As String, ExcludeLines As String, _
  IncludeLines As String)

  Dim Components As Dictionary
  Dim SearchFolders As Dictionary
  Dim Drawings As Dictionary
  Dim Pattern As RegExp
  Dim NotFound As Dictionary
  Dim Copied As Dictionary
  Dim Exclude As Collection
  Dim Include As Collection
  Dim CurrentDocConf As String
  Dim CurrentAsm As AssemblyDoc
  
  Set Components = New Dictionary
  Set SearchFolders = New Dictionary
  Set Exclude = SplitLine(ExcludeLines)
  Set Include = SplitLine(IncludeLines)
  CurrentDocConf = gCurrentDoc.ConfigurationManager.ActiveConfiguration.Name
  
  MainForm.Output "������� ����������� ������..."
  Set CurrentAsm = gCurrentDoc
  CurrentAsm.ResolveAllLightWeightComponents False
  
  MainForm.Append " OK" + vbNewLine + "������ ����������� ������..."
  AddComponent gCurrentDoc, CurrentDocConf, Components, SearchFolders, Exclude, "", ""
  ComponentResearch gCurrentDoc, Components, SearchFolders, Exclude, CurrentDocConf
  AddUserSearchFolders Include, SearchFolders
  
  Set Drawings = New Dictionary
  Set Pattern = CreatePattern(SoughtForExtensions)
  
  MainForm.Append " OK" + vbNewLine + "����� ��������..."
  CollectAllDrawings Pattern, SearchFolders, Drawings
  
  MainForm.Append " OK" + vbNewLine + "������������� �������� �����������..."
  MatchFiles Components, Drawings
  
  Set NotFound = New Dictionary
  Set Copied = New Dictionary
  
  MainForm.Append " OK" + vbNewLine + "����������� ��������..."
  UniqueCopiedFiles Components, Copied, NotFound, SoughtForExtensions
  CopyFiles Copied, Target
  
  MainForm.Output CreateOutput(Copied.Count, NotFound)
    
End Sub

Function GetDefaultTarget() As String

    GetDefaultTarget = gFSO.GetParentFolderName(gCurrentDoc.GetPathName) + "\������� � �����"
    
End Function

Function GetLogFileName() As String

   GetLogFileName = gFSO.GetParentFolderName(gCurrentDoc.GetPathName) + "\�� ������� �������.txt"
   
End Function
