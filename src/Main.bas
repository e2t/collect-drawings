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

Sub Run(SoughtForExtensions As Dictionary, target As String, excludeLines As String, _
        includeLines As String)

    Dim components As Dictionary
    Dim searchFolders As Dictionary
    Dim Drawings As Dictionary
    Dim pattern As RegExp
    Dim NotFound As Dictionary
    Dim copied As Dictionary
    Dim exclude As Collection
    Dim include As Collection
    Dim currentDocConf As String
    Dim currentAsm As AssemblyDoc
    
    Set components = New Dictionary
    Set searchFolders = New Dictionary
    Set exclude = SplitLine(excludeLines)
    Set include = SplitLine(includeLines)
    currentDocConf = gCurrentDoc.ConfigurationManager.ActiveConfiguration.Name
    
    MainForm.Output "������� ����������� ������..."
    Set currentAsm = gCurrentDoc
    currentAsm.ResolveAllLightWeightComponents False
    
    MainForm.Append " OK" + vbNewLine + "������ ����������� ������..."
    AddComponent gCurrentDoc, currentDocConf, components, searchFolders, exclude, "", ""
    ComponentResearch gCurrentDoc, components, searchFolders, exclude, currentDocConf
    AddUserSearchFolders include, searchFolders
    
    Set Drawings = New Dictionary
    Set pattern = CreatePattern(SoughtForExtensions)
    
    MainForm.Append " OK" + vbNewLine + "����� ��������..."
    CollectAllDrawings pattern, searchFolders, Drawings
    
    MainForm.Append " OK" + vbNewLine + "������������� �������� �����������..."
    MatchFiles components, Drawings
    
    Set NotFound = New Dictionary
    Set copied = New Dictionary
    
    MainForm.Append " OK" + vbNewLine + "����������� ��������..."
    UniqueCopiedFiles components, copied, NotFound, SoughtForExtensions
    CopyFiles copied, target
    
    MainForm.Output CreateOutput(copied.Count, NotFound)
    
End Sub

Function GetDefaultTarget() As String

    GetDefaultTarget = gFSO.GetParentFolderName(gCurrentDoc.GetPathName) + "\������� � �����"
    
End Function

Function GetLogFileName() As String

   GetLogFileName = gFSO.GetParentFolderName(gCurrentDoc.GetPathName) + "\�� ������� �������.txt"
   
End Function
