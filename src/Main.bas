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
    
    Set gCurrentDoc = swApp.ActiveDoc
    If gCurrentDoc Is Nothing Then Exit Sub
    If gCurrentDoc.GetType <> swDocASSEMBLY Then
        MsgBox "Запускать в сборках!", vbCritical
        Exit Sub
    End If
    
    MainForm.Show
End Sub

Sub Run(SoughtForExtensions As Dictionary, target As String, excludeLines As String)
    Dim components As Dictionary
    Dim searchFolders As Dictionary
    Dim Drawings As Collection
    Dim pattern As RegExp
    Dim NotFound As Dictionary
    Dim copied As Dictionary
    Dim exclude As Collection
    Dim currentDocConf As String
    
    Set components = New Dictionary
    Set searchFolders = New Dictionary
    Set exclude = CreateExclude(excludeLines)
    currentDocConf = gCurrentDoc.ConfigurationManager.ActiveConfiguration.name
    MainForm.Output "Анализ компонентов сборки..."
    AddComponent gCurrentDoc, currentDocConf, components, searchFolders, exclude, "", ""
    ComponentResearch gCurrentDoc, components, searchFolders, exclude, currentDocConf
    
    Set Drawings = New Collection
    Set pattern = CreatePattern(SoughtForExtensions)
    MainForm.Output "Поиск чертежей..."
    CollectAllDrawings pattern, searchFolders, Drawings
    
    MainForm.Output "Сопоставление чертежей компонентам..."
    MatchFiles components, Drawings
    
    Set NotFound = New Dictionary
    Set copied = New Dictionary
    MainForm.Output "Копирование чертежей..."
    
    UniqueCopiedFiles components, copied, NotFound, SoughtForExtensions
    CopyFiles copied, target
    
    MainForm.Output CreateOutput(copied.count, NotFound)
End Sub

Function GetDefaultTarget() As String
    GetDefaultTarget = gFSO.GetParentFolderName(gCurrentDoc.GetPathName) + "\Чертежи в архив"
End Function

Function GetLogFileName() As String
   GetLogFileName = gFSO.GetParentFolderName(gCurrentDoc.GetPathName) + "\Не найдены чертежи.txt"
End Function
