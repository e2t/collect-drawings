Attribute VB_Name = "Tools"
Option Explicit

Const gPropertyDesignation = "Обозначение"
Const gPropertyName = "Наименование"
Const gPropertyBlank = "Заготовка"

Dim gRegex As RegExp

Function Init()

  Set gRegex = New RegExp
  gRegex.Global = True
  gRegex.IgnoreCase = True
   
End Function

Sub MatchFiles(ByRef Components As Dictionary, ByRef Drawings As Dictionary)

  Dim Ci_ As Variant
  Dim Ci As ComponentInfo
  Dim Drw_ As Variant
  Dim Drw As String
  Dim Key As String
  
  For Each Ci_ In Components.Items
    Set Ci = Ci_
    Key = Ci.Doc.GetPathName
    
    For Each Drw_ In Drawings.Keys
      Drw = Drw_
      Select Case Drawings(Drw)
        Case ""
          If CheckMatchFileAndComponent(Drw, Ci) Then
            Drawings(Drw) = Key
          End If
        Case Key
          CheckMatchFileAndComponent Drw, Ci
      End Select
    Next
  Next
   
End Sub

Function HaveThisDrawingFormat(Drawings As Dictionary, Extension As String) As Boolean

  Dim I As Variant
  
  HaveThisDrawingFormat = False
  For Each I In Drawings.Keys
    If LCase(I) = LCase(Extension) Then
      HaveThisDrawingFormat = True
      Exit For
    End If
  Next
   
End Function

Sub UniqueCopiedFiles( _
  Components As Dictionary, ByRef Copied As Dictionary, _
  ByRef NotFound As Dictionary, SoughtForExtensions As Dictionary)
                      
  Dim Ci_ As Variant
  Dim Ci As ComponentInfo
  Dim IsSheetDetail As RegExp
  Dim IsPipeDetail As RegExp
  
  Set IsSheetDetail = New RegExp
  IsSheetDetail.IgnoreCase = True
  IsSheetDetail.Pattern = ".*(лист|sheet).*"
  
  Set IsPipeDetail = New RegExp
  IsPipeDetail.IgnoreCase = True
  IsPipeDetail.Pattern = ".*(труба|pipe|tube).*"
  
  'Lower case fo extension is required!

  NotFound.Add "pdf", New Dictionary
  NotFound.Add "dwg|dxf", New Dictionary
  NotFound.Add "xls", New Dictionary

  For Each Ci_ In Components.Items
    Set Ci = Ci_
    If SoughtForExtensions.Exists("pdf") Then
      AssignItem Ci, "pdf", Copied, NotFound 'PDF - Для всех.
    End If
    If SoughtForExtensions.Exists("dwg") Or SoughtForExtensions.Exists("dxf") Then
      If IsSheetDetail.Test(Ci.PropertyBlank) Then 'DWG/DXF - Для листовых деталей.
        AssignItem Ci, "dwg|dxf", Copied, NotFound
      End If
    End If
    If SoughtForExtensions.Exists("xls") Then
      If Ci.Doc.GetType = swDocASSEMBLY Then 'XLS - Для сборок.
        AssignItem Ci, "xls", Copied, NotFound
      End If
    End If
  Next
   
End Sub

Sub AssignItem( _
  Ci As ComponentInfo, VerticalLineSeparatedExtensions As String, _
  ByRef Copied As Dictionary, ByRef NotFound As Dictionary)
               
  Dim Rec As DrawingRecord
  Dim Extension As Variant
  Dim Found As Boolean
  
  Found = False
  For Each Extension In Split(VerticalLineSeparatedExtensions, "|")
    If Ci.Drawings.Exists(Extension) Then
      Set Rec = Ci.Drawings(Extension)
      AddUniqueItemInDict Rec.Path, Ci.Parent, Copied
      Found = True
    End If
  Next
  If Not Found Then
    AddUniqueItemInDict Ci, 0, NotFound(VerticalLineSeparatedExtensions)
  End If
   
End Sub

Sub CopyFiles(Copied As Dictionary, Target As String)

  Dim F_ As Variant
  Dim F As String
  Dim NewName As String
  Dim FolderName As String
  
  If Not gFSO.FolderExists(Target) Then
    gFSO.CreateFolder Target
  End If
  
  For Each F_ In Copied.Keys
    F = F_
    FolderName = Target + "\" + Trim(Copied(F))
    If Not gFSO.FolderExists(FolderName) Then
      gFSO.CreateFolder FolderName
    End If
    
    NewName = FolderName + "\" + gFSO.GetFileName(F)
    
    If LCase(F) <> LCase(NewName) Then
      gFSO.CopyFile F, NewName
    End If
  Next
    
End Sub

Function CreateOutput(foundCount As Integer, NotFound As Dictionary) As String

  Dim Ci_ As Variant
  Dim Ci As ComponentInfo
  Dim Text As String
  Dim NotFoundArray() As String
  Dim I As Long
  Dim Extension As Variant
  Dim NotFoundSomeFormat As Dictionary
  Dim FirstTime As Boolean
  
  Text = "Найдено чертежей: " + Str(foundCount)
  FirstTime = True
  
  For Each Extension In NotFound
    Set NotFoundSomeFormat = NotFound(Extension)
    If NotFoundSomeFormat.Count > 0 Then
      ReDim NotFoundArray(NotFoundSomeFormat.Count - 1)
      Text = Text + vbNewLine + vbNewLine + "Не найдены чертежи " + UCase(Extension) + " для:"
      I = -1
      For Each Ci_ In NotFoundSomeFormat.Keys
        Set Ci = Ci_
        I = I + 1
        NotFoundArray(I) = Trim(Ci.PropertyDesignation + " " + Ci.PropertyName) + _
                           " [файл: " + gFSO.GetFileName(Ci.Doc.GetPathName) + " @ " + Ci.Conf + _
                           " ]"
      Next
      QuickSort NotFoundArray, 0, I
      LogNotFoundArray NotFoundArray, FirstTime
      Text = Text + vbNewLine + Join(NotFoundArray, vbNewLine)
    End If
  Next
  CreateOutput = Text
   
End Function

Sub LogNotFoundArray(NotFoundArray() As String, ByRef FirstTime As Boolean)

  Dim FileStream As TextStream
  Dim I As Variant
  Dim Mode As IOMode

  If FirstTime Then
    Mode = ForWriting
    FirstTime = False
  Else
    Mode = ForAppending
  End If
  Set FileStream = gFSO.OpenTextFile(GetLogFileName, Mode, True)
  If Mode = ForAppending Then
    FileStream.WriteBlankLines 1
  End If
  For Each I In NotFoundArray
    FileStream.WriteLine I
  Next
  FileStream.Close
   
End Sub

'Функция-сокращение для любых словарей.
'True - если был создан новый объект, False - если объект существовал.
Function AddUniqueItemInDict(Key As Variant, Item As Variant, ByRef Dict As Dictionary) As Boolean

  AddUniqueItemInDict = Not Dict.Exists(Key)
  If AddUniqueItemInDict Then
    Dict.Add Key, Item
  End If
    
End Function

Function IsArrayEmpty(ByRef anArray As Variant) As Boolean

  IsArrayEmpty = True
  On Error Resume Next
  IsArrayEmpty = LBound(anArray) > UBound(anArray)

End Function

Sub ComponentResearch( _
  AsmDoc As ModelDoc2, ByRef Components As Dictionary, ByRef SearchFolders As Dictionary, _
  Exclude As Collection, AsmConf As String)
                      
  Dim Comp_ As Variant
  Dim Comp As Component2
  Dim Doc As ModelDoc2
  Dim Asm As AssemblyDoc
  Dim ComponentArray As Variant
  Dim Conf As String
  
  Set Asm = AsmDoc
  ComponentArray = Asm.GetComponents(True)
  If Not IsArrayEmpty(ComponentArray) Then 'бывают вспомогательные пустые сборки
    For Each Comp_ In ComponentArray
      Set Comp = Comp_
      If Comp.IsSuppressed Then  'погашен
        GoTo NextComp
      End If
      If Comp.ExcludeFromBOM Then  'исключен из спецификации
        GoTo NextComp
      End If
      If Comp.IsEnvelope Then  'конверт
        GoTo NextComp
      End If
      
      Set Doc = Comp.GetModelDoc2
      If Doc Is Nothing Then  'не найден
        GoTo NextComp
      End If
      Conf = Comp.ReferencedConfiguration
      
      If AddComponent(Doc, Conf, Components, SearchFolders, Exclude, AsmDoc.GetPathName, AsmConf) Then
        If Doc.GetType = swDocASSEMBLY Then
          Doc.ShowConfiguration2 Conf
          ComponentResearch Doc, Components, SearchFolders, Exclude, Conf
        End If
      End If
NextComp:
    Next
  End If
    
End Sub

Function AddComponent( _
  Doc As ModelDoc2, Conf As String, ByRef Components As Dictionary, _
  ByRef SearchFolders As Dictionary, Exclude As Collection, _
  AsmPath As String, AsmConf As String) As Boolean
                 
  Dim Key As String
  Dim Ci As ComponentInfo
  Dim ParentCi As ComponentInfo
  Dim ParentKey As String
  Dim FolderName As String
  Dim I As Variant
  
  AddComponent = False
  
  For Each I In Exclude
    If LCase(Doc.GetPathName) Like I Then
      Exit Function
    End If
  Next
  
  Key = CreateComponentKey(Doc.GetPathName, Conf)
  
  If Not Components.Exists(Key) Then
    Set Ci = CreateComponentInfo(Doc, Conf)
    
    If Doc.GetType = swDocASSEMBLY Then
      Ci.Parent = CreateBaseDesigname(Ci)
    Else
      ParentKey = CreateComponentKey(AsmPath, AsmConf)
      Set ParentCi = Components(ParentKey)
      Ci.Parent = CreateBaseDesigname(ParentCi)
    End If

    Components.Add Key, Ci
    FolderName = gFSO.GetParentFolderName(Doc.GetPathName)
    AddSearchFolder FolderName, SearchFolders, False
    AddComponent = True
  End If
   
End Function

Sub AddSearchFolder(FolderName As String, ByRef SearchFolders As Dictionary, IsRecursively As Boolean)

  Dim AFolder As Folder
  Dim Sf As SearchFolder
  
  On Error Resume Next
  Set AFolder = gFSO.GetFolder(FolderName)
  If Not AFolder Is Nothing Then
    Set Sf = New SearchFolder
    Set Sf.AFolder = AFolder
    Sf.IsRecursively = IsRecursively
    AddUniqueItemInDict LCase(FolderName), Sf, SearchFolders
  End If

End Sub

Sub AddUserSearchFolders(Include As Collection, ByRef SearchFolders As Dictionary)

  Dim I As Variant
  Dim FolderName As String
  
  For Each I In Include
    FolderName = I
    AddSearchFolder FolderName, SearchFolders, True
  Next

End Sub

Function CreateBaseDesigname(Ci As ComponentInfo) As String

  CreateBaseDesigname = Ci.BaseDesignation + " " + Ci.PropertyName
    
End Function

Function CreateComponentKey(DocPath As String, Conf As String)

  CreateComponentKey = gFSO.GetBaseName(DocPath) + "@" + Conf
    
End Function

Function CreatePattern(SoughtForExtensions As Dictionary) As RegExp

  Dim Regex As RegExp
  Dim I As Integer
  Dim Arr As Variant
  
  Arr = SoughtForExtensions.Keys
  For I = 0 To UBound(Arr)
    Arr(I) = RTrim(Arr(I))  'RegEscape?
  Next
  
  Set Regex = New RegExp
  Regex.Global = True
  Regex.IgnoreCase = True
  Regex.Pattern = ".*\.(" + Join(Arr, "|") + ")"
  Set CreatePattern = Regex
    
End Function

Function CreateComponentInfo(Doc As ModelDoc2, Conf As String) As ComponentInfo

  Dim Ci As ComponentInfo
  
  Set Ci = New ComponentInfo
  Set Ci.Doc = Doc
  Ci.Conf = Conf
  Ci.PropertyDesignation = GetProperty(gPropertyDesignation, Doc.Extension, Conf)
  Ci.PropertyName = GetProperty(gPropertyName, Doc.Extension, Conf)
  Ci.PropertyBlank = GetProperty(gPropertyBlank, Doc.Extension, Conf)
  Ci.BaseDesignation = GetBaseDesignation(Ci.PropertyDesignation)
  Set Ci.Drawings = New Dictionary
  
  Set CreateComponentInfo = Ci
    
End Function

Function GetProperty(Property As String, DocExt As ModelDocExtension, Conf As String) As String

  Dim ResultGetProp As swCustomInfoGetResult_e
  Dim RawProp As String
  Dim ResolvedValue As String
  Dim WasResolved As Boolean
  
  ResultGetProp = DocExt.CustomPropertyManager(Conf).Get5(Property, True, RawProp, ResolvedValue, WasResolved)
  If ResultGetProp = swCustomInfoGetResult_NotPresent Then
    DocExt.CustomPropertyManager("").Get5 Property, True, RawProp, ResolvedValue, WasResolved
  End If
  GetProperty = ResolvedValue
    
End Function

Sub CollectAllDrawings(Pattern As RegExp, SearchFolders As Dictionary, ByRef Drawings As Dictionary)

  Dim I As Variant
  
  For Each I In SearchFolders.Items
    CollectFolderDrawings I, Pattern, Drawings
  Next

End Sub

'Sf меняется внутри функции, на коллекцию это не должно влиять (уменьшение расхода памяти).
Sub CollectFolderDrawings(ByVal Sf As SearchFolder, Pattern As RegExp, ByRef Drawings As Dictionary)

  Dim I As Variant
  Dim F As File
  Dim SubFolder As Folder
  
  For Each I In Sf.AFolder.Files
    Set F = I
    If Pattern.Test(F.Path) Then
      Drawings.Add F.Path, ""
    End If
  Next
  
  If Sf.IsRecursively Then
    For Each I In Sf.AFolder.SubFolders
      Set Sf.AFolder = I
      CollectFolderDrawings Sf, Pattern, Drawings
    Next
  End If

End Sub

Function GetBaseDesignation(Designation As String) As String

  Dim LastFullstopPosition As Integer
  Dim FirstHyphenPosition As Integer
  
  GetBaseDesignation = Designation
  LastFullstopPosition = InStrRev(Designation, ".")
  If LastFullstopPosition > 0 Then
    FirstHyphenPosition = InStr(LastFullstopPosition, Designation, "-")
    If FirstHyphenPosition > 0 Then
      GetBaseDesignation = Left(Designation, FirstHyphenPosition - 1)
    End If
  End If
    
End Function

Function CheckMatchFileAndComponent(Fpath As String, ByRef Ci As ComponentInfo) As Boolean

  Const RegAnyPath = ".*\\"
  Const RegRev = " \((изм|rev)\.([0-9]{2})\)"
  Const RegCode = "(( *|\.)?(Р|СБ|РСБ|ВО|ТЧ|ГЧ|МЭ|МЧ|УЧ|ЭСБ|ПЭ|ПЗ|ТБ|РР|И|ТУ|ПМ|ВС|ВД|ВП|ВИ|ДП|ПТ|ЭП|ТП|ВДЭ|AD|ID))?"
  
  Dim Extension As String
  Dim RegExtension As String
  Dim RegDesignation As String
  Dim RegBaseDesignation As String
  Dim RegName As String
  Dim Revision As Integer
  Dim Priority As Integer
  
  CheckMatchFileAndComponent = False
  
  'Lower case fo extension is required!
  
  Extension = LCase(gFSO.GetExtensionName(Fpath))
  RegExtension = "\." + RegEscape(Extension)
  RegDesignation = RegEscape(Ci.PropertyDesignation) + RegCode
  RegBaseDesignation = RegEscape(Ci.BaseDesignation) + RegCode
  RegName = " +" + RegEscape(Ci.PropertyName) + " *"
  Priority = 10
  
  'Обозначение-01 Наименование
  ChangeRegex Priority, gRegex, RegAnyPath + RegDesignation + RegName + RegExtension
  If gRegex.Test(Fpath) Then
    CheckMatchFileAndComponent = AddMatchingDrawing(0, Extension, Fpath, Ci, Priority)
    Exit Function
  End If
  
  'Обозначение-01 Наименование (изм.##)
  ChangeRegex Priority, gRegex, RegAnyPath + RegDesignation + RegName + RegRev + RegExtension
  If gRegex.Test(Fpath) Then
    Revision = CInt(gRegex.Execute(Fpath)(0).SubMatches(4))
    'MsgBox revision & vbNewLine & fpath
    CheckMatchFileAndComponent = AddMatchingDrawing(Revision, Extension, Fpath, Ci, Priority)
    Exit Function
  End If
  
  'Обозначение-01
  ChangeRegex Priority, gRegex, RegAnyPath + RegDesignation + RegExtension
  If gRegex.Test(Fpath) Then
    CheckMatchFileAndComponent = AddMatchingDrawing(0, Extension, Fpath, Ci, Priority)
    Exit Function
  End If
  
  'Обозначение-01 (изм.##)
  ChangeRegex Priority, gRegex, RegAnyPath + RegDesignation + RegRev + RegExtension
  If gRegex.Test(Fpath) Then
    Revision = CInt(gRegex.Execute(Fpath)(0).SubMatches(4))
    'MsgBox revision & vbNewLine & fpath
    CheckMatchFileAndComponent = AddMatchingDrawing(Revision, Extension, Fpath, Ci, Priority)
    Exit Function
  End If
  
  'Для разверток не допускаются чертежи с базовым обозначением.
  If Extension <> "dwg" And Extension <> "dxf" Then
  
    'БазовоеОбозначение Наименование
    ChangeRegex Priority, gRegex, RegAnyPath + RegBaseDesignation + RegName + RegExtension
    If gRegex.Test(Fpath) Then
      CheckMatchFileAndComponent = AddMatchingDrawing(0, Extension, Fpath, Ci, Priority)
      Exit Function
    End If
    
    'БазовоеОбозначение Наименование (изм.##)
    ChangeRegex Priority, gRegex, RegAnyPath + RegBaseDesignation + RegName + RegRev + RegExtension
    If gRegex.Test(Fpath) Then
      Revision = CInt(gRegex.Execute(Fpath)(0).SubMatches(4))
      'MsgBox revision & vbNewLine & fpath
      CheckMatchFileAndComponent = AddMatchingDrawing(Revision, Extension, Fpath, Ci, Priority)
      Exit Function
    End If
    
    'БазовоеОбозначение
    ChangeRegex Priority, gRegex, RegAnyPath + RegBaseDesignation + RegExtension
    If gRegex.Test(Fpath) Then
      CheckMatchFileAndComponent = AddMatchingDrawing(0, Extension, Fpath, Ci, Priority)
      Exit Function
    End If
    
    'БазовоеОбозначение (изм.##)
    ChangeRegex Priority, gRegex, RegAnyPath + RegBaseDesignation + RegRev + RegExtension
    If gRegex.Test(Fpath) Then
      Revision = CInt(gRegex.Execute(Fpath)(0).SubMatches(4))
      'MsgBox revision & vbNewLine & fpath
      CheckMatchFileAndComponent = AddMatchingDrawing(Revision, Extension, Fpath, Ci, Priority)
      Exit Function
    End If

  End If
   
End Function

Sub ChangeRegex(ByRef Priority As Integer, ByRef Regex As RegExp, Pattern As String)

  Priority = Priority - 1
  Regex.Pattern = Pattern
    
End Sub

Function AddMatchingDrawing( _
  Revision As Integer, Extension As String, Fpath As String, _
  ByRef Ci As ComponentInfo, Priority As Integer) As Boolean
                      
  Dim Dr As DrawingRecord
  
  AddMatchingDrawing = False
  
  If Not Ci.Drawings.Exists(Extension) Then
    Set Dr = New DrawingRecord
    Dr.Path = Fpath
    Dr.Rev = Revision
    Dr.Priority = Priority
    Ci.Drawings.Add Extension, Dr
    AddMatchingDrawing = True
  Else
    Set Dr = Ci.Drawings(Extension)
    If (Revision > Dr.Rev) Or (Revision = Dr.Rev And Priority > Dr.Priority) Then
      Dr.Path = Fpath
      Dr.Rev = Revision
      Dr.Priority = Priority
      AddMatchingDrawing = True
    End If
  End If
   
End Function

Function SplitLine(Text As String) As Collection

  Dim I As Variant
  Dim Line As String
  Dim Col As Collection
  
  Set Col = New Collection
  For Each I In Split(Text, vbNewLine)
    Line = Trim(I)
    If Line <> "" Then
      Col.Add LCase(Line)
    End If
  Next
  Set SplitLine = Col
    
End Function

'See: https://docs.microsoft.com/en-us/dotnet/api/System.Text.RegularExpressions.Regex.Escape
Function RegEscape(ByVal Line As String) As String

  Line = Replace(Line, "\", "\\")  'MUST be first!
  Line = Replace(Line, ".", "\.")
  Line = Replace(Line, "[", "\[")
  'line = Replace(line, "]", "\]")
  Line = Replace(Line, "|", "\|")
  Line = Replace(Line, "^", "\^")
  Line = Replace(Line, "$", "\$")
  Line = Replace(Line, "?", "\?")
  Line = Replace(Line, "+", "\+")
  Line = Replace(Line, "*", "\*")
  Line = Replace(Line, "{", "\{")
  'line = Replace(line, "}", "\}")
  Line = Replace(Line, "(", "\(")
  Line = Replace(Line, ")", "\)")
  Line = Replace(Line, "#", "\#")
  'and white space??
  RegEscape = Line
    
End Function

Function ExitApp() 'hide

  Unload MainForm
  End
    
End Function

