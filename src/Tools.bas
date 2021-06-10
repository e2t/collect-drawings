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

Sub MatchFiles(ByRef components As Dictionary, ByRef Drawings As Dictionary)

   Dim ci_ As Variant
   Dim ci As ComponentInfo
   Dim drw_ As Variant
   Dim drw As String
   Dim key As String
   
   For Each ci_ In components.Items
      Set ci = ci_
      key = ci.Doc.GetPathName
      
      For Each drw_ In Drawings.Keys
         drw = drw_
         Select Case Drawings(drw)
            Case ""
               If CheckMatchFileAndComponent(drw, ci) Then
                  Drawings(drw) = key
               End If
            Case key
               CheckMatchFileAndComponent drw, ci
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

Sub UniqueCopiedFiles(components As Dictionary, ByRef copied As Dictionary, _
                      ByRef NotFound As Dictionary, SoughtForExtensions As Dictionary)
                      
   Dim ci_ As Variant
   Dim ci As ComponentInfo
   Dim IsSheetDetail As RegExp
   
   Set IsSheetDetail = New RegExp
   IsSheetDetail.Global = True
   IsSheetDetail.IgnoreCase = True
   IsSheetDetail.pattern = ".*(лист|sheet).*"
   
   'Lower case fo extension is required!

   NotFound.Add "pdf", New Dictionary
   NotFound.Add "dwg|dxf", New Dictionary
   NotFound.Add "xls", New Dictionary

   For Each ci_ In components.Items
      Set ci = ci_
      If SoughtForExtensions.Exists("pdf") Then
         AssignItem ci, "pdf", copied, NotFound 'PDF - Для всех.
      End If
      If SoughtForExtensions.Exists("dwg") Or SoughtForExtensions.Exists("dxf") Then
         If IsSheetDetail.Test(ci.PropertyBlank) Then 'DWG/DXF - Для листовых деталей.
            AssignItem ci, "dwg|dxf", copied, NotFound
         End If
      End If
      If SoughtForExtensions.Exists("xls") Then
         If ci.Doc.GetType = swDocASSEMBLY Then 'XLS - Для сборок.
            AssignItem ci, "xls", copied, NotFound
         End If
      End If
   Next
   
End Sub

Sub AssignItem(ci As ComponentInfo, VerticalLineSeparatedExtensions As String, _
               ByRef copied As Dictionary, ByRef NotFound As Dictionary)
               
   Dim rec As DrawingRecord
   Dim Extension As Variant
   Dim Found As Boolean
   
   Found = False
   For Each Extension In Split(VerticalLineSeparatedExtensions, "|")
      If ci.Drawings.Exists(Extension) Then
         Set rec = ci.Drawings(Extension)
         AddUniqueItemInDict rec.path, ci.Parent, copied
         Found = True
      End If
   Next
   If Not Found Then
      AddUniqueItemInDict ci, 0, NotFound(VerticalLineSeparatedExtensions)
   End If
   
End Sub

Sub CopyFiles(copied As Dictionary, target As String)

   Dim f_ As Variant
   Dim F As String
   Dim newname As String
   Dim FolderName As String
   
   If Not gFSO.FolderExists(target) Then
      gFSO.CreateFolder target
   End If
   
   For Each f_ In copied.Keys
      F = f_
      FolderName = target + "\" + Trim(copied(F))
      If Not gFSO.FolderExists(FolderName) Then
         gFSO.CreateFolder FolderName
      End If
      
      newname = FolderName + "\" + gFSO.GetFileName(F)
      
      If LCase(F) <> LCase(newname) Then
         gFSO.CopyFile F, newname
      End If
   Next
    
End Sub

Function CreateOutput(foundCount As Integer, NotFound As Dictionary) As String

   Dim ci_ As Variant
   Dim ci As ComponentInfo
   Dim text As String
   Dim notfoundArray() As String
   Dim I As Long
   Dim Extension As Variant
   Dim NotFoundSomeFormat As Dictionary
   Dim FirstTime As Boolean
   
   text = "Найдено чертежей: " + Str(foundCount)
   FirstTime = True
   
   For Each Extension In NotFound
      Set NotFoundSomeFormat = NotFound(Extension)
      If NotFoundSomeFormat.Count > 0 Then
         ReDim notfoundArray(NotFoundSomeFormat.Count - 1)
         text = text + vbNewLine + vbNewLine + "Не найдены чертежи " + UCase(Extension) + " для:"
         I = -1
         For Each ci_ In NotFoundSomeFormat.Keys
            Set ci = ci_
            I = I + 1
            notfoundArray(I) = Trim(ci.PropertyDesignation + " " + ci.PropertyName) + _
                               " [файл: " + gFSO.GetFileName(ci.Doc.GetPathName) + " @ " + ci.conf + _
                               " ]"
         Next
         QuickSort notfoundArray, 0, I
         LogNotFoundArray notfoundArray, FirstTime
         text = text + vbNewLine + Join(notfoundArray, vbNewLine)
      End If
   Next
   CreateOutput = text
   
End Function

Sub LogNotFoundArray(notfoundArray() As String, ByRef FirstTime As Boolean)

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
   For Each I In notfoundArray
      FileStream.WriteLine I
   Next
   FileStream.Close
   
End Sub

'Функция-сокращение для любых словарей.
'True - если был создан новый объект, False - если объект существовал.
Function AddUniqueItemInDict(key As Variant, item As Variant, ByRef dict As Dictionary) As Boolean

   AddUniqueItemInDict = Not dict.Exists(key)
   If AddUniqueItemInDict Then
      dict.Add key, item
   End If
    
End Function

Function IsArrayEmpty(ByRef anArray As Variant) As Boolean

   Dim I As Integer
 
   On Error GoTo ArrayIsEmpty
   IsArrayEmpty = LBound(anArray) > UBound(anArray)
   Exit Function
ArrayIsEmpty:
   IsArrayEmpty = True

End Function

Sub ComponentResearch(asmDoc As ModelDoc2, ByRef components As Dictionary, ByRef searchFolders As Dictionary, _
                      exclude As Collection, asmConf As String)
                      
   Dim comp_ As Variant
   Dim comp As Component2
   Dim Doc As ModelDoc2
   Dim asm As AssemblyDoc
   Dim componentArray As Variant
   Dim conf As String
   
   Set asm = asmDoc
   componentArray = asm.GetComponents(True)
   If Not IsArrayEmpty(componentArray) Then 'бывают вспомогательные пустые сборки
      For Each comp_ In componentArray
         Set comp = comp_
         If comp.IsSuppressed Then  'погашен
            GoTo NextComp
         End If
         Set Doc = comp.GetModelDoc2
         If Doc Is Nothing Then  'не найден
            GoTo NextComp
         End If
         conf = comp.ReferencedConfiguration
         
         If AddComponent(Doc, conf, components, searchFolders, exclude, asmDoc.GetPathName, asmConf) Then
            If Doc.GetType = swDocASSEMBLY Then
               Doc.ShowConfiguration2 conf
               ComponentResearch Doc, components, searchFolders, exclude, conf
            End If
         End If
NextComp:
      Next
   End If
    
End Sub

Function AddComponent(Doc As ModelDoc2, conf As String, ByRef components As Dictionary, _
                      ByRef searchFolders As Dictionary, exclude As Collection, _
                      asmPath As String, asmConf As String) As Boolean
                 
   Dim key As String
   Dim ci As ComponentInfo
   Dim parentCi As ComponentInfo
   Dim parentKey As String
   Dim FolderName As String
   Dim I As Variant
   
   AddComponent = False
   
   For Each I In exclude
      If LCase(Doc.GetPathName) Like I Then
         Exit Function
      End If
   Next
   
   key = CreateComponentKey(Doc.GetPathName, conf)
   
   If Not components.Exists(key) Then
      Set ci = CreateComponentInfo(Doc, conf)
      
      If Doc.GetType = swDocASSEMBLY Then
         ci.Parent = CreateBaseDesigname(ci)
      Else
         parentKey = CreateComponentKey(asmPath, asmConf)
         Set parentCi = components(parentKey)
         ci.Parent = CreateBaseDesigname(parentCi)
      End If

      components.Add key, ci
      FolderName = gFSO.GetParentFolderName(Doc.GetPathName)
      AddSearchFolder FolderName, searchFolders, False
      AddComponent = True
   End If
   
End Function

Sub AddSearchFolder(FolderName As String, ByRef searchFolders As Dictionary, IsRecursively As Boolean)

  Dim AFolder As Folder
  Dim Sf As SearchFolder
  
  On Error Resume Next
  Set AFolder = gFSO.GetFolder(FolderName)
  If Not AFolder Is Nothing Then
    Set Sf = New SearchFolder
    Set Sf.AFolder = AFolder
    Sf.IsRecursively = IsRecursively
    AddUniqueItemInDict LCase(FolderName), Sf, searchFolders
  End If

End Sub

Sub AddUserSearchFolders(include As Collection, ByRef searchFolders As Dictionary)

  Dim I As Variant
  Dim FolderName As String
  
  For Each I In include
    FolderName = I
    AddSearchFolder FolderName, searchFolders, True
  Next

End Sub

Function CreateBaseDesigname(ci As ComponentInfo) As String
    CreateBaseDesigname = ci.BaseDesignation + " " + ci.PropertyName
End Function

Function CreateComponentKey(docPath As String, conf As String)
    CreateComponentKey = gFSO.GetBaseName(docPath) + "@" + conf
End Function

Function CreatePattern(SoughtForExtensions As Dictionary) As RegExp
    Dim Regex As RegExp
    Dim I As Integer
    Dim arr As Variant
    
    arr = SoughtForExtensions.Keys
    For I = 0 To UBound(arr)
        arr(I) = RTrim(arr(I))  'RegEscape?
    Next
    
    Set Regex = New RegExp
    Regex.Global = True
    Regex.IgnoreCase = True
    Regex.pattern = ".*\.(" + Join(arr, "|") + ")"
    Set CreatePattern = Regex
End Function

Function CreateComponentInfo(Doc As ModelDoc2, conf As String) As ComponentInfo
    Dim ci As ComponentInfo
    
    Set ci = New ComponentInfo
    Set ci.Doc = Doc
    ci.conf = conf
    ci.PropertyDesignation = GetProperty(gPropertyDesignation, Doc.Extension, conf)
    ci.PropertyName = GetProperty(gPropertyName, Doc.Extension, conf)
    ci.PropertyBlank = GetProperty(gPropertyBlank, Doc.Extension, conf)
    ci.BaseDesignation = GetBaseDesignation(ci.PropertyDesignation)
    Set ci.Drawings = New Dictionary
    
    Set CreateComponentInfo = ci
End Function

Function GetProperty(property As String, docext As ModelDocExtension, conf As String) As String
    Dim resultGetProp As swCustomInfoGetResult_e
    Dim rawProp As String, resolvedValue As String
    Dim wasResolved As Boolean
    
    resultGetProp = docext.CustomPropertyManager(conf).Get5(property, True, rawProp, resolvedValue, wasResolved)
    If resultGetProp = swCustomInfoGetResult_NotPresent Then
        docext.CustomPropertyManager("").Get5 property, True, rawProp, resolvedValue, wasResolved
    End If
    GetProperty = resolvedValue
End Function

Sub CollectAllDrawings(pattern As RegExp, searchFolders As Dictionary, ByRef Drawings As Dictionary)

  Dim I As Variant
  
  For Each I In searchFolders.Items
    CollectFolderDrawings I, pattern, Drawings
  Next

End Sub

'Sf меняется внутри функции, на коллекцию это не должно влиять (уменьшение расхода памяти).
Sub CollectFolderDrawings(ByVal Sf As SearchFolder, pattern As RegExp, ByRef Drawings As Dictionary)

  Dim I As Variant
  Dim F As File
  Dim SubFolder As Folder
  
  For Each I In Sf.AFolder.Files
    Set F = I
    If pattern.Test(F.path) Then
      Drawings.Add F.path, ""
    End If
  Next
  
  If Sf.IsRecursively Then
    For Each I In Sf.AFolder.SubFolders
      Set Sf.AFolder = I
      CollectFolderDrawings Sf, pattern, Drawings
    Next
  End If

End Sub

Function GetBaseDesignation(designation As String) As String
    Dim lastFullstopPosition As Integer
    Dim firstHyphenPosition As Integer
    
    GetBaseDesignation = designation
    lastFullstopPosition = InStrRev(designation, ".")
    If lastFullstopPosition > 0 Then
        firstHyphenPosition = InStr(lastFullstopPosition, designation, "-")
        If firstHyphenPosition > 0 Then
            GetBaseDesignation = Left(designation, firstHyphenPosition - 1)
        End If
    End If
End Function

Function CheckMatchFileAndComponent(fpath As String, ByRef ci As ComponentInfo) As Boolean
   Const regAnyPath = ".*\\"
   Const regRev = " \((изм|rev)\.([0-9]{2})\)"
   Const regCode = "(( *|\.)?(Р|СБ|РСБ|ВО|ТЧ|ГЧ|МЭ|МЧ|УЧ|ЭСБ|ПЭ|ПЗ|ТБ|РР|И|ТУ|ПМ|ВС|ВД|ВП|ВИ|ДП|ПТ|ЭП|ТП|ВДЭ|AD|ID))?"
   
   Dim Extension As String
   Dim regExtension As String
   Dim regDesignation As String
   Dim regBaseDesignation As String
   Dim regName As String
   Dim revision As Integer
   Dim priority As Integer
   
   CheckMatchFileAndComponent = False
   
   'Lower case fo extension is required!
   
   Extension = LCase(gFSO.GetExtensionName(fpath))
   regExtension = "\." + RegEscape(Extension)
   regDesignation = RegEscape(ci.PropertyDesignation) + regCode
   regBaseDesignation = RegEscape(ci.BaseDesignation) + regCode
   regName = " +" + RegEscape(ci.PropertyName) + " *"
   priority = 10
   
   'Обозначение-01 Наименование
   ChangeRegex priority, gRegex, regAnyPath + regDesignation + regName + regExtension
   If gRegex.Test(fpath) Then
      CheckMatchFileAndComponent = AddMatchingDrawing(0, Extension, fpath, ci, priority)
      Exit Function
   End If
   
   'Обозначение-01 Наименование (изм.##)
   ChangeRegex priority, gRegex, regAnyPath + regDesignation + regName + regRev + regExtension
   If gRegex.Test(fpath) Then
      revision = CInt(gRegex.Execute(fpath)(0).SubMatches(4))
      'MsgBox revision & vbNewLine & fpath
      CheckMatchFileAndComponent = AddMatchingDrawing(revision, Extension, fpath, ci, priority)
      Exit Function
   End If
   
   'Обозначение-01
   ChangeRegex priority, gRegex, regAnyPath + regDesignation + regExtension
   If gRegex.Test(fpath) Then
      CheckMatchFileAndComponent = AddMatchingDrawing(0, Extension, fpath, ci, priority)
      Exit Function
   End If
   
   'Обозначение-01 (изм.##)
   ChangeRegex priority, gRegex, regAnyPath + regDesignation + regRev + regExtension
   If gRegex.Test(fpath) Then
      revision = CInt(gRegex.Execute(fpath)(0).SubMatches(4))
      'MsgBox revision & vbNewLine & fpath
      CheckMatchFileAndComponent = AddMatchingDrawing(revision, Extension, fpath, ci, priority)
      Exit Function
   End If
   
   'Для разверток не допускаются чертежи с базовым обозначением.
   If Extension <> "dwg" And Extension <> "dxf" Then
   
      'БазовоеОбозначение Наименование
      ChangeRegex priority, gRegex, regAnyPath + regBaseDesignation + regName + regExtension
      If gRegex.Test(fpath) Then
         CheckMatchFileAndComponent = AddMatchingDrawing(0, Extension, fpath, ci, priority)
         Exit Function
      End If
      
      'БазовоеОбозначение Наименование (изм.##)
      ChangeRegex priority, gRegex, regAnyPath + regBaseDesignation + regName + regRev + regExtension
      If gRegex.Test(fpath) Then
         revision = CInt(gRegex.Execute(fpath)(0).SubMatches(4))
         'MsgBox revision & vbNewLine & fpath
         CheckMatchFileAndComponent = AddMatchingDrawing(revision, Extension, fpath, ci, priority)
         Exit Function
      End If
      
      'БазовоеОбозначение
      ChangeRegex priority, gRegex, regAnyPath + regBaseDesignation + regExtension
      If gRegex.Test(fpath) Then
         CheckMatchFileAndComponent = AddMatchingDrawing(0, Extension, fpath, ci, priority)
         Exit Function
      End If
      
      'БазовоеОбозначение (изм.##)
      ChangeRegex priority, gRegex, regAnyPath + regBaseDesignation + regRev + regExtension
      If gRegex.Test(fpath) Then
         revision = CInt(gRegex.Execute(fpath)(0).SubMatches(4))
         'MsgBox revision & vbNewLine & fpath
         CheckMatchFileAndComponent = AddMatchingDrawing(revision, Extension, fpath, ci, priority)
         Exit Function
      End If

   End If
   
End Function

Sub ChangeRegex(ByRef priority As Integer, ByRef Regex As RegExp, pattern As String)
    priority = priority - 1
    Regex.pattern = pattern
End Sub

Function AddMatchingDrawing( _
   revision As Integer, Extension As String, fpath As String, _
   ByRef ci As ComponentInfo, priority As Integer) As Boolean
                       
   Dim dr As DrawingRecord
   
   AddMatchingDrawing = False
   
   If Not ci.Drawings.Exists(Extension) Then
      Set dr = New DrawingRecord
      dr.path = fpath
      dr.rev = revision
      dr.priority = priority
      ci.Drawings.Add Extension, dr
      AddMatchingDrawing = True
   Else
      Set dr = ci.Drawings(Extension)
      If (revision > dr.rev) Or (revision = dr.rev And priority > dr.priority) Then
         dr.path = fpath
         dr.rev = revision
         dr.priority = priority
         AddMatchingDrawing = True
      End If
   End If
   
End Function

Function SplitLine(text As String) As Collection
    Dim I As Variant
    Dim line As String
    Dim col As Collection
    
    Set col = New Collection
    For Each I In Split(text, vbNewLine)
        line = Trim(I)
        If line <> "" Then
            col.Add LCase(line)
        End If
    Next
    Set SplitLine = col
End Function

'See: https://docs.microsoft.com/en-us/dotnet/api/System.Text.RegularExpressions.Regex.Escape
Function RegEscape(ByVal line As String) As String
    line = Replace(line, "\", "\\")  'MUST be first!
    line = Replace(line, ".", "\.")
    line = Replace(line, "[", "\[")
    'line = Replace(line, "]", "\]")
    line = Replace(line, "|", "\|")
    line = Replace(line, "^", "\^")
    line = Replace(line, "$", "\$")
    line = Replace(line, "?", "\?")
    line = Replace(line, "+", "\+")
    line = Replace(line, "*", "\*")
    line = Replace(line, "{", "\{")
    'line = Replace(line, "}", "\}")
    line = Replace(line, "(", "\(")
    line = Replace(line, ")", "\)")
    line = Replace(line, "#", "\#")
    'and white space??
    RegEscape = line
End Function

Function ExitApp() 'masked
    Unload MainForm
    End
End Function

