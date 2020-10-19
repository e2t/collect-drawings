Attribute VB_Name = "Tools"
Option Explicit

Const gPropertyDesignation = "Обозначение"
Const gPropertyName = "Наименование"
Const gPropertyBlank = "Заготовка"

Sub MatchFiles(ByRef components As Dictionary, Drawings As Collection)
    Dim ci_ As Variant
    Dim ci As ComponentInfo
    Dim drw_ As Variant
    Dim drw As String
    
    For Each ci_ In components.Items
        Set ci = ci_
        For Each drw_ In Drawings
            drw = drw_
            CheckMatchFileAndComponent drw, ci
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

Sub AssignItem(ci As ComponentInfo, VeritacalLineSeparatedExtensions As String, _
               ByRef copied As Dictionary, ByRef NotFound As Dictionary)
   Dim rec As DrawingRecord
   Dim Extension As Variant
   Dim Found As Boolean
   
   Found = False
   For Each Extension In Split(VeritacalLineSeparatedExtensions, "|")
      If ci.Drawings.Exists(Extension) Then
         Set rec = ci.Drawings(Extension)
         AddUniqueItemInDict rec.path, ci.Parent, copied
         Found = True
      End If
   Next
   If Not Found Then
      AddUniqueItemInDict ci, 0, NotFound(VeritacalLineSeparatedExtensions)
   End If
End Sub

Sub CopyFiles(copied As Dictionary, target As String)
    Dim f_ As Variant
    Dim f As String
    Dim newname As String
    Dim FolderName As String
    
    If Not gFSO.FolderExists(target) Then
        gFSO.CreateFolder target
    End If
    
    For Each f_ In copied.Keys
        f = f_
        FolderName = target + "\" + copied(f)
        If Not gFSO.FolderExists(FolderName) Then
            gFSO.CreateFolder FolderName
        End If
        
        newname = FolderName + "\" + gFSO.GetFileName(f)
        
        If LCase(f) <> LCase(newname) Then
            gFSO.CopyFile f, newname
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
   
   text = "Найдено чертежей: " + str(foundCount)
   
   For Each Extension In NotFound
      Set NotFoundSomeFormat = NotFound(Extension)
      If NotFoundSomeFormat.count > 0 Then
         ReDim notfoundArray(NotFoundSomeFormat.count - 1)
         text = text + vbNewLine + vbNewLine + "Не найдены чертежи " + UCase(Extension) + " для:"
         I = -1
         For Each ci_ In NotFoundSomeFormat.Keys
            Set ci = ci_
            I = I + 1
            notfoundArray(I) = Trim(ci.PropertyDesignation + " " + ci.PropertyName) + _
                               " [файл: " + gFSO.GetFileName(ci.Doc.GetPathName) + " @ " + ci.Conf + _
                               " ]"
         Next
         QuickSort notfoundArray, 0, I
         LogNotFoundArray notfoundArray
         text = text + vbNewLine + Join(notfoundArray, vbNewLine)
      End If
   Next
   CreateOutput = text
End Function

Sub LogNotFoundArray(notfoundArray() As String)
   Dim FileStream As TextStream
   Dim I As Variant
   
   Set FileStream = gFSO.OpenTextFile(GetLogFileName, ForWriting, True)
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

Sub ComponentResearch(asmDoc As ModelDoc2, ByRef components As Dictionary, ByRef searchFolders As Dictionary, _
                      exclude As Collection, asmConf As String)
    Dim comp_ As Variant
    Dim comp As Component2
    Dim Doc As ModelDoc2
    Dim asm As AssemblyDoc
    
    Set asm = asmDoc
    For Each comp_ In asm.GetComponents(True)
        Set comp = comp_
        If comp.IsSuppressed Then  'погашен
            GoTo NextFor
        End If
        Set Doc = comp.GetModelDoc2
        If Doc Is Nothing Then  'не найден
            GoTo NextFor
        End If
        AddComponent comp.GetModelDoc2, comp.ReferencedConfiguration, components, searchFolders, _
                     exclude, asmDoc.GetPathName, asmConf
        If Doc.GetType = swDocASSEMBLY Then
            ComponentResearch Doc, components, searchFolders, exclude, comp.ReferencedConfiguration
        End If
NextFor:
    Next
End Sub

Sub AddComponent(Doc As ModelDoc2, Conf As String, ByRef components As Dictionary, _
                 ByRef searchFolders As Dictionary, exclude As Collection, _
                 asmPath As String, asmConf As String)
   Dim key As String
   Dim excludePattern As Variant
   Dim ci As ComponentInfo
   Dim parentCi As ComponentInfo
   Dim parentKey As String
   Dim FolderName As String
   
   For Each excludePattern In exclude
      If LCase(Doc.GetPathName) Like excludePattern Then
         Exit Sub
      End If
   Next
   key = CreateComponentKey(Doc.GetPathName, Conf)
   
   If Not components.Exists(key) Then
      Set ci = CreateComponentInfo(Doc, Conf)
      
      If Doc.GetType = swDocASSEMBLY Then
         ci.Parent = CreateBaseDesigname(ci)
      Else
         parentKey = CreateComponentKey(asmPath, asmConf)
         Set parentCi = components(parentKey)
         ci.Parent = CreateBaseDesigname(parentCi)
      End If

      components.Add key, ci
      FolderName = gFSO.GetParentFolderName(Doc.GetPathName)
      AddUniqueItemInDict LCase(FolderName), gFSO.GetFolder(FolderName), searchFolders
   End If
End Sub

Function CreateBaseDesigname(ci As ComponentInfo) As String
    CreateBaseDesigname = ci.BaseDesignation + " " + ci.PropertyName
End Function

Function CreateComponentKey(docPath As String, Conf As String)
    CreateComponentKey = gFSO.GetBaseName(docPath) + "@" + Conf
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

Function CreateComponentInfo(Doc As ModelDoc2, Conf As String) As ComponentInfo
    Dim ci As ComponentInfo
    
    Set ci = New ComponentInfo
    Set ci.Doc = Doc
    ci.Conf = Conf
    ci.PropertyDesignation = GetProperty(gPropertyDesignation, Doc.Extension, Conf)
    ci.PropertyName = GetProperty(gPropertyName, Doc.Extension, Conf)
    ci.PropertyBlank = GetProperty(gPropertyBlank, Doc.Extension, Conf)
    ci.BaseDesignation = GetBaseDesignation(ci.PropertyDesignation)
    Set ci.Drawings = New Dictionary
    
    Set CreateComponentInfo = ci
End Function

Function GetProperty(property As String, docext As ModelDocExtension, Conf As String) As String
    Dim resultGetProp As swCustomInfoGetResult_e
    Dim rawProp As String, resolvedValue As String
    Dim wasResolved As Boolean
    
    resultGetProp = docext.CustomPropertyManager(Conf).Get5(property, True, rawProp, resolvedValue, wasResolved)
    If resultGetProp = swCustomInfoGetResult_NotPresent Then
        docext.CustomPropertyManager("").Get5 property, True, rawProp, resolvedValue, wasResolved
    End If
    GetProperty = resolvedValue
End Function

Sub CollectAllDrawings(pattern As RegExp, searchFolders As Dictionary, ByRef Drawings As Collection)
    Dim aFolder_ As Variant
    Dim aFolder As Folder
    Dim f_ As Variant
    Dim f As File
    
    For Each aFolder_ In searchFolders.Items
        Set aFolder = aFolder_
        For Each f_ In aFolder.Files
            Set f = f_
            If pattern.Test(f.path) Then
                Drawings.Add f.path
            End If
        Next
    Next
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

Sub CheckMatchFileAndComponent(fpath As String, ByRef ci As ComponentInfo)
   Const regAnyPath = ".*\\"
   Const regRev = " \(изм\.([0-9]{2})\)"
   Const regCode = "( *(Р|СБ|РСБ|ВО|ТЧ|ГЧ|МЭ|МЧ|УЧ|ЭСБ|ПЭ|ПЗ|ТБ|РР|И|ТУ|ПМ|ВС|ВД|ВП|ВИ|ДП|ПТ|ЭП|ТП|ВДЭ|AD|ID))?"
   
   Dim Regex As RegExp
   Dim Extension As String
   Dim regExtension As String
   Dim regDesignation As String
   Dim regBaseDesignation As String
   Dim regName As String
   Dim revision As Integer
   Dim priority As Integer
   
   Set Regex = New RegExp
   Regex.Global = True
   Regex.IgnoreCase = True
   
   'Lower case fo extension is required!
   
   Extension = LCase(gFSO.GetExtensionName(fpath))
   regExtension = "\." + RegEscape(Extension)
   regDesignation = RegEscape(ci.PropertyDesignation) + regCode
   regBaseDesignation = RegEscape(ci.BaseDesignation) + regCode
   regName = " +" + RegEscape(ci.PropertyName) + " *"
   priority = 10
   
   'Обозначение-01 Наименование
   ChangeRegex priority, Regex, regAnyPath + regDesignation + regName + regExtension
   If Regex.Test(fpath) Then
      AddMatchingDrawing 0, Extension, fpath, ci, priority
      Exit Sub
   End If
   
   'Обозначение-01 Наименование (изм.##)
   ChangeRegex priority, Regex, regAnyPath + regDesignation + regName + regRev + regExtension
   If Regex.Test(fpath) Then
      revision = CInt(Regex.Execute(fpath)(0).SubMatches(2))
      AddMatchingDrawing revision, Extension, fpath, ci, priority
      Exit Sub
   End If
   
   'Обозначение-01
   ChangeRegex priority, Regex, regAnyPath + regDesignation + regExtension
   If Regex.Test(fpath) Then
      AddMatchingDrawing 0, Extension, fpath, ci, priority
      Exit Sub
   End If
   
   'Обозначение-01 (изм.##)
   ChangeRegex priority, Regex, regAnyPath + regDesignation + regRev + regExtension
   If Regex.Test(fpath) Then
      revision = CInt(Regex.Execute(fpath)(0).SubMatches(2))
      AddMatchingDrawing revision, Extension, fpath, ci, priority
      Exit Sub
   End If

   'Для разверток не допускаются чертежи с базовым обозначением.
   If Extension <> "dwg" And Extension <> "dxf" Then
    
      'БазовоеОбозначение Наименование
      ChangeRegex priority, Regex, regAnyPath + regBaseDesignation + regName + regExtension
      If Regex.Test(fpath) Then
         AddMatchingDrawing 0, Extension, fpath, ci, priority
         Exit Sub
      End If
      
      'БазовоеОбозначение Наименование (изм.##)
      ChangeRegex priority, Regex, regAnyPath + regBaseDesignation + regName + regRev + regExtension
      If Regex.Test(fpath) Then
         revision = CInt(Regex.Execute(fpath)(0).SubMatches(2))
         AddMatchingDrawing revision, Extension, fpath, ci, priority
         Exit Sub
      End If
      
      'БазовоеОбозначение
      ChangeRegex priority, Regex, regAnyPath + regBaseDesignation + regExtension
      If Regex.Test(fpath) Then
         AddMatchingDrawing 0, Extension, fpath, ci, priority
         Exit Sub
      End If
      
      'БазовоеОбозначение (изм.##)
      ChangeRegex priority, Regex, regAnyPath + regBaseDesignation + regRev + regExtension
      If Regex.Test(fpath) Then
         revision = CInt(Regex.Execute(fpath)(0).SubMatches(2))
         AddMatchingDrawing revision, Extension, fpath, ci, priority
         Exit Sub
      End If
      
   End If
End Sub

Sub ChangeRegex(ByRef priority As Integer, ByRef Regex As RegExp, pattern As String)
    priority = priority - 1
    Regex.pattern = pattern
End Sub

Sub AddMatchingDrawing(revision As Integer, Extension As String, _
                       fpath As String, ByRef ci As ComponentInfo, priority As Integer)
    Dim dr As DrawingRecord
    
    If Not ci.Drawings.Exists(Extension) Then
        Set dr = New DrawingRecord
        dr.path = fpath
        dr.rev = revision
        dr.priority = priority
        ci.Drawings.Add Extension, dr
    Else
        Set dr = ci.Drawings(Extension)
        If (revision > dr.rev) Or (revision = dr.rev And priority > dr.priority) Then
            dr.path = fpath
            dr.rev = revision
            dr.priority = priority
        End If
    End If
End Sub

Function CreateExclude(text As String) As Collection
    Dim line_ As Variant
    Dim line As String
    Dim col As Collection
    
    Set col = New Collection
    For Each line_ In Split(text, vbNewLine)
        line = Trim(line_)
        If line <> "" Then
            col.Add LCase(line)
        End If
    Next
    Set CreateExclude = col
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

