'********************************************************************************
'                           Factorio Mods Locale Tool
'                   Coded by Mr.Jos   Email: sd7056333@163.com
'********************************************************************************

'---------------------------- The MIT License (MIT) -----------------------------

'Copyright (c) 2016 Mr.Jos

'Permission is hereby granted, free of charge, to any person obtaining a copy of 
'this software and associated documentation files (the "Software"), to deal in 
'the Software without restriction, including without limitation the rights to 
'use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies 
'of the Software, and to permit persons to whom the Software is furnished to do 
'so, subject to the following conditions:

'The above copyright notice and this permission notice shall be included in all 
'copies or substantial portions of the Software.

'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR 
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS 
'FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR 
'COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER 
'IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN 
'CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

'--------------------------------------------------------------------------------

Option Explicit
CONST SCRIPT_VERSION = 283  'Update date: 2016.07.22

'----------------------------------- Options ------------------------------------

'WARNING: THE FOLLOWING OPTIONS ARE SET ONLY FOR LIBRARY AUTHORS!

CONST NAME_LIBRARY = "[FMLT] Library for zh-CN"
'    This is the name of locale library in the same directory as this script, 
'which is either a zip file or a folder with valid library-info.json & 
'script-echo.json inside.

CONST TEXT_PRIORITY = False
'    This setting is used to toggle the priority of the text source to translate 
'each file in mods for some particular cases. The following uses zh-CN as an 
'example of the locale language.
'    If True, the priority is [zh-CN in mod]>>[library]>>[en in mod]. In this 
'case, only the untranslated items in mods will be rewrited and translated, which 
'is usually applied to auto-collect the translated texts in mods.
'    If False (default), the priority is [library]>>[zh-CN in mod]>>[en in mod]. 

CONST UPDATE_LIBRARY = False
'    This setting is used to toggle whether to copy the locale files back to the 
'loaded library after translation.
'    If True, the newly-generated locale files will be copied back to library, 
'which is usually applied to auto-collect translations in a large number of mods.
'    If False (default), the library will keep locked.

'--------------------------------------------------------------------------------

Dim M : Set M = New Main

Class Main

    Private Sub Class_Initialize()
        Call RebootAsCscript(True)
        Call Main()
    End Sub

    Private Sub RebootAsCscript(ByRef blnHoldWindow)
        If LCase(Right(Wscript.FullName, 11)) = "cscript.exe" Then Exit Sub
        Dim arrCommand(), i
        ReDim arrCommand(1)
        arrCommand(0) = "%comspec% /c Color 3F & Cscript.exe //NoLogo"
        arrCommand(1) = Chr(34) & Wscript.ScriptFullName & Chr(34)
        If Wscript.Arguments.Count > 0 Then
            ReDim Preserve arrCommand(Wscript.Arguments.Count + 1)
            For i = 0 To Wscript.Arguments.Count - 1
                arrCommand(i + 2) = Chr(34) & Wscript.Arguments(i) & Chr(34)
            Next
        End If
        If blnHoldWindow Then
            ReDim Preserve arrCommand(UBound(arrCommand) + 1)
            arrCommand(UBound(arrCommand)) = "& Pause"
        End If
        CreateObject("Wscript.Shell").Run Join(arrCommand, " "), 5
        WScript.Quit
    End Sub

    Private Sub Main()
        Dim FSO, MT, strPath
        Set FSO = CreateObject("Scripting.FileSystemObject")
        Set MT = New MODs_Translator
        strPath = FSO.GetParentFolderName(WScript.ScriptFullName)
        Call MT.Load_Library(strPath, SCRIPT_VERSION, _
            Array(NAME_LIBRARY, TEXT_PRIORITY, UPDATE_LIBRARY))
        If Wscript.Arguments.Count = 0 Then
            Call MT.Scan_Paths(Array(FSO.BuildPath(strPath, "mods")))
            'Call MT.Generate_Mods_List(strPath)
        Else
            Call MT.Scan_Paths(Args_Array())
        End If
    End Sub

    Private Function Args_Array()
        Dim i, arrReturn
        If Wscript.Arguments.Count = 0 Then
            arrReturn = Array()
        Else
            Redim arrReturn(Wscript.Arguments.Count - 1)
            For i = 0 To Wscript.Arguments.Count - 1
                arrReturn(i) = Wscript.Arguments(i)
            Next
        End If
        Args_Array = arrReturn
    End Function

End Class

Class MODs_Translator

    Private NAME_LibInfo, NAME_ScrEcho, NAME_Log, NAME_List
    Private LABEL_DefGroup, LABEL_InvLine
    Private FZ, JA, ML, strLocale, objFolder_Lib_Root, arrOptions

    Private Sub Class_Initialize()
        NAME_LibInfo = "library-info.json"
        NAME_ScrEcho = "script-echo.json"
        NAME_Log = "fmlt-running-log.log"
        NAME_List = "mods-list.txt"
        LABEL_DefGroup = "[Default-Group]"
        LABEL_InvLine = "#Invalid-Line#"
        Set FZ = New FileSystem_throughZip
        Set JA = New JSON_Advanced
        Set ML = New Multilanguage_Echo
    End Sub

    Public Sub Load_Library(ByRef strPath_Work, ByRef intVer, ByRef arrOpt)
        Dim objFolder, arrFiles, objFile_Echo, dicInfo, dicEcho
        strLocale = Empty
        Set objFolder_Lib_Root = Nothing
        If Not FZ.FolderExists(objFolder, strPath_Work, arrOpt(0), False) Then
            MsgBox objFolder.Self.Path & vbCrlf & vbCrlf & _
                " does not exist.", 48, "Script Loading Error"
            WScript.Quit
        End If
        Set arrFiles = Locate_Files(objFolder, NAME_LibInfo, False, False)
        If arrFiles.Count = 0 Then
            '[skip line]
        ElseIf Not Valid_JSON(dicInfo, FZ.Read(arrFiles(0))) Then
            '[skip line]
        Elseif Not dicInfo.Exists("locale") Then
            '[skip line]
        ElseIf LCase(dicInfo("locale")) = "en" Then
            '[skip line]
        ElseIf Not FZ.FileExists(objFile_Echo, arrFiles(0).Parent, NAME_ScrEcho) Then
            '[skip line]
        ElseIf Not Valid_JSON(dicEcho, FZ.Read(objFile_Echo)) Then
            '[skip line]
        ElseIf ML.Initialize(dicEcho, strPath_Work, NAME_Log) Then
            strLocale = dicInfo("locale")
            Set objFolder_Lib_Root = arrFiles(0).Parent
            arrOptions = arrOpt
            Echo_Title dicInfo, intVer
        End If
        If IsEmpty(strLocale) Then
            MsgBox "This locale library has mistakes: " & vbCrlf & _
                vbCrlf & objFolder.Self.Path, 48, "Script Loading Error"
            WScript.Quit
        End If
    End Sub

    Private Sub Echo_Title(ByRef dicInfo, ByRef intVer)
        Dim dicOrg, dicMem, strEcho
        ML.Print -1, "*"
        ML.Print 0, ML.Echo("script_title", Array())
        ML.Print 0, ML.Echo("script_author", Array(intVer))
        ML.Print 0, ML.Echo("script_setting", Array(_
            FormatDateTime(Date, 2), FormatDateTime(Time, 4), _
            CStr(CBool(arrOptions(1))), CStr(CBool(arrOptions(2)))))
        ML.Print -1, "-"
        ML.Print 0, ML.Echo("library_name", Array(arrOptions(0)))
        ML.Print 0, ML.Echo("library_info", Array(dicInfo("locale"), _
            dicInfo("name"), dicInfo("version"), dicInfo("update")))
        If dicInfo.Exists("organization") Then
            Set dicOrg = dicInfo("organization")
            ML.Print 0, ML.Echo("library_member", _
                Array(dicOrg("name"), dicOrg("homepage")))
            For Each dicMem In dicOrg("members")
                If strEcho = "" Then
                    strEcho = ML.Echo("library_member", _
                        Array(dicMem("name"), dicMem("contact")))
                Else
                    ML.Print 0, strEcho & " " & ML.Echo("library_member", _
                    Array(dicMem("name"), dicMem("contact")))
                    strEcho = ""
                End If
            Next
            If strEcho <> "" Then ML.Print 0, strEcho
        End If
        ML.Print -1, "-"
        ML.Print 1, ML.Echo("tip_1", Array())
        ML.Print 1, ML.Echo("tip_2", Array())
        ML.Print 1, ML.Echo("tip_3", Array(strLocale))
        ML.Print -1, "*"
    End Sub

    Public Sub Scan_Paths(ByRef arrPaths)
        Dim i, objFolder, dicFiles, objFile, intCount, sngTime
        If IsEmpty(strLocale) Then Exit Sub
        sngTime = sngTime - Timer
        For i = 0 To UBound(arrPaths)
            ML.Print 1, ML.Echo("path_scanning", Array(arrPaths(i)))
            If Not FZ.FolderExists(objFolder, arrPaths(i), Null, False) Then
                ML.Print 2, ML.Echo("path_not_exist", Array())
            Else
                Set dicFiles = Locate_Files(objFolder, "info.json", True, True)
                If dicFiles.Count = 0 Then
                    ML.Print 2, ML.Echo("path_no_mod", Array())
                Else
                    For Each objFile In dicFiles.Items
                        intCount = intCount + Translate_Mod(objFile)
                    Next
                End If
            End If
        Next
        sngTime = sngTime + Timer
        ML.Print 1, ML.Echo("path_finished", Array(Int2Str(intCount, 1), _
            FormatNumber(sngTime, 2, -1)))
    End Sub

    Public Sub Generate_Mods_List(ByRef strPath_Work)
        Dim sngTime, dicFiles, objFile, dicInfo, dicOutput, intCount, objFolder
        If IsEmpty(strLocale) Then Exit Sub
        If Not FZ.FolderExists(objFolder, strPath_Work, Null, False) Then Exit Sub
        sngTime = sngTime - Timer
        ML.Print 1, ML.Echo("list_generating", Array(NAME_List))
        Set dicFiles = Locate_Files(objFolder_Lib_Root, "info.json", True, False)
        Set dicOutput = CreateObject("Scripting.Dictionary")
        For Each objFile In dicFiles.Items
            If Valid_JSON(dicInfo, FZ.Read(objFile)) Then
                dicOutput.Add dicOutput.Count, ML.Echo("list_format", _
                    Array(dicInfo("name"), dicInfo("title"), dicInfo("description")))
                intCount = intCount + 1
            End If
        Next
        Call FZ.Write(objFolder, NAME_List, Join(dicOutput.Items, vbCrlf))
        sngTime = sngTime + Timer
        ML.Print 1, ML.Echo("list_finish", Array(Int2Str(intCount, 1), _
            FormatNumber(sngTime, 2, -1)))
    End Sub

    Private Function Translate_Mod(ByRef objFile_Info)
        Dim objFolder_Lib, dicInfo, arrCount(1)
        If Not Valid_JSON(dicInfo, FZ.Read(objFile_Info)) Then Exit Function
        If Len(dicInfo("name")) = 0 Then Exit Function
        ML.Print 2, ML.Echo("trans_start", Array(dicInfo("name"), dicInfo("version")))
        Call FZ.FolderExists(objFolder_Lib, objFolder_Lib_Root, _
                dicInfo("name"), arrOptions(2))
        Add_Arr arrCount, Translate_Info(objFile_Info.Parent, objFolder_Lib, dicInfo)
        Add_Arr arrCount, Translate_Locale(objFile_Info.Parent, objFolder_Lib)
        Add_Arr arrCount, Translate_Script_Locale(objFile_Info.Parent, objFolder_Lib)
        If arrCount(0) = 0 Then
            ML.Print 3, ML.Echo("trans_none", Array(Int2Str(0, 3), Int2Str(0, 3)))
        Else
            ML.Print 3, ML.Echo("trans_finished", Array(_
                Int2Str(arrCount(1), 3), Int2Str(arrCount(0), 3), _
                FormatPercent(arrCount(1)/arrCount(0), 1, -1, 0)))
            Translate_Mod = 1
        End If
    End Function

    Private Function Translate_Info(ByRef objFolder_Mod_Root, ByRef objFolder_Lib, _
            ByRef dicInfo_inMod)
        Dim objFile, dicInfo_inLib, dicInfo_New, strKey, strKey_ori, arrCount(1)
        If Not arrOptions(1) And FZ.FileExists(objFile, objFolder_Lib, "info.json") Then
            Call Valid_JSON(dicInfo_inLib, FZ.Read(objFile))
        Else
            Set dicInfo_inLib = CreateObject("Scripting.Dictionary")
        End If
        Set dicInfo_New = CreateObject("Scripting.Dictionary")
        For Each strKey In dicInfo_inMod.Keys
            Select Case strKey
                Case "title", "description"
                    If Len(dicInfo_inLib(strKey)) > 0 Then
                        dicInfo_New(strKey) = dicInfo_inLib(strKey)
                    Else
                        dicInfo_New(strKey) = dicInfo_inMod(strKey)
                    End If
                    strKey_ori = strKey & "_original"
                    If Len(dicInfo_inMod(strKey_ori)) > 0 Then
                        dicInfo_New(strKey_ori) = dicInfo_inMod(strKey_ori)
                    Else
                        dicInfo_New(strKey_ori) = dicInfo_inMod(strKey)
                    End If
                Case "title_original", "description_original"
                    '[skip line]
                Case Else
                    dicInfo_New(strKey) = dicInfo_inMod(strKey)
            End Select
        Next
        Call FZ.Write(objFolder_Mod_Root, "info.json", JA.EncodeJSON(dicInfo_New, 1))
        If arrOptions(2) Then
            Call FZ.FileExists(objFile, objFolder_Mod_Root, "info.json")
            Call FZ.Copy(objFile, objFolder_Lib, True)
        End If
        arrCount(0) = 2
        arrCount(1) = - (dicInfo_New("title") <> dicInfo_New("title_original")) _
            - (dicInfo_New("description") <> dicInfo_New("description_original"))
        ML.Print 3, ML.Echo("trans_file", _
            Array(Int2Str(arrCount(1), 3), Int2Str(arrCount(0), 3), "\info.json"))
        Translate_Info = arrCount
    End Function

    Private Function Translate_Locale(ByRef objFolder_Mod_Root, ByRef objFolder_Lib)
        Dim objFolder, objFolder_en, objFolder_Loc, objFile_en, objFile
        Dim strName, dicItems, arrResult, arrCount(1)
        If Not FZ.FolderExists(objFolder, objFolder_Mod_Root, "locale", False) Then
            '[skip line]
        ElseIf Not FZ.FolderExists(objFolder_en, objFolder, "en", False) Then
            '[skip line]
        Else
            Call FZ.FolderExists(objFolder_Loc, objFolder, strLocale, True)
            For Each objFile_en In objFolder_en.Items
                strName = FZ.Name(objFile_en)
                If objFile_en.IsFolder Then
                    '[skip line]
                ElseIf LCase(Right(strName, 3)) = "cfg" Then
                    Set dicItems = Load_Items(objFile_en)
                    If arrOptions(1) Then
                        Update_Items dicItems, objFolder_Lib, strName
                        Update_Items dicItems, objFolder_Loc, strName
                    Else
                        Update_Items dicItems, objFolder_Loc, strName
                        Update_Items dicItems, objFolder_Lib, strName
                    End If
                    arrResult = Write_Items(dicItems, objFolder_Loc, strName)
                    If arrOptions(2) Then
                        Call FZ.FileExists(objFile, objFolder_Loc, strName)
                        Call FZ.Copy(objFile, objFolder_Lib, True)
                    End If
                    ML.Print 3, ML.Echo("trans_file", Array(_
                        Int2Str(arrResult(1), 3), Int2Str(arrResult(0), 3), _
                        "\locale\" & strLocale & "\" & strName))
                    Add_Arr arrCount, arrResult
                End If
            Next
        End If
        Translate_Locale = arrCount
    End Function

    Private Function Translate_Script_Locale(ByRef objFolder_Mod_Root, _
            ByRef objFolder_Lib)
        Dim objFolder_Loc, objFile, dicItems, strName, arrResult, arrCount(1)
        If Not FZ.FolderExists(objFolder_Loc, objFolder_Mod_Root, _
                "script-locale", False) Then
            '[skip line]
        ElseIf Not FZ.FileExists(objFile, objFolder_Loc, "en.cfg") Then
            '[skip line]
        Else
            strName = strLocale & ".cfg"
            Set dicItems = Load_Items(objFile)
            If arrOptions(1) Then
                Update_Items dicItems, objFolder_Lib, strName
                Update_Items dicItems, objFolder_Loc, strName
            Else
                Update_Items dicItems, objFolder_Loc, strName
                Update_Items dicItems, objFolder_Lib, strName
            End If
            arrResult = Write_Items(dicItems, objFolder_Loc, strName)
            If arrOptions(2) Then
                Call FZ.FileExists(objFile, objFolder_Loc, strName)
                Call FZ.Copy(objFile, objFolder_Lib, True)
            End If
            ML.Print 3, ML.Echo("trans_file", Array(_
                Int2Str(arrResult(1), 3), Int2Str(arrResult(0), 3), _
                "\script-locale\" & strName))
            Add_Arr arrCount, arrResult
        End If
        Translate_Script_Locale = arrCount
    End Function

    Private Function Load_Items(ByRef objFile_Target)
        Dim dicItems, arrLines, strGroup, dicGroups, i, intPos
        Set dicItems = CreateObject("Scripting.Dictionary")
        strGroup = LABEL_DefGroup
        Redim dicGroups(0)
        Set dicGroups(0) = CreateObject("Scripting.Dictionary")
        dicItems.Add strGroup, dicGroups(0)
        arrLines = Text2Lines(FZ.Read(objFile_Target))
        For i = 0 To UBound(arrLines)
            intPos = InStr(arrLines(i), "=")
            If intPos > 0 Then
                dicItems(strGroup).Add Trim(Left(arrLines(i), intPos-1)), _
                    Array(Trim(Mid(arrLines(i), intPos+1)), "")
            ElseIf Left(arrLines(i), 1) = "[" And Right(arrLines(i), 1) = "]" Then
                strGroup = arrLines(i)
                If Not dicItems.Exists(arrLines(i)) Then
                    Redim dicGroups(dicItems.Count)
                    Set dicGroups(dicItems.Count) = CreateObject("Scripting.Dictionary")
                    dicItems.Add arrLines(i), dicGroups(dicItems.Count)
                End If
            Else
                dicItems(strGroup).Add LABEL_InvLine & i, arrLines(i)
            End If
        Next
        Set Load_Items = dicItems
    End Function

    Private Sub Update_Items(ByRef dicItems, ByRef objFolder_Target, ByRef strName_cfg)
        Dim objFile, strGroup, arrLines, i, intPos, strKey, arrText
        If Not FZ.FileExists(objFile, objFolder_Target, strName_cfg) Then Exit Sub
        arrLines = Text2Lines(FZ.Read(objFile))
        strGroup = LABEL_DefGroup
        For i = 0 To UBound(arrLines)
            intPos = InStr(arrLines(i), "=")
            If intPos > 0 And Len(strGroup) > 0 Then
                strKey = Trim(Left(arrLines(i), intPos-1))
                If dicItems(strGroup).Exists(strKey) Then
                    arrText = dicItems(strGroup)(strKey)
                    arrText(1) = Trim(Mid(arrLines(i), intPos+1))
                    If Len(arrText(1)) > 0 Then dicItems(strGroup)(strKey) = arrText
                End If
            ElseIf Left(arrLines(i), 1) = "[" And Right(arrLines(i), 1) = "]" Then
                strGroup = ""
                If dicItems.Exists(arrLines(i)) Then strGroup = arrLines(i)
            End If
        Next
    End Sub

    Private Function Write_Items(ByRef dicItems, ByRef objFolder_Target, _
            ByRef strName_cfg)
        Dim dicLines, strGroup, strKey, arrText, arrCount(1)
        Set dicLines = CreateObject("Scripting.Dictionary")
        For Each strGroup In dicItems.Keys
            dicLines.Add dicLines.Count, strGroup
            For Each strKey In dicItems(strGroup).Keys
                If Left(strKey, Len(LABEL_InvLine)) = LABEL_InvLine Then
                    dicLines.Add dicLines.Count, dicItems(strGroup)(strKey)
                Else
                    arrCount(0) = arrCount(0) + 1   'Amount of total items
                    arrText = dicItems(strGroup)(strKey)
                    If Len(arrText(1)) = 0 Then
                        dicLines.Add dicLines.Count, strKey & " = " & arrText(0)
                    ElseIf arrText(0) = arrText(1) Then
                        dicLines.Add dicLines.Count, strKey & " = " & arrText(0)
                    Else
                        arrCount(1) = arrCount(1) + 1   'Amount of translated items
                        dicLines.Add dicLines.Count, strKey & " = " & arrText(1)
                    End If
                End If
            Next
        Next
        dicLines.Remove(0)
        Call FZ.Write(objFolder_Target, strName_cfg, Join(dicLines.Items, vbCrlf))
        Write_Items = arrCount
    End Function

    Private Function Locate_Files(ByRef objFolder_Root, ByRef strName_File, _
            ByRef blnFindAll, ByRef blnAvoidLib)
        Dim arrReturn, blnExec, objFile, objItem
        Set arrReturn = CreateObject("Scripting.Dictionary")
        If Not blnAvoidLib Then
            blnExec = True
        ElseIf objFolder_Lib_Root Is Nothing Then
            blnExec = True
        ElseIf InStr(objFolder_Root.Self.Path, objFolder_Lib_Root.Self.Path) = 0 Then
            blnExec = True
        End If
        If blnExec Then
            If FZ.FileExists(objFile, objFolder_Root, strName_File) Then
                arrReturn.Add arrReturn.Count, objFile
            End If
            If Not blnFindAll And arrReturn.Count > 0 Then
                '[skip line]
            Else
                For Each objItem In objFolder_Root.Items
                    If objItem.IsFolder Then
                        For Each objFile In Locate_Files(objItem.GetFolder, _
                                strName_File, blnFindAll, blnAvoidLib).Items
                            arrReturn.Add arrReturn.Count, objFile
                        Next
                    End If
                Next
            End If
        End If
        Set Locate_Files = arrReturn
    End Function

    Private Function Text2Lines(ByRef strText)
        Dim Reg, objMatches, i, arrReturn()
        strText = Replace(strText, vbCrlf, vbLf)
        strText = Replace(strText, vbCr, vbLf)
        Set Reg = New RegExp
        Reg.Global = True
        Reg.MultiLine = True
        Reg.Pattern = "^.*$"
        Set objMatches = Reg.Execute(strText)
        ReDim arrReturn(objMatches.Count - 1)
        For i = 0 To objMatches.Count - 1
            arrReturn(i) = Trim(objMatches(i).Value)
        Next
        Text2Lines = arrReturn
    End Function

    Private Function Valid_JSON(ByRef dicReturn, ByRef strText)
        On Error Resume Next
            Err.Clear
            JA.DecodeJSON strText, dicReturn
            If Err.Number <> 0 Then dicReturn = Empty
        On Error Goto 0
        If TypeName(dicReturn) <> "Dictionary" Then
            Set dicReturn = CreateObject("Scripting.Dictionary")
        End If
        Valid_JSON = (dicReturn.Count > 0)
    End Function

    Private Function Int2Str(ByVal intNum, ByVal intDigit)
        intNum = CStr(CInt(intNum))
        intDigit = intDigit - Len(intNum)
        Int2Str = Space(-(intDigit>0)*intDigit) & intNum
    End Function

    Private Sub Add_Arr(ByRef arrBase, ByRef arrAdd)
        Dim i
        For i = 0 To UBound(arrBase)
            arrBase(i) = arrBase(i) + arrAdd(i)
        Next
    End Sub

End Class

Class Multilanguage_Echo
    
    Private dicText, strLang(1), intWidth, objOutput

    Public Function Initialize(ByRef dicLocale, ByRef strPath, ByRef strName_Log)
        Set dicText = Nothing
        If Not dicLocale.Exists("echo_text") Then
            '[skip line]
        ElseIf Not dicLocale.Exists("echo_setting") Then
            '[skip line]
        ElseIf Not dicLocale("echo_setting").Exists("default_language") Then
            '[skip line]
        ElseIf Not dicLocale("echo_setting").Exists("applied_language") Then
            '[skip line]
        ElseIf Not dicLocale("echo_setting").Exists("line_width") Then
            '[skip line]
        Else
            strLang(0) = dicLocale("echo_setting")("default_language")
            strLang(1) = dicLocale("echo_setting")("applied_language")
            intWidth = dicLocale("echo_setting")("line_width")
            Set dicText = dicLocale("echo_text")
        End If
        If Not dicText Is Nothing Then
            Dim FSO
            Set FSO = CreateObject("Scripting.FileSystemObject")
            Set objOutput = FSO.CreateTextFile(FSO.BuildPath(strPath, strName_Log), True)
            Initialize = True
        End If
    End Function

    Public Function Echo(ByRef strName, ByRef arrSep)
        Dim i
        If Not dicText.Exists(strName) Then
            Echo = "[Echo Missing: " & strName & "]"
        ElseIf dicText(strName).Exists(strLang(1)) Then
            Echo = dicText(strName)(strLang(1))
        ElseIf dicText(strName).Exists(strLang(0)) Then
            Echo = dicText(strName)(strLang(0))
        Else
            Echo = "[Echo Missing: " & strName & "]"
        End If
        For i = 0 To UBound(arrSep)
            Echo = Replace(Echo, "__" & CStr(i+1) & "__", arrSep(i))
        Next
    End Function

    Public Sub Print(ByRef intType, ByVal strEcho)
        Dim intLen, intPos(1), strChar, intTmp
        If intType >= 1 Then
            'Left align with levels
            strEcho = Replace(strEcho, vbCrlf, vbLf)
            strEcho = Replace(strEcho, vbCr, vbLf)
            strEcho = Space(3*(intType-1)) & ">> " & _
                Replace(strEcho, vbLf, vbLf & Space(3*intType))
            intPos(1) = 1
            Do While intPos(1) <= Len(strEcho)
                strChar = Mid(strEcho, intPos(1), 1)
                If strChar = vbTab Then
                    intTmp = 5
                Else
                    intTmp = LenW(strChar)
                End If
                If strChar = vbLf Then
                    intPos(0) = intPos(1) + 1
                    intPos(1) = intPos(0) + 3*intType
                    intLen = 3*intType
                ElseIf intLen + intTmp > intWidth Then
                    strEcho = Left(strEcho, intPos(1)-1) & vbLf & _
                        Space(3*intType) & Mid(strEcho, intPos(1))
                    intPos(0) = intPos(1) + 3*intType + 1
                    intPos(1) = intPos(0) + 3*intType
                    intLen = 3*intType
                Else
                    intLen = intLen + intTmp
                    intPos(1) = intPos(1) + 1
                End If
            Loop
            strEcho = Replace(strEcho, vbLf, vbCrlf)
        ElseIf intType = 0 Then
            'Center align
            intTmp = intWidth - LenW(strEcho)
            If intTmp < 0 Then intTmp = 0
            strEcho = Space(Int(intTmp/2)) & strEcho
        ElseIf intType = -1 Then
            'Fill a row
            If LenW(strEcho) > 0 Then strEcho = _
                Left(String(Int(intWidth/LenW(strEcho))+1, strEcho), intWidth)
        End If
        Wscript.Echo strEcho
        objOutput.WriteLine strEcho
    End Sub

    Private Function LenW(ByRef strText)
        Dim i, intReturn
        For i = 2 To LenB(strText) Step 2
            If AscB(MidB(strText, i, 1)) > 0 Then
                intReturn = intReturn + 2
            Else
                intReturn = intReturn + 1
            End If
        Next
        LenW = intReturn
    End Function

End Class

Class FileSystem_throughZip
    'Author: Mr.Jos(sd7056333)
    'Date: 2016/07/13
    'Website: http://blog.csdn.net/sd7056333
    'Description: 
    '   This class offer some methods for file & folder operations which can 
    '   see the Zip files as common folders.
    '   The variable types of the files & folders in this class are inherited 
    '   from Windows Shell, which are indicated by the prefixes of variable 
    '   names of those parameters in each functions. This also allows you to 
    '   use the methods and properties offered by Windows Shell.
    '   The following table shows which variable types the prefixes represent: 
    '   +-------------+------------------+---------------------+
    '   | Prefix      | Represent        | Variable Types      |
    '   +=============+==================+=====================+
    '   | objFolder   | Folder           |    Folder3          |
    '   +-------------+------------------+---------------------+
    '   | objFile     | File             |    FolderItem2      |
    '   +-------------+------------------+---------------------+
    '   | varFolder   | Folder           |    Folder3          |
    '   +-------------+------------------+ Or FolderItem2      |
    '   | varItem     | File Or Folder   | Or String(Path)     |
    '   +-------------+------------------+---------------------+
    'Main References:
    '   https://msdn.microsoft.com/en-us/library/windows/desktop/bb787868
    '   https://msdn.microsoft.com/en-us/library/windows/desktop/bb787810

    Private FSO, Shell, objFolder_Work, TE

    Private Sub Class_Initialize()
        Dim strPath_Work
        Set FSO = CreateObject("Scripting.FileSystemObject")
        Set Shell = CreateObject("Shell.Application")
        strPath_Work = FSO.BuildPath(FSO.GetSpecialFolder(2), FSO.GetTempName)
        If FSO.FolderExists(strPath_Work) Then FSO.DeleteFolder strPath_Work, True
        FSO.CreateFolder strPath_Work
        Set objFolder_Work = Shell.NameSpace(strPath_Work)
        '[Custom Class] Text reading & wrting class
        Set TE = New Text_Encode
    End Sub

    Private Sub Class_Terminate()
        FSO.DeleteFolder objFolder_Work.Self.Path, True
    End Sub

    Public Function FolderExists(ByRef objFolder_Return, ByRef varFolder_Parent, _
            ByRef strName_SubFolder, ByRef blnCreate)
        Dim objFolder_Parent, strPath_Target, strPath
        Set objFolder_Parent = GetObject_Folder3(varFolder_Parent)
        If IsNull(strName_SubFolder) Or strName_SubFolder = "" Then
            Set objFolder_Return = objFolder_Parent
        ElseIf objFolder_Parent Is Nothing Then
            Set objFolder_Return = Nothing
        Else
            strPath_Target = FSO.BuildPath(objFolder_Parent.Self.Path, strName_SubFolder)
            If FSO.FolderExists(objFolder_Parent.Self.Path) Then
                If Not blnCreate Then
                    '[skip line]
                ElseIf FSO.FolderExists(strPath_Target) Then
                    '[skip line]
                ElseIf FSO.FileExists(strPath_Target) Then
                    '[skip line]
                ElseIf LCase(FSO.GetExtensionName(strPath_Target)) = "zip" Then
                    With FSO.CreateTextFile(strPath_Target, True)
                        .Write "PK" & Chr(5) & Chr(6) & String(18, Chr(0))
                        .Close
                    End With
                Else
                    FSO.CreateFolder strPath_Target
                End If
                Set objFolder_Return = GetObject_Folder3(strPath_Target)
            Else
                If Not objFolder_Parent.ParseName(strName_SubFolder) Is Nothing Then
                    Set objFolder_Return = GetObject_Folder3(strPath_Target)
                ElseIf blnCreate Then
                    strPath = FSO.BuildPath(objFolder_Work.Self.Path, strName_SubFolder)
                    Delete strPath
                    FSO.CreateFolder strPath
                    FSO.CreateTextFile FSO.BuildPath(strPath, "Tmp")
                    Copy strPath, objFolder_Parent, True
                    Set objFolder_Return = Shell.NameSpace(strPath_Target)
                    Delete objFolder_Return.ParseName("Tmp")
                    Delete strPath
                Else
                    Set objFolder_Return = Nothing
                End If
            End If
        End If
        FolderExists = Not objFolder_Return Is Nothing
    End Function

    Public Function FileExists(ByRef objFile_Return, ByRef objFolder_Parent, _
            ByRef strName_File)
        Dim objItem
        Set objFile_Return = Nothing
        Select Case TypeName(objFolder_Parent)
            Case "Folder3"
                Set objItem = objFolder_Parent.ParseName(strName_File)
                If objItem Is Nothing Then
                    '[skip line]
                ElseIf Not objItem.IsFolder Then
                    Set objFile_Return = objItem
                End If
            Case "Nothing"
                '[skip line]
            Case Else
                Err.Raise 8731, "File Operation Error", "Invalid parameter type"
        End Select
        FileExists = Not objFile_Return Is Nothing
    End Function

    Public Sub Copy(ByRef varItem_Source, ByRef objFolder_Destination, _
            ByRef blnOverride)
        Dim objItem_Source
        Select Case TypeName(objFolder_Destination)
            Case "Folder3"
                Set objItem_Source = GetObject_FolderItem2(varItem_Source)
                If objItem_Source Is Nothing Then
                    '[skip line]
                ElseIf FSO.FolderExists(objItem_Source.Parent.Self.Path) Or _
                        FSO.FolderExists(objFolder_Destination.Self.Path) Then
                    CopyItem objItem_Source, objFolder_Destination, blnOverride
                Else
                    If blnOverride Then Delete FSO.BuildPath(_
                        objFolder_Destination.Self.Path, Name(objItem_Source))
                    Delete FSO.BuildPath(objFolder_Work.Self.Path, Name(objItem_Source))
                    CopyItem objItem_Source, objFolder_Work, False
                    CopyItem objFolder_Work.ParseName(Name(objItem_Source)), _
                        objFolder_Destination, blnOverride
                    Delete FSO.BuildPath(objFolder_Work.Self.Path, Name(objItem_Source))
                End If
            Case "Nothing"
                '[skip line]
            Case Else
                Err.Raise 8731, "File Operation Error", "Invalid parameter type"
        End Select
    End Sub

    Public Sub Delete(ByRef varItem_Delete)
        Dim objItem_Delete, strPath, intCount
        Set objItem_Delete = GetObject_FolderItem2(varItem_Delete)
        If objItem_Delete Is Nothing Then
            '[skip line]
        ElseIf FSO.FolderExists(objItem_Delete.Path) Then
            FSO.DeleteFolder objItem_Delete.Path, True
        ElseIf FSO.FileExists(objItem_Delete.Path) Then
            FSO.DeleteFile objItem_Delete.Path, True
        Else
            strPath = FSO.BuildPath(objFolder_Work.Self.Path, Name(objItem_Delete))
            If FSO.FileExists(strPath) Then FSO.DeleteFile strPath, True
            If FSO.FolderExists(strPath) Then FSO.DeleteFolder strPath, True
            intCount = objItem_Delete.Parent.Items.Count
            objFolder_Work.MoveHere objItem_Delete, 4 + 16
            Do
                WScript.Sleep 2
                If objItem_Delete.Parent.Items.Count = intCount - 1 Then
                    Exit Do
                ElseIf Not objItem_Delete.IsFolder Then
                    '[skip line]
                ElseIf objItem_Delete.GetFolder.Items.Count = 0 Then
                    Exit Do
                End If
            Loop
            If FSO.FileExists(strPath) Then FSO.DeleteFile strPath, True
            If FSO.FolderExists(strPath) Then FSO.DeleteFolder strPath, True
        End If
    End Sub

    Public Function Read(ByRef objFile_Read)
        Dim strPath
        If TypeName(objFile_Read) <> "FolderItem2" Then
            Err.Raise 8731, "File Operation Error", "Invalid parameter type"
        ElseIf objFile_Read.IsFolder Then
            '[skip line]
        ElseIf FSO.FolderExists(objFile_Read.Parent.Self.Path) Then
            '[Custom Function] Read text file
            Read = TE.ReadText(objFile_Read.Path, Null)
        Else
            Copy objFile_Read, objFolder_Work, True
            strPath = FSO.BuildPath(objFolder_Work.Self.Path, Name(objFile_Read))
            '[Custom Function] Read text file
            Read = TE.ReadText(strPath, Null)
            Delete strPath
        End If
    End Function

    Public Sub Write(ByRef objFolder_Parent, ByRef strName_File, ByRef strText_Write)
        Dim strPath
        If TypeName(objFolder_Parent) <> "Folder3" Then
            Err.Raise 8731, "File Operation Error", "Invalid parameter type"
        ElseIf FSO.FolderExists(objFolder_Parent.Self.Path) Then
            strPath = FSO.BuildPath(objFolder_Parent.Self.Path, strName_File)
            Delete strPath
            '[Custom Function] Write text file
            TE.WriteText strPath, strText_Write, "UTF-8"
        Else
            Delete FSO.BuildPath(objFolder_Parent.Self.Path, strName_File)
            strPath = FSO.BuildPath(objFolder_Work.Self.Path, strName_File)
            Delete strPath
            '[Custom Function] Write text file
            TE.WriteText strPath, strText_Write, "UTF-8"
            Copy strPath, objFolder_Parent, False
            Delete strPath
        End If
    End Sub

    Public Function Name(ByRef varItem)
        Select Case TypeName(varItem)
            Case "Folder3"
                Name = varItem.Name
            Case "FolderItem2"
                Name = FSO.GetFileName(varItem.Path)
            Case Else
                Name = FSO.GetFileName(varItem)
        End Select 
    End Function

    Private Function GetObject_FolderItem2(ByRef varItem)
        Dim objFolder, strPath
        Set GetObject_FolderItem2 = Nothing
        Select Case TypeName(varItem)
            Case "Folder3"
                Set GetObject_FolderItem2 = varItem.Self
            Case "FolderItem2"
                Set GetObject_FolderItem2 = varItem
            Case Else
                If Len(CStr(varItem)) = 0 Then Exit Function
                strPath = FSO.GetParentFolderName(CStr(varItem))
                If Len(strPath) > 0 Then
                    Set objFolder = GetObject_Folder3(strPath)
                    If Not objFolder Is Nothing Then
                        Set GetObject_FolderItem2 = _
                            objFolder.ParseName(FSO.GetFileName(CStr(varItem)))
                    End If
                End If
        End Select
    End Function

    Private Function GetObject_Folder3(ByRef varFolder)
        Select Case TypeName(varFolder)
            Case "Folder3"
                Set GetObject_Folder3 = varFolder
            Case "FolderItem2"
                If varFolder.IsFolder Then
                    Set GetObject_Folder3 = varFolder.GetFolder
                Else
                    Set GetObject_Folder3 = Nothing
                End If
            Case Else
                If LCase(FSO.GetExtensionName(CStr(varFolder))) <> "zip" _
                        And FSO.FileExists(CStr(varFolder)) Then
                    Set GetObject_Folder3 = Nothing
                Else
                    Set GetObject_Folder3 = Shell.NameSpace(CStr(varFolder))
                    If GetObject_Folder3 Is Nothing Then
                        '[skip line]
                    ElseIf Not GetObject_Folder3.Self.IsFolder Then
                        Set GetObject_Folder3 = Nothing
                    End If
                End If
        End Select
    End Function

    Private Sub CopyItem(ByRef objItem_Source, ByRef objFolder_Destination, _
            ByRef blnOverride)
        Dim objItem_Destination, blnExec, objItem, intCount
        Set objItem_Destination = objFolder_Destination.ParseName(Name(objItem_Source))
        If objItem_Source.IsFolder Then
            If objItem_Source.GetFolder.Items.Count = 0 Then
                '[skip line]
            ElseIf objItem_Destination Is Nothing Then
                blnExec = True
            ElseIf objItem_Destination.IsFolder Then
                For Each objItem In objItem_Source.GetFolder.Items
                    CopyItem objItem, objItem_Destination.GetFolder
                Next
            ElseIf blnOverride Then
                Delete objItem_Destination
                blnExec = True
            End If
        Else
            If objItem_Destination Is Nothing Then
                blnExec = True
            ElseIf blnOverride Then
                Delete objItem_Destination
                blnExec = True
            End If
        End If
        If blnExec Then
            intCount = objFolder_Destination.Items.Count
            objFolder_Destination.CopyHere objItem_Source, 4 + 16
            Do
                WScript.Sleep 2
            Loop Until objFolder_Destination.Items.Count = intCount + 1
        End If
    End Sub

End Class

Class Text_Encode
    'Author: Mr.Jos(sd7056333)
    'Date: 2016/04/23
    'Website: http://blog.csdn.net/sd7056333
    'Description: 
    '   This is a class for identifying/reading/writing text files of various 
    '   encoding character sets.
    '   The followling table shows the charsets that can be distinguished.
    '   +-----------------------+---------------------------------------+
    '   | Charset               | Identifiable Other Names              |
    '   +=======================+=======================================+
    '   | Unicode Little Endian | Unicode, Unicode LE, UTF-16, UTF-16LE |
    '   +-----------------------+---------------------------------------+
    '   | Unicode Big Endian    | Unicode BE, UTF-16BE                  |
    '   +-----------------------+---------------------------------------+
    '   | UTF-8 without BOM     | UTF-8                                 |
    '   +-----------------------+---------------------------------------+
    '   | UTF-8 with BOM        |                                       |
    '   +-----------------------+---------------------------------------+
    '   | ANSI                  | (ANSI Charset: GB2312/Big5/...)       |
    '   +-----------------------+---------------------------------------+
    'Main References:
    '   http://demon.tw/programming/vbs-file-unicode.html
    '   http://demon.tw/programming/vbs-validate-utf8.html
    '   http://www.ruanyifeng.com/blog/2007/10/ascii_unicode_and_utf-8.html
    
    Private ADO_Bin, ADO_Text, Reg_UTF8
    Public ANSI_Charset
    
    Private Sub Class_Initialize()
        Set ADO_Bin = CreateObject("ADODB.Stream")
        ADO_Bin.Mode = 3
        ADO_Bin.Type = 1
        Set ADO_Text = CreateObject("ADODB.Stream")
        ADO_Text.Mode = 3
        ADO_Text.Type = 2
        Set Reg_UTF8 = New Regexp
        Reg_UTF8.Pattern = Join(Array("", _
                "[\xC0-\xDF]([^\x80-\xBF]|$)", _
                "|[\xE0-\xEF].{0,1}([^\x80-\xBF]|$)", _
                "|[\xF0-\xF7].{0,2}([^\x80-\xBF]|$)", _
                "|[\xF8-\xFB].{0,3}([^\x80-\xBF]|$)", _
                "|[\xFC-\xFD].{0,4}([^\x80-\xBF]|$)", _
                "|[\xFE-\xFE].{0,5}([^\x80-\xBF]|$)", _
                "|[\x00-\x7F][\x80-\xBF]", _
                "|[\xC0-\xDF].[\x80-\xBF]", _
                "|[\xE0-\xEF]..[\x80-\xBF]", _
                "|[\xF0-\xF7]...[\x80-\xBF]", _
                "|[\xF8-\xFB]....[\x80-\xBF]", _
                "|[\xFC-\xFD].....[\x80-\xBF]", _
                "|[\xFE-\xFE]......[\x80-\xBF]", _
                "|^[\x80-\xBF]"), "")
        ANSI_Charset = "GB2312"     'Default ANSI charset
    End Sub
    
    Public Function IdentifyCharset(ByRef strPath)
        Dim bytTest, arrBin, i
        ADO_Bin.Open
        ADO_Bin.LoadFromFile strPath
        bytTest = ADO_Bin.Read(3)
        If BinComp(bytTest, Array(&HEF, &HBB, &HBF)) Then
            IdentifyCharset = "UTF-8 with BOM"
        ElseIf BinComp(bytTest, Array(&HFF, &HFE)) Then
            IdentifyCharset = "Unicode Little Endian"
        ElseIf BinComp(bytTest, Array(&HFE, &HFF)) Then
            IdentifyCharset = "Unicode Big Endian"
        Else
            ADO_Bin.Position = 0
            ReDim arrBin(ADO_Bin.Size - 1)
            For i = 0 To ADO_Bin.Size - 1
                arrBin(i) = ChrW(AscB(ADO_Bin.Read(1)))
            Next
            If Not Reg_UTF8.Test(Join(arrBin, "")) Then
                IdentifyCharset = "UTF-8 without BOM"
            Else
                IdentifyCharset = "ANSI"
            End If
        End If
        ADO_Bin.Close
    End Function
    
    Public Function ReadText(ByRef strPath, ByRef strCharset)
        If IsNull(strCharset) Then strCharset = IdentifyCharset(strPath)
        Select Case ModifyCharset(strCharset)
            Case "Unicode Little Endian", "Unicode Big Endian"
                ADO_Text.Charset = "Unicode"
            Case "UTF-8 without BOM", "UTF-8 with BOM"
                ADO_Text.Charset = "UTF-8"
            Case "ANSI"
                ADO_Text.Charset = ANSI_Charset
            Case Else
                ADO_Text.Charset = strCharset
        End Select
        ADO_Text.Open
        ADO_Text.LoadFromFile strPath
        ReadText = ADO_Text.ReadText
        ADO_Text.Close
    End Function
    
    Public Sub WriteText(ByRef strPath, ByRef strText, ByRef strCharset)
        Dim blnSkipCommonWrite
        Select Case ModifyCharset(strCharset)
            Case "Unicode Little Endian"
                ADO_Text.Charset = "Unicode"
            Case "Unicode Big Endian"
                blnSkipCommonWrite = True
                Call WriteAsUnicodeBigEndian(strPath, strText)
            Case "UTF-8 without BOM"
                blnSkipCommonWrite = True
                Call WriteAsUTF8noBOM(strPath, strText)
            Case "UTF-8 with BOM"
                ADO_Text.Charset = "UTF-8"
            Case "ANSI"
                ADO_Text.Charset = ANSI_Charset
            Case Else
                ADO_Text.Charset = strCharset
        End Select
        If Not blnSkipCommonWrite Then
            ADO_Text.Open
            ADO_Text.WriteText strText
            ADO_Text.SaveToFile strPath, 2
            ADO_Text.Close
        End If
    End Sub
    
    Private Function ModifyCharset(strCharset)
        Select Case Trim(LCase(strCharset))
            Case "unicode little endian", "unicode le", "unicode", "utf-16le", "utf-16"
                ModifyCharset = "Unicode Little Endian"
            Case "unicode big endian", "unicode be", "utf-16be"
                ModifyCharset = "Unicode Big Endian"
            Case "utf-8 without bom", "utf-8"
                ModifyCharset = "UTF-8 without BOM"
            Case "utf-8 with bom"
                ModifyCharset = "UTF-8 with BOM"
            Case "ansi"
                ModifyCharset = "ANSI"
            Case Else
                ModifyCharset = strCharset
        End Select
    End Function
    
    Private Sub WriteAsUTF8noBOM(ByRef strPath, ByRef strText)
        ADO_Bin.Open()
        With ADO_Text
            .Charset = "UTF-8"
            .Open
            .WriteText strText
            .Position = 3
            .CopyTo ADO_Bin
            .Close
        End With
        ADO_Bin.SaveToFile strPath, 2
        ADO_Bin.Close
    End Sub
    
    Private Sub WriteAsUnicodeBigEndian(ByRef strPath, ByRef strText)
        Dim i, bytTrans(1)
        ADO_Bin.Open()
        With ADO_Text
            .Charset = "Unicode"
            .Open
            .WriteText strText
            .Position = 0
            .CopyTo ADO_Bin
            .Close
        End With
        ADO_Bin.Position = 0
        For i = 1 To ADO_Bin.Size Step 2
            bytTrans(0) = ADO_Bin.Read(1)
            bytTrans(1) = ADO_Bin.Read(1)
            ADO_Bin.Position = i - 1
            ADO_Bin.Write bytTrans(1)
            ADO_Bin.Write bytTrans(0)
        Next
        ADO_Bin.SaveToFile strPath, 2
        ADO_Bin.Close
    End Sub
    
    Private Function BinComp(ByRef strBin, ByRef arrHex)
        Dim intComp, i
        intComp = &HFF
        For i = 0 To UBound(arrHex)
            intComp = intComp And (AscB(MidB(strBin, i + 1, 1)) Eqv arrHex(i))
        Next
        BinComp = (intComp = &HFF)
    End Function
    
End Class

Class JSON_Advanced
    'Author: Mr.Jos(sd7056333)
    'Date: 2016/07/07
    'Website: http://blog.csdn.net/sd7056333
    'Description: 
    '   This is a class for transforming a representation of specified data between 
    '   JSON string and VBScript structure. Here are the corresponding data types: 
    '       +-----------------+-----------------+
    '       | VBScript        | JSON            |
    '       +=================+=================+
    '       | Dictionary      | object          |
    '       +-----------------+-----------------+
    '       | Array           | array           |
    '       +-----------------+-----------------+
    '       | String          | string          |
    '       +-----------------+-----------------+
    '       | Number          | number          |
    '       +-----------------+-----------------+
    '       | True/False/Null | true/false/null |
    '       +-----------------+-----------------+
    'Main references: 
    '   http://www.json.org/
    '   http://demon.tw/my-work/vbs-json.html
    '   http://www.ediy.co.nz/vbjson-json-parser-library-in-vb6-xidc55680.html
    
    Public Strict_Standard  'False(Default)
        'Setting: Boolean, apply strict-JSON-standards or not. Details:
        '    For EncodeJSON, "/" --> "\/" ; Non-ASCII-Charactors(ASC>127) --> "\uXXXX"
    
    Public Function EncodeJSON(ByRef varStruct, ByRef intFormat)
        'Function: VBScript structure -> JSON string
        '  [varStruct] Dictionary or Array, data of a VBScript structure.
        '  [intFormat] Integer, number of structure layers to be formatted.
        Select Case VarType(varStruct)
            Case vbObject
                EncodeJSON = EncodeObject(varStruct, intFormat, 1)
            Case vbArray + vbVariant
                EncodeJSON = EncodeArray(varStruct, intFormat, 1)
            Case Else
                Err.Raise 8732,,"Invalid VBScript structure type"
        End Select
    End Function
    
    Public Sub DecodeJSON(ByRef strText, ByRef varReturn)
        'Function: JSON string -> VBScript structure
        '  [strText] String, JSON text to be parsed.
        '  [varReturn] Return value, parsed VBScript structure.
        'Identifiable violations: 
        '  Comments in the space ("// ...", "/* ... */"). --> SKIP
        '  "true", "false" and "null" with capital letters.
        '  Horizontal tabs (vbTab) in string.
        '  Strings in single quotes.
        Dim arrPos(2)
        arrPos(0) = 1   'Current position of the reading text
        Call SkipSpace(strText, arrPos)
        Select Case Mid(strText, arrPos(0), 1)
            Case "{"
                Set varReturn = DecodeObject(strText, arrPos)
            Case "["
                varReturn = DecodeArray(strText, arrPos)
            Case Else
                Call ErrRaise(arrPos, "No JSON structure has been found")
        End Select
    End Sub
    
    Private Function EncodeVariant(ByRef varData, ByRef intFormat, ByRef intLayer)
        Select Case VarType(varData)
            Case vbNull
                EncodeVariant = "null"
            Case vbBoolean
                Select Case varData
                    Case True  EncodeVariant = "true"
                    Case False EncodeVariant = "false"
                End Select
            Case vbInteger, vbLong, vbSingle, vbDouble
                EncodeVariant = CStr(varData)
            Case vbString
                EncodeVariant = EncodeString(varData)
            Case vbArray + vbVariant
                EncodeVariant = EncodeArray(varData, intFormat, intLayer)
            Case vbObject
                EncodeVariant = EncodeObject(varData, intFormat, intLayer)
            Case Else
                EncodeVariant = """" & CStr(varData) & """"
        End Select
    End Function
    
    Private Function EncodeObject(ByRef dicData, ByRef intFormat, ByRef intLayer)
        Dim arrText, i, varEle, strSep
        If TypeName(dicData) <> "Dictionary" Then Err.Raise 8732,,"Non-dictionary object"
        Redim arrText(3 * dicData.Count)
        If intLayer <= intFormat Then
            strSep = "," & vbCrlf & String(intLayer, vbTab)
            arrText(0) = "{" & vbCrlf & String(intLayer, vbTab)
        Else
            strSep = ", "
            arrText(0) = "{"
        End If
        For Each varEle In dicData
            arrText(3 * i + 1) = """" & varEle & """: "
            arrText(3 * i + 2) = EncodeVariant(dicData(varEle), intFormat, intLayer + 1)
            arrText(3 * i + 3) = strSep
            i = i + 1
        Next
        If dicData.Count = 0 Then
            arrText(0) = "{}"
        ElseIf intLayer <= intFormat Then
            arrText(UBound(arrText)) = vbCrlf & String(intLayer - 1, vbTab) & "}"
        Else
            arrText(UBound(arrText)) = "}"
        End If
        EncodeObject = Join(arrText, "")
    End Function
    
    Private Function EncodeArray(ByRef arrData, ByRef intFormat, ByRef intLayer)
        Dim arrText, i, varEle, strSep
        Redim arrText(0)
        If intLayer <= intFormat Then
            strSep = "," & vbCrlf & String(intLayer, vbTab)
            arrText(0) = "[" & vbCrlf & String(intLayer, vbTab)
        Else
            strSep = ", "
            arrText(0) = "["
        End If
        For Each varEle In arrData
            Redim Preserve arrText(2 * i + 2)
            arrText(2 * i + 1) = EncodeVariant(varEle, intFormat, intLayer + 1)
            arrText(2 * i + 2) = strSep
            i = i + 1
        Next
        If i = 0 Then
            arrText(0) = "[]"
        ElseIf intLayer <= intFormat Then
            arrText(UBound(arrText)) = vbCrlf & String(intLayer - 1, vbTab) & "]"
        Else
            arrText(UBound(arrText)) = "]"
        End If
        EncodeArray = Join(arrText, "")
    End Function
    
    Private Function EncodeString(ByRef strValue)
        Dim arrText, i, intAsc
        Redim arrText(Len(strValue) + 1)
        arrText(0) = """"
        For i = 1 To UBound(arrText) - 1
            arrText(i) = Mid(strValue, i, 1)
            intAsc = AscW(arrText(i))
            Select Case intAsc
                Case &H08 arrText(i) = "\b"  'backspace
                Case &H09 arrText(i) = "\t"  'horizontal tab
                Case &H0A arrText(i) = "\n"  'newline
                Case &H0C arrText(i) = "\f"  'formfeed
                Case &H0D arrText(i) = "\r"  'carriage return
                Case &H22 arrText(i) = "\""" ' "
                Case &H5C arrText(i) = "\\"  ' \
                Case &H2F
                    If Strict_Standard Then arrText(i) = "\/"
                Case Else
                    If intAsc >= 0 And intAsc <= 31 Then
                        arrText(i) = "\u" & Right("0000" & Hex(intAsc), 4)
                    ElseIf Strict_Standard Then
                        If intAsc < 0 Or intAsc > 127 Then
                            arrText(i) = "\u" & Right("0000" & Hex(intAsc), 4)
                        End If
                    End If
            End Select
        Next
        arrText(UBound(arrText)) = """"
        EncodeString = Join(arrText, "")
    End Function
    
    Private Sub DecodeVariant(ByRef strText, ByRef arrPos, ByRef varData)
        varData = Empty
        Call SkipSpace(strText, arrPos)
        Select Case Mid(strText, arrPos(0), 1)
            Case "{"
                Set varData = DecodeObject(strText, arrPos)
            Case "["
                varData = DecodeArray(strText, arrPos)
            Case """", "'"
                varData = DecodeString(strText, arrPos)
            Case "t", "T"
                If LCase(Mid(strText, arrPos(0), 4)) = "true" Then
                    varData = True
                    arrPos(0) = arrPos(0) + 4
                End If
            Case "f", "F"
                If LCase(Mid(strText, arrPos(0), 5)) = "false" Then
                    varData = False
                    arrPos(0) = arrPos(0) + 5
                End If
            Case "n", "N"
                If LCase(Mid(strText, arrPos(0), 4)) = "null" Then
                    varData = Null
                    arrPos(0) = arrPos(0) + 4
                End If
            Case Else
                Dim intPos, strNum
                intPos = arrPos(0)
                Do While intPos <= Len(strText)
                    If InStr("+-0123456789.eE", Mid(strText, intPos, 1)) = 0 Then Exit Do
                    intPos = intPos + 1
                Loop
                strNum = Mid(strText, arrPos(0), intPos - arrPos(0))
                If IsNumeric(strNum) Then
                    varData = CDbl(strNum)
                    arrPos(0) = intPos
                End If
        End Select
        If IsEmpty(varData) Then Call ErrRaise(arrPos, "No JSON value has been parsed")
    End Sub
    
    Private Function DecodeObject(ByRef strText, ByRef arrPos)
        Dim strKey, varValue
        Set DecodeObject = CreateObject("Scripting.Dictionary")
        arrPos(0) = arrPos(0) + 1
        Call SkipSpace(strText, arrPos)
        If Mid(strText, arrPos(0), 1) = "}" Then
            arrPos(0) = arrPos(0) + 1
            Exit Function
        End If
        Do
            If arrPos(0) > Len(strText) Then
                Call ErrRaise(arrPos(0), "Missing '}'")
            ElseIf Mid(strText, arrPos(0), 1) <> """" Then
                Call ErrRaise(arrPos, "Expecting property name")
            End If
            strKey = DecodeString(strText, arrPos)
            Call SkipSpace(strText, arrPos)
            If Len(strKey) = 0 Then
                Call ErrRaise(arrPos, "Property name cannot be empty")
            ElseIf Mid(strText, arrPos(0), 1) <> ":" Then
                Call ErrRaise(arrPos, "Expecting ':' delimiter")
            End If
            arrPos(0) = arrPos(0) + 1
            Call DecodeVariant(strText, arrPos, varValue)
            If DecodeObject.Exists(strKey) Then DecodeObject.Remove(strKey)
            DecodeObject.Add strKey, varValue
            Call SkipSpace(strText, arrPos)
            Select Case Mid(strText, arrPos(0), 1)
                Case "}" Exit Do
                Case ","
                Case Else Call ErrRaise(arrPos, "Expecting ',' delimiter")
            End Select
            arrPos(0) = arrPos(0) + 1
            Call SkipSpace(strText, arrPos)
        Loop
        arrPos(0) = arrPos(0) + 1
    End Function
    
    Private Function DecodeArray(ByRef strText, ByRef arrPos)
        Dim dicArray, varValue
        Set dicArray = CreateObject("Scripting.Dictionary")
        arrPos(0) = arrPos(0) + 1
        Call SkipSpace(strText, arrPos)
        If Mid(strText, arrPos(0), 1) <> "]" Then
            Do
                If arrPos(0) > Len(strText) Then Call ErrRaise(arrPos, "Missing ']'")
                Call DecodeVariant(strText, arrPos, varValue)
                dicArray.Add dicArray.Count, varValue
                Call SkipSpace(strText, arrPos)
                Select Case Mid(strText, arrPos(0), 1)
                    Case "]" Exit Do
                    Case ","
                    Case Else Call ErrRaise(arrPos, "Expecting ',' delimiter")
                End Select
                arrPos(0) = arrPos(0) + 1
            Loop
        End If
        arrPos(0) = arrPos(0) + 1
        DecodeArray = dicArray.Items
    End Function
    
    Private Function DecodeString(ByRef strText, ByRef arrPos)
        Dim dicString, strQuote, vbBack, strChar
        Set dicString = CreateObject("Scripting.Dictionary")
        strQuote = Mid(strText, arrPos(0), 1)
        vbBack = ChrW(8)
        arrPos(0) = arrPos(0) + 1
        Do
            If arrPos(0) > Len(strText) Then Call ErrRaise(arrPos, "Missing quote")
            strChar = Mid(strText, arrPos(0), 1)
            Select Case strChar
            Case strQuote
                If Mid(strText, arrPos(0) + 1, 1) = strQuote Then
                    dicString.Add dicString.Count, strChar
                    arrPos(0) = arrPos(0) + 2
                Else
                    arrPos(0) = arrPos(0) + 1
                    Exit Do
                End If
            Case "\"
                arrPos(0) = arrPos(0) + 1
                strChar = Mid(strText, arrPos(0), 1)
                Select Case strChar
                    Case """", "\", "/", "'"
                    Case "b" strChar = vbBack
                    Case "t" strChar = vbTab
                    Case "n" strChar = vbLf
                    Case "f" strChar = vbFormFeed
                    Case "r" strChar = vbCr
                    Case "u"
                        strChar = "&H" & Mid(strText, arrPos(0) + 1, 4)
                        If IsNumeric(strChar) Then
                            strChar = ChrW(strChar)
                            arrPos(0) = arrPos(0) + 4
                        Else
                            Call ErrRaise(arrPos, "Incorrect hex")
                        End If
                    Case Else
                        Call ErrRaise(arrPos, "Invalid control character")
                End Select
                dicString.Add dicString.Count, strChar
                arrPos(0) = arrPos(0) + 1
            'Case vbLf, vbCr, vbBack, vbFormFeed, vbTab
            '   Call ErrRaise(arrPos, "Invalid character in a string")
            Case Else
                dicString.Add dicString.Count, strChar
                arrPos(0) = arrPos(0) + 1
            End Select
        Loop
        DecodeString = Join(dicString.Items, "")
    End Function
    
    Private Sub SkipSpace(ByRef strText, ByRef arrPos)
        Dim blnComment_Status, blnComment_Long
        Do While arrPos(0) <= Len(strText)
            Select Case Mid(strText, arrPos(0), 1)
                Case vbTab, " ", "(", ")"
                    
                Case vbCr, vbLf
                    If blnComment_Status Then
                        If Not blnComment_Long Then blnComment_Status = False
                    End If
                    If Mid(strText, arrPos(0), 2) = vbCrlf Then arrPos(0) = arrPos(0) + 1
                    arrPos(1) = arrPos(0)       'Position of the end of last row
                    arrPos(2) = arrPos(2) + 1   'Number of rows having been read
                Case "/"
                    If Not blnComment_Status Then
                        If arrPos(0) = Len(strText) Then Exit Do
                        arrPos(0) = arrPos(0) + 1
                        Select Case Mid(strText, arrPos(0), 1)
                            Case "/"
                                blnComment_Status = True
                                blnComment_Long = False
                            Case "*"
                                blnComment_Status = True
                                blnComment_Long = True
                            Case Else
                                Call ErrRaise(arrPos, "Invalid comment characters")
                        End Select
                    End If
                Case "*"
                    If blnComment_Long Then
                        If arrPos(0) = Len(strText) Then Exit Do
                        arrPos(0) = arrPos(0) + 1
                        If Mid(strText, arrPos(0), 1) = "/" Then
                            blnComment_Status = False
                            blnComment_Long = False
                        End If
                    End If
                Case Else
                    If Not blnComment_Status Then Exit Do
            End Select
            arrPos(0) = arrPos(0) + 1
        Loop
    End Sub
    
    Private Sub ErrRaise(ByRef arrPos, ByRef strDescription)
        Err.Raise 8732, "JSON Format Error", Join(Array(strDescription, _
            " at (", arrPos(2) + 1, ", ", arrPos(0) - arrPos(1), ")"), "")
    End Sub
    
End Class