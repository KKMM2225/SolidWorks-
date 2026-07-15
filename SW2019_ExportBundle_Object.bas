Attribute VB_Name = "ExportBundleObject"
Option Explicit

' SolidWorks 2019 compatible macro.
' All SolidWorks API objects are late-bound as Object to avoid VBA reference/version issues.

Private Const PDF_PRINTER_NAME As String = "Microsoft Print to PDF"
Private Const MACRO_VERSION As String = "v2026-07-14-SW2019-FOLDER-OR-PATH"
Private Const SHOW_VERSION_ON_START As Boolean = True

Private Const swDocPART As Long = 1
Private Const swDocASSEMBLY As Long = 2
Private Const swDocDRAWING As Long = 3
Private Const swOpenDocOptions_Silent As Long = 1
Private Const swSaveAsCurrentVersion As Long = 0
Private Const swSaveAsOptions_Silent As Long = 1
Private Const swSaveAsOptions_Copy As Long = 2

Private swApp As Object

Sub main()
    On Error GoTo Fail

    LogLine "START " & MACRO_VERSION

    Set swApp = Application.SldWorks

    If SHOW_VERSION_ON_START Then
        MsgBox "Export bundle " & MACRO_VERSION & vbCrLf & _
               "SW2019 Object late-bound build.", vbInformation, "Export bundle " & MACRO_VERSION
    End If

    Dim swModel As Object
    Set swModel = swApp.ActiveDoc

    If swModel Is Nothing Then
        MsgBox "Please open the target SolidWorks drawing first.", vbExclamation, "Export bundle " & MACRO_VERSION
        Exit Sub
    End If

    If swModel.GetType <> swDocDRAWING Then
        MsgBox "The active document is not a drawing. Please activate a .slddrw file first.", vbExclamation, "Export bundle " & MACRO_VERSION
        Exit Sub
    End If

    Dim rootFolder As String
    rootFolder = PickFolderOrTypePath("Paste/type the parent folder path, or leave blank and click OK to choose a folder.")
    If Len(rootFolder) = 0 Then Exit Sub
    LogLine "Root folder: " & rootFolder

    Dim defaultName As String
    defaultName = GetDefaultExportName(swModel)

    Dim exportName As String
    exportName = InputBox("Enter part/export name without extension:", "Export bundle " & MACRO_VERSION, defaultName)
    exportName = Trim$(exportName)
    If Len(exportName) = 0 Then Exit Sub

    exportName = CleanFileName(exportName)
    LogLine "Export name: " & exportName
    If Len(exportName) = 0 Then
        MsgBox "Invalid export name.", vbExclamation, "Export bundle " & MACRO_VERSION
        Exit Sub
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim drwFolder As String
    Dim dwgFolder As String
    Dim pdfFolder As String
    Dim stepFolder As String

    drwFolder = fso.BuildPath(rootFolder, DrawingFolderName())
    dwgFolder = fso.BuildPath(rootFolder, "DWG")
    pdfFolder = fso.BuildPath(rootFolder, "PDF")
    stepFolder = fso.BuildPath(rootFolder, "STEP")

    EnsureFolderTree drwFolder
    EnsureFolderTree dwgFolder
    EnsureFolderTree pdfFolder
    EnsureFolderTree stepFolder
    LogLine "Subfolders ready."

    Dim drwPath As String
    Dim dwgPath As String
    Dim pdfPath As String
    Dim stepPath As String

    drwPath = fso.BuildPath(drwFolder, exportName & ".slddrw")
    dwgPath = fso.BuildPath(dwgFolder, exportName & ".dwg")
    pdfPath = fso.BuildPath(pdfFolder, exportName & ".pdf")
    stepPath = fso.BuildPath(stepFolder, exportName & ".step")

    Dim referencedModelPath As String
    referencedModelPath = GetFirstReferencedModelPath(swModel)
    LogLine "Referenced model: " & referencedModelPath

    LogLine "Export drawing start: " & drwPath
    SaveActiveDrawingCopy swModel, drwPath
    LogLine "Export drawing done."

    LogLine "Export DWG start: " & dwgPath
    SaveActiveDrawingAsDwg swModel, dwgPath
    LogLine "Export DWG done."

    LogLine "Export PDF start: " & pdfPath
    SaveActiveDrawingAsPdf swModel, pdfPath
    LogLine "Export PDF done."

    Dim stepMessage As String
    If Len(referencedModelPath) > 0 Then
        LogLine "Export STEP start: " & stepPath
        ExportReferencedModelToStep referencedModelPath, stepPath
        LogLine "Export STEP done."
        stepMessage = "STEP: " & stepPath
    Else
        LogLine "STEP skipped: no referenced model."
        stepMessage = "STEP: skipped, no referenced model was found."
    End If

    MsgBox "Export complete:" & vbCrLf & _
           "Drawing: " & drwPath & vbCrLf & _
           "DWG: " & dwgPath & vbCrLf & _
           "PDF: " & pdfPath & vbCrLf & _
           stepMessage, vbInformation, "Export bundle " & MACRO_VERSION
    Exit Sub

Fail:
    LogLine "FAILED: " & Err.Description
    MsgBox "Export failed:" & vbCrLf & Err.Description, vbCritical, "Export bundle " & MACRO_VERSION
End Sub

Private Sub LogLine(ByVal message As String)
    On Error Resume Next

    Dim fso As Object
    Dim stream As Object
    Dim logPath As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    logPath = fso.BuildPath(CreateObject("WScript.Shell").ExpandEnvironmentStrings("%TEMP%"), "SW2019_ExportBundle_Log.txt")

    Set stream = fso.OpenTextFile(logPath, 8, True)
    stream.WriteLine Format$(Now, "yyyy-mm-dd hh:nn:ss") & "  " & message
    stream.Close
End Sub

Private Function DrawingFolderName() As String
    DrawingFolderName = ChrW(&H5DE5) & ChrW(&H7A0B) & ChrW(&H56FE)
End Function

Private Function PickFolderOrTypePath(ByVal prompt As String) As String
    Dim typedPath As String
    typedPath = InputBox(prompt & vbCrLf & vbCrLf & _
                         "Example: C:\Users\22915\Desktop\ExportTest" & vbCrLf & _
                         "Tip: leave this box blank to open folder selection." & vbCrLf & _
                         "Do not include the final subfolder name such as DWG/PDF/STEP.", _
                         "Export bundle " & MACRO_VERSION, "")
    typedPath = NormalizeFolderPath(typedPath)

    If Len(typedPath) = 0 Then
        LogLine "Folder path box was blank. Opening folder picker."
        typedPath = PickFolder("Choose the parent folder. Subfolders will be used or created automatically.")
        typedPath = NormalizeFolderPath(typedPath)

        If Len(typedPath) = 0 Then
            LogLine "Folder picker cancelled."
            PickFolderOrTypePath = ""
            Exit Function
        End If
    Else
        LogLine "Folder path typed or pasted."
    End If

    EnsureFolderTree typedPath
    PickFolderOrTypePath = typedPath
End Function

Private Function PickFolder(ByVal prompt As String) As String
    On Error GoTo Done

    Dim shellApp As Object
    Dim folderObj As Object

    Set shellApp = CreateObject("Shell.Application")
    Set folderObj = shellApp.BrowseForFolder(0, prompt, 0, 0)

    If Not folderObj Is Nothing Then
        PickFolder = folderObj.Self.Path
    End If

Done:
End Function

Private Function NormalizeFolderPath(ByVal folderPath As String) As String
    Dim result As String
    result = Trim$(folderPath)

    If Len(result) >= 2 Then
        If Left$(result, 1) = """" And Right$(result, 1) = """" Then
            result = Mid$(result, 2, Len(result) - 2)
        End If
    End If

    Do While Len(result) > 3 And Right$(result, 1) = "\"
        result = Left$(result, Len(result) - 1)
    Loop

    NormalizeFolderPath = result
End Function

Private Sub EnsureFolderTree(ByVal folderPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(folderPath) Then Exit Sub

    Dim parentPath As String
    parentPath = fso.GetParentFolderName(folderPath)

    If Len(parentPath) > 0 Then
        If Not fso.FolderExists(parentPath) Then
            EnsureFolderTree parentPath
        End If
    End If

    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
End Sub

Private Function CleanFileName(ByVal rawName As String) As String
    Dim invalidChars As Variant
    invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")

    Dim result As String
    result = Trim$(rawName)

    Dim i As Long
    For i = LBound(invalidChars) To UBound(invalidChars)
        result = Replace(result, CStr(invalidChars(i)), "_")
    Next i

    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop

    CleanFileName = result
End Function

Private Function GetDefaultExportName(ByVal swModel As Object) As String
    Dim modelPath As String
    modelPath = swModel.GetPathName

    If Len(modelPath) > 0 Then
        GetDefaultExportName = FileNameWithoutExtension(modelPath)
    Else
        GetDefaultExportName = FileNameWithoutExtension(swModel.GetTitle)
    End If
End Function

Private Function FileNameWithoutExtension(ByVal filePathOrTitle As String) As String
    Dim nameOnly As String
    nameOnly = filePathOrTitle

    Dim slashPos As Long
    slashPos = InStrRev(nameOnly, "\")
    If slashPos > 0 Then nameOnly = Mid$(nameOnly, slashPos + 1)

    Dim dotPos As Long
    dotPos = InStrRev(nameOnly, ".")
    If dotPos > 0 Then nameOnly = Left$(nameOnly, dotPos - 1)

    FileNameWithoutExtension = CleanFileName(nameOnly)
End Function

Private Sub SaveActiveDrawingCopy(ByVal swModel As Object, ByVal outPath As String)
    LogLine "Drawing method: Extension SaveAs Copy"
    If TrySaveModelAs(swModel, outPath, swSaveAsOptions_Silent + swSaveAsOptions_Copy) Then Exit Sub

    LogLine "Drawing method: file copy fallback"
    If CopyCurrentFileAsFallback(swModel, outPath) Then Exit Sub

    LogLine "Drawing method: Extension SaveAs normal"
    If TrySaveModelAs(swModel, outPath, swSaveAsOptions_Silent) Then Exit Sub

    Err.Raise vbObjectError + 1001, , "Drawing export failed." & vbCrLf & "Path: " & outPath
End Sub

Private Sub SaveActiveDrawingAsDwg(ByVal swModel As Object, ByVal outPath As String)
    LogLine "DWG method: SaveAs fallbacks"
    If TryExportBySaveAs(swModel, outPath) Then Exit Sub

    Err.Raise vbObjectError + 1006, , "DWG export failed after multiple SaveAs methods." & vbCrLf & _
              "Path: " & outPath & vbCrLf & _
              "Try manually saving this drawing as DWG once in SolidWorks, confirm the DXF/DWG export settings, then run the macro again."
End Sub

Private Sub SaveActiveDrawingAsPdf(ByVal swModel As Object, ByVal pdfPath As String)
    DeleteFileIfSafe pdfPath, ""

    LogLine "PDF method: SolidWorks SaveAs PDF"
    If TrySaveModelAs(swModel, pdfPath, swSaveAsOptions_Silent) Then
        If OutputFileExists(pdfPath) Then Exit Sub
    End If

    Err.Raise vbObjectError + 1003, , "PDF SaveAs failed." & vbCrLf & "Target: " & pdfPath
End Sub

Private Function TryExportBySaveAs(ByVal swModel As Object, ByVal outPath As String) As Boolean
    DeleteFileIfSafe outPath, swModel.GetPathName
    ActivateModel swModel

    LogLine "TryExportBySaveAs: ModelDoc SaveAs chain: " & outPath
    If TryModelDocSaveAs(swModel, outPath) Then
        TryExportBySaveAs = OutputFileExists(outPath)
        If TryExportBySaveAs Then Exit Function
    End If

    LogLine "TryExportBySaveAs: Extension SaveAs silent: " & outPath
    If TrySaveModelAs(swModel, outPath, swSaveAsOptions_Silent) Then
        TryExportBySaveAs = OutputFileExists(outPath)
        If TryExportBySaveAs Then Exit Function
    End If

    LogLine "TryExportBySaveAs: Extension SaveAs copy: " & outPath
    If TrySaveModelAs(swModel, outPath, swSaveAsOptions_Silent + swSaveAsOptions_Copy) Then
        TryExportBySaveAs = OutputFileExists(outPath)
    End If
End Function

Private Function TrySaveModelAs(ByVal swModel As Object, ByVal outPath As String, ByVal saveOptions As Long) As Boolean
    Dim errors As Long
    Dim warnings As Long

    DeleteFileIfSafe outPath, swModel.GetPathName

    If TryExtensionSaveAs3LateBound(swModel, outPath, saveOptions, errors, warnings) Then
        LogLine "Extension.SaveAs3 succeeded. Errors=" & CStr(errors) & ", Warnings=" & CStr(warnings)
        TrySaveModelAs = True
        Exit Function
    End If
    LogLine "Extension.SaveAs3 failed. Errors=" & CStr(errors) & ", Warnings=" & CStr(warnings)

    errors = 0
    warnings = 0
    If TryExtensionSaveAs(swModel, outPath, saveOptions, errors, warnings) Then
        LogLine "Extension.SaveAs succeeded. Errors=" & CStr(errors) & ", Warnings=" & CStr(warnings)
        TrySaveModelAs = True
    Else
        LogLine "Extension.SaveAs failed. Errors=" & CStr(errors) & ", Warnings=" & CStr(warnings)
    End If
End Function

Private Function TryExtensionSaveAs3LateBound(ByVal swModel As Object, ByVal outPath As String, ByVal saveOptions As Long, ByRef errors As Long, ByRef warnings As Long) As Boolean
    On Error Resume Next
    Err.Clear

    Dim result As Variant
    result = CallByName(swModel.Extension, "SaveAs3", VbMethod, outPath, swSaveAsCurrentVersion, saveOptions, Nothing, Nothing, errors, warnings)

    If Err.Number = 0 Then
        TryExtensionSaveAs3LateBound = CBool(result)
    End If

    Err.Clear
    On Error GoTo 0
End Function

Private Function TryExtensionSaveAs(ByVal swModel As Object, ByVal outPath As String, ByVal saveOptions As Long, ByRef errors As Long, ByRef warnings As Long) As Boolean
    On Error Resume Next
    Err.Clear
    TryExtensionSaveAs = swModel.Extension.SaveAs(outPath, swSaveAsCurrentVersion, saveOptions, Nothing, errors, warnings)
    If Err.Number <> 0 Then
        Err.Clear
        TryExtensionSaveAs = False
    End If
    On Error GoTo 0
End Function

Private Function TryModelDocSaveAs(ByVal swModel As Object, ByVal outPath As String) As Boolean
    Dim result As Variant
    Dim errors As Long
    Dim warnings As Long

    On Error Resume Next

    Err.Clear
    result = CallByName(swModel, "SaveAs3", VbMethod, outPath, swSaveAsCurrentVersion, swSaveAsOptions_Silent)
    If Err.Number = 0 Then
        If CBool(result) Then
            LogLine "ModelDoc.SaveAs3 succeeded."
            TryModelDocSaveAs = True
            On Error GoTo 0
            Exit Function
        End If
    End If
    LogLine "ModelDoc.SaveAs3 failed or returned False."

    Err.Clear
    errors = 0
    warnings = 0
    result = CallByName(swModel, "SaveAs2", VbMethod, outPath, swSaveAsCurrentVersion, swSaveAsOptions_Silent, False, errors, warnings)
    If Err.Number = 0 Then
        If CBool(result) Then
            LogLine "ModelDoc.SaveAs2 succeeded. Errors=" & CStr(errors) & ", Warnings=" & CStr(warnings)
            TryModelDocSaveAs = True
            On Error GoTo 0
            Exit Function
        End If
    End If
    LogLine "ModelDoc.SaveAs2 failed or returned False. Errors=" & CStr(errors) & ", Warnings=" & CStr(warnings)

    Err.Clear
    result = CallByName(swModel, "SaveAs", VbMethod, outPath)
    If Err.Number = 0 Then
        If CBool(result) Then
            LogLine "ModelDoc.SaveAs succeeded."
            TryModelDocSaveAs = True
        Else
            LogLine "ModelDoc.SaveAs returned False."
        End If
    Else
        LogLine "ModelDoc.SaveAs raised error."
    End If

    On Error GoTo 0
End Function

Private Function CopyCurrentFileAsFallback(ByVal swModel As Object, ByVal outPath As String) As Boolean
    Dim sourcePath As String
    sourcePath = swModel.GetPathName

    If Len(sourcePath) = 0 Then Exit Function
    If LCase$(sourcePath) = LCase$(outPath) Then
        CopyCurrentFileAsFallback = True
        Exit Function
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(sourcePath) Then Exit Function

    DeleteFileIfSafe outPath, sourcePath
    fso.CopyFile sourcePath, outPath, True

    CopyCurrentFileAsFallback = fso.FileExists(outPath)
End Function

Private Sub ActivateModel(ByVal swModel As Object)
    On Error Resume Next
    Dim activateErrors As Long
    swApp.ActivateDoc2 swModel.GetTitle, False, activateErrors
    On Error GoTo 0
End Sub

Private Function OutputFileExists(ByVal filePath As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(filePath) Then
        OutputFileExists = (fso.GetFile(filePath).Size > 0)
    End If
End Function

Private Function WaitForFile(ByVal filePath As String, ByVal timeoutSeconds As Long) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim deadline As Date
    deadline = DateAdd("s", timeoutSeconds, Now)

    Do While Now < deadline
        DoEvents
        If fso.FileExists(filePath) Then
            If fso.GetFile(filePath).Size > 0 Then
                WaitForFile = True
                Exit Function
            End If
        End If
    Loop
End Function

Private Function GetFirstReferencedModelPath(ByVal swDrawing As Object) As String
    Dim swView As Object
    Set swView = swDrawing.GetFirstView

    If Not swView Is Nothing Then
        Set swView = swView.GetNextView
    End If

    Do While Not swView Is Nothing
        Dim candidate As String
        candidate = ""

        On Error Resume Next
        candidate = swView.GetReferencedModelName
        On Error GoTo 0

        If Len(candidate) = 0 Then
            Dim refDoc As Object
            On Error Resume Next
            Set refDoc = swView.ReferencedDocument
            On Error GoTo 0

            If Not refDoc Is Nothing Then candidate = refDoc.GetPathName
        End If

        If Len(candidate) > 0 Then
            GetFirstReferencedModelPath = candidate
            Exit Function
        End If

        Set swView = swView.GetNextView
    Loop
End Function

Private Sub ExportReferencedModelToStep(ByVal modelPath As String, ByVal stepPath As String)
    Dim docType As Long
    docType = DocTypeFromPath(modelPath)

    If docType = 0 Then
        Err.Raise vbObjectError + 1004, , "Cannot detect referenced model type for STEP export: " & modelPath
    End If

    Dim refModel As Object
    Dim openedHere As Boolean

    Set refModel = swApp.GetOpenDocumentByName(modelPath)

    If refModel Is Nothing Then
        Dim openErrors As Long
        Dim openWarnings As Long
        Set refModel = swApp.OpenDoc6(modelPath, docType, swOpenDocOptions_Silent, "", openErrors, openWarnings)
        openedHere = True

        If refModel Is Nothing Then
            Err.Raise vbObjectError + 1005, , "Cannot open referenced model for STEP export:" & vbCrLf & _
                      modelPath & vbCrLf & _
                      "Errors=" & CStr(openErrors) & ", Warnings=" & CStr(openWarnings)
        End If
    End If

    If Not TryExportBySaveAs(refModel, stepPath) Then
        Err.Raise vbObjectError + 1007, , "STEP export failed after multiple SaveAs methods." & vbCrLf & _
                  "Path: " & stepPath & vbCrLf & _
                  "Referenced model: " & refModel.GetPathName & vbCrLf & _
                  "Make sure the referenced part or assembly is saved and can be manually saved as STEP."
    End If

    If openedHere Then
        swApp.CloseDoc refModel.GetTitle
    End If
End Sub

Private Function DocTypeFromPath(ByVal filePath As String) As Long
    Dim ext As String
    ext = LCase$(Mid$(filePath, InStrRev(filePath, ".") + 1))

    Select Case ext
        Case "sldprt"
            DocTypeFromPath = swDocPART
        Case "sldasm"
            DocTypeFromPath = swDocASSEMBLY
        Case Else
            DocTypeFromPath = 0
    End Select
End Function

Private Sub DeleteFileIfSafe(ByVal filePath As String, ByVal currentModelPath As String)
    On Error GoTo Done

    If Len(filePath) = 0 Then Exit Sub
    If Len(currentModelPath) > 0 Then
        If LCase$(filePath) = LCase$(currentModelPath) Then Exit Sub
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(filePath) Then
        fso.DeleteFile filePath, True
    End If

Done:
End Sub
