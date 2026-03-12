Attribute VB_Name = "RKM_Excel"
Option Explicit

Private Const DEFAULT_SHEET_NAME As String = "ALBUM"
Private Const HEADER_ROW_INDEX As Long = 1
Private Const HEADER_SCAN_ROWS As Long = 10

Public Function LoadAlbumItemsFromExcel(ByVal excelPath As String, Optional ByVal workspacePath As String = "", Optional ByVal sheetName As String = DEFAULT_SHEET_NAME) As Collection
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    Dim headerMap As Object
    Dim rowsCount As Long
    Dim i As Long
    Dim modelPathRaw As String
    Dim resolvedModelPath As String
    Dim item As Object
    Dim promptMap As Object

    On Error GoTo EH

    If Len(Trim$(excelPath)) = 0 Then Exit Function

    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = False

    Set xlBook = xlApp.Workbooks.Open(excelPath)
    Set xlSheet = ResolveAlbumWorksheet(xlBook, sheetName)
    If xlSheet Is Nothing Then
        Err.Raise vbObjectError + 3100, "LoadAlbumItemsFromExcel", "Worksheet '" & sheetName & "' was not found in " & excelPath
    End If

    Set headerMap = ReadHeaderMap(xlSheet, DetectHeaderRowIndex(xlSheet))
    If Not headerMap.Exists("MODEL_PATH") Then
        Err.Raise vbObjectError + 3101, "LoadAlbumItemsFromExcel", "Required header MODEL_PATH is missing on worksheet '" & sheetName & "'."
    End If

    rowsCount = LastUsedRowByColumn(xlSheet, CLng(headerMap("MODEL_PATH")))
    Set LoadAlbumItemsFromExcel = New Collection

    For i = CLng(headerMap("__HEADER_ROW")) + 1 To rowsCount
        modelPathRaw = Trim$(ReadCellText(xlSheet, i, CLng(headerMap("MODEL_PATH"))))
        If Len(modelPathRaw) = 0 Then GoTo ContinueLoop

        resolvedModelPath = ResolveModelPath(modelPathRaw, workspacePath, excelPath)
        If Len(resolvedModelPath) = 0 Then GoTo ContinueLoop

        Set item = CreateObject("Scripting.Dictionary")
        item.CompareMode = vbTextCompare
        item("MODEL_PATH") = resolvedModelPath

        Set promptMap = CreatePromptMapFromRow(xlSheet, headerMap, i)
        item("PROMPTS") = promptMap

        If headerMap.Exists("SHEET") Then item("SHEET") = Trim$(ReadCellText(xlSheet, i, CLng(headerMap("SHEET"))))
        If headerMap.Exists("SHEETS") Then item("SHEETS") = Trim$(ReadCellText(xlSheet, i, CLng(headerMap("SHEETS"))))

        LoadAlbumItemsFromExcel.Add item

ContinueLoop:
    Next i

CleanExit:
    CloseExcelObjects xlBook, xlApp
    Set xlSheet = Nothing
    Set headerMap = Nothing
    Exit Function
EH:
    Debug.Print "LOG: Excel load failed. Err=" & Err.Number & "; " & Err.Description & "; File=" & excelPath
    Set LoadAlbumItemsFromExcel = Nothing
    Resume CleanExit
End Function

Private Function ResolveAlbumWorksheet(ByVal xlBook As Object, ByVal sheetName As String) As Object
    Dim oSheet As Object

    On Error Resume Next
    Set oSheet = xlBook.Worksheets(sheetName)
    On Error GoTo 0

    If Not oSheet Is Nothing Then
        Set ResolveAlbumWorksheet = oSheet
        Exit Function
    End If

    If xlBook.Worksheets.Count > 0 Then
        Set ResolveAlbumWorksheet = xlBook.Worksheets(1)
    End If
End Function

Private Function DetectHeaderRowIndex(ByVal xlSheet As Object) As Long
    Dim rowIndex As Long
    Dim map As Object

    DetectHeaderRowIndex = HEADER_ROW_INDEX
    For rowIndex = 1 To HEADER_SCAN_ROWS
        Set map = ReadHeaderMap(xlSheet, rowIndex)
        If Not map Is Nothing Then
            If map.Exists("MODEL_PATH") Then
                DetectHeaderRowIndex = rowIndex
                Exit Function
            End If
        End If
    Next rowIndex
End Function

Private Function ReadHeaderMap(ByVal xlSheet As Object, ByVal headerRow As Long) As Object
    Dim map As Object
    Dim col As Long
    Dim headerValue As String
    Dim canonicalName As String
    Dim lastCol As Long

    Set map = CreateObject("Scripting.Dictionary")
    map.CompareMode = vbTextCompare

    lastCol = LastUsedColumn(xlSheet, headerRow)
    For col = 1 To lastCol
        headerValue = ReadCellText(xlSheet, headerRow, col)
        canonicalName = ResolveHeaderAlias(headerValue)
        If Len(canonicalName) = 0 Then
            canonicalName = NormalizeHeaderKey(headerValue)
        End If

        If Len(canonicalName) > 0 Then
            If Not map.Exists(canonicalName) Then
                map.Add canonicalName, col
            End If
        End If
    Next col

    map("__HEADER_ROW") = headerRow

    Set ReadHeaderMap = map
End Function

Private Function ResolveHeaderAlias(ByVal rawHeader As String) As String
    Dim normalized As String

    normalized = NormalizeHeaderKey(rawHeader)

    Select Case normalized
        Case "MODEL_PATH", "MODEL", "MODELFILE", "MODELPATH", "FILE", "PATH", "FILEPATH", "MODELIPT", "MODELDOC", "P"
            ResolveHeaderAlias = "MODEL_PATH"
        Case "CODE", "DOCNUMBER", "DOCUMENTNUMBER", "NDOC", "NDOC", "N", "N_DOK", "NODOC", "NOMERDOC", "NOMERDOK", "NOMERDOKUMENTA", "DOKNUMBER", "DOK"
            ResolveHeaderAlias = "CODE"
        Case "PROJECT_NAME", "PROJECT", "PROJECTTITLE", "OBJECT", "OBJECTNAME", "OBJECTTITLE", "NAIMENOVANIEOBEKTA", "NAIMENOVANIEPROEKTA", "PROEKT", "OBEKT", "OBECT", "NAIMPROEKTA"
            ResolveHeaderAlias = "PROJECT_NAME"
        Case "DRAWING_NAME", "DRAWING", "DRAWINGTITLE", "SHEETNAME", "DESCRIPTION", "NAIMENOVANIELISTA", "NAIMENOVANIE", "LISTNAME", "TITLE", "NAME"
            ResolveHeaderAlias = "DRAWING_NAME"
        Case "ORG_NAME", "ORG", "ORGANIZATION", "ORGANISATION", "COMPANY", "AUTHORORG", "ORGANIZACIYA", "ORGANIZATSIYA"
            ResolveHeaderAlias = "ORG_NAME"
        Case "STAGE", "PHASE", "STADIYA", "STADIA", "ST"
            ResolveHeaderAlias = "STAGE"
        Case "SHEET", "LIST", "LISTN", "SHEETN", "SHEETNO", "SHEETNUMBER", "LISTNOMER", "NUMSHEET"
            ResolveHeaderAlias = "SHEET"
        Case "SHEETS", "LISTOV", "TOTALSHEETS", "ALLSHEETS", "SHEETSTOTAL", "SHEETCOUNT", "LISTCOUNT"
            ResolveHeaderAlias = "SHEETS"
    End Select
End Function

Private Function NormalizeHeaderKey(ByVal rawHeader As String) As String
    Dim valueText As String

    valueText = UCase$(Trim$(rawHeader))
    valueText = Replace(valueText, "№", "N")
    valueText = Replace(valueText, ".", " ")
    valueText = Replace(valueText, ",", " ")
    valueText = Replace(valueText, ":", " ")
    valueText = Replace(valueText, ";", " ")
    valueText = Replace(valueText, "-", " ")
    valueText = Replace(valueText, "/", " ")
    valueText = Replace(valueText, "\\", " ")
    valueText = Replace(valueText, vbTab, " ")
    valueText = Replace(valueText, vbCr, " ")
    valueText = Replace(valueText, vbLf, " ")

    valueText = TransliterateRuToAscii(valueText)
    valueText = CollapseSpaces(valueText)
    valueText = Replace(valueText, " ", "_")

    Do While InStr(valueText, "__") > 0
        valueText = Replace(valueText, "__", "_")
    Loop

    If Left$(valueText, 1) = "_" Then valueText = Mid$(valueText, 2)
    If Right$(valueText, 1) = "_" Then valueText = Left$(valueText, Len(valueText) - 1)

    NormalizeHeaderKey = valueText
End Function

Private Function CollapseSpaces(ByVal valueText As String) As String
    valueText = Trim$(valueText)
    Do While InStr(valueText, "  ") > 0
        valueText = Replace(valueText, "  ", " ")
    Loop
    CollapseSpaces = valueText
End Function

Private Function TransliterateRuToAscii(ByVal valueText As String) As String
    Dim src As Variant
    Dim dst As Variant
    Dim i As Long

    src = Array("А", "Б", "В", "Г", "Д", "Е", "Ё", "Ж", "З", "И", "Й", "К", "Л", "М", "Н", "О", "П", "Р", "С", "Т", "У", "Ф", "Х", "Ц", "Ч", "Ш", "Щ", "Ъ", "Ы", "Ь", "Э", "Ю", "Я")
    dst = Array("A", "B", "V", "G", "D", "E", "E", "ZH", "Z", "I", "Y", "K", "L", "M", "N", "O", "P", "R", "S", "T", "U", "F", "H", "C", "CH", "SH", "SCH", "", "Y", "", "E", "YU", "YA")

    For i = LBound(src) To UBound(src)
        valueText = Replace(valueText, CStr(src(i)), CStr(dst(i)))
    Next i

    TransliterateRuToAscii = valueText
End Function

Private Function CreatePromptMapFromRow(ByVal xlSheet As Object, ByVal headerMap As Object, ByVal rowIndex As Long) As Object
    Dim promptMap As Object

    Set promptMap = CreateObject("Scripting.Dictionary")
    promptMap.CompareMode = vbTextCompare

    LoadPromptValue xlSheet, headerMap, rowIndex, promptMap, "CODE"
    LoadPromptValue xlSheet, headerMap, rowIndex, promptMap, "PROJECT_NAME"
    LoadPromptValue xlSheet, headerMap, rowIndex, promptMap, "DRAWING_NAME"
    LoadPromptValue xlSheet, headerMap, rowIndex, promptMap, "ORG_NAME"
    LoadPromptValue xlSheet, headerMap, rowIndex, promptMap, "STAGE"
    LoadPromptValue xlSheet, headerMap, rowIndex, promptMap, "SHEET"
    LoadPromptValue xlSheet, headerMap, rowIndex, promptMap, "SHEETS"

    Set CreatePromptMapFromRow = promptMap
End Function

Private Sub LoadPromptValue(ByVal xlSheet As Object, ByVal headerMap As Object, ByVal rowIndex As Long, ByVal promptMap As Object, ByVal keyName As String)
    If headerMap.Exists(keyName) Then
        promptMap(keyName) = Trim$(ReadCellText(xlSheet, rowIndex, CLng(headerMap(keyName))))
    End If
End Sub

Private Function ResolveModelPath(ByVal inputPath As String, ByVal workspacePath As String, ByVal excelPath As String) As String
    Dim fso As Object
    Dim excelFolder As String

    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(inputPath) Then
        ResolveModelPath = fso.GetAbsolutePathName(inputPath)
        Exit Function
    End If

    If Len(workspacePath) > 0 Then
        If fso.FileExists(fso.BuildPath(workspacePath, inputPath)) Then
            ResolveModelPath = fso.GetAbsolutePathName(fso.BuildPath(workspacePath, inputPath))
            Exit Function
        End If
    End If

    excelFolder = fso.GetParentFolderName(excelPath)
    If Len(excelFolder) > 0 Then
        If fso.FileExists(fso.BuildPath(excelFolder, inputPath)) Then
            ResolveModelPath = fso.GetAbsolutePathName(fso.BuildPath(excelFolder, inputPath))
            Exit Function
        End If
    End If

    Debug.Print "LOG: MODEL_PATH not found, skip row: " & inputPath
End Function

Private Function LastUsedColumn(ByVal xlSheet As Object, ByVal rowIndex As Long) As Long
    LastUsedColumn = xlSheet.Cells(rowIndex, xlSheet.Columns.Count).End(-4159).Column
End Function

Private Function LastUsedRowByColumn(ByVal xlSheet As Object, ByVal colIndex As Long) As Long
    LastUsedRowByColumn = xlSheet.Cells(xlSheet.Rows.Count, colIndex).End(-4162).Row
End Function

Private Function ReadCellText(ByVal xlSheet As Object, ByVal rowIndex As Long, ByVal colIndex As Long) As String
    On Error Resume Next
    ReadCellText = CStr(xlSheet.Cells(rowIndex, colIndex).Value)
    On Error GoTo 0
End Function

Private Sub CloseExcelObjects(ByRef xlBook As Object, ByRef xlApp As Object)
    On Error Resume Next
    If Not xlBook Is Nothing Then xlBook.Close False
    If Not xlApp Is Nothing Then xlApp.Quit
    Set xlBook = Nothing
    Set xlApp = Nothing
    On Error GoTo 0
End Sub
