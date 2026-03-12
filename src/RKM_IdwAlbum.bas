Attribute VB_Name = "RKM_IdwAlbum"
Option Explicit

Private Const ALBUM_SHEET_PREFIX As String = "ALB_"
Private Const MODEL_EXT As String = ".ipt"
Private Const GAP_MM As Double = 8#
Private Const TITLE_SAFE_TOP_MM As Double = 60#
Private Const TITLE_SAFE_LEFT_MM As Double = 230#
Private Const LAYOUT_PAD_MM As Double = 6#

Public Sub Rkm_BuildOrUpdateIdwAlbum()
    Dim oDoc As DrawingDocument
    Dim modelPaths() As String
    Dim modelItems As Collection
    Dim modelCount As Long
    Dim i As Long
    Dim item As Object

    Set oDoc = GetActiveDrawingDocument(ThisApplication)
    If oDoc Is Nothing Then Exit Sub

    modelCount = CollectNumberedIptPaths(modelPaths)
    If modelCount = 0 Then
        Debug.Print "LOG: No numbered IPT files found in active project workspace."
        Exit Sub
    End If

    Set modelItems = New Collection
    For i = 1 To modelCount
        Set item = CreateObject("Scripting.Dictionary")
        item.CompareMode = vbTextCompare
        item("MODEL_PATH") = modelPaths(i)
        item("PROMPTS") = CreateObject("Scripting.Dictionary")
        modelItems.Add item
    Next i

    BuildOrUpdateAlbumCore oDoc, modelItems
End Sub

Public Sub Rkm_BuildOrUpdateIdwAlbum_FromExcel(ByVal oDoc As DrawingDocument, ByVal excelPath As String)
    Dim modelItems As Collection
    Dim workspacePath As String

    If oDoc Is Nothing Then Exit Sub

    workspacePath = GetProjectWorkspacePath()
    Set modelItems = LoadAlbumItemsFromExcel(excelPath, workspacePath, "ALBUM")
    If modelItems Is Nothing Then
        MsgBox "Album build failed (Err 0): Excel parsing returned no data.", vbCritical
        Exit Sub
    End If

    If modelItems.Count = 0 Then
        MsgBox "Excel does not contain valid MODEL_PATH rows.", vbExclamation
        Exit Sub
    End If

    BuildOrUpdateAlbumCore oDoc, modelItems
End Sub

Public Sub BuildOrUpdateAlbumCore(ByVal oDoc As DrawingDocument, ByVal modelItems As Collection)
    Dim i As Long
    Dim oSheet As Sheet
    Dim oModelDoc As Document
    Dim borderDef As BorderDefinition
    Dim titleDef As TitleBlockDefinition
    Dim hadFatalError As Boolean
    Dim fatalMessage As String
    Dim fatalErrNumber As Long
    Dim item As Object
    Dim modelPath As String
    Dim promptMap As Object
    Dim activeStage As String

    On Error GoTo EH
    ThisApplication.SilentOperation = True

    If oDoc Is Nothing Then GoTo CleanUp
    If modelItems Is Nothing Then GoTo CleanUp
    If modelItems.Count = 0 Then GoTo CleanUp
    If Not CanEditDrawingResources(ThisApplication) Then GoTo CleanUp

    Set borderDef = EnsureRkmBorderDefinition(oDoc)
    If borderDef Is Nothing Then
        Err.Raise vbObjectError + 3300, "BuildOrUpdateAlbumCore", "BorderDefinition was not created."
    End If

    Set titleDef = EnsureRkmTitleBlockDefinition(oDoc)
    If titleDef Is Nothing Then
        Err.Raise vbObjectError + 3301, "BuildOrUpdateAlbumCore", "TitleBlockDefinition was not created."
    End If

    For i = 1 To modelItems.Count
        Set item = modelItems.Item(i)
        modelPath = CStr(item("MODEL_PATH"))
        Set promptMap = ResolvePromptMap(item, i, modelItems.Count)

        activeStage = "ensure sheet"
        Set oSheet = EnsureAlbumSheet(oDoc, modelPath)
        If oSheet Is Nothing Then
            Debug.Print "LOG: skip, sheet failed; model=" & modelPath
            GoTo ContinueLoop
        End If

        activeStage = "prepare sheet"
        oSheet.Activate
        On Error Resume Next
        oSheet.Size = kA3DrawingSheetSize
        oSheet.Orientation = kLandscapePageOrientation
        On Error GoTo EH

        activeStage = "remove views"
        RemoveAllDrawingViews oSheet

        activeStage = "apply border"
        ApplyRkmBorderToSheetSafe oSheet, borderDef

        activeStage = "apply title"
        ApplyRkmTitleBlockToSheetWithPrompts oSheet, titleDef, promptMap

        activeStage = "open model"
        Set oModelDoc = OpenModelDocument(modelPath)
        If oModelDoc Is Nothing Then
            Debug.Print "LOG: Skip model (open failed): " & modelPath
            GoTo ContinueLoop
        End If

        activeStage = "add views"
        BuildSheetViews oDoc, oSheet, oModelDoc

ContinueLoop:
        Debug.Print "LOG: done row=" & CStr(i) & "; stage=" & activeStage & "; sheet=" & SafeSheetName(oSheet) & "; model=" & modelPath
        Set oModelDoc = Nothing
        Set oSheet = Nothing
    Next i

    RemoveStaleAlbumSheetsByItems oDoc, modelItems

    Debug.Print "LOG: IDW album build/update completed: " & CStr(modelItems.Count) & " sheets."
    GoTo CleanUp
EH:
    ThisApplication.SilentOperation = False
    hadFatalError = True
    fatalErrNumber = Err.Number
    fatalMessage = Err.Description
    Debug.Print "LOG: Album build failed; stage=" & activeStage & "; sheet=" & SafeSheetName(oSheet) & "; model=" & modelPath & "; Err=" & CStr(fatalErrNumber) & "; " & fatalMessage

CleanUp:
    ThisApplication.SilentOperation = False
    If hadFatalError Then
        MsgBox "Album build failed (Err " & CStr(fatalErrNumber) & "): " & fatalMessage, vbCritical
    End If
End Sub

Private Sub ApplyRkmBorderToSheetSafe(ByVal oSheet As Sheet, ByVal oDef As BorderDefinition)
    On Error GoTo EH
    ApplyRkmBorderToSheet oSheet, oDef
    Exit Sub
EH:
    Debug.Print "LOG: Apply border failed; sheet=" & SafeSheetName(oSheet) & "; Err=" & Err.Number & "; " & Err.Description
    Err.Raise Err.Number, "ApplyRkmBorderToSheetSafe", Err.Description
End Sub

Private Function ResolvePromptMap(ByVal item As Object, ByVal itemIndex As Long, ByVal totalItems As Long) As Object
    Dim result As Object

    Set result = DefaultPromptMap()

    If Not item Is Nothing Then
        If item.Exists("PROMPTS") Then MergePromptMaps result, item("PROMPTS")
        If item.Exists("SHEET") Then
            If Len(Trim$(CStr(item("SHEET")))) > 0 Then result("SHEET") = Trim$(CStr(item("SHEET")))
        End If
        If item.Exists("SHEETS") Then
            If Len(Trim$(CStr(item("SHEETS")))) > 0 Then result("SHEETS") = Trim$(CStr(item("SHEETS")))
        End If
    End If

    If Len(Trim$(CStr(result("SHEET")))) = 0 Then result("SHEET") = CStr(itemIndex)
    If Len(Trim$(CStr(result("SHEETS")))) = 0 Then result("SHEETS") = CStr(totalItems)

    Set ResolvePromptMap = result
End Function

Private Sub MergePromptMaps(ByVal targetMap As Object, ByVal sourceMap As Object)
    Dim key As Variant
    Dim keyName As String

    If targetMap Is Nothing Then Exit Sub
    If sourceMap Is Nothing Then Exit Sub

    For Each key In sourceMap.Keys
        keyName = CStr(key)
        If targetMap.Exists(keyName) Then
            If Len(Trim$(CStr(sourceMap(key)))) > 0 Then
                targetMap(keyName) = CStr(sourceMap(key))
            End If
        End If
    Next key
End Sub

Private Function SafeSheetName(ByVal oSheet As Sheet) As String
    If oSheet Is Nothing Then
        SafeSheetName = "<none>"
    Else
        SafeSheetName = oSheet.Name
    End If
End Function

Private Sub BuildSheetViews(ByVal oDoc As DrawingDocument, ByVal oSheet As Sheet, ByVal oModelDoc As Document)
    Dim scaleCandidates As Variant
    Dim i As Long
    Dim scaleValue As Double
    Dim trialVariant As Long
    Dim bestVariant As Long
    Dim bestScale As Double
    Dim bestHasRight As Boolean
    Dim bestHasTop As Boolean
    Dim baseFits As Boolean
    Dim rightFits As Boolean
    Dim topFits As Boolean
    Dim trialArea As Double
    Dim bestArea As Double
    Dim blockedRect As Object
    Dim mainRect As Object
    Dim rightRect As Object
    Dim topRect As Object

    scaleCandidates = Array(5#, 4#, 3#, 2#, 1.5, 1.25, 1#, 0.9, 0.8, 0.75, 0.67, 0.5, 0.4, 0.33, 0.25, 0.2, 0.1)

    Set blockedRect = GetTitleBlockBlockedRectCm(oDoc)
    Set mainRect = GetMainViewRectCm(oDoc)
    Set rightRect = GetRightProjectedRectCm(oDoc)
    Set topRect = GetTopProjectedRectCm(oDoc)

    For i = LBound(scaleCandidates) To UBound(scaleCandidates)
        scaleValue = CDbl(scaleCandidates(i))
        EvaluateLayoutTrial oSheet, oModelDoc, scaleValue, mainRect, rightRect, topRect, blockedRect, baseFits, rightFits, topFits

        Debug.Print "LOG: layout trial; scale=" & CStr(scaleValue) & "; baseFits=" & CStr(baseFits) & "; rightFits=" & CStr(rightFits) & "; topFits=" & CStr(topFits)

        If baseFits Then
            trialVariant = 1
            trialArea = 0#
            If rightFits Then trialVariant = trialVariant + 1
            If topFits Then trialVariant = trialVariant + 1

            trialArea = scaleValue

            If (trialVariant > bestVariant) Or ((trialVariant = bestVariant) And (trialArea > bestArea)) Then
                bestVariant = trialVariant
                bestScale = scaleValue
                bestHasRight = rightFits
                bestHasTop = topFits
                bestArea = trialArea
            End If

            If bestVariant = 3 Then Exit For
        End If
    Next i

    If bestVariant = 0 Then
        Debug.Print "LOG: Skipping model " & oModelDoc.DisplayName & " - View placement failed."
        Exit Sub
    End If

    PlaceSelectedLayout oSheet, oModelDoc, bestScale, bestHasRight, bestHasTop, mainRect, rightRect, topRect, blockedRect

    If bestVariant = 3 Then
        Debug.Print "LOG: layout selected; variant=BASE+RIGHT+TOP; scale=" & CStr(bestScale)
    ElseIf bestVariant = 2 Then
        If bestHasRight Then
            Debug.Print "LOG: layout selected; variant=BASE+RIGHT; scale=" & CStr(bestScale)
        Else
            Debug.Print "LOG: layout selected; variant=BASE+TOP; scale=" & CStr(bestScale)
        End If
    Else
        Debug.Print "LOG: layout selected; variant=BASE; scale=" & CStr(bestScale)
    End If
End Sub

Private Sub EvaluateLayoutTrial(ByVal oSheet As Sheet, ByVal oModelDoc As Document, ByVal scaleValue As Double, _
                                ByVal mainRect As Object, ByVal rightRect As Object, ByVal topRect As Object, ByVal blockedRect As Object, _
                                ByRef baseFits As Boolean, ByRef rightFits As Boolean, ByRef topFits As Boolean)
    Dim baseView As DrawingView
    Dim rightView As DrawingView
    Dim topView As DrawingView

    baseFits = False
    rightFits = False
    topFits = False

    Set baseView = TryCreateBaseView(oSheet, oModelDoc, scaleValue)
    If baseView Is Nothing Then Exit Sub

    baseView.Center = Pt(RectCenterX(mainRect), RectCenterY(mainRect))
    If DoesViewFitRect(baseView, mainRect) And (Not IsViewIntersectingBlockedArea(baseView, blockedRect)) Then
        baseFits = True

        Set rightView = TryAddOneProjected(oSheet, baseView, rightRect, blockedRect, Pt(RectCenterX(rightRect), RectCenterY(rightRect)))
        Set topView = TryAddOneProjected(oSheet, baseView, topRect, blockedRect, Pt(RectCenterX(topRect), RectCenterY(topRect)))

        rightFits = Not rightView Is Nothing
        topFits = Not topView Is Nothing
    End If

    On Error Resume Next
    If Not topView Is Nothing Then topView.Delete
    If Not rightView Is Nothing Then rightView.Delete
    If Not baseView Is Nothing Then baseView.Delete
    On Error GoTo 0
End Sub

Private Sub PlaceSelectedLayout(ByVal oSheet As Sheet, ByVal oModelDoc As Document, ByVal scaleValue As Double, _
                                ByVal addRight As Boolean, ByVal addTop As Boolean, _
                                ByVal mainRect As Object, ByVal rightRect As Object, ByVal topRect As Object, ByVal blockedRect As Object)
    Dim baseView As DrawingView

    Set baseView = TryCreateBaseView(oSheet, oModelDoc, scaleValue)
    If baseView Is Nothing Then Exit Sub

    baseView.Center = Pt(RectCenterX(mainRect), RectCenterY(mainRect))
    If Not DoesViewFitRect(baseView, mainRect) Or IsViewIntersectingBlockedArea(baseView, blockedRect) Then
        baseView.Delete
        Exit Sub
    End If

    If addRight Then
        Call TryAddOneProjected(oSheet, baseView, rightRect, blockedRect, Pt(RectCenterX(rightRect), RectCenterY(rightRect)))
    End If

    If addTop Then
        Call TryAddOneProjected(oSheet, baseView, topRect, blockedRect, Pt(RectCenterX(topRect), RectCenterY(topRect)))
    End If
End Sub

Private Function TryCreateBaseView(ByVal oSheet As Sheet, ByVal oModelDoc As Document, ByVal scaleValue As Double) As DrawingView
    Dim centerPt As Point2d

    On Error GoTo EH
    Set centerPt = Pt(oSheet.Width / 2#, oSheet.Height / 2#)

    Set TryCreateBaseView = oSheet.DrawingViews.AddBaseView( _
        oModelDoc, centerPt, scaleValue, kFrontViewOrientation, kHiddenLineRemovedDrawingViewStyle)
    Exit Function
EH:
    ThisApplication.SilentOperation = False
    Debug.Print "LOG: AddBaseView failed; sheet=" & SafeSheetName(oSheet) & "; model=" & oModelDoc.DisplayName & "; Err=" & Err.Number & "; " & Err.Description
    Set TryCreateBaseView = Nothing
End Function

Private Function TryAddOneProjected(ByVal oSheet As Sheet, ByVal baseView As DrawingView, ByVal targetRect As Object, ByVal blockedRect As Object, ByVal targetPt As Point2d) As DrawingView
    Dim projView As DrawingView

    On Error GoTo EH

    Set projView = oSheet.DrawingViews.AddProjectedView(baseView, targetPt, kHiddenLineRemovedDrawingViewStyle)
    If projView Is Nothing Then Exit Function

    If Not DoesViewFitRect(projView, targetRect) Then
        projView.Delete
        Set projView = Nothing
        Exit Function
    End If

    If IsViewIntersectingBlockedArea(projView, blockedRect) Then
        projView.Delete
        Set projView = Nothing
    End If

    Set TryAddOneProjected = projView
    Exit Function
EH:
    ThisApplication.SilentOperation = False
    Debug.Print "LOG: AddProjectedView failed; sheet=" & SafeSheetName(oSheet) & "; Err=" & Err.Number & "; " & Err.Description
    On Error Resume Next
    If Not projView Is Nothing Then projView.Delete
    On Error GoTo 0
    Set TryAddOneProjected = Nothing
End Sub

Private Function IsViewInsideSafeArea(ByVal oDoc As DrawingDocument, ByVal oView As DrawingView) As Boolean
    IsViewInsideSafeArea = DoesViewFitRect(oView, GetSheetSafeRectCm(oDoc))
End Function

Private Function IsViewInTitleArea(ByVal oDoc As DrawingDocument, ByVal oView As DrawingView) As Boolean
    IsViewInTitleArea = IsViewIntersectingBlockedArea(oView, GetTitleBlockBlockedRectCm(oDoc))
End Function

Private Function GetSheetSafeRectCm(ByVal oDoc As DrawingDocument) As Object
    Dim oSheet As Sheet

    Set oSheet = oDoc.ActiveSheet
    Set GetSheetSafeRectCm = CreateRect( _
        MmToCm(oDoc, FRAME_LEFT_MM), _
        oSheet.Width - MmToCm(oDoc, FRAME_OTHER_MM), _
        MmToCm(oDoc, FRAME_OTHER_MM), _
        oSheet.Height - MmToCm(oDoc, FRAME_OTHER_MM))
End Function

Private Function GetTitleBlockBlockedRectCm(ByVal oDoc As DrawingDocument) As Object
    Dim safeRect As Object

    Set safeRect = GetSheetSafeRectCm(oDoc)
    Set GetTitleBlockBlockedRectCm = CreateRect( _
        MmToCm(oDoc, TITLE_SAFE_LEFT_MM), _
        safeRect("Right"), _
        safeRect("Bottom"), _
        MmToCm(oDoc, TITLE_SAFE_TOP_MM))
End Function

Private Function GetMainViewRectCm(ByVal oDoc As DrawingDocument) As Object
    Dim safeRect As Object
    Dim splitX As Double
    Dim splitY As Double
    Dim padCm As Double

    Set safeRect = InsetRect(GetSheetSafeRectCm(oDoc), MmToCm(oDoc, LAYOUT_PAD_MM))
    padCm = MmToCm(oDoc, GAP_MM)
    splitX = safeRect("Left") + RectWidth(safeRect) * 0.67
    splitY = safeRect("Bottom") + RectHeight(safeRect) * 0.6

    Set GetMainViewRectCm = CreateRect(safeRect("Left"), splitX - padCm, safeRect("Bottom"), splitY - padCm)
End Function

Private Function GetTopProjectedRectCm(ByVal oDoc As DrawingDocument) As Object
    Dim safeRect As Object
    Dim splitX As Double
    Dim splitY As Double
    Dim padCm As Double

    Set safeRect = InsetRect(GetSheetSafeRectCm(oDoc), MmToCm(oDoc, LAYOUT_PAD_MM))
    padCm = MmToCm(oDoc, GAP_MM)
    splitX = safeRect("Left") + RectWidth(safeRect) * 0.67
    splitY = safeRect("Bottom") + RectHeight(safeRect) * 0.6

    Set GetTopProjectedRectCm = CreateRect(safeRect("Left"), splitX - padCm, splitY + padCm, safeRect("Top"))
End Function

Private Function GetRightProjectedRectCm(ByVal oDoc As DrawingDocument) As Object
    Dim safeRect As Object
    Dim blockedRect As Object
    Dim splitX As Double
    Dim splitY As Double
    Dim padCm As Double
    Dim bottomLimit As Double

    Set safeRect = InsetRect(GetSheetSafeRectCm(oDoc), MmToCm(oDoc, LAYOUT_PAD_MM))
    Set blockedRect = GetTitleBlockBlockedRectCm(oDoc)
    padCm = MmToCm(oDoc, GAP_MM)
    splitX = safeRect("Left") + RectWidth(safeRect) * 0.67
    splitY = safeRect("Bottom") + RectHeight(safeRect) * 0.6
    bottomLimit = blockedRect("Top") + padCm

    If bottomLimit < safeRect("Bottom") Then bottomLimit = safeRect("Bottom")

    Set GetRightProjectedRectCm = CreateRect(splitX + padCm, safeRect("Right"), bottomLimit, safeRect("Top"))
End Function

Private Function RectWidth(ByVal rect As Object) As Double
    RectWidth = rect("Right") - rect("Left")
End Function

Private Function RectHeight(ByVal rect As Object) As Double
    RectHeight = rect("Top") - rect("Bottom")
End Function

Private Function RectCenterX(ByVal rect As Object) As Double
    RectCenterX = rect("Left") + RectWidth(rect) / 2#
End Function

Private Function RectCenterY(ByVal rect As Object) As Double
    RectCenterY = rect("Bottom") + RectHeight(rect) / 2#
End Function

Private Function InsetRect(ByVal rect As Object, ByVal deltaCm As Double) As Object
    Set InsetRect = CreateRect(rect("Left") + deltaCm, rect("Right") - deltaCm, rect("Bottom") + deltaCm, rect("Top") - deltaCm)
End Function

Private Function RectanglesIntersect(ByVal rectA As Object, ByVal rectB As Object) As Boolean
    RectanglesIntersect = Not (rectA("Right") <= rectB("Left") Or rectB("Right") <= rectA("Left") Or rectA("Top") <= rectB("Bottom") Or rectB("Top") <= rectA("Bottom"))
End Function

Private Function DoesViewFitRect(ByVal oView As DrawingView, ByVal rect As Object) As Boolean
    Dim viewRect As Object

    If oView Is Nothing Then Exit Function
    Set viewRect = CreateRect(oView.Left, oView.Left + oView.Width, oView.Top - oView.Height, oView.Top)
    DoesViewFitRect = (viewRect("Left") >= rect("Left")) And (viewRect("Right") <= rect("Right")) And (viewRect("Bottom") >= rect("Bottom")) And (viewRect("Top") <= rect("Top"))
End Function

Private Function IsViewIntersectingBlockedArea(ByVal oView As DrawingView, ByVal blockedRect As Object) As Boolean
    Dim viewRect As Object

    If oView Is Nothing Then Exit Function
    Set viewRect = CreateRect(oView.Left, oView.Left + oView.Width, oView.Top - oView.Height, oView.Top)
    IsViewIntersectingBlockedArea = RectanglesIntersect(viewRect, blockedRect)
End Function

Private Function CreateRect(ByVal leftCm As Double, ByVal rightCm As Double, ByVal bottomCm As Double, ByVal topCm As Double) As Object
    Dim rect As Object

    Set rect = CreateObject("Scripting.Dictionary")
    rect.CompareMode = vbTextCompare
    rect("Left") = leftCm
    rect("Right") = rightCm
    rect("Bottom") = bottomCm
    rect("Top") = topCm
    rect("Width") = rightCm - leftCm
    rect("Height") = topCm - bottomCm
    Set CreateRect = rect
End Function

Private Sub RemoveAllDrawingViews(ByVal oSheet As Sheet)
    Dim i As Long

    If oSheet Is Nothing Then Exit Sub

    For i = oSheet.DrawingViews.Count To 1 Step -1
        oSheet.DrawingViews.Item(i).Delete
    Next i
End Sub

Private Function EnsureAlbumSheet(ByVal oDoc As DrawingDocument, ByVal modelPath As String) As Sheet
    Dim sheetName As String
    Dim oSheet As Sheet

    sheetName = MakeAlbumSheetName(modelPath)
    Set oSheet = FindSheetByName(oDoc, sheetName)

    If oSheet Is Nothing Then
        Set oSheet = oDoc.Sheets.Add(kA3DrawingSheetSize, kLandscapePageOrientation)
        oSheet.Name = sheetName
    End If

    Set EnsureAlbumSheet = oSheet
End Function

Private Function FindSheetByName(ByVal oDoc As DrawingDocument, ByVal sheetName As String) As Sheet
    Dim i As Long

    If oDoc Is Nothing Then Exit Function

    For i = 1 To oDoc.Sheets.Count
        If StrComp(Split(oDoc.Sheets.Item(i).Name, ":")(0), sheetName, vbTextCompare) = 0 Then
            Set FindSheetByName = oDoc.Sheets.Item(i)
            Exit Function
        End If
    Next i
End Function

Private Sub RemoveStaleAlbumSheetsByItems(ByVal oDoc As DrawingDocument, ByVal modelItems As Collection)
    Dim i As Long
    Dim oSheet As Sheet

    If oDoc Is Nothing Then Exit Sub

    For i = oDoc.Sheets.Count To 1 Step -1
        Set oSheet = oDoc.Sheets.Item(i)
        If IsAlbumSheet(oSheet.Name) Then
            If Not IsSheetBackedByItems(oSheet.Name, modelItems) Then
                oSheet.Delete
            End If
        End If
    Next i
End Sub

Private Function IsSheetBackedByItems(ByVal sheetName As String, ByVal modelItems As Collection) As Boolean
    Dim i As Long
    Dim item As Object

    If modelItems Is Nothing Then Exit Function

    For i = 1 To modelItems.Count
        Set item = modelItems.Item(i)
        If StrComp(Split(sheetName, ":")(0), MakeAlbumSheetName(CStr(item("MODEL_PATH"))), vbTextCompare) = 0 Then
            IsSheetBackedByItems = True
            Exit Function
        End If
    Next i
End Function

Private Function IsAlbumSheet(ByVal sheetName As String) As Boolean
    IsAlbumSheet = (UCase$(Left$(sheetName, Len(ALBUM_SHEET_PREFIX))) = ALBUM_SHEET_PREFIX)
End Function

Private Function MakeAlbumSheetName(ByVal modelPath As String) As String
    MakeAlbumSheetName = ALBUM_SHEET_PREFIX & BaseName(modelPath)
End Function

Private Function BaseName(ByVal filePath As String) As String
    Dim fileName As String
    Dim dotPos As Long
    Dim slashPos As Long

    slashPos = InStrRev(filePath, "\")
    If slashPos > 0 Then
        fileName = Mid$(filePath, slashPos + 1)
    Else
        fileName = filePath
    End If

    dotPos = InStrRev(fileName, ".")
    If dotPos > 0 Then
        BaseName = Left$(fileName, dotPos - 1)
    Else
        BaseName = fileName
    End If
End Function

Private Function CollectNumberedIptPaths(ByRef outPaths() As String) As Long
    Dim projectRoot As String
    Dim fso As Object
    Dim rootFolder As Object
    Dim bag As Collection

    projectRoot = GetProjectWorkspacePath()
    If Len(projectRoot) = 0 Then Exit Function

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(projectRoot) Then Exit Function

    Set rootFolder = fso.GetFolder(projectRoot)
    Set bag = New Collection

    CollectIptRecursive rootFolder, bag
    CopySortedCollectionToArray bag, outPaths

    CollectNumberedIptPaths = bag.Count
End Function

Private Sub CollectIptRecursive(ByVal folderObj As Object, ByVal bag As Collection)
    Dim subFolder As Object
    Dim fileObj As Object
    Dim baseNameText As String

    For Each fileObj In folderObj.Files
        If LCase$(Right$(fileObj.Name, Len(MODEL_EXT))) = MODEL_EXT Then
            baseNameText = BaseName(CStr(fileObj.Name))
            If HasNumericPrefix(baseNameText) And Not IsVersionedNumericPattern(baseNameText) Then
                bag.Add CStr(fileObj.Path)
            End If
        End If
    Next fileObj

    For Each subFolder In folderObj.SubFolders
        CollectIptRecursive subFolder, bag
    Next subFolder
End Sub

Private Function IsVersionedNumericPattern(ByVal baseNameText As String) As Boolean
    Dim parts() As String

    parts = Split(baseNameText, ".")
    If UBound(parts) <> 1 Then Exit Function

    IsVersionedNumericPattern = (Len(parts(0)) = 3 And IsDigitsOnly(parts(0)) And Len(parts(1)) = 4 And IsDigitsOnly(parts(1)))
End Function

Private Function IsDigitsOnly(ByVal valueText As String) As Boolean
    Dim i As Long
    Dim ch As String

    If Len(valueText) = 0 Then Exit Function

    IsDigitsOnly = True
    For i = 1 To Len(valueText)
        ch = Mid$(valueText, i, 1)
        If ch < "0" Or ch > "9" Then
            IsDigitsOnly = False
            Exit Function
        End If
    Next i
End Function

Private Function HasNumericPrefix(ByVal fileNameWithoutExt As String) As Boolean
    Dim i As Long
    Dim ch As String

    For i = 1 To Len(fileNameWithoutExt)
        ch = Mid$(fileNameWithoutExt, i, 1)
        If ch < "0" Or ch > "9" Then Exit For
        HasNumericPrefix = True
    Next i
End Function

Private Sub CopySortedCollectionToArray(ByVal bag As Collection, ByRef outPaths() As String)
    Dim count As Long
    Dim i As Long
    Dim j As Long
    Dim temp As String

    count = bag.Count
    If count = 0 Then Exit Sub

    ReDim outPaths(1 To count)

    For i = 1 To count
        outPaths(i) = CStr(bag.Item(i))
    Next i

    For i = 1 To count - 1
        For j = i + 1 To count
            If CompareModelPaths(outPaths(i), outPaths(j)) > 0 Then
                temp = outPaths(i)
                outPaths(i) = outPaths(j)
                outPaths(j) = temp
            End If
        Next j
    Next i
End Sub

Private Function CompareModelPaths(ByVal aPath As String, ByVal bPath As String) As Long
    Dim aName As String
    Dim bName As String
    Dim aNum As Long
    Dim bNum As Long

    aName = BaseName(aPath)
    bName = BaseName(bPath)

    aNum = LeadingNumber(aName)
    bNum = LeadingNumber(bName)

    If aNum < bNum Then
        CompareModelPaths = -1
    ElseIf aNum > bNum Then
        CompareModelPaths = 1
    Else
        CompareModelPaths = StrComp(aName, bName, vbTextCompare)
    End If
End Function

Private Function LeadingNumber(ByVal valueText As String) As Long
    Dim i As Long
    Dim ch As String
    Dim numText As String

    For i = 1 To Len(valueText)
        ch = Mid$(valueText, i, 1)
        If ch >= "0" And ch <= "9" Then
            numText = numText & ch
        Else
            Exit For
        End If
    Next i

    If Len(numText) = 0 Then
        LeadingNumber = 0
    Else
        LeadingNumber = CLng(numText)
    End If
End Function

Public Function GetProjectWorkspacePath() As String
    Dim oProj As DesignProject

    On Error GoTo EH
    Set oProj = ThisApplication.DesignProjectManager.ActiveDesignProject
    If oProj Is Nothing Then Exit Function

    GetProjectWorkspacePath = oProj.WorkspacePath
    Exit Function
EH:
    ThisApplication.SilentOperation = False
    GetProjectWorkspacePath = ""
End Function

Private Function OpenModelDocument(ByVal modelPath As String) As Document
    Dim i As Long
    Dim oDoc As Document
    Dim previousSilentOperation As Boolean

    For i = 1 To ThisApplication.Documents.Count
        Set oDoc = ThisApplication.Documents.Item(i)
        If StrComp(oDoc.FullFileName, modelPath, vbTextCompare) = 0 Then
            Set OpenModelDocument = oDoc
            Exit Function
        End If
    Next i

    On Error GoTo EH
    previousSilentOperation = ThisApplication.SilentOperation
    ThisApplication.SilentOperation = True
    Set OpenModelDocument = ThisApplication.Documents.Open(modelPath, False)
    ThisApplication.SilentOperation = previousSilentOperation
    Exit Function
EH:
    ThisApplication.SilentOperation = False
    Debug.Print "LOG: Open model failed; path=" & modelPath & "; Err=" & Err.Number & "; " & Err.Description
    Set OpenModelDocument = Nothing
End Function
