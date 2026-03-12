Attribute VB_Name = "RKM_IdwAlbum"
Option Explicit

Private Const ALBUM_SHEET_PREFIX As String = "ALB_"
Private Const MODEL_EXT As String = ".ipt"
Private Const GAP_MM As Double = 8#
Private Const TITLE_SAFE_TOP_MM As Double = 60#
Private Const TITLE_SAFE_LEFT_MM As Double = 230#

Public Sub Rkm_BuildOrUpdateIdwAlbum()
    Dim oDoc As DrawingDocument
    Dim modelPaths() As String
    Dim modelCount As Long
    Dim i As Long
    Dim oSheet As Sheet
    Dim oModelDoc As Document
    Dim borderDef As BorderDefinition
    Dim titleDef As TitleBlockDefinition

    On Error GoTo EH

    Set oDoc = GetActiveDrawingDocument(ThisApplication)
    If oDoc Is Nothing Then Exit Sub

    If Not CanEditDrawingResources(ThisApplication) Then Exit Sub

    modelCount = CollectNumberedIptPaths(modelPaths)
    If modelCount = 0 Then
        MsgBox "No numbered IPT files found in active project workspace.", vbExclamation
        Exit Sub
    End If

    Set borderDef = EnsureRkmBorderDefinition(oDoc)
    If borderDef Is Nothing Then Exit Sub

    Set titleDef = EnsureRkmTitleBlockDefinition(oDoc)
    If titleDef Is Nothing Then Exit Sub

    For i = 1 To modelCount
        Set oSheet = EnsureAlbumSheet(oDoc, modelPaths(i))
        If oSheet Is Nothing Then GoTo ContinueLoop

        oSheet.Activate
        oSheet.ChangeSize kA3DrawingSheetSize, kLandscapePageOrientation

        RemoveAllDrawingViews oSheet
        ApplyRkmBorderToSheet oSheet, borderDef
        ApplyRkmTitleBlockToSheet oSheet, titleDef

        Set oModelDoc = OpenModelDocument(modelPaths(i))
        If oModelDoc Is Nothing Then
            Debug.Print "Skip model (open failed): " & modelPaths(i)
            GoTo ContinueLoop
        End If

        BuildSheetViews oDoc, oSheet, oModelDoc

ContinueLoop:
        Set oModelDoc = Nothing
    Next i

    RemoveStaleAlbumSheets oDoc, modelPaths, modelCount

    MsgBox "IDW album build/update completed: " & CStr(modelCount) & " sheets.", vbInformation
    Exit Sub
EH:
    MsgBox "Album build failed: " & Err.Description, vbCritical
End Sub

Private Sub BuildSheetViews(ByVal oDoc As DrawingDocument, ByVal oSheet As Sheet, ByVal oModelDoc As Document)
    Dim scaleCandidates As Variant
    Dim i As Long
    Dim baseView As DrawingView
    Dim placed As Boolean

    scaleCandidates = Array(2#, 1.5, 1#, 0.75, 0.5, 0.4, 0.33, 0.25, 0.2, 0.1)

    For i = LBound(scaleCandidates) To UBound(scaleCandidates)
        Set baseView = TryCreateBaseView(oSheet, oModelDoc, CDbl(scaleCandidates(i)))
        If Not baseView Is Nothing Then
            If IsViewInsideSafeArea(oDoc, baseView) And Not IsViewInTitleArea(oDoc, baseView) Then
                placed = True
                Exit For
            End If

            baseView.Delete
            Set baseView = Nothing
        End If
    Next i

    If Not placed Then
        MsgBox "Cannot place base view for model: " & oModelDoc.DisplayName, vbExclamation
        Exit Sub
    End If

    TryAddProjectedViews oDoc, oSheet, baseView
End Sub

Private Function TryCreateBaseView(ByVal oSheet As Sheet, ByVal oModelDoc As Document, ByVal scaleValue As Double) As DrawingView
    Dim centerPt As Point2d

    On Error GoTo EH
    Set centerPt = Pt(21.75, 17.6)

    Set TryCreateBaseView = oSheet.DrawingViews.AddBaseView( _
        oModelDoc, centerPt, scaleValue, kFrontViewOrientation, kHiddenLineRemovedDrawingViewStyle)
    Exit Function
EH:
    Set TryCreateBaseView = Nothing
End Function

Private Sub TryAddProjectedViews(ByVal oDoc As DrawingDocument, ByVal oSheet As Sheet, ByVal baseView As DrawingView)
    Dim xMin As Double, xMax As Double, yMin As Double, yMax As Double
    Dim gapCm As Double

    If baseView Is Nothing Then Exit Sub

    xMin = MmToCm(oDoc, FRAME_LEFT_MM)
    xMax = MmToCm(oDoc, A3_WIDTH_MM - FRAME_OTHER_MM)
    yMin = MmToCm(oDoc, FRAME_OTHER_MM)
    yMax = MmToCm(oDoc, A3_HEIGHT_MM - FRAME_OTHER_MM)
    gapCm = MmToCm(oDoc, GAP_MM)

    TryAddOneProjected oDoc, oSheet, baseView, Pt(baseView.Center.X + baseView.Width / 2# + gapCm + baseView.Width / 2#, baseView.Center.Y)
    TryAddOneProjected oDoc, oSheet, baseView, Pt(baseView.Center.X, baseView.Center.Y + baseView.Height / 2# + gapCm + baseView.Height / 2#)
End Sub

Private Sub TryAddOneProjected(ByVal oDoc As DrawingDocument, ByVal oSheet As Sheet, ByVal baseView As DrawingView, ByVal targetPt As Point2d)
    Dim projView As DrawingView

    On Error GoTo EH

    Set projView = oSheet.DrawingViews.AddProjectedView(baseView, targetPt, kHiddenLineRemovedDrawingViewStyle)
    If projView Is Nothing Then Exit Sub

    If Not IsViewInsideSafeArea(oDoc, projView) Then
        projView.Delete
        Exit Sub
    End If

    If IsViewInTitleArea(oDoc, projView) Then
        projView.Delete
    End If

    Exit Sub
EH:
    On Error Resume Next
    If Not projView Is Nothing Then projView.Delete
    On Error GoTo 0
End Sub

Private Function IsViewInsideSafeArea(ByVal oDoc As DrawingDocument, ByVal oView As DrawingView) As Boolean
    Dim leftCm As Double, rightCm As Double, bottomCm As Double, topCm As Double
    Dim xMin As Double, xMax As Double, yMin As Double, yMax As Double

    If oView Is Nothing Then Exit Function

    leftCm = oView.Left
    rightCm = oView.Left + oView.Width
    bottomCm = oView.Top - oView.Height
    topCm = oView.Top

    xMin = MmToCm(oDoc, FRAME_LEFT_MM)
    xMax = MmToCm(oDoc, A3_WIDTH_MM - FRAME_OTHER_MM)
    yMin = MmToCm(oDoc, FRAME_OTHER_MM)
    yMax = MmToCm(oDoc, A3_HEIGHT_MM - FRAME_OTHER_MM)

    IsViewInsideSafeArea = (leftCm >= xMin) And (rightCm <= xMax) And (bottomCm >= yMin) And (topCm <= yMax)
End Function

Private Function IsViewInTitleArea(ByVal oDoc As DrawingDocument, ByVal oView As DrawingView) As Boolean
    Dim leftCm As Double, rightCm As Double, bottomCm As Double, topCm As Double
    Dim titleLeftCm As Double, titleRightCm As Double, titleBottomCm As Double, titleTopCm As Double

    If oView Is Nothing Then Exit Function

    leftCm = oView.Left
    rightCm = oView.Left + oView.Width
    bottomCm = oView.Top - oView.Height
    topCm = oView.Top

    titleLeftCm = MmToCm(oDoc, TITLE_SAFE_LEFT_MM)
    titleRightCm = MmToCm(oDoc, A3_WIDTH_MM - FRAME_OTHER_MM)
    titleBottomCm = MmToCm(oDoc, FRAME_OTHER_MM)
    titleTopCm = MmToCm(oDoc, TITLE_SAFE_TOP_MM)

    IsViewInTitleArea = RectanglesIntersect(leftCm, rightCm, bottomCm, topCm, titleLeftCm, titleRightCm, titleBottomCm, titleTopCm)
End Function

Private Function RectanglesIntersect(ByVal l1 As Double, ByVal r1 As Double, ByVal b1 As Double, ByVal t1 As Double, _
                                     ByVal l2 As Double, ByVal r2 As Double, ByVal b2 As Double, ByVal t2 As Double) As Boolean
    RectanglesIntersect = Not (r1 <= l2 Or r2 <= l1 Or t1 <= b2 Or t2 <= b1)
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
        If StrComp(oDoc.Sheets.Item(i).Name, sheetName, vbTextCompare) = 0 Then
            Set FindSheetByName = oDoc.Sheets.Item(i)
            Exit Function
        End If
    Next i
End Function

Private Sub RemoveStaleAlbumSheets(ByVal oDoc As DrawingDocument, ByRef modelPaths() As String, ByVal modelCount As Long)
    Dim i As Long
    Dim oSheet As Sheet

    If oDoc Is Nothing Then Exit Sub

    For i = oDoc.Sheets.Count To 1 Step -1
        Set oSheet = oDoc.Sheets.Item(i)
        If IsAlbumSheet(oSheet.Name) Then
            If Not IsSheetBackedByModel(oSheet.Name, modelPaths, modelCount) Then
                oSheet.Delete
            End If
        End If
    Next i
End Sub

Private Function IsSheetBackedByModel(ByVal sheetName As String, ByRef modelPaths() As String, ByVal modelCount As Long) As Boolean
    Dim i As Long

    For i = 1 To modelCount
        If StrComp(sheetName, MakeAlbumSheetName(modelPaths(i)), vbTextCompare) = 0 Then
            IsSheetBackedByModel = True
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

    For Each fileObj In folderObj.Files
        If LCase$(Right$(fileObj.Name, Len(MODEL_EXT))) = MODEL_EXT Then
            If HasNumericPrefix(BaseName(CStr(fileObj.Name))) Then
                bag.Add CStr(fileObj.Path)
            End If
        End If
    Next fileObj

    For Each subFolder In folderObj.SubFolders
        CollectIptRecursive subFolder, bag
    Next subFolder
End Sub

Private Function HasNumericPrefix(ByVal fileNameWithoutExt As String) As Boolean
    Dim i As Long
    Dim ch As String

    For i = 1 To Len(fileNameWithoutExt)
        ch = Mid$(fileNameWithoutExt, i, 1)
        If ch < "0" Or ch > "9" Then
            Exit For
        End If
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

Private Function GetProjectWorkspacePath() As String
    Dim oProj As DesignProject

    On Error GoTo EH
    Set oProj = ThisApplication.DesignProjectManager.ActiveDesignProject
    If oProj Is Nothing Then Exit Function

    GetProjectWorkspacePath = oProj.WorkspacePath
    Exit Function
EH:
    GetProjectWorkspacePath = ""
End Function

Private Function OpenModelDocument(ByVal modelPath As String) As Document
    Dim i As Long
    Dim oDoc As Document

    For i = 1 To ThisApplication.Documents.Count
        Set oDoc = ThisApplication.Documents.Item(i)
        If StrComp(oDoc.FullFileName, modelPath, vbTextCompare) = 0 Then
            Set OpenModelDocument = oDoc
            Exit Function
        End If
    Next i

    On Error GoTo EH
    Set OpenModelDocument = ThisApplication.Documents.Open(modelPath, False)
    Exit Function
EH:
    Set OpenModelDocument = Nothing
End Function
