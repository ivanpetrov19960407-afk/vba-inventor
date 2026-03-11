Attribute VB_Name = "RKM_SPDS_A3_Standalone"
Option Explicit

Private Const A3_W_MM As Double = 420#
Private Const A3_H_MM As Double = 297#

Private Const FRAME_LEFT_MM As Double = 20#
Private Const FRAME_OTHER_MM As Double = 5#

' Form 3 geometry constants (mm). Keep these values grouped for easy adjustment.
Private Const TB_W_MM As Double = 185#   ' Full title block width.
Private Const TB_H_MM As Double = 55#    ' Full title block height.
Private Const TB_C1_MM As Double = 110#  ' Vertical split from title block left edge.
Private Const TB_C2_MM As Double = 150#
Private Const TB_C3_MM As Double = 170#
Private Const TB_R1_MM As Double = 15#   ' Horizontal split from title block bottom edge.
Private Const TB_R2_MM As Double = 30#
Private Const TB_R3_MM As Double = 45#

Private Const A3_TOL_MM As Double = 0.05

Public Sub Rkm_CreateOrApplyA3Frame_SPDS()
    Dim oDoc As DrawingDocument
    Dim oSheet As Sheet
    Dim oBorderDef As BorderDefinition
    Dim oTitleDef As TitleBlockDefinition

    On Error GoTo EH

    Set oDoc = EnsureDrawingDocument()
    If oDoc Is Nothing Then Exit Sub

    If Not CanEditDrawingResources(ThisApplication) Then Exit Sub

    Set oSheet = EnsureA3LandscapeSheet(oDoc)
    If oSheet Is Nothing Then Exit Sub

    PrintSheetDiagnostics oDoc, oSheet
    Debug.Print "Sheet.Width (mm): " & Fmt(oDoc.UnitsOfMeasure.ConvertUnits(oSheet.Width, kCentimeterLengthUnits, kMillimeterLengthUnits))
    Debug.Print "Sheet.Height (mm): " & Fmt(oDoc.UnitsOfMeasure.ConvertUnits(oSheet.Height, kCentimeterLengthUnits, kMillimeterLengthUnits))

    Set oBorderDef = EnsureSpdsA3BorderDefinition(oDoc)
    If oBorderDef Is Nothing Then Exit Sub

    Set oTitleDef = EnsureSpdsForm3TitleBlockDefinition(oDoc)
    If oTitleDef Is Nothing Then Exit Sub

    If Not CanEditDrawingResources(ThisApplication) Then Exit Sub

    ApplySpdsBorderToSheet oSheet, oBorderDef
    ApplySpdsTitleBlockToSheet oSheet, oTitleDef

    Debug.Print "Applied BorderDefinition: " & oBorderDef.Name
    Debug.Print "Applied TitleBlockDefinition: " & oTitleDef.Name

    MsgBox "SPDS A3 frame and form 3 title block applied.", vbInformation
    Exit Sub
EH:
    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

Private Function EnsureDrawingDocument() As DrawingDocument
    If ThisApplication.ActiveDocument Is Nothing Then
        MsgBox "Open a drawing document first.", vbExclamation
        Exit Function
    End If

    If ThisApplication.ActiveDocument.DocumentType <> kDrawingDocumentObject Then
        MsgBox "Active document is not DrawingDocument.", vbExclamation
        Exit Function
    End If

    Set EnsureDrawingDocument = ThisApplication.ActiveDocument
End Function

Private Function EnsureA3LandscapeSheet(ByVal oDoc As DrawingDocument) As Sheet
    Dim oSheet As Sheet
    Dim wMm As Double
    Dim hMm As Double

    If oDoc Is Nothing Then Exit Function

    Set oSheet = oDoc.ActiveSheet
    If oSheet Is Nothing Then Exit Function

    On Error Resume Next
    oSheet.ChangeSize kA3DrawingSheetSize, kLandscapePageOrientation
    If Err.Number <> 0 Then
        Err.Clear
        Set oSheet = oDoc.Sheets.Add(kA3DrawingSheetSize, kLandscapePageOrientation)
    End If
    On Error GoTo 0

    If oSheet Is Nothing Then
        MsgBox "Could not set/create A3 landscape sheet.", vbCritical
        Exit Function
    End If

    oSheet.Activate

    wMm = CmToMm(oDoc, oSheet.Width)
    hMm = CmToMm(oDoc, oSheet.Height)
    If Abs(wMm - A3_W_MM) > A3_TOL_MM Or Abs(hMm - A3_H_MM) > A3_TOL_MM Then
        MsgBox "Active sheet is not A3 landscape after resize/create.", vbCritical
        Exit Function
    End If

    Set EnsureA3LandscapeSheet = oSheet
End Function

Private Function EnsureSpdsA3BorderDefinition(ByVal oDoc As DrawingDocument) As BorderDefinition
    Dim oDef As BorderDefinition
    Dim oSketch As DrawingSketch
    Dim isEditing As Boolean

    On Error GoTo EH

    Set oDef = FindBorderDefinition(oDoc, RKM_BORDER_NAME)
    If oDef Is Nothing Then
        Set oDef = oDoc.BorderDefinitions.Add(RKM_BORDER_NAME)
    End If

    oDef.Edit oSketch
    isEditing = True

    ClearSketch oSketch
    DrawSpdsBorderGeometry oSketch

    oDef.ExitEdit True
    isEditing = False

    Set EnsureSpdsA3BorderDefinition = oDef
    Exit Function
EH:
    If isEditing Then
        On Error Resume Next
        oDef.ExitEdit False
        On Error GoTo 0
    End If
    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical
End Function

Private Function EnsureSpdsForm3TitleBlockDefinition(ByVal oDoc As DrawingDocument) As TitleBlockDefinition
    Dim oDef As TitleBlockDefinition
    Dim oSketch As DrawingSketch
    Dim isEditing As Boolean

    On Error GoTo EH

    Set oDef = FindTitleBlockDefinition(oDoc, RKM_TITLEBLOCK_NAME)
    If oDef Is Nothing Then
        Set oDef = oDoc.TitleBlockDefinitions.Add(RKM_TITLEBLOCK_NAME)
    End If

    oDef.Edit oSketch
    isEditing = True

    ClearSketch oSketch
    DrawForm3Geometry oSketch
    AddForm3StaticText oSketch

    oDef.ExitEdit True
    isEditing = False

    Set EnsureSpdsForm3TitleBlockDefinition = oDef
    Exit Function
EH:
    If isEditing Then
        On Error Resume Next
        oDef.ExitEdit False
        On Error GoTo 0
    End If
    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical
End Function

Private Sub ApplySpdsBorderToSheet(ByVal oSheet As Sheet, ByVal oDef As BorderDefinition)
    If oSheet Is Nothing Then Exit Sub
    If oDef Is Nothing Then Exit Sub

    On Error Resume Next
    If Not oSheet.Border Is Nothing Then oSheet.Border.Delete
    On Error GoTo 0

    oSheet.AddBorder oDef
End Sub

Private Sub ApplySpdsTitleBlockToSheet(ByVal oSheet As Sheet, ByVal oDef As TitleBlockDefinition)
    Dim prompts As Variant

    If oSheet Is Nothing Then Exit Sub
    If oDef Is Nothing Then Exit Sub

    On Error Resume Next
    If Not oSheet.TitleBlock Is Nothing Then oSheet.TitleBlock.Delete
    On Error GoTo 0

    prompts = BuildPromptStringsIfNeeded(oDef)

    If IsEmpty(prompts) Then
        oSheet.AddTitleBlock oDef
    Else
        oSheet.AddTitleBlock oDef, , prompts
    End If
End Sub

Private Sub DrawSpdsBorderGeometry(ByVal oSketch As DrawingSketch)
    Dim x0 As Double
    Dim y0 As Double
    Dim xMax As Double
    Dim yMax As Double
    Dim ix1 As Double
    Dim iy1 As Double
    Dim ix2 As Double
    Dim iy2 As Double
    Dim tbX1 As Double
    Dim tbY1 As Double
    Dim tbX2 As Double
    Dim tbY2 As Double

    x0 = 0#
    y0 = 0#
    xMax = Mm(A3_W_MM)
    yMax = Mm(A3_H_MM)

    ix1 = Mm(FRAME_LEFT_MM)
    iy1 = Mm(FRAME_OTHER_MM)
    ix2 = Mm(A3_W_MM - FRAME_OTHER_MM)
    iy2 = Mm(A3_H_MM - FRAME_OTHER_MM)

    tbX2 = ix2
    tbY1 = iy1
    tbX1 = tbX2 - Mm(TB_W_MM)
    tbY2 = tbY1 + Mm(TB_H_MM)

    Debug.Print "Inner frame (cm): (" & Fmt(ix1) & "," & Fmt(iy1) & ") - (" & Fmt(ix2) & "," & Fmt(iy2) & ")"
    Debug.Print "Title zone (cm): (" & Fmt(tbX1) & "," & Fmt(tbY1) & ") - (" & Fmt(tbX2) & "," & Fmt(tbY2) & ")"

    oSketch.SketchLines.AddAsTwoPointRectangle P2d(x0, y0), P2d(xMax, yMax)
    oSketch.SketchLines.AddAsTwoPointRectangle P2d(ix1, iy1), P2d(ix2, iy2)

    oSketch.SketchLines.AddByTwoPoints P2d(tbX1, iy1), P2d(tbX1, tbY2)
    oSketch.SketchLines.AddByTwoPoints P2d(tbX1, tbY2), P2d(ix2, tbY2)
End Sub

Private Sub DrawForm3Geometry(ByVal oSketch As DrawingSketch)
    Dim x1 As Double
    Dim y1 As Double
    Dim x2 As Double
    Dim y2 As Double
    Dim c1 As Double
    Dim c2 As Double
    Dim c3 As Double
    Dim r1 As Double
    Dim r2 As Double
    Dim r3 As Double

    x2 = Mm(A3_W_MM - FRAME_OTHER_MM)
    y1 = Mm(FRAME_OTHER_MM)
    x1 = x2 - Mm(TB_W_MM)
    y2 = y1 + Mm(TB_H_MM)

    c1 = x1 + Mm(TB_C1_MM)
    c2 = x1 + Mm(TB_C2_MM)
    c3 = x1 + Mm(TB_C3_MM)

    r1 = y1 + Mm(TB_R1_MM)
    r2 = y1 + Mm(TB_R2_MM)
    r3 = y1 + Mm(TB_R3_MM)

    Debug.Print "Title block box (cm): (" & Fmt(x1) & "," & Fmt(y1) & ") - (" & Fmt(x2) & "," & Fmt(y2) & ")"
    Debug.Print "Title block size (mm): " & Fmt(CmToMm(Nothing, x2 - x1)) & " x " & Fmt(CmToMm(Nothing, y2 - y1))

    oSketch.SketchLines.AddAsTwoPointRectangle P2d(x1, y1), P2d(x2, y2)
    oSketch.SketchLines.AddByTwoPoints P2d(c1, y1), P2d(c1, y2)
    oSketch.SketchLines.AddByTwoPoints P2d(c2, y1), P2d(c2, y2)
    oSketch.SketchLines.AddByTwoPoints P2d(c3, y1), P2d(c3, y2)

    oSketch.SketchLines.AddByTwoPoints P2d(x1, r1), P2d(x2, r1)
    oSketch.SketchLines.AddByTwoPoints P2d(x1, r2), P2d(x2, r2)
    oSketch.SketchLines.AddByTwoPoints P2d(x1, r3), P2d(x2, r3)
End Sub

Private Sub AddForm3StaticText(ByVal oSketch As DrawingSketch)
    Dim x1 As Double
    Dim y1 As Double
    Dim x2 As Double

    x2 = Mm(A3_W_MM - FRAME_OTHER_MM)
    y1 = Mm(FRAME_OTHER_MM)
    x1 = x2 - Mm(TB_W_MM)

    oSketch.TextBoxes.AddFitted P2d(x1 + Mm(2#), y1 + Mm(47#)), "Project"
    oSketch.TextBoxes.AddFitted P2d(x1 + Mm(2#), y1 + Mm(32#)), "Drawing"
    oSketch.TextBoxes.AddFitted P2d(x1 + Mm(112#), y1 + Mm(47#)), "Stage"
    oSketch.TextBoxes.AddFitted P2d(x1 + Mm(152#), y1 + Mm(47#)), "Sheet"
    oSketch.TextBoxes.AddFitted P2d(x1 + Mm(172#), y1 + Mm(47#)), "Sheets"
    oSketch.TextBoxes.AddFitted P2d(x1 + Mm(112#), y1 + Mm(2#)), "A3"
End Sub

Private Function BuildPromptStringsIfNeeded(ByVal oTitleDef As TitleBlockDefinition) As Variant
    Dim oSketch As DrawingSketch
    Dim oTextBox As TextBox
    Dim promptCount As Long
    Dim i As Long
    Dim promptValues() As String

    If oTitleDef Is Nothing Then Exit Function

    On Error Resume Next
    Set oSketch = oTitleDef.Sketch
    On Error GoTo 0
    If oSketch Is Nothing Then Exit Function

    For Each oTextBox In oSketch.TextBoxes
        If InStr(1, oTextBox.Text, "<Prompt>", vbTextCompare) > 0 Then
            promptCount = promptCount + 1
        End If
    Next oTextBox

    If promptCount = 0 Then Exit Function

    ReDim promptValues(1 To promptCount)
    For i = 1 To promptCount
        promptValues(i) = ""
    Next i

    BuildPromptStringsIfNeeded = promptValues
End Function

Public Sub Rkm_SelfTest_SPDS_A3()
    Dim oDoc As DrawingDocument
    Dim oSheet As Sheet
    Dim widthMm As Double
    Dim heightMm As Double

    Set oDoc = EnsureDrawingDocument()
    If oDoc Is Nothing Then Exit Sub

    Set oSheet = EnsureA3LandscapeSheet(oDoc)
    If oSheet Is Nothing Then Exit Sub

    widthMm = oDoc.UnitsOfMeasure.ConvertUnits(oSheet.Width, kCentimeterLengthUnits, kMillimeterLengthUnits)
    heightMm = oDoc.UnitsOfMeasure.ConvertUnits(oSheet.Height, kCentimeterLengthUnits, kMillimeterLengthUnits)

    Debug.Print "SelfTest: doc=" & oDoc.DisplayName
    Debug.Print "SelfTest: sheet(mm)=" & Fmt(widthMm) & " x " & Fmt(heightMm)

    If Abs(widthMm - A3_W_MM) <= A3_TOL_MM And Abs(heightMm - A3_H_MM) <= A3_TOL_MM Then
        Debug.Print "SelfTest: PASS"
    Else
        Debug.Print "SelfTest: FAIL"
    End If
End Sub

Private Function Mm(ByVal valueMm As Double) As Double
    Mm = valueMm * 0.1
End Function

Private Function CmToMm(ByVal oDoc As DrawingDocument, ByVal valueCm As Double) As Double
    If oDoc Is Nothing Then
        CmToMm = valueCm * 10#
    Else
        CmToMm = oDoc.UnitsOfMeasure.ConvertUnits(valueCm, kCentimeterLengthUnits, kMillimeterLengthUnits)
    End If
End Function

Private Function FindBorderDefinition(ByVal oDoc As DrawingDocument, ByVal defName As String) As BorderDefinition
    On Error Resume Next
    Set FindBorderDefinition = oDoc.BorderDefinitions.Item(defName)
    On Error GoTo 0
End Function

Private Function FindTitleBlockDefinition(ByVal oDoc As DrawingDocument, ByVal defName As String) As TitleBlockDefinition
    On Error Resume Next
    Set FindTitleBlockDefinition = oDoc.TitleBlockDefinitions.Item(defName)
    On Error GoTo 0
End Function

Private Sub ClearSketch(ByVal oSketch As DrawingSketch)
    Dim i As Long

    If oSketch Is Nothing Then Exit Sub

    For i = oSketch.TextBoxes.Count To 1 Step -1
        oSketch.TextBoxes.Item(i).Delete
    Next i

    For i = oSketch.SketchLines.Count To 1 Step -1
        oSketch.SketchLines.Item(i).Delete
    Next i
End Sub

Private Function P2d(ByVal x As Double, ByVal y As Double) As Point2d
    Set P2d = ThisApplication.TransientGeometry.CreatePoint2d(x, y)
End Function

Private Function Fmt(ByVal value As Double) As String
    Fmt = Format$(value, "0.000")
End Function
