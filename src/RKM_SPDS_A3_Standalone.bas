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

Private Const TOP_TABLE_H_MM As Double = 15#
Private Const TOP_COL_1_MM As Double = 60#
Private Const TOP_COL_2_MM As Double = 60#
Private Const TOP_COL_3_MM As Double = 25#

Public Sub Rkm_CreateOrApplyA3Frame_SPDS()
    Dim oDoc As DrawingDocument
    Dim oSheet As Sheet
    Dim oBorderDef As BorderDefinition
    Dim oTitleDef As TitleBlockDefinition

    On Error GoTo EH

    Set oDoc = GetActiveDrawingDocument(ThisApplication)
    If oDoc Is Nothing Then Exit Sub

    If Not CanEditDrawingResources(ThisApplication) Then Exit Sub

    Set oSheet = EnsureA3LandscapeSheet(oDoc)
    If oSheet Is Nothing Then Exit Sub

    If Not ValidateSpdsA3Sheet(oDoc, oSheet) Then Exit Sub

    PrintSheetDiagnostics oDoc, oSheet
    Debug.Print "Sheet.Width (mm): " & Fmt(oDoc.UnitsOfMeasure.ConvertUnits(oSheet.Width, kCentimeterLengthUnits, kMillimeterLengthUnits))
    Debug.Print "Sheet.Height (mm): " & Fmt(oDoc.UnitsOfMeasure.ConvertUnits(oSheet.Height, kCentimeterLengthUnits, kMillimeterLengthUnits))

    Set oBorderDef = EnsureSpdsA3BorderDefinition(oDoc)
    If oBorderDef Is Nothing Then Exit Sub

    Set oTitleDef = EnsureRkmTitleBlockDefinition(oDoc)
    If oTitleDef Is Nothing Then Exit Sub

    If Not CanEditDrawingResources(ThisApplication) Then Exit Sub

    ApplySpdsBorderToSheet oSheet, oBorderDef
    ApplyRkmTitleBlockToSheet oSheet, oTitleDef

    Debug.Print "Applied BorderDefinition: " & oBorderDef.Name
    Debug.Print "Applied TitleBlockDefinition: " & oTitleDef.Name

    MsgBox "SPDS A3 frame and form 3 title block applied.", vbInformation
    Exit Sub
EH:
    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

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

Private Sub ApplySpdsBorderToSheet(ByVal oSheet As Sheet, ByVal oDef As BorderDefinition)
    If oSheet Is Nothing Then Exit Sub
    If oDef Is Nothing Then Exit Sub

    On Error Resume Next
    If Not oSheet.Border Is Nothing Then oSheet.Border.Delete
    On Error GoTo 0

    oSheet.AddBorder oDef
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

    Dim topY1 As Double
    Dim topX1 As Double
    Dim topX2 As Double
    Dim topX3 As Double

    oSketch.SketchLines.AddAsTwoPointRectangle P2d(x0, y0), P2d(xMax, yMax)
    oSketch.SketchLines.AddAsTwoPointRectangle P2d(ix1, iy1), P2d(ix2, iy2)

    ' Bottom-right title-block placement zone outline (grid belongs to TitleBlockDefinition only).
    oSketch.SketchLines.AddByTwoPoints P2d(tbX1, iy1), P2d(tbX1, tbY2)
    oSketch.SketchLines.AddByTwoPoints P2d(tbX1, tbY2), P2d(ix2, tbY2)

    ' Top service table (SPDS upper fields).
    topY1 = iy2 - Mm(TOP_TABLE_H_MM)
    topX1 = ix2 - Mm(TOP_COL_3_MM)
    topX2 = topX1 - Mm(TOP_COL_2_MM)
    topX3 = topX2 - Mm(TOP_COL_1_MM)

    oSketch.SketchLines.AddByTwoPoints P2d(ix1, topY1), P2d(ix2, topY1)
    oSketch.SketchLines.AddByTwoPoints P2d(topX1, topY1), P2d(topX1, iy2)
    oSketch.SketchLines.AddByTwoPoints P2d(topX2, topY1), P2d(topX2, iy2)
    oSketch.SketchLines.AddByTwoPoints P2d(topX3, topY1), P2d(topX3, iy2)
End Sub

Public Sub Rkm_SelfTest_SPDS_A3()
    Dim oDoc As DrawingDocument
    Dim oSheet As Sheet
    Dim widthMm As Double
    Dim heightMm As Double

    Set oDoc = GetActiveDrawingDocument(ThisApplication)
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

Private Function FindBorderDefinition(ByVal oDoc As DrawingDocument, ByVal defName As String) As BorderDefinition
    On Error Resume Next
    Set FindBorderDefinition = oDoc.BorderDefinitions.Item(defName)
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
