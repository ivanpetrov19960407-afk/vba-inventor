Attribute VB_Name = "RKM_FrameBorder"
Option Explicit

Private Const SHEET_WIDTH_CM As Double = 42#
Private Const SHEET_HEIGHT_CM As Double = 29.7

Private Const OUTER_MARGIN_LEFT_CM As Double = 0.1
Private Const OUTER_MARGIN_BOTTOM_CM As Double = 0.1
Private Const OUTER_MARGIN_RIGHT_CM As Double = 0.1
Private Const OUTER_MARGIN_TOP_CM As Double = 0.1

Private Const INNER_LEFT_CM As Double = 2#
Private Const INNER_BOTTOM_CM As Double = 0.5
Private Const INNER_RIGHT_CM As Double = 41.5
Private Const INNER_TOP_CM As Double = 29.1

Private Const LEFT_ZONE_RIGHT_CM As Double = 1.55
Private Const STAMP_WIDTH_CM As Double = 17.8
Private Const STAMP_HEIGHT_CM As Double = 5.5

Public Function EnsureRkmBorderDefinition(ByVal oDoc As DrawingDocument) As BorderDefinition
    Dim defName As String
    Dim oDef As BorderDefinition
    Dim oSketch As DrawingSketch

    If oDoc Is Nothing Then Exit Function

    defName = SafeBorderDefinitionName(oDoc, RKM_BORDER_NAME)
    Set oDef = oDoc.BorderDefinitions.Add(defName)

    Call oDef.Edit(oSketch)
    Call DrawFrameGeometry(oSketch)
    Call DrawLeftStrip(oSketch)
    Call DrawCompactStamp(oSketch)
    Call oDef.ExitEdit(True)

    Set EnsureRkmBorderDefinition = oDef
End Function

Public Sub ApplyRkmBorderToSheet(ByVal oSheet As Sheet, ByVal oDef As BorderDefinition)
    If oSheet Is Nothing Then Exit Sub
    If oDef Is Nothing Then Exit Sub

    Call RemoveSheetBorder(oSheet)
    Call oSheet.AddBorder(oDef)
End Sub

Private Sub DrawFrameGeometry(ByVal oSketch As DrawingSketch)
    Dim outerX1 As Double
    Dim outerY1 As Double
    Dim outerX2 As Double
    Dim outerY2 As Double

    outerX1 = OUTER_MARGIN_LEFT_CM
    outerY1 = OUTER_MARGIN_BOTTOM_CM
    outerX2 = SHEET_WIDTH_CM - OUTER_MARGIN_RIGHT_CM
    outerY2 = SHEET_HEIGHT_CM - OUTER_MARGIN_TOP_CM

    Call oSketch.SketchLines.AddAsTwoPointRectangle(Pt(outerX1, outerY1), Pt(outerX2, outerY2))
    Call oSketch.SketchLines.AddAsTwoPointRectangle(Pt(INNER_LEFT_CM, INNER_BOTTOM_CM), Pt(INNER_RIGHT_CM, INNER_TOP_CM))
End Sub

Private Sub DrawLeftStrip(ByVal oSketch As DrawingSketch)
    Dim xValues As Variant
    Dim yValues As Variant
    Dim i As Long

    xValues = Array(0.55, 0.85, 1.05, 1.35, 1.55)
    yValues = Array(17#, 16#, 14#, 12#, 9#, 6.5, 3.5)

    For i = LBound(xValues) To UBound(xValues)
        Call oSketch.SketchLines.AddByTwoPoints(Pt(CDbl(xValues(i)), OUTER_MARGIN_BOTTOM_CM), Pt(CDbl(xValues(i)), SHEET_HEIGHT_CM - OUTER_MARGIN_TOP_CM))
    Next i

    For i = LBound(yValues) To UBound(yValues)
        Call oSketch.SketchLines.AddByTwoPoints(Pt(OUTER_MARGIN_LEFT_CM, CDbl(yValues(i))), Pt(LEFT_ZONE_RIGHT_CM, CDbl(yValues(i))))
    Next i

    Call oSketch.TextBoxes.AddFitted(Pt(0.2, 19#), "SEC 1")
    Call oSketch.TextBoxes.AddFitted(Pt(0.2, 11#), "SEC 2")
    Call oSketch.TextBoxes.AddFitted(Pt(0.2, 2#), "SEC 3")
End Sub

Private Sub DrawCompactStamp(ByVal oSketch As DrawingSketch)
    Dim x1 As Double
    Dim y1 As Double
    Dim x2 As Double
    Dim y2 As Double

    x2 = INNER_RIGHT_CM
    y1 = INNER_BOTTOM_CM
    x1 = x2 - STAMP_WIDTH_CM
    y2 = y1 + STAMP_HEIGHT_CM

    Call oSketch.SketchLines.AddAsTwoPointRectangle(Pt(x1, y1), Pt(x2, y2))

    Call oSketch.SketchLines.AddByTwoPoints(Pt(x1 + 11.2, y1), Pt(x1 + 11.2, y2))
    Call oSketch.SketchLines.AddByTwoPoints(Pt(x1 + 13.8, y1), Pt(x1 + 13.8, y2))
    Call oSketch.SketchLines.AddByTwoPoints(Pt(x1 + 15.2, y1), Pt(x1 + 15.2, y2))
    Call oSketch.SketchLines.AddByTwoPoints(Pt(x1 + 16.5, y1), Pt(x1 + 16.5, y2))

    Call oSketch.SketchLines.AddByTwoPoints(Pt(x1, y1 + 1.2), Pt(x2, y1 + 1.2))
    Call oSketch.SketchLines.AddByTwoPoints(Pt(x1, y1 + 2.3), Pt(x2, y1 + 2.3))
    Call oSketch.SketchLines.AddByTwoPoints(Pt(x1, y1 + 3.5), Pt(x2, y1 + 3.5))
    Call oSketch.SketchLines.AddByTwoPoints(Pt(x1, y1 + 4.4), Pt(x2, y1 + 4.4))

    Call oSketch.TextBoxes.AddFitted(Pt(x1 + 0.2, y1 + 4.65), "RKM-000")
    Call oSketch.TextBoxes.AddFitted(Pt(x1 + 0.2, y1 + 3.65), "OBJ 1")
    Call oSketch.TextBoxes.AddFitted(Pt(x1 + 0.2, y1 + 2.75), "OBJ 2")
    Call oSketch.TextBoxes.AddFitted(Pt(x1 + 0.2, y1 + 1.55), "OBJ 3")

    Call oSketch.TextBoxes.AddFitted(Pt(x1 + 11.35, y1 + 4.65), "P")
    Call oSketch.TextBoxes.AddFitted(Pt(x1 + 13.95, y1 + 4.65), "1")
    Call oSketch.TextBoxes.AddFitted(Pt(x1 + 15.35, y1 + 4.65), "A3")
    Call oSketch.TextBoxes.AddFitted(Pt(x1 + 16.6, y1 + 4.65), "ORG")

    Call oSketch.TextBoxes.AddFitted(Pt(x1 + 11.35, y1 + 3.65), "IVANOV")
    Call oSketch.TextBoxes.AddFitted(Pt(x1 + 11.35, y1 + 2.75), "01.01.2026")
End Sub
