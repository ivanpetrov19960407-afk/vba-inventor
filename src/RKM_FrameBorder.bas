Attribute VB_Name = "RKM_FrameBorder"
Option Explicit

Private Const OUTER_X1 As Double = 0.1
Private Const OUTER_Y1 As Double = 0.1
Private Const OUTER_X2 As Double = 41.9
Private Const OUTER_Y2 As Double = 29.6

Private Const MAIN_X1 As Double = 2#
Private Const MAIN_Y1 As Double = 0.5
Private Const MAIN_X2 As Double = 41.5
Private Const MAIN_Y2 As Double = 29.1

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
    Call oSketch.SketchLines.AddAsTwoPointRectangle(Pt(OUTER_X1, OUTER_Y1), Pt(OUTER_X2, OUTER_Y2))
    Call oSketch.SketchLines.AddAsTwoPointRectangle(Pt(MAIN_X1, MAIN_Y1), Pt(MAIN_X2, MAIN_Y2))
End Sub

Private Sub DrawLeftStrip(ByVal oSketch As DrawingSketch)
    Dim xValues As Variant
    Dim yValues As Variant
    Dim i As Long

    xValues = Array(0.55, 0.85, 1.05, 1.35, 1.55)
    yValues = Array(17#, 16#, 14#, 12#, 9#, 6.5, 3.5)

    For i = LBound(xValues) To UBound(xValues)
        Call oSketch.SketchLines.AddByTwoPoints(Pt(CDbl(xValues(i)), OUTER_Y1), Pt(CDbl(xValues(i)), OUTER_Y2))
    Next i

    For i = LBound(yValues) To UBound(yValues)
        Call oSketch.SketchLines.AddByTwoPoints(Pt(OUTER_X1, CDbl(yValues(i))), Pt(1.55, CDbl(yValues(i))))
    Next i

    Call oSketch.TextBoxes.AddFitted(Pt(0.2, 19#), "Согласовано:")
    Call oSketch.TextBoxes.AddFitted(Pt(0.2, 11#), "Взамен инв. №")
    Call oSketch.TextBoxes.AddFitted(Pt(0.2, 8#), "Подпись и дата")
    Call oSketch.TextBoxes.AddFitted(Pt(0.2, 2#), "Инв. № подл.")
End Sub
