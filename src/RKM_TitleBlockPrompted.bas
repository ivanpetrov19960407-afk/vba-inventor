Attribute VB_Name = "RKM_TitleBlockPrompted"
Option Explicit

Private Const TB_X_DOC_RIGHT_MM As Double = 110#
Private Const TB_X_DESIGN_RIGHT_MM As Double = 140#
Private Const TB_X_STAGE_RIGHT_MM As Double = 162#
Private Const TB_X_SHEET_RIGHT_MM As Double = 174#
Private Const TB_X_SHEETS_RIGHT_MM As Double = 185#

Private Const TB_Y_ROW1_MM As Double = 10#
Private Const TB_Y_ROW2_MM As Double = 25#
Private Const TB_Y_ROW3_MM As Double = 40#
Private Const TB_Y_TOP_MM As Double = 55#

Public Function EnsureRkmTitleBlockDefinition(ByVal oDoc As DrawingDocument) As TitleBlockDefinition
    Dim oDef As TitleBlockDefinition
    Dim oSketch As DrawingSketch
    Dim isEditing As Boolean

    If oDoc Is Nothing Then Exit Function

    On Error GoTo EH

    Set oDef = TitleBlockDefinitionByName(oDoc, RKM_TITLEBLOCK_NAME)
    If oDef Is Nothing Then
        Set oDef = oDoc.TitleBlockDefinitions.Add(RKM_TITLEBLOCK_NAME)
    End If

    oDef.Edit oSketch
    isEditing = True

    ClearSketch oSketch
    DrawTitleBlockGeometry oDoc, oSketch
    AddTitleBlockLabels oDoc, oSketch

    oDef.ExitEdit True
    isEditing = False

    Set EnsureRkmTitleBlockDefinition = oDef
    Exit Function

EH:
    If isEditing Then
        On Error Resume Next
        oDef.ExitEdit False
        On Error GoTo 0
    End If

    MsgBox "Title block definition update failed." & vbCrLf & _
           "Err.Number: " & CStr(Err.Number) & vbCrLf & _
           "Err.Description: " & Err.Description, vbCritical
End Function

Public Sub ApplyRkmTitleBlockToSheet(ByVal oSheet As Sheet, ByVal oDef As TitleBlockDefinition)
    On Error GoTo AddTitleBlockFailed

    If oSheet Is Nothing Then Exit Sub
    If oDef Is Nothing Then Exit Sub

    RemoveSheetTitleBlock oSheet
    oSheet.AddTitleBlock oDef

    Debug.Print "Applied title block: " & oDef.Name
    Exit Sub

AddTitleBlockFailed:
    MsgBox "Title block insertion failed." & vbCrLf & _
           "Err.Number: " & CStr(Err.Number) & vbCrLf & _
           "Err.Description: " & Err.Description, vbExclamation
End Sub

Private Sub DrawTitleBlockGeometry(ByVal oDoc As DrawingDocument, ByVal oSketch As DrawingSketch)
    Dim x1 As Double
    Dim y1 As Double
    Dim x2 As Double
    Dim y2 As Double

    x2 = MmToCm(oDoc, A3_WIDTH_MM - FRAME_OTHER_MM)
    y1 = MmToCm(oDoc, FRAME_OTHER_MM)
    x1 = x2 - MmToCm(oDoc, TITLE_W_MM)
    y2 = y1 + MmToCm(oDoc, TITLE_H_MM)

    oSketch.SketchLines.AddAsTwoPointRectangle Pt(x1, y1), Pt(x2, y2)

    DrawVLineMm oDoc, oSketch, x1, y1, TB_X_DOC_RIGHT_MM, 0#, TB_Y_TOP_MM
    DrawVLineMm oDoc, oSketch, x1, y1, TB_X_DESIGN_RIGHT_MM, 0#, TB_Y_TOP_MM
    DrawVLineMm oDoc, oSketch, x1, y1, TB_X_STAGE_RIGHT_MM, 0#, TB_Y_TOP_MM
    DrawVLineMm oDoc, oSketch, x1, y1, TB_X_SHEET_RIGHT_MM, 0#, TB_Y_TOP_MM

    DrawHLineMm oDoc, oSketch, x1, y1, 0#, TB_X_SHEETS_RIGHT_MM, TB_Y_ROW1_MM
    DrawHLineMm oDoc, oSketch, x1, y1, 0#, TB_X_SHEETS_RIGHT_MM, TB_Y_ROW2_MM
    DrawHLineMm oDoc, oSketch, x1, y1, 0#, TB_X_SHEETS_RIGHT_MM, TB_Y_ROW3_MM
End Sub

Private Sub AddTitleBlockLabels(ByVal oDoc As DrawingDocument, ByVal oSketch As DrawingSketch)
    Dim x2 As Double
    Dim y1 As Double
    Dim x1 As Double

    x2 = MmToCm(oDoc, A3_WIDTH_MM - FRAME_OTHER_MM)
    y1 = MmToCm(oDoc, FRAME_OTHER_MM)
    x1 = x2 - MmToCm(oDoc, TITLE_W_MM)

    AddLabelBox oDoc, oSketch, x1, y1, 1#, TB_Y_ROW3_MM + 0.8, TB_X_DOC_RIGHT_MM - 1#, TB_Y_TOP_MM - 0.8, RuText(1053, 1072, 1080, 1084, 1077, 1085, 1086, 1074, 1072, 1085, 1080, 1077)
    AddLabelBox oDoc, oSketch, x1, y1, 1#, TB_Y_ROW2_MM + 0.8, TB_X_DOC_RIGHT_MM - 1#, TB_Y_ROW3_MM - 0.8, RuText(1054, 1073, 1086, 1079, 1085, 1072, 1095, 1077, 1085, 1080, 1077)

    AddLabelBox oDoc, oSketch, x1, y1, TB_X_DOC_RIGHT_MM + 0.8, TB_Y_ROW3_MM + 0.8, TB_X_DESIGN_RIGHT_MM - 0.8, TB_Y_TOP_MM - 0.8, RuText(1057, 1090, 1072, 1076, 1080, 1103)
    AddLabelBox oDoc, oSketch, x1, y1, TB_X_DESIGN_RIGHT_MM + 0.5, TB_Y_ROW3_MM + 0.8, TB_X_STAGE_RIGHT_MM - 0.5, TB_Y_TOP_MM - 0.8, RuText(1051, 1080, 1089, 1090)
    AddLabelBox oDoc, oSketch, x1, y1, TB_X_STAGE_RIGHT_MM + 0.5, TB_Y_ROW3_MM + 0.8, TB_X_SHEET_RIGHT_MM - 0.5, TB_Y_TOP_MM - 0.8, RuText(1051, 1080, 1089, 1090)
    AddLabelBox oDoc, oSketch, x1, y1, TB_X_SHEET_RIGHT_MM + 0.5, TB_Y_ROW3_MM + 0.8, TB_X_SHEETS_RIGHT_MM - 0.5, TB_Y_TOP_MM - 0.8, RuText(1051, 1080, 1089, 1090, 1086, 1074)

    AddLabelBox oDoc, oSketch, x1, y1, TB_X_DOC_RIGHT_MM + 1#, 1#, TB_X_DESIGN_RIGHT_MM - 1#, TB_Y_ROW1_MM - 1#, "A3"
End Sub

Private Sub AddLabelBox(ByVal oDoc As DrawingDocument, ByVal oSketch As DrawingSketch, ByVal x0 As Double, ByVal y0 As Double, _
                        ByVal leftMm As Double, ByVal bottomMm As Double, ByVal rightMm As Double, ByVal topMm As Double, _
                        ByVal valueText As String)
    oSketch.TextBoxes.AddByRectangle _
        Pt(x0 + MmToCm(oDoc, leftMm), y0 + MmToCm(oDoc, bottomMm)), _
        Pt(x0 + MmToCm(oDoc, rightMm), y0 + MmToCm(oDoc, topMm)), _
        valueText
End Sub

Private Sub DrawVLineMm(ByVal oDoc As DrawingDocument, ByVal oSketch As DrawingSketch, ByVal x0 As Double, ByVal y0 As Double, _
                        ByVal atMm As Double, ByVal yFromMm As Double, ByVal yToMm As Double)
    oSketch.SketchLines.AddByTwoPoints _
        Pt(x0 + MmToCm(oDoc, atMm), y0 + MmToCm(oDoc, yFromMm)), _
        Pt(x0 + MmToCm(oDoc, atMm), y0 + MmToCm(oDoc, yToMm))
End Sub

Private Sub DrawHLineMm(ByVal oDoc As DrawingDocument, ByVal oSketch As DrawingSketch, ByVal x0 As Double, ByVal y0 As Double, _
                        ByVal xFromMm As Double, ByVal xToMm As Double, ByVal atMm As Double)
    oSketch.SketchLines.AddByTwoPoints _
        Pt(x0 + MmToCm(oDoc, xFromMm), y0 + MmToCm(oDoc, atMm)), _
        Pt(x0 + MmToCm(oDoc, xToMm), y0 + MmToCm(oDoc, atMm))
End Sub
