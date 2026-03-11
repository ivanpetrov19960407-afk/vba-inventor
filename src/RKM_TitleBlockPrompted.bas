Attribute VB_Name = "RKM_TitleBlockPrompted"
Option Explicit

Public Function EnsureRkmTitleBlockDefinition(ByVal oDoc As DrawingDocument) As TitleBlockDefinition
    Dim oDef As TitleBlockDefinition
    Dim oSketch As DrawingSketch

    If oDoc Is Nothing Then Exit Function

    Set oDef = TitleBlockDefinitionByName(oDoc, RKM_TITLEBLOCK_NAME)
    If oDef Is Nothing Then
        Set oDef = oDoc.TitleBlockDefinitions.Add(RKM_TITLEBLOCK_NAME)
    End If

    oDef.Edit oSketch
    ClearSketch oSketch
    DrawTitleBlockGeometry oSketch
    AddTitleBlockLabels oSketch
    oDef.ExitEdit True

    Set EnsureRkmTitleBlockDefinition = oDef
End Function

Public Sub ApplyRkmTitleBlockToSheet(ByVal oSheet As Sheet, ByVal oDef As TitleBlockDefinition)
    On Error GoTo AddTitleBlockFailed

    If oSheet Is Nothing Then Exit Sub
    If oDef Is Nothing Then Exit Sub

    RemoveSheetTitleBlock oSheet
    oSheet.AddTitleBlock oDef
    Exit Sub

AddTitleBlockFailed:
    MsgBox "Title block insertion failed." & vbCrLf & _
           "Err.Number: " & CStr(Err.Number) & vbCrLf & _
           "Err.Description: " & Err.Description, vbExclamation
End Sub

Private Sub DrawTitleBlockGeometry(ByVal oSketch As DrawingSketch)
    Dim x1 As Double
    Dim y1 As Double
    Dim x2 As Double
    Dim y2 As Double

    x2 = MmToCm(A3_WIDTH_MM - FRAME_OTHER_MM)
    y1 = MmToCm(FRAME_OTHER_MM)
    x1 = x2 - MmToCm(TITLE_W_MM)
    y2 = y1 + MmToCm(TITLE_H_MM)

    oSketch.SketchLines.AddAsTwoPointRectangle Pt(x1, y1), Pt(x2, y2)

    oSketch.SketchLines.AddByTwoPoints Pt(x1 + MmToCm(110#), y1), Pt(x1 + MmToCm(110#), y2)
    oSketch.SketchLines.AddByTwoPoints Pt(x1 + MmToCm(150#), y1), Pt(x1 + MmToCm(150#), y2)
    oSketch.SketchLines.AddByTwoPoints Pt(x1 + MmToCm(170#), y1), Pt(x1 + MmToCm(170#), y2)

    oSketch.SketchLines.AddByTwoPoints Pt(x1, y1 + MmToCm(15#)), Pt(x2, y1 + MmToCm(15#))
    oSketch.SketchLines.AddByTwoPoints Pt(x1, y1 + MmToCm(30#)), Pt(x2, y1 + MmToCm(30#))
    oSketch.SketchLines.AddByTwoPoints Pt(x1, y1 + MmToCm(45#)), Pt(x2, y1 + MmToCm(45#))
End Sub

Private Sub AddTitleBlockLabels(ByVal oSketch As DrawingSketch)
    Dim x2 As Double
    Dim y1 As Double
    Dim x1 As Double

    x2 = MmToCm(A3_WIDTH_MM - FRAME_OTHER_MM)
    y1 = MmToCm(FRAME_OTHER_MM)
    x1 = x2 - MmToCm(TITLE_W_MM)

    oSketch.TextBoxes.AddFitted Pt(x1 + MmToCm(2#), y1 + MmToCm(47#)), RuText(1053, 1072, 1080, 1084, 1077, 1085, 1086, 1074, 1072, 1085, 1080, 1077)
    oSketch.TextBoxes.AddFitted Pt(x1 + MmToCm(2#), y1 + MmToCm(32#)), RuText(1054, 1073, 1086, 1079, 1085, 1072, 1095, 1077, 1085, 1080, 1077)
    oSketch.TextBoxes.AddFitted Pt(x1 + MmToCm(112#), y1 + MmToCm(47#)), RuText(1057, 1090, 1072, 1076, 1080, 1103)
    oSketch.TextBoxes.AddFitted Pt(x1 + MmToCm(152#), y1 + MmToCm(47#)), RuText(1051, 1080, 1089, 1090)
    oSketch.TextBoxes.AddFitted Pt(x1 + MmToCm(172#), y1 + MmToCm(47#)), RuText(1051, 1080, 1089, 1090, 1086, 1074)
    oSketch.TextBoxes.AddFitted Pt(x1 + MmToCm(112#), y1 + MmToCm(2#)), "A3"
End Sub
