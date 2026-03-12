Attribute VB_Name = "RKM_FrameBorder"
Option Explicit

Public Function EnsureRkmBorderDefinition(ByVal oDoc As DrawingDocument) As BorderDefinition
    Dim oDef As BorderDefinition
    Dim oSketch As DrawingSketch
    Dim isEditing As Boolean
    Dim targetName As String

    If oDoc Is Nothing Then Exit Function
    On Error GoTo EH

    RemoveSheetBorder oDoc.ActiveSheet

    ' Версия 12 (сброс сломанного кэша)
    targetName = RKM_BORDER_NAME & "_V12"

    Set oDef = BorderDefinitionByName(oDoc, targetName)
    If oDef Is Nothing Then
        On Error GoTo EH
        ThisApplication.SilentOperation = True
        Set oDef = oDoc.BorderDefinitions.Add(targetName)
        ThisApplication.SilentOperation = False
    End If

    oDef.Edit oSketch
    isEditing = True

    ClearSketch oSketch
    DrawSpdsInnerFrame oDoc, oSketch

    oDef.ExitEdit True
    isEditing = False

    Set EnsureRkmBorderDefinition = oDef
    Exit Function
EH:
    ThisApplication.SilentOperation = False
    If isEditing Then
        On Error Resume Next
        oDef.ExitEdit False
        On Error GoTo 0
    End If
    MsgBox "Border update failed: " & Err.Description, vbCritical
End Function

Public Sub ApplyRkmBorderToSheet(ByVal oSheet As Sheet, ByVal oDef As BorderDefinition)
    Dim newBorder As Border

    If oSheet Is Nothing Or oDef Is Nothing Then Exit Sub

    ThisApplication.SilentOperation = True
    RemoveSheetBorder oSheet
    Set newBorder = oSheet.AddCustomBorder(oDef)
    ThisApplication.SilentOperation = False
End Sub

Private Sub DrawSpdsInnerFrame(ByVal oDoc As DrawingDocument, ByVal oSketch As DrawingSketch)
    Dim oLines As SketchLines
    Set oLines = oSketch.SketchLines

    ' --- МИКРО-ЯКОРЯ (0.001 мм) ---
    ' Делают габарит рамки строго равным листу А3. Inventor не сможет ничего "отцентрировать" и сместить.
    oLines.AddByTwoPoints Pt(0, 0), Pt(0.0001, 0.0001)
    oLines.AddByTwoPoints Pt(MmToCm(oDoc, A3_WIDTH_MM), MmToCm(oDoc, A3_HEIGHT_MM)), _
                          Pt(MmToCm(oDoc, A3_WIDTH_MM) - 0.0001, MmToCm(oDoc, A3_HEIGHT_MM) - 0.0001)

    ' --- ВНУТРЕННЯЯ РАМКА ---
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
    x1 = MmToCm(oDoc, FRAME_LEFT_MM) ' 20 мм
    y1 = MmToCm(oDoc, FRAME_OTHER_MM) ' 5 мм
    x2 = MmToCm(oDoc, A3_WIDTH_MM - FRAME_OTHER_MM) ' 415 мм
    y2 = MmToCm(oDoc, A3_HEIGHT_MM - FRAME_OTHER_MM) ' 292 мм

    oLines.AddAsTwoPointRectangle Pt(x1, y1), Pt(x2, y2)
End Sub
