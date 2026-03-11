Attribute VB_Name = "RKM_FrameBorder"
Option Explicit

Public Function EnsureRkmBorderDefinition(ByVal oDoc As DrawingDocument) As BorderDefinition
    Dim oDef As BorderDefinition
    Dim oSketch As DrawingSketch
    Dim isEditing As Boolean

    If oDoc Is Nothing Then Exit Function

    On Error GoTo EH

    Set oDef = BorderDefinitionByName(oDoc, RKM_BORDER_NAME)
    If oDef Is Nothing Then
        Set oDef = oDoc.BorderDefinitions.Add(RKM_BORDER_NAME)
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
    If isEditing Then
        On Error Resume Next
        oDef.ExitEdit False
        On Error GoTo 0
    End If

    MsgBox "Border definition update failed." & vbCrLf & _
           "Err.Number: " & CStr(Err.Number) & vbCrLf & _
           "Err.Description: " & Err.Description, vbCritical
End Function

Public Sub ApplyRkmBorderToSheet(ByVal oSheet As Sheet, ByVal oDef As BorderDefinition)
    If oSheet Is Nothing Then Exit Sub
    If oDef Is Nothing Then Exit Sub

    RemoveSheetBorder oSheet
    oSheet.AddBorder oDef
End Sub

Private Sub DrawSpdsInnerFrame(ByVal oDoc As DrawingDocument, ByVal oSketch As DrawingSketch)
    Dim x1 As Double
    Dim y1 As Double
    Dim x2 As Double
    Dim y2 As Double

    x1 = MmToCm(oDoc, FRAME_LEFT_MM)
    y1 = MmToCm(oDoc, FRAME_OTHER_MM)
    x2 = MmToCm(oDoc, A3_WIDTH_MM - FRAME_OTHER_MM)
    y2 = MmToCm(oDoc, A3_HEIGHT_MM - FRAME_OTHER_MM)

    oSketch.SketchLines.AddAsTwoPointRectangle Pt(x1, y1), Pt(x2, y2)
End Sub
