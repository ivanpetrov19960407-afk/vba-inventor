Attribute VB_Name = "RKM_TitleBlockPrompted"
Option Explicit

' Title block total width (A3 bottom-right zone), cm.
Private Const TB_W As Double = 17.8
' Title block total height (A3 bottom-right zone), cm.
Private Const TB_H As Double = 5.5
Private Const LEFT_W As Double = 6#

Public Function EnsureRkmTitleBlockDefinition(ByVal oDoc As DrawingDocument) As TitleBlockDefinition
    Dim defName As String
    Dim oDef As TitleBlockDefinition
    Dim oSketch As DrawingSketch

    If oDoc Is Nothing Then Exit Function

    defName = SafeTitleBlockDefinitionName(oDoc, RKM_TITLEBLOCK_NAME)
    Set oDef = oDoc.TitleBlockDefinitions.Add(defName)

    Call oDef.Edit(oSketch)
    Call DrawTitleBlockGrid(oSketch)
    Call AddStaticLabels(oSketch)
    Call AddPromptedFields(oSketch)
    Call oDef.ExitEdit(True)

    Set EnsureRkmTitleBlockDefinition = oDef
End Function

Public Sub ApplyRkmTitleBlockToSheet(ByVal oSheet As Sheet, ByVal oDef As TitleBlockDefinition, ByVal promptValues As Variant)
    If oSheet Is Nothing Then Exit Sub
    If oDef Is Nothing Then Exit Sub

    On Error GoTo AddTitleBlockFailed
    Call RemoveSheetTitleBlock(oSheet)
    Call oSheet.AddTitleBlock(oDef)
    Exit Sub

AddTitleBlockFailed:
    MsgBox "Title block insertion failed" & vbCrLf & _
           "Err.Number: " & CStr(Err.Number) & vbCrLf & _
           "Err.Description: " & Err.Description, vbExclamation
End Sub

Private Function NormalizePromptValues(ByVal promptValues As Variant) As Variant
    Dim values(0 To 14) As String
    Dim defaults As Variant
    Dim srcLower As Long
    Dim srcUpper As Long
    Dim i As Long

    defaults = DefaultPromptValues()
    For i = 0 To 14
        values(i) = SafePromptString(defaults(i), "")
    Next i

    If IsEmpty(promptValues) Or IsNull(promptValues) Then
        NormalizePromptValues = values
        Exit Function
    End If

    If Not IsArray(promptValues) Then
        NormalizePromptValues = values
        Exit Function
    End If

    On Error GoTo InvalidPromptArray
    srcLower = LBound(promptValues)
    srcUpper = UBound(promptValues)

    If srcLower > srcUpper Then GoTo InvalidPromptArray

    For i = 0 To 14
        If (srcLower + i) <= srcUpper Then
            values(i) = SafePromptString(promptValues(srcLower + i), values(i))
        End If
    Next i

    On Error GoTo 0
    NormalizePromptValues = values
    Exit Function

InvalidPromptArray:
    On Error GoTo 0
    NormalizePromptValues = values
End Function

Private Function SafePromptString(ByVal value As Variant, ByVal fallbackValue As String) As String
    If IsError(value) Then
        SafePromptString = fallbackValue
        Exit Function
    End If

    If IsNull(value) Or IsEmpty(value) Then
        SafePromptString = fallbackValue
        Exit Function
    End If

    On Error GoTo ConvertFailed
    SafePromptString = CStr(value)
    Exit Function

ConvertFailed:
    SafePromptString = fallbackValue
End Function

Private Sub DrawTitleBlockGrid(ByVal oSketch As DrawingSketch)
    ' Inventor controls title block placement when inserting the definition.
    ' Keep internal sketch geometry local and compact from (0,0) to (TB_W, TB_H).
    ' Outer box.
    Call oSketch.SketchLines.AddAsTwoPointRectangle(Pt(0#, 0#), Pt(TB_W, TB_H))

    ' Left mini-table.
    Call oSketch.SketchLines.AddByTwoPoints(Pt(LEFT_W, 0#), Pt(LEFT_W, TB_H))
    Call oSketch.SketchLines.AddByTwoPoints(Pt(0#, 4.7), Pt(LEFT_W, 4.7))
    Call oSketch.SketchLines.AddByTwoPoints(Pt(0#, 4.0), Pt(LEFT_W, 4.0))
    Call oSketch.SketchLines.AddByTwoPoints(Pt(0#, 3.3), Pt(LEFT_W, 3.3))
    Call oSketch.SketchLines.AddByTwoPoints(Pt(0#, 2.4), Pt(LEFT_W, 2.4))
    Call oSketch.SketchLines.AddByTwoPoints(Pt(0#, 1.2), Pt(LEFT_W, 1.2))

    ' Revision sub-columns.
    Call oSketch.SketchLines.AddByTwoPoints(Pt(1.0, 3.3), Pt(1.0, 5.5))
    Call oSketch.SketchLines.AddByTwoPoints(Pt(2.0, 3.3), Pt(2.0, 5.5))
    Call oSketch.SketchLines.AddByTwoPoints(Pt(3.0, 3.3), Pt(3.0, 5.5))
    Call oSketch.SketchLines.AddByTwoPoints(Pt(4.0, 3.3), Pt(4.0, 5.5))
    Call oSketch.SketchLines.AddByTwoPoints(Pt(5.0, 3.3), Pt(5.0, 5.5))

    ' Right content block row guides.
    Call oSketch.SketchLines.AddByTwoPoints(Pt(LEFT_W, 4.7), Pt(TB_W, 4.7))
    Call oSketch.SketchLines.AddByTwoPoints(Pt(LEFT_W, 3.6), Pt(TB_W, 3.6))
    Call oSketch.SketchLines.AddByTwoPoints(Pt(LEFT_W, 2.0), Pt(TB_W, 2.0))
    Call oSketch.SketchLines.AddByTwoPoints(Pt(LEFT_W, 1.0), Pt(TB_W, 1.0))

    ' Stage/sheet/sheets cells on right.
    Call oSketch.SketchLines.AddByTwoPoints(Pt(13.4, 1.0), Pt(13.4, 3.6))
    Call oSketch.SketchLines.AddByTwoPoints(Pt(14.8, 1.0), Pt(14.8, 3.6))
    Call oSketch.SketchLines.AddByTwoPoints(Pt(16.2, 1.0), Pt(16.2, 3.6))
End Sub

Private Sub AddStaticLabels(ByVal oSketch As DrawingSketch)
    Call oSketch.TextBoxes.AddFitted(Pt(6.2, 5.0), ":")

    Call oSketch.TextBoxes.AddFitted(Pt(13.45, 3.15), "")
    Call oSketch.TextBoxes.AddFitted(Pt(14.85, 3.15), "")
    Call oSketch.TextBoxes.AddFitted(Pt(16.25, 3.15), "")

    Call oSketch.TextBoxes.AddFitted(Pt(0.1, 5.0), ".")
    Call oSketch.TextBoxes.AddFitted(Pt(1.05, 5.0), ".")
    Call oSketch.TextBoxes.AddFitted(Pt(2.05, 5.0), "")
    Call oSketch.TextBoxes.AddFitted(Pt(3.05, 5.0), ".")
    Call oSketch.TextBoxes.AddFitted(Pt(4.05, 5.0), "")
    Call oSketch.TextBoxes.AddFitted(Pt(5.05, 5.0), "")

    Call oSketch.TextBoxes.AddFitted(Pt(0.15, 0.4), "")
    Call oSketch.TextBoxes.AddFitted(Pt(17.0, 0.2), "A3")
End Sub

Private Sub AddPromptedFields(ByVal oSketch As DrawingSketch)
    ' Prompt dependency intentionally disabled as a stabilization step.
    Call oSketch.TextBoxes.AddFitted(Pt(8.0, 5.0), " ")
    Call oSketch.TextBoxes.AddFitted(Pt(6.2, 4.0), "RKM-000")

    Call oSketch.TextBoxes.AddFitted(Pt(6.2, 3.1), " ,  1")
    Call oSketch.TextBoxes.AddFitted(Pt(6.2, 2.7), " ,  2")
    Call oSketch.TextBoxes.AddFitted(Pt(6.2, 2.3), " ,  3")

    Call oSketch.TextBoxes.AddFitted(Pt(6.2, 1.55), ",  1")
    Call oSketch.TextBoxes.AddFitted(Pt(6.2, 1.25), ",  2")
    Call oSketch.TextBoxes.AddFitted(Pt(6.2, 0.95), ",  3")

    Call oSketch.TextBoxes.AddFitted(Pt(13.55, 2.55), "")
    Call oSketch.TextBoxes.AddFitted(Pt(14.95, 2.55), "1")
    Call oSketch.TextBoxes.AddFitted(Pt(16.35, 2.55), "1")

    Call oSketch.TextBoxes.AddFitted(Pt(6.2, 0.45), " ")
    Call oSketch.TextBoxes.AddFitted(Pt(11.2, 0.45), " ")

    Call oSketch.TextBoxes.AddFitted(Pt(2.7, 0.4), ".. ")
    Call oSketch.TextBoxes.AddFitted(Pt(5.15, 0.4), "01.01.2026")
End Sub
