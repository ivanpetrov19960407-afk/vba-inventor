Attribute VB_Name = "RKM_TitleBlockPrompted"
Option Explicit

' Title block total width (A3 bottom-right zone), cm.
Private Const TB_W As Double = 17.8
' Title block total height (A3 bottom-right zone), cm.
Private Const TB_H As Double = 5.5

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
    ' Compact GOST Form 3 stamp only.
    Call oSketch.SketchLines.AddAsTwoPointRectangle(Pt(0#, 0#), Pt(TB_W, TB_H))

    Call oSketch.SketchLines.AddByTwoPoints(Pt(0#, 4.2), Pt(TB_W, 4.2))
    Call oSketch.SketchLines.AddByTwoPoints(Pt(0#, 3.0), Pt(TB_W, 3.0))
    Call oSketch.SketchLines.AddByTwoPoints(Pt(0#, 2.0), Pt(TB_W, 2.0))
    Call oSketch.SketchLines.AddByTwoPoints(Pt(0#, 1.0), Pt(TB_W, 1.0))

    Call oSketch.SketchLines.AddByTwoPoints(Pt(12.4, 0#), Pt(12.4, TB_H))
    Call oSketch.SketchLines.AddByTwoPoints(Pt(14.2, 0#), Pt(14.2, 3.0))
    Call oSketch.SketchLines.AddByTwoPoints(Pt(16.0, 0#), Pt(16.0, 3.0))
End Sub

Private Sub AddStaticLabels(ByVal oSketch As DrawingSketch)
    Call oSketch.TextBoxes.AddFitted(Pt(14.35, 2.35), "P")
    Call oSketch.TextBoxes.AddFitted(Pt(16.15, 2.35), "1")
    Call oSketch.TextBoxes.AddFitted(Pt(16.15, 0.25), "A3")
End Sub

Private Sub AddPromptedFields(ByVal oSketch As DrawingSketch)
    Call oSketch.TextBoxes.AddFitted(Pt(0.3, 4.45), "RKM-000")
    Call oSketch.TextBoxes.AddFitted(Pt(0.3, 3.25), "OBJ 1")
    Call oSketch.TextBoxes.AddFitted(Pt(0.3, 2.25), "OBJ 2")
    Call oSketch.TextBoxes.AddFitted(Pt(0.3, 1.25), "OBJ 3")

    Call oSketch.TextBoxes.AddFitted(Pt(12.6, 4.45), "SEC 1")
    Call oSketch.TextBoxes.AddFitted(Pt(12.6, 3.25), "SEC 2")
    Call oSketch.TextBoxes.AddFitted(Pt(12.6, 2.25), "SEC 3")

    Call oSketch.TextBoxes.AddFitted(Pt(0.3, 0.25), "ORG")
    Call oSketch.TextBoxes.AddFitted(Pt(6.0, 0.25), "IVANOV")
    Call oSketch.TextBoxes.AddFitted(Pt(9.6, 0.25), "01.01.2026")
End Sub
