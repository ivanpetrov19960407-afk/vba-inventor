Attribute VB_Name = "RKM_TitleBlockPrompted"
Option Explicit

Private Const PROMPT_DOC_NAME As String = "DOC_NAME"
Private Const PROMPT_OBJ_NAME As String = "OBJ_NAME"
Private Const PROMPT_STAGE As String = "STAGE"
Private Const PROMPT_SHEET As String = "SHEET"
Private Const PROMPT_SHEETS As String = "SHEETS"

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
    AddPromptedFields oDoc, oSketch

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
    Dim prompts As Variant

    On Error GoTo AddTitleBlockFailed

    If oSheet Is Nothing Then Exit Sub
    If oDef Is Nothing Then Exit Sub

    prompts = BuildPromptStringsIfNeeded(oDef)

    RemoveSheetTitleBlock oSheet
    If IsEmpty(prompts) Then
        oSheet.AddTitleBlock oDef
    Else
        oSheet.AddTitleBlock oDef, , , prompts
    End If

    Debug.Print "Applied title block: " & oDef.Name
    Exit Sub

AddTitleBlockFailed:
    MsgBox "Title block insertion failed." & vbCrLf & _
           "Err.Number: " & CStr(Err.Number) & vbCrLf & _
           "Err.Description: " & Err.Description, vbExclamation
End Sub

Private Function BuildPromptStringsIfNeeded(ByVal oDef As TitleBlockDefinition) As Variant
    Dim oSketch As DrawingSketch
    Dim oTextBox As TextBox
    Dim promptNames() As String
    Dim promptValues() As String
    Dim promptCount As Long
    Dim promptName As String
    Dim i As Long

    If oDef Is Nothing Then Exit Function

    On Error Resume Next
    Set oSketch = oDef.Sketch
    On Error GoTo 0
    If oSketch Is Nothing Then Exit Function

    For Each oTextBox In oSketch.TextBoxes
        promptName = ExtractPromptName(oTextBox.Text)
        If Len(promptName) > 0 Then
            promptCount = promptCount + 1
            ReDim Preserve promptNames(1 To promptCount)
            promptNames(promptCount) = promptName
        End If
    Next oTextBox

    If promptCount = 0 Then Exit Function

    ReDim promptValues(1 To promptCount)
    For i = 1 To promptCount
        promptValues(i) = GetPromptDefaultValue(promptNames(i))
    Next i

    Debug.Print "Prompted fields detected: " & CStr(promptCount)
    BuildPromptStringsIfNeeded = promptValues
End Function

Private Function GetPromptDefaultValue(ByVal promptName As String) As String
    Select Case UCase$(promptName)
        Case PROMPT_DOC_NAME, PROMPT_OBJ_NAME, PROMPT_STAGE, PROMPT_SHEET, PROMPT_SHEETS
            GetPromptDefaultValue = ""
        Case Else
            GetPromptDefaultValue = ""
    End Select
End Function

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

    oSketch.SketchLines.AddByTwoPoints Pt(x1 + MmToCm(oDoc, 110#), y1), Pt(x1 + MmToCm(oDoc, 110#), y2)
    oSketch.SketchLines.AddByTwoPoints Pt(x1 + MmToCm(oDoc, 150#), y1), Pt(x1 + MmToCm(oDoc, 150#), y2)
    oSketch.SketchLines.AddByTwoPoints Pt(x1 + MmToCm(oDoc, 170#), y1), Pt(x1 + MmToCm(oDoc, 170#), y2)

    oSketch.SketchLines.AddByTwoPoints Pt(x1, y1 + MmToCm(oDoc, 15#)), Pt(x2, y1 + MmToCm(oDoc, 15#))
    oSketch.SketchLines.AddByTwoPoints Pt(x1, y1 + MmToCm(oDoc, 30#)), Pt(x2, y1 + MmToCm(oDoc, 30#))
    oSketch.SketchLines.AddByTwoPoints Pt(x1, y1 + MmToCm(oDoc, 45#)), Pt(x2, y1 + MmToCm(oDoc, 45#))
End Sub

Private Sub AddTitleBlockLabels(ByVal oDoc As DrawingDocument, ByVal oSketch As DrawingSketch)
    Dim x2 As Double
    Dim y1 As Double
    Dim x1 As Double

    x2 = MmToCm(oDoc, A3_WIDTH_MM - FRAME_OTHER_MM)
    y1 = MmToCm(oDoc, FRAME_OTHER_MM)
    x1 = x2 - MmToCm(oDoc, TITLE_W_MM)

    oSketch.TextBoxes.AddFitted Pt(x1 + MmToCm(oDoc, 2#), y1 + MmToCm(oDoc, 47#)), RuText(1053, 1072, 1080, 1084, 1077, 1085, 1086, 1074, 1072, 1085, 1080, 1077)
    oSketch.TextBoxes.AddFitted Pt(x1 + MmToCm(oDoc, 2#), y1 + MmToCm(oDoc, 32#)), RuText(1054, 1073, 1086, 1079, 1085, 1072, 1095, 1077, 1085, 1080, 1077)
    oSketch.TextBoxes.AddFitted Pt(x1 + MmToCm(oDoc, 112#), y1 + MmToCm(oDoc, 47#)), RuText(1057, 1090, 1072, 1076, 1080, 1103)
    oSketch.TextBoxes.AddFitted Pt(x1 + MmToCm(oDoc, 152#), y1 + MmToCm(oDoc, 47#)), RuText(1051, 1080, 1089, 1090)
    oSketch.TextBoxes.AddFitted Pt(x1 + MmToCm(oDoc, 172#), y1 + MmToCm(oDoc, 47#)), RuText(1051, 1080, 1089, 1090, 1086, 1074)
    oSketch.TextBoxes.AddFitted Pt(x1 + MmToCm(oDoc, 112#), y1 + MmToCm(oDoc, 2#)), "A3"
End Sub

Private Sub AddPromptedFields(ByVal oDoc As DrawingDocument, ByVal oSketch As DrawingSketch)
    Dim x2 As Double
    Dim y1 As Double
    Dim x1 As Double

    x2 = MmToCm(oDoc, A3_WIDTH_MM - FRAME_OTHER_MM)
    y1 = MmToCm(oDoc, FRAME_OTHER_MM)
    x1 = x2 - MmToCm(oDoc, TITLE_W_MM)

    oSketch.TextBoxes.AddByRectangle Pt(x1 + MmToCm(oDoc, 2#), y1 + MmToCm(oDoc, 31#)), Pt(x1 + MmToCm(oDoc, 108#), y1 + MmToCm(oDoc, 44#)), PromptToken(PROMPT_DOC_NAME)
    oSketch.TextBoxes.AddByRectangle Pt(x1 + MmToCm(oDoc, 2#), y1 + MmToCm(oDoc, 16#)), Pt(x1 + MmToCm(oDoc, 108#), y1 + MmToCm(oDoc, 29#)), PromptToken(PROMPT_OBJ_NAME)
    oSketch.TextBoxes.AddByRectangle Pt(x1 + MmToCm(oDoc, 112#), y1 + MmToCm(oDoc, 31#)), Pt(x1 + MmToCm(oDoc, 148#), y1 + MmToCm(oDoc, 44#)), PromptToken(PROMPT_STAGE)
    oSketch.TextBoxes.AddByRectangle Pt(x1 + MmToCm(oDoc, 152#), y1 + MmToCm(oDoc, 31#)), Pt(x1 + MmToCm(oDoc, 168#), y1 + MmToCm(oDoc, 44#)), PromptToken(PROMPT_SHEET)
    oSketch.TextBoxes.AddByRectangle Pt(x1 + MmToCm(oDoc, 172#), y1 + MmToCm(oDoc, 31#)), Pt(x1 + MmToCm(oDoc, 183#), y1 + MmToCm(oDoc, 44#)), PromptToken(PROMPT_SHEETS)
End Sub

Private Function PromptToken(ByVal name As String) As String
    PromptToken = "<Prompt>" & name & "</Prompt>"
End Function

Private Function ExtractPromptName(ByVal textValue As String) As String
    Const OPEN_TAG As String = "<Prompt>"
    Const CLOSE_TAG As String = "</Prompt>"

    Dim p1 As Long
    Dim p2 As Long

    p1 = InStr(1, textValue, OPEN_TAG, vbTextCompare)
    If p1 = 0 Then Exit Function

    p1 = p1 + Len(OPEN_TAG)
    p2 = InStr(p1, textValue, CLOSE_TAG, vbTextCompare)
    If p2 <= p1 Then Exit Function

    ExtractPromptName = Mid$(textValue, p1, p2 - p1)
End Function
