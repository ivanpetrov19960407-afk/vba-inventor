Attribute VB_Name = "RKM_TitleBlockPrompted"
Option Explicit

Public Const RKM_TITLEBLOCK_NAME As String = "RKM_SPDS_A3_FORM3"

Private Const TITLE_W_MM As Double = 185#
Private Const TITLE_H_MM As Double = 55#

Public Function EnsureRkmTitleBlockDefinition(ByVal oDoc As DrawingDocument) As TitleBlockDefinition
    Dim oDef As TitleBlockDefinition
    Dim oSketch As DrawingSketch
    Dim isEditing As Boolean
    Dim targetName As String

    If oDoc Is Nothing Then Exit Function
    On Error GoTo EH

    RemoveSheetTitleBlock oDoc.ActiveSheet

    ' Версия 16 - Полное удаление XML-форматирования шрифтов
    targetName = RKM_TITLEBLOCK_NAME & "_V16"

    Set oDef = TitleBlockDefinitionByName(oDoc, targetName)
    If oDef Is Nothing Then
        Set oDef = oDoc.TitleBlockDefinitions.Add(targetName)
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
    MsgBox "Title block update failed: " & Err.Description, vbCritical
End Function

Public Sub ApplyRkmTitleBlockToSheet(ByVal oSheet As Sheet, ByVal oDef As TitleBlockDefinition)
    On Error GoTo AddTitleBlockFailed

    If oSheet Is Nothing Or oDef Is Nothing Then Exit Sub
    RemoveSheetTitleBlock oSheet

    Dim sPrompts(1 To 7) As String
    sPrompts(1) = "000-2026-АР"
    sPrompts(2) = "Многоквартирный жилой дом"
    sPrompts(3) = "План на отм. 0.000"
    sPrompts(4) = "ООО 'Ромашка'"
    sPrompts(5) = "П"
    sPrompts(6) = "1"
    sPrompts(7) = "10"

    oSheet.AddTitleBlock oDef, , sPrompts
    Exit Sub

AddTitleBlockFailed:
    MsgBox "Ошибка вставки штампа: " & Err.Description, vbExclamation
End Sub

Private Sub DrawTitleBlockGeometry(ByVal oDoc As DrawingDocument, ByVal oSketch As DrawingSketch)
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
    Dim y As Double

    oSketch.SketchLines.AddByTwoPoints Pt(0, 0), Pt(-0.0001, 0.0001)

    x2 = -MmToCm(oDoc, 5)
    y1 = MmToCm(oDoc, 5)
    x1 = x2 - MmToCm(oDoc, TITLE_W_MM)
    y2 = y1 + MmToCm(oDoc, TITLE_H_MM)

    oSketch.SketchLines.AddAsTwoPointRectangle Pt(x1, y1), Pt(x2, y2)

    DrawVLineMm oDoc, oSketch, x1, y1, 7, 0, 55
    DrawVLineMm oDoc, oSketch, x1, y1, 17, 0, 55
    DrawVLineMm oDoc, oSketch, x1, y1, 27, 0, 55
    DrawVLineMm oDoc, oSketch, x1, y1, 42, 0, 55
    DrawVLineMm oDoc, oSketch, x1, y1, 57, 0, 55
    DrawVLineMm oDoc, oSketch, x1, y1, 67, 0, 55

    DrawVLineMm oDoc, oSketch, x1, y1, 137, 0, 40
    DrawVLineMm oDoc, oSketch, x1, y1, 152, 15, 40
    DrawVLineMm oDoc, oSketch, x1, y1, 167, 15, 40

    For y = 5 To 30 Step 5
        DrawHLineMm oDoc, oSketch, x1, y1, 0, 67, y
    Next y

    DrawHLineMm oDoc, oSketch, x1, y1, 0, 185, 15
    DrawHLineMm oDoc, oSketch, x1, y1, 0, 67, 35
    DrawHLineMm oDoc, oSketch, x1, y1, 137, 185, 35
    DrawHLineMm oDoc, oSketch, x1, y1, 0, 185, 40

    DrawHLineMm oDoc, oSketch, x1, y1, 0, 67, 45
    DrawHLineMm oDoc, oSketch, x1, y1, 0, 67, 50
End Sub

Private Sub AddTitleBlockLabels(ByVal oDoc As DrawingDocument, ByVal oSketch As DrawingSketch)
    Dim x2 As Double, y1 As Double, x1 As Double
    x2 = -MmToCm(oDoc, 5)
    y1 = MmToCm(oDoc, 5)
    x1 = x2 - MmToCm(oDoc, TITLE_W_MM)

    ' --- СТАТИЧЕСКИЕ НАДПИСИ (Голый текст) ---
    AddLabelBox oDoc, oSketch, x1, y1, 0, 35, 7, 40, "Изм."
    AddLabelBox oDoc, oSketch, x1, y1, 7, 35, 17, 40, "Кол.уч"
    AddLabelBox oDoc, oSketch, x1, y1, 17, 35, 27, 40, "Лист"
    AddLabelBox oDoc, oSketch, x1, y1, 27, 35, 42, 40, "№ док."
    AddLabelBox oDoc, oSketch, x1, y1, 42, 35, 57, 40, "Подп."
    AddLabelBox oDoc, oSketch, x1, y1, 57, 35, 67, 40, "Дата"

    AddLabelBox oDoc, oSketch, x1, y1, 137, 35, 152, 40, "Стадия"
    AddLabelBox oDoc, oSketch, x1, y1, 152, 35, 167, 40, "Лист"
    AddLabelBox oDoc, oSketch, x1, y1, 167, 35, 185, 40, "Листов"

    ' --- ИНТЕРАКТИВНЫЕ ПОЛЯ (Голый текст с тегом Prompt) ---
    AddPromptBox oDoc, oSketch, x1, y1, 67, 40, 185, 55, "ШИФР"
    AddPromptBox oDoc, oSketch, x1, y1, 67, 15, 137, 40, "НАИМЕНОВАНИЕ_ПРОЕКТА"
    AddPromptBox oDoc, oSketch, x1, y1, 67, 0, 137, 15, "НАИМЕНОВАНИЕ_ЧЕРТЕЖА"
    AddPromptBox oDoc, oSketch, x1, y1, 137, 0, 185, 15, "ОРГАНИЗАЦИЯ"
    AddPromptBox oDoc, oSketch, x1, y1, 137, 15, 152, 35, "СТАДИЯ"
    AddPromptBox oDoc, oSketch, x1, y1, 152, 15, 167, 35, "ЛИСТ"
    AddPromptBox oDoc, oSketch, x1, y1, 167, 15, 185, 35, "ЛИСТОВ"
End Sub

Private Sub AddLabelBox(ByVal oDoc As DrawingDocument, ByVal oSketch As DrawingSketch, ByVal x0 As Double, ByVal y0 As Double, _
                        ByVal leftMm As Double, ByVal bottomMm As Double, ByVal rightMm As Double, ByVal topMm As Double, _
                        ByVal valueText As String)
    Dim oBox As TextBox
    Set oBox = oSketch.TextBoxes.AddByRectangle(Pt(x0 + MmToCm(oDoc, leftMm), y0 + MmToCm(oDoc, bottomMm)), _
                                               Pt(x0 + MmToCm(oDoc, rightMm), y0 + MmToCm(oDoc, topMm)), valueText)
    oBox.HorizontalJustification = kAlignTextCenter
    oBox.VerticalJustification = kAlignTextMiddle
End Sub

Private Sub AddPromptBox(ByVal oDoc As DrawingDocument, ByVal oSketch As DrawingSketch, ByVal x0 As Double, ByVal y0 As Double, _
                        ByVal leftMm As Double, ByVal bottomMm As Double, ByVal rightMm As Double, ByVal topMm As Double, _
                        ByVal promptName As String)
    Dim oBox As TextBox
    Set oBox = oSketch.TextBoxes.AddByRectangle(Pt(x0 + MmToCm(oDoc, leftMm), y0 + MmToCm(oDoc, bottomMm)), _
                                               Pt(x0 + MmToCm(oDoc, rightMm), y0 + MmToCm(oDoc, topMm)), _
                                               "<Prompt>" & promptName & "</Prompt>")
    oBox.HorizontalJustification = kAlignTextCenter
    oBox.VerticalJustification = kAlignTextMiddle
End Sub

Private Sub DrawVLineMm(ByVal oDoc As DrawingDocument, ByVal oSketch As DrawingSketch, ByVal x0 As Double, ByVal y0 As Double, _
                        ByVal atMm As Double, ByVal yFromMm As Double, ByVal yToMm As Double)
    oSketch.SketchLines.AddByTwoPoints Pt(x0 + MmToCm(oDoc, atMm), y0 + MmToCm(oDoc, yFromMm)), _
                                      Pt(x0 + MmToCm(oDoc, atMm), y0 + MmToCm(oDoc, yToMm))
End Sub

Private Sub DrawHLineMm(ByVal oDoc As DrawingDocument, ByVal oSketch As DrawingSketch, ByVal x0 As Double, ByVal y0 As Double, _
                        ByVal xFromMm As Double, ByVal xToMm As Double, ByVal atMm As Double)
    oSketch.SketchLines.AddByTwoPoints Pt(x0 + MmToCm(oDoc, xFromMm), y0 + MmToCm(oDoc, atMm)), _
                                      Pt(x0 + MmToCm(oDoc, xToMm), y0 + MmToCm(oDoc, atMm))
End Sub
