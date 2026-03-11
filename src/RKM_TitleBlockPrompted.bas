Attribute VB_Name = "RKM_TitleBlockPrompted"
Option Explicit

' Local title block coordinates, lower-left origin.
Private Const TB_W As Double = 18.5
Private Const TB_H As Double = 5.5
Private Const LEFT_W As Double = 6.5

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

    Call RemoveSheetTitleBlock(oSheet)
    Call oSheet.AddTitleBlock(oDef, promptValues)
End Sub

Private Sub DrawTitleBlockGrid(ByVal oSketch As DrawingSketch)
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
    Call oSketch.SketchLines.AddByTwoPoints(Pt(4.1, 3.3), Pt(4.1, 5.5))
    Call oSketch.SketchLines.AddByTwoPoints(Pt(5.2, 3.3), Pt(5.2, 5.5))

    ' Right content block row guides.
    Call oSketch.SketchLines.AddByTwoPoints(Pt(LEFT_W, 4.7), Pt(TB_W, 4.7))
    Call oSketch.SketchLines.AddByTwoPoints(Pt(LEFT_W, 3.6), Pt(TB_W, 3.6))
    Call oSketch.SketchLines.AddByTwoPoints(Pt(LEFT_W, 2.0), Pt(TB_W, 2.0))
    Call oSketch.SketchLines.AddByTwoPoints(Pt(LEFT_W, 1.0), Pt(TB_W, 1.0))

    ' Stage/sheet/sheets cells on right.
    Call oSketch.SketchLines.AddByTwoPoints(Pt(14.0, 1.0), Pt(14.0, 3.6))
    Call oSketch.SketchLines.AddByTwoPoints(Pt(15.5, 1.0), Pt(15.5, 3.6))
    Call oSketch.SketchLines.AddByTwoPoints(Pt(17.0, 1.0), Pt(17.0, 3.6))
End Sub

Private Sub AddStaticLabels(ByVal oSketch As DrawingSketch)
    Call oSketch.TextBoxes.AddFitted(Pt(6.7, 5.0), "Заказчик:")

    Call oSketch.TextBoxes.AddFitted(Pt(14.1, 3.15), "Стадия")
    Call oSketch.TextBoxes.AddFitted(Pt(15.6, 3.15), "Лист")
    Call oSketch.TextBoxes.AddFitted(Pt(17.05, 3.15), "Листов")

    Call oSketch.TextBoxes.AddFitted(Pt(0.1, 5.0), "Изм.")
    Call oSketch.TextBoxes.AddFitted(Pt(1.05, 5.0), "Кол.уч")
    Call oSketch.TextBoxes.AddFitted(Pt(2.05, 5.0), "Лист")
    Call oSketch.TextBoxes.AddFitted(Pt(3.05, 5.0), "№док.")
    Call oSketch.TextBoxes.AddFitted(Pt(4.15, 5.0), "Подпись")
    Call oSketch.TextBoxes.AddFitted(Pt(5.25, 5.0), "Дата")

    Call oSketch.TextBoxes.AddFitted(Pt(0.15, 0.4), "Разработал")
    Call oSketch.TextBoxes.AddFitted(Pt(17.7, 0.2), "A3")
End Sub

Private Sub AddPromptedFields(ByVal oSketch As DrawingSketch)
    ' Prompt order must match docs/prompt-fields.md and DefaultPromptValues().
    Call oSketch.TextBoxes.AddFitted(Pt(8.6, 5.0), "<Prompt>Заказчик</Prompt>")                 ' 1 Customer
    Call oSketch.TextBoxes.AddFitted(Pt(6.7, 4.0), "<Prompt>Обозначение</Prompt>")             ' 2 Designation

    Call oSketch.TextBoxes.AddFitted(Pt(6.7, 3.1), "<Prompt>Описание объекта 1</Prompt>")      ' 3 ObjectDescription1
    Call oSketch.TextBoxes.AddFitted(Pt(6.7, 2.7), "<Prompt>Описание объекта 2</Prompt>")      ' 4 ObjectDescription2
    Call oSketch.TextBoxes.AddFitted(Pt(6.7, 2.3), "<Prompt>Описание объекта 3</Prompt>")      ' 5 ObjectDescription3

    Call oSketch.TextBoxes.AddFitted(Pt(6.7, 1.55), "<Prompt>Заголовок раздела 1</Prompt>")    ' 6 SectionTitle1
    Call oSketch.TextBoxes.AddFitted(Pt(6.7, 1.25), "<Prompt>Заголовок раздела 2</Prompt>")    ' 7 SectionTitle2
    Call oSketch.TextBoxes.AddFitted(Pt(6.7, 0.95), "<Prompt>Заголовок раздела 3</Prompt>")    ' 8 SectionTitle3

    Call oSketch.TextBoxes.AddFitted(Pt(14.2, 2.55), "<Prompt>Стадия</Prompt>")                 ' 9 Stage
    Call oSketch.TextBoxes.AddFitted(Pt(15.6, 2.55), "<Prompt>Лист</Prompt>")                   ' 10 SheetNumber
    Call oSketch.TextBoxes.AddFitted(Pt(17.05, 2.55), "<Prompt>Листов</Prompt>")                ' 11 TotalSheets

    Call oSketch.TextBoxes.AddFitted(Pt(6.7, 0.45), "<Prompt>Наименование листа</Prompt>")     ' 12 SheetName
    Call oSketch.TextBoxes.AddFitted(Pt(11.8, 0.45), "<Prompt>Организация</Prompt>")            ' 13 Organization

    Call oSketch.TextBoxes.AddFitted(Pt(2.7, 0.4), "<Prompt>Разработал ФИО</Prompt>")          ' 14 DeveloperName
    Call oSketch.TextBoxes.AddFitted(Pt(5.35, 0.4), "<Prompt>Дата разработал</Prompt>")        ' 15 DeveloperDate
End Sub
