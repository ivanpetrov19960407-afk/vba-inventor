Attribute VB_Name = "RKM_Utils"
Option Explicit

Public Const RKM_BORDER_NAME As String = "RKM_A3_BORDER"
Public Const RKM_TITLEBLOCK_NAME As String = "RKM_A3_TITLEBLOCK"

Public Function CanEditDrawingResources(ByVal oApp As Inventor.Application) As Boolean
    CanEditDrawingResources = False

    If oApp Is Nothing Then Exit Function

    If Not oApp.ActiveEditObject Is Nothing Then
        MsgBox "Finish active edit mode before running this macro.", vbExclamation
        Exit Function
    End If

    CanEditDrawingResources = True
End Function

Public Function GetActiveDrawingDocument(ByVal oApp As Inventor.Application) As DrawingDocument
    If oApp Is Nothing Then Exit Function
    If oApp.ActiveDocument Is Nothing Then Exit Function

    If oApp.ActiveDocument.DocumentType <> kDrawingDocumentObject Then
        MsgBox "Active document is not a drawing.", vbExclamation
        Exit Function
    End If

    Set GetActiveDrawingDocument = oApp.ActiveDocument
End Function

Public Function CreateNewA3LandscapeDrawing(ByVal oApp As Inventor.Application) As DrawingDocument
    Dim oDoc As DrawingDocument
    Dim oSheet As Sheet

    Set oDoc = oApp.Documents.Add(kDrawingDocumentObject, "", True)
    If oDoc Is Nothing Then Exit Function

    Set oSheet = oDoc.ActiveSheet
    On Error Resume Next
    Call oSheet.ChangeSize(kA3DrawingSheetSize, kLandscapePageOrientation)
    On Error GoTo 0

    Set CreateNewA3LandscapeDrawing = oDoc
End Function

Public Function GetActiveSheet(ByVal oDoc As DrawingDocument) As Sheet
    If oDoc Is Nothing Then Exit Function
    Set GetActiveSheet = oDoc.ActiveSheet
End Function

Public Function Pt(ByVal x As Double, ByVal y As Double) As Point2d
    Set Pt = ThisApplication.TransientGeometry.CreatePoint2d(x, y)
End Function

Public Function BorderDefinitionByName(ByVal oDoc As DrawingDocument, ByVal defName As String) As BorderDefinition
    On Error Resume Next
    Set BorderDefinitionByName = oDoc.BorderDefinitions.Item(defName)
    On Error GoTo 0
End Function

Public Function TitleBlockDefinitionByName(ByVal oDoc As DrawingDocument, ByVal defName As String) As TitleBlockDefinition
    On Error Resume Next
    Set TitleBlockDefinitionByName = oDoc.TitleBlockDefinitions.Item(defName)
    On Error GoTo 0
End Function

Public Function SafeBorderDefinitionName(ByVal oDoc As DrawingDocument, ByVal baseName As String) As String
    Dim i As Long
    Dim candidate As String
    Dim oDef As BorderDefinition

    candidate = baseName
    Set oDef = BorderDefinitionByName(oDoc, candidate)

    If oDef Is Nothing Then
        SafeBorderDefinitionName = candidate
        Exit Function
    End If

    If Not oDef.IsReferenced Then
        Call oDef.Delete
        SafeBorderDefinitionName = candidate
        Exit Function
    End If

    i = 2
    Do
        candidate = baseName & "_v" & CStr(i)
        Set oDef = BorderDefinitionByName(oDoc, candidate)
        If oDef Is Nothing Then
            SafeBorderDefinitionName = candidate
            Exit Do
        End If
        i = i + 1
    Loop
End Function

Public Function SafeTitleBlockDefinitionName(ByVal oDoc As DrawingDocument, ByVal baseName As String) As String
    Dim i As Long
    Dim candidate As String
    Dim oDef As TitleBlockDefinition

    candidate = baseName
    Set oDef = TitleBlockDefinitionByName(oDoc, candidate)

    If oDef Is Nothing Then
        SafeTitleBlockDefinitionName = candidate
        Exit Function
    End If

    If Not oDef.IsReferenced Then
        Call oDef.Delete
        SafeTitleBlockDefinitionName = candidate
        Exit Function
    End If

    i = 2
    Do
        candidate = baseName & "_v" & CStr(i)
        Set oDef = TitleBlockDefinitionByName(oDoc, candidate)
        If oDef Is Nothing Then
            SafeTitleBlockDefinitionName = candidate
            Exit Do
        End If
        i = i + 1
    Loop
End Function

Public Sub RemoveSheetBorder(ByVal oSheet As Sheet)
    If oSheet Is Nothing Then Exit Sub
    On Error Resume Next
    If Not oSheet.Border Is Nothing Then oSheet.Border.Delete
    On Error GoTo 0
End Sub

Public Sub RemoveSheetTitleBlock(ByVal oSheet As Sheet)
    If oSheet Is Nothing Then Exit Sub
    On Error Resume Next
    If Not oSheet.TitleBlock Is Nothing Then oSheet.TitleBlock.Delete
    On Error GoTo 0
End Sub

Public Function DefaultPromptValues() As Variant
    Dim values(0 To 14) As String

    values(0) = "ООО Заказчик"
    values(1) = "RKM-000"
    values(2) = "Наименование объекта, строка 1"
    values(3) = "Наименование объекта, строка 2"
    values(4) = "Наименование объекта, строка 3"
    values(5) = "Раздел, строка 1"
    values(6) = "Раздел, строка 2"
    values(7) = "Раздел, строка 3"
    values(8) = "П"
    values(9) = "1"
    values(10) = "1"
    values(11) = "Общий вид"
    values(12) = "Проектная организация"
    values(13) = "И.И. Иванов"
    values(14) = "01.01.2026"

    DefaultPromptValues = values
End Function
