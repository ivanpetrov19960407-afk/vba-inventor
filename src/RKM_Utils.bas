Attribute VB_Name = "RKM_Utils"
Option Explicit

Public Const RKM_BORDER_NAME As String = "RKM_SPDS_A3_BORDER"
Public Const RKM_TITLEBLOCK_NAME As String = "RKM_SPDS_A3_FORM3_TITLEBLOCK"

Public Const MM_TO_CM As Double = 0.1
Public Const A3_WIDTH_MM As Double = 420#
Public Const A3_HEIGHT_MM As Double = 297#
Public Const FRAME_LEFT_MM As Double = 20#
Public Const FRAME_OTHER_MM As Double = 5#
Public Const TITLE_W_MM As Double = 185#
Public Const TITLE_H_MM As Double = 55#

Public Const DIM_TOLERANCE_MM As Double = 0.05

Public Function MmToCm(ByVal oDoc As DrawingDocument, ByVal valueMm As Double) As Double
    Dim oUom As UnitsOfMeasure

    If oDoc Is Nothing Then
        MmToCm = valueMm * MM_TO_CM
        Exit Function
    End If

    Set oUom = oDoc.UnitsOfMeasure
    MmToCm = oUom.ConvertUnits(valueMm, kMillimeterLengthUnits, kCentimeterLengthUnits)
End Function

Public Function CmToMm(ByVal oDoc As DrawingDocument, ByVal valueCm As Double) As Double
    Dim oUom As UnitsOfMeasure

    If oDoc Is Nothing Then
        CmToMm = valueCm / MM_TO_CM
        Exit Function
    End If

    Set oUom = oDoc.UnitsOfMeasure
    CmToMm = oUom.ConvertUnits(valueCm, kCentimeterLengthUnits, kMillimeterLengthUnits)
End Function

Public Function Pt(ByVal x As Double, ByVal y As Double) As Point2d
    Set Pt = ThisApplication.TransientGeometry.CreatePoint2d(x, y)
End Function

Public Function CanEditDrawingResources(ByVal oApp As Inventor.Application) As Boolean
    Dim eo As Object

    CanEditDrawingResources = False

    If oApp Is Nothing Then Exit Function

    On Error Resume Next
    Set eo = oApp.ActiveEditObject
    On Error GoTo 0

    If Not eo Is Nothing Then
        If TypeOf eo Is DrawingSketch Or TypeOf eo Is Sketch Then
            Debug.Print "ActiveEditObject=" & TypeName(eo)
            MsgBox "Finish active sketch/resource edit before running macro.", vbExclamation
            Exit Function
        End If
    End If

    CanEditDrawingResources = True
End Function

Public Function GetActiveDrawingDocument(ByVal oApp As Inventor.Application) As DrawingDocument
    If oApp Is Nothing Then Exit Function
    If oApp.ActiveDocument Is Nothing Then
        MsgBox "Open a drawing document first.", vbExclamation
        Exit Function
    End If

    If oApp.ActiveDocument.DocumentType <> kDrawingDocumentObject Then
        MsgBox "Active document is not a drawing.", vbExclamation
        Exit Function
    End If

    Set GetActiveDrawingDocument = oApp.ActiveDocument
End Function

Public Function EnsureA3LandscapeSheet(ByVal oDoc As DrawingDocument) As Sheet
    Dim oSheet As Sheet

    If oDoc Is Nothing Then Exit Function
    Set oSheet = oDoc.ActiveSheet
    If oSheet Is Nothing Then Exit Function

    On Error GoTo TryCreateSheet
    oSheet.ChangeSize kA3DrawingSheetSize, kLandscapePageOrientation
    oSheet.Activate
    On Error GoTo 0

    Set EnsureA3LandscapeSheet = oSheet
    Exit Function

TryCreateSheet:
    On Error GoTo ResizeFailed
    Err.Clear
    Set oSheet = oDoc.Sheets.Add(kA3DrawingSheetSize, kLandscapePageOrientation)
    oSheet.Activate
    On Error GoTo 0

    Set EnsureA3LandscapeSheet = oSheet
    Exit Function

ResizeFailed:
    On Error GoTo 0
    MsgBox "Could not set or create A3 Landscape sheet.", vbCritical
End Function

Public Function ValidateSpdsA3Sheet(ByVal oDoc As DrawingDocument, ByVal oSheet As Sheet) As Boolean
    Dim widthMm As Double
    Dim heightMm As Double

    ValidateSpdsA3Sheet = False

    If oDoc Is Nothing Then Exit Function
    If oSheet Is Nothing Then Exit Function

    widthMm = CmToMm(oDoc, oSheet.Width)
    heightMm = CmToMm(oDoc, oSheet.Height)

    If Abs(widthMm - A3_WIDTH_MM) > DIM_TOLERANCE_MM Or Abs(heightMm - A3_HEIGHT_MM) > DIM_TOLERANCE_MM Then
        MsgBox "Active sheet is not A3 landscape after resize/create." & vbCrLf & _
               "Expected (mm): 420 x 297" & vbCrLf & _
               "Actual (mm): " & FormatNumber(widthMm, 2) & " x " & FormatNumber(heightMm, 2), vbCritical
        Exit Function
    End If

    ValidateSpdsA3Sheet = True
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

Public Sub ClearSketch(ByVal oSketch As DrawingSketch)
    Dim i As Long

    If oSketch Is Nothing Then Exit Sub

    For i = oSketch.TextBoxes.Count To 1 Step -1
        oSketch.TextBoxes.Item(i).Delete
    Next i

    For i = oSketch.SketchLines.Count To 1 Step -1
        oSketch.SketchLines.Item(i).Delete
    Next i
End Sub

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

Public Function RuText(ByVal ParamArray cps() As Variant) As String
    Dim i As Long
    Dim result As String

    result = ""
    For i = LBound(cps) To UBound(cps)
        result = result & ChrW(CLng(cps(i)))
    Next i

    RuText = result
End Function
