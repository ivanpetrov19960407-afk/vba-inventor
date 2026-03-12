Attribute VB_Name = "RKM_SelfTests"
Option Explicit

Public Sub Rkm_SelfTest_Create3ViewsOnActiveSheet()
    Dim oDoc As DrawingDocument
    Dim oSheet As Sheet
    Dim oModelDoc As Document
    Dim blockedRect As Object
    Dim i As Long
    Dim oView As DrawingView
    Dim collisions As Long

    Set oDoc = GetActiveDrawingDocument(ThisApplication)
    If oDoc Is Nothing Then Exit Sub

    Set oSheet = oDoc.ActiveSheet
    If oSheet Is Nothing Then Exit Sub

    Set oModelDoc = GetFirstReferencedModel(oDoc)
    If oModelDoc Is Nothing Then
        Debug.Print "SELFTEST: no referenced model found on active drawing."
        Exit Sub
    End If

    RemoveAllDrawingViewsFromSheet oSheet
    BuildSheetViews_Orthographic3 oDoc, oSheet, oModelDoc, Nothing

    Debug.Print "SELFTEST: views on active sheet = " & CStr(oSheet.DrawingViews.Count)
    Debug.Assert oSheet.DrawingViews.Count = 3

    Set blockedRect = SelfTest_GetTitleBlockBlockedRectCm(oDoc)
    collisions = 0

    For i = 1 To oSheet.DrawingViews.Count
        Set oView = oSheet.DrawingViews.Item(i)
        If SelfTest_ViewIntersectsRect(oView, blockedRect) Then
            collisions = collisions + 1
            Debug.Print "SELFTEST: blocked area collision; view=" & oView.Name
        End If
        Debug.Assert Not SelfTest_ViewIntersectsRect(oView, blockedRect)
    Next i

    Debug.Print "SELFTEST: blocked collisions = " & CStr(collisions)
    Debug.Assert collisions = 0
End Sub

Private Function GetFirstReferencedModel(ByVal oDoc As DrawingDocument) As Document
    Dim i As Long
    Dim oView As DrawingView

    If oDoc Is Nothing Then Exit Function
    If oDoc.ActiveSheet Is Nothing Then Exit Function

    For i = 1 To oDoc.ActiveSheet.DrawingViews.Count
        Set oView = oDoc.ActiveSheet.DrawingViews.Item(i)
        On Error Resume Next
        Set GetFirstReferencedModel = oView.ReferencedDocumentDescriptor.ReferencedDocument
        On Error GoTo 0
        If Not GetFirstReferencedModel Is Nothing Then Exit Function
    Next i
End Function

Private Sub RemoveAllDrawingViewsFromSheet(ByVal oSheet As Sheet)
    Dim i As Long

    If oSheet Is Nothing Then Exit Sub
    For i = oSheet.DrawingViews.Count To 1 Step -1
        oSheet.DrawingViews.Item(i).Delete
    Next i
End Sub

Private Function SelfTest_GetTitleBlockBlockedRectCm(ByVal oDoc As DrawingDocument) As Object
    Dim oSheet As Sheet
    Dim safeLeft As Double
    Dim safeRight As Double
    Dim safeBottom As Double

    Set oSheet = oDoc.ActiveSheet

    safeLeft = MmToCm(oDoc, FRAME_LEFT_MM)
    safeRight = oSheet.Width - MmToCm(oDoc, FRAME_OTHER_MM)
    safeBottom = MmToCm(oDoc, FRAME_OTHER_MM)

    Set SelfTest_GetTitleBlockBlockedRectCm = SelfTest_CreateRect( _
        safeRight - MmToCm(oDoc, 185#), _
        safeRight, _
        safeBottom, _
        safeBottom + MmToCm(oDoc, 55#))
End Function

Private Function SelfTest_CreateRect(ByVal leftCm As Double, ByVal rightCm As Double, ByVal bottomCm As Double, ByVal topCm As Double) As Object
    Dim rect As Object

    Set rect = CreateObject("Scripting.Dictionary")
    rect.CompareMode = vbTextCompare
    rect("Left") = leftCm
    rect("Right") = rightCm
    rect("Bottom") = bottomCm
    rect("Top") = topCm

    Set SelfTest_CreateRect = rect
End Function

Private Function SelfTest_ViewIntersectsRect(ByVal oView As DrawingView, ByVal rect As Object) As Boolean
    Dim viewRect As Object

    If oView Is Nothing Then Exit Function

    Set viewRect = SelfTest_CreateRect(oView.Left, oView.Left + oView.Width, oView.Top - oView.Height, oView.Top)
    SelfTest_ViewIntersectsRect = Not (viewRect("Right") <= rect("Left") Or rect("Right") <= viewRect("Left") Or viewRect("Top") <= rect("Bottom") Or rect("Top") <= viewRect("Bottom"))
End Function
