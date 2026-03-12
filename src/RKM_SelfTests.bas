Attribute VB_Name = "RKM_SelfTests"
Option Explicit

Public Sub Rkm_SelfTest_Create3ViewsOnActiveSheet()
    Dim oDoc As DrawingDocument
    Dim oSheet As Sheet
    Dim oModelDoc As Document
    Dim blockedRect As Object
    Dim frontRect As Object
    Dim topRect As Object
    Dim sideRect As Object
    Dim firstAngle As Boolean
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
    firstAngle = SelfTest_GetProjectionStandard(oDoc)
    Set frontRect = SelfTest_GetFrontViewRectCm(oDoc, firstAngle)
    Set topRect = SelfTest_GetTopProjectedRectCm(oDoc, firstAngle)
    Set sideRect = SelfTest_GetSideProjectedRectCm(oDoc, firstAngle)
    Debug.Print "SELFTEST: firstAngle=" & CStr(firstAngle) & _
                "; frontRect L=" & CStr(frontRect("Left")) & ", R=" & CStr(frontRect("Right")) & ", B=" & CStr(frontRect("Bottom")) & ", T=" & CStr(frontRect("Top")) & _
                "; topRect L=" & CStr(topRect("Left")) & ", R=" & CStr(topRect("Right")) & ", B=" & CStr(topRect("Bottom")) & ", T=" & CStr(topRect("Top")) & _
                "; sideRect L=" & CStr(sideRect("Left")) & ", R=" & CStr(sideRect("Right")) & ", B=" & CStr(sideRect("Bottom")) & ", T=" & CStr(sideRect("Top"))

    BuildSheetViews_Orthographic3 oDoc, oSheet, oModelDoc, Nothing

    Debug.Print "SELFTEST: views on active sheet = " & CStr(oSheet.DrawingViews.Count)
    Call SelfTest_PrintSheetViews(oSheet)

    Set blockedRect = SelfTest_GetTitleBlockBlockedRectCm(oDoc)
    collisions = 0

    For i = 1 To oSheet.DrawingViews.Count
        Set oView = oSheet.DrawingViews.Item(i)
        If SelfTest_ViewIntersectsRect(oView, blockedRect) Then
            collisions = collisions + 1
            Debug.Print "SELFTEST: blocked area collision; view=" & oView.Name
        End If
    Next i

    Debug.Print "SELFTEST: blocked collisions = " & CStr(collisions)

    If oSheet.DrawingViews.Count <> 3 Then
        MsgBox "SELFTEST FAILED: expected 3 views, actual = " & CStr(oSheet.DrawingViews.Count), vbExclamation
        Exit Sub
    End If

    If collisions = 0 Then
        MsgBox "SELFTEST PASSED", vbInformation
    Else
        MsgBox "SELFTEST FAILED: views=" & CStr(oSheet.DrawingViews.Count) & "; collisions=" & CStr(collisions) & "; details in Immediate window.", vbExclamation
    End If
End Sub

Private Sub SelfTest_PrintSheetViews(ByVal oSheet As Sheet)
    Dim i As Long
    Dim oView As DrawingView

    If oSheet Is Nothing Then Exit Sub

    Debug.Print "SELFTEST: view list begin"
    For i = 1 To oSheet.DrawingViews.Count
        Set oView = oSheet.DrawingViews.Item(i)
        Debug.Print "SELFTEST: view[" & CStr(i) & "] name=" & oView.Name & _
                    "; Left=" & CStr(oView.Left) & _
                    "; Top=" & CStr(oView.Top) & _
                    "; Width=" & CStr(oView.Width) & _
                    "; Height=" & CStr(oView.Height)
    Next i
    Debug.Print "SELFTEST: view list end"
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

Private Function SelfTest_GetProjectionStandard(ByVal oDoc As DrawingDocument) As Boolean
    On Error GoTo EH
    SelfTest_GetProjectionStandard = oDoc.StylesManager.ActiveStandardStyle.FirstAngleProjection
    Exit Function
EH:
    SelfTest_GetProjectionStandard = True
End Function

Private Function SelfTest_GetSheetSafeRectCm(ByVal oDoc As DrawingDocument) As Object
    Dim oSheet As Sheet

    Set oSheet = oDoc.ActiveSheet
    Set SelfTest_GetSheetSafeRectCm = SelfTest_CreateRect( _
        MmToCm(oDoc, FRAME_LEFT_MM), _
        oSheet.Width - MmToCm(oDoc, FRAME_OTHER_MM), _
        MmToCm(oDoc, FRAME_OTHER_MM), _
        oSheet.Height - MmToCm(oDoc, FRAME_OTHER_MM))
End Function

Private Function SelfTest_GetFrontViewRectCm(ByVal oDoc As DrawingDocument, ByVal firstAngle As Boolean) As Object
    Dim safeRect As Object
    Dim splitX As Double
    Dim splitY As Double
    Dim padCm As Double

    Set safeRect = SelfTest_InsetRect(SelfTest_GetSheetSafeRectCm(oDoc), MmToCm(oDoc, 6#))
    padCm = MmToCm(oDoc, 8#)
    splitX = safeRect("Right") - (safeRect("Right") - safeRect("Left")) * 0.34
    splitY = safeRect("Top") - (safeRect("Top") - safeRect("Bottom")) * 0.36

    If firstAngle Then
        Set SelfTest_GetFrontViewRectCm = SelfTest_CreateRect(safeRect("Left"), splitX - padCm, splitY + padCm, safeRect("Top") - padCm)
    Else
        Set SelfTest_GetFrontViewRectCm = SelfTest_CreateRect(safeRect("Left"), splitX - padCm, safeRect("Bottom"), splitY - padCm)
    End If
End Function

Private Function SelfTest_GetTopProjectedRectCm(ByVal oDoc As DrawingDocument, ByVal firstAngle As Boolean) As Object
    Dim safeRect As Object
    Dim splitX As Double
    Dim splitY As Double
    Dim padCm As Double

    Set safeRect = SelfTest_InsetRect(SelfTest_GetSheetSafeRectCm(oDoc), MmToCm(oDoc, 6#))
    padCm = MmToCm(oDoc, 8#)
    splitX = safeRect("Right") - (safeRect("Right") - safeRect("Left")) * 0.34
    splitY = safeRect("Top") - (safeRect("Top") - safeRect("Bottom")) * 0.36

    If firstAngle Then
        Set SelfTest_GetTopProjectedRectCm = SelfTest_CreateRect(safeRect("Left"), splitX - padCm, safeRect("Bottom"), splitY - padCm)
    Else
        Set SelfTest_GetTopProjectedRectCm = SelfTest_CreateRect(safeRect("Left"), splitX - padCm, splitY + padCm, safeRect("Top") - padCm)
    End If
End Function

Private Function SelfTest_GetSideProjectedRectCm(ByVal oDoc As DrawingDocument, ByVal firstAngle As Boolean) As Object
    Dim safeRect As Object
    Dim frontRect As Object
    Dim splitX As Double
    Dim padCm As Double
    Dim topLimit As Double
    Dim bottomLimit As Double

    Set safeRect = SelfTest_InsetRect(SelfTest_GetSheetSafeRectCm(oDoc), MmToCm(oDoc, 6#))
    Set frontRect = SelfTest_GetFrontViewRectCm(oDoc, firstAngle)
    padCm = MmToCm(oDoc, 8#)
    splitX = safeRect("Right") - (safeRect("Right") - safeRect("Left")) * 0.34

    bottomLimit = frontRect("Bottom")
    topLimit = frontRect("Top")

    If bottomLimit < safeRect("Bottom") Then bottomLimit = safeRect("Bottom")
    If topLimit > safeRect("Top") Then topLimit = safeRect("Top")

    Set SelfTest_GetSideProjectedRectCm = SelfTest_CreateRect(splitX + padCm, safeRect("Right") - padCm, bottomLimit, topLimit)
End Function

Private Function SelfTest_InsetRect(ByVal rect As Object, ByVal deltaCm As Double) As Object
    Set SelfTest_InsetRect = SelfTest_CreateRect(rect("Left") + deltaCm, rect("Right") - deltaCm, rect("Bottom") + deltaCm, rect("Top") - deltaCm)
End Function
