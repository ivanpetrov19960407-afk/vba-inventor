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

    Set oModelDoc = GetFirstReferencedModelAnywhere(oDoc)
    If oModelDoc Is Nothing Then
        MsgBox "SELFTEST FAILED: no referenced model found on any sheet.", vbExclamation
        Debug.Print "SELFTEST: no referenced model found on any sheet."
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
        If oSheet.DrawingViews.Count = 0 Then
            MsgBox "SELFTEST FAILED: 0 views. Check Immediate window. Most likely base view creation/fit failed.", vbExclamation
        Else
            MsgBox "SELFTEST FAILED: expected 3 views, actual = " & CStr(oSheet.DrawingViews.Count), vbExclamation
        End If
        Exit Sub
    End If

    If collisions = 0 Then
        MsgBox "SELFTEST PASSED", vbInformation
    Else
        MsgBox "SELFTEST FAILED: views=" & CStr(oSheet.DrawingViews.Count) & "; collisions=" & CStr(collisions) & "; details in Immediate window.", vbExclamation
    End If
End Sub

Public Sub Rkm_SelfTest_BaseViewOnly_OnActiveSheet()
    Dim oDoc As DrawingDocument
    Dim oSheet As Sheet
    Dim oModelDoc As Document
    Dim firstAngle As Boolean
    Dim blockedRect As Object
    Dim frontRect As Object
    Dim baseView As DrawingView
    Dim scaleValue As Double

    Set oDoc = GetActiveDrawingDocument(ThisApplication)
    If oDoc Is Nothing Then Exit Sub

    Set oSheet = oDoc.ActiveSheet
    If oSheet Is Nothing Then Exit Sub

    Set oModelDoc = GetFirstReferencedModelAnywhere(oDoc)
    If oModelDoc Is Nothing Then
        MsgBox "BASEVIEW FAILED: no referenced model found on any sheet.", vbExclamation
        Debug.Print "SELFTEST BASEVIEW: no referenced model found on any sheet."
        Exit Sub
    End If

    RemoveAllDrawingViewsFromSheet oSheet

    firstAngle = SelfTest_GetProjectionStandard(oDoc)
    Set frontRect = SelfTest_GetFrontViewRectCm(oDoc, firstAngle)
    Set blockedRect = SelfTest_GetTitleBlockBlockedRectCm(oDoc)
    scaleValue = 1#

    Debug.Print "SELFTEST BASEVIEW: firstAngle=" & CStr(firstAngle) & _
                "; model=" & oModelDoc.DisplayName & _
                "; modelPath=" & oModelDoc.FullFileName & _
                "; scale=" & CStr(scaleValue)
    Debug.Print "SELFTEST BASEVIEW: frontRect L=" & CStr(frontRect("Left")) & ", R=" & CStr(frontRect("Right")) & ", B=" & CStr(frontRect("Bottom")) & ", T=" & CStr(frontRect("Top"))
    Debug.Print "SELFTEST BASEVIEW: blockedRect L=" & CStr(blockedRect("Left")) & ", R=" & CStr(blockedRect("Right")) & ", B=" & CStr(blockedRect("Bottom")) & ", T=" & CStr(blockedRect("Top"))

    On Error Resume Next
    Set baseView = oSheet.DrawingViews.AddBaseView( _
        oModelDoc, Pt((frontRect("Left") + frontRect("Right")) / 2#, (frontRect("Bottom") + frontRect("Top")) / 2#), _
        scaleValue, kFrontViewOrientation, kHiddenLineRemovedDrawingViewStyle)
    On Error GoTo 0

    If baseView Is Nothing Then
        MsgBox "BASEVIEW FAILED: AddBaseView returned Nothing", vbExclamation
        Exit Sub
    End If

    Debug.Print "SELFTEST BASEVIEW: view Left=" & CStr(baseView.Left) & "; Top=" & CStr(baseView.Top) & "; Width=" & CStr(baseView.Width) & "; Height=" & CStr(baseView.Height)

    If Not SelfTest_ViewFitsRect(baseView, frontRect) Then
        MsgBox "BASEVIEW FAILED: frontRect fit failed", vbExclamation
        Exit Sub
    End If

    If SelfTest_ViewIntersectsRect(baseView, blockedRect) Then
        MsgBox "BASEVIEW FAILED: blockedRect collision", vbExclamation
        Exit Sub
    End If

    MsgBox "BASEVIEW OK", vbInformation
End Sub

Public Sub Rkm_SelfTest_BaseView_FromPickedModel()
    Dim oDoc As DrawingDocument
    Dim oSheet As Sheet
    Dim oModelDoc As Document
    Dim firstAngle As Boolean
    Dim frontRect As Object
    Dim blockedRect As Object
    Dim baseView As DrawingView
    Dim scaleValue As Double
    Dim reasonText As String

    Set oDoc = GetActiveDrawingDocument(ThisApplication)
    If oDoc Is Nothing Then Exit Sub

    Set oSheet = oDoc.ActiveSheet
    If oSheet Is Nothing Then Exit Sub

    Set oModelDoc = SelfTest_PickModelDocument(ThisApplication)
    If oModelDoc Is Nothing Then
        MsgBox "BASEVIEW CANCELLED: model not selected", vbExclamation
        Debug.Print "SELFTEST BASEVIEW PICKED: failure reason=model not selected"
        Exit Sub
    End If

    RemoveAllDrawingViewsFromSheet oSheet

    firstAngle = SelfTest_GetProjectionStandard(oDoc)
    Set frontRect = SelfTest_GetFrontViewRectCm(oDoc, firstAngle)
    Set blockedRect = SelfTest_GetTitleBlockBlockedRectCm(oDoc)
    scaleValue = 1#

    Debug.Print "SELFTEST BASEVIEW PICKED: modelPath=" & oModelDoc.FullFileName
    Debug.Print "SELFTEST BASEVIEW PICKED: activeSheet=" & oSheet.Name
    Debug.Print "SELFTEST BASEVIEW PICKED: firstAngle=" & CStr(firstAngle)
    Debug.Print "SELFTEST BASEVIEW PICKED: frontRect L=" & CStr(frontRect("Left")) & ", R=" & CStr(frontRect("Right")) & ", B=" & CStr(frontRect("Bottom")) & ", T=" & CStr(frontRect("Top"))
    Debug.Print "SELFTEST BASEVIEW PICKED: blockedRect L=" & CStr(blockedRect("Left")) & ", R=" & CStr(blockedRect("Right")) & ", B=" & CStr(blockedRect("Bottom")) & ", T=" & CStr(blockedRect("Top"))

    On Error Resume Next
    Set baseView = oSheet.DrawingViews.AddBaseView( _
        oModelDoc, Pt(oSheet.Width / 2#, oSheet.Height / 2#), _
        scaleValue, kFrontViewOrientation, kHiddenLineRemovedDrawingViewStyle)
    On Error GoTo 0

    If baseView Is Nothing Then
        reasonText = "AddBaseView returned Nothing"
        Debug.Print "SELFTEST BASEVIEW PICKED: failure reason=" & reasonText
        MsgBox "BASEVIEW FAILED: AddBaseView returned Nothing. Check Immediate window.", vbExclamation
        Exit Sub
    End If

    Debug.Print "SELFTEST BASEVIEW PICKED: view Left=" & CStr(baseView.Left) & "; Top=" & CStr(baseView.Top) & "; Width=" & CStr(baseView.Width) & "; Height=" & CStr(baseView.Height)

    If Not SelfTest_ViewFitsRect(baseView, frontRect) Then
        reasonText = "frontRect fit failed"
        Debug.Print "SELFTEST BASEVIEW PICKED: failure reason=" & reasonText
        MsgBox "BASEVIEW FAILED: frontRect fit failed. Check Immediate window.", vbExclamation
        Exit Sub
    End If

    If SelfTest_ViewIntersectsRect(baseView, blockedRect) Then
        reasonText = "blockedRect collision"
        Debug.Print "SELFTEST BASEVIEW PICKED: failure reason=" & reasonText
        MsgBox "BASEVIEW FAILED: blockedRect collision. Check Immediate window.", vbExclamation
        Exit Sub
    End If

    MsgBox "BASEVIEW PASSED", vbInformation
End Sub

Public Sub Rkm_SelfTest_Create3Views_FromPickedModel()
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

    Set oModelDoc = SelfTest_PickModelDocument(ThisApplication)
    If oModelDoc Is Nothing Then
        MsgBox "SELFTEST CANCELLED: model not selected", vbExclamation
        Exit Sub
    End If

    RemoveAllDrawingViewsFromSheet oSheet

    firstAngle = SelfTest_GetProjectionStandard(oDoc)
    Set frontRect = SelfTest_GetFrontViewRectCm(oDoc, firstAngle)
    Set topRect = SelfTest_GetTopProjectedRectCm(oDoc, firstAngle)
    Set sideRect = SelfTest_GetSideProjectedRectCm(oDoc, firstAngle)
    Set blockedRect = SelfTest_GetTitleBlockBlockedRectCm(oDoc)

    Debug.Print "SELFTEST PICKED MODEL: " & oModelDoc.FullFileName
    Debug.Print "SELFTEST SHEET: " & oSheet.Name
    Debug.Print "SELFTEST firstAngle=" & CStr(firstAngle)
    Debug.Print "SELFTEST frontRect L=" & CStr(frontRect("Left")) & ", R=" & CStr(frontRect("Right")) & ", B=" & CStr(frontRect("Bottom")) & ", T=" & CStr(frontRect("Top"))
    Debug.Print "SELFTEST topRect L=" & CStr(topRect("Left")) & ", R=" & CStr(topRect("Right")) & ", B=" & CStr(topRect("Bottom")) & ", T=" & CStr(topRect("Top"))
    Debug.Print "SELFTEST sideRect L=" & CStr(sideRect("Left")) & ", R=" & CStr(sideRect("Right")) & ", B=" & CStr(sideRect("Bottom")) & ", T=" & CStr(sideRect("Top"))
    Debug.Print "SELFTEST blockedRect L=" & CStr(blockedRect("Left")) & ", R=" & CStr(blockedRect("Right")) & ", B=" & CStr(blockedRect("Bottom")) & ", T=" & CStr(blockedRect("Top"))

    BuildSheetViews_Orthographic3 oDoc, oSheet, oModelDoc, Nothing

    Debug.Print "SELFTEST: views on active sheet = " & CStr(oSheet.DrawingViews.Count)
    For i = 1 To oSheet.DrawingViews.Count
        Set oView = oSheet.DrawingViews.Item(i)
        Debug.Print "SELFTEST: view[" & CStr(i) & "] name=" & oView.Name & _
                    "; Left=" & CStr(oView.Left) & _
                    "; Top=" & CStr(oView.Top) & _
                    "; Width=" & CStr(oView.Width) & _
                    "; Height=" & CStr(oView.Height)
    Next i

    collisions = 0
    For i = 1 To oSheet.DrawingViews.Count
        Set oView = oSheet.DrawingViews.Item(i)
        If SelfTest_ViewIntersectsRect(oView, blockedRect) Then collisions = collisions + 1
    Next i

    If oSheet.DrawingViews.Count = 3 Then
        MsgBox "SELFTEST PASSED", vbInformation
    Else
        MsgBox "SELFTEST FAILED: expected 3 views, actual = " & CStr(oSheet.DrawingViews.Count) & ". Check Immediate window.", vbExclamation
    End If
End Sub

Private Function SelfTest_PickModelDocument(ByVal oApp As Inventor.Application) As Document
    Dim oDlg As FileDialog
    Dim filePath As String

    If oApp Is Nothing Then Exit Function

    On Error GoTo EH
    oApp.CreateFileDialog oDlg
    If oDlg Is Nothing Then Exit Function

    oDlg.DialogTitle = "Select model for self-test"
    oDlg.Filter = "Inventor Models (*.ipt;*.iam)|*.ipt;*.iam"
    oDlg.FilterIndex = 1
    oDlg.ShowOpen

    filePath = Trim$(oDlg.FileName)
    If Len(filePath) = 0 Then Exit Function

    Set SelfTest_PickModelDocument = oApp.Documents.Open(filePath, True)
    Exit Function
EH:
    Set SelfTest_PickModelDocument = Nothing
End Function

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

Private Function GetFirstReferencedModelAnywhere(ByVal oDoc As DrawingDocument) As Document
    Dim sheetIndex As Long
    Dim viewIndex As Long
    Dim oSheet As Sheet
    Dim oView As DrawingView

    If oDoc Is Nothing Then Exit Function

    For sheetIndex = 1 To oDoc.Sheets.Count
        Set oSheet = oDoc.Sheets.Item(sheetIndex)
        For viewIndex = 1 To oSheet.DrawingViews.Count
            Set oView = oSheet.DrawingViews.Item(viewIndex)
            On Error Resume Next
            Set GetFirstReferencedModelAnywhere = oView.ReferencedDocumentDescriptor.ReferencedDocument
            On Error GoTo 0
            If Not GetFirstReferencedModelAnywhere Is Nothing Then Exit Function
        Next viewIndex
    Next sheetIndex
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

Private Function SelfTest_ViewFitsRect(ByVal oView As DrawingView, ByVal rect As Object) As Boolean
    Dim viewRect As Object

    If oView Is Nothing Then Exit Function

    Set viewRect = SelfTest_CreateRect(oView.Left, oView.Left + oView.Width, oView.Top - oView.Height, oView.Top)
    SelfTest_ViewFitsRect = (viewRect("Left") >= rect("Left")) And (viewRect("Right") <= rect("Right")) And (viewRect("Bottom") >= rect("Bottom")) And (viewRect("Top") <= rect("Top"))
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
