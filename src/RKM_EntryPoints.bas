Attribute VB_Name = "RKM_EntryPoints"
Option Explicit

Public Sub RKM_CreateNewDrawingAndApplyFrameAndTitleBlock()
    Dim oDoc As DrawingDocument
    Dim oSheet As Sheet

    If Not CanEditDrawingResources(ThisApplication) Then Exit Sub

    Set oDoc = CreateNewA3LandscapeDrawing(ThisApplication)
    If oDoc Is Nothing Then
        MsgBox "Could not create a new drawing document.", vbCritical
        Exit Sub
    End If

    Set oSheet = GetActiveSheet(oDoc)
    Call ApplyRkmResourcesToSheet(oDoc, oSheet)
End Sub

Public Sub RKM_ApplyFrameAndTitleBlockToActiveSheet()
    Dim oDoc As DrawingDocument
    Dim oSheet As Sheet

    If Not CanEditDrawingResources(ThisApplication) Then Exit Sub

    Set oDoc = GetActiveDrawingDocument(ThisApplication)
    If oDoc Is Nothing Then Exit Sub

    Set oSheet = GetActiveSheet(oDoc)
    Call ApplyRkmResourcesToSheet(oDoc, oSheet)
End Sub

Public Sub RKM_RebuildDefinitionsInActiveDrawing()
    Dim oDoc As DrawingDocument
    Dim oBorderDef As BorderDefinition

    If Not CanEditDrawingResources(ThisApplication) Then Exit Sub

    Set oDoc = GetActiveDrawingDocument(ThisApplication)
    If oDoc Is Nothing Then Exit Sub

    Set oBorderDef = EnsureRkmBorderDefinition(oDoc)

    If oBorderDef Is Nothing Then
        MsgBox "Could not rebuild border definition.", vbExclamation
        Exit Sub
    End If

    MsgBox "Border definition is rebuilt or versioned in the active drawing.", vbInformation
End Sub

Private Sub ApplyRkmResourcesToSheet(ByVal oDoc As DrawingDocument, ByVal oSheet As Sheet)
    Dim oBorderDef As BorderDefinition

    If oDoc Is Nothing Then Exit Sub
    If oSheet Is Nothing Then Exit Sub

    Set oBorderDef = EnsureRkmBorderDefinition(oDoc)

    Call ApplyRkmBorderToSheet(oSheet, oBorderDef)

    MsgBox "RKM A3 border was applied.", vbInformation
End Sub
