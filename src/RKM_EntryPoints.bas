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
    Dim oTitleDef As TitleBlockDefinition

    If Not CanEditDrawingResources(ThisApplication) Then Exit Sub

    Set oDoc = GetActiveDrawingDocument(ThisApplication)
    If oDoc Is Nothing Then Exit Sub

    Set oBorderDef = EnsureRkmBorderDefinition(oDoc)
    Set oTitleDef = EnsureRkmTitleBlockDefinition(oDoc)

    If oBorderDef Is Nothing Or oTitleDef Is Nothing Then
        MsgBox "Could not rebuild one or more definitions.", vbExclamation
        Exit Sub
    End If

    MsgBox "Definitions are rebuilt or versioned in the active drawing.", vbInformation
End Sub

Private Sub ApplyRkmResourcesToSheet(ByVal oDoc As DrawingDocument, ByVal oSheet As Sheet)
    Dim oBorderDef As BorderDefinition
    Dim oTitleDef As TitleBlockDefinition
    Dim prompts As Variant

    If oDoc Is Nothing Then Exit Sub
    If oSheet Is Nothing Then Exit Sub

    Set oBorderDef = EnsureRkmBorderDefinition(oDoc)
    Set oTitleDef = EnsureRkmTitleBlockDefinition(oDoc)

    prompts = DefaultPromptValues()

    Call ApplyRkmBorderToSheet(oSheet, oBorderDef)
    Call ApplyRkmTitleBlockToSheet(oSheet, oTitleDef, prompts)

    MsgBox "RKM border and prompted title block were applied.", vbInformation
End Sub
