Attribute VB_Name = "RKM_EntryPoints"
Option Explicit

Public Sub Rkm_CreateOrApplyA3Frame_SPDS()
    Dim oDoc As DrawingDocument
    Dim oSheet As Sheet
    Dim oBorderDef As BorderDefinition
    Dim oTitleDef As TitleBlockDefinition

    On Error GoTo FailHandler

    If Not CanEditDrawingResources(ThisApplication) Then Exit Sub

    Set oDoc = GetActiveDrawingDocument(ThisApplication)
    If oDoc Is Nothing Then Exit Sub

    Set oSheet = EnsureA3LandscapeSheet(oDoc)
    If oSheet Is Nothing Then Exit Sub

    If Not ValidateSpdsA3Sheet(oDoc, oSheet) Then Exit Sub

    Set oBorderDef = EnsureRkmBorderDefinition(oDoc)
    If oBorderDef Is Nothing Then Exit Sub

    Set oTitleDef = EnsureRkmTitleBlockDefinition(oDoc)
    If oTitleDef Is Nothing Then Exit Sub

    If Not CanEditDrawingResources(ThisApplication) Then Exit Sub

    ApplyRkmBorderToSheet oSheet, oBorderDef
    ApplyRkmTitleBlockToSheet oSheet, oTitleDef

    MsgBox "SPDS A3 frame and form 3 title block applied.", vbInformation
    Exit Sub

FailHandler:
    MsgBox "Rkm_CreateOrApplyA3Frame_SPDS failed." & vbCrLf & _
           "Err.Number: " & CStr(Err.Number) & vbCrLf & _
           "Err.Description: " & Err.Description, vbCritical
End Sub

Public Sub Rkm_CreateOrApplyA3Frame()
    Rkm_CreateOrApplyA3Frame_SPDS
End Sub
