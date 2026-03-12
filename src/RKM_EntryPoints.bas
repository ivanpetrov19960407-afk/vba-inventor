Attribute VB_Name = "RKM_EntryPoints"
Option Explicit

Public Sub Rkm_CreateOrApplyA3Frame()
    Rkm_CreateOrApplyA3Frame_SPDS
End Sub

Public Sub Rkm_BuildOrUpdateAlbum()
    Rkm_BuildOrUpdateIdwAlbum
End Sub

Public Sub Rkm_BuildAlbumFromExcel_OnActiveDrawing()
    Dim oDoc As DrawingDocument
    Dim excelPath As String

    Set oDoc = GetActiveDrawingDocument(ThisApplication)
    If oDoc Is Nothing Then Exit Sub

    excelPath = PickExcelFilePath()
    If Len(excelPath) = 0 Then Exit Sub

    Rkm_BuildOrUpdateIdwAlbum_FromExcel oDoc, excelPath
End Sub

Public Sub Rkm_BuildAlbumFromExcel_AndSaveAs()
    Dim excelPath As String
    Dim savePath As String
    Dim oDoc As DrawingDocument
    Dim templatePath As String

    excelPath = PickExcelFilePath()
    If Len(excelPath) = 0 Then Exit Sub

    savePath = PickIdwSavePath("Album.idw")
    If Len(savePath) = 0 Then Exit Sub

    templatePath = ThisApplication.FileManager.GetTemplateFile(kDrawingDocumentObject)
    Set oDoc = ThisApplication.Documents.Add(kDrawingDocumentObject, templatePath, True)

    Rkm_BuildOrUpdateIdwAlbum_FromExcel oDoc, excelPath

    On Error GoTo EH
    oDoc.SaveAs savePath, False
    Exit Sub
EH:
    MsgBox "SaveAs failed (Err " & Err.Number & "): " & Err.Description, vbCritical
End Sub

Public Sub Rkm_RunSelfTest_Create3ViewsOnActiveSheet()
    Rkm_SelfTest_Create3ViewsOnActiveSheet
End Sub
