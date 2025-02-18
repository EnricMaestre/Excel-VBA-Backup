Attribute VB_Name = "AutoBackup"

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    ActiveWorkbook.Save
    If Err.Number <> 0 Then
        MsgBox "Error al guardar el archivo antes de cerrar.", vbExclamation, "Error"
    End If
    On Error GoTo 0
End Sub

Private Sub Workbook_Open()
    Dim ruta1 As String
    Dim ruta2 As String
    Dim nombreArchivo As String
    Dim fechaHora As String
    
    fechaHora = Format(Now, "yyyy-mm-dd_hhmmss")
    nombreArchivo = "COP.SEG._repartos_plataformas_" & fechaHora & ".xlsm"
    
    ' 📂 Rutas de ejemplo para guardar las copias de seguridad
    ruta1 = "C:\Backup_Excel\Seguridad1\" & nombreArchivo
    ruta2 = "D:\Copias_Excel\Seguridad2\" & nombreArchivo

    On Error Resume Next

    ActiveWorkbook.SaveCopyAs Filename:=ruta1
    If Err.Number <> 0 Then
        MsgBox "⚠️ Error al guardar la copia de seguridad en: " & ruta1, vbExclamation, "Error"
    End If

    Err.Clear
    ActiveWorkbook.SaveCopyAs Filename:=ruta2
    If Err.Number <> 0 Then
        MsgBox "⚠️ Error al guardar la copia de seguridad en: " & ruta2, vbExclamation, "Error"
    End If

    Err.Clear
    On Error GoTo 0

    On Error Resume Next
    ActiveWorkbook.RefreshAll
    If Err.Number <> 0 Then
        MsgBox "⚠️ Error al actualizar las conexiones de datos.", vbExclamation, "Error"
    End If
    On Error GoTo 0
End Sub
