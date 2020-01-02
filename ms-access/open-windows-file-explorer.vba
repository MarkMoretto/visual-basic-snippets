Private Sub OpenExplorer_btn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

''' Notes:
''' Subroutine that should open a Windows file expolorer when clicked.
''' This particular function is used in MS Access
''' Me.OutPath_txt.Value is a textbox on the targeted user form.

Dim f_dialog As Office.FileDialog
Dim sel_fldr As String
Dim curr_val As String
Dim init_fldr As String
Dim v_file

init_fldr = "C:\Users\" & Environ("USERNAME") & "\Desktop"
curr_val = Nz(Me.OutPath_txt.Value, "")

Set f_dialog = Application.FileDialog(msoFileDialogFolderPicker)

With f_dialog
    .Title = "Select Export Destination Folder"
    .AllowMultiSelect = False
    If Len(Nz(Me.OutPath_txt.Value, "")) = 0 Then
        .InitialFileName = init_fldr
    Else
        .InitialFileName = Me.OutPath_txt.Value
    End If
    If .Show <> -1 Then
        If Len(Nz(curr_val, "")) = 0 Then
            Me.OutPath_txt.Value = curr_val
        Else
            Me.OutPath_txt.Value = ""
        End If
    Else
        Me.OutPath_txt.Value = .SelectedItems(1)
    End If
End With

export_folder = Nz(Me.OutPath_txt.Value, "")

Set f_dialog = Nothing

End Sub
