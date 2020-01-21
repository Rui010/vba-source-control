' =====================================
' �i�[���Ă��郂�W���[����ǂݍ���
' ThisWorkbook�ɓ\��t���Ďg��
' =====================================
Private Sub Workbook_Open()
    ' ���W���[�����i�[���Ă���t�H���_
    Const MODULE_DIR As String = ".\module"
    
    Dim file_dir As String
    Dim file_name As String
    Dim file_path As String

    If Left(MODULE_DIR, 1) = "." Then
        file_dir = ThisWorkbook.Path & Mid(MODULE_DIR, 2, Len(MODULE_DIR) - 1)
    Else
        file_dir = MODULE_DIR
    
    ' ���W���[���̍폜
    ' 1: �W�����W���[��, 2: �N���X���W���[��
    For Each component In ThisWorkbook.VBProject.VBComponents
        If component.Type = 1 Or component.Type = 2 Then
            ThisWorkbook.VBProject.VBComponents.Remove component
        End If
    Next component
    
    ' ���W���[���̒ǉ�
    file_name = Dir(file_dir)
    Do While file_name <> ""
        file_path = file_dir & "\" & file_name
        ThisWorkbook.VBProject.VBComponents.Import file_path
        file_name = Dir()
    Loop
    
    ThisWorkbook.Save
    
End Sub
