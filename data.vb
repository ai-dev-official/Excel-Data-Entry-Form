Private Sub Label14_Click()

End Sub

Private Sub btnAdd_Click()

    Dim wks As Worksheet
    Dim AddNew As Range
    Set wks = Sheet1

    Set AddNew = wks.Range("A65356").End(xlUp).Offset(1, 0)

    AddNew.Offset(0, 0).Value = txtName.Text
    AddNew.Offset(0, 1).Value = txtDob.Text
    AddNew.Offset(0, 2).Value = txtGender.Text
    AddNew.Offset(0, 3).Value = txtPhone.Text
    AddNew.Offset(0, 4).Value = txtEmail.Text
    AddNew.Offset(0, 5).Value = txtStudentId.Text
    AddNew.Offset(0, 6).Value = txtUniversity.Text
    AddNew.Offset(0, 7).Value = txtMajor.Text
    AddNew.Offset(0, 8).Value = txtYear.Text
    AddNew.Offset(0, 9).Value = txtEmergency.Text
    AddNew.Offset(0, 10).Value = txtRelation.Text
    AddNew.Offset(0, 11).Value = txtAllergies.Text
    AddNew.Offset(0, 12).Value = txtNationality.Text
    AddNew.Offset(0, 13).Value = txtCitizenship.Text
    AddNew.Offset(0, 14).Value = txtLanguages.Text
    AddNew.Offset(0, 15).Value = txtInterests.Text


    displayBox.ColumnCount = 16
    displayBox.RowSource = "A1: P65356"

End Sub

Private Sub btnDelete_Click()


    Dim i As Integer

    For i = 0 To Range("A65356").End(xlUp).Row - 1
        If displayBox.Selected(i) Then
            Rows(i).Select
            Selection.Delete
        End If

    Next i



End Sub

Private Sub btnExit_Click()

    Dim iExit As VbMsgBoxResult

    iExit = MsgBox("Confirm If you want To Exit", vbQuestion + vbYesNo, "Data Entry Form")

    If iExit = vbYes Then
        Unload Me
    End If

End Sub

Private Sub btnReset_Click()

    Dim iControl As Control

    For Each iControl In Me.Controls

        If iControl.Name Like "txt*" Then iControl = vbNullString

            Next

End Sub



