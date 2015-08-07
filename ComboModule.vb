Module ComboModule


    Public Sub GenericCombo(sender As Object, e As EventArgs)

        If sender.SelectedValue.ToString = vbNullString Then Exit Sub

        If OverClass.UnloadData() = True Then Exit Sub
        OverClass.ResetCollection()
        Call SubCombo(sender)


    End Sub

    Private Sub SubCombo(sender As ComboBox)

        Select Case sender.Name.ToString

            'Case "ComboBox4"
            'StartCombo(Form1.ComboBox3)

            Case Else
                ComboRefreshData(sender)

        End Select

    End Sub

    Public Sub StartCombo(ctl As ComboBox)

        Select Case ctl.Name.ToString()


            'Case "ComboBox18"
            '   If IsNothing(Form1.ComboBox17.SelectedValue) Then Exit Sub
            '   ctl.DataSource = OverClass.TempDataTable("SELECT CohortID, " & _
            '                                                  "CohortName FROM Cohort WHERE StudyID=" _
            '                                                  & Form1.ComboBox17.SelectedValue.ToString & _
            '                                                  " AND Generated=True " & _
            '                                                  " ORDER BY CohortName ASC")
            '   ctl.ValueMember = "CohortID"
            '   ctl.DisplayMember = "CohortName"

        End Select

        ComboRefreshData(ctl)

    End Sub

    Public Sub ComboRefreshData(sender As ComboBox)

        Dim Grid As DataGridView = Nothing

        Select Case sender.Name.ToString()

            'Case "ComboBox22"
            '   Grid = Form1.DataGridView13


        End Select


        'If Not IsNothing(Grid) Then Call Form1.Specifics(Grid)

    End Sub


End Module
