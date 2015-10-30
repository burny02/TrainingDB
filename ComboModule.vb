Module ComboModule

    Public Sub GenericCombo(sender As Object, e As EventArgs)

        If OverClass.UnloadData() = True Then Exit Sub
        OverClass.ResetCollection()
        Call SubCombo(sender)


    End Sub

    Public Sub SubCombo(sender As ComboBox)

        Select Case sender.Name.ToString

            Case "ComboBox1", "ComboBox2", "ComboBox3", "ComboBox4"
                Form1.Specifics(Form1.ReportViewer1)
                StartCombo(Form1.ComboBox1)
                StartCombo(Form1.ComboBox2)
                StartCombo(Form1.ComboBox3)
                StartCombo(Form1.ComboBox4)

        End Select

    End Sub

    Public Sub StartCombo(ctl As ComboBox)

        Select Case ctl.Name.ToString()

            Case "ComboBox5"

                If ctl.SelectedValue <> "" Then Exit Sub

                Dim dt As DataTable = OverClass.TempDataTable("SELECT DISTINCT * FROM ( " & _
                                                              "SELECT '' AS TrainingName " & _
                                                              "UNION ALL " & _
                                                              "SELECT TrainingName " & _
                                                                "FROM ExpiredTraining) ORDER BY TrainingName ASC")



                ctl.DataSource = dt
                ctl.DisplayMember = "TrainingName"
                ctl.ValueMember = "TrainingName"

            Case "ComboBox2"

                If ctl.SelectedValue <> "" Then Exit Sub
                Dim Fielder As String = "Site"

                Dim dt As DataTable = OverClass.TempDataTable( _
                "SELECT Site FROM Staff GROUP BY Site")

                Dim QueryResult = ((From a In dt.AsEnumerable() _
                                    Select "").Union _
                                    (From a In dt.AsEnumerable()
                                     Where a.Field(Of String)(Fielder) <> ""
                                     Where a.Field(Of String)(Fielder) <> Nothing
                                        Where a.Field(Of String)(Fielder) <> " "
                                     Order By a.Field(Of String)(Fielder) Ascending
                                    Select a.Field(Of String)(Fielder))).Distinct()

                Dim dt2 As New DataTable
                dt2.Columns.Add(Fielder)
                For Each row In QueryResult
                    dt2.Rows.Add(row)
                Next

                ctl.DataSource = dt2
                ctl.DisplayMember = Fielder
                ctl.ValueMember = Fielder

            Case "ComboBox3"

                If ctl.SelectedValue <> "" Then Exit Sub
                Dim Fielder As String = "Role"

                Dim dt As DataTable = OverClass.TempDataTable( _
                "SELECT Role FROM Staff GROUP BY Role")

                Dim QueryResult = ((From a In dt.AsEnumerable() _
                                    Select "").Union _
                                    (From a In dt.AsEnumerable()
                                     Where a.Field(Of String)(Fielder) <> ""
                                     Where a.Field(Of String)(Fielder) <> Nothing
                                        Where a.Field(Of String)(Fielder) <> " "
                                     Order By a.Field(Of String)(Fielder) Ascending
                                    Select a.Field(Of String)(Fielder))).Distinct()

                Dim dt2 As New DataTable
                dt2.Columns.Add(Fielder)
                For Each row In QueryResult
                    dt2.Rows.Add(row)
                Next

                ctl.DataSource = dt2
                ctl.DisplayMember = Fielder
                ctl.ValueMember = Fielder

            Case "ComboBox4"

                If ctl.SelectedValue <> "" Then Exit Sub
                Dim Fielder As String = "Contract"

                Dim dt As DataTable = OverClass.TempDataTable( _
                "SELECT Contract FROM Staff GROUP BY Contract")

                Dim QueryResult = ((From a In dt.AsEnumerable() _
                                    Select "").Union _
                                    (From a In dt.AsEnumerable()
                                     Where a.Field(Of String)(Fielder) <> ""
                                     Where a.Field(Of String)(Fielder) <> Nothing
                                        Where a.Field(Of String)(Fielder) <> " "
                                     Order By a.Field(Of String)(Fielder) Ascending
                                    Select a.Field(Of String)(Fielder))).Distinct()

                Dim dt2 As New DataTable
                dt2.Columns.Add(Fielder)
                For Each row In QueryResult
                    dt2.Rows.Add(row)
                Next

                ctl.DataSource = dt2
                ctl.DisplayMember = Fielder
                ctl.ValueMember = Fielder


        End Select

    End Sub


End Module
