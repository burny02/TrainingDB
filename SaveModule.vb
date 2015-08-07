Module SaveModule

    Public Sub Saver(ctl As Object)

        'Get a generic command list first - Ignore errors (Multi table)
        Dim cb As New OleDb.OleDbCommandBuilder(OverClass.CurrentDataAdapter)

        Try
            OverClass.CurrentDataAdapter.UpdateCommand = cb.GetUpdateCommand()
        Catch
        End Try
        Try
            OverClass.CurrentDataAdapter.InsertCommand = cb.GetInsertCommand()
        Catch
        End Try
        Try
            OverClass.CurrentDataAdapter.DeleteCommand = cb.GetDeleteCommand()
        Catch
        End Try


        'Create and overwrite a custom one if needed (More than 1 table) ...OLEDB Parameters must be added in the order they are used
        Select Case ctl.name

            Case "DataGridView4"

                Dim PKey As Double = Form1.TextBox1.Tag
                'SET THE Commands, with Parameters (OLDB Parameters must be added in the order they are used in the statement)
                OverClass.CurrentDataAdapter.InsertCommand = New OleDb.OleDbCommand("INSERT INTO CourseAttendees (CourseID, StaffID) " & _
                                                                          "VALUES (" & PKey & ", @P1)")


                'Add parameters with the source columns in the dataset
                With OverClass.CurrentDataAdapter.InsertCommand.Parameters
                    .Add("@P1", OleDb.OleDbType.VarChar, 255, "StaffID")
                End With


        End Select


        Call OverClass.SetCommandConnection()
        Call OverClass.UpdateBackend(ctl)


    End Sub



End Module
