Module SaveModule

    Public Sub Saver(ctl As Object)

        'Get a generic command list first - Ignore errors (Multi table)
        Dim cb As New OleDb.OleDbCommandBuilder(Central.CurrentDataAdapter)

        Try
            Central.CurrentDataAdapter.UpdateCommand = cb.GetUpdateCommand()
        Catch
        End Try
        Try
            Central.CurrentDataAdapter.InsertCommand = cb.GetInsertCommand()
        Catch
        End Try
        Try
            Central.CurrentDataAdapter.DeleteCommand = cb.GetDeleteCommand()
        Catch
        End Try


        'Create and overwrite a custom one if needed (More than 1 table) ...OLEDB Parameters must be added in the order they are used
        Select Case ctl.name

            Case "DataGridView4"

                Dim PKey As Double = Form1.TextBox1.Tag
                'SET THE Commands, with Parameters (OLDB Parameters must be added in the order they are used in the statement)
                Central.CurrentDataAdapter.InsertCommand = New OleDb.OleDbCommand("INSERT INTO CourseAttendees (CourseID, StaffID) " & _
                                                                          "VALUES (" & PKey & ", @P1)")


                'Add parameters with the source columns in the dataset
                With Central.CurrentDataAdapter.InsertCommand.Parameters
                    .Add("@P1", OleDb.OleDbType.VarChar, 255, "StaffID")
                End With


        End Select


        Call Central.SetCommandConnection()
        Call Central.UpdateBackend(ctl)


    End Sub



End Module
