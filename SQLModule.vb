'Imports needed for OLEDB connections (Access backend)
Imports System.IO
Imports System.Data
Imports System.Data.OleDb.OleDbConnection
Module SQLModule
    'Store the current dataset, adapter and binding source so that pubically accessable (Save operations etc) - Also then 1 main dataset per form view
    Public CurrentDataSet As DataSet = Nothing
    Public CurrentDataAdapter As OleDb.OleDbDataAdapter = Nothing
    Public CurrentBindingSource As BindingSource = Nothing
    'Connection information privately accessible 
    Private Const TablePath As String = "M:\VOLUNTEER SCREENING SERVICES\DavidBurnside\\Training\Backend.accdb"
    Private Const PWord As String = "Crypto*Dave02"
    Private Const Connect As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & TablePath & ";Jet OLEDB:Database Password=" & PWord
    Public Const SolutionName As String = "Training Tool"

    Public Function QueryTest(SQLCode As String) As Long
        'Execute a SQL Command and return the number of records

        Dim Counter As Long
        Dim dt As New DataTable
        Dim da As New OleDb.OleDbDataAdapter(SQLCode, Connect)

        Try
            'Connect
            da.Fill(dt)
            'Assign
            Counter = dt.Rows.Count

        Catch ex As Exception
            MsgBox(ex.Message)

        Finally
            'Close Off & Clean up
            dt = Nothing
            da = Nothing

        End Try
        QueryTest = Counter

    End Function

    Public Sub ExecuteSQL(SQLCode As String)
        'Execute a SQL Command - No return

        'Create connection & Command
        Dim con As New OleDb.OleDbConnection(Connect)
        Dim cmd As New OleDb.OleDbCommand(SQLCode, con)

        Try
            'Open connection 
            con.Open()
            'Execute SQL Command
            cmd.ExecuteNonQuery()

        Catch ex As Exception
            MsgBox(ex.Message)

        Finally

            'Close Off & Clean up
            Try
                con.Close()
            Catch ex As Exception
                ex = Nothing
            Finally
                con = Nothing
                cmd = Nothing
            End Try

        End Try

    End Sub

    Public Sub CreateDataSet(SQLCode As String, BindSource As BindingSource, ctl As Object)
        'Create a new dataset, set a bindining source and object to that binding source

        'Create Connection object
        Dim con As New OleDb.OleDbConnection(Connect)

        Try
            'Open connection
            con.Open()
            'Create New Dataset & adapter
            CurrentDataAdapter = New OleDb.OleDbDataAdapter(SQLCode, con)
            CurrentDataSet = New DataSet()
            CurrentBindingSource = BindSource

            'Use adapter to fill dataset
            CurrentDataAdapter.Fill(CurrentDataSet)

            'Set bindsource datasource as dataset, set object datasource as bindsource
            BindSource.DataSource = CurrentDataSet.Tables(0)
            ctl.datasource = BindSource

        Catch ex As Exception
            MsgBox(ex.Message)

        Finally

            'Close off & Clean up
            Try
                con.Close()
            Catch ex As Exception
                ex = Nothing
            Finally
                con = Nothing
            End Try

        End Try

    End Sub

    Public Sub UpdateBackend(ctl As Object)
        'Saving function to update access backend

        'New conecction 
        Dim con As New OleDb.OleDbConnection(Connect)

        'Set INSERT, UPDATE COMMANDS
        Call CustomCommand(ctl, con)


        'Is the data dirty / has errors that have auto-undone
        If CurrentDataSet.HasChanges() = False Then
            MsgBox("Errors present/No changes to upload")
            Exit Sub
        End If


        Try
            'Cancel current edit
            CurrentBindingSource.EndEdit()
            'Open connection
            con.Open()
            'Use dataadapter to update the backend (Commands already set)
            CurrentDataAdapter.Update(CurrentDataSet)
            MsgBox("Table Updated")
            'Remove any error messages & accept changes
            CurrentDataSet.AcceptChanges()

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally

            'Close off & clean up
            Try
                con.Close()
                'Requery
                Call Refresher(ctl)
            Catch ex As Exception
                ex = Nothing
            Finally
                con = Nothing
            End Try

        End Try

    End Sub

    Public Function UnloadData() As Boolean
        'Close down currnt dataset, dataadapter & bindinsource

        'Variable if user wants to save
        Dim Cancel As Boolean = False


        'Is there currently a dataset to close?
        If IsNothing(CurrentDataSet) Then
            UnloadData = False
            Exit Function
        End If

        Try

            'Is the dataset dirty?
            If CurrentDataSet.HasChanges() Then

                'Ask user if they want to proceed and lose data?
                If (MsgBox("Changes to data will be lost unless saved first. Do you wish to discard changes?", vbYesNo) = vbNo) Then Cancel = True

            End If


            'If want to continue, clear all current data items
            If Cancel = False Then
                CurrentDataSet = Nothing
                CurrentDataAdapter = Nothing
                CurrentBindingSource = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            'Pass back whether clean up happened
            UnloadData = Cancel
        End Try

    End Function

    Public Sub CustomCommand(ctl As Object, connection As OleDb.OleDbConnection)
        'Create Custom INSERT, UPDATE COMMANDS for saving dataset (More than 1 table in select)


        Select Case ctl.name

            Case Else

                'If not specified - Select commands with one table can auto generate INSERT, UPDATE commands
                Dim cb As New OleDb.OleDbCommandBuilder(CurrentDataAdapter)

        End Select

    End Sub

    Public Function TempDataSet(SQLCode As String) As DataSet
        'Create a temporary dataset for things such as combo box which arent based on the initial query

        'New connection
        Dim con As New OleDb.OleDbConnection(Connect)

        Try
            'Open connection
            con.Open()
            'New temporary data adapter and dataset
            Dim TempDataAdapter = New OleDb.OleDbDataAdapter(SQLCode, con)
            TempDataSet = New DataSet()
            'Use temp adapter to fill temp dataset
            TempDataAdapter.Fill(TempDataSet)

        Catch ex As Exception
            MsgBox(ex.Message)
            TempDataSet = Nothing
        Finally

            'Close off & Clean up
            Try
                con.Close()
            Catch ex As Exception
                ex = Nothing
            Finally
                con = Nothing
            End Try

        End Try

    End Function

    Public Function CreateCSVString(SQLCode As String) As String

        Dim da As New OleDb.OleDbDataAdapter(SQLCode, Connect)
        Dim dt As New DataTable
        Dim Output As String = vbNullString

        Try 
            da.Fill(dt)

            For Each row In dt.Rows

                If Not IsNothing(row.item(0)) And Not IsDBNull(row.item(0)) Then
                    Output = Output & row.Item(0) & ","
                End If

            Next

            Output = Left(Output, Len(Output) - 1)

        Catch ex As Exception
            MsgBox(ex.Message)

        Finally
            CreateCSVString = Output
            dt = Nothing
            da = Nothing
        End Try

    End Function

    Public Sub Refresher(DataItem As Object)

        Try
            Call CreateDataSet(CurrentDataAdapter.SelectCommand.CommandText, CurrentBindingSource, DataItem)
            DataItem.parent.refresh()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

End Module
