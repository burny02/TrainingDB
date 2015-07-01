Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Call LockCheck()

        Call LoginCheck()

        Me.Label2.Text = "Developed by David Burnside" & vbNewLine & vbTab & "Version: " & My.Application.Info.Version.ToString()

    End Sub


    Private Sub TabControl1_Selected(sender As Object, e As TabControlEventArgs) Handles TabControl1.Selected

        Dim SQLCode As String = vbNullString
        Dim Bind As BindingSource = BindingSource1
        Dim ctl As Object = Nothing


        Select Case e.TabPageIndex

            Case 1
                SQLCode = "SELECT * FROM STAFF ORDER BY SName ASC"
                CreateDataSet(SQLCode, Bind, Me.DataGridView1)
                Me.DataGridView1.Columns(0).Visible = False

            Case 2
                SQLCode = "SELECT * FROM TrainType ORDER BY TrainingName ASC"
                CreateDataSet(SQLCode, Bind, Me.DataGridView2)
                Me.DataGridView2.Columns(0).Visible = False

        End Select

    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs)

        Call UpdateBackend()

    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing

        If UnloadData() = True Then e.Cancel = True

    End Sub

    Private Sub DataGridView2_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DataGridView2.DataError

        e.Cancel = False
        Call ErrorHandler(sender, e)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Call UpdateBackend()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Call UpdateBackend()

    End Sub

    Private Sub DataGridView1_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError

        e.Cancel = False
        Call ErrorHandler(sender, e)

    End Sub
End Class
