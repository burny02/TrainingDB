Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Call LockCheck()

        Call LoginCheck()

        Me.Label2.Text = "Developed by David Burnside" & vbNewLine & vbTab & "Version: " & My.Application.Info.Version.ToString()

        Me.Text = SolutionName

    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing

        If UnloadData() = True Then e.Cancel = True
        Call Quitter(True)

    End Sub

    Private Sub DataGridView2_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DataGridView2.DataError

        e.Cancel = False
        Call ErrorHandler(sender, e)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Call UpdateBackend(Me.DataGridView2)

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Call UpdateBackend(Me.DataGridView1)

    End Sub

    Private Sub DataGridView1_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError

        e.Cancel = False
        Call ErrorHandler(sender, e)

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Call UpdateBackend(Me.DataGridView3)

    End Sub

    Private Sub TabControl1_Selecting(sender As Object, e As TabControlCancelEventArgs) Handles TabControl1.Selecting

        Dim SQLCode As String = vbNullString
        Dim Bind As BindingSource = BindingSource1
        Dim ctl As Object = Nothing

        If UnloadData() = True Then
            e.Cancel = True
            Exit Sub
        End If

        Call ResetDataGrid()

        Select Case e.TabPageIndex

            Case 1
                ctl = Me.DataGridView1
                SQLCode = "SELECT ID, FName, SName FROM STAFF ORDER BY SName ASC"
                CreateDataSet(SQLCode, Bind, ctl)

            Case 2
                ctl = Me.DataGridView2
                SQLCode = "SELECT ID, TrainingName, ValidLength FROM TrainType ORDER BY TrainingName ASC"
                CreateDataSet(SQLCode, Bind, ctl)


            Case 3
                ctl = Me.DataGridView3
                SQLCode = "SELECT ID, TypeID, CourseDate FROM TrainingCourse  ORDER BY CourseDate ASC"
                CreateDataSet(SQLCode, Bind, ctl)

        End Select


        Call Specifics(ctl)

    End Sub

    Private Sub DataGridView3_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DataGridView3.DataError

        e.Cancel = False
        Call ErrorHandler(sender, e)

    End Sub

    Private Sub ResetDataGrid()

        Me.DataGridView1.Columns.Clear()
        Me.DataGridView2.Columns.Clear()
        Me.DataGridView3.Columns.Clear()
        Me.DataGridView1.DataSource = Nothing
        Me.DataGridView2.DataSource = Nothing
        Me.DataGridView3.DataSource = Nothing

    End Sub

    Private Sub Specifics(ctl As Object)

        If IsNothing(ctl) Then Exit Sub

        Select Case ctl.name

            Case "DataGridView1"
                ctl.Columns(0).Visible = False
                ctl.columns(1).headertext = "Name"
                ctl.columns(2).headertext = "Surname"

            Case "DataGridView2"
                ctl.Columns(0).Visible = False
                ctl.columns(1).headertext = "Training Name"
                ctl.columns(2).headertext = "Valid Length (Months)"

            Case "DataGridView3"
                ctl.Columns(0).Visible = False
                ctl.columns(1).visible = False
                ctl.columns(2).DefaultCellStyle.Format = "dd-MMM-yyyy"
                Dim cmb As New DataGridViewComboBoxColumn()
                cmb.DataSource = TempDataSet("SELECT ID, TrainingName FROM TrainType ORDER BY TrainingName ASC").Tables(0)
                cmb.DataPropertyName = CurrentDataSet.Tables(0).Columns(1).ToString
                cmb.ValueMember = "ID"
                cmb.DisplayMember = "TrainingName"
                ctl.Columns.Add(cmb)
                ctl.columns(2).headertext = "Course Date"
                ctl.columns(3).headertext = "Training Name"

        End Select

    End Sub
End Class
