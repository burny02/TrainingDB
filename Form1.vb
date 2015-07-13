Public Class Form1

    Private colRemovedTabs As New Collection()

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.WindowState = FormWindowState.Maximized

        Call LockCheck()

        Call LoginCheck()

        Try
            Me.Label2.Text = "Training Tool " & vbNewLine & "Developed by David Burnside" & vbNewLine & "Version: " & System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString
        Catch ex As Exception
            Me.Label2.Text = "Training Tool " & vbNewLine & "Developed by David Burnside"
        End Try

        Me.Text = SolutionName

        colRemovedTabs.Add(Me.TabPage5, Me.TabPage5.Name)
        TabControl1.Controls.Remove(Me.TabPage5)


    End Sub



    Private Sub DataGridView2_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DataGridView2.DataError

        e.Cancel = False
        Call ErrorHandler(sender, e)

    End Sub

    Private Sub DataGridView1_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError

        e.Cancel = False
        Call ErrorHandler(sender, e)

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
                TabControl1.Controls.Remove(Me.TabPage5)

            Case 2
                ctl = Me.DataGridView2
                SQLCode = "SELECT ID, TrainingName, ValidLength FROM TrainType ORDER BY TrainingName ASC"
                CreateDataSet(SQLCode, Bind, ctl)
                TabControl1.Controls.Remove(Me.TabPage5)


            Case 3
                ctl = Me.DataGridView3
                SQLCode = "SELECT ID, TypeID, CourseDate FROM TrainingCourse  ORDER BY CourseDate ASC"
                CreateDataSet(SQLCode, Bind, ctl)
                TabControl1.Controls.Remove(Me.TabPage5)

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
                Dim cmb2 As New DataGridViewImageColumn()
                cmb2.Image = My.Resources.social_networking_users_icon
                cmb2.ImageLayout = DataGridViewImageCellLayout.Zoom
                cmb2.HeaderText = "Attendees"
                ctl.columns.add(cmb2)
        End Select

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        Call UpdateBackend(Me.DataGridView1)

    End Sub


    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Call UpdateBackend(Me.DataGridView3)
    End Sub

    Private Sub DataGridView3_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellClick
        If Not sender.columns(e.ColumnIndex).HeaderText = "Attendees" Then Exit Sub
        If IsNothing(sender.rows(e.RowIndex).cells("ID").value) _
        Or IsDBNull(sender.rows(e.RowIndex).cells("ID").value) _
        Or sender.CurrentRow.IsNewRow Then Exit Sub

        If Not IsNothing(CurrentBindingSource) Then CurrentBindingSource.EndEdit()

        If UnloadData() = False Then
            If QueryTest("SELECT ID FROM TrainingCourse WHERE ID=" & sender.rows(e.RowIndex).cells("ID").value) = 0 Then
                MsgBox("No Record found")
            Else
                TabControl1.Controls.Add(colRemovedTabs("TabPage5"))
                TabControl1.SelectedTab = TabPage5
                Call ResetDataGrid()

            End If
        End If




    End Sub
End Class
