Public Class Form1

    Private colRemovedTabs As New Collection()

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.WindowState = FormWindowState.Maximized

        Call StartUpCentral()

        Central.LockCheck()

        Central.LoginCheck()

        Try
            Me.Label2.Text = SolutionName & vbNewLine & "Developed by David Burnside" & vbNewLine & "Version: " & System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString
        Catch
            Me.Label2.Text = SolutionName & vbNewLine & "Developed by David Burnside"
        End Try

        Me.Text = SolutionName

        colRemovedTabs.Add(Me.TabPage5, Me.TabPage5.Name)
        TabControl1.Controls.Remove(Me.TabPage5)


    End Sub

    Private Sub DataGridView2_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DataGridView2.DataError

        e.Cancel = False
        Call Central.ErrorHandler(sender, e)

    End Sub

    Private Sub DataGridView1_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError

        e.Cancel = False
        Call Central.ErrorHandler(sender, e)

    End Sub

    Private Sub TabControl1_Selecting(sender As Object, e As TabControlCancelEventArgs) Handles TabControl1.Selecting

        Dim SQLCode As String = vbNullString
        Dim Bind As BindingSource = BindingSource1
        Dim ctl As Object = Nothing

        If Central.UnloadData() = True Then
            e.Cancel = True
            Exit Sub
        End If

        Call ResetDataGrid()

        Select Case e.TabPageIndex

            Case 1
                ctl = Me.DataGridView1
                SQLCode = "SELECT ID, FName, SName FROM STAFF ORDER BY SName ASC"
                Central.CreateDataSet(SQLCode, Bind, ctl)
                TabControl1.Controls.Remove(Me.TabPage5)

            Case 2
                ctl = Me.DataGridView2
                SQLCode = "SELECT ID, TrainingName, ValidLength FROM TrainType ORDER BY TrainingName ASC"
                Central.CreateDataSet(SQLCode, Bind, ctl)
                TabControl1.Controls.Remove(Me.TabPage5)


            Case 3
                ctl = Me.DataGridView3
                SQLCode = "SELECT ID, TypeID, CourseDate FROM TrainingCourse  ORDER BY CourseDate ASC"
                Central.CreateDataSet(SQLCode, Bind, ctl)
                TabControl1.Controls.Remove(Me.TabPage5)

        End Select


        Call Specifics(ctl)

    End Sub

    Private Sub DataGridView3_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DataGridView3.DataError

        e.Cancel = False
        Call Central.ErrorHandler(sender, e)

    End Sub

    Private Sub ResetDataGrid()

        Me.DataGridView1.Columns.Clear()
        Me.DataGridView2.Columns.Clear()
        Me.DataGridView3.Columns.Clear()
        Me.DataGridView4.Columns.Clear()
        Me.TextBox1.DataBindings.Clear()
        Me.TextBox2.DataBindings.Clear()
        Me.DataGridView1.DataSource = Nothing
        Me.DataGridView2.DataSource = Nothing
        Me.DataGridView3.DataSource = Nothing
        Me.DataGridView4.DataSource = Nothing


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
                cmb.DataSource = Central.TempDataSet("SELECT ID, TrainingName FROM TrainType ORDER BY TrainingName ASC").Tables(0)
                cmb.DataPropertyName = Central.CurrentDataSet.Tables(0).Columns(1).ToString
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

        Call Saver(DataGridView1)

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Call Saver(DataGridView3)
    End Sub

    Private Sub DataGridView3_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellClick
        If Not sender.columns(e.ColumnIndex).HeaderText = "Attendees" Then Exit Sub
        If IsNothing(sender.rows(e.RowIndex).cells("ID").value) _
        Or IsDBNull(sender.rows(e.RowIndex).cells("ID").value) _
        Or sender.CurrentRow.IsNewRow Then Exit Sub

        If Not IsNothing(Central.CurrentBindingSource) Then Central.CurrentBindingSource.EndEdit()



        If Central.UnloadData() = False Then
            Dim CourseID As Double
            CourseID = sender.rows(e.RowIndex).cells("ID").value

            If Central.QueryTest("SELECT ID FROM TrainingCourse WHERE ID=" & CourseID) = 0 Then
                MsgBox("No Record found")
            Else
                Dim SQLCode As String
                SQLCode = "SELECT ID, CourseID, StaffID " & _
                "FROM CourseAttendees " & _
                "WHERE CourseID=" & CourseID

                TabControl1.Controls.Add(colRemovedTabs("TabPage5"))
                TabControl1.SelectedTab = TabPage5
                Call ResetDataGrid()
                Central.CreateDataSet(SQLCode, BindingSource1, Me.DataGridView4)
                Me.TextBox1.DataBindings.Add("Text", Central.TempDataSet("SELECT TrainingName FROM TrainType a " & _
                                                                 "INNER JOIN TrainingCourse b ON a.ID=b.TypeID " & _
                                                                 "WHERE b.ID=" & CourseID).Tables(0), "TrainingName")
                Me.TextBox1.DataBindings.Add("Tag", Central.TempDataSet("SELECT ID FROM TrainingCourse " & _
                                                                "WHERE ID=" & CourseID).Tables(0), "ID")
                Me.TextBox2.DataBindings.Add("Text", Central.TempDataSet("SELECT format(CourseDate,'dd-mmm-yyyy') AS Dater FROM TrainingCourse " & _
                                                                  "WHERE ID=" & CourseID).Tables(0), "Dater")
                Central.CurrentDataSet.Tables(0).Columns("CourseID").DefaultValue = CourseID
                Dim cmb As New DataGridViewComboBoxColumn()
                cmb.DataSource = Central.TempDataSet("SELECT ID, FName & ' ' & SName As FullName FROM STAFF ORDER BY FName & ' ' & SName ASC").Tables(0)
                cmb.DataPropertyName = Central.CurrentDataSet.Tables(0).Columns("StaffID").ToString
                cmb.ValueMember = "ID"
                cmb.DisplayMember = "FullName"
                Me.DataGridView4.Columns(0).Visible = False
                Me.DataGridView4.Columns(1).Visible = False
                Me.DataGridView4.Columns(2).Visible = False
                Me.DataGridView4.Columns.Add(cmb)


            End If

        End If


    End Sub

    Private Sub DataGridView4_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DataGridView4.DataError
        e.Cancel = False
        Call Central.ErrorHandler(sender, e)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call Saver(DataGridView4)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Call Saver(DataGridView2)
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If Central.UnloadData() = True Then e.Cancel = True
        Call Central.Quitter(True)
    End Sub

    Private Sub DataGridView3_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellEnter
        Call Central.SingleClick(sender, e)
    End Sub

    Private Sub DataGridView4_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView4.CellEnter
        Call Central.SingleClick(sender, e)
    End Sub
End Class
