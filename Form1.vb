Public Class Form1

    Private colRemovedTabs As New Collection()

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.WindowState = FormWindowState.Maximized

        Call StartUpCentral()

        Try
            Me.Label2.Text = SolutionName & vbNewLine & "Developed by David Burnside" & vbNewLine & "Version: " & System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString
        Catch
            Me.Label2.Text = SolutionName & vbNewLine & "Developed by David Burnside"
        End Try

        Me.Text = SolutionName

        colRemovedTabs.Add(Me.TabPage5, Me.TabPage5.Name)
        TabControl1.Controls.Remove(Me.TabPage5)


    End Sub

    Private Sub TabControl1_Selecting(sender As Object, e As TabControlCancelEventArgs) Handles TabControl1.Selecting

        Dim SQLCode As String = vbNullString
        Dim Bind As BindingSource = BindingSource1
        Dim ctl As Object = Nothing

        If OverClass.UnloadData() = True Then
            e.Cancel = True
            Exit Sub
        End If

        OverClass.ResetCollection()

        Select Case e.TabPageIndex

            Case 1
                ctl = Me.DataGridView1
                SQLCode = "SELECT ID, FName, SName FROM STAFF ORDER BY SName ASC"
                OverClass.CreateDataSet(SQLCode, Bind, ctl)
                TabControl1.Controls.Remove(Me.TabPage5)

            Case 2
                ctl = Me.DataGridView2
                SQLCode = "SELECT ID, TrainingName, ValidLength FROM TrainType ORDER BY TrainingName ASC"
                OverClass.CreateDataSet(SQLCode, Bind, ctl)
                TabControl1.Controls.Remove(Me.TabPage5)


            Case 3
                ctl = Me.DataGridView3
                SQLCode = "SELECT ID, TypeID, CourseDate FROM TrainingCourse  ORDER BY CourseDate ASC"
                OverClass.CreateDataSet(SQLCode, Bind, ctl)
                TabControl1.Controls.Remove(Me.TabPage5)

        End Select


        Call Specifics(ctl)

    End Sub

    Private Sub Specifics(ctl As Object)

        If IsNothing(ctl) Then Exit Sub

        Select Case ctl.name

            Case "DataGridView1"
                ctl.Columns(0).Visible = False
                ctl.columns(1).headertext = "Name"
                ctl.columns(2).headertext = "Surname"
                Dim cmb As New DataGridViewImageColumn
                cmb.DisplayIndex = 10
                cmb.HeaderText = "View Training"
                cmb.Image = My.Resources.training
                cmb.ImageLayout = DataGridViewImageCellLayout.Zoom
                ctl.columns.add(cmb)
                cmb.Name = "ViewTraining"

            Case "DataGridView2"
                ctl.Columns(0).Visible = False
                ctl.columns(1).headertext = "Training Name"
                ctl.columns(2).headertext = "Valid Length (Months)"

            Case "DataGridView3"
                ctl.Columns(0).Visible = False
                ctl.columns(1).visible = False
                ctl.columns(2).DefaultCellStyle.Format = "dd-MMM-yyyy"
                Dim cmb As New DataGridViewComboBoxColumn()
                cmb.DataSource = OverClass.TempDataTable("SELECT ID, TrainingName FROM TrainType ORDER BY TrainingName ASC")
                cmb.DataPropertyName = OverClass.CurrentDataSet.Tables(0).Columns(1).ToString
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

    Private Sub DataGridView3_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellClick
        If Not sender.columns(e.ColumnIndex).HeaderText = "Attendees" Then Exit Sub
        If IsNothing(sender.rows(e.RowIndex).cells("ID").value) _
        Or IsDBNull(sender.rows(e.RowIndex).cells("ID").value) _
        Or sender.CurrentRow.IsNewRow Then Exit Sub

        If Not IsNothing(OverClass.CurrentBindingSource) Then OverClass.CurrentBindingSource.EndEdit()



        If OverClass.UnloadData() = False Then
            Dim CourseID As Double
            CourseID = sender.rows(e.RowIndex).cells("ID").value

            If OverClass.QueryTest("SELECT ID FROM TrainingCourse WHERE ID=" & CourseID) = 0 Then
                MsgBox("No Record found")
            Else
                Dim SQLCode As String
                SQLCode = "SELECT ID, CourseID, StaffID " & _
                "FROM CourseAttendees " & _
                "WHERE CourseID=" & CourseID

                TabControl1.Controls.Add(colRemovedTabs("TabPage5"))
                TabControl1.SelectedTab = TabPage5
                OverClass.ResetCollection()
                OverClass.CreateDataSet(SQLCode, BindingSource1, Me.DataGridView4)
                Me.TextBox1.DataBindings.Clear()
                Me.TextBox2.DataBindings.Clear()
                Me.TextBox1.DataBindings.Add("Text", OverClass.TempDataTable("SELECT TrainingName FROM TrainType a " & _
                                                                 "INNER JOIN TrainingCourse b ON a.ID=b.TypeID " & _
                                                                 "WHERE b.ID=" & CourseID), "TrainingName")
                Me.TextBox1.DataBindings.Add("Tag", OverClass.TempDataTable("SELECT ID FROM TrainingCourse " & _
                                                                "WHERE ID=" & CourseID), "ID")
                Me.TextBox2.DataBindings.Add("Text", OverClass.TempDataTable("SELECT format(CourseDate,'dd-mmm-yyyy') AS Dater FROM TrainingCourse " & _
                                                                  "WHERE ID=" & CourseID), "Dater")
                OverClass.CurrentDataSet.Tables(0).Columns("CourseID").DefaultValue = CourseID
                Dim cmb As New DataGridViewComboBoxColumn()
                cmb.DataSource = OverClass.TempDataTable("SELECT ID, FName & ' ' & SName As FullName FROM STAFF ORDER BY FName & ' ' & SName ASC")
                cmb.DataPropertyName = OverClass.CurrentDataSet.Tables(0).Columns("StaffID").ToString
                cmb.ValueMember = "ID"
                cmb.DisplayMember = "FullName"
                Me.DataGridView4.Columns(0).Visible = False
                Me.DataGridView4.Columns(1).Visible = False
                Me.DataGridView4.Columns(2).Visible = False
                Me.DataGridView4.Columns.Add(cmb)


            End If

        End If


    End Sub


    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

        If e.ColumnIndex <> sender.columns("ViewTraining").index Then Exit Sub
        If IsDBNull(sender.item("ID", e.RowIndex).value) Then Exit Sub

        Dim Dt As DataTable = OverClass.TempDataTable("SELECT CourseDate, TrainingName, DateAdd('M',iif(isnull(ValidLength),120,ValidLength),CourseDate) As Expiry " & _
                                                      "FROM ((CourseAttendees a INNER JOIN TrainingCourse b " & _
                                                        "ON a.CourseID=b.ID) INNER JOIN TrainType c ON b.TypeID=c.ID) " & _
                                                        "WHERE StaffID=" & sender.item("ID", e.RowIndex).value & _
                                                        " ORDER BY CourseDate ASC")
        If Dt.Rows.Count = 0 Then
            MsgBox("No training found")
        Else
            Dim Viewer As New StaffTraining
            Viewer.DataGridView1.DataSource = Dt
            Viewer.DataGridView1.AllowUserToAddRows = False
            Viewer.DataGridView1.ReadOnly = True
            Viewer.DataGridView1.Columns("CourseDate").DefaultCellStyle.Format = "dd-MMM-yyyy"
            Viewer.DataGridView1.Columns("Expiry").DefaultCellStyle.Format = "dd-MMM-yyyy"
            Viewer.Text = sender.item("FName", e.RowIndex).value & " " & sender.item("SName", e.RowIndex).value
            Viewer.Visible = True
        End If
        

    End Sub
End Class
