Imports Microsoft.Reporting.WinForms

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


        Me.ReportViewer1.RefreshReport()
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
                SQLCode = "SELECT ID, FName, SName, Role, Site, Hidden, Contract FROM STAFF WHERE Hidden = False ORDER BY FName ASC"
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

            Case 4

                Call Specifics(Me.ReportViewer1)
                StartCombo(Me.ComboBox1)
                StartCombo(Me.ComboBox2)
                StartCombo(Me.ComboBox3)
                StartCombo(Me.ComboBox4)
                StartCombo(Me.ComboBox5)

        End Select


        Call Specifics(ctl)

    End Sub

    Public Sub Specifics(ctl As Object)

        If IsNothing(ctl) Then Exit Sub

        Select Case ctl.name

            Case "DataGridView1"
                ctl.Columns(0).Visible = False
                ctl.columns("Hidden").visible = False
                ctl.columns("Site").visible = False
                ctl.columns("Contract").visible = False
                ctl.columns(1).headertext = "Name"
                ctl.columns(2).headertext = "Surname"
                Dim cmb3 As New DataGridViewComboBoxColumn()
                cmb3.DataPropertyName = "Site"
                cmb3.Items.Add("MAN")
                cmb3.Items.Add("WHC")
                cmb3.Items.Add("QUA")
                cmb3.HeaderText = "Site"
                ctl.Columns.Add(cmb3)
                Dim cmb4 As New DataGridViewComboBoxColumn()
                cmb4.DataPropertyName = "Contract"
                cmb4.Items.Add("Permanent")
                cmb4.Items.Add("Bank")
                cmb4.HeaderText = "Contract Type"
                ctl.Columns.Add(cmb4)
                Dim cmb As New DataGridViewImageColumn
                cmb.DisplayIndex = 10
                cmb.HeaderText = "View Training"
                cmb.Image = My.Resources.training
                cmb.ImageLayout = DataGridViewImageCellLayout.Zoom
                ctl.columns.add(cmb)
                cmb.Name = "ViewTraining"
                Dim cmb2 As New DataGridViewImageColumn
                cmb2.HeaderText = "Hide Staff Member"
                cmb2.Image = My.Resources.hide
                cmb2.ImageLayout = DataGridViewImageCellLayout.Zoom
                ctl.columns.add(cmb2)
                cmb2.Name = "HideStaff"

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

            Case "ReportViewer1"

                Dim SiteCrit As String = "'%' OR Site IS NULL OR Site='' "
                Dim MonthCrit As Date = Date.Now
                Dim ContCrit As String = "'%' OR Contract IS NULL OR Contract='' "
                Dim RoleCrit As String = "'%' OR Role IS NULL Or Role='' "
                Dim CourseCrit As String = "'%'"
                Dim MonthAdd As Integer = 0

                If Me.ComboBox2.SelectedValue <> "" Then SiteCrit = "'" & Me.ComboBox2.SelectedValue & "'"
                If Me.ComboBox1.SelectedItem <> "" Then
                    MonthAdd = Me.ComboBox1.SelectedItem
                    MonthCrit = DateAdd(DateInterval.Month, MonthAdd, MonthCrit)
                End If
                If Me.ComboBox3.SelectedValue <> "" Then ContCrit = "'" & Me.ComboBox3.SelectedValue & "'"
                If Me.ComboBox4.SelectedValue <> "" Then RoleCrit = "'" & Me.ComboBox4.SelectedValue & "'"
                If Me.ComboBox5.SelectedValue <> "" Then CourseCrit = "'" & Me.ComboBox5.SelectedValue & "'"

                Dim SQLCode As String =
                "SELECT FullName, Expires, TrainingName FROM ExpiredTraining " &
                "WHERE (Site LIKE " & SiteCrit & ") AND " &
                "(Contract LIKE " & ContCrit & ") AND " &
                "(Role LIKE " & RoleCrit & ") AND " &
                "(TrainingName LIKE " & CourseCrit & ") AND " &
                "(Expires <=" & OverClass.SQLDate(MonthCrit) & ")"

                Me.ReportViewer1.Visible = True
                Me.ReportViewer1.LocalReport.DataSources.Clear()
                Me.ReportViewer1.LocalReport.ReportEmbeddedResource = "TrainingDB.ExpiredTraining.rdlc"
                Me.ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("ReportDataSet", _
                                                          OverClass.TempDataTable(SQLCode)))
                Me.ReportViewer1.RefreshReport()


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

            If OverClass.SELECTCount("SELECT ID FROM TrainingCourse WHERE ID=" & CourseID) = 0 Then
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
                cmb.DataSource = OverClass.TempDataTable("SELECT ID, FName & ' ' & SName As FullName FROM STAFF WHERE Hidden=False ORDER BY FName & ' ' & SName ASC")
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





        If e.ColumnIndex = sender.columns("ViewTraining").index Then

            If IsDBNull(sender.item("ID", e.RowIndex).value) Then
                MsgBox("Please save record first")
                Exit Sub
            End If


            Dim Dt As DataTable = OverClass.TempDataTable("SELECT CourseDate, TrainingName, DateAdd('M',iif(isnull(ValidLength),120,ValidLength),CourseDate) As Expiry " &
                                                          "FROM ((CourseAttendees a INNER JOIN TrainingCourse b " &
                                                            "ON a.CourseID=b.ID) INNER JOIN TrainType c ON b.TypeID=c.ID) " &
                                                            "WHERE StaffID=" & sender.item("ID", e.RowIndex).value &
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

        End If

        If e.ColumnIndex = sender.columns("HideStaff").index Then

            If IsDBNull(sender.item("ID", e.RowIndex).value) Then
                MsgBox("Please save record first")
                Exit Sub
            End If

            If MsgBox("Are you sure you want to hide this staff member?" & vbNewLine &
                      "Please save to commit changes", vbYesNo) = MsgBoxResult.Yes Then

                sender.item("Hidden", e.RowIndex).value = True
                sender.CurrentCell = Nothing
                sender.Rows(e.RowIndex).Visible = False

            End If


        End If
    End Sub

End Class
