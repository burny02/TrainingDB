Public Class StaffTraining

    Private Sub DataGridView1_RowPrePaint(sender As Object, e As DataGridViewRowPrePaintEventArgs) Handles DataGridView1.RowPrePaint
        If Me.DataGridView1.Rows(e.RowIndex).Cells("Expiry").Value < DateAndTime.Today Then
            Me.DataGridView1.Rows(e.RowIndex).Cells("Expiry").Style.BackColor = Color.Red
        Else
            Me.DataGridView1.Rows(e.RowIndex).Cells("Expiry").Style.BackColor = Color.White
        End If
    End Sub
End Class