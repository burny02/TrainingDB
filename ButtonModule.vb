Module ButtonModule


    Public Sub ButtonSpecifics(sender As Object, e As EventArgs)

        Dim ctl As Object = Nothing

        Select Case sender.name.ToString

            Case "Button1"
                Call Saver(Form1.DataGridView4)

            Case "Button2"
                Call Saver(Form1.DataGridView2)

            Case "Button4"
                Call Saver(Form1.DataGridView1)

            Case "Button5"
                Call Saver(Form1.DataGridView3)

        End Select

    End Sub

End Module
