Public Class Form1

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click


        Dim myForm As PackMain
        myForm = New PackMain

        myForm.Show()


    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click

        Dim myForm As ProPackMain
        myForm = New ProPackMain

        myForm.Show()
    End Sub
End Class