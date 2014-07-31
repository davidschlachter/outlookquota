Public Class Progress

    Private Sub FirstRunProgress_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Paint
        Label1.Text = QuotaTool.theMessage
        ProgressBar1.Minimum = 0
        ProgressBar1.Maximum = QuotaTool.allFolders
        ProgressBar1.Value = QuotaTool.atFolder
        Label2.Text = "Now processing the folder """ & QuotaTool.currentFolder & """ (" & QuotaTool.atFolder & "/" & QuotaTool.allFolders & ")."
    End Sub

End Class