Public Class FirstRunProgress

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click

    End Sub

    Private Sub ProgressBar1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProgressBar1.Click

    End Sub

    Private Sub FirstRunProgress_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Paint
        ProgressBar1.Minimum = 0
        ProgressBar1.Maximum = ThisAddIn.allFolders
        ProgressBar1.Value = ThisAddIn.atFolderCount
        Label2.Text = "Now processing the folder """ & ThisAddIn.currentFolder & """ (" & ThisAddIn.atFolderCount & "/" & ThisAddIn.allFolders & ")."
    End Sub
End Class