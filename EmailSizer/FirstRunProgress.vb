Public Class FirstRunProgress

    Private Sub FirstRunProgress_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Paint
        Label1.Text = ThisAddIn.theMessage
        ProgressBar1.Minimum = 0
        ProgressBar1.Maximum = ThisAddIn.allFolders
        ProgressBar1.Value = ThisAddIn.atFolderCount
        Label2.Text = "Now processing the folder """ & ThisAddIn.currentFolder & """ (" & ThisAddIn.atFolderCount & "/" & ThisAddIn.allFolders & ")."
    End Sub

End Class