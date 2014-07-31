Public Class Details
    Private Sub Details_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Paint
        SizeLabel.Text = "Total usage is " & QuotaTool.NumberUsage & " out of " _
            & QuotaTool.Quota / 1000000 & " MB, or " & QuotaTool.PercentageQuota & " %."
    End Sub
End Class