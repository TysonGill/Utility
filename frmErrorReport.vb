''' <summary>Display a common form for reporting unanticipated errors.</summary>
Friend Class frmErrorReport

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCopy.Click
        System.Windows.Forms.Clipboard.SetText(txtDesc.Text)
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
End Class
