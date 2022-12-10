''' <summary>A general use form to select an item from a list.</summary>
''' <remarks> Use ui.SelectFromList to display.</remarks>
Public Class frmListSelector

Public ListData As DataTable
Public ListField As String
Public ReturnedField As String
Public Selection As String = ""

Private Sub frmListSelector_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    For i As Integer = 0 To ListData.Rows.Count - 1
        lst.Items.Add(ListData.Rows(i)(ListField).ToString)
    Next
    If lst.Items.Count > 0 Then lst.SelectedIndex = 0
    btnOK.Enabled = (lst.Items.Count > 0)
End Sub

Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
    Selection = ListData.Rows(lst.SelectedIndex)(ReturnedField).ToString
    Hide()
End Sub

Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
    Hide()
End Sub

Private Sub lst_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lst.DoubleClick
    Selection = ListData.Rows(lst.SelectedIndex)(ReturnedField).ToString
    Hide()
End Sub

End Class