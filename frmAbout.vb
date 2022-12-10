Imports System.Windows.Forms

''' <summary>
''' A generic about box. Displayed information is taken from the project assembly information settings.
''' </summary>
Public NotInheritable Class frmAbout

    Private Sub frmAbout_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Cursor.Current = Cursors.WaitCursor
            Dim ApplicationTitle As String
            If My.Application.Info.Title <> "" Then
                ApplicationTitle = My.Application.Info.Title
            Else
                ApplicationTitle = System.IO.Path.GetFileNameWithoutExtension(My.Application.Info.AssemblyName)
            End If
            Me.Text = String.Format("About {0}", ApplicationTitle)
            lblProduct.Text = My.Application.Info.ProductName
            lblVersion.Text = String.Format("Version {0}", My.Application.Info.Version.ToString)
            lblCopyright.Text = My.Application.Info.Copyright
            lblCompany.Text = My.Application.Info.CompanyName
            lblDescription.Text = My.Application.Info.Description
        Catch ex As Exception
            Utility.ReportError(ex)
        Finally
            Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Me.Close()
    End Sub

    Private Sub lblCompany_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblCompany.Click

    End Sub
End Class
