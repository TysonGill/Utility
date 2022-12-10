<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmErrorReport
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.btnCopy = New System.Windows.Forms.Button()
        Me.lblFatal = New System.Windows.Forms.Label()
        Me.txtDesc = New System.Windows.Forms.TextBox()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'btnCopy
        '
        Me.btnCopy.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnCopy.Location = New System.Drawing.Point(12, 293)
        Me.btnCopy.Name = "btnCopy"
        Me.btnCopy.Size = New System.Drawing.Size(67, 23)
        Me.btnCopy.TabIndex = 1
        Me.btnCopy.Text = "Copy"
        '
        'lblFatal
        '
        Me.lblFatal.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblFatal.AutoSize = True
        Me.lblFatal.ForeColor = System.Drawing.Color.Red
        Me.lblFatal.Location = New System.Drawing.Point(231, 298)
        Me.lblFatal.Name = "lblFatal"
        Me.lblFatal.Size = New System.Drawing.Size(245, 13)
        Me.lblFatal.TabIndex = 7
        Me.lblFatal.Text = "The application will terminate after closing this form"
        Me.lblFatal.Visible = False
        '
        'txtDesc
        '
        Me.txtDesc.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDesc.Location = New System.Drawing.Point(0, 0)
        Me.txtDesc.Multiline = True
        Me.txtDesc.Name = "txtDesc"
        Me.txtDesc.ReadOnly = True
        Me.txtDesc.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtDesc.Size = New System.Drawing.Size(561, 273)
        Me.txtDesc.TabIndex = 8
        Me.txtDesc.TabStop = False
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(482, 293)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(67, 23)
        Me.btnClose.TabIndex = 9
        Me.btnClose.Text = "Close"
        '
        'frmErrorReport
        '
        Me.AcceptButton = Me.btnClose
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(561, 333)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.txtDesc)
        Me.Controls.Add(Me.btnCopy)
        Me.Controls.Add(Me.lblFatal)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(435, 304)
        Me.Name = "frmErrorReport"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Application Error Report"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnCopy As System.Windows.Forms.Button
    Friend WithEvents lblFatal As System.Windows.Forms.Label
    Friend WithEvents txtDesc As System.Windows.Forms.TextBox
    Friend WithEvents btnClose As System.Windows.Forms.Button

End Class
