<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Frm_DGV
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
        Me.DGV_Consulta = New System.Windows.Forms.DataGridView()
        CType(Me.DGV_Consulta, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DGV_Consulta
        '
        Me.DGV_Consulta.AllowUserToAddRows = False
        Me.DGV_Consulta.AllowUserToDeleteRows = False
        Me.DGV_Consulta.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DGV_Consulta.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DGV_Consulta.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGV_Consulta.Location = New System.Drawing.Point(12, 12)
        Me.DGV_Consulta.Name = "DGV_Consulta"
        Me.DGV_Consulta.ReadOnly = True
        Me.DGV_Consulta.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DGV_Consulta.Size = New System.Drawing.Size(857, 440)
        Me.DGV_Consulta.TabIndex = 8
        '
        'Frm_DGV
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(881, 464)
        Me.Controls.Add(Me.DGV_Consulta)
        Me.Name = "Frm_DGV"
        Me.Text = "Consulta"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.DGV_Consulta, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents DGV_Consulta As System.Windows.Forms.DataGridView
End Class
