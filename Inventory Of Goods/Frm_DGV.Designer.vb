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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Frm_DGV))
        Me.DGV_Consulta = New System.Windows.Forms.DataGridView()
        Me.CMS_DGV = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ExcluirDadosToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.LblLinhas = New System.Windows.Forms.Label()
        Me.CopiarDadosToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        CType(Me.DGV_Consulta, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.CMS_DGV.SuspendLayout()
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
        Me.DGV_Consulta.Location = New System.Drawing.Point(16, 15)
        Me.DGV_Consulta.Margin = New System.Windows.Forms.Padding(4)
        Me.DGV_Consulta.Name = "DGV_Consulta"
        Me.DGV_Consulta.ReadOnly = True
        Me.DGV_Consulta.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DGV_Consulta.Size = New System.Drawing.Size(1143, 520)
        Me.DGV_Consulta.TabIndex = 8
        '
        'CMS_DGV
        '
        Me.CMS_DGV.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.CMS_DGV.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ExcluirDadosToolStripMenuItem, Me.CopiarDadosToolStripMenuItem})
        Me.CMS_DGV.Name = "CMS_DGV"
        Me.CMS_DGV.Size = New System.Drawing.Size(176, 80)
        '
        'ExcluirDadosToolStripMenuItem
        '
        Me.ExcluirDadosToolStripMenuItem.Name = "ExcluirDadosToolStripMenuItem"
        Me.ExcluirDadosToolStripMenuItem.Size = New System.Drawing.Size(175, 24)
        Me.ExcluirDadosToolStripMenuItem.Text = "Excluir Dados"
        '
        'LblLinhas
        '
        Me.LblLinhas.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblLinhas.Location = New System.Drawing.Point(952, 545)
        Me.LblLinhas.Name = "LblLinhas"
        Me.LblLinhas.Size = New System.Drawing.Size(207, 23)
        Me.LblLinhas.TabIndex = 9
        Me.LblLinhas.Text = "Total de Registros: 0"
        Me.LblLinhas.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'CopiarDadosToolStripMenuItem
        '
        Me.CopiarDadosToolStripMenuItem.Name = "CopiarDadosToolStripMenuItem"
        Me.CopiarDadosToolStripMenuItem.Size = New System.Drawing.Size(175, 24)
        Me.CopiarDadosToolStripMenuItem.Text = "Copiar Dados"
        '
        'Frm_DGV
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1175, 571)
        Me.Controls.Add(Me.LblLinhas)
        Me.Controls.Add(Me.DGV_Consulta)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "Frm_DGV"
        Me.Text = "Consulta"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.DGV_Consulta, System.ComponentModel.ISupportInitialize).EndInit()
        Me.CMS_DGV.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents DGV_Consulta As System.Windows.Forms.DataGridView
    Friend WithEvents CMS_DGV As ContextMenuStrip
    Friend WithEvents ExcluirDadosToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents LblLinhas As Label
    Friend WithEvents CopiarDadosToolStripMenuItem As ToolStripMenuItem
End Class
