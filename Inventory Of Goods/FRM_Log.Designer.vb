<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FRM_Log
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
        Me.BtnCarregado = New System.Windows.Forms.Button()
        Me.BtnNovo = New System.Windows.Forms.Button()
        Me.SFD = New System.Windows.Forms.SaveFileDialog()
        Me.OFD = New System.Windows.Forms.OpenFileDialog()
        Me.Folderbd = New System.Windows.Forms.FolderBrowserDialog()
        Me.SuspendLayout()
        '
        'BtnCarregado
        '
        Me.BtnCarregado.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnCarregado.Location = New System.Drawing.Point(12, 41)
        Me.BtnCarregado.Name = "BtnCarregado"
        Me.BtnCarregado.Size = New System.Drawing.Size(149, 23)
        Me.BtnCarregado.TabIndex = 3
        Me.BtnCarregado.Text = "Existente"
        Me.BtnCarregado.UseVisualStyleBackColor = True
        '
        'BtnNovo
        '
        Me.BtnNovo.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnNovo.Location = New System.Drawing.Point(12, 12)
        Me.BtnNovo.Name = "BtnNovo"
        Me.BtnNovo.Size = New System.Drawing.Size(149, 23)
        Me.BtnNovo.TabIndex = 2
        Me.BtnNovo.Text = "Novo"
        Me.BtnNovo.UseVisualStyleBackColor = True
        '
        'OFD
        '
        Me.OFD.Title = "Escolha o Arquivo Excel (Inventário)"
        '
        'Folderbd
        '
        Me.Folderbd.Description = "Escolha a pasta de Fotos"
        '
        'FRM_Log
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(170, 77)
        Me.Controls.Add(Me.BtnCarregado)
        Me.Controls.Add(Me.BtnNovo)
        Me.Name = "FRM_Log"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Início"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents BtnCarregado As System.Windows.Forms.Button
    Friend WithEvents BtnNovo As System.Windows.Forms.Button
    Friend WithEvents SFD As System.Windows.Forms.SaveFileDialog
    Friend WithEvents OFD As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Folderbd As System.Windows.Forms.FolderBrowserDialog
End Class
