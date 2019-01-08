<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmData
    Inherits System.Windows.Forms.Form

    'Descartar substituições de formulário para limpar a lista de componentes.
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

    'Exigido pelo Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'OBSERVAÇÃO: o procedimento a seguir é exigido pelo Windows Form Designer
    'Pode ser modificado usando o Windows Form Designer.  
    'Não o modifique usando o editor de códigos.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmData))
        Me.DTP = New System.Windows.Forms.DateTimePicker()
        Me.BtnData = New System.Windows.Forms.Button()
        Me.LblData = New System.Windows.Forms.Label()
        Me.LblDataAtual = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'DTP
        '
        Me.DTP.Cursor = System.Windows.Forms.Cursors.Hand
        Me.DTP.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DTP.Location = New System.Drawing.Point(15, 35)
        Me.DTP.Name = "DTP"
        Me.DTP.Size = New System.Drawing.Size(99, 20)
        Me.DTP.TabIndex = 0
        '
        'BtnData
        '
        Me.BtnData.Location = New System.Drawing.Point(120, 36)
        Me.BtnData.Name = "BtnData"
        Me.BtnData.Size = New System.Drawing.Size(75, 23)
        Me.BtnData.TabIndex = 1
        Me.BtnData.Text = "Atualizar"
        Me.BtnData.UseVisualStyleBackColor = True
        '
        'LblData
        '
        Me.LblData.AutoSize = True
        Me.LblData.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblData.Location = New System.Drawing.Point(12, 9)
        Me.LblData.Name = "LblData"
        Me.LblData.Size = New System.Drawing.Size(241, 13)
        Me.LblData.TabIndex = 2
        Me.LblData.Text = "Selecione a data para expirar o software:"
        '
        'LblDataAtual
        '
        Me.LblDataAtual.AutoSize = True
        Me.LblDataAtual.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblDataAtual.ForeColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.LblDataAtual.Location = New System.Drawing.Point(12, 82)
        Me.LblDataAtual.Name = "LblDataAtual"
        Me.LblDataAtual.Size = New System.Drawing.Size(0, 13)
        Me.LblDataAtual.TabIndex = 3
        '
        'FrmData
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(264, 104)
        Me.Controls.Add(Me.LblDataAtual)
        Me.Controls.Add(Me.LblData)
        Me.Controls.Add(Me.BtnData)
        Me.Controls.Add(Me.DTP)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "FrmData"
        Me.Text = "Data Software"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents DTP As DateTimePicker
    Friend WithEvents BtnData As Button
    Friend WithEvents LblData As Label
    Friend WithEvents LblDataAtual As Label
End Class
