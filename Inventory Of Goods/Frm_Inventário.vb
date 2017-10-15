Imports System.IO
Public Class Frm_Inventário
    Dim I_E As New Inventário_Excel
    Public TUC As Integer
    Public TI As Integer
    Public TI_Cod_Geral_Todos As Integer

    Dim ID As Integer
    Public Sequencial As String

    Public UAR As Integer
    Public A1 As String
    Public A2 As String
    Public A3 As String
    Public A4 As String
    Public A5 As String
    Public A6 As String

    Public CM1 As String
    Public CM2 As String
    Public CM3 As String

    Public consultor As String
    Public lider As String

    Dim N_Fotos As Integer
    Dim Imagem As String
    Dim Nome_Imagem As String

    Public Fotos_Array As New ArrayList
    Public Add_Fotos_Array As Integer = 0

    Dim Invalidos As Boolean

    Private Sub Validacao_Salvar()
        If TxtLocal.Text = "" Or TxtODI.Text = "" Or CmbTI.Text = "" Or CmbTI_Geral.Text = "" Then
            MsgBox("Dados Incompletos na aba local", MsgBoxStyle.Exclamation)
            Invalidos = True
        End If
        If CmbTUC.Text = "" Then
            MsgBox("Preencha TUC", MsgBoxStyle.Exclamation)
            Invalidos = True
        End If
        If CmbA1.Text = "" Then
            MsgBox("Preencha Tipo de Bem", MsgBoxStyle.Exclamation)
            Invalidos = True
        End If
        If CmbUAR.Text = "" Then
            MsgBox("Preencha UAR", MsgBoxStyle.Exclamation)
            Invalidos = True
        End If
        If CmbA2.Text = "" And CmbA2.Enabled = True Then
            MsgBox("Preencha A2", MsgBoxStyle.Exclamation)
            Invalidos = True
        End If
        If CmbA3.Text = "" And CmbA3.Enabled = True Then
            MsgBox("Preencha A3", MsgBoxStyle.Exclamation)
            Invalidos = True
        End If
        If CmbA4.Text = "" And CmbA4.Enabled = True Then
            MsgBox("Preencha A4", MsgBoxStyle.Exclamation)
            Invalidos = True
        End If
        If CmbA5.Text = "" And CmbA5.Enabled = True Then
            MsgBox("Preencha A5", MsgBoxStyle.Exclamation)
            Invalidos = True
        End If
        If CmbA6.Text = "" And CmbA6.Enabled = True Then
            MsgBox("Preencha A6", MsgBoxStyle.Exclamation)
            Invalidos = True
        End If
        If CmbCm1.Text = "" Then
            MsgBox("Preencha cm1", MsgBoxStyle.Exclamation)
            Invalidos = True
        End If
        If CmbCm2.Text = "" Then
            MsgBox("Preencha cm2", MsgBoxStyle.Exclamation)
            Invalidos = True
        End If
        If CmbCm3.Text = "" Then
            MsgBox("Preencha cm3", MsgBoxStyle.Exclamation)
            Invalidos = True
        End If
        If TxtDesc.Text = "" Then
            MsgBox("Preencha descrição", MsgBoxStyle.Exclamation)
            Invalidos = True
        End If
        If TxtQtd.Text = "" Then
            MsgBox("Preencha quantidade", MsgBoxStyle.Exclamation)
            Invalidos = True
        End If
        If CmbUm.Text = "" Then
            MsgBox("Preencha unidade de medida", MsgBoxStyle.Exclamation)
            Invalidos = True
        End If
        If CmbStatus.Text = "" Then
            MsgBox("Preencha status do bem", MsgBoxStyle.Exclamation)
            Invalidos = True
        End If
        If CmbEstado.Text = "" Then
            MsgBox("Preencha estado do bem", MsgBoxStyle.Exclamation)
            Invalidos = True
        End If
        If TxtConsultor.Text = "" Then
            MsgBox("Preencha o e-mail do consultor", MsgBoxStyle.Exclamation)
            Invalidos = True
        End If
        If TxtLider.Text = "" Then
            MsgBox("Preencha o e-mail do líder", MsgBoxStyle.Exclamation)
            Invalidos = True
        End If
    End Sub

    Private Sub Limpar_Tudo()
        TxtBay.Text = ""
        CmbTUC.Text = ""
        CmbA1.Text = ""
        CmbA2.Text = ""
        CmbA3.Text = ""
        CmbA4.Text = ""
        CmbA5.Text = ""
        CmbA6.Text = ""
        CmbUAR.Text = ""
        CmbCm1.Text = ""
        CmbCm2.Text = ""
        CmbCm3.Text = ""
        TxtDesc.Text = ""
        TxtFabricante.Text = ""
        TxtModelo.Text = ""
        TxtSerie.Text = ""
        TxtObs.Text = ""
        TxtQtd.Text = 1
        CmbUm.Text = "UN"
        CmbAno.Text = ""
        CmbMes.Text = ""
        CmbDia.Text = ""
        CmbStatus.Text = ""
        CmbEstado.Text = ""
        TxtAltura.Text = ""
        TxtLargura.Text = ""
        TxtComprimento.Text = ""
        TxtArea.Text = ""
        TxtPe.Text = ""
        TxtObsCivil.Text = ""
        CmbA2.Enabled = True
        CmbA3.Enabled = True
        CmbA4.Enabled = True
        CmbA5.Enabled = True
        CmbA6.Enabled = True
        CmbA2.Items.Clear()
        CmbA3.Items.Clear()
        CmbA4.Items.Clear()
        CmbA5.Items.Clear()
        CmbA6.Items.Clear()
        CmbUAR.Items.Clear()
        CmbA1.Items.Clear()
        LblA2.Text = "A2:"
        LblA3.Text = "A3:"
        LblA4.Text = "A4:"
        LblA5.Text = "A5:"
        LblA6.Text = "A6:"
    End Sub

    Private Sub Limpar_Parcial()
        TxtSerie.Text = ""
        CmbStatus.Text = ""
        CmbEstado.Text = ""
        CmbAno.Text = ""
        CmbMes.Text = ""
        CmbDia.Text = ""
        TxtObs.Text = ""
        TxtObsCivil.Text = ""
    End Sub

    Public Sub Ajuste_A1()
        A1 = I_E.Buscar_A1(CmbA1, TUC)
        I_E.Consulta_A2_A6(CmbA2, TUC, A1, "Desc_A2")
        I_E.Consulta_A2_A6(CmbA3, TUC, A1, "Desc_A3")
        I_E.Consulta_A2_A6(CmbA4, TUC, A1, "Desc_A4")
        I_E.Consulta_A2_A6(CmbA5, TUC, A1, "Desc_A5")
        I_E.Consulta_A2_A6(CmbA6, TUC, A1, "Desc_A6")
        I_E.Buscar_Tabela(LblA2, TUC, A1, "Desc_A2")
        I_E.Buscar_Tabela(LblA3, TUC, A1, "Desc_A3")
        I_E.Buscar_Tabela(LblA4, TUC, A1, "Desc_A4")
        I_E.Buscar_Tabela(LblA5, TUC, A1, "Desc_A5")
        I_E.Buscar_Tabela(LblA6, TUC, A1, "Desc_A6")
    End Sub

    Public Sub Mostrar_Imagem()
        'I_E.Buscar_Fotos()
        If Fotos_Array.Count = 0 Then
            Exit Sub
        End If
        'Buscar caminho da imagem DS e colocar na Picture Box
        For Each dr As DataRow In I_E.DS.Tables("TB_Foto").Rows
            If dr(0).ToString = Fotos_Array(0) Then
                PictureBox_Consulta.ImageLocation = dr(1).ToString
                Add_Fotos_Array = 1
                Exit For
            End If
        Next
    End Sub

    Private Sub FrmInventario_Novo_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        FRM_Log.Close()
    End Sub

    Private Sub FrmInventario_Novo_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        I_E.Consulta_TUC(CmbTUC)
        I_E.Consulta_CM(CmbCm1, CmbCm2, CmbCm3)
        I_E.Consulta_TI_Geral(CmbTI_Geral)
        'If Nome_Imagem Is Nothing Then
        '    Exit Sub
        'End If
        'Nome_Imagem = I_E.DS.Tables("TB_Foto").Rows(0)(0)
        'Imagem = I_E.DS.Tables("TB_Foto").Rows(0)(1)
        N_Fotos = 0
        'PictureBox.ImageLocation = Imagem
    End Sub

    Private Sub CmbTUC_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbTUC.SelectedIndexChanged
        TUC = I_E.Buscar_TUC(CmbTUC)
        I_E.Consulta_UAR(CmbUAR, TUC)
        I_E.Consulta_A1(CmbA1, TUC)
        'Limpar
        CmbUAR.Text = ""
        CmbA1.Text = ""
        CmbA2.Text = ""
        CmbA3.Text = ""
        CmbA4.Text = ""
        CmbA5.Text = ""
        CmbA6.Text = ""
        CmbA2.Items.Clear()
        CmbA3.Items.Clear()
        CmbA4.Items.Clear()
        CmbA5.Items.Clear()
        CmbA6.Items.Clear()
        CmbA2.Enabled = True
        CmbA3.Enabled = True
        CmbA4.Enabled = True
        CmbA5.Enabled = True
        CmbA6.Enabled = True
        LblA2.Text = "A2:"
        LblA3.Text = "A3:"
        LblA4.Text = "A4:"
        LblA5.Text = "A5:"
        LblA6.Text = "A6:"
    End Sub

    Private Sub CmbCm1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbCm1.SelectedIndexChanged
        If CmbCm1.Text = "NÃO APLICÁVEL" Then
            CM1 = ""
        Else
            CM1 = I_E.Buscar_CM1(CmbCm1)
        End If
    End Sub

    Private Sub CmbCm2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbCm2.SelectedIndexChanged
        If CmbCm2.Text = "NÃO APLICÁVEL" Then
            CM2 = ""
        Else
            CM2 = I_E.Buscar_CM2(CmbCm2)
        End If
    End Sub

    Private Sub CmbCm3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbCm3.SelectedIndexChanged
        If CmbCm3.Text = "NÃO APLICÁVEL" Then
            CM3 = ""
        Else
            CM3 = I_E.Buscar_CM3(CmbCm3)
        End If
    End Sub

    Private Sub CmbTI_Geral_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbTI_Geral.SelectedIndexChanged
        TI_Cod_Geral_Todos = I_E.Buscar_TI_Geral(CmbTI_Geral)
        I_E.Consulta_TI(CmbTI, TI_Cod_Geral_Todos)
        CmbTI.Text = ""
    End Sub

    Private Sub CmbTI_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbTI.SelectedIndexChanged
        TI = I_E.Buscar_TI(CmbTI, TI_Cod_Geral_Todos)
    End Sub

    Private Sub CmbUAR_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbUAR.SelectedIndexChanged
        UAR = I_E.Buscar_UAR(CmbUAR, TUC)
    End Sub

    Private Sub CmbA1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbA1.SelectedIndexChanged
        Me.Cursor = Cursors.WaitCursor
        Ajuste_A1()
        'Limpar
        CmbA2.Text = ""
        CmbA3.Text = ""
        CmbA4.Text = ""
        CmbA5.Text = ""
        CmbA6.Text = ""

        Me.Cursor = Cursors.Default
    End Sub

    Private Sub CmbA2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbA2.SelectedIndexChanged
        A2 = I_E.Buscar_A2_A6(CmbA2)
    End Sub

    Private Sub CmbA3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbA3.SelectedIndexChanged
        A3 = I_E.Buscar_A2_A6(CmbA3)
    End Sub

    Private Sub CmbA4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbA4.SelectedIndexChanged
        A4 = I_E.Buscar_A2_A6(CmbA4)
    End Sub

    Private Sub CmbA5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbA5.SelectedIndexChanged
        A5 = I_E.Buscar_A2_A6(CmbA5)
    End Sub

    Private Sub CmbA6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbA6.SelectedIndexChanged
        A6 = I_E.Buscar_A2_A6(CmbA6)
    End Sub


    Private Sub BtnAnterior_Click(sender As Object, e As EventArgs) Handles BtnAnterior.Click
        If N_Fotos = 0 Then
            Exit Sub
        End If
        N_Fotos -= 1
        Nome_Imagem = I_E.DS.Tables("TB_Foto").Rows(N_Fotos)(0)
        Imagem = I_E.DS.Tables("TB_Foto").Rows(N_Fotos)(1)
        PictureBox.ImageLocation = Imagem
    End Sub

    Private Sub BtnProximo_Click(sender As Object, e As EventArgs) Handles BtnProximo.Click
        N_Fotos += 1
        If N_Fotos > I_E.DS.Tables("TB_Foto").Rows.Count - 1 Then
            N_Fotos -= 1
            Exit Sub
        End If

        Nome_Imagem = I_E.DS.Tables("TB_Foto").Rows(N_Fotos)(0)
        Imagem = I_E.DS.Tables("TB_Foto").Rows(N_Fotos)(1)
        PictureBox.ImageLocation = Imagem
    End Sub

    Private Sub BtnZoom_Click(sender As Object, e As EventArgs) Handles BtnZoom.Click
        Dim p As New Process()
        p.StartInfo.FileName = "rundll32.exe"
        p.StartInfo.Arguments = Path.Combine(Environment.SystemDirectory, "shimgvw.dll" & ",ImageView_Fullscreen " & PictureBox.ImageLocation)
        p.Start()
    End Sub

    Private Sub BtnVisualizar_Consulta_Click(sender As Object, e As EventArgs) Handles BtnVisualizar_Consulta.Click
        Dim p As New Process()
        p.StartInfo.FileName = "rundll32.exe"
        p.StartInfo.Arguments = Path.Combine(Environment.SystemDirectory, "shimgvw.dll" & ",ImageView_Fullscreen " & PictureBox_Consulta.ImageLocation)
        p.Start()
    End Sub

    Private Sub BtnSalvar_Click(sender As Object, e As EventArgs) Handles BtnSalvar.Click
        'Validação de Dados
        Validacao_Salvar()
        If Invalidos = True Then
            Invalidos = False
            Exit Sub
        End If

        'Update BD
        Sequencial = TxtLocal.Text & " - " & TxtSeq_Local.Text
        consultor = TxtConsultor.Text
        lider = TxtLider.Text
        ID = TxtSeq_Civil.Text

        If TxtAltura.Text = "" Then
            TxtAltura.Text = 0
        End If
        If TxtLargura.Text = "" Then
            TxtLargura.Text = 0
        End If
        If TxtComprimento.Text = "" Then
            TxtComprimento.Text = 0
        End If
        If TxtArea.Text = "" Then
            TxtArea.Text = 0
        End If
        If TxtPe.Text = "" Then
            TxtPe.Text = 0
        End If

        I_E.Update_Inventario(ID, Sequencial, TxtLocal.Text, TxtODI.Text, TI, CmbTI.Text, TxtBay.Text, TUC, CmbTUC.Text, A1, CmbA1.Text,
                              UAR, CmbUAR.Text, A2, CmbA2.Text, A3, CmbA3.Text, A4, CmbA4.Text, A5, CmbA5.Text, A6, CmbA6.Text, CM1, CmbCm1.Text,
                              CM2, CmbCm2.Text, CM3, CmbCm3.Text, TxtDesc.Text, TxtFabricante.Text, TxtModelo.Text, TxtSerie.Text, TxtObs.Text,
                              TxtQtd.Text, CmbUm.Text, CmbAno.Text, CmbMes.Text, CmbDia.Text, CmbStatus.Text, CmbEstado.Text, TxtAltura.Text,
                              TxtLargura.Text, TxtComprimento.Text, TxtArea.Text, TxtPe.Text, TxtObsCivil.Text, "", consultor, lider, Now())


        BtnNovo.Enabled = True
        BtnCopiar.Enabled = True
    End Sub

    Private Sub BtnNovo_Click(sender As Object, e As EventArgs) Handles BtnNovo.Click
        'Consultar ID +1
        ID = I_E.Buscar_Ultimo_ID()
        ID += 1
        TxtSeq_Civil.Text = ID
        TxtSeq_Desc.Text = ID
        TxtSeq_Local.Text = ID
        'Inserir no BD
        I_E.Inserir_ID(ID)
        'Limpar Dados
        Limpar_Tudo()
        Fotos_Array.Clear()
        PictureBox_Consulta.ImageLocation = ""
        Add_Fotos_Array = 0

        BtnNovo.Enabled = False
        BtnCopiar.Enabled = False
    End Sub

    Private Sub BtnCopiar_Click(sender As Object, e As EventArgs) Handles BtnCopiar.Click
        'Consultar ID +1
        ID = I_E.Buscar_Ultimo_ID()
        ID += 1
        TxtSeq_Civil.Text = ID
        TxtSeq_Desc.Text = ID
        TxtSeq_Local.Text = ID
        'Inserir no BD
        I_E.Inserir_ID(ID)
        'Limpar Alguns dados
        Limpar_Parcial()
        Fotos_Array.Clear()
        PictureBox_Consulta.ImageLocation = ""
        Add_Fotos_Array = 0

        BtnNovo.Enabled = False
        BtnCopiar.Enabled = False

        'Consultar A2 a A6
        Ajuste_A1()
    End Sub

    Private Sub TxtQtd_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtQtd.KeyPress
        If e.KeyChar = ","c Then
            e.Handled = (CType(sender, TextBox).Text.IndexOf(","c) <> -1)
        ElseIf e.KeyChar <> ControlChars.Back Then
            e.Handled = ("0123456789".IndexOf(e.KeyChar) = -1)
        End If
    End Sub

    Private Sub TxtAltura_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtAltura.KeyPress
        If e.KeyChar = ","c Then
            e.Handled = (CType(sender, TextBox).Text.IndexOf(","c) <> -1)
        ElseIf e.KeyChar <> ControlChars.Back Then
            e.Handled = ("0123456789".IndexOf(e.KeyChar) = -1)
        End If
    End Sub

    Private Sub TxtLargura_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtLargura.KeyPress
        If e.KeyChar = ","c Then
            e.Handled = (CType(sender, TextBox).Text.IndexOf(","c) <> -1)
        ElseIf e.KeyChar <> ControlChars.Back Then
            e.Handled = ("0123456789".IndexOf(e.KeyChar) = -1)
        End If
    End Sub

    Private Sub TxtComprimento_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtComprimento.KeyPress
        If e.KeyChar = ","c Then
            e.Handled = (CType(sender, TextBox).Text.IndexOf(","c) <> -1)
        ElseIf e.KeyChar <> ControlChars.Back Then
            e.Handled = ("0123456789".IndexOf(e.KeyChar) = -1)
        End If
    End Sub

    Private Sub TxtArea_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtArea.KeyPress
        If e.KeyChar = ","c Then
            e.Handled = (CType(sender, TextBox).Text.IndexOf(","c) <> -1)
        ElseIf e.KeyChar <> ControlChars.Back Then
            e.Handled = ("0123456789".IndexOf(e.KeyChar) = -1)
        End If
    End Sub

    Private Sub TxtPe_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtPe.KeyPress
        If e.KeyChar = ","c Then
            e.Handled = (CType(sender, TextBox).Text.IndexOf(","c) <> -1)
        ElseIf e.KeyChar <> ControlChars.Back Then
            e.Handled = ("0123456789".IndexOf(e.KeyChar) = -1)
        End If
    End Sub

    Private Sub BtnAdd_Click(sender As Object, e As EventArgs) Handles BtnAdd.Click
        If Fotos_Array.Count = 10 Then
            MsgBox("O limite de imagens por cadastro são 10", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        Fotos_Array.Add(Nome_Imagem)
        'Buscar caminho da imagem DS e colocar na Picture Box
        For Each dr As DataRow In I_E.DS.Tables("TB_Foto").Rows
            If dr(0).ToString = Fotos_Array(Fotos_Array.Count - 1) Then
                PictureBox_Consulta.ImageLocation = dr(1).ToString
                Add_Fotos_Array = Fotos_Array.Count - 1
                Exit For
            End If
        Next
        MsgBox("Fotos adicionadas com sucesso", MsgBoxStyle.Information)
    End Sub
    Private Sub BtnRemover_Fotos_Click(sender As Object, e As EventArgs) Handles BtnRemover_Fotos.Click
        If Fotos_Array.Count = 0 Then
            PictureBox_Consulta.ImageLocation = ""
            Exit Sub
        End If

        Fotos_Array.RemoveAt(Add_Fotos_Array)

        If Add_Fotos_Array > 0 Then
            Add_Fotos_Array -= 1
        End If

        If Fotos_Array.Count = 0 Then
            PictureBox_Consulta.ImageLocation = ""
            Exit Sub
        End If
        'Buscar caminho da imagem DS e colocar na Picture Box
        For Each dr As DataRow In I_E.DS.Tables("TB_Foto").Rows
            If dr(0).ToString = Fotos_Array(Add_Fotos_Array) Then
                PictureBox_Consulta.ImageLocation = dr(1).ToString
            End If
        Next
    End Sub

    Private Sub BtnAnterior_Consulta_Click(sender As Object, e As EventArgs) Handles BtnAnterior_Consulta.Click
        If Add_Fotos_Array > Fotos_Array.Count - 1 Then
            Exit Sub
        End If
        If Add_Fotos_Array = 0 Then
            'Buscar caminho da imagem DS e colocar na Picture Box
            For Each dr As DataRow In I_E.DS.Tables("TB_Foto").Rows
                If dr(0).ToString = Fotos_Array(0) Then
                    PictureBox_Consulta.ImageLocation = dr(1).ToString
                End If
            Next
            Exit Sub
        End If
        Add_Fotos_Array -= 1
        'Buscar caminho da imagem DS e colocar na Picture Box
        For Each dr As DataRow In I_E.DS.Tables("TB_Foto").Rows
            If dr(0).ToString = Fotos_Array(Add_Fotos_Array) Then
                PictureBox_Consulta.ImageLocation = dr(1).ToString
            End If
        Next
    End Sub

    Private Sub BtnProximo_Consulta_Click(sender As Object, e As EventArgs) Handles BtnProximo_Consulta.Click
        If Add_Fotos_Array > Fotos_Array.Count - 1 Then
            Exit Sub
        End If
        If Add_Fotos_Array = Fotos_Array.Count - 1 Then
            'Buscar caminho da imagem DS e colocar na Picture Box
            For Each dr As DataRow In I_E.DS.Tables("TB_Foto").Rows
                If dr(0).ToString = Fotos_Array(Add_Fotos_Array) Then
                    PictureBox_Consulta.ImageLocation = dr(1).ToString
                End If
            Next
            Exit Sub
        End If
        Add_Fotos_Array += 1
        'Buscar caminho da imagem DS e colocar na Picture Box
        For Each dr As DataRow In I_E.DS.Tables("TB_Foto").Rows
            If dr(0).ToString = Fotos_Array(Add_Fotos_Array) Then
                PictureBox_Consulta.ImageLocation = dr(1).ToString
            End If
        Next
    End Sub

    Private Sub BtnConsultar_Click(sender As Object, e As EventArgs) Handles BtnConsultar.Click
        Frm_DGV.Show()
        Me.Hide()
    End Sub

End Class
