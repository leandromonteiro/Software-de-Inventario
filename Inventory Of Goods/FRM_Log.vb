Imports System
Imports System.IO
Imports System.Collections

Public Class FRM_Log
    Dim I_E As New Inventário_Excel
    Public Caminho_Arquivo As String = ""
    Public Novo_Registro As Boolean
    Private Sub Caminho_Excel()
        Dim Caminho As String
        'Procurar Pasta de Fotos
        Folderbd.ShowDialog()
        'Inserir nome das Fotos no BD
        Caminho = Folderbd.SelectedPath

        If Caminho = "" Then
            Exit Sub
        End If

        Dim fileEntries As String() = Directory.GetFiles(Caminho)
        ' Inserindo as fotos no Excel.
        Dim fileName As String
        For Each fileName In fileEntries
            'Inserir caminho e no das fotos no Excel
            'I_E.Inserir_Fotos(fileName, Path.GetFileNameWithoutExtension(fileName))
        Next fileName

    End Sub
    Private Sub BtnNovo_Click(sender As Object, e As EventArgs) Handles BtnNovo.Click
        Caminho_Excel()
        'Criando Excel
        I_E.Modelo_Excel(Me, SFD)
        Caminho_Arquivo = SFD.FileName
        If Caminho_Arquivo = "" Then
            MsgBox("Nenhum arquivo foi selecionado, a aplicação será encerrada", MsgBoxStyle.Critical)
            Application.Exit()
        End If
        My.Settings.DataSource = Caminho_Arquivo
        Novo_Registro = True
        Frm_Inventário.Show()
        Frm_Inventário.BtnCopiar.Enabled = False
        Frm_Inventário.BtnNovo.Enabled = False
        Me.Hide()
    End Sub

    Private Sub BtnCarregado_Click(sender As Object, e As EventArgs) Handles BtnCarregado.Click
        Caminho_Excel()
        'Seleciona Excel para Abrir
        OFD.ShowDialog()
        Caminho_Arquivo = OFD.FileName
        If Caminho_Arquivo = "" Then
            MsgBox("Nenhum arquivo foi selecionado, a aplicação será encerrada", MsgBoxStyle.Critical)
            Application.Exit()
        End If
        My.Settings.DataSource = Caminho_Arquivo
        Novo_Registro = False

        Frm_DGV.Show()
        Me.Hide()
    End Sub

End Class