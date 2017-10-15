Imports System.Net.Mail
Imports Microsoft.Office.Interop

Public Class Inventário_Excel
    Dim connstr As String = "Data Source=C:\Users\Public\INVENTÁRIO.db;"
    Public DS As New DataSet
    Public Foto1 As String
    Public Foto2 As String
    Public Foto3 As String
    Public Foto4 As String
    Public Foto5 As String
    Public Foto6 As String
    Public Foto7 As String
    Public Foto8 As String
    Public Foto9 As String
    Public Foto10 As String

    Public Function Consulta_TI_Geral(TI As Integer)
        Try
            Dim leitor As SQLite.SQLiteDataReader
            Dim connection As New SQLite.SQLiteConnection(connstr)
            Dim cmd As New SQLite.SQLiteCommand
            connection.Open()
            cmd.Connection = connection
            cmd.CommandText = "select Desc_Geral from [TI$] where Cod_Geral=" &
                "(select Cod_Geral_Todos from [TI$] where Cod=" & TI & ");"
            leitor = cmd.ExecuteReader
            leitor.Read()
            Return leitor("Desc_Geral")
            connection.Close()

        Catch
            MsgBox("Erro ao buscar dados no Excel", MsgBoxStyle.Critical)
            Return Nothing
        End Try
    End Function

    Public Sub Consulta_Fotos(ID As Integer)
        Try
            Dim leitor As SQLite.SQLiteDataReader
            Dim connection As New SQLite.SQLiteConnection(connstr)
            Dim cmd As New SQLite.SQLiteCommand
            connection.Open()
            cmd.Connection = connection
            cmd.CommandText = "select FOTO 1,FOTO 2,FOTO 3,FOTO 4,FOTO 5,FOTO 6,FOTO 7,FOTO 8,FOTO 9,FOTO 10 from [Inventario$] where ID=" & ID & ";"
            leitor = cmd.ExecuteReader
            leitor.Read()
            Foto1 = leitor("FOTO 1")
            Foto2 = leitor("FOTO 2")
            Foto3 = leitor("FOTO 3")
            Foto4 = leitor("FOTO 4")
            Foto5 = leitor("FOTO 5")
            Foto6 = leitor("FOTO 6")
            Foto7 = leitor("FOTO 7")
            Foto8 = leitor("FOTO 8")
            Foto9 = leitor("FOTO 9")
            Foto10 = leitor("FOTO 10")
            connection.Close()

        Catch
            MsgBox("Erro ao buscar dados no Excel", MsgBoxStyle.Critical)
        End Try
    End Sub

    'Public Sub Update_Fotos(FOTO1 As String, FOTO2 As String, FOTO3 As String, FOTO4 As String, FOTO5 As String, FOTO6 As String,
    '                        FOTO7 As String, FOTO8 As String, FOTO9 As String, FOTO10 As String, ID As Integer)
    '    Try
    '        Dim connection As New OleDb.OleDbConnection(connstr_consulta)
    '        Dim cmd As New OleDb.OleDbCommand
    '        connection.Open()
    '        cmd.Connection = connection
    '        cmd.CommandText = "update [Inventario$] set [FOTO 1]='" & FOTO1 & "',[FOTO 2]='" & FOTO2 & "',[FOTO 3]='" & FOTO3 & _
    '        "',[FOTO 4]='" & FOTO4 & "',[FOTO 5]='" & FOTO5 & "',[FOTO 6]='" & FOTO6 & "',[FOTO 7]='" & FOTO7 & _
    '        "',[FOTO 8]='" & FOTO8 & "',[FOTO 9]='" & FOTO9 & "',[FOTO 10]='" & FOTO10 & "' where [ID]=" & ID & ";"
    '        cmd.ExecuteNonQuery()
    '        connection.Close()

    '    Catch
    '        MsgBox("Erro ao inserir foto no Excel", MsgBoxStyle.Critical)
    '    End Try
    'End Sub

    Public Sub Consulta_TUC(cmb As ComboBox)
        cmb.Items.Clear()
        Try
            Dim leitor As SQLite.SQLiteDataReader
            Dim connection As New SQLite.SQLiteConnection(connstr)
            Dim cmd As New SQLite.SQLiteCommand
            connection.Open()
            cmd.Connection = connection
            cmd.CommandText = "select Descricao from [TUC$];"
            leitor = cmd.ExecuteReader
            Do While leitor.Read
                cmb.Items.Add(leitor("Descricao"))
            Loop
            connection.Close()

        Catch
            MsgBox("Erro ao buscar dados no Excel", MsgBoxStyle.Critical)
        End Try
    End Sub

    Public Sub Consulta_TI_Geral(cmb As ComboBox)
        cmb.Items.Clear()
        Try
            Dim leitor As SQLite.SQLiteDataReader
            Dim connection As New SQLite.SQLiteConnection(connstr)
            Dim cmd As New SQLite.SQLiteCommand
            connection.Open()
            cmd.Connection = connection
            cmd.CommandText = "select Desc_Geral from [TI$];"
            leitor = cmd.ExecuteReader
            Do While leitor.Read
                If Not IsDBNull(leitor("Desc_Geral")) Then
                    cmb.Items.Add(leitor("Desc_Geral"))
                End If
            Loop
            connection.Close()

        Catch
            MsgBox("Erro ao buscar dados no Excel", MsgBoxStyle.Critical)
        End Try
    End Sub

    Public Sub Consulta_TI(cmb As ComboBox, cod_geral As Integer)
        cmb.Items.Clear()
        Try
            Dim leitor As SQLite.SQLiteDataReader
            Dim connection As New SQLite.SQLiteConnection(connstr)
            Dim cmd As New SQLite.SQLiteCommand
            connection.Open()
            cmd.Connection = connection
            cmd.CommandText = "select Descricao from [TI$] where Cod_Geral_Todos=" & cod_geral & ";"
            leitor = cmd.ExecuteReader
            Do While leitor.Read
                cmb.Items.Add(leitor("Descricao"))
            Loop
            connection.Close()

        Catch
            MsgBox("Erro ao buscar dados no Excel", MsgBoxStyle.Critical)
        End Try
    End Sub

    Public Sub Consulta_CM(cmb1 As ComboBox, cmb2 As ComboBox, cmb3 As ComboBox)
        cmb1.Items.Clear()
        cmb2.Items.Clear()
        cmb3.Items.Clear()
        Try
            Dim leitor As SQLite.SQLiteDataReader
            Dim connection As New SQLite.SQLiteConnection(connstr)
            Dim cmd As New SQLite.SQLiteCommand
            connection.Open()
            cmd.Connection = connection
            cmd.CommandText = "select CM1,CM2,CM3 from [CM$];"
            leitor = cmd.ExecuteReader
            Do While leitor.Read
                If Not IsDBNull(leitor("CM1")) Then
                    cmb1.Items.Add(leitor("CM1"))
                End If
                cmb2.Items.Add(leitor("CM2"))
                If Not IsDBNull(leitor("CM3")) Then
                    cmb3.Items.Add(leitor("CM3"))
                End If
            Loop
            connection.Close()

        Catch
            MsgBox("Erro ao buscar dados no Excel", MsgBoxStyle.Critical)
        End Try
    End Sub

    Public Sub Consulta_UAR(cmb As ComboBox, TUC As Integer)
        cmb.Items.Clear()
        Try
            Dim leitor As SQLite.SQLiteDataReader
            Dim connection As New SQLite.SQLiteConnection(connstr)
            Dim cmd As New SQLite.SQLiteCommand
            connection.Open()
            cmd.Connection = connection
            cmd.CommandText = "select Descricao from [UAR$] where TUC=" & TUC & ";"
            leitor = cmd.ExecuteReader
            Do While leitor.Read
                cmb.Items.Add(leitor("Descricao"))
            Loop
            connection.Close()

        Catch
            MsgBox("Erro ao buscar dados no Excel", MsgBoxStyle.Critical)
        End Try
    End Sub

    Public Sub Consulta_A1(cmb As ComboBox, TUC As Integer)
        cmb.Items.Clear()
        Try
            Dim leitor As SQLite.SQLiteDataReader
            Dim connection As New SQLite.SQLiteConnection(connstr)
            Dim cmd As New SQLite.SQLiteCommand
            connection.Open()
            cmd.Connection = connection
            cmd.CommandText = "select Desc_A1 from [ATRIBUTOS$] where TUC=" & TUC & ";"
            leitor = cmd.ExecuteReader
            Do While leitor.Read
                cmb.Items.Add(leitor("Desc_A1"))
            Loop
            connection.Close()

        Catch
            MsgBox("Erro ao buscar dados no Excel", MsgBoxStyle.Critical)
        End Try
    End Sub

    Public Sub Consulta_A2_A6(cmb As ComboBox, TUC As Integer, A1 As String, Desc_A As String)
        Dim N As Integer = 0
        cmb.Items.Clear()
        cmb.Enabled = True
        Try
            Dim leitor As SQLite.SQLiteDataReader
            Dim connection As New SQLite.SQLiteConnection(connstr)
            Dim cmd As New SQLite.SQLiteCommand
            connection.Open()
            cmd.Connection = connection
            cmd.CommandText = "select Descricao from [ATRIBUTOS$] where Cod_Tabela_Geral=(select Cod_Tabela from [ATRIBUTOS$] where Tabela=(select " & _
            Desc_A & " from [ATRIBUTOS$] where TUC=" & TUC & " and A1='" & A1 & "'));"
            leitor = cmd.ExecuteReader
            Do While leitor.Read
                N = 1
                cmb.Items.Add(leitor("Descricao"))
            Loop
            If N = 0 Then
                cmb.Enabled = False
            End If
            connection.Close()

        Catch
            MsgBox("Erro ao buscar dados no Excel", MsgBoxStyle.Critical)
        End Try
    End Sub

    Public Sub Buscar_Tabela(Lbl As Label, TUC As Integer, A1 As String, Desc_A As String)
        Try
            Dim leitor As SQLite.SQLiteDataReader
            Dim connection As New SQLite.SQLiteConnection(connstr)
            Dim cmd As New SQLite.SQLiteCommand
            connection.Open()
            cmd.Connection = connection
            cmd.CommandText = "select " & _
            Desc_A & " from [ATRIBUTOS$] where TUC=" & TUC & " and A1='" & A1 & "';"
            leitor = cmd.ExecuteReader
            Do While leitor.Read
                Lbl.Text = leitor(Desc_A)
            Loop
            connection.Close()

        Catch
            MsgBox("Erro ao buscar dados no Excel", MsgBoxStyle.Critical)
        End Try
    End Sub

    Public Function Buscar_A2_A6(cmb As ComboBox)
        Try
            Dim leitor As SQLite.SQLiteDataReader
            Dim connection As New SQLite.SQLiteConnection(connstr)
            Dim cmd As New SQLite.SQLiteCommand
            connection.Open()
            cmd.Connection = connection
            cmd.CommandText = "select Cod_Atributos from [ATRIBUTOS$] where Descricao='" & cmb.Text & "';"
            leitor = cmd.ExecuteReader
            leitor.Read()
            Return leitor("Cod_Atributos")
            connection.Close()

        Catch
            MsgBox("Erro ao buscar dados no Excel", MsgBoxStyle.Critical)
            Return 0
        End Try
    End Function

    Public Function Buscar_TI_Geral(cmb As ComboBox)
        Try
            Dim leitor As SQLite.SQLiteDataReader
            Dim connection As New SQLite.SQLiteConnection(connstr)
            Dim cmd As New SQLite.SQLiteCommand
            connection.Open()
            cmd.Connection = connection
            cmd.CommandText = "select Cod_Geral from [TI$] where Desc_Geral='" & cmb.Text & "';"
            leitor = cmd.ExecuteReader
            leitor.Read()
            Return leitor("Cod_Geral")
            connection.Close()

        Catch
            MsgBox("Erro ao buscar dados no Excel", MsgBoxStyle.Critical)
            Return 0
        End Try
    End Function

    Public Function Buscar_TI(cmb As ComboBox, cod_geral_todos As Integer)
        Try
            Dim leitor As SQLite.SQLiteDataReader
            Dim connection As New SQLite.SQLiteConnection(connstr)
            Dim cmd As New SQLite.SQLiteCommand
            connection.Open()
            cmd.Connection = connection
            cmd.CommandText = "select Cod from [TI$] where Cod_Geral_Todos=" & cod_geral_todos & " and Descricao='" & cmb.Text & "';"
            leitor = cmd.ExecuteReader
            leitor.Read()
            Return leitor("Cod")
            connection.Close()

        Catch
            MsgBox("Erro ao buscar dados no Excel", MsgBoxStyle.Critical)
            Return 0
        End Try
    End Function

    Public Function Buscar_UAR(cmb As ComboBox, TUC As Integer)
        Try
            Dim leitor As SQLite.SQLiteDataReader
            Dim connection As New SQLite.SQLiteConnection(connstr)
            Dim cmd As New SQLite.SQLiteCommand
            connection.Open()
            cmd.Connection = connection
            cmd.CommandText = "select Cod from [UAR$] where TUC=" & TUC & " and Descricao='" & cmb.Text & "';"
            leitor = cmd.ExecuteReader
            leitor.Read()
            Return leitor("Cod")
            connection.Close()

        Catch
            MsgBox("Erro ao buscar dados no Excel", MsgBoxStyle.Critical)
            Return 0
        End Try
    End Function

    Public Function Buscar_A1(cmb As ComboBox, TUC As Integer)
        Try
            Dim leitor As SQLite.SQLiteDataReader
            Dim connection As New SQLite.SQLiteConnection(connstr)
            Dim cmd As New SQLite.SQLiteCommand
            connection.Open()
            cmd.Connection = connection
            cmd.CommandText = "select A1 from [ATRIBUTOS$] where TUC=" & TUC & " and Desc_A1='" & cmb.Text & "';"
            leitor = cmd.ExecuteReader
            leitor.Read()
            Return leitor("A1")
            connection.Close()

        Catch
            MsgBox("Erro ao buscar dados no Excel", MsgBoxStyle.Critical)
            Return 0
        End Try
    End Function

    Public Function Buscar_TUC(cmb As ComboBox)
        Try
            Dim leitor As SQLite.SQLiteDataReader
            Dim connection As New SQLite.SQLiteConnection(connstr)
            Dim cmd As New SQLite.SQLiteCommand
            connection.Open()
            cmd.Connection = connection
            cmd.CommandText = "select TUC from [TUC$] where Descricao='" & cmb.Text & "';"
            leitor = cmd.ExecuteReader
            leitor.Read()
            Return leitor("TUC")
        connection.Close()

        Catch
            MsgBox("Erro ao buscar dados no Excel", MsgBoxStyle.Critical)
            Return 0
        End Try
    End Function

    Public Function Buscar_CM1(cmb As ComboBox)
        Try
            Dim leitor As SQLite.SQLiteDataReader
            Dim connection As New SQLite.SQLiteConnection(connstr)
            Dim cmd As New SQLite.SQLiteCommand
            connection.Open()
            cmd.Connection = connection
            cmd.CommandText = "select C1 from [CM$] where CM1='" & cmb.Text & "';"
            leitor = cmd.ExecuteReader
            leitor.Read()
            Return leitor("C1")
            connection.Close()

        Catch
            MsgBox("Erro ao buscar dados no Excel", MsgBoxStyle.Critical)
            Return 0
        End Try
    End Function

    Public Function Buscar_CM2(cmb As ComboBox)
        Try
            Dim leitor As SQLite.SQLiteDataReader
            Dim connection As New SQLite.SQLiteConnection(connstr)
            Dim cmd As New SQLite.SQLiteCommand
            connection.Open()
            cmd.Connection = connection
            cmd.CommandText = "select C2 from [CM$] where CM2='" & cmb.Text & "';"
            leitor = cmd.ExecuteReader
            leitor.Read()
            Return leitor("C2")
            connection.Close()

        Catch
            MsgBox("Erro ao buscar dados no Excel", MsgBoxStyle.Critical)
            Return 0
        End Try
    End Function

    Public Function Buscar_CM3(cmb As ComboBox)
        Try
            Dim leitor As SQLite.SQLiteDataReader
            Dim connection As New SQLite.SQLiteConnection(connstr)
            Dim cmd As New SQLite.SQLiteCommand
            connection.Open()
            cmd.Connection = connection
            cmd.CommandText = "select C3 from [CM$] where CM3='" & cmb.Text & "';"
            leitor = cmd.ExecuteReader
            leitor.Read()
            Return leitor("C3")
            connection.Close()

        Catch
            MsgBox("Erro ao buscar dados no Excel", MsgBoxStyle.Critical)
            Return 0
        End Try
    End Function

    Public Sub Excluir_Fotos()
        Dim xlApp As New Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As New Excel.Worksheet

        xlWorkBook = xlApp.Workbooks.Open("C:\Users\Public\INVENTÁRIO_BD.xlsx")
        xlWorkSheet = xlWorkBook.Sheets("ADM")
        xlWorkSheet.Range("A2:B10000").Clear()
        xlWorkBook.Save()
        xlWorkBook.Close()
    End Sub

    Public Sub Inserir_Fotos(Caminho_Foto As String, Nome_Foto As String)
        Try
            Dim connection As New SQLite.SQLiteConnection(connstr)
            Dim cmd As New SQLite.SQLiteCommand
            connection.Open()
            cmd.Connection = connection
            cmd.CommandText = "insert into [ADM$] ([CAMINHO FOTOS],[FOTOS]) values('" & Caminho_Foto & "','" & Nome_Foto & "');"
            cmd.ExecuteNonQuery()
            connection.Close()

        Catch
            MsgBox("Erro ao buscar dados no Excel", MsgBoxStyle.Critical)
        End Try
    End Sub

    Public Sub Buscar_Fotos()
        Try
            DS.Clear()
            Dim DA As New SQLite.SQLiteDataAdapter
            Dim connection As New SQLite.SQLiteConnection(connstr)
            connection.Open()
            DA.SelectCommand = New SQLite.SQLiteCommand("select [FOTOS],[CAMINHO FOTOS] from [ADM$];", connection)
            DA.Fill(DS, "TB_Foto")
            connection.Close()

        Catch
            MsgBox("Erro ao buscar dados no Excel", MsgBoxStyle.Critical)
        End Try
    End Sub

    Public Function Buscar_Ultimo_ID()
        Try
            Dim leitor As SQLite.SQLiteDataReader
            Dim connection As New SQLite.SQLiteConnection(connstr)
            Dim cmd As New SQLite.SQLiteCommand
            connection.Open()
            cmd.Connection = connection
            cmd.CommandText = "select ID from [Inventario$] order by ID DESC;"
            leitor = cmd.ExecuteReader
            leitor.Read()
            Return leitor("ID")
            connection.Close()

        Catch
            MsgBox("Erro ao buscar dados no Excel", MsgBoxStyle.Critical)
            Return 0
        End Try
    End Function

    Public Sub Inserir_ID(ID As Integer)
        Try
            Dim connection As New SQLite.SQLiteConnection(connstr)
            Dim cmd As New SQLite.SQLiteCommand
            connection.Open()
            cmd.Connection = connection
            cmd.CommandText = "insert into [Inventario$] (ID) values('" & ID & "');"
            cmd.ExecuteNonQuery()
            connection.Close()

        Catch
            MsgBox("Erro ao inserir dados no Excel", MsgBoxStyle.Critical)
        End Try
    End Sub

    Public Sub Update_Inventario(ID As Integer, seq As String, local As String, odi As String, cod_ti As String, ti As String, bay As String,
                                 cod_tuc As String, tuc As String, cod_tipo_bem As String, tipo_bem As String, cod_uar As String, uar As String,
                                 cod_a2 As String, a2 As String, cod_a3 As String, a3 As String, cod_a4 As String, a4 As String, cod_a5 As String,
                                 a5 As String, cod_a6 As String, a6 As String, cod_cm1 As String, cm1 As String, cod_cm2 As String, cm2 As String,
                                 cod_cm3 As String, cm3 As String, descricao As String, fabricante As String, modelo As String, serie As String,
                                 obs As String, qtd As Decimal, um As String, ano As String, mes As String, dia As String, status As String,
                                 estado_bem As String, altura As Decimal, largura As Decimal, comprimento As Decimal, area As Decimal, pe As Decimal,
                                 obs_civil As String, consultor As String, lider As String, data_hora As String)
        Try
            Dim connection As New SQLite.SQLiteConnection(connstr)
            Dim cmd As New SQLite.SQLiteCommand
            connection.Open()
            cmd.Connection = connection
            cmd.CommandText = "update [Inventario$] set [Sequencial]='" & seq & "',[Local]='" & local & "',[ODI]='" & odi & "',[Código TI]='" &
                cod_ti & "',[TI]='" & ti & "',[Bay]='" & bay & "',[Código TUC]='" & cod_tuc & "',[Descrição TUC]='" & tuc & "',[Código Tipo de Bem]='" &
                cod_tipo_bem & "',[Descrição Tipo de Bem]='" & tipo_bem & "',[Código UAR]='" & cod_uar & "',[Descrição UAR]='" & uar &
                "',[Código A2]='" & cod_a2 & "',[Descrição A2]='" & a2 & "',[Código A3]='" & cod_a3 & "',[Descrição A3]='" & a3 & "',[Código A4]='" &
                cod_a4 & "',[Descrição A4]='" & a4 & "',[Código A5]='" & cod_a5 & "',[Descrição A5]='" & a5 & "',[Código A6]='" & cod_a6 &
                "',[Descrição A6]='" & a6 & "',[Código CM1]='" & cod_cm1 & "',[Descrição CM1]='" & cm1 & "',[Código CM2]='" & cod_cm2 &
                "',[Descrição CM2]='" & cm2 & "',[Código CM3]='" & cod_cm3 & "',[Descrição CM3]='" & cm3 & "',[Descrição]='" & descricao &
                "',[Fabricante]='" & fabricante & "',[Modelo]='" & modelo & "',[N° de Série]='" & serie & "',[Observação]='" & obs & "',[Quantidade]=" &
                qtd & ",[Unidade de Medida]='" & um & "',[Ano de Fabricação]='" & ano & "',[Mês de Fabricação]='" & mes & "',[Dia de Fabricação]='" & dia &
                "',[Status do Bem]='" & status & "',[Estado do Bem]='" & estado_bem & "',[Altura]=" & altura & ",[Largura]=" & largura & ",[Comprimento]=" &
                comprimento & ",[Área]=" & area & ",[Pé Direito]=" & pe & ",[Observacao Civil]='" & obs_civil & "',[Consultor]='" & consultor & "',[Líder]='" &
                lider & "',[Data/Hora]='" & data_hora & "' where [ID]=" & ID & ";"

            cmd.ExecuteNonQuery()
            connection.Close()
            MsgBox("Dados Salvos com Sucesso", MsgBoxStyle.Information)
        Catch
            MsgBox("Erro ao inserir dados no Excel", MsgBoxStyle.Critical)
        End Try
    End Sub

    Public Sub Update_Fotos(ID As Integer, FOTO_Query As String, Foto As String)
        Try
            Dim connection As New SQLite.SQLiteConnection(connstr)
            Dim cmd As New SQLite.SQLiteCommand
            connection.Open()
            cmd.Connection = connection
            cmd.CommandText = "update [Inventario$] set [" & FOTO_Query & "]='" & Foto & "' where [ID]=" & ID & ";"
            cmd.ExecuteNonQuery()
            connection.Close()

        Catch
            MsgBox("Erro ao inserir fotos no Excel", MsgBoxStyle.Critical)
        End Try
    End Sub

    Public Sub Consulta_Excel(DGV As DataGridView)
        Try
            Dim connection As New SQLite.SQLiteConnection(connstr)
            Dim DA As New SQLite.SQLiteDataAdapter
            Dim DT As New DataTable
            connection.Open()
            DA.SelectCommand = New SQLite.SQLiteCommand("select * from [Inventario$];", connection)
            DA.Fill(DT)
            connection.Close()
            DGV.DataSource = DT
        Catch
            MsgBox("Erro na consulta", MsgBoxStyle.Critical)
        End Try
    End Sub
    Public Sub Modelo_Excel(FRM As Form, SFD As SaveFileDialog)
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim Sh_T As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim salvar As String
        Try
            xlApp = New Excel.Application
            xlWorkBook = xlApp.Workbooks.Add(misValue)
            Sh_T = xlWorkBook.Sheets(1)
            Sh_T.Name = "Inventario"
            Sh_T.Range("a1").Value = "ID"
            Sh_T.Range("b1").Value = "Sequencial"
            Sh_T.Range("c1").Value = "Local"
            Sh_T.Range("d1").Value = "ODI"
            Sh_T.Range("e1").Value = "Código TI"
            Sh_T.Range("f1").Value = "TI"
            Sh_T.Range("g1").Value = "Bay"
            Sh_T.Range("h1").Value = "Código TUC"
            Sh_T.Range("i1").Value = "Descrição TUC"
            Sh_T.Range("j1").Value = "Código Tipo de Bem"
            Sh_T.Range("k1").Value = "Descrição Tipo de Bem"
            Sh_T.Range("l1").Value = "Código UAR"
            Sh_T.Range("m1").Value = "Descrição UAR"
            Sh_T.Range("n1").Value = "Código A2"
            Sh_T.Range("o1").Value = "Descrição A2"
            Sh_T.Range("p1").Value = "Código A3"
            Sh_T.Range("q1").Value = "Descrição A3"
            Sh_T.Range("r1").Value = "Código A4"
            Sh_T.Range("s1").Value = "Descrição A4"
            Sh_T.Range("t1").Value = "Código A5"
            Sh_T.Range("u1").Value = "Descrição A5"
            Sh_T.Range("v1").Value = "Código A6"
            Sh_T.Range("w1").Value = "Descrição A6"
            Sh_T.Range("x1").Value = "Código CM1"
            Sh_T.Range("y1").Value = "Descrição CM1"
            Sh_T.Range("z1").Value = "Código CM2"
            Sh_T.Range("aa1").Value = "Descrição CM2"
            Sh_T.Range("ab1").Value = "Código CM3"
            Sh_T.Range("ac1").Value = "Descrição CM3"
            Sh_T.Range("ad1").Value = "Descrição"
            Sh_T.Range("ae1").Value = "Fabricante"
            Sh_T.Range("af1").Value = "Modelo"
            Sh_T.Range("ag1").Value = "N° de Série"
            Sh_T.Range("ah1").Value = "Observação"
            Sh_T.Range("ai1").Value = "Quantidade"
            Sh_T.Range("aj1").Value = "Unidade de Medida"
            Sh_T.Range("ak1").Value = "Ano de Fabricação"
            Sh_T.Range("al1").Value = "Mês de Fabricação"
            Sh_T.Range("am1").Value = "Dia de Fabricação"
            Sh_T.Range("an1").Value = "Status do Bem"
            Sh_T.Range("ao1").Value = "Estado do Bem"
            Sh_T.Range("ap1").Value = "Altura"
            Sh_T.Range("aq1").Value = "Largura"
            Sh_T.Range("ar1").Value = "Comprimento"
            Sh_T.Range("as1").Value = "Área"
            Sh_T.Range("at1").Value = "Pé Direito"
            Sh_T.Range("au1").Value = "Observacao Civil"
            Sh_T.Range("av1").Value = "Consultor"
            Sh_T.Range("aw1").Value = "Líder"
            Sh_T.Range("ax1").Value = "Data/Hora"
            Sh_T.Range("ay1").Value = "Foto 1"
            Sh_T.Range("az1").Value = "Foto 2"
            Sh_T.Range("ba1").Value = "Foto 3"
            Sh_T.Range("bb1").Value = "Foto 4"
            Sh_T.Range("bc1").Value = "Foto 5"
            Sh_T.Range("bd1").Value = "Foto 6"
            Sh_T.Range("be1").Value = "Foto 7"
            Sh_T.Range("bf1").Value = "Foto 8"
            Sh_T.Range("bg1").Value = "Foto 9"
            Sh_T.Range("bh1").Value = "Foto 10"

            Sh_T.Columns.AutoFit()
            Sh_T.Range("a1:bh1").Font.Bold = True
            Sh_T.Range("a1:bh1").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray)

            Sh_T.Range("a2").Value = 1

            SFD.Title = "Selecione o local para Salvar em Excel"
            SFD.Filter = "Excel (*.xlsx)|*.xlsx"
            If SFD.ShowDialog = Windows.Forms.DialogResult.OK Then
                salvar = SFD.FileName
                xlWorkBook.SaveAs(salvar)
                xlWorkBook.Close()
            Else
                xlWorkBook.Close()
                MsgBox("O Excel não foi salvo. O software será fechado", MsgBoxStyle.Exclamation)
                FRM.Close()
            End If
        Catch
            MsgBox("Erro ao Carregar Excel!", MsgBoxStyle.Critical)
        End Try
    End Sub
End Class
