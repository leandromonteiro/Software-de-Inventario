Imports System.Net.Mail
Imports Microsoft.Office.Interop

Public Class Inventário_Excel
    Dim connstr As String = "Data Source=C:\Users\Public\INVENTARIO.db;;Version=3;New=True;Compress=True;Pooling=True"
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
        'Try
        Dim leitor As SQLite.SQLiteDataReader
            Dim connection As New SQLite.SQLiteConnection(connstr)
            Dim cmd As New SQLite.SQLiteCommand
            connection.Open()
            cmd.Connection = connection
        cmd.CommandText = "select desc_geral from ti_geral where cod_geral=(select cod_geral_todos from ti where cod='" & TI & "');"
        leitor = cmd.ExecuteReader
        leitor.Read()
        Dim a As String
        a = leitor("desc_geral")
        cmd.Dispose()
        connection.Close()
        connection.Dispose()
        Return a

        'Catch
        'MsgBox("Erro ao buscar dados ", MsgBoxStyle.Critical)
        'Return Nothing
        'End Try
    End Function

    Public Sub Consulta_Fotos(ID As Integer)
        'Try
        '    Dim leitor As SQLite.SQLiteDataReader
        '    Dim connection As New SQLite.SQLiteConnection(connstr)
        '    Dim cmd As New SQLite.SQLiteCommand
        '    connection.Open()
        '    cmd.Connection = connection
        '    cmd.CommandText = "select FOTO 1,FOTO 2,FOTO 3,FOTO 4,FOTO 5,FOTO 6,FOTO 7,FOTO 8,FOTO 9,FOTO 10 from [Inventario$] where ID=" & ID & ";"
        '    leitor = cmd.ExecuteReader
        '    leitor.Read()
        '    Foto1 = leitor("FOTO 1")
        '    Foto2 = leitor("FOTO 2")
        '    Foto3 = leitor("FOTO 3")
        '    Foto4 = leitor("FOTO 4")
        '    Foto5 = leitor("FOTO 5")
        '    Foto6 = leitor("FOTO 6")
        '    Foto7 = leitor("FOTO 7")
        '    Foto8 = leitor("FOTO 8")
        '    Foto9 = leitor("FOTO 9")
        '    Foto10 = leitor("FOTO 10")
        '    connection.Close()
        '    connection.Dispose()
        'Catch
        '    MsgBox("Erro ao buscar dados ", MsgBoxStyle.Critical)
        'End Try
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
    'cmd.Dispose()
    '        connection.Close()

    '    Catch
    '        MsgBox("Erro ao inserir foto ", MsgBoxStyle.Critical)
    '    End Try
    'End Sub

    Public Sub Consulta_TUC(cmb As ComboBox)
        cmb.Items.Clear()
        'Try
        Dim leitor As SQLite.SQLiteDataReader
        Dim connection As New SQLite.SQLiteConnection(connstr)
        Dim cmd As New SQLite.SQLiteCommand
        connection.Open()
        cmd.Connection = connection
        cmd.CommandText = "select descricao from tuc;"
        leitor = cmd.ExecuteReader

        Do While leitor.Read
            cmb.Items.Add(leitor("Descricao"))
        Loop
        cmd.Dispose()
        connection.Close()
        connection.Dispose()
        'Catch
        'MsgBox("Erro ao buscar dados ", MsgBoxStyle.Critical)
        'End Try
    End Sub

    Public Sub Consulta_TI_Geral(cmb As ComboBox)
        cmb.Items.Clear()
        Try
            Dim leitor As SQLite.SQLiteDataReader
            Dim connection As New SQLite.SQLiteConnection(connstr)
            Dim cmd As New SQLite.SQLiteCommand
            connection.Open()
            cmd.Connection = connection
            cmd.CommandText = "select desc_geral from ti_geral;"
            leitor = cmd.ExecuteReader
            Do While leitor.Read
                If Not IsDBNull(leitor("desc_geral")) Then
                    cmb.Items.Add(leitor("desc_geral"))
                End If
            Loop
            cmd.Dispose()
            connection.Close()
            connection.Dispose()
        Catch
            MsgBox("Erro ao buscar dados ", MsgBoxStyle.Critical)
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
            cmd.CommandText = "select descricao from ti where cod_geral_todos='" & cod_geral & "';"
            leitor = cmd.ExecuteReader
            Do While leitor.Read
                cmb.Items.Add(leitor("descricao"))
            Loop
            cmd.Dispose()
            connection.Close()
            connection.Dispose()
            GC.Collect()
        Catch
            MsgBox("Erro ao buscar dados ", MsgBoxStyle.Critical)
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
            cmd.CommandText = "select CM1,CM2,CM3 from CM;"
            leitor = cmd.ExecuteReader
            Do While leitor.Read
                If Not leitor("CM1") = "" Then
                    cmb1.Items.Add(leitor("CM1"))
                End If
                cmb2.Items.Add(leitor("CM2"))
                If Not leitor("CM3") = "" Then
                    cmb3.Items.Add(leitor("CM3"))
                End If
            Loop
            cmd.Dispose()
            connection.Close()
            connection.Dispose()
            GC.Collect()
        Catch
            MsgBox("Erro ao buscar dados ", MsgBoxStyle.Critical)
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
            cmd.CommandText = "select Descricao from UAR where TUC=" & TUC & ";"
            leitor = cmd.ExecuteReader
            Do While leitor.Read
                cmb.Items.Add(leitor("Descricao"))
            Loop
            cmd.Dispose()
            connection.Close()
            connection.Dispose()
            GC.Collect()
        Catch
            MsgBox("Erro ao buscar dados ", MsgBoxStyle.Critical)
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
            cmd.CommandText = "select Desc_A1 from RELACIONAMENTO_ATRIBUTOS where TUC='" & TUC & "';"
            leitor = cmd.ExecuteReader
            Do While leitor.Read
                cmb.Items.Add(leitor("Desc_A1"))
            Loop
            cmd.Dispose()
            connection.Close()
            connection.Dispose()
            GC.Collect()
        Catch
            MsgBox("Erro ao buscar dados ", MsgBoxStyle.Critical)
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
            cmd.CommandText = "select Descricao from DESC_TABELA where Cod_Tabela_Geral=(select Cod_Tabela from TABELA where Tabela=(select " &
            Desc_A & " from RELACIONAMENTO_ATRIBUTOS where TUC='" & TUC & "' and A1='" & A1 & "'));"
            leitor = cmd.ExecuteReader
            Do While leitor.Read
                N = 1
                cmb.Items.Add(leitor("Descricao"))
            Loop
            If N = 0 Then
                cmb.Enabled = False
            End If
            cmd.Dispose()
            connection.Close()
            connection.Dispose()
            GC.Collect()
        Catch
            MsgBox("Erro ao buscar dados ", MsgBoxStyle.Critical)
        End Try
    End Sub

    Public Sub Buscar_Tabela(Lbl As Label, TUC As Integer, A1 As String, Desc_A As String)
        Try
            Dim leitor As SQLite.SQLiteDataReader
            Dim connection As New SQLite.SQLiteConnection(connstr)
            Dim cmd As New SQLite.SQLiteCommand
            connection.Open()
            cmd.Connection = connection
            cmd.CommandText = "select " &
            Desc_A & " from RELACIONAMENTO_ATRIBUTOS where TUC='" & TUC & "' and A1='" & A1 & "';"
            leitor = cmd.ExecuteReader
            Do While leitor.Read
                Lbl.Text = leitor(Desc_A)
            Loop
            cmd.Dispose()
            connection.Close()
            connection.Dispose()
            GC.Collect()
        Catch
            MsgBox("Erro ao buscar dados ", MsgBoxStyle.Critical)
        End Try
    End Sub

    Public Function Buscar_A2_A6(cmb As ComboBox)
        Try
            Dim leitor As SQLite.SQLiteDataReader
            Dim connection As New SQLite.SQLiteConnection(connstr)
            Dim cmd As New SQLite.SQLiteCommand
            connection.Open()
            cmd.Connection = connection
            cmd.CommandText = "select Cod_Atributos from DESC_TABELA where Descricao='" & cmb.Text & "';"
            leitor = cmd.ExecuteReader
            leitor.Read()
            Dim a As String
            a = leitor("Cod_Atributos")
            cmd.Dispose()
            connection.Close()
            connection.Dispose()
            Return a

        Catch
            MsgBox("Erro ao buscar dados ", MsgBoxStyle.Critical)
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
            cmd.CommandText = "select Cod_Geral from ti_geral where Desc_Geral='" & cmb.Text & "';"
            leitor = cmd.ExecuteReader
            leitor.Read()
            Dim a As String
            a = leitor("Cod_Geral")
            cmd.Dispose()
            connection.Close()
            connection.Dispose()
            Return a

        Catch
            MsgBox("Erro ao buscar dados ", MsgBoxStyle.Critical)
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
            cmd.CommandText = "select Cod from TI where Cod_Geral_Todos='" & cod_geral_todos & "' and Descricao='" & cmb.Text & "';"
            leitor = cmd.ExecuteReader
            leitor.Read()
            Dim a As String
            a = leitor("Cod")
            cmd.Dispose()
            connection.Close()
            connection.Dispose()
            Return a
        Catch
            MsgBox("Erro ao buscar dados ", MsgBoxStyle.Critical)
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
            cmd.CommandText = "select Cod from UAR where TUC=" & TUC & " and Descricao='" & cmb.Text & "';"
            leitor = cmd.ExecuteReader
            leitor.Read()
            Dim a As String
            a = leitor("Cod")
            cmd.Dispose()
            connection.Close()
            connection.Dispose()
            Return a
        Catch
            MsgBox("Erro ao buscar dados ", MsgBoxStyle.Critical)
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
            cmd.CommandText = "select A1 from RELACIONAMENTO_ATRIBUTOS where TUC='" & TUC & "' and Desc_A1='" & cmb.Text & "';"
            leitor = cmd.ExecuteReader
            leitor.Read()
            Dim a As String
            a = leitor("A1")
            cmd.Dispose()
            connection.Close()
            connection.Dispose()
            Return a
        Catch
            MsgBox("Erro ao buscar dados ", MsgBoxStyle.Critical)
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
            cmd.CommandText = "select TUC from TUC where Descricao='" & cmb.Text & "';"
            leitor = cmd.ExecuteReader
            leitor.Read()
            Dim a As String
            a = leitor("TUC")
            cmd.Dispose()
            connection.Close()
            connection.Dispose()
            Return a
        Catch
            MsgBox("Erro ao buscar dados ", MsgBoxStyle.Critical)
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
            cmd.CommandText = "select C1 from CM where CM1='" & cmb.Text & "';"
            leitor = cmd.ExecuteReader
            leitor.Read()
            Dim a As String
            a = leitor("C1")
            cmd.Dispose()
            connection.Close()
            connection.Dispose()
            Return a
        Catch
            MsgBox("Erro ao buscar dados ", MsgBoxStyle.Critical)
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
            cmd.CommandText = "select C2 from CM where CM2='" & cmb.Text & "';"
            leitor = cmd.ExecuteReader
            leitor.Read()
            Dim a As String
            a = leitor("C2")
            cmd.Dispose()
            connection.Close()
            connection.Dispose()
            Return a
        Catch
            MsgBox("Erro ao buscar dados ", MsgBoxStyle.Critical)
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
            cmd.CommandText = "select C3 from CM where CM3='" & cmb.Text & "';"
            leitor = cmd.ExecuteReader
            leitor.Read()
            Dim a As String
            a = leitor("C3")
            cmd.Dispose()
            connection.Close()
            connection.Dispose()
            Return a
        Catch
            MsgBox("Erro ao buscar dados", MsgBoxStyle.Critical)
            Return 0
        End Try
    End Function

    Public Function Buscar_Ultimo_ID()
        Try
            Dim leitor As SQLite.SQLiteDataReader
            Dim connection As New SQLite.SQLiteConnection(connstr)
            Dim cmd As New SQLite.SQLiteCommand
            connection.Open()
            cmd.Connection = connection
            cmd.CommandText = "select ID from INVENTARIO order by ID DESC;"
            leitor = cmd.ExecuteReader
            leitor.Read()
            Dim a As String
            a = leitor("ID")
            cmd.Dispose()
            connection.Close()
            connection.Dispose()
            Return a
        Catch
            Return 0
        End Try
    End Function

    Public Sub Inserir_Dados(ID As Integer, Sequencial As String, Local As String, odi As String, cod_ti As Integer, ti As String, bay As String, cod_tuc As Integer, tuc As String,
                             cod_tipo_bem As String, tipo_bem As String, cod_uar As String, uar As String, cod_a2 As String, desc_a2 As String, cod_a3 As String, desc_a3 As String,
                             cod_a4 As String, desc_a4 As String, cod_a5 As String, desc_a5 As String, cod_a6 As String, desc_a6 As String, cod_cm1 As Integer, desc_cm1 As String,
                             cod_cm2 As Integer, desc_cm2 As String, cod_cm3 As Integer, desc_cm3 As String, descricao As String, fabricante As String, modelo As String, serie As String,
                             observacao As String, quantidade As Decimal, unidade As String, ano As Integer, mes As Integer, dia As Integer, status_bem As String, estado_bem As String,
                             altura As Decimal, largura As Decimal, comprimento As Decimal, area As Decimal, pe As Decimal, esforco As Decimal, obs_civil As String, consultor As String,
                             lider As String, Num_Manutencao As String, Foto As String)
        Try
            Using connection As New SQLite.SQLiteConnection(connstr)
                SQLite.SQLiteConnection.ClearAllPools()
                Dim cmd As New SQLite.SQLiteCommand
                If connection.State = ConnectionState.Closed Then
                    connection.Open()
                End If
                cmd.Connection = connection
                cmd.CommandText = "insert into INVENTARIO (ID,Sequencial,Local,odi,cod_ti,ti,bay,cod_tuc,desc_tuc,cod_tipo_bem,desc_tipo_bem,cod_uar,desc_uar,cod_a2,desc_a2," &
                                  "cod_a3,desc_a3,cod_a4,desc_a4,cod_a5,desc_a5,cod_a6,desc_a6,cod_cm1,desc_cm1,cod_cm2,desc_cm2,cod_cm3,desc_cm3,descricao,fabricante,modelo,serie," &
                                  "observacao,quantidade,unidade_medida,ano,mes,dia,status_bem,estado_bem,altura,largura,comprimento,area,pe_direito,esforco,obs_civil,consultor," &
                                  "lider,data_hora,Numero_Manutencao,foto) values(" & ID & ",'" & Sequencial & "','" & Local & "','" & odi & "'," & cod_ti & ",'" & ti & "','" & bay &
                                  "'," & cod_tuc & ",'" & tuc & "','" & cod_tipo_bem & "','" & tipo_bem & "', '" & cod_uar & "','" & uar & "','" & cod_a2 & "','" & desc_a2 & "','" &
                                  cod_a3 & "','" & desc_a3 & "','" & cod_a4 & "','" & desc_a4 & "','" & cod_a5 & "','" & desc_a5 & "','" & cod_a6 & "','" & desc_a6 & "'," & cod_cm1 &
                                  ",'" & desc_cm1 & "'," & cod_cm2 & ",'" & desc_cm2 & "'," & cod_cm3 & ",'" & desc_cm3 & "','" & descricao & "','" & fabricante & "','" & modelo & "','" &
                                  serie & "','" & observacao & "'," & quantidade & ", '" & unidade & "'," & ano & ", " & mes & "," & dia & ",'" & status_bem & "','" & estado_bem &
                                  "'," & altura & "," & largura & "," & comprimento & "," & area & "," & pe & "," & esforco & ",'" & obs_civil & "','" & consultor & "','" &
                                  lider & "','" & Now & "','" & Num_Manutencao & "','" & Foto & "');"
                cmd.ExecuteNonQuery()
                cmd.Dispose()
                connection.Close()
                connection.Dispose()
                GC.Collect()
            End Using
        Catch
            MsgBox("Erro ao inserir dados", MsgBoxStyle.Critical)
        End Try
    End Sub

    Public Sub Update_Inventario(ID As Integer, seq As String, local As String, odi As String, cod_ti As String, ti As String, bay As String,
                                 cod_tuc As String, tuc As String, cod_tipo_bem As String, tipo_bem As String, cod_uar As String, uar As String,
                                 cod_a2 As String, a2 As String, cod_a3 As String, a3 As String, cod_a4 As String, a4 As String, cod_a5 As String,
                                 a5 As String, cod_a6 As String, a6 As String, cod_cm1 As String, cm1 As String, cod_cm2 As String, cm2 As String,
                                 cod_cm3 As String, cm3 As String, descricao As String, fabricante As String, modelo As String, serie As String,
                                 obs As String, qtd As Decimal, um As String, ano As String, mes As String, dia As String, status As String,
                                 estado_bem As String, altura As Decimal, largura As Decimal, comprimento As Decimal, area As Decimal, pe As Decimal,
                                 obs_civil As String, foto As String, consultor As String, lider As String, TAG As String, esforco As Decimal)
        Try
            Using connection As New SQLite.SQLiteConnection(connstr)
                SQLite.SQLiteConnection.ClearAllPools()
                Dim cmd As New SQLite.SQLiteCommand
                If connection.State = ConnectionState.Closed Then
                    connection.Open()
                End If
                cmd.Connection = connection
                cmd.CommandText = "update Inventario set Sequencial='" & seq & "',Local='" & local & "',ODI='" & odi & "',cod_ti=" &
                    cod_ti & ",ti='" & ti & "',Bay='" & bay & "',cod_tuc='" & cod_tuc & "',desc_tuc='" & tuc & "',cod_tipo_bem='" &
                    cod_tipo_bem & "',desc_tipo_bem='" & tipo_bem & "',cod_uar='" & cod_uar & "',desc_uar='" & uar &
                    "',Cod_A2='" & cod_a2 & "',desc_A2='" & a2 & "',Cod_A3='" & cod_a3 & "',desc_A3='" & a3 & "',cod_A4='" &
                    cod_a4 & "',Desc_A4='" & a4 & "',Cod_A5='" & cod_a5 & "',Desc_A5='" & a5 & "',Cod_A6='" & cod_a6 &
                    "',Desc_A6='" & a6 & "',Cod_CM1='" & cod_cm1 & "',Desc_CM1='" & cm1 & "',Cod_CM2='" & cod_cm2 &
                    "',Desc_CM2='" & cm2 & "',Cod_CM3='" & cod_cm3 & "',Desc_CM3='" & cm3 & "',Descricao='" & descricao &
                    "',Fabricante='" & fabricante & "',Modelo='" & modelo & "',serie='" & serie & "',Observacao='" & obs & "',Quantidade=" &
                    qtd & ",Unidade_Medida='" & um & "',Ano='" & ano & "',Mes='" & mes & "',Dia='" & dia &
                    "',Status_Bem='" & status & "',Estado_Bem='" & estado_bem & "',Altura=" & altura & ",Largura=" & largura & ",Comprimento=" &
                    comprimento & ",area=" & area & ",Pe_direito=" & pe & ",Obs_Civil='" & obs_civil & "',foto='" & foto & "',Consultor='" & consultor & "',Lider='" &
                    lider & "',Data_Hora='" & Now & "',Numero_Manutencao='" & TAG & "',esforco=" & esforco & " where ID=" & ID & ";"

                cmd.ExecuteNonQuery()
                cmd.Dispose()
                connection.Close()
                connection.Dispose()
            End Using
            'MsgBox("Dados Salvos com Sucesso", MsgBoxStyle.Information)
        Catch
            MsgBox("Erro ao atualizar dados", MsgBoxStyle.Critical)
        End Try
    End Sub
    Public Sub Excluir_Tudo()
        Try
            Using connection As New SQLite.SQLiteConnection(connstr)
                SQLite.SQLiteConnection.ClearAllPools()
                Dim cmd As New SQLite.SQLiteCommand
                If connection.State = ConnectionState.Closed Then
                    connection.Open()
                End If
                cmd.Connection = connection
                cmd.CommandText = "Delete from INVENTARIO;"
                cmd.ExecuteNonQuery()
                cmd.Dispose()
                connection.Close()
                connection.Dispose()
            End Using
            MsgBox("Dados Excluídos com Sucesso", MsgBoxStyle.Information)
        Catch
            MsgBox("Erro ao excluir dados", MsgBoxStyle.Critical)
        End Try
    End Sub

    Public Sub Excluir(ID As Integer)
        Try
            Using connection As New SQLite.SQLiteConnection(connstr)
                SQLite.SQLiteConnection.ClearAllPools()
                Dim cmd As New SQLite.SQLiteCommand
                If connection.State = ConnectionState.Closed Then
                    connection.Open()
                End If
                cmd.Connection = connection
                cmd.CommandText = "Delete from INVENTARIO where ID=" & ID & ";"
                cmd.ExecuteNonQuery()
                cmd.Dispose()
                connection.Close()
                connection.Dispose()
            End Using
            MsgBox("Dados Excluídos com Sucesso", MsgBoxStyle.Information)
        Catch
            MsgBox("Erro ao excluir dados", MsgBoxStyle.Critical)
        End Try
    End Sub

    Public Sub Consulta_Excel(DGV As DataGridView)
        Try
            Dim connection As New SQLite.SQLiteConnection(connstr)
            Dim DA As New SQLite.SQLiteDataAdapter
            Dim DT As New DataTable
            connection.Open()
            DA.SelectCommand = New SQLite.SQLiteCommand("select * from Inventario;", connection)
            DA.Fill(DT)
            connection.Close()
            connection.Dispose()
            DA.Dispose()
            GC.Collect()
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
            Sh_T.Range("ay1").Value = "Foto"

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
