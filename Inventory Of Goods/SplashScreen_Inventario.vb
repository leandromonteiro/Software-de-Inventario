Public NotInheritable Class SplashScreen_Inventario

    'TODO: Este formulário pode ser facilmente configurado como a tela inicial da aplicação através da edição da aba "Aplicação"
    '  no Designer de Projeto ("Propriedades" dentro do menu "Projetos").


    Private Sub SplashScreen_Inventario_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ''Título da Aplicação
        'If My.Application.Info.Title <> "" Then
        '    ApplicationTitle.Text = My.Application.Info.Title
        'Else
        '    'Se o título da aplicação estiver faltando, utiliza o nome da aplicação sem a extensão
        '    ApplicationTitle.Text = System.IO.Path.GetFileNameWithoutExtension(My.Application.Info.AssemblyName)
        'End If

        Version.Text = System.String.Format(Version.Text, My.Application.Info.Version.Major, My.Application.Info.Version.Minor)

        'Informação de Copyright
        'Copyright.Text = My.Application.Info.Copyright

        'Limite de tempo
        If Today.Day >= 1 And Today.Month >= 6 And Today.Year >= 2018 Then
            MsgBox("Tempo de teste do software expirado.", vbCritical)
            Application.Exit()
        End If

    End Sub
End Class
