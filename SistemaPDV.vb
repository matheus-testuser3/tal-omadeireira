Imports System.Windows.Forms
Imports System.Drawing
Imports System.Configuration

''' <summary>
''' Formulário principal do Sistema PDV - Madeireira Maria Luiza
''' Interface moderna com menu lateral e acesso às funcionalidades
''' </summary>
Public Class MainForm
    Inherits Form

    ' Controles da interface
    Private WithEvents pnlSidebar As Panel
    Private WithEvents pnlMain As Panel
    Private WithEvents pnlHeader As Panel
    Private WithEvents lblTitle As Label
    Private WithEvents lblSubtitle As Label
    Private WithEvents btnGerarTalao As Button
    Private WithEvents btnRelatorios As Button
    Private WithEvents btnConfiguracoes As Button
    Private WithEvents btnSobre As Button
    Private WithEvents btnSair As Button
    Private WithEvents picLogo As PictureBox

    ' Dados da madeireira
    Private ReadOnly nomeMadeireira As String = ConfigurationManager.AppSettings("NomeMadeireira")
    Private ReadOnly enderecoMadeireira As String = ConfigurationManager.AppSettings("EnderecoMadeireira")

    ''' <summary>
    ''' Construtor do formulário principal
    ''' </summary>
    Public Sub New()
        InitializeComponent()
        ConfigurarInterface()
    End Sub

    ''' <summary>
    ''' Inicializa os componentes da interface
    ''' </summary>
    Private Sub InitializeComponent()
        ' Configurações do formulário
        Me.Text = "Sistema PDV - " & nomeMadeireira
        Me.Size = New Size(1200, 800)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.BackColor = Color.WhiteSmoke
        Me.Font = New Font("Segoe UI", 10.0F, FontStyle.Regular)

        ' Painel lateral (sidebar)
        pnlSidebar = New Panel()
        pnlSidebar.Size = New Size(250, Me.Height)
        pnlSidebar.Dock = DockStyle.Left
        pnlSidebar.BackColor = Color.FromArgb(41, 53, 65)
        Me.Controls.Add(pnlSidebar)

        ' Painel principal
        pnlMain = New Panel()
        pnlMain.Dock = DockStyle.Fill
        pnlMain.BackColor = Color.WhiteSmoke
        pnlMain.Padding = New Padding(20)
        Me.Controls.Add(pnlMain)

        ' Header do painel principal
        pnlHeader = New Panel()
        pnlHeader.Size = New Size(pnlMain.Width - 40, 120)
        pnlHeader.BackColor = Color.White
        pnlHeader.Location = New Point(20, 20)
        pnlHeader.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        pnlMain.Controls.Add(pnlHeader)

        ' Logo da madeireira
        picLogo = New PictureBox()
        picLogo.Size = New Size(80, 80)
        picLogo.Location = New Point(200, 10)
        picLogo.BackColor = Color.LightGray
        picLogo.BorderStyle = BorderStyle.FixedSingle
        pnlSidebar.Controls.Add(picLogo)

        ' Título principal
        lblTitle = New Label()
        lblTitle.Text = nomeMadeireira
        lblTitle.Font = New Font("Segoe UI", 18.0F, FontStyle.Bold)
        lblTitle.ForeColor = Color.FromArgb(52, 73, 94)
        lblTitle.Size = New Size(600, 40)
        lblTitle.Location = New Point(20, 20)
        pnlHeader.Controls.Add(lblTitle)

        ' Subtítulo
        lblSubtitle = New Label()
        lblSubtitle.Text = "Sistema de Ponto de Venda Integrado com Geração Automática de Talões"
        lblSubtitle.Font = New Font("Segoe UI", 11.0F, FontStyle.Regular)
        lblSubtitle.ForeColor = Color.FromArgb(127, 140, 141)
        lblSubtitle.Size = New Size(600, 25)
        lblSubtitle.Location = New Point(20, 65)
        pnlHeader.Controls.Add(lblSubtitle)

        ' Botão Gerar Talão (principal)
        btnGerarTalao = New Button()
        btnGerarTalao.Text = "🧾 GERAR TALÃO (F2)"
        btnGerarTalao.Size = New Size(200, 50)
        btnGerarTalao.Location = New Point(25, 120)
        btnGerarTalao.BackColor = Color.FromArgb(46, 204, 113)
        btnGerarTalao.ForeColor = Color.White
        btnGerarTalao.Font = New Font("Segoe UI", 12.0F, FontStyle.Bold)
        btnGerarTalao.FlatStyle = FlatStyle.Flat
        btnGerarTalao.FlatAppearance.BorderSize = 0
        btnGerarTalao.Cursor = Cursors.Hand
        pnlSidebar.Controls.Add(btnGerarTalao)

        ' Botão Relatórios
        btnRelatorios = New Button()
        btnRelatorios.Text = "📊 RELATÓRIOS (F5)"
        btnRelatorios.Size = New Size(200, 40)
        btnRelatorios.Location = New Point(25, 180)
        btnRelatorios.BackColor = Color.FromArgb(52, 152, 219)
        btnRelatorios.ForeColor = Color.White
        btnRelatorios.Font = New Font("Segoe UI", 10.0F, FontStyle.Regular)
        btnRelatorios.FlatStyle = FlatStyle.Flat
        btnRelatorios.FlatAppearance.BorderSize = 0
        btnRelatorios.Cursor = Cursors.Hand
        pnlSidebar.Controls.Add(btnRelatorios)

        ' Botão Configurações
        btnConfiguracoes = New Button()
        btnConfiguracoes.Text = "⚙️ Configurações"
        btnConfiguracoes.Size = New Size(200, 40)
        btnConfiguracoes.Location = New Point(25, 230)
        btnConfiguracoes.BackColor = Color.FromArgb(52, 73, 94)
        btnConfiguracoes.ForeColor = Color.White
        btnConfiguracoes.Font = New Font("Segoe UI", 10.0F, FontStyle.Regular)
        btnConfiguracoes.FlatStyle = FlatStyle.Flat
        btnConfiguracoes.FlatAppearance.BorderSize = 0
        btnConfiguracoes.Cursor = Cursors.Hand
        pnlSidebar.Controls.Add(btnConfiguracoes)

        ' Botão Sobre
        btnSobre = New Button()
        btnSobre.Text = "ℹ️ Sobre o Sistema"
        btnSobre.Size = New Size(200, 40)
        btnSobre.Location = New Point(25, 280)
        btnSobre.BackColor = Color.FromArgb(52, 73, 94)
        btnSobre.ForeColor = Color.White
        btnSobre.Font = New Font("Segoe UI", 10.0F, FontStyle.Regular)
        btnSobre.FlatStyle = FlatStyle.Flat
        btnSobre.FlatAppearance.BorderSize = 0
        btnSobre.Cursor = Cursors.Hand
        pnlSidebar.Controls.Add(btnSobre)

        ' Botão Sair
        btnSair = New Button()
        btnSair.Text = "🚪 Sair"
        btnSair.Size = New Size(200, 40)
        btnSair.Location = New Point(25, Me.Height - 80)
        btnSair.BackColor = Color.FromArgb(231, 76, 60)
        btnSair.ForeColor = Color.White
        btnSair.Font = New Font("Segoe UI", 10.0F, FontStyle.Regular)
        btnSair.FlatStyle = FlatStyle.Flat
        btnSair.FlatAppearance.BorderSize = 0
        btnSair.Cursor = Cursors.Hand
        btnSair.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
        pnlSidebar.Controls.Add(btnSair)
    End Sub

    ''' <summary>
    ''' Configura detalhes adicionais da interface
    ''' </summary>
    Private Sub ConfigurarInterface()
        ' Adicionar informações da madeireira no painel principal
        Dim lblInfo As New Label()
        lblInfo.Text = enderecoMadeireira & vbCrLf & 
                      "📞 " & ConfigurationManager.AppSettings("TelefoneMadeireira") & vbCrLf &
                      "📋 CNPJ: " & ConfigurationManager.AppSettings("CNPJMadeireira")
        lblInfo.Font = New Font("Segoe UI", 10.0F, FontStyle.Regular)
        lblInfo.ForeColor = Color.FromArgb(127, 140, 141)
        lblInfo.AutoSize = True
        lblInfo.Location = New Point(20, 160)
        pnlMain.Controls.Add(lblInfo)

        ' Adicionar instruções
        Dim lblInstrucoes As New Label()
        lblInstrucoes.Text = "📋 INSTRUÇÕES DE USO:" & vbCrLf & vbCrLf &
                            "1. Clique em 'GERAR TALÃO' (F2) para nova venda" & vbCrLf &
                            "2. Preencha os dados do cliente e produtos" & vbCrLf &
                            "3. O sistema irá gerar e imprimir automaticamente" & vbCrLf &
                            "4. Use 'RELATÓRIOS' (F5) para consultar vendas" & vbCrLf & vbCrLf &
                            "⌨️ ATALHOS DE TECLADO:" & vbCrLf &
                            "• F2 = Nova Venda  • F5 = Relatórios" & vbCrLf &
                            "• F1 = Sobre  • ESC = Sair" & vbCrLf & vbCrLf &
                            "✅ Sistema profissional com logs e backup automático" & vbCrLf &
                            "✅ Validação inteligente de dados" & vbCrLf &
                            "✅ Histórico completo de vendas" & vbCrLf &
                            "✅ Todo o processo é automático e seguro!"
        lblInstrucoes.Font = New Font("Segoe UI", 11.0F, FontStyle.Regular)
        lblInstrucoes.ForeColor = Color.FromArgb(52, 73, 94)
        lblInstrucoes.Size = New Size(700, 300)
        lblInstrucoes.Location = New Point(20, 250)
        pnlMain.Controls.Add(lblInstrucoes)

        ' Efeitos visuais nos botões
        AdicionarEfeitosBotoes()
    End Sub

    ''' <summary>
    ''' Adiciona efeitos visuais aos botões (hover, etc.)
    ''' </summary>
    Private Sub AdicionarEfeitosBotoes()
        ' Efeito hover para o botão principal
        AddHandler btnGerarTalao.MouseEnter, Sub() btnGerarTalao.BackColor = Color.FromArgb(39, 174, 96)
        AddHandler btnGerarTalao.MouseLeave, Sub() btnGerarTalao.BackColor = Color.FromArgb(46, 204, 113)

        ' Efeito hover para outros botões
        AddHandler btnRelatorios.MouseEnter, Sub() btnRelatorios.BackColor = Color.FromArgb(41, 128, 185)
        AddHandler btnRelatorios.MouseLeave, Sub() btnRelatorios.BackColor = Color.FromArgb(52, 152, 219)
        
        AddHandler btnConfiguracoes.MouseEnter, Sub() btnConfiguracoes.BackColor = Color.FromArgb(44, 62, 80)
        AddHandler btnConfiguracoes.MouseLeave, Sub() btnConfiguracoes.BackColor = Color.FromArgb(52, 73, 94)

        AddHandler btnSobre.MouseEnter, Sub() btnSobre.BackColor = Color.FromArgb(44, 62, 80)
        AddHandler btnSobre.MouseLeave, Sub() btnSobre.BackColor = Color.FromArgb(52, 73, 94)

        AddHandler btnSair.MouseEnter, Sub() btnSair.BackColor = Color.FromArgb(192, 57, 43)
        AddHandler btnSair.MouseLeave, Sub() btnSair.BackColor = Color.FromArgb(231, 76, 60)
    End Sub

    ''' <summary>
    ''' Evento click do botão Gerar Talão - função principal do sistema
    ''' </summary>
    Private Sub btnGerarTalao_Click(sender As Object, e As EventArgs) Handles btnGerarTalao.Click
        Try
            ' Abrir formulário de entrada de dados
            Dim formPDV As New FormPDV()
            If formPDV.ShowDialog() = DialogResult.OK Then
                ' Os dados foram preenchidos, processar talão
                ProcessarTalao(formPDV.DadosColetados)
            End If
        Catch ex As Exception
            MessageBox.Show("Erro ao abrir formulário de entrada de dados:" & vbCrLf & ex.Message, 
                          "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' Processa a geração do talão com integração Excel/VBA
    ''' </summary>
    Private Sub ProcessarTalao(dados As DadosTalao)
        Try
            ' Log do início do processamento
            Logger.Instance.Info($"Iniciando processamento de talão para cliente: {dados.NomeCliente}")
            
            ' Validar dados usando novo sistema
            Dim erros = CompatibilityAdapter.ValidarDadosTalao(dados)
            If erros.Count > 0 Then
                Dim mensagemErro = "Erros encontrados:" & vbCrLf & String.Join(vbCrLf, erros)
                MessageBox.Show(mensagemErro, "Dados Inválidos", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Logger.Instance.Warning($"Dados inválidos para talão: {String.Join(", ", erros)}")
                Return
            End If
            
            ' Formatar dados automaticamente
            CompatibilityAdapter.FormatarDadosCliente(dados)
            
            ' Converter para nova arquitetura
            Dim venda = CompatibilityAdapter.ConvertToVenda(dados)
            If venda Is Nothing Then
                MessageBox.Show("Erro ao processar dados da venda.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If
            
            ' Mostrar loading com melhor feedback
            Dim loading As New Form()
            loading.Text = "Processando..."
            loading.Size = New Size(450, 200)
            loading.StartPosition = FormStartPosition.CenterParent
            loading.FormBorderStyle = FormBorderStyle.FixedDialog
            loading.MaximizeBox = False
            loading.MinimizeBox = False

            Dim lblLoading As New Label()
            lblLoading.Text = "🔄 Gerando talão com sistema otimizado..." & vbCrLf & 
                             "• Validando dados" & vbCrLf &
                             "• Iniciando Excel em segundo plano" & vbCrLf &
                             "• Criando template profissional" & vbCrLf &
                             "• Preenchendo dados do cliente" & vbCrLf &
                             "• Configurando impressão" & vbCrLf &
                             "• Executando impressão automática"
            lblLoading.AutoSize = True
            lblLoading.Location = New Point(20, 20)
            lblLoading.Font = New Font("Segoe UI", 10.0F)
            loading.Controls.Add(lblLoading)

            loading.Show()
            Application.DoEvents()

            ' Usar novo serviço para processar venda
            Dim vendaService = New VendaService()
            Dim sucesso = vendaService.ProcessarVenda(venda)

            loading.Close()

            If sucesso Then
                ' Sucesso com estatísticas
                MessageBox.Show("✅ Talão gerado e impresso com sucesso!" & vbCrLf & vbCrLf &
                              $"Talão: {venda.NumeroTalao}" & vbCrLf &
                              $"Cliente: {venda.Cliente.Nome}" & vbCrLf &
                              $"Produtos: {venda.Itens.Count}" & vbCrLf &
                              $"Valor Total: {venda.ValorTotal:C}" & vbCrLf &
                              $"Vendedor: {venda.Vendedor}",
                              "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                              
                Logger.Instance.Audit("TALAO_GERADO_SUCESSO", 
                    $"Talão: {venda.NumeroTalao}, Cliente: {venda.Cliente.Nome}, Valor: {venda.ValorTotal:C}",
                    venda.Vendedor)
            Else
                MessageBox.Show("❌ Erro ao gerar talão." & vbCrLf & vbCrLf &
                              "Verifique:" & vbCrLf &
                              "• Se o Microsoft Excel está instalado" & vbCrLf &
                              "• Se há uma impressora configurada" & vbCrLf &
                              "• Os logs do sistema para mais detalhes",
                              "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                              
                Logger.Instance.Error($"Falha ao processar venda {venda.NumeroTalao}")
            End If

        Catch ex As Exception
            Logger.Instance.Error("Erro crítico ao processar talão", ex)
            MessageBox.Show("❌ Erro crítico ao gerar talão:" & vbCrLf & vbCrLf & ex.Message & vbCrLf & vbCrLf &
                          "O erro foi registrado nos logs do sistema.",
                          "Erro Crítico", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' Evento click do botão Relatórios
    ''' </summary>
    Private Sub btnRelatorios_Click(sender As Object, e As EventArgs) Handles btnRelatorios.Click
        Try
            Logger.Instance.Info("Abrindo formulário de relatórios")
            Dim formRelatorios = New RelatoriosForm()
            formRelatorios.ShowDialog(Me)
        Catch ex As Exception
            Logger.Instance.Error("Erro ao abrir relatórios", ex)
            MessageBox.Show("Erro ao abrir relatórios: " & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' Evento click do botão Configurações
    ''' </summary>
    Private Sub btnConfiguracoes_Click(sender As Object, e As EventArgs) Handles btnConfiguracoes.Click
        MessageBox.Show("🔧 Módulo de Configurações" & vbCrLf & vbCrLf &
                       "Em desenvolvimento. Funcionalidades planejadas:" & vbCrLf &
                       "• Configuração de impressora padrão" & vbCrLf &
                       "• Dados da madeireira" & vbCrLf &
                       "• Layout do talão" & vbCrLf &
                       "• Produtos cadastrados",
                       "Configurações", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    ''' <summary>
    ''' Evento click do botão Sobre
    ''' </summary>
    Private Sub btnSobre_Click(sender As Object, e As EventArgs) Handles btnSobre.Click
        MessageBox.Show("📋 Sistema PDV - " & nomeMadeireira & vbCrLf & vbCrLf &
                       "Versão: 1.0.0" & vbCrLf &
                       "Desenvolvido por: matheus-testuser3" & vbCrLf & vbCrLf &
                       "🎯 Características:" & vbCrLf &
                       "• Interface moderna em VB.NET" & vbCrLf &
                       "• Integração automática com Excel" & vbCrLf &
                       "• Geração de talões profissionais" & vbCrLf &
                       "• Execução de VBA incorporado" & vbCrLf &
                       "• Impressão automática" & vbCrLf & vbCrLf &
                       "© 2024 - Todos os direitos reservados",
                       "Sobre o Sistema", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    ''' <summary>
    ''' Evento click do botão Sair
    ''' </summary>
    Private Sub btnSair_Click(sender As Object, e As EventArgs) Handles btnSair.Click
        If MessageBox.Show("Tem certeza que deseja sair do sistema?", 
                          "Confirmar Saída", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            Application.Exit()
        End If
    End Sub

    ''' <summary>
    ''' Evento de carregamento do formulário
    ''' </summary>
    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Verificar se Excel está instalado
        Try
            Dim excel As Object = CreateObject("Excel.Application")
            excel.Quit()
            excel = Nothing
        Catch ex As Exception
            MessageBox.Show("⚠️ ATENÇÃO: Microsoft Excel não foi detectado!" & vbCrLf & vbCrLf &
                          "O sistema PDV requer o Microsoft Excel para funcionar." & vbCrLf &
                          "Por favor, instale o Microsoft Excel e reinicie o sistema." & vbCrLf & vbCrLf &
                          "Erro: " & ex.Message,
                          "Excel Não Encontrado", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub
    
    ''' <summary>
    ''' Processa atalhos de teclado
    ''' </summary>
    Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean
        Try
            Select Case keyData
                Case Keys.F2
                    ' F2 = Nova Venda
                    btnGerarTalao_Click(Nothing, Nothing)
                    Return True
                Case Keys.F5
                    ' F5 = Relatórios
                    btnRelatorios_Click(Nothing, Nothing)
                    Return True
                Case Keys.F1
                    ' F1 = Sobre
                    btnSobre_Click(Nothing, Nothing)
                    Return True
                Case Keys.Alt Or Keys.F4
                    ' Alt+F4 = Sair
                    btnSair_Click(Nothing, Nothing)
                    Return True
                Case Keys.Escape
                    ' ESC = Sair com confirmação
                    btnSair_Click(Nothing, Nothing)
                    Return True
            End Select
        Catch ex As Exception
            Logger.Instance.Error("Erro ao processar atalho de teclado", ex)
        End Try
        
        Return MyBase.ProcessCmdKey(msg, keyData)
    End Function
End Class

''' <summary>
''' Estrutura de dados para o talão
''' </summary>
Public Class DadosTalao
    Public Property NomeCliente As String
    Public Property EnderecoCliente As String
    Public Property CEP As String
    Public Property Cidade As String
    Public Property Telefone As String
    Public Property Produtos As List(Of ProdutoTalao)
    Public Property FormaPagamento As String
    Public Property Vendedor As String
    Public Property DataVenda As Date
    Public Property NumeroTalao As String

    Public Sub New()
        Produtos = New List(Of ProdutoTalao)()
        DataVenda = Date.Now
        NumeroTalao = Date.Now.ToString("yyyyMMddHHmmss")
    End Sub
End Class

''' <summary>
''' Estrutura de dados para produtos do talão
''' </summary>
Public Class ProdutoTalao
    Public Property Descricao As String
    Public Property Quantidade As Double
    Public Property Unidade As String
    Public Property PrecoUnitario As Double
    Public Property PrecoTotal As Double

    Public Sub New()
        Unidade = "UN"
    End Sub
End Class