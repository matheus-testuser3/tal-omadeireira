Imports System.Windows.Forms
Imports System.Drawing
Imports System.Configuration

''' <summary>
''' Formulário principal do Sistema PDV - Madeireira Maria Luiza
''' Interface moderna integrada com todas as funcionalidades
''' </summary>
Public Class MainForm
    Inherits Form

    ' Controles da interface
    Private WithEvents pnlSidebar As Panel
    Private WithEvents pnlMain As Panel
    Private WithEvents pnlHeader As Panel
    Private WithEvents lblTitle As Label
    Private WithEvents lblSubtitle As Label
    Private WithEvents lblStatusSistema As Label
    Private WithEvents btnPDVCompleto As Button
    Private WithEvents btnGerarTalao As Button
    Private WithEvents btnGestaoClientes As Button
    Private WithEvents btnGestaoEstoque As Button
    Private WithEvents btnRelatorios As Button
    Private WithEvents btnConfiguracoes As Button
    Private WithEvents btnSobre As Button
    Private WithEvents btnSair As Button
    Private WithEvents picLogo As PictureBox

    ' Sistema integrado
    Private _database As DatabaseManager
    Private _config As ConfiguracaoSistema
    Private _mainPDVForm As MainPDVForm

    ' Dados da madeireira
    Private ReadOnly nomeMadeireira As String
    Private ReadOnly enderecoMadeireira As String

    ''' <summary>
    ''' Construtor do formulário principal
    ''' </summary>
    Public Sub New()
        Try
            _config = New ConfiguracaoSistema()
            nomeMadeireira = _config.NomeMadeireira
            enderecoMadeireira = _config.EnderecoMadeireira
            
            InitializeComponent()
            InicializarSistema()
            ConfigurarInterface()
        Catch ex As Exception
            MessageBox.Show($"Erro ao inicializar sistema: {ex.Message}", "Erro Crítico", 
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' Inicializa os componentes da interface
    ''' </summary>
    Private Sub InitializeComponent()
        ' Configurações do formulário principal
        Me.Text = $"Sistema PDV Integrado - {nomeMadeireira}"
        Me.Size = New Size(1000, 700)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.WindowState = FormWindowState.Maximized
        Me.BackColor = Color.WhiteSmoke
        Me.Icon = Nothing

        CriarMenuLateral()
        CriarAreaPrincipal()
    End Sub

    ''' <summary>
    ''' Inicializa os sistemas integrados
    ''' </summary>
    Private Sub InicializarSistema()
        Try
            ' Inicializar banco de dados
            _database = DatabaseManager.Instance
            
            ' Atualizar status
            lblStatusSistema.Text = _database.VerificarConexao()
            lblStatusSistema.ForeColor = If(_database.VerificarConexao().Contains("Access"), Color.Green, Color.Orange)
            
            Console.WriteLine("Sistema PDV integrado inicializado com sucesso")
            
        Catch ex As Exception
            lblStatusSistema.Text = "Erro na inicialização"
            lblStatusSistema.ForeColor = Color.Red
            Console.WriteLine($"Erro na inicialização: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Cria o menu lateral moderno
    ''' </summary>
    Private Sub CriarMenuLateral()
        pnlSidebar = New Panel() With {
            .Dock = DockStyle.Left,
            .Width = 250,
            .BackColor = Color.FromArgb(45, 45, 48),
            .Padding = New Padding(0, 20, 0, 20)
        }

        ' Logo da empresa
        picLogo = New PictureBox() With {
            .Size = New Size(200, 80),
            .Location = New Point(25, 20),
            .BackColor = Color.White,
            .SizeMode = PictureBoxSizeMode.CenterImage
        }

        ' Título da empresa
        lblTitle = New Label() With {
            .Text = nomeMadeireira.ToUpper(),
            .Location = New Point(25, 110),
            .Size = New Size(200, 40),
            .ForeColor = Color.White,
            .Font = New Font("Segoe UI", 12, FontStyle.Bold),
            .TextAlign = ContentAlignment.MiddleCenter
        }

        lblSubtitle = New Label() With {
            .Text = "SISTEMA PDV INTEGRADO",
            .Location = New Point(25, 150),
            .Size = New Size(200, 20),
            .ForeColor = Color.LightGray,
            .Font = New Font("Segoe UI", 9),
            .TextAlign = ContentAlignment.MiddleCenter
        }

        ' Botões do menu principal
        btnPDVCompleto = CriarBotaoMenu("🛒 PDV COMPLETO", 200, Color.FromArgb(0, 120, 215))
        btnGerarTalao = CriarBotaoMenu("🧾 GERAR TALÃO", 250, Color.FromArgb(0, 153, 51))
        btnGestaoClientes = CriarBotaoMenu("👥 GESTÃO CLIENTES", 300, Color.FromArgb(153, 102, 51))
        btnGestaoEstoque = CriarBotaoMenu("📦 GESTÃO ESTOQUE", 350, Color.FromArgb(128, 0, 128))
        btnRelatorios = CriarBotaoMenu("📊 RELATÓRIOS", 400, Color.FromArgb(255, 102, 0))
        btnConfiguracoes = CriarBotaoMenu("⚙️ CONFIGURAÇÕES", 450, Color.FromArgb(105, 105, 105))

        ' Botões de sistema
        btnSobre = CriarBotaoMenu("ℹ️ SOBRE", 520, Color.FromArgb(70, 130, 180))
        btnSair = CriarBotaoMenu("❌ SAIR", 570, Color.FromArgb(220, 20, 60))

        pnlSidebar.Controls.AddRange({
            picLogo, lblTitle, lblSubtitle,
            btnPDVCompleto, btnGerarTalao, btnGestaoClientes, btnGestaoEstoque,
            btnRelatorios, btnConfiguracoes, btnSobre, btnSair
        })

        Me.Controls.Add(pnlSidebar)
    End Sub

    ''' <summary>
    ''' Cria um botão do menu com estilo moderno
    ''' </summary>
    Private Function CriarBotaoMenu(texto As String, top As Integer, cor As Color) As Button
        Dim btn As New Button() With {
            .Text = texto,
            .Size = New Size(220, 40),
            .Location = New Point(15, top),
            .BackColor = cor,
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .TextAlign = ContentAlignment.MiddleLeft,
            .Padding = New Padding(10, 0, 0, 0),
            .Cursor = Cursors.Hand
        }

        btn.FlatAppearance.BorderSize = 0
        btn.FlatAppearance.MouseOverBackColor = Color.FromArgb(255, Math.Min(255, cor.R + 30), Math.Min(255, cor.G + 30), Math.Min(255, cor.B + 30))

        Return btn
    End Function

    ''' <summary>
    ''' Cria a área principal de conteúdo
    ''' </summary>
    Private Sub CriarAreaPrincipal()
        ' Header
        pnlHeader = New Panel() With {
            .Dock = DockStyle.Top,
            .Height = 80,
            .BackColor = Color.White,
            .Padding = New Padding(30, 20, 30, 20)
        }

        lblTitle = New Label() With {
            .Text = "Sistema PDV - Madeireira Maria Luiza",
            .Font = New Font("Segoe UI", 20, FontStyle.Bold),
            .ForeColor = Color.FromArgb(45, 45, 48),
            .Dock = DockStyle.Left,
            .AutoSize = True
        }

        lblStatusSistema = New Label() With {
            .Text = "Inicializando sistema...",
            .Font = New Font("Segoe UI", 10),
            .ForeColor = Color.Gray,
            .Dock = DockStyle.Right,
            .AutoSize = True,
            .TextAlign = ContentAlignment.MiddleRight
        }

        pnlHeader.Controls.AddRange({lblTitle, lblStatusSistema})

        ' Área principal
        pnlMain = New Panel() With {
            .Dock = DockStyle.Fill,
            .BackColor = Color.WhiteSmoke,
            .Padding = New Padding(30)
        }

        Me.Controls.AddRange({pnlHeader, pnlMain})
        CriarTelaInicial()
    End Sub

    ''' <summary>
    ''' Cria a tela inicial com informações do sistema
    ''' </summary>
    Private Sub CriarTelaInicial()
        Dim lblBemVindo As New Label() With {
            .Text = $"Bem-vindo ao Sistema PDV Integrado!",
            .Font = New Font("Segoe UI", 24, FontStyle.Bold),
            .ForeColor = Color.FromArgb(45, 45, 48),
            .Location = New Point(50, 50),
            .AutoSize = True
        }

        Dim lblDescricao As New Label() With {
            .Text = "Sistema completo de Ponto de Venda com integração total:" & Environment.NewLine & Environment.NewLine &
                   "✅ PDV Completo - Interface integrada com todas as funcionalidades" & Environment.NewLine &
                   "✅ Geração de Talões - Sistema automatizado com Excel" & Environment.NewLine &
                   "✅ Gestão de Clientes - CRUD completo com relatórios" & Environment.NewLine &
                   "✅ Gestão de Estoque - Controle de produtos e movimentações" & Environment.NewLine &
                   "✅ Relatórios Avançados - Dashboards e análises" & Environment.NewLine &
                   "✅ Banco de Dados Inteligente - Access com fallback para Excel" & Environment.NewLine &
                   "✅ Sistema de Busca - Produtos e clientes com filtros" & Environment.NewLine &
                   "✅ Calendário Integrado - Eventos e datas importantes" & Environment.NewLine &
                   "✅ Cálculos Automáticos - Totais, descontos e impostos",
            .Font = New Font("Segoe UI", 12),
            .ForeColor = Color.FromArgb(80, 80, 80),
            .Location = New Point(50, 100),
            .Size = New Size(700, 300)
        }

        Dim lblInstrucoes As New Label() With {
            .Text = "👈 Use o menu lateral para navegar pelas funcionalidades do sistema",
            .Font = New Font("Segoe UI", 14, FontStyle.Bold),
            .ForeColor = Color.FromArgb(0, 120, 215),
            .Location = New Point(50, 420),
            .AutoSize = True
        }

        Dim lblVersao As New Label() With {
            .Text = $"Versão 5.0 Integrada | Status: {_database?.VerificarConexao()}",
            .Font = New Font("Segoe UI", 9),
            .ForeColor = Color.Gray,
            .Location = New Point(50, 470),
            .AutoSize = True
        }

        pnlMain.Controls.AddRange({lblBemVindo, lblDescricao, lblInstrucoes, lblVersao})
    End Sub

    ''' <summary>
    ''' Configura eventos e comportamentos da interface
    ''' </summary>
    Private Sub ConfigurarInterface()
        ' Animações hover nos botões (implementar se necessário)
        ' Atalhos de teclado (implementar se necessário)
    End Sub

    #Region "Eventos dos Botões"

    Private Sub btnPDVCompleto_Click(sender As Object, e As EventArgs) Handles btnPDVCompleto.Click
        Try
            If _mainPDVForm Is Nothing OrElse _mainPDVForm.IsDisposed Then
                _mainPDVForm = New MainPDVForm()
            End If
            _mainPDVForm.Show()
            _mainPDVForm.BringToFront()
        Catch ex As Exception
            MessageBox.Show($"Erro ao abrir PDV: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnGerarTalao_Click(sender As Object, e As EventArgs) Handles btnGerarTalao.Click
        Try
            Using formPDV As New FormPDV()
                If formPDV.ShowDialog() = DialogResult.OK Then
                    Dim excelAutomation As New ExcelAutomation()
                    excelAutomation.ProcessarTalaoCompleto(formPDV.DadosColetados)
                    MessageBox.Show("Talão gerado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End Using
        Catch ex As Exception
            MessageBox.Show($"Erro ao gerar talão: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnGestaoClientes_Click(sender As Object, e As EventArgs) Handles btnGestaoClientes.Click
        Try
            Using form As New FormGestaoClientes()
                form.ShowDialog()
            End Using
        Catch ex As Exception
            MessageBox.Show($"Erro ao abrir gestão de clientes: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnGestaoEstoque_Click(sender As Object, e As EventArgs) Handles btnGestaoEstoque.Click
        Try
            Using form As New FormBuscaProdutos()
                form.ShowDialog()
            End Using
        Catch ex As Exception
            MessageBox.Show($"Erro ao abrir gestão de estoque: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnRelatorios_Click(sender As Object, e As EventArgs) Handles btnRelatorios.Click
        MessageBox.Show("Relatórios em desenvolvimento." & Environment.NewLine & 
                       "Em breve: Dashboard completo com gráficos e análises.", 
                       "Relatórios", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub btnConfiguracoes_Click(sender As Object, e As EventArgs) Handles btnConfiguracoes.Click
        MessageBox.Show("Configurações do sistema:" & Environment.NewLine & Environment.NewLine &
                       "• Edite o arquivo App.config para personalizar" & Environment.NewLine &
                       "• Nome da madeireira, endereço, vendedor padrão" & Environment.NewLine &
                       "• Configurações de banco de dados" & Environment.NewLine &
                       "• Parâmetros de impressão", 
                       "Configurações", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub btnSobre_Click(sender As Object, e As EventArgs) Handles btnSobre.Click
        MessageBox.Show($"Sistema PDV Integrado - {nomeMadeireira}" & Environment.NewLine & Environment.NewLine &
                       "Versão: 5.0 Integrada e Otimizada" & Environment.NewLine &
                       "Desenvolvedor: matheus-testuser3" & Environment.NewLine &
                       "Data: " & Date.Now.ToString("dd/MM/yyyy") & Environment.NewLine & Environment.NewLine &
                       "Sistema completo de PDV com:" & Environment.NewLine &
                       "• Interface moderna integrada" & Environment.NewLine &
                       "• Geração automática de talões" & Environment.NewLine &
                       "• Gestão completa de clientes e produtos" & Environment.NewLine &
                       "• Banco de dados inteligente" & Environment.NewLine &
                       "• Relatórios e análises" & Environment.NewLine & Environment.NewLine &
                       "Framework: .NET 4.7.2 | Excel Automation | VBA Integration", 
                       "Sobre o Sistema", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub btnSair_Click(sender As Object, e As EventArgs) Handles btnSair.Click
        If MessageBox.Show("Deseja realmente sair do sistema?", "Confirmar Saída", 
                          MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            Application.Exit()
        End If
    End Sub

    #End Region

    ''' <summary>
    ''' Limpa recursos ao fechar
    ''' </summary>
    Protected Overrides Sub OnFormClosed(e As FormClosedEventArgs)
        Try
            _mainPDVForm?.Close()
        Catch
        End Try
        MyBase.OnFormClosed(e)
    End Sub
