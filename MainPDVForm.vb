Imports System.Windows.Forms
Imports System.Drawing

''' <summary>
''' Formulário principal integrado do Sistema PDV - Madeireira Maria Luiza
''' Interface completa com todas as funcionalidades do sistema
''' </summary>
Public Class MainPDVForm
    Inherits Form
    
    #Region "Controles da Interface"
    
    ' Menu lateral
    Private WithEvents pnlMenuLateral As Panel
    Private WithEvents btnVendas As Button
    Private WithEvents btnClientes As Button
    Private WithEvents btnProdutos As Button
    Private WithEvents btnRelatorios As Button
    Private WithEvents btnConfiguracoes As Button
    Private WithEvents btnSair As Button
    
    ' Área principal
    Private WithEvents pnlPrincipal As Panel
    Private WithEvents pnlHeader As Panel
    Private WithEvents lblTitulo As Label
    Private WithEvents lblStatusConexao As Label
    
    ' Controles de venda
    Private WithEvents pnlVenda As Panel
    Private WithEvents dgvItens As DataGridView
    Private WithEvents pnlDadosCliente As Panel
    Private WithEvents txtNomeCliente As TextBox
    Private WithEvents txtEnderecoCliente As TextBox
    Private WithEvents txtTelefoneCliente As TextBox
    Private WithEvents btnBuscarCliente As Button
    
    ' Controles de produto
    Private WithEvents pnlProduto As Panel
    Private WithEvents txtBuscaProduto As TextBox
    Private WithEvents btnBuscarProduto As Button
    Private WithEvents cmbSecaoProduto As ComboBox
    Private WithEvents txtQuantidade As TextBox
    Private WithEvents txtPrecoUnitario As TextBox
    Private WithEvents btnAdicionarItem As Button
    
    ' Totais e pagamento
    Private WithEvents pnlTotais As Panel
    Private WithEvents lblSubtotal As Label
    Private WithEvents lblDesconto As Label
    Private WithEvents lblFrete As Label
    Private WithEvents lblTotal As Label
    Private WithEvents txtDescontoGeral As TextBox
    Private WithEvents txtFrete As TextBox
    Private WithEvents cmbFormaPagamento As ComboBox
    Private WithEvents cmbVendedor As ComboBox
    
    ' Botões de ação
    Private WithEvents pnlAcoes As Panel
    Private WithEvents btnGerarTalao As Button
    Private WithEvents btnLimparVenda As Button
    Private WithEvents btnSalvarVenda As Button
    Private WithEvents btnCarregarVenda As Button
    
    #End Region
    
    #Region "Propriedades e Variáveis"
    
    Private _calculadora As CalculadoraMadeireira
    Private _vendaAtual As Venda
    Private _database As DatabaseManager
    Private _searchManager As ProductSearchManager
    Private _config As ConfiguracaoSistema
    Private _calendarioManager As CalendarioManager
    
    #End Region
    
    #Region "Construtor e Inicialização"
    
    Public Sub New()
        InitializeComponent()
        InicializarSistema()
        ConfigurarInterface()
        CarregarDados()
    End Sub
    
    Private Sub InicializarSistema()
        Try
            _database = DatabaseManager.Instance
            _searchManager = New ProductSearchManager()
            _config = New ConfiguracaoSistema()
            _calendarioManager = New CalendarioManager()
            
            NovaVenda()
            
            ' Atualizar status da conexão
            lblStatusConexao.Text = _database.VerificarConexao()
            lblStatusConexao.ForeColor = If(_database.VerificarConexao().Contains("Access"), Color.Green, Color.Orange)
            
        Catch ex As Exception
            MessageBox.Show($"Erro ao inicializar sistema: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    Private Sub NovaVenda()
        _vendaAtual = New Venda()
        _calculadora = New CalculadoraMadeireira(_vendaAtual.Itens)
        AtualizarInterface()
    End Sub
    
    #End Region
    
    #Region "Interface Design"
    
    Private Sub InitializeComponent()
        Me.Text = $"Sistema PDV - {_config?.NomeMadeireira}"
        Me.Size = New Size(1200, 800)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.WindowState = FormWindowState.Maximized
        Me.BackColor = Color.WhiteSmoke
        
        CriarMenuLateral()
        CriarAreaPrincipal()
        CriarPainelVenda()
        CriarPainelTotais()
        CriarPainelAcoes()
    End Sub
    
    Private Sub CriarMenuLateral()
        pnlMenuLateral = New Panel() With {
            .Dock = DockStyle.Left,
            .Width = 200,
            .BackColor = Color.DarkBlue,
            .Padding = New Padding(10)
        }
        
        ' Logo/Título
        Dim lblLogo As New Label() With {
            .Text = "PDV INTEGRADO",
            .Dock = DockStyle.Top,
            .Height = 50,
            .ForeColor = Color.White,
            .Font = New Font("Segoe UI", 12, FontStyle.Bold),
            .TextAlign = ContentAlignment.MiddleCenter
        }
        
        ' Botões do menu
        btnVendas = CriarBotaoMenu("🛒 VENDAS", 60)
        btnClientes = CriarBotaoMenu("👥 CLIENTES", 120)
        btnProdutos = CriarBotaoMenu("📦 PRODUTOS", 180)
        btnRelatorios = CriarBotaoMenu("📊 RELATÓRIOS", 240)
        btnConfiguracoes = CriarBotaoMenu("⚙️ CONFIGURAÇÕES", 300)
        btnSair = CriarBotaoMenu("❌ SAIR", 500)
        
        btnSair.BackColor = Color.DarkRed
        btnVendas.BackColor = Color.Green ' Ativo por padrão
        
        pnlMenuLateral.Controls.AddRange({lblLogo, btnVendas, btnClientes, btnProdutos, btnRelatorios, btnConfiguracoes, btnSair})
        Me.Controls.Add(pnlMenuLateral)
    End Sub
    
    Private Function CriarBotaoMenu(texto As String, top As Integer) As Button
        Return New Button() With {
            .Text = texto,
            .Size = New Size(180, 40),
            .Location = New Point(10, top),
            .BackColor = Color.SteelBlue,
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .TextAlign = ContentAlignment.MiddleLeft,
            .Padding = New Padding(10, 0, 0, 0)
        }
    End Function
    
    Private Sub CriarAreaPrincipal()
        ' Header
        pnlHeader = New Panel() With {
            .Dock = DockStyle.Top,
            .Height = 60,
            .BackColor = Color.White,
            .Padding = New Padding(20, 10, 20, 10)
        }
        
        lblTitulo = New Label() With {
            .Text = "NOVA VENDA",
            .Font = New Font("Segoe UI", 16, FontStyle.Bold),
            .ForeColor = Color.DarkBlue,
            .Dock = DockStyle.Left,
            .AutoSize = True
        }
        
        lblStatusConexao = New Label() With {
            .Text = "Verificando conexão...",
            .Font = New Font("Segoe UI", 9),
            .Dock = DockStyle.Right,
            .AutoSize = True,
            .ForeColor = Color.Gray
        }
        
        pnlHeader.Controls.AddRange({lblTitulo, lblStatusConexao})
        
        ' Área principal
        pnlPrincipal = New Panel() With {
            .Dock = DockStyle.Fill,
            .BackColor = Color.WhiteSmoke,
            .Padding = New Padding(20)
        }
        
        Me.Controls.AddRange({pnlHeader, pnlPrincipal})
    End Sub
    
    Private Sub CriarPainelVenda()
        ' Painel de dados do cliente
        pnlDadosCliente = New Panel() With {
            .Dock = DockStyle.Top,
            .Height = 100,
            .BackColor = Color.White,
            .Padding = New Padding(10)
        }
        
        Dim lblCliente As New Label() With {
            .Text = "DADOS DO CLIENTE:",
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .Location = New Point(10, 10),
            .AutoSize = True
        }
        
        txtNomeCliente = New TextBox() With {
            .Location = New Point(10, 35),
            .Size = New Size(250, 25),
            .Font = New Font("Segoe UI", 10)
        }
        
        txtEnderecoCliente = New TextBox() With {
            .Location = New Point(270, 35),
            .Size = New Size(300, 25),
            .Font = New Font("Segoe UI", 10)
        }
        
        txtTelefoneCliente = New TextBox() With {
            .Location = New Point(580, 35),
            .Size = New Size(150, 25),
            .Font = New Font("Segoe UI", 10)
        }
        
        btnBuscarCliente = New Button() With {
            .Text = "🔍",
            .Location = New Point(740, 35),
            .Size = New Size(30, 25),
            .BackColor = Color.LightBlue,
            .FlatStyle = FlatStyle.Flat
        }
        
        pnlDadosCliente.Controls.AddRange({lblCliente, txtNomeCliente, txtEnderecoCliente, txtTelefoneCliente, btnBuscarCliente})
        
        ' Painel de produtos
        pnlProduto = New Panel() With {
            .Dock = DockStyle.Top,
            .Height = 80,
            .BackColor = Color.LightGray,
            .Padding = New Padding(10)
        }
        
        Dim lblProduto As New Label() With {
            .Text = "ADICIONAR PRODUTO:",
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .Location = New Point(10, 10),
            .AutoSize = True
        }
        
        txtBuscaProduto = New TextBox() With {
            .Location = New Point(10, 35),
            .Size = New Size(200, 25),
            .Font = New Font("Segoe UI", 10)
        }
        
        btnBuscarProduto = New Button() With {
            .Text = "🔍 Buscar",
            .Location = New Point(220, 35),
            .Size = New Size(80, 25),
            .BackColor = Color.DodgerBlue,
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
        
        cmbSecaoProduto = New ComboBox() With {
            .Location = New Point(310, 35),
            .Size = New Size(120, 25),
            .DropDownStyle = ComboBoxStyle.DropDownList
        }
        
        txtQuantidade = New TextBox() With {
            .Location = New Point(440, 35),
            .Size = New Size(60, 25),
            .Font = New Font("Segoe UI", 10),
            .Text = "1"
        }
        
        txtPrecoUnitario = New TextBox() With {
            .Location = New Point(510, 35),
            .Size = New Size(80, 25),
            .Font = New Font("Segoe UI", 10)
        }
        
        btnAdicionarItem = New Button() With {
            .Text = "➕ Adicionar",
            .Location = New Point(600, 35),
            .Size = New Size(100, 25),
            .BackColor = Color.Green,
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
        
        pnlProduto.Controls.AddRange({lblProduto, txtBuscaProduto, btnBuscarProduto, cmbSecaoProduto, txtQuantidade, txtPrecoUnitario, btnAdicionarItem})
        
        ' Grid de itens
        dgvItens = New DataGridView() With {
            .Dock = DockStyle.Fill,
            .AllowUserToAddRows = False,
            .AllowUserToDeleteRows = True,
            .BackgroundColor = Color.White,
            .BorderStyle = BorderStyle.None,
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
        }
        
        ConfigurarGridItens()
        
        pnlPrincipal.Controls.AddRange({pnlDadosCliente, pnlProduto, dgvItens})
    End Sub
    
    Private Sub ConfigurarGridItens()
        dgvItens.Columns.Clear()
        dgvItens.Columns.Add("Codigo", "Código")
        dgvItens.Columns.Add("Descricao", "Descrição")
        dgvItens.Columns.Add("Quantidade", "Qtd")
        dgvItens.Columns.Add("Unidade", "Un")
        dgvItens.Columns.Add("PrecoUnitario", "Preço Unit.")
        dgvItens.Columns.Add("Desconto", "Desconto")
        dgvItens.Columns.Add("Subtotal", "Subtotal")
        
        ' Configurar larguras
        dgvItens.Columns("Codigo").Width = 80
        dgvItens.Columns("Descricao").Width = 300
        dgvItens.Columns("Quantidade").Width = 80
        dgvItens.Columns("Unidade").Width = 60
        dgvItens.Columns("PrecoUnitario").Width = 100
        dgvItens.Columns("Desconto").Width = 100
        dgvItens.Columns("Subtotal").Width = 120
        
        ' Formatação
        dgvItens.Columns("PrecoUnitario").DefaultCellStyle.Format = "C2"
        dgvItens.Columns("Desconto").DefaultCellStyle.Format = "C2"
        dgvItens.Columns("Subtotal").DefaultCellStyle.Format = "C2"
        dgvItens.Columns("Quantidade").DefaultCellStyle.Format = "N3"
        
        ' Configurar cálculo automático
        CalculosUtilities.ConfigurarCalculoAutomatico(dgvItens, _calculadora)
    End Sub
    
    Private Sub CriarPainelTotais()
        pnlTotais = New Panel() With {
            .Dock = DockStyle.Bottom,
            .Height = 120,
            .BackColor = Color.White,
            .Padding = New Padding(20)
        }
        
        ' Labels de totais
        Dim lblSubtotalTxt As New Label() With {
            .Text = "Subtotal:",
            .Location = New Point(20, 20),
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .AutoSize = True
        }
        
        lblSubtotal = New Label() With {
            .Location = New Point(120, 20),
            .Font = New Font("Segoe UI", 10),
            .AutoSize = True,
            .Text = "R$ 0,00"
        }
        
        Dim lblDescontoTxt As New Label() With {
            .Text = "Desconto:",
            .Location = New Point(20, 45),
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .AutoSize = True
        }
        
        lblDesconto = New Label() With {
            .Location = New Point(120, 45),
            .Font = New Font("Segoe UI", 10),
            .AutoSize = True,
            .Text = "R$ 0,00"
        }
        
        Dim lblFreteTxt As New Label() With {
            .Text = "Frete:",
            .Location = New Point(20, 70),
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .AutoSize = True
        }
        
        lblFrete = New Label() With {
            .Location = New Point(120, 70),
            .Font = New Font("Segoe UI", 10),
            .AutoSize = True,
            .Text = "R$ 0,00"
        }
        
        Dim lblTotalTxt As New Label() With {
            .Text = "TOTAL:",
            .Location = New Point(250, 40),
            .Font = New Font("Segoe UI", 14, FontStyle.Bold),
            .ForeColor = Color.DarkBlue,
            .AutoSize = True
        }
        
        lblTotal = New Label() With {
            .Location = New Point(350, 40),
            .Font = New Font("Segoe UI", 16, FontStyle.Bold),
            .ForeColor = Color.DarkBlue,
            .AutoSize = True,
            .Text = "R$ 0,00"
        }
        
        ' Controles de entrada
        txtDescontoGeral = New TextBox() With {
            .Location = New Point(600, 20),
            .Size = New Size(100, 25),
            .Font = New Font("Segoe UI", 10)
        }
        
        txtFrete = New TextBox() With {
            .Location = New Point(600, 50),
            .Size = New Size(100, 25),
            .Font = New Font("Segoe UI", 10)
        }
        
        cmbFormaPagamento = New ComboBox() With {
            .Location = New Point(720, 20),
            .Size = New Size(150, 25),
            .DropDownStyle = ComboBoxStyle.DropDownList
        }
        
        cmbVendedor = New ComboBox() With {
            .Location = New Point(720, 50),
            .Size = New Size(150, 25),
            .DropDownStyle = ComboBoxStyle.DropDownList
        }
        
        pnlTotais.Controls.AddRange({lblSubtotalTxt, lblSubtotal, lblDescontoTxt, lblDesconto, 
                                   lblFreteTxt, lblFrete, lblTotalTxt, lblTotal,
                                   txtDescontoGeral, txtFrete, cmbFormaPagamento, cmbVendedor})
        
        pnlPrincipal.Controls.Add(pnlTotais)
    End Sub
    
    Private Sub CriarPainelAcoes()
        pnlAcoes = New Panel() With {
            .Dock = DockStyle.Bottom,
            .Height = 60,
            .BackColor = Color.LightGray,
            .Padding = New Padding(20, 10, 20, 10)
        }
        
        btnGerarTalao = New Button() With {
            .Text = "🧾 GERAR TALÃO",
            .Size = New Size(150, 40),
            .Location = New Point(20, 10),
            .BackColor = Color.Green,
            .ForeColor = Color.White,
            .Font = New Font("Segoe UI", 11, FontStyle.Bold),
            .FlatStyle = FlatStyle.Flat
        }
        
        btnSalvarVenda = New Button() With {
            .Text = "💾 SALVAR",
            .Size = New Size(100, 40),
            .Location = New Point(180, 10),
            .BackColor = Color.Blue,
            .ForeColor = Color.White,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .FlatStyle = FlatStyle.Flat
        }
        
        btnLimparVenda = New Button() With {
            .Text = "🗑️ LIMPAR",
            .Size = New Size(100, 40),
            .Location = New Point(290, 10),
            .BackColor = Color.Orange,
            .ForeColor = Color.White,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .FlatStyle = FlatStyle.Flat
        }
        
        btnCarregarVenda = New Button() With {
            .Text = "📂 CARREGAR",
            .Size = New Size(120, 40),
            .Location = New Point(400, 10),
            .BackColor = Color.Purple,
            .ForeColor = Color.White,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .FlatStyle = FlatStyle.Flat
        }
        
        pnlAcoes.Controls.AddRange({btnGerarTalao, btnSalvarVenda, btnLimparVenda, btnCarregarVenda})
        Me.Controls.Add(pnlAcoes)
    End Sub
    
    #End Region
    
    #Region "Configuração da Interface"
    
    Private Sub ConfigurarInterface()
        ' Configurar validação numérica
        CalculosUtilities.ValidarEntradaNumerica(txtQuantidade, True)
        CalculosUtilities.ValidarEntradaNumerica(txtPrecoUnitario, True)
        CalculosUtilities.ValidarEntradaNumerica(txtDescontoGeral, True)
        CalculosUtilities.ValidarEntradaNumerica(txtFrete, True)
        
        ' Configurar eventos de mudança para cálculo automático
        AddHandler txtDescontoGeral.TextChanged, AddressOf RecalcularTotais
        AddHandler txtFrete.TextChanged, AddressOf RecalcularTotais
    End Sub
    
    Private Sub CarregarDados()
        Try
            ' Carregar seções de produtos
            cmbSecaoProduto.Items.Clear()
            cmbSecaoProduto.Items.Add("Todas")
            For Each secao In _searchManager.ObterSecoes()
                cmbSecaoProduto.Items.Add(secao)
            Next
            cmbSecaoProduto.SelectedIndex = 0
            
            ' Carregar formas de pagamento
            cmbFormaPagamento.Items.AddRange({"À Vista", "Cartão Débito", "Cartão Crédito", "PIX", "Boleto", "Fiado"})
            cmbFormaPagamento.SelectedIndex = 0
            
            ' Carregar vendedores
            cmbVendedor.Items.Add(_config.VendedorPadrao)
            cmbVendedor.SelectedIndex = 0
            
        Catch ex As Exception
            MessageBox.Show($"Erro ao carregar dados: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    #End Region
    
    #Region "Eventos dos Controles"
    
    Private Sub btnBuscarProduto_Click(sender As Object, e As EventArgs) Handles btnBuscarProduto.Click
        Try
            Using formBusca As New FormBuscaProdutos()
                If formBusca.ShowDialog() = DialogResult.OK Then
                    Dim produto = formBusca.ProdutoSelecionado
                    If produto IsNot Nothing Then
                        txtBuscaProduto.Text = produto.Descricao
                        txtPrecoUnitario.Text = produto.PrecoVenda.ToString("F2")
                        txtQuantidade.Focus()
                    End If
                End If
            End Using
        Catch ex As Exception
            MessageBox.Show($"Erro ao buscar produto: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    Private Sub btnAdicionarItem_Click(sender As Object, e As EventArgs) Handles btnAdicionarItem.Click
        Try
            If String.IsNullOrEmpty(txtBuscaProduto.Text) Then
                MessageBox.Show("Selecione um produto primeiro.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            
            Dim quantidade As Double
            Dim preco As Decimal
            
            If Not Double.TryParse(txtQuantidade.Text, quantidade) OrElse quantidade <= 0 Then
                MessageBox.Show("Quantidade inválida.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            
            If Not Decimal.TryParse(txtPrecoUnitario.Text, preco) OrElse preco < 0 Then
                MessageBox.Show("Preço inválido.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            
            ' Criar produto temporário para o item
            Dim produto As New Produto() With {
                .Codigo = "TEMP",
                .Descricao = txtBuscaProduto.Text,
                .Unidade = "UN",
                .PrecoVenda = preco
            }
            
            ' Criar item da venda
            Dim item As New ItemVenda() With {
                .Produto = produto,
                .Quantidade = quantidade,
                .PrecoUnitario = preco,
                .Desconto = 0
            }
            
            _vendaAtual.Itens.Add(item)
            _calculadora.Itens = _vendaAtual.Itens
            
            AtualizarGridItens()
            AtualizarTotais()
            LimparCamposProduto()
            
        Catch ex As Exception
            MessageBox.Show($"Erro ao adicionar item: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    Private Sub btnGerarTalao_Click(sender As Object, e As EventArgs) Handles btnGerarTalao.Click
        Try
            If _vendaAtual.Itens.Count = 0 Then
                MessageBox.Show("Adicione pelo menos um item à venda.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            
            ' Atualizar dados da venda com dados da interface
            _vendaAtual.Cliente.Nome = txtNomeCliente.Text
            _vendaAtual.Cliente.Endereco = txtEnderecoCliente.Text
            _vendaAtual.Cliente.Telefone = txtTelefoneCliente.Text
            _vendaAtual.FormaPagamento = cmbFormaPagamento.Text
            _vendaAtual.VendedorNome = cmbVendedor.Text
            
            ' Atualizar calculadora
            Decimal.TryParse(txtDescontoGeral.Text, _calculadora.DescontoGeral)
            Decimal.TryParse(txtFrete.Text, _calculadora.Frete)
            
            ' Abrir formulário de confirmação
            Using formConfirmacao As New FormConfirmacaoPedido(_vendaAtual, _calculadora)
                Dim resultado = formConfirmacao.ShowDialog()
                
                If resultado = DialogResult.OK Then
                    ' Pedido confirmado
                    MessageBox.Show("Pedido confirmado e talão gerado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    NovaVenda() ' Limpar para nova venda
                ElseIf resultado = DialogResult.Retry Then
                    ' Usuário quer editar - manter dados atuais
                    Return
                End If
            End Using
            
        Catch ex As Exception
            MessageBox.Show($"Erro ao processar pedido: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    Private Sub btnLimparVenda_Click(sender As Object, e As EventArgs) Handles btnLimparVenda.Click
        If MessageBox.Show("Deseja limpar todos os dados da venda atual?", "Confirmar", 
                          MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            NovaVenda()
        End If
    End Sub
    
    Private Sub btnSair_Click(sender As Object, e As EventArgs) Handles btnSair.Click
        If MessageBox.Show("Deseja sair do sistema?", "Confirmar", 
                          MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            Application.Exit()
        End If
    End Sub
    
    #End Region
    
    #Region "Métodos Auxiliares"
    
    Private Sub AtualizarInterface()
        AtualizarGridItens()
        AtualizarTotais()
        LimparCamposProduto()
        LimparCamposCliente()
    End Sub
    
    Private Sub AtualizarGridItens()
        dgvItens.Rows.Clear()
        
        For Each item In _vendaAtual.Itens
            dgvItens.Rows.Add(
                item.Produto.Codigo,
                item.Produto.Descricao,
                item.Quantidade,
                item.Produto.Unidade,
                item.PrecoUnitario,
                item.Desconto,
                _calculadora.CalcularSubtotalItem(item)
            )
            
            ' Armazenar referência do item
            dgvItens.Rows(dgvItens.Rows.Count - 1).Tag = item
        Next
    End Sub
    
    Private Sub AtualizarTotais()
        CalculosUtilities.AtualizarTotaisInterface(_calculadora, lblSubtotal, lblDesconto, lblTotal, lblFrete)
    End Sub
    
    Private Sub RecalcularTotais(sender As Object, e As EventArgs)
        Try
            ' Atualizar valores na calculadora
            Decimal.TryParse(txtDescontoGeral.Text, _calculadora.DescontoGeral)
            Decimal.TryParse(txtFrete.Text, _calculadora.Frete)
            
            AtualizarTotais()
        Catch ex As Exception
            Console.WriteLine($"Erro no recálculo: {ex.Message}")
        End Try
    End Sub
    
    Private Sub LimparCamposProduto()
        txtBuscaProduto.Clear()
        txtQuantidade.Text = "1"
        txtPrecoUnitario.Clear()
        txtBuscaProduto.Focus()
    End Sub
    
    Private Sub LimparCamposCliente()
        txtNomeCliente.Clear()
        txtEnderecoCliente.Clear()
        txtTelefoneCliente.Clear()
    End Sub
    
    #End Region
End Class