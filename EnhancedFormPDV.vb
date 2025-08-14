Imports System.Windows.Forms
Imports System.Drawing
Imports System.Configuration

''' <summary>
''' Formulário de entrada de dados para geração de talão
''' Interface moderna com validação em tempo real e UX otimizada
''' </summary>
Public Class EnhancedFormPDV
    Inherits Form

    ' Controles da interface principal
    Private WithEvents pnlHeader As Panel
    Private WithEvents pnlCliente As Panel
    Private WithEvents pnlProdutos As Panel
    Private WithEvents pnlResumo As Panel
    Private WithEvents pnlBotoes As Panel
    
    ' Controles do cliente com validação
    Private WithEvents txtNomeCliente As TextBox
    Private WithEvents txtEnderecoCliente As TextBox
    Private WithEvents txtCEP As MaskedTextBox
    Private WithEvents txtCidade As TextBox
    Private WithEvents txtTelefone As MaskedTextBox
    Private WithEvents cmbFormaPagamento As ComboBox
    Private WithEvents txtVendedor As TextBox
    
    ' Labels de status de validação
    Private lblStatusNome As Label
    Private lblStatusCEP As Label
    Private lblStatusTelefone As Label
    
    ' Controles de produtos
    Private WithEvents dgvProdutos As DataGridView
    Private WithEvents txtDescricaoProduto As TextBox
    Private WithEvents txtQuantidade As NumericUpDown
    Private WithEvents cmbUnidade As ComboBox
    Private WithEvents txtPrecoUnitario As NumericUpDown
    Private WithEvents btnAdicionarProduto As Button
    Private WithEvents btnRemoverProduto As Button
    Private WithEvents btnEditarProduto As Button
    
    ' Resumo e totais
    Private lblSubtotal As Label
    Private lblDesconto As Label
    Private lblTotal As Label
    Private txtDesconto As NumericUpDown
    
    ' Botões principais
    Private WithEvents btnConfirmar As Button
    Private WithEvents btnCancelar As Button
    Private WithEvents btnDadosTeste As Button
    Private WithEvents btnLimpar As Button
    
    ' Sistema
    Private ReadOnly _logger As LoggingSystem = LoggingSystem.Instance
    Private ReadOnly _config As EnhancedConfigurationManager = EnhancedConfigurationManager.Instance
    Private _validationErrors As New Dictionary(Of String, String)()
    Private WithEvents _validationTimer As New Timer()
    
    ' Propriedades
    Public Property DadosColetados As DadosTalao
    Public Property ValidacaoTempoReal As Boolean = True
    
    ''' <summary>
    ''' Construtor do formulário
    ''' </summary>
    Public Sub New()
        InitializeComponent()
        ConfigurarInterface()
        ConfigurarValidacao()
        DadosColetados = New DadosTalao()
        
        _logger.LogInfo("EnhancedFormPDV", "Formulário inicializado")
    End Sub
    
    #Region "Inicialização"
    
    ''' <summary>
    ''' Inicializa os componentes da interface
    ''' </summary>
    Private Sub InitializeComponent()
        Me.Text = "🏪 Sistema PDV - Entrada de Dados"
        Me.Size = New Size(1000, 800)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.BackColor = Color.FromArgb(236, 240, 241)
        Me.Font = New Font("Segoe UI", 9.0F, FontStyle.Regular)
        
        CriarHeader()
        CriarPainelCliente()
        CriarPainelProdutos()
        CriarPainelResumo()
        CriarPainelBotoes()
    End Sub
    
    ''' <summary>
    ''' Cria cabeçalho do formulário
    ''' </summary>
    Private Sub CriarHeader()
        pnlHeader = New Panel()
        pnlHeader.Size = New Size(980, 60)
        pnlHeader.Location = New Point(10, 10)
        pnlHeader.BackColor = Color.FromArgb(52, 73, 94)
        pnlHeader.ForeColor = Color.White
        Me.Controls.Add(pnlHeader)
        
        Dim lblTitulo As New Label()
        lblTitulo.Text = "📋 NOVA VENDA - TALÃO"
        lblTitulo.Font = New Font("Segoe UI", 16.0F, FontStyle.Bold)
        lblTitulo.ForeColor = Color.White
        lblTitulo.Location = New Point(20, 15)
        lblTitulo.Size = New Size(300, 30)
        pnlHeader.Controls.Add(lblTitulo)
        
        Dim lblSubtitulo As New Label()
        lblSubtitulo.Text = $"🏪 {_config.NomeMadeireira}"
        lblSubtitulo.Font = New Font("Segoe UI", 10.0F, FontStyle.Regular)
        lblSubtitulo.ForeColor = Color.FromArgb(189, 195, 199)
        lblSubtitulo.Location = New Point(400, 20)
        lblSubtitulo.Size = New Size(400, 20)
        pnlHeader.Controls.Add(lblSubtitulo)
    End Sub
    
    ''' <summary>
    ''' Cria painel de dados do cliente
    ''' </summary>
    Private Sub CriarPainelCliente()
        pnlCliente = New Panel()
        pnlCliente.Size = New Size(980, 220)
        pnlCliente.Location = New Point(10, 80)
        pnlCliente.BackColor = Color.White
        pnlCliente.BorderStyle = BorderStyle.FixedSingle
        Me.Controls.Add(pnlCliente)
        
        ' Título da seção
        Dim lblTituloCliente As New Label()
        lblTituloCliente.Text = "👤 DADOS DO CLIENTE"
        lblTituloCliente.Font = New Font("Segoe UI", 12.0F, FontStyle.Bold)
        lblTituloCliente.ForeColor = Color.FromArgb(52, 73, 94)
        lblTituloCliente.Size = New Size(200, 25)
        lblTituloCliente.Location = New Point(15, 10)
        pnlCliente.Controls.Add(lblTituloCliente)
        
        ' Nome do Cliente com validação
        CriarCampoTexto(pnlCliente, "Nome do Cliente *:", txtNomeCliente, lblStatusNome, 15, 45, 300, True)
        
        ' Endereço
        CriarCampoTexto(pnlCliente, "Endereço:", txtEnderecoCliente, Nothing, 15, 85, 400, False)
        
        ' CEP com máscara e validação
        CriarCampoCEP()
        
        ' Cidade
        CriarCampoTexto(pnlCliente, "Cidade:", txtCidade, Nothing, 15, 165, 200, False)
        
        ' Telefone com máscara
        CriarCampoTelefone()
        
        ' Forma de Pagamento
        CriarCampoFormaPagamento()
        
        ' Vendedor
        CriarCampoVendedor()
    End Sub
    
    ''' <summary>
    ''' Cria campo de texto genérico com validação
    ''' </summary>
    Private Sub CriarCampoTexto(painel As Panel, labelText As String, ByRef textBox As TextBox, ByRef statusLabel As Label, x As Integer, y As Integer, width As Integer, isRequired As Boolean)
        ' Label
        Dim lbl As New Label()
        lbl.Text = labelText
        lbl.Location = New Point(x, y)
        lbl.Size = New Size(140, 20)
        lbl.Font = New Font("Segoe UI", 9.0F, FontStyle.Regular)
        If isRequired Then lbl.ForeColor = Color.FromArgb(231, 76, 60)
        painel.Controls.Add(lbl)
        
        ' TextBox
        textBox = New TextBox()
        textBox.Location = New Point(x + 150, y - 2)
        textBox.Size = New Size(width, 23)
        textBox.Font = New Font("Segoe UI", 9.0F)
        painel.Controls.Add(textBox)
        
        ' Status label se for campo obrigatório
        If isRequired Then
            statusLabel = New Label()
            statusLabel.Location = New Point(x + width + 160, y)
            statusLabel.Size = New Size(20, 20)
            statusLabel.Text = "⚠️"
            statusLabel.ForeColor = Color.FromArgb(231, 76, 60)
            statusLabel.Font = New Font("Segoe UI", 8.0F)
            painel.Controls.Add(statusLabel)
            
            ' Tooltip
            Dim tooltip As New ToolTip()
            tooltip.SetToolTip(statusLabel, "Campo obrigatório")
        End If
    End Sub
    
    ''' <summary>
    ''' Cria campo CEP com máscara
    ''' </summary>
    Private Sub CriarCampoCEP()
        Dim lblCEP As New Label()
        lblCEP.Text = "CEP:"
        lblCEP.Location = New Point(450, 85)
        lblCEP.Size = New Size(40, 20)
        pnlCliente.Controls.Add(lblCEP)
        
        txtCEP = New MaskedTextBox()
        txtCEP.Mask = "00000-000"
        txtCEP.Location = New Point(500, 83)
        txtCEP.Size = New Size(100, 23)
        txtCEP.Font = New Font("Segoe UI", 9.0F)
        pnlCliente.Controls.Add(txtCEP)
        
        lblStatusCEP = New Label()
        lblStatusCEP.Location = New Point(610, 85)
        lblStatusCEP.Size = New Size(20, 20)
        lblStatusCEP.Text = "ℹ️"
        lblStatusCEP.Visible = False
        pnlCliente.Controls.Add(lblStatusCEP)
        
        ' Botão consultar CEP
        Dim btnConsultarCEP As New Button()
        btnConsultarCEP.Text = "🔍"
        btnConsultarCEP.Size = New Size(30, 25)
        btnConsultarCEP.Location = New Point(635, 82)
        btnConsultarCEP.UseVisualStyleBackColor = True
        AddHandler btnConsultarCEP.Click, AddressOf ConsultarCEP
        pnlCliente.Controls.Add(btnConsultarCEP)
        
        Dim tooltip As New ToolTip()
        tooltip.SetToolTip(btnConsultarCEP, "Consultar CEP automaticamente")
    End Sub
    
    ''' <summary>
    ''' Cria campo telefone com máscara
    ''' </summary>
    Private Sub CriarCampoTelefone()
        Dim lblTelefone As New Label()
        lblTelefone.Text = "Telefone:"
        lblTelefone.Location = New Point(250, 165)
        lblTelefone.Size = New Size(80, 20)
        pnlCliente.Controls.Add(lblTelefone)
        
        txtTelefone = New MaskedTextBox()
        txtTelefone.Mask = "(00) 00000-0000"
        txtTelefone.Location = New Point(340, 163)
        txtTelefone.Size = New Size(120, 23)
        txtTelefone.Font = New Font("Segoe UI", 9.0F)
        pnlCliente.Controls.Add(txtTelefone)
        
        lblStatusTelefone = New Label()
        lblStatusTelefone.Location = New Point(470, 165)
        lblStatusTelefone.Size = New Size(20, 20)
        lblStatusTelefone.Visible = False
        pnlCliente.Controls.Add(lblStatusTelefone)
    End Sub
    
    ''' <summary>
    ''' Cria campo forma de pagamento
    ''' </summary>
    Private Sub CriarCampoFormaPagamento()
        Dim lblFormaPagamento As New Label()
        lblFormaPagamento.Text = "Forma Pagamento:"
        lblFormaPagamento.Location = New Point(500, 165)
        lblFormaPagamento.Size = New Size(110, 20)
        pnlCliente.Controls.Add(lblFormaPagamento)
        
        cmbFormaPagamento = New ComboBox()
        cmbFormaPagamento.Location = New Point(620, 163)
        cmbFormaPagamento.Size = New Size(150, 23)
        cmbFormaPagamento.DropDownStyle = ComboBoxStyle.DropDownList
        cmbFormaPagamento.Items.AddRange({"À Vista", "Cartão Débito", "Cartão Crédito", "PIX", "Dinheiro", "Cheque", "Boleto", "Parcelado"})
        cmbFormaPagamento.SelectedIndex = 0
        pnlCliente.Controls.Add(cmbFormaPagamento)
    End Sub
    
    ''' <summary>
    ''' Cria campo vendedor
    ''' </summary>
    Private Sub CriarCampoVendedor()
        Dim lblVendedor As New Label()
        lblVendedor.Text = "Vendedor:"
        lblVendedor.Location = New Point(15, 195)
        lblVendedor.Size = New Size(80, 20)
        pnlCliente.Controls.Add(lblVendedor)
        
        txtVendedor = New TextBox()
        txtVendedor.Location = New Point(165, 193)
        txtVendedor.Size = New Size(200, 23)
        txtVendedor.Font = New Font("Segoe UI", 9.0F)
        txtVendedor.Text = _config.VendedorPadrao
        pnlCliente.Controls.Add(txtVendedor)
    End Sub
    
    ''' <summary>
    ''' Cria painel de produtos
    ''' </summary>
    Private Sub CriarPainelProdutos()
        pnlProdutos = New Panel()
        pnlProdutos.Size = New Size(980, 300)
        pnlProdutos.Location = New Point(10, 310)
        pnlProdutos.BackColor = Color.White
        pnlProdutos.BorderStyle = BorderStyle.FixedSingle
        Me.Controls.Add(pnlProdutos)
        
        ' Título da seção
        Dim lblTituloProdutos As New Label()
        lblTituloProdutos.Text = "🛒 PRODUTOS"
        lblTituloProdutos.Font = New Font("Segoe UI", 12.0F, FontStyle.Bold)
        lblTituloProdutos.ForeColor = Color.FromArgb(52, 73, 94)
        lblTituloProdutos.Size = New Size(150, 25)
        lblTituloProdutos.Location = New Point(15, 10)
        pnlProdutos.Controls.Add(lblTituloProdutos)
        
        ' Campos de entrada de produto
        CriarCamposProduto()
        
        ' DataGridView para produtos
        CriarGridProdutos()
    End Sub
    
    ''' <summary>
    ''' Cria campos de entrada de produto
    ''' </summary>
    Private Sub CriarCamposProduto()
        ' Descrição do produto
        Dim lblDescricao As New Label()
        lblDescricao.Text = "Descrição:"
        lblDescricao.Location = New Point(15, 45)
        lblDescricao.Size = New Size(80, 20)
        pnlProdutos.Controls.Add(lblDescricao)
        
        txtDescricaoProduto = New TextBox()
        txtDescricaoProduto.Location = New Point(100, 43)
        txtDescricaoProduto.Size = New Size(300, 23)
        txtDescricaoProduto.Font = New Font("Segoe UI", 9.0F)
        pnlProdutos.Controls.Add(txtDescricaoProduto)
        
        ' Quantidade
        Dim lblQuantidade As New Label()
        lblQuantidade.Text = "Qtd:"
        lblQuantidade.Location = New Point(420, 45)
        lblQuantidade.Size = New Size(30, 20)
        pnlProdutos.Controls.Add(lblQuantidade)
        
        txtQuantidade = New NumericUpDown()
        txtQuantidade.Location = New Point(455, 43)
        txtQuantidade.Size = New Size(80, 23)
        txtQuantidade.Minimum = 0.01
        txtQuantidade.Maximum = 9999
        txtQuantidade.DecimalPlaces = 2
        txtQuantidade.Value = 1
        pnlProdutos.Controls.Add(txtQuantidade)
        
        ' Unidade
        Dim lblUnidade As New Label()
        lblUnidade.Text = "Un:"
        lblUnidade.Location = New Point(550, 45)
        lblUnidade.Size = New Size(25, 20)
        pnlProdutos.Controls.Add(lblUnidade)
        
        cmbUnidade = New ComboBox()
        cmbUnidade.Location = New Point(580, 43)
        cmbUnidade.Size = New Size(60, 23)
        cmbUnidade.DropDownStyle = ComboBoxStyle.DropDownList
        cmbUnidade.Items.AddRange({"UN", "M", "M²", "M³", "KG", "L", "PC", "CX", "SC"})
        cmbUnidade.SelectedIndex = 0
        pnlProdutos.Controls.Add(cmbUnidade)
        
        ' Preço unitário
        Dim lblPreco As New Label()
        lblPreco.Text = "Preço:"
        lblPreco.Location = New Point(660, 45)
        lblPreco.Size = New Size(40, 20)
        pnlProdutos.Controls.Add(lblPreco)
        
        txtPrecoUnitario = New NumericUpDown()
        txtPrecoUnitario.Location = New Point(705, 43)
        txtPrecoUnitario.Size = New Size(80, 23)
        txtPrecoUnitario.Minimum = 0.01
        txtPrecoUnitario.Maximum = 999999
        txtPrecoUnitario.DecimalPlaces = 2
        txtPrecoUnitario.ThousandsSeparator = True
        pnlProdutos.Controls.Add(txtPrecoUnitario)
        
        ' Botões de ação
        btnAdicionarProduto = New Button()
        btnAdicionarProduto.Text = "➕ Adicionar"
        btnAdicionarProduto.Location = New Point(800, 42)
        btnAdicionarProduto.Size = New Size(80, 26)
        btnAdicionarProduto.BackColor = Color.FromArgb(46, 204, 113)
        btnAdicionarProduto.ForeColor = Color.White
        btnAdicionarProduto.FlatStyle = FlatStyle.Flat
        btnAdicionarProduto.FlatAppearance.BorderSize = 0
        pnlProdutos.Controls.Add(btnAdicionarProduto)
        
        ' Botão remover (só fica visível quando há seleção)
        btnRemoverProduto = New Button()
        btnRemoverProduto.Text = "🗑️ Remover"
        btnRemoverProduto.Location = New Point(890, 42)
        btnRemoverProduto.Size = New Size(80, 26)
        btnRemoverProduto.BackColor = Color.FromArgb(231, 76, 60)
        btnRemoverProduto.ForeColor = Color.White
        btnRemoverProduto.FlatStyle = FlatStyle.Flat
        btnRemoverProduto.FlatAppearance.BorderSize = 0
        btnRemoverProduto.Enabled = False
        pnlProdutos.Controls.Add(btnRemoverProduto)
    End Sub
    
    ''' <summary>
    ''' Cria grid de produtos
    ''' </summary>
    Private Sub CriarGridProdutos()
        dgvProdutos = New DataGridView()
        dgvProdutos.Location = New Point(15, 80)
        dgvProdutos.Size = New Size(950, 200)
        dgvProdutos.BackgroundColor = Color.White
        dgvProdutos.BorderStyle = BorderStyle.Fixed3D
        dgvProdutos.AllowUserToAddRows = False
        dgvProdutos.AllowUserToDeleteRows = False
        dgvProdutos.ReadOnly = True
        dgvProdutos.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvProdutos.MultiSelect = False
        dgvProdutos.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        
        ' Configurar colunas
        dgvProdutos.Columns.Add("Descricao", "Descrição")
        dgvProdutos.Columns.Add("Quantidade", "Qtd")
        dgvProdutos.Columns.Add("Unidade", "Un")
        dgvProdutos.Columns.Add("PrecoUnitario", "Preço Unit.")
        dgvProdutos.Columns.Add("PrecoTotal", "Total")
        
        ' Formatação das colunas
        dgvProdutos.Columns("Quantidade").DefaultCellStyle.Format = "N2"
        dgvProdutos.Columns("PrecoUnitario").DefaultCellStyle.Format = "C2"
        dgvProdutos.Columns("PrecoTotal").DefaultCellStyle.Format = "C2"
        dgvProdutos.Columns("Quantidade").Width = 80
        dgvProdutos.Columns("Unidade").Width = 60
        dgvProdutos.Columns("PrecoUnitario").Width = 100
        dgvProdutos.Columns("PrecoTotal").Width = 100
        
        ' Estilo zebra
        dgvProdutos.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(248, 249, 250)
        dgvProdutos.DefaultCellStyle.SelectionBackColor = Color.FromArgb(52, 152, 219)
        
        pnlProdutos.Controls.Add(dgvProdutos)
    End Sub
    
    ''' <summary>
    ''' Cria painel de resumo
    ''' </summary>
    Private Sub CriarPainelResumo()
        pnlResumo = New Panel()
        pnlResumo.Size = New Size(980, 80)
        pnlResumo.Location = New Point(10, 620)
        pnlResumo.BackColor = Color.FromArgb(236, 240, 241)
        pnlResumo.BorderStyle = BorderStyle.FixedSingle
        Me.Controls.Add(pnlResumo)
        
        ' Subtotal
        Dim lblSubtotalText As New Label()
        lblSubtotalText.Text = "Subtotal:"
        lblSubtotalText.Font = New Font("Segoe UI", 10.0F, FontStyle.Regular)
        lblSubtotalText.Location = New Point(15, 15)
        lblSubtotalText.Size = New Size(80, 20)
        pnlResumo.Controls.Add(lblSubtotalText)
        
        lblSubtotal = New Label()
        lblSubtotal.Text = "R$ 0,00"
        lblSubtotal.Font = New Font("Segoe UI", 10.0F, FontStyle.Bold)
        lblSubtotal.ForeColor = Color.FromArgb(52, 73, 94)
        lblSubtotal.Location = New Point(100, 15)
        lblSubtotal.Size = New Size(100, 20)
        pnlResumo.Controls.Add(lblSubtotal)
        
        ' Desconto
        Dim lblDescontoText As New Label()
        lblDescontoText.Text = "Desconto:"
        lblDescontoText.Font = New Font("Segoe UI", 10.0F, FontStyle.Regular)
        lblDescontoText.Location = New Point(220, 15)
        lblDescontoText.Size = New Size(80, 20)
        pnlResumo.Controls.Add(lblDescontoText)
        
        txtDesconto = New NumericUpDown()
        txtDesconto.Location = New Point(310, 13)
        txtDesconto.Size = New Size(80, 23)
        txtDesconto.Minimum = 0
        txtDesconto.Maximum = 999999
        txtDesconto.DecimalPlaces = 2
        txtDesconto.ThousandsSeparator = True
        pnlResumo.Controls.Add(txtDesconto)
        
        ' Total
        Dim lblTotalText As New Label()
        lblTotalText.Text = "TOTAL:"
        lblTotalText.Font = New Font("Segoe UI", 14.0F, FontStyle.Bold)
        lblTotalText.ForeColor = Color.FromArgb(231, 76, 60)
        lblTotalText.Location = New Point(500, 10)
        lblTotalText.Size = New Size(80, 30)
        pnlResumo.Controls.Add(lblTotalText)
        
        lblTotal = New Label()
        lblTotal.Text = "R$ 0,00"
        lblTotal.Font = New Font("Segoe UI", 16.0F, FontStyle.Bold)
        lblTotal.ForeColor = Color.FromArgb(231, 76, 60)
        lblTotal.Location = New Point(590, 8)
        lblTotal.Size = New Size(150, 32)
        pnlResumo.Controls.Add(lblTotal)
        
        ' Status de produtos
        lblStatusProdutos = New Label()
        lblStatusProdutos.Text = "⚠️ Nenhum produto adicionado"
        lblStatusProdutos.Font = New Font("Segoe UI", 9.0F, FontStyle.Italic)
        lblStatusProdutos.ForeColor = Color.FromArgb(231, 76, 60)
        lblStatusProdutos.Location = New Point(15, 50)
        lblStatusProdutos.Size = New Size(300, 20)
        pnlResumo.Controls.Add(lblStatusProdutos)
    End Sub
    
    ''' <summary>
    ''' Cria painel de botões
    ''' </summary>
    Private Sub CriarPainelBotoes()
        pnlBotoes = New Panel()
        pnlBotoes.Size = New Size(980, 60)
        pnlBotoes.Location = New Point(10, 710)
        pnlBotoes.BackColor = Color.Transparent
        Me.Controls.Add(pnlBotoes)
        
        ' Botão Dados de Teste
        btnDadosTeste = New Button()
        btnDadosTeste.Text = "📝 Dados de Teste"
        btnDadosTeste.Location = New Point(20, 15)
        btnDadosTeste.Size = New Size(120, 35)
        btnDadosTeste.BackColor = Color.FromArgb(241, 196, 15)
        btnDadosTeste.ForeColor = Color.White
        btnDadosTeste.FlatStyle = FlatStyle.Flat
        btnDadosTeste.FlatAppearance.BorderSize = 0
        btnDadosTeste.Font = New Font("Segoe UI", 9.0F, FontStyle.Regular)
        pnlBotoes.Controls.Add(btnDadosTeste)
        
        ' Botão Limpar
        btnLimpar = New Button()
        btnLimpar.Text = "🗑️ Limpar Tudo"
        btnLimpar.Location = New Point(160, 15)
        btnLimpar.Size = New Size(120, 35)
        btnLimpar.BackColor = Color.FromArgb(149, 165, 166)
        btnLimpar.ForeColor = Color.White
        btnLimpar.FlatStyle = FlatStyle.Flat
        btnLimpar.FlatAppearance.BorderSize = 0
        btnLimpar.Font = New Font("Segoe UI", 9.0F, FontStyle.Regular)
        pnlBotoes.Controls.Add(btnLimpar)
        
        ' Botão Cancelar
        btnCancelar = New Button()
        btnCancelar.Text = "❌ Cancelar"
        btnCancelar.Location = New Point(700, 15)
        btnCancelar.Size = New Size(120, 35)
        btnCancelar.BackColor = Color.FromArgb(231, 76, 60)
        btnCancelar.ForeColor = Color.White
        btnCancelar.FlatStyle = FlatStyle.Flat
        btnCancelar.FlatAppearance.BorderSize = 0
        btnCancelar.Font = New Font("Segoe UI", 9.0F, FontStyle.Regular)
        btnCancelar.DialogResult = DialogResult.Cancel
        pnlBotoes.Controls.Add(btnCancelar)
        
        ' Botão Confirmar
        btnConfirmar = New Button()
        btnConfirmar.Text = "✅ Confirmar e Gerar Talão"
        btnConfirmar.Location = New Point(840, 15)
        btnConfirmar.Size = New Size(140, 35)
        btnConfirmar.BackColor = Color.FromArgb(46, 204, 113)
        btnConfirmar.ForeColor = Color.White
        btnConfirmar.FlatStyle = FlatStyle.Flat
        btnConfirmar.FlatAppearance.BorderSize = 0
        btnConfirmar.Font = New Font("Segoe UI", 9.0F, FontStyle.Bold)
        btnConfirmar.DialogResult = DialogResult.OK
        btnConfirmar.Enabled = False
        pnlBotoes.Controls.Add(btnConfirmar)
    End Sub
    
    #End Region
    
    #Region "Configuração e Validação"
    
    ''' <summary>
    ''' Configura a interface inicial
    ''' </summary>
    Private Sub ConfigurarInterface()
        ' Configura timer de validação
        _validationTimer.Interval = 500 ' 500ms depois da última digitação
        _validationTimer.Enabled = ValidacaoTempoReal
        
        ' Configurar tooltips
        ConfigurarTooltips()
        
        ' Aplicar tema
        AplicarTema()
    End Sub
    
    ''' <summary>
    ''' Configura validação em tempo real
    ''' </summary>
    Private Sub ConfigurarValidacao()
        ' Associar eventos de texto alterado
        AddHandler txtNomeCliente.TextChanged, AddressOf Campo_TextChanged
        AddHandler txtCEP.TextChanged, AddressOf Campo_TextChanged
        AddHandler txtTelefone.TextChanged, AddressOf Campo_TextChanged
        AddHandler txtDescricaoProduto.TextChanged, AddressOf Campo_TextChanged
        AddHandler txtQuantidade.ValueChanged, AddressOf Campo_ValueChanged
        AddHandler txtPrecoUnitario.ValueChanged, AddressOf Campo_ValueChanged
        AddHandler txtDesconto.ValueChanged, AddressOf Campo_ValueChanged
        
        ' Eventos específicos
        AddHandler txtCEP.Leave, AddressOf ValidarCEP
        AddHandler txtTelefone.Leave, AddressOf ValidarTelefone
    End Sub
    
    ''' <summary>
    ''' Configura tooltips para ajuda ao usuário
    ''' </summary>
    Private Sub ConfigurarTooltips()
        Dim tooltip As New ToolTip()
        tooltip.AutoPopDelay = 5000
        tooltip.InitialDelay = 1000
        tooltip.ReshowDelay = 500
        tooltip.ShowAlways = True
        
        tooltip.SetToolTip(txtNomeCliente, "Nome completo do cliente (obrigatório)")
        tooltip.SetToolTip(txtCEP, "CEP no formato 00000-000. Clique na lupa para consulta automática")
        tooltip.SetToolTip(txtTelefone, "Telefone no formato (00) 00000-0000")
        tooltip.SetToolTip(txtDescricaoProduto, "Descrição detalhada do produto")
        tooltip.SetToolTip(txtQuantidade, "Quantidade do produto (permite decimais)")
        tooltip.SetToolTip(txtPrecoUnitario, "Preço unitário do produto")
        tooltip.SetToolTip(txtDesconto, "Desconto em reais a ser aplicado no total")
        tooltip.SetToolTip(btnDadosTeste, "Carrega dados de exemplo para teste")
        tooltip.SetToolTip(btnLimpar, "Remove todos os dados do formulário")
    End Sub
    
    ''' <summary>
    ''' Aplica tema visual moderno
    ''' </summary>
    Private Sub AplicarTema()
        ' Aplicar bordas arredondadas nos painéis (simulação)
        For Each panel As Panel In {pnlCliente, pnlProdutos, pnlResumo}
            panel.Paint += Sub(sender, e)
                               Dim rect = New Rectangle(0, 0, panel.Width - 1, panel.Height - 1)
                               e.Graphics.DrawRectangle(New Pen(Color.FromArgb(189, 195, 199)), rect)
                           End Sub
        Next
    End Sub
    
    #End Region
    
    #Region "Eventos de Validação"
    
    ''' <summary>
    ''' Evento disparado quando texto é alterado
    ''' </summary>
    Private Sub Campo_TextChanged(sender As Object, e As EventArgs)
        If ValidacaoTempoReal Then
            _validationTimer.Stop()
            _validationTimer.Start()
        End If
    End Sub
    
    ''' <summary>
    ''' Evento disparado quando valor numérico é alterado
    ''' </summary>
    Private Sub Campo_ValueChanged(sender As Object, e As EventArgs)
        If ValidacaoTempoReal Then
            AtualizarTotais()
            ValidarFormulario()
        End If
    End Sub
    
    ''' <summary>
    ''' Timer de validação disparado
    ''' </summary>
    Private Sub _validationTimer_Tick(sender As Object, e As EventArgs) Handles _validationTimer.Tick
        _validationTimer.Stop()
        ValidarFormulario()
    End Sub
    
    ''' <summary>
    ''' Valida todo o formulário
    ''' </summary>
    Private Sub ValidarFormulario()
        _validationErrors.Clear()
        
        ' Validar nome do cliente
        Dim resultadoNome = ValidationSystem.ValidateRequired(txtNomeCliente.Text, "Nome do Cliente")
        If Not resultadoNome.IsValid Then
            _validationErrors("nome") = resultadoNome.ErrorMessage
            lblStatusNome.Text = "❌"
            lblStatusNome.ForeColor = Color.FromArgb(231, 76, 60)
        Else
            lblStatusNome.Text = "✅"
            lblStatusNome.ForeColor = Color.FromArgb(46, 204, 113)
        End If
        
        ' Validar produtos
        If dgvProdutos.Rows.Count = 0 Then
            _validationErrors("produtos") = "Pelo menos um produto deve ser adicionado"
            lblStatusProdutos.Text = "⚠️ Nenhum produto adicionado"
            lblStatusProdutos.ForeColor = Color.FromArgb(231, 76, 60)
        Else
            lblStatusProdutos.Text = $"✅ {dgvProdutos.Rows.Count} produto(s) adicionado(s)"
            lblStatusProdutos.ForeColor = Color.FromArgb(46, 204, 113)
        End If
        
        ' Habilitar/Desabilitar botão confirmar
        btnConfirmar.Enabled = _validationErrors.Count = 0
        
        ' Log de validação se houver erros
        If _validationErrors.Count > 0 Then
            _logger.LogDebug("EnhancedFormPDV", $"Erros de validação: {String.Join(", ", _validationErrors.Values)}")
        End If
    End Sub
    
    ''' <summary>
    ''' Valida CEP quando sai do campo
    ''' </summary>
    Private Sub ValidarCEP(sender As Object, e As EventArgs)
        If Not String.IsNullOrWhiteSpace(txtCEP.Text) Then
            Dim resultado = ValidationSystem.ValidateCEP(txtCEP.Text)
            If resultado.IsValid Then
                lblStatusCEP.Text = "✅"
                lblStatusCEP.ForeColor = Color.FromArgb(46, 204, 113)
                lblStatusCEP.Visible = True
                
                ' Se configurado, consultar CEP automaticamente
                If _config.GetConfigValuePublic("ConsultarCEPAutomatico", True) Then
                    ConsultarCEP(sender, e)
                End If
            Else
                lblStatusCEP.Text = "❌"
                lblStatusCEP.ForeColor = Color.FromArgb(231, 76, 60)
                lblStatusCEP.Visible = True
                
                Dim tooltip As New ToolTip()
                tooltip.SetToolTip(lblStatusCEP, resultado.ErrorMessage)
            End If
        Else
            lblStatusCEP.Visible = False
        End If
    End Sub
    
    ''' <summary>
    ''' Valida telefone quando sai do campo
    ''' </summary>
    Private Sub ValidarTelefone(sender As Object, e As EventArgs)
        If Not String.IsNullOrWhiteSpace(txtTelefone.Text) Then
            Dim resultado = ValidationSystem.ValidatePhone(txtTelefone.Text)
            If resultado.IsValid Then
                lblStatusTelefone.Text = "✅"
                lblStatusTelefone.ForeColor = Color.FromArgb(46, 204, 113)
                lblStatusTelefone.Visible = True
            Else
                lblStatusTelefone.Text = "❌"
                lblStatusTelefone.ForeColor = Color.FromArgb(231, 76, 60)
                lblStatusTelefone.Visible = True
                
                Dim tooltip As New ToolTip()
                tooltip.SetToolTip(lblStatusTelefone, resultado.ErrorMessage)
            End If
        Else
            lblStatusTelefone.Visible = False
        End If
    End Sub
    
    #End Region
    
    #Region "Eventos de Ações"
    
    ''' <summary>
    ''' Adiciona produto à lista
    ''' </summary>
    Private Sub btnAdicionarProduto_Click(sender As Object, e As EventArgs) Handles btnAdicionarProduto.Click
        Try
            ' Validar campos do produto
            If String.IsNullOrWhiteSpace(txtDescricaoProduto.Text) Then
                MessageBox.Show("Descrição do produto é obrigatória", "Validação", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtDescricaoProduto.Focus()
                Return
            End If
            
            If txtQuantidade.Value <= 0 Then
                MessageBox.Show("Quantidade deve ser maior que zero", "Validação", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtQuantidade.Focus()
                Return
            End If
            
            If txtPrecoUnitario.Value <= 0 Then
                MessageBox.Show("Preço unitário deve ser maior que zero", "Validação", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtPrecoUnitario.Focus()
                Return
            End If
            
            ' Adicionar produto ao grid
            Dim precoTotal = txtQuantidade.Value * txtPrecoUnitario.Value
            dgvProdutos.Rows.Add(
                txtDescricaoProduto.Text,
                txtQuantidade.Value,
                cmbUnidade.Text,
                txtPrecoUnitario.Value,
                precoTotal
            )
            
            ' Limpar campos
            txtDescricaoProduto.Clear()
            txtQuantidade.Value = 1
            txtPrecoUnitario.Value = 0
            txtDescricaoProduto.Focus()
            
            ' Atualizar totais
            AtualizarTotais()
            ValidarFormulario()
            
            _logger.LogInfo("EnhancedFormPDV", $"Produto adicionado: {txtDescricaoProduto.Text}")
            
        Catch ex As Exception
            _logger.LogError("EnhancedFormPDV", "Erro ao adicionar produto", ex)
            MessageBox.Show($"Erro ao adicionar produto: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    ''' <summary>
    ''' Remove produto selecionado
    ''' </summary>
    Private Sub btnRemoverProduto_Click(sender As Object, e As EventArgs) Handles btnRemoverProduto.Click
        Try
            If dgvProdutos.SelectedRows.Count > 0 Then
                Dim resultado = MessageBox.Show("Confirma a remoção do produto selecionado?", "Confirmação", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If resultado = DialogResult.Yes Then
                    dgvProdutos.Rows.RemoveAt(dgvProdutos.SelectedRows(0).Index)
                    AtualizarTotais()
                    ValidarFormulario()
                    _logger.LogInfo("EnhancedFormPDV", "Produto removido")
                End If
            End If
        Catch ex As Exception
            _logger.LogError("EnhancedFormPDV", "Erro ao remover produto", ex)
            MessageBox.Show($"Erro ao remover produto: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    ''' <summary>
    ''' Atualiza totais no painel de resumo
    ''' </summary>
    Private Sub AtualizarTotais()
        Try
            Dim subtotal As Decimal = 0
            For Each row As DataGridViewRow In dgvProdutos.Rows
                If row.Cells("PrecoTotal").Value IsNot Nothing Then
                    subtotal += Convert.ToDecimal(row.Cells("PrecoTotal").Value)
                End If
            Next
            
            Dim desconto = txtDesconto.Value
            Dim total = subtotal - desconto
            
            lblSubtotal.Text = subtotal.ToString("C2")
            lblTotal.Text = total.ToString("C2")
            
            ' Mudar cor do total se negativo
            If total < 0 Then
                lblTotal.ForeColor = Color.FromArgb(231, 76, 60)
            Else
                lblTotal.ForeColor = Color.FromArgb(46, 204, 113)
            End If
            
        Catch ex As Exception
            _logger.LogError("EnhancedFormPDV", "Erro ao atualizar totais", ex)
        End Try
    End Sub
    
    ''' <summary>
    ''' Consulta CEP automaticamente
    ''' </summary>
    Private Sub ConsultarCEP(sender As Object, e As EventArgs)
        Try
            If Not String.IsNullOrWhiteSpace(txtCEP.Text) AndAlso txtCEP.MaskCompleted Then
                ' TODO: Implementar consulta real de CEP via API
                ' Por enquanto, simular preenchimento
                If txtCEP.Text.StartsWith("53") Then ' CEP de Paulista/PE
                    If String.IsNullOrWhiteSpace(txtCidade.Text) Then
                        txtCidade.Text = "Paulista/PE"
                    End If
                End If
                
                _logger.LogDebug("EnhancedFormPDV", $"CEP consultado: {txtCEP.Text}")
            End If
        Catch ex As Exception
            _logger.LogError("EnhancedFormPDV", "Erro ao consultar CEP", ex)
        End Try
    End Sub
    
    ''' <summary>
    ''' Carrega dados de teste
    ''' </summary>
    Private Sub btnDadosTeste_Click(sender As Object, e As EventArgs) Handles btnDadosTeste.Click
        Try
            ' Dados do cliente
            txtNomeCliente.Text = "João Silva Santos - TESTE"
            txtEnderecoCliente.Text = "Rua das Madeiras, 456 - Centro"
            txtCEP.Text = "53401-123"
            txtCidade.Text = "Paulista/PE"
            txtTelefone.Text = "(81) 98765-4321"
            cmbFormaPagamento.SelectedItem = "À Vista"
            
            ' Limpar produtos existentes
            dgvProdutos.Rows.Clear()
            
            ' Adicionar produtos de teste
            AdicionarProdutoTeste("Tábua de Pinus 2,5x30x300cm", 10, "UN", 28.50)
            AdicionarProdutoTeste("Ripão 5x5x300cm", 20, "UN", 18.75)
            AdicionarProdutoTeste("Compensado Naval 15mm", 3, "M²", 89.90)
            AdicionarProdutoTeste("Caibro 5x6x300cm", 15, "UN", 22.30)
            
            AtualizarTotais()
            ValidarFormulario()
            
            MessageBox.Show("✅ Dados de teste carregados com sucesso!", "Teste", MessageBoxButtons.OK, MessageBoxIcon.Information)
            _logger.LogInfo("EnhancedFormPDV", "Dados de teste carregados")
            
        Catch ex As Exception
            _logger.LogError("EnhancedFormPDV", "Erro ao carregar dados de teste", ex)
            MessageBox.Show($"Erro ao carregar dados de teste: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    ''' <summary>
    ''' Adiciona produto de teste ao grid
    ''' </summary>
    Private Sub AdicionarProdutoTeste(descricao As String, quantidade As Decimal, unidade As String, preco As Decimal)
        Dim total = quantidade * preco
        dgvProdutos.Rows.Add(descricao, quantidade, unidade, preco, total)
    End Sub
    
    ''' <summary>
    ''' Limpa todos os dados do formulário
    ''' </summary>
    Private Sub btnLimpar_Click(sender As Object, e As EventArgs) Handles btnLimpar.Click
        Dim resultado = MessageBox.Show("Isso irá limpar todos os dados do formulário. Confirma?", "Limpar Dados", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If resultado = DialogResult.Yes Then
            LimparFormulario()
            _logger.LogInfo("EnhancedFormPDV", "Formulário limpo pelo usuário")
        End If
    End Sub
    
    ''' <summary>
    ''' Limpa todos os campos do formulário
    ''' </summary>
    Private Sub LimparFormulario()
        ' Dados do cliente
        txtNomeCliente.Clear()
        txtEnderecoCliente.Clear()
        txtCEP.Clear()
        txtCidade.Clear()
        txtTelefone.Clear()
        cmbFormaPagamento.SelectedIndex = 0
        txtVendedor.Text = _config.VendedorPadrao
        
        ' Produtos
        txtDescricaoProduto.Clear()
        txtQuantidade.Value = 1
        cmbUnidade.SelectedIndex = 0
        txtPrecoUnitario.Value = 0
        dgvProdutos.Rows.Clear()
        
        ' Totais
        txtDesconto.Value = 0
        AtualizarTotais()
        
        ' Status
        lblStatusNome.Text = "⚠️"
        lblStatusCEP.Visible = False
        lblStatusTelefone.Visible = False
        
        ValidarFormulario()
    End Sub
    
    #End Region
    
    #Region "Finalização"
    
    ''' <summary>
    ''' Coleta dados finais quando confirma
    ''' </summary>
    Private Sub btnConfirmar_Click(sender As Object, e As EventArgs) Handles btnConfirmar.Click
        Try
            ' Validação final
            If Not ValidarFormularioFinal() Then
                Return
            End If
            
            ' Coletar dados
            ColetarDados()
            
            _logger.LogInfo("EnhancedFormPDV", "Dados coletados com sucesso para geração de talão")
            
        Catch ex As Exception
            _logger.LogError("EnhancedFormPDV", "Erro ao coletar dados", ex)
            MessageBox.Show($"Erro ao processar dados: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    ''' <summary>
    ''' Validação final antes de confirmar
    ''' </summary>
    Private Function ValidarFormularioFinal() As Boolean
        Dim erros As New List(Of String)()
        
        ' Validações obrigatórias
        If String.IsNullOrWhiteSpace(txtNomeCliente.Text) Then
            erros.Add("Nome do cliente é obrigatório")
        End If
        
        If dgvProdutos.Rows.Count = 0 Then
            erros.Add("Pelo menos um produto deve ser adicionado")
        End If
        
        ' Validar total
        Dim total = Convert.ToDecimal(lblTotal.Text.Replace("R$", "").Replace(".", "").Replace(",", "."))
        If total <= 0 Then
            erros.Add("Total da venda deve ser maior que zero")
        End If
        
        If erros.Count > 0 Then
            MessageBox.Show($"Corrija os seguintes erros:{Environment.NewLine}{String.Join(Environment.NewLine, erros)}", 
                          "Validação", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return False
        End If
        
        Return True
    End Function
    
    ''' <summary>
    ''' Coleta dados do formulário
    ''' </summary>
    Private Sub ColetarDados()
        DadosColetados = New DadosTalao()
        
        ' Dados do cliente
        DadosColetados.NomeCliente = txtNomeCliente.Text.Trim()
        DadosColetados.EnderecoCliente = txtEnderecoCliente.Text.Trim()
        DadosColetados.CEP = txtCEP.Text.Trim()
        DadosColetados.Cidade = txtCidade.Text.Trim()
        DadosColetados.Telefone = txtTelefone.Text.Trim()
        DadosColetados.FormaPagamento = cmbFormaPagamento.Text
        DadosColetados.Vendedor = txtVendedor.Text.Trim()
        DadosColetados.DataVenda = Date.Now
        DadosColetados.NumeroTalao = Date.Now.ToString("yyyyMMddHHmmss")
        DadosColetados.Desconto = txtDesconto.Value
        
        ' Produtos
        DadosColetados.Produtos = New List(Of ProdutoTalao)()
        For Each row As DataGridViewRow In dgvProdutos.Rows
            Dim produto As New ProdutoTalao()
            produto.Descricao = row.Cells("Descricao").Value.ToString()
            produto.Quantidade = Convert.ToDouble(row.Cells("Quantidade").Value)
            produto.Unidade = row.Cells("Unidade").Value.ToString()
            produto.PrecoUnitario = Convert.ToDecimal(row.Cells("PrecoUnitario").Value)
            produto.PrecoTotal = Convert.ToDecimal(row.Cells("PrecoTotal").Value)
            DadosColetados.Produtos.Add(produto)
        Next
        
        ' Calcular total geral
        DadosColetados.TotalGeral = DadosColetados.Produtos.Sum(Function(p) p.PrecoTotal) - DadosColetados.Desconto
    End Sub
    
    ''' <summary>
    ''' Evento ao selecionar produto no grid
    ''' </summary>
    Private Sub dgvProdutos_SelectionChanged(sender As Object, e As EventArgs) Handles dgvProdutos.SelectionChanged
        btnRemoverProduto.Enabled = dgvProdutos.SelectedRows.Count > 0
    End Sub
    
    #End Region
End Class