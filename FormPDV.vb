Imports System.Windows.Forms
Imports System.Drawing
Imports System.Configuration

''' <summary>
''' Formulário de entrada de dados para geração de talão
''' Interface para coleta de dados do cliente e produtos
''' Versão otimizada com validação em tempo real e melhor UX
''' </summary>
Public Class FormPDV
    Inherits Form

    ' Controles da interface
    Private WithEvents pnlCliente As Panel
    Private WithEvents pnlProdutos As Panel
    Private WithEvents pnlBotoes As Panel
    Private WithEvents dgvProdutos As DataGridView

    ' Dados do cliente
    Private WithEvents txtNomeCliente As TextBox
    Private WithEvents txtEnderecoCliente As TextBox
    Private WithEvents txtCEP As TextBox
    Private WithEvents txtCidade As TextBox
    Private WithEvents txtTelefone As TextBox

    ' Produtos
    Private WithEvents txtDescricaoProduto As TextBox
    Private WithEvents txtQuantidade As TextBox
    Private WithEvents cmbUnidade As ComboBox
    Private WithEvents txtPrecoUnitario As TextBox
    Private WithEvents btnAdicionarProduto As Button
    Private WithEvents btnRemoverProduto As Button

    ' Forma de pagamento
    Private WithEvents cmbFormaPagamento As ComboBox
    Private WithEvents txtVendedor As TextBox

    ' Botões principais
    Private WithEvents btnConfirmar As Button
    Private WithEvents btnCancelar As Button
    Private WithEvents btnDadosTeste As Button
    
    ' Labels de status de validação
    Private lblStatusNome As Label
    Private lblStatusCEP As Label
    Private lblStatusTelefone As Label
    Private lblStatusProdutos As Label

    ' Propriedade para retornar dados coletados
    Public Property DadosColetados As DadosTalao
    
    ' Sistema de logging e validação
    Private ReadOnly _logger As LoggingSystem = LoggingSystem.Instance
    Private _validationErrors As New Dictionary(Of String, String)()

    ''' <summary>
    ''' Construtor do formulário
    ''' </summary>
    Public Sub New()
        InitializeComponent()
        ConfigurarInterface()
        DadosColetados = New DadosTalao()
    End Sub

    ''' <summary>
    ''' Inicializa os componentes da interface
    ''' </summary>
    Private Sub InitializeComponent()
        ' Configurações do formulário
        Me.Text = "Entrada de Dados - Talão de Venda"
        Me.Size = New Size(900, 700)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.BackColor = Color.WhiteSmoke
        Me.Font = New Font("Segoe UI", 9.0F, FontStyle.Regular)

        ' Painel dados do cliente
        pnlCliente = New Panel()
        pnlCliente.Size = New Size(860, 200)
        pnlCliente.Location = New Point(20, 20)
        pnlCliente.BackColor = Color.White
        pnlCliente.BorderStyle = BorderStyle.FixedSingle
        Me.Controls.Add(pnlCliente)

        ' Título seção cliente
        Dim lblCliente As New Label()
        lblCliente.Text = "👤 DADOS DO CLIENTE"
        lblCliente.Font = New Font("Segoe UI", 12.0F, FontStyle.Bold)
        lblCliente.ForeColor = Color.FromArgb(52, 73, 94)
        lblCliente.Size = New Size(200, 25)
        lblCliente.Location = New Point(10, 10)
        pnlCliente.Controls.Add(lblCliente)

        ' Campo Nome do Cliente
        Dim lblNome As New Label()
        lblNome.Text = "Nome do Cliente:"
        lblNome.Location = New Point(20, 50)
        lblNome.Size = New Size(120, 20)
        pnlCliente.Controls.Add(lblNome)

        txtNomeCliente = New TextBox()
        txtNomeCliente.Location = New Point(150, 48)
        txtNomeCliente.Size = New Size(400, 23)
        txtNomeCliente.Font = New Font("Segoe UI", 10.0F)
        pnlCliente.Controls.Add(txtNomeCliente)

        ' Campo Endereço
        Dim lblEndereco As New Label()
        lblEndereco.Text = "Endereço:"
        lblEndereco.Location = New Point(20, 80)
        lblEndereco.Size = New Size(120, 20)
        pnlCliente.Controls.Add(lblEndereco)

        txtEnderecoCliente = New TextBox()
        txtEnderecoCliente.Location = New Point(150, 78)
        txtEnderecoCliente.Size = New Size(500, 23)
        txtEnderecoCliente.Font = New Font("Segoe UI", 10.0F)
        pnlCliente.Controls.Add(txtEnderecoCliente)

        ' Campo CEP
        Dim lblCEP As New Label()
        lblCEP.Text = "CEP:"
        lblCEP.Location = New Point(20, 110)
        lblCEP.Size = New Size(120, 20)
        pnlCliente.Controls.Add(lblCEP)

        txtCEP = New TextBox()
        txtCEP.Location = New Point(150, 108)
        txtCEP.Size = New Size(120, 23)
        txtCEP.Font = New Font("Segoe UI", 10.0F)
        pnlCliente.Controls.Add(txtCEP)

        ' Campo Cidade
        Dim lblCidade As New Label()
        lblCidade.Text = "Cidade:"
        lblCidade.Location = New Point(300, 110)
        lblCidade.Size = New Size(50, 20)
        pnlCliente.Controls.Add(lblCidade)

        txtCidade = New TextBox()
        txtCidade.Location = New Point(360, 108)
        txtCidade.Size = New Size(200, 23)
        txtCidade.Font = New Font("Segoe UI", 10.0F)
        pnlCliente.Controls.Add(txtCidade)

        ' Campo Telefone
        Dim lblTelefone As New Label()
        lblTelefone.Text = "Telefone:"
        lblTelefone.Location = New Point(20, 140)
        lblTelefone.Size = New Size(120, 20)
        pnlCliente.Controls.Add(lblTelefone)

        txtTelefone = New TextBox()
        txtTelefone.Location = New Point(150, 138)
        txtTelefone.Size = New Size(150, 23)
        txtTelefone.Font = New Font("Segoe UI", 10.0F)
        pnlCliente.Controls.Add(txtTelefone)

        ' Botão dados de teste
        btnDadosTeste = New Button()
        btnDadosTeste.Text = "📝 Carregar Dados de Teste"
        btnDadosTeste.Location = New Point(580, 50)
        btnDadosTeste.Size = New Size(160, 30)
        btnDadosTeste.BackColor = Color.FromArgb(241, 196, 15)
        btnDadosTeste.ForeColor = Color.White
        btnDadosTeste.FlatStyle = FlatStyle.Flat
        btnDadosTeste.FlatAppearance.BorderSize = 0
        pnlCliente.Controls.Add(btnDadosTeste)

        ' Painel produtos
        pnlProdutos = New Panel()
        pnlProdutos.Size = New Size(860, 350)
        pnlProdutos.Location = New Point(20, 240)
        pnlProdutos.BackColor = Color.White
        pnlProdutos.BorderStyle = BorderStyle.FixedSingle
        Me.Controls.Add(pnlProdutos)

        ' Título seção produtos
        Dim lblProdutos As New Label()
        lblProdutos.Text = "📦 PRODUTOS"
        lblProdutos.Font = New Font("Segoe UI", 12.0F, FontStyle.Bold)
        lblProdutos.ForeColor = Color.FromArgb(52, 73, 94)
        lblProdutos.Size = New Size(150, 25)
        lblProdutos.Location = New Point(10, 10)
        pnlProdutos.Controls.Add(lblProdutos)

        ' Campos para adicionar produto
        Dim lblDescricao As New Label()
        lblDescricao.Text = "Descrição:"
        lblDescricao.Location = New Point(20, 50)
        lblDescricao.Size = New Size(70, 20)
        pnlProdutos.Controls.Add(lblDescricao)

        txtDescricaoProduto = New TextBox()
        txtDescricaoProduto.Location = New Point(95, 48)
        txtDescricaoProduto.Size = New Size(300, 23)
        txtDescricaoProduto.Font = New Font("Segoe UI", 10.0F)
        pnlProdutos.Controls.Add(txtDescricaoProduto)

        Dim lblQtd As New Label()
        lblQtd.Text = "Qtd:"
        lblQtd.Location = New Point(410, 50)
        lblQtd.Size = New Size(30, 20)
        pnlProdutos.Controls.Add(lblQtd)

        txtQuantidade = New TextBox()
        txtQuantidade.Location = New Point(445, 48)
        txtQuantidade.Size = New Size(60, 23)
        txtQuantidade.Font = New Font("Segoe UI", 10.0F)
        txtQuantidade.Text = "1"
        pnlProdutos.Controls.Add(txtQuantidade)

        Dim lblUnidade As New Label()
        lblUnidade.Text = "Un:"
        lblUnidade.Location = New Point(520, 50)
        lblUnidade.Size = New Size(25, 20)
        pnlProdutos.Controls.Add(lblUnidade)

        cmbUnidade = New ComboBox()
        cmbUnidade.Location = New Point(550, 48)
        cmbUnidade.Size = New Size(60, 23)
        cmbUnidade.DropDownStyle = ComboBoxStyle.DropDownList
        cmbUnidade.Items.AddRange({"UN", "M", "M²", "M³", "PC", "CX", "KG"})
        cmbUnidade.SelectedIndex = 0
        pnlProdutos.Controls.Add(cmbUnidade)

        Dim lblPreco As New Label()
        lblPreco.Text = "Preço:"
        lblPreco.Location = New Point(625, 50)
        lblPreco.Size = New Size(40, 20)
        pnlProdutos.Controls.Add(lblPreco)

        txtPrecoUnitario = New TextBox()
        txtPrecoUnitario.Location = New Point(670, 48)
        txtPrecoUnitario.Size = New Size(80, 23)
        txtPrecoUnitario.Font = New Font("Segoe UI", 10.0F)
        txtPrecoUnitario.Text = "0,00"
        pnlProdutos.Controls.Add(txtPrecoUnitario)

        btnAdicionarProduto = New Button()
        btnAdicionarProduto.Text = "+"
        btnAdicionarProduto.Location = New Point(760, 48)
        btnAdicionarProduto.Size = New Size(30, 23)
        btnAdicionarProduto.BackColor = Color.FromArgb(46, 204, 113)
        btnAdicionarProduto.ForeColor = Color.White
        btnAdicionarProduto.FlatStyle = FlatStyle.Flat
        btnAdicionarProduto.FlatAppearance.BorderSize = 0
        pnlProdutos.Controls.Add(btnAdicionarProduto)

        btnRemoverProduto = New Button()
        btnRemoverProduto.Text = "-"
        btnRemoverProduto.Location = New Point(800, 48)
        btnRemoverProduto.Size = New Size(30, 23)
        btnRemoverProduto.BackColor = Color.FromArgb(231, 76, 60)
        btnRemoverProduto.ForeColor = Color.White
        btnRemoverProduto.FlatStyle = FlatStyle.Flat
        btnRemoverProduto.FlatAppearance.BorderSize = 0
        pnlProdutos.Controls.Add(btnRemoverProduto)

        ' DataGridView para lista de produtos
        dgvProdutos = New DataGridView()
        dgvProdutos.Location = New Point(20, 85)
        dgvProdutos.Size = New Size(820, 200)
        dgvProdutos.AllowUserToAddRows = False
        dgvProdutos.AllowUserToDeleteRows = False
        dgvProdutos.ReadOnly = True
        dgvProdutos.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvProdutos.BackgroundColor = Color.White
        dgvProdutos.BorderStyle = BorderStyle.Fixed3D
        pnlProdutos.Controls.Add(dgvProdutos)

        ' Configurar colunas do DataGridView
        dgvProdutos.Columns.Add("Descricao", "Descrição")
        dgvProdutos.Columns.Add("Quantidade", "Qtd")
        dgvProdutos.Columns.Add("Unidade", "Un")
        dgvProdutos.Columns.Add("PrecoUnitario", "Preço Unit.")
        dgvProdutos.Columns.Add("PrecoTotal", "Total")

        dgvProdutos.Columns(0).Width = 350
        dgvProdutos.Columns(1).Width = 80
        dgvProdutos.Columns(2).Width = 60
        dgvProdutos.Columns(3).Width = 100
        dgvProdutos.Columns(4).Width = 100

        ' Forma de pagamento e vendedor
        Dim lblFormaPgto As New Label()
        lblFormaPgto.Text = "Forma de Pagamento:"
        lblFormaPgto.Location = New Point(20, 300)
        lblFormaPgto.Size = New Size(130, 20)
        pnlProdutos.Controls.Add(lblFormaPgto)

        cmbFormaPagamento = New ComboBox()
        cmbFormaPagamento.Location = New Point(155, 298)
        cmbFormaPagamento.Size = New Size(150, 23)
        cmbFormaPagamento.DropDownStyle = ComboBoxStyle.DropDownList
        cmbFormaPagamento.Items.AddRange({"Dinheiro", "Cartão de Débito", "Cartão de Crédito", "PIX", "Boleto", "Fiado", "Cheque"})
        cmbFormaPagamento.SelectedIndex = 0
        pnlProdutos.Controls.Add(cmbFormaPagamento)

        Dim lblVendedor As New Label()
        lblVendedor.Text = "Vendedor:"
        lblVendedor.Location = New Point(330, 300)
        lblVendedor.Size = New Size(60, 20)
        pnlProdutos.Controls.Add(lblVendedor)

        txtVendedor = New TextBox()
        txtVendedor.Location = New Point(395, 298)
        txtVendedor.Size = New Size(200, 23)
        txtVendedor.Font = New Font("Segoe UI", 10.0F)
        txtVendedor.Text = ConfigurationManager.AppSettings("VendedorPadrao")
        pnlProdutos.Controls.Add(txtVendedor)

        ' Painel botões
        pnlBotoes = New Panel()
        pnlBotoes.Size = New Size(860, 60)
        pnlBotoes.Location = New Point(20, 610)
        pnlBotoes.BackColor = Color.WhiteSmoke
        Me.Controls.Add(pnlBotoes)

        btnConfirmar = New Button()
        btnConfirmar.Text = "✅ CONFIRMAR E GERAR TALÃO"
        btnConfirmar.Location = New Point(500, 15)
        btnConfirmar.Size = New Size(200, 35)
        btnConfirmar.BackColor = Color.FromArgb(46, 204, 113)
        btnConfirmar.ForeColor = Color.White
        btnConfirmar.Font = New Font("Segoe UI", 10.0F, FontStyle.Bold)
        btnConfirmar.FlatStyle = FlatStyle.Flat
        btnConfirmar.FlatAppearance.BorderSize = 0
        pnlBotoes.Controls.Add(btnConfirmar)

        btnCancelar = New Button()
        btnCancelar.Text = "❌ Cancelar"
        btnCancelar.Location = New Point(720, 15)
        btnCancelar.Size = New Size(120, 35)
        btnCancelar.BackColor = Color.FromArgb(231, 76, 60)
        btnCancelar.ForeColor = Color.White
        btnCancelar.Font = New Font("Segoe UI", 10.0F, FontStyle.Bold)
        btnCancelar.FlatStyle = FlatStyle.Flat
        btnCancelar.FlatAppearance.BorderSize = 0
        pnlBotoes.Controls.Add(btnCancelar)
    End Sub

    ''' <summary>
    ''' Configura detalhes adicionais da interface
    ''' </summary>
    Private Sub ConfigurarInterface()
        ' Adicionar alguns produtos de exemplo para a madeireira
        CarregarProdutosTeste()
    End Sub

    ''' <summary>
    ''' Carrega dados de teste para facilitar demonstração
    ''' </summary>
    Private Sub btnDadosTeste_Click(sender As Object, e As EventArgs) Handles btnDadosTeste.Click
        txtNomeCliente.Text = "João Silva - TESTE"
        txtEnderecoCliente.Text = "Rua das Árvores, 123 - Centro"
        txtCEP.Text = "55431-165"
        txtCidade.Text = "Paulista/PE"
        txtTelefone.Text = "(81) 9876-5432"

        ' Limpar produtos existentes
        dgvProdutos.Rows.Clear()

        ' Adicionar produtos de teste
        AdicionarProdutoNaGrid("Tábua de Pinus 2x4m", 5, "UN", 25.0)
        AdicionarProdutoNaGrid("Ripão 3x3x3m", 10, "UN", 15.0)
        AdicionarProdutoNaGrid("Compensado 18mm", 2, "M²", 45.0)

        MessageBox.Show("✅ Dados de teste carregados!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    ''' <summary>
    ''' Adiciona produto no DataGridView
    ''' </summary>
    Private Sub AdicionarProdutoNaGrid(descricao As String, quantidade As Double, unidade As String, precoUnit As Double)
        Dim precoTotal As Double = quantidade * precoUnit
        dgvProdutos.Rows.Add(descricao, quantidade.ToString("0.##"), unidade, 
                           precoUnit.ToString("C"), precoTotal.ToString("C"))
    End Sub

    ''' <summary>
    ''' Carrega produtos típicos de madeireira para facilitar seleção
    ''' </summary>
    Private Sub CarregarProdutosTeste()
        ' Esta função poderia carregar de um banco de dados
        ' Por enquanto, deixaremos para entrada manual
    End Sub

    ''' <summary>
    ''' Adiciona produto à lista
    ''' </summary>
    Private Sub btnAdicionarProduto_Click(sender As Object, e As EventArgs) Handles btnAdicionarProduto.Click
        Try
            If String.IsNullOrWhiteSpace(txtDescricaoProduto.Text) Then
                MessageBox.Show("Digite a descrição do produto", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtDescricaoProduto.Focus()
                Return
            End If

            Dim quantidade As Double = Convert.ToDouble(txtQuantidade.Text.Replace(",", "."))
            Dim precoUnit As Double = Convert.ToDouble(txtPrecoUnitario.Text.Replace("R$", "").Replace(",", ".").Trim())

            AdicionarProdutoNaGrid(txtDescricaoProduto.Text, quantidade, cmbUnidade.Text, precoUnit)

            ' Limpar campos
            txtDescricaoProduto.Text = ""
            txtQuantidade.Text = "1"
            txtPrecoUnitario.Text = "0,00"
            txtDescricaoProduto.Focus()

        Catch ex As Exception
            MessageBox.Show("Erro ao adicionar produto: " & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' Remove produto selecionado
    ''' </summary>
    Private Sub btnRemoverProduto_Click(sender As Object, e As EventArgs) Handles btnRemoverProduto.Click
        If dgvProdutos.SelectedRows.Count > 0 Then
            dgvProdutos.Rows.RemoveAt(dgvProdutos.SelectedRows(0).Index)
        Else
            MessageBox.Show("Selecione um produto para remover", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub

    ''' <summary>
    ''' Confirma dados e fecha formulário
    ''' </summary>
    Private Sub btnConfirmar_Click(sender As Object, e As EventArgs) Handles btnConfirmar.Click
        Try
            ' Validar dados obrigatórios
            If String.IsNullOrWhiteSpace(txtNomeCliente.Text) Then
                MessageBox.Show("Digite o nome do cliente", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtNomeCliente.Focus()
                Return
            End If

            If dgvProdutos.Rows.Count = 0 Then
                MessageBox.Show("Adicione pelo menos um produto", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtDescricaoProduto.Focus()
                Return
            End If

            ' Coletar dados do formulário
            DadosColetados.NomeCliente = txtNomeCliente.Text
            DadosColetados.EnderecoCliente = txtEnderecoCliente.Text
            DadosColetados.CEP = txtCEP.Text
            DadosColetados.Cidade = txtCidade.Text
            DadosColetados.Telefone = txtTelefone.Text
            DadosColetados.FormaPagamento = cmbFormaPagamento.Text
            DadosColetados.Vendedor = txtVendedor.Text

            ' Coletar produtos
            DadosColetados.Produtos.Clear()
            For Each row As DataGridViewRow In dgvProdutos.Rows
                Dim produto As New ProdutoTalao()
                produto.Descricao = row.Cells("Descricao").Value.ToString()
                produto.Quantidade = Convert.ToDouble(row.Cells("Quantidade").Value.ToString())
                produto.Unidade = row.Cells("Unidade").Value.ToString()
                produto.PrecoUnitario = Convert.ToDouble(row.Cells("PrecoUnitario").Value.ToString().Replace("R$", "").Replace(",", ".").Trim())
                produto.PrecoTotal = produto.Quantidade * produto.PrecoUnitario
                DadosColetados.Produtos.Add(produto)
            Next

            ' Fechar formulário com sucesso
            Me.DialogResult = DialogResult.OK
            Me.Close()

        Catch ex As Exception
            MessageBox.Show("Erro ao validar dados: " & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' Cancela operação
    ''' </summary>
    Private Sub btnCancelar_Click(sender As Object, e As EventArgs) Handles btnCancelar.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub

    ''' <summary>
    ''' Evento ao pressionar Enter nos campos de texto
    ''' </summary>
    Private Sub txtDescricaoProduto_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtDescricaoProduto.KeyPress
        If e.KeyChar = Chr(13) Then ' Enter
            txtQuantidade.Focus()
        End If
    End Sub

    Private Sub txtQuantidade_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtQuantidade.KeyPress
        If e.KeyChar = Chr(13) Then ' Enter
            txtPrecoUnitario.Focus()
        End If
    End Sub

    Private Sub txtPrecoUnitario_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPrecoUnitario.KeyPress
        If e.KeyChar = Chr(13) Then ' Enter
            btnAdicionarProduto_Click(sender, New EventArgs())
        End If
    End Sub
End Class