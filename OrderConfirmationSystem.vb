Imports System.Windows.Forms
Imports System.Drawing

''' <summary>
''' Sistema de confirmação de pedidos integrado
''' Gerencia preenchimento automático, validação e confirmação de vendas
''' </summary>
Public Class OrderConfirmationManager
    Private _venda As Venda
    Private _calculadora As CalculadoraMadeireira
    Private _database As DatabaseManager
    
    Public Sub New()
        _database = DatabaseManager.Instance
    End Sub
    
    ''' <summary>
    ''' Configura o gerenciador para uma venda específica
    ''' </summary>
    Public Sub ConfigurarVenda(venda As Venda, calculadora As CalculadoraMadeireira)
        _venda = venda
        _calculadora = calculadora
    End Sub
    
    ''' <summary>
    ''' Valida todos os dados da venda antes da confirmação
    ''' </summary>
    Public Function ValidarVenda() As List(Of String)
        Dim erros As New List(Of String)()
        
        Try
            ' Validar dados do cliente
            If String.IsNullOrEmpty(_venda.Cliente.Nome) Then
                erros.Add("Nome do cliente é obrigatório")
            End If
            
            ' Validar itens
            If _venda.Itens.Count = 0 Then
                erros.Add("Adicione pelo menos um item à venda")
            End If
            
            ' Validar cada item
            For Each item In _venda.Itens
                If item.Quantidade <= 0 Then
                    erros.Add($"Quantidade inválida para {item.Produto.Descricao}")
                End If
                
                If item.PrecoUnitario <= 0 Then
                    erros.Add($"Preço inválido para {item.Produto.Descricao}")
                End If
            Next
            
            ' Validar cálculos
            If _calculadora IsNot Nothing Then
                erros.AddRange(_calculadora.ValidarCalculos())
            End If
            
            ' Validar forma de pagamento
            If String.IsNullOrEmpty(_venda.FormaPagamento) Then
                erros.Add("Selecione uma forma de pagamento")
            End If
            
            ' Validar vendedor
            If String.IsNullOrEmpty(_venda.VendedorNome) Then
                erros.Add("Informe o vendedor")
            End If
            
        Catch ex As Exception
            erros.Add($"Erro na validação: {ex.Message}")
        End Try
        
        Return erros
    End Function
    
    ''' <summary>
    ''' Preenche automaticamente dados padrão da venda
    ''' </summary>
    Public Sub PreencherDadosAutomaticos()
        Try
            ' Número do talão
            If String.IsNullOrEmpty(_venda.NumeroTalao) Then
                _venda.NumeroTalao = GerarNumeroTalao()
            End If
            
            ' Data da venda
            If _venda.DataVenda = Date.MinValue Then
                _venda.DataVenda = Date.Now
            End If
            
            ' Status padrão
            If String.IsNullOrEmpty(_venda.Status) Then
                _venda.Status = "Confirmado"
            End If
            
            ' Forma de pagamento padrão
            If String.IsNullOrEmpty(_venda.FormaPagamento) Then
                _venda.FormaPagamento = "À Vista"
            End If
            
            ' Vendedor padrão se não informado
            If String.IsNullOrEmpty(_venda.VendedorNome) Then
                Dim config As New ConfiguracaoSistema()
                _venda.VendedorNome = config.VendedorPadrao
            End If
            
        Catch ex As Exception
            Console.WriteLine($"Erro ao preencher dados automáticos: {ex.Message}")
        End Try
    End Sub
    
    ''' <summary>
    ''' Confirma a venda e executa todas as operações necessárias
    ''' </summary>
    Public Function ConfirmarVenda() As Boolean
        Try
            ' Preencher dados automáticos
            PreencherDadosAutomaticos()
            
            ' Validar venda
            Dim erros = ValidarVenda()
            If erros.Count > 0 Then
                MessageBox.Show($"Corrija os seguintes erros:{Environment.NewLine}{String.Join(Environment.NewLine, erros)}", 
                              "Erros de Validação", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return False
            End If
            
            ' Salvar venda no banco
            If Not _database.SalvarVenda(_venda) Then
                MessageBox.Show("Erro ao salvar venda no banco de dados.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End If
            
            ' Atualizar estoque (se implementado)
            AtualizarEstoque()
            
            ' Log da operação
            LogarConfirmacao()
            
            Return True
            
        Catch ex As Exception
            MessageBox.Show($"Erro ao confirmar venda: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' Gera número único para o talão
    ''' </summary>
    Private Function GerarNumeroTalao() As String
        Return $"TAL{Date.Now:yyyyMMddHHmmss}"
    End Function
    
    ''' <summary>
    ''' Atualiza estoque dos produtos vendidos
    ''' </summary>
    Private Sub AtualizarEstoque()
        Try
            ' TODO: Implementar atualização real do estoque
            For Each item In _venda.Itens
                Console.WriteLine($"Estoque atualizado: {item.Produto.Descricao} - Qtd: {item.Quantidade}")
            Next
        Catch ex As Exception
            Console.WriteLine($"Erro ao atualizar estoque: {ex.Message}")
        End Try
    End Sub
    
    ''' <summary>
    ''' Registra log da confirmação
    ''' </summary>
    Private Sub LogarConfirmacao()
        Try
            Dim logMsg = $"Venda confirmada - Talão: {_venda.NumeroTalao}, Cliente: {_venda.Cliente.Nome}, Total: {_calculadora?.CalcularTotalLiquido():C2}"
            Console.WriteLine($"{Date.Now:yyyy-MM-dd HH:mm:ss} - {logMsg}")
        Catch ex As Exception
            Console.WriteLine($"Erro ao registrar log: {ex.Message}")
        End Try
    End Sub
End Class

''' <summary>
''' Formulário de confirmação de pedido
''' Interface para revisar e confirmar vendas antes da finalização
''' </summary>
Public Class FormConfirmacaoPedido
    Inherits Form
    
    Private WithEvents dgvItens As DataGridView
    Private WithEvents lblClienteInfo As Label
    Private WithEvents lblTotalInfo As Label
    Private WithEvents lblFormaPagamento As Label
    Private WithEvents lblVendedor As Label
    Private WithEvents lblDataVenda As Label
    Private WithEvents lblNumeroTalao As Label
    Private WithEvents txtObservacoes As TextBox
    Private WithEvents btnConfirmar As Button
    Private WithEvents btnCancelar As Button
    Private WithEvents btnEditar As Button
    Private WithEvents btnImprimir As Button
    
    Private _venda As Venda
    Private _calculadora As CalculadoraMadeireira
    Private _confirmationManager As OrderConfirmationManager
    
    Public Property VendaConfirmada As Boolean = False
    
    Public Sub New(venda As Venda, calculadora As CalculadoraMadeireira)
        _venda = venda
        _calculadora = calculadora
        _confirmationManager = New OrderConfirmationManager()
        _confirmationManager.ConfigurarVenda(_venda, _calculadora)
        
        InitializeComponent()
        ConfigurarInterface()
        CarregarDados()
    End Sub
    
    Private Sub InitializeComponent()
        Me.Text = "Confirmação de Pedido"
        Me.Size = New Size(900, 700)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.BackColor = Color.WhiteSmoke
        
        ' Painel header
        Dim pnlHeader As New Panel() With {
            .Dock = DockStyle.Top,
            .Height = 120,
            .BackColor = Color.White,
            .Padding = New Padding(20)
        }
        
        Dim lblTitulo As New Label() With {
            .Text = "CONFIRMAÇÃO DO PEDIDO",
            .Font = New Font("Segoe UI", 16, FontStyle.Bold),
            .ForeColor = Color.DarkBlue,
            .Location = New Point(20, 10),
            .AutoSize = True
        }
        
        lblNumeroTalao = New Label() With {
            .Font = New Font("Segoe UI", 12, FontStyle.Bold),
            .ForeColor = Color.DarkGreen,
            .Location = New Point(20, 40),
            .AutoSize = True
        }
        
        lblDataVenda = New Label() With {
            .Font = New Font("Segoe UI", 10),
            .ForeColor = Color.Gray,
            .Location = New Point(20, 65),
            .AutoSize = True
        }
        
        pnlHeader.Controls.AddRange({lblTitulo, lblNumeroTalao, lblDataVenda})
        
        ' Painel de informações do cliente
        Dim pnlCliente As New Panel() With {
            .Dock = DockStyle.Top,
            .Height = 80,
            .BackColor = Color.LightBlue,
            .Padding = New Padding(20, 10, 20, 10)
        }
        
        Dim lblClienteTitulo As New Label() With {
            .Text = "CLIENTE:",
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .Location = New Point(20, 10),
            .AutoSize = True
        }
        
        lblClienteInfo = New Label() With {
            .Font = New Font("Segoe UI", 10),
            .Location = New Point(20, 30),
            .Size = New Size(800, 40)
        }
        
        pnlCliente.Controls.AddRange({lblClienteTitulo, lblClienteInfo})
        
        ' Grid de itens
        dgvItens = New DataGridView() With {
            .Dock = DockStyle.Fill,
            .AllowUserToAddRows = False,
            .AllowUserToDeleteRows = False,
            .ReadOnly = True,
            .MultiSelect = False,
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            .BackgroundColor = Color.White,
            .BorderStyle = BorderStyle.None
        }
        
        ' Painel de totais
        Dim pnlTotais As New Panel() With {
            .Dock = DockStyle.Bottom,
            .Height = 100,
            .BackColor = Color.LightGray,
            .Padding = New Padding(20, 10, 20, 10)
        }
        
        lblTotalInfo = New Label() With {
            .Font = New Font("Segoe UI", 12, FontStyle.Bold),
            .ForeColor = Color.DarkBlue,
            .Location = New Point(20, 10),
            .Size = New Size(400, 80)
        }
        
        lblFormaPagamento = New Label() With {
            .Font = New Font("Segoe UI", 10),
            .Location = New Point(450, 10),
            .AutoSize = True
        }
        
        lblVendedor = New Label() With {
            .Font = New Font("Segoe UI", 10),
            .Location = New Point(450, 35),
            .AutoSize = True
        }
        
        pnlTotais.Controls.AddRange({lblTotalInfo, lblFormaPagamento, lblVendedor})
        
        ' Painel de observações
        Dim pnlObservacoes As New Panel() With {
            .Dock = DockStyle.Bottom,
            .Height = 80,
            .BackColor = Color.White,
            .Padding = New Padding(20, 10, 20, 10)
        }
        
        Dim lblObservacoes As New Label() With {
            .Text = "Observações:",
            .Location = New Point(20, 10),
            .AutoSize = True
        }
        
        txtObservacoes = New TextBox() With {
            .Location = New Point(20, 30),
            .Size = New Size(820, 40),
            .Multiline = True,
            .ScrollBars = ScrollBars.Vertical
        }
        
        pnlObservacoes.Controls.AddRange({lblObservacoes, txtObservacoes})
        
        ' Painel de botões
        Dim pnlBotoes As New Panel() With {
            .Dock = DockStyle.Bottom,
            .Height = 60,
            .BackColor = Color.DarkGray,
            .Padding = New Padding(20, 10, 20, 10)
        }
        
        btnConfirmar = New Button() With {
            .Text = "✅ CONFIRMAR PEDIDO",
            .Size = New Size(150, 40),
            .Location = New Point(680, 10),
            .BackColor = Color.Green,
            .ForeColor = Color.White,
            .Font = New Font("Segoe UI", 11, FontStyle.Bold),
            .FlatStyle = FlatStyle.Flat
        }
        
        btnImprimir = New Button() With {
            .Text = "🖨️ IMPRIMIR",
            .Size = New Size(120, 40),
            .Location = New Point(550, 10),
            .BackColor = Color.Blue,
            .ForeColor = Color.White,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .FlatStyle = FlatStyle.Flat
        }
        
        btnEditar = New Button() With {
            .Text = "✏️ EDITAR",
            .Size = New Size(100, 40),
            .Location = New Point(440, 10),
            .BackColor = Color.Orange,
            .ForeColor = Color.White,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .FlatStyle = FlatStyle.Flat
        }
        
        btnCancelar = New Button() With {
            .Text = "❌ CANCELAR",
            .Size = New Size(120, 40),
            .Location = New Point(20, 10),
            .BackColor = Color.Red,
            .ForeColor = Color.White,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .FlatStyle = FlatStyle.Flat
        }
        
        pnlBotoes.Controls.AddRange({btnConfirmar, btnImprimir, btnEditar, btnCancelar})
        
        Me.Controls.AddRange({pnlHeader, pnlCliente, dgvItens, pnlObservacoes, pnlTotais, pnlBotoes})
    End Sub
    
    Private Sub ConfigurarInterface()
        ' Configurar grid
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
    End Sub
    
    Private Sub CarregarDados()
        Try
            ' Informações do cabeçalho
            lblNumeroTalao.Text = $"Talão Nº: {_venda.NumeroTalao}"
            lblDataVenda.Text = $"Data: {_venda.DataVenda:dd/MM/yyyy HH:mm}"
            
            ' Informações do cliente
            lblClienteInfo.Text = $"{_venda.Cliente.Nome}" & Environment.NewLine &
                                 $"{_venda.Cliente.Endereco} - {_venda.Cliente.Cidade}/{_venda.Cliente.UF} - Tel: {_venda.Cliente.Telefone}"
            
            ' Informações de pagamento
            lblFormaPagamento.Text = $"Forma de Pagamento: {_venda.FormaPagamento}"
            lblVendedor.Text = $"Vendedor: {_venda.VendedorNome}"
            
            ' Totais
            lblTotalInfo.Text = $"Subtotal: {_calculadora.CalcularSubtotalItens():C2}" & Environment.NewLine &
                               $"Desconto: {_calculadora.CalcularDescontoTotal():C2}" & Environment.NewLine &
                               $"Frete: {_calculadora.Frete:C2}" & Environment.NewLine &
                               $"TOTAL GERAL: {_calculadora.CalcularTotalLiquido():C2}"
            
            ' Observações
            txtObservacoes.Text = _venda.Observacoes
            
            ' Carregar itens
            CarregarItens()
            
        Catch ex As Exception
            MessageBox.Show($"Erro ao carregar dados: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    Private Sub CarregarItens()
        dgvItens.Rows.Clear()
        
        For Each item In _venda.Itens
            dgvItens.Rows.Add(
                item.Produto.Codigo,
                item.Produto.Descricao,
                item.Quantidade,
                item.Produto.Unidade,
                item.PrecoUnitario,
                item.Desconto,
                _calculadora.CalcularSubtotalItem(item)
            )
        Next
    End Sub
    
    Private Sub btnConfirmar_Click(sender As Object, e As EventArgs) Handles btnConfirmar.Click
        Try
            ' Atualizar observações
            _venda.Observacoes = txtObservacoes.Text
            
            ' Confirmar venda
            If _confirmationManager.ConfirmarVenda() Then
                VendaConfirmada = True
                MessageBox.Show($"Pedido confirmado com sucesso!" & Environment.NewLine & Environment.NewLine &
                              $"Talão: {_venda.NumeroTalao}" & Environment.NewLine &
                              $"Cliente: {_venda.Cliente.Nome}" & Environment.NewLine &
                              $"Total: {_calculadora.CalcularTotalLiquido():C2}",
                              "Pedido Confirmado", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.DialogResult = DialogResult.OK
                Me.Close()
            End If
            
        Catch ex As Exception
            MessageBox.Show($"Erro ao confirmar pedido: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    Private Sub btnImprimir_Click(sender As Object, e As EventArgs) Handles btnImprimir.Click
        Try
            ' Preparar dados para impressão
            Dim dadosTalao As New DadosTalao() With {
                .NomeCliente = _venda.Cliente.Nome,
                .EnderecoCliente = _venda.Cliente.Endereco,
                .Telefone = _venda.Cliente.Telefone,
                .FormaPagamento = _venda.FormaPagamento,
                .Vendedor = _venda.VendedorNome,
                .NumeroTalao = _venda.NumeroTalao
            }
            
            ' Converter itens
            For Each item In _venda.Itens
                dadosTalao.Produtos.Add(New ProdutoTalao() With {
                    .Descricao = item.Produto.Descricao,
                    .Quantidade = item.Quantidade,
                    .Unidade = item.Produto.Unidade,
                    .PrecoUnitario = item.PrecoUnitario,
                    .PrecoTotal = _calculadora.CalcularSubtotalItem(item)
                })
            Next
            
            ' Gerar e imprimir talão
            Dim excelAutomation As New ExcelAutomation()
            excelAutomation.ProcessarTalaoCompleto(dadosTalao)
            
            MessageBox.Show("Talão impresso com sucesso!", "Impressão", MessageBoxButtons.OK, MessageBoxIcon.Information)
            
        Catch ex As Exception
            MessageBox.Show($"Erro ao imprimir talão: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    Private Sub btnEditar_Click(sender As Object, e As EventArgs) Handles btnEditar.Click
        Me.DialogResult = DialogResult.Retry ' Sinalizar que quer editar
        Me.Close()
    End Sub
    
    Private Sub btnCancelar_Click(sender As Object, e As EventArgs) Handles btnCancelar.Click
        If MessageBox.Show("Deseja cancelar este pedido?", "Confirmar", 
                          MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            Me.DialogResult = DialogResult.Cancel
            Me.Close()
        End If
    End Sub
End Class