Imports System.Windows.Forms
Imports System.Drawing

''' <summary>
''' Sistema de busca avançada de produtos
''' Interface moderna com filtros e busca dinâmica
''' </summary>
Public Class ProductSearchManager
    Private _produtos As List(Of Produto)
    Private _database As DatabaseManager
    Private _produtosFiltrados As List(Of Produto)
    
    Public Sub New()
        _database = DatabaseManager.Instance
        _produtos = New List(Of Produto)()
        _produtosFiltrados = New List(Of Produto)()
        CarregarProdutos()
    End Sub
    
    ''' <summary>
    ''' Carrega todos os produtos do banco
    ''' </summary>
    Public Sub CarregarProdutos()
        Try
            _produtos = _database.ObterTodosProdutos()
            _produtosFiltrados = New List(Of Produto)(_produtos)
        Catch ex As Exception
            MessageBox.Show($"Erro ao carregar produtos: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    ''' <summary>
    ''' Busca produtos por termo
    ''' </summary>
    Public Function BuscarProdutos(termo As String, Optional secao As String = "") As List(Of Produto)
        Try
            If String.IsNullOrEmpty(termo) AndAlso String.IsNullOrEmpty(secao) Then
                _produtosFiltrados = New List(Of Produto)(_produtos)
            Else
                _produtosFiltrados = _produtos.Where(Function(p) 
                    (String.IsNullOrEmpty(termo) OrElse 
                     p.Descricao.ToUpper().Contains(termo.ToUpper()) OrElse
                     p.Codigo.ToUpper().Contains(termo.ToUpper())) AndAlso
                    (String.IsNullOrEmpty(secao) OrElse 
                     p.Secao.Equals(secao, StringComparison.OrdinalIgnoreCase))
                ).ToList()
            End If
            
            Return _produtosFiltrados
        Catch ex As Exception
            Console.WriteLine($"Erro na busca de produtos: {ex.Message}")
            Return New List(Of Produto)()
        End Try
    End Function
    
    ''' <summary>
    ''' Obtém sugestões de busca
    ''' </summary>
    Public Function ObterSugestoes(termo As String) As List(Of String)
        Try
            If String.IsNullOrEmpty(termo) OrElse termo.Length < 2 Then
                Return New List(Of String)()
            End If
            
            Dim sugestoes = _produtos.Where(Function(p) 
                p.Descricao.ToUpper().Contains(termo.ToUpper()) OrElse
                p.Codigo.ToUpper().Contains(termo.ToUpper())
            ).Select(Function(p) p.Descricao).Distinct().Take(10).ToList()
            
            Return sugestoes
        Catch ex As Exception
            Console.WriteLine($"Erro ao obter sugestões: {ex.Message}")
            Return New List(Of String)()
        End Try
    End Function
    
    ''' <summary>
    ''' Obtém todas as seções disponíveis
    ''' </summary>
    Public Function ObterSecoes() As List(Of String)
        Try
            Return _database.ObterSecoesProdutos()
        Catch ex As Exception
            Console.WriteLine($"Erro ao obter seções: {ex.Message}")
            Return New List(Of String)()
        End Try
    End Function
    
    ''' <summary>
    ''' Filtra produtos por preço
    ''' </summary>
    Public Function FiltrarPorPreco(precoMin As Decimal, precoMax As Decimal) As List(Of Produto)
        Try
            _produtosFiltrados = _produtosFiltrados.Where(Function(p) 
                p.PrecoVenda >= precoMin AndAlso p.PrecoVenda <= precoMax
            ).ToList()
            
            Return _produtosFiltrados
        Catch ex As Exception
            Console.WriteLine($"Erro ao filtrar por preço: {ex.Message}")
            Return New List(Of Produto)()
        End Try
    End Function
    
    ''' <summary>
    ''' Filtra produtos por estoque
    ''' </summary>
    Public Function FiltrarPorEstoque(estoqueMinimo As Double) As List(Of Produto)
        Try
            _produtosFiltrados = _produtosFiltrados.Where(Function(p) 
                p.EstoqueAtual >= estoqueMinimo
            ).ToList()
            
            Return _produtosFiltrados
        Catch ex As Exception
            Console.WriteLine($"Erro ao filtrar por estoque: {ex.Message}")
            Return New List(Of Produto)()
        End Try
    End Function
    
    ''' <summary>
    ''' Ordena produtos
    ''' </summary>
    Public Function OrdenarProdutos(criterio As String, crescente As Boolean) As List(Of Produto)
        Try
            Select Case criterio.ToUpper()
                Case "CODIGO"
                    If crescente Then
                        _produtosFiltrados = _produtosFiltrados.OrderBy(Function(p) p.Codigo).ToList()
                    Else
                        _produtosFiltrados = _produtosFiltrados.OrderByDescending(Function(p) p.Codigo).ToList()
                    End If
                Case "DESCRICAO"
                    If crescente Then
                        _produtosFiltrados = _produtosFiltrados.OrderBy(Function(p) p.Descricao).ToList()
                    Else
                        _produtosFiltrados = _produtosFiltrados.OrderByDescending(Function(p) p.Descricao).ToList()
                    End If
                Case "PRECO"
                    If crescente Then
                        _produtosFiltrados = _produtosFiltrados.OrderBy(Function(p) p.PrecoVenda).ToList()
                    Else
                        _produtosFiltrados = _produtosFiltrados.OrderByDescending(Function(p) p.PrecoVenda).ToList()
                    End If
                Case "ESTOQUE"
                    If crescente Then
                        _produtosFiltrados = _produtosFiltrados.OrderBy(Function(p) p.EstoqueAtual).ToList()
                    Else
                        _produtosFiltrados = _produtosFiltrados.OrderByDescending(Function(p) p.EstoqueAtual).ToList()
                    End If
                Case Else
                    ' Ordem padrão por descrição
                    _produtosFiltrados = _produtosFiltrados.OrderBy(Function(p) p.Descricao).ToList()
            End Select
            
            Return _produtosFiltrados
        Catch ex As Exception
            Console.WriteLine($"Erro ao ordenar produtos: {ex.Message}")
            Return _produtosFiltrados
        End Try
    End Function
    
    ''' <summary>
    ''' Limpa todos os filtros
    ''' </summary>
    Public Sub LimparFiltros()
        _produtosFiltrados = New List(Of Produto)(_produtos)
    End Sub
    
    ''' <summary>
    ''' Propriedade para acessar produtos filtrados
    ''' </summary>
    Public ReadOnly Property ProdutosFiltrados As List(Of Produto)
        Get
            Return _produtosFiltrados
        End Get
    End Property
    
    ''' <summary>
    ''' Propriedade para acessar todos os produtos
    ''' </summary>
    Public ReadOnly Property TodosProdutos As List(Of Produto)
        Get
            Return _produtos
        End Get
    End Property
End Class

''' <summary>
''' Formulário de busca de produtos com interface moderna
''' </summary>
Public Class FormBuscaProdutos
    Inherits Form
    
    Private WithEvents txtBusca As TextBox
    Private WithEvents cmbSecao As ComboBox
    Private WithEvents dgvProdutos As DataGridView
    Private WithEvents btnBuscar As Button
    Private WithEvents btnLimpar As Button
    Private WithEvents btnSelecionar As Button
    Private WithEvents btnCancelar As Button
    
    Private _searchManager As ProductSearchManager
    Private _produtoSelecionado As Produto
    
    Public Property ProdutoSelecionado As Produto
        Get
            Return _produtoSelecionado
        End Get
        Set(value As Produto)
            _produtoSelecionado = value
        End Set
    End Property
    
    Public Sub New()
        InitializeComponent()
        _searchManager = New ProductSearchManager()
        ConfigurarInterface()
        CarregarDados()
    End Sub
    
    Private Sub InitializeComponent()
        Me.Text = "Busca de Produtos"
        Me.Size = New Size(800, 600)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.BackColor = Color.WhiteSmoke
        
        ' Painel de busca
        Dim pnlBusca As New Panel() With {
            .Dock = DockStyle.Top,
            .Height = 80,
            .BackColor = Color.White,
            .Padding = New Padding(10)
        }
        
        ' Campo de busca
        Dim lblBusca As New Label() With {
            .Text = "Buscar produto:",
            .Location = New Point(10, 10),
            .AutoSize = True
        }
        
        txtBusca = New TextBox() With {
            .Location = New Point(10, 30),
            .Size = New Size(300, 25),
            .Font = New Font("Segoe UI", 10)
        }
        
        ' Combo seção
        Dim lblSecao As New Label() With {
            .Text = "Seção:",
            .Location = New Point(320, 10),
            .AutoSize = True
        }
        
        cmbSecao = New ComboBox() With {
            .Location = New Point(320, 30),
            .Size = New Size(150, 25),
            .DropDownStyle = ComboBoxStyle.DropDownList
        }
        
        ' Botões
        btnBuscar = New Button() With {
            .Text = "🔍 Buscar",
            .Location = New Point(480, 30),
            .Size = New Size(80, 25),
            .BackColor = Color.DodgerBlue,
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
        
        btnLimpar = New Button() With {
            .Text = "🗑️ Limpar",
            .Location = New Point(570, 30),
            .Size = New Size(80, 25),
            .BackColor = Color.Gray,
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
        
        pnlBusca.Controls.AddRange({lblBusca, txtBusca, lblSecao, cmbSecao, btnBuscar, btnLimpar})
        
        ' Grid de produtos
        dgvProdutos = New DataGridView() With {
            .Dock = DockStyle.Fill,
            .AllowUserToAddRows = False,
            .AllowUserToDeleteRows = False,
            .ReadOnly = True,
            .MultiSelect = False,
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            .BackgroundColor = Color.White,
            .BorderStyle = BorderStyle.None
        }
        
        ' Painel de botões
        Dim pnlBotoes As New Panel() With {
            .Dock = DockStyle.Bottom,
            .Height = 50,
            .BackColor = Color.White,
            .Padding = New Padding(10)
        }
        
        btnSelecionar = New Button() With {
            .Text = "✅ Selecionar",
            .Location = New Point(600, 10),
            .Size = New Size(100, 30),
            .BackColor = Color.Green,
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
        
        btnCancelar = New Button() With {
            .Text = "❌ Cancelar",
            .Location = New Point(490, 10),
            .Size = New Size(100, 30),
            .BackColor = Color.Gray,
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
        
        pnlBotoes.Controls.AddRange({btnSelecionar, btnCancelar})
        
        Me.Controls.AddRange({pnlBusca, dgvProdutos, pnlBotoes})
    End Sub
    
    Private Sub ConfigurarInterface()
        ' Configurar grid
        dgvProdutos.Columns.Add("Codigo", "Código")
        dgvProdutos.Columns.Add("Descricao", "Descrição")
        dgvProdutos.Columns.Add("Secao", "Seção")
        dgvProdutos.Columns.Add("Unidade", "Unidade")
        dgvProdutos.Columns.Add("PrecoVenda", "Preço de Venda")
        dgvProdutos.Columns.Add("EstoqueAtual", "Estoque")
        
        ' Configurar larguras
        dgvProdutos.Columns("Codigo").Width = 80
        dgvProdutos.Columns("Descricao").Width = 250
        dgvProdutos.Columns("Secao").Width = 100
        dgvProdutos.Columns("Unidade").Width = 70
        dgvProdutos.Columns("PrecoVenda").Width = 100
        dgvProdutos.Columns("EstoqueAtual").Width = 80
        
        ' Formatação
        dgvProdutos.Columns("PrecoVenda").DefaultCellStyle.Format = "C2"
        dgvProdutos.Columns("EstoqueAtual").DefaultCellStyle.Format = "N2"
    End Sub
    
    Private Sub CarregarDados()
        ' Carregar seções
        cmbSecao.Items.Clear()
        cmbSecao.Items.Add("Todas")
        For Each secao In _searchManager.ObterSecoes()
            cmbSecao.Items.Add(secao)
        Next
        cmbSecao.SelectedIndex = 0
        
        ' Carregar produtos
        AtualizarGrid(_searchManager.TodosProdutos)
    End Sub
    
    Private Sub AtualizarGrid(produtos As List(Of Produto))
        dgvProdutos.Rows.Clear()
        
        For Each produto In produtos
            dgvProdutos.Rows.Add(
                produto.Codigo,
                produto.Descricao,
                produto.Secao,
                produto.Unidade,
                produto.PrecoVenda,
                produto.EstoqueAtual
            )
            
            ' Armazenar objeto produto na tag da linha
            dgvProdutos.Rows(dgvProdutos.Rows.Count - 1).Tag = produto
        Next
    End Sub
    
    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click
        Dim termo = txtBusca.Text.Trim()
        Dim secao = If(cmbSecao.SelectedIndex = 0, "", cmbSecao.Text)
        
        Dim produtos = _searchManager.BuscarProdutos(termo, secao)
        AtualizarGrid(produtos)
    End Sub
    
    Private Sub btnLimpar_Click(sender As Object, e As EventArgs) Handles btnLimpar.Click
        txtBusca.Clear()
        cmbSecao.SelectedIndex = 0
        _searchManager.LimparFiltros()
        AtualizarGrid(_searchManager.TodosProdutos)
    End Sub
    
    Private Sub btnSelecionar_Click(sender As Object, e As EventArgs) Handles btnSelecionar.Click
        If dgvProdutos.CurrentRow IsNot Nothing Then
            _produtoSelecionado = CType(dgvProdutos.CurrentRow.Tag, Produto)
            Me.DialogResult = DialogResult.OK
            Me.Close()
        Else
            MessageBox.Show("Selecione um produto.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub
    
    Private Sub btnCancelar_Click(sender As Object, e As EventArgs) Handles btnCancelar.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub
    
    Private Sub txtBusca_TextChanged(sender As Object, e As EventArgs) Handles txtBusca.TextChanged
        ' Busca em tempo real
        If txtBusca.Text.Length >= 2 Then
            Dim timer As New Timer() With {.Interval = 500}
            AddHandler timer.Tick, Sub(s, ev)
                                       timer.Stop()
                                       btnBuscar_Click(sender, e)
                                   End Sub
            timer.Start()
        End If
    End Sub
    
    Private Sub dgvProdutos_DoubleClick(sender As Object, e As EventArgs) Handles dgvProdutos.DoubleClick
        btnSelecionar_Click(sender, e)
    End Sub
End Class