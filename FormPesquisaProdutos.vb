Imports System.Windows.Forms
Imports System.Drawing
Imports Microsoft.Office.Interop.Excel

''' <summary>
''' Formulário de pesquisa integrada de produtos
''' Interface compacta para busca em tempo real na planilha Excel
''' </summary>
Public Class FormPesquisaProdutos
    Inherits Form

    ' Controles da interface
    Private WithEvents txtPesquisa As TextBox
    Private WithEvents dgvResultados As DataGridView
    Private WithEvents lblStatus As Label
    Private WithEvents btnSelecionar As Button
    Private WithEvents btnCancelar As Button

    ' Dados e propriedades
    Public Property ProdutoSelecionado As ProdutoTalao
    Private produtosEncontrados As List(Of ProdutoTalao)
    Private excelApp As Application
    Private planilhaProdutos As Worksheet

    ''' <summary>
    ''' Construtor do formulário de pesquisa
    ''' </summary>
    Public Sub New()
        InitializeComponent()
        ConfigurarInterface()
        produtosEncontrados = New List(Of ProdutoTalao)()
        CarregarProdutosDaPlanilha()
    End Sub

    ''' <summary>
    ''' Inicializa os componentes da interface
    ''' </summary>
    Private Sub InitializeComponent()
        ' Configurações do formulário
        Me.Text = "🔍 Pesquisa de Produtos"
        Me.Size = New Size(700, 500)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.BackColor = Color.WhiteSmoke
        Me.Font = New Font("Segoe UI", 9.0F, FontStyle.Regular)

        ' Campo de pesquisa
        Dim lblPesquisa As New Label()
        lblPesquisa.Text = "Digite para pesquisar:"
        lblPesquisa.Location = New Point(20, 20)
        lblPesquisa.Size = New Size(150, 20)
        lblPesquisa.Font = New Font("Segoe UI", 10.0F, FontStyle.Bold)
        Me.Controls.Add(lblPesquisa)

        txtPesquisa = New TextBox()
        txtPesquisa.Location = New Point(20, 45)
        txtPesquisa.Size = New Size(650, 25)
        txtPesquisa.Font = New Font("Segoe UI", 12.0F)
        txtPesquisa.PlaceholderText = "🔍 Digite código, nome, material ou qualquer parte da descrição..."
        Me.Controls.Add(txtPesquisa)

        ' Status da pesquisa
        lblStatus = New Label()
        lblStatus.Text = "💡 Carregando produtos da planilha..."
        lblStatus.Location = New Point(20, 80)
        lblStatus.Size = New Size(650, 20)
        lblStatus.Font = New Font("Segoe UI", 9.0F, FontStyle.Italic)
        lblStatus.ForeColor = Color.FromArgb(52, 73, 94)
        Me.Controls.Add(lblStatus)

        ' DataGridView para resultados
        dgvResultados = New DataGridView()
        dgvResultados.Location = New Point(20, 110)
        dgvResultados.Size = New Size(650, 280)
        dgvResultados.AllowUserToAddRows = False
        dgvResultados.AllowUserToDeleteRows = False
        dgvResultados.ReadOnly = True
        dgvResultados.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvResultados.MultiSelect = False
        dgvResultados.BackgroundColor = Color.White
        dgvResultados.BorderStyle = BorderStyle.Fixed3D
        Me.Controls.Add(dgvResultados)

        ' Configurar colunas do DataGridView
        dgvResultados.Columns.Add("Codigo", "Código")
        dgvResultados.Columns.Add("Descricao", "Descrição")
        dgvResultados.Columns.Add("Material", "Material")
        dgvResultados.Columns.Add("Unidade", "Un.")
        dgvResultados.Columns.Add("PrecoReal", "Preço Real")
        dgvResultados.Columns.Add("PrecoVisual", "Preço Visual (x1000)")

        dgvResultados.Columns(0).Width = 80  ' Código
        dgvResultados.Columns(1).Width = 250 ' Descrição
        dgvResultados.Columns(2).Width = 100 ' Material
        dgvResultados.Columns(3).Width = 50  ' Unidade
        dgvResultados.Columns(4).Width = 85  ' Preço Real
        dgvResultados.Columns(5).Width = 85  ' Preço Visual

        ' Botões
        btnSelecionar = New Button()
        btnSelecionar.Text = "✅ Selecionar Produto"
        btnSelecionar.Location = New Point(420, 410)
        btnSelecionar.Size = New Size(150, 35)
        btnSelecionar.BackColor = Color.FromArgb(46, 204, 113)
        btnSelecionar.ForeColor = Color.White
        btnSelecionar.Font = New Font("Segoe UI", 10.0F, FontStyle.Bold)
        btnSelecionar.FlatStyle = FlatStyle.Flat
        btnSelecionar.FlatAppearance.BorderSize = 0
        btnSelecionar.Enabled = False
        Me.Controls.Add(btnSelecionar)

        btnCancelar = New Button()
        btnCancelar.Text = "❌ Cancelar"
        btnCancelar.Location = New Point(580, 410)
        btnCancelar.Size = New Size(90, 35)
        btnCancelar.BackColor = Color.FromArgb(231, 76, 60)
        btnCancelar.ForeColor = Color.White
        btnCancelar.Font = New Font("Segoe UI", 10.0F, FontStyle.Bold)
        btnCancelar.FlatStyle = FlatStyle.Flat
        btnCancelar.FlatAppearance.BorderSize = 0
        Me.Controls.Add(btnCancelar)
    End Sub

    ''' <summary>
    ''' Configura detalhes adicionais da interface
    ''' </summary>
    Private Sub ConfigurarInterface()
        ' Configurar eventos de pesquisa em tempo real
        AddHandler txtPesquisa.TextChanged, AddressOf PesquisarEmTempoReal
        AddHandler dgvResultados.SelectionChanged, AddressOf AtualizarBotaoSelecionar
        AddHandler dgvResultados.DoubleClick, AddressOf SelecionarProdutoComDuploClick
    End Sub

    ''' <summary>
    ''' Carrega produtos da planilha Excel
    ''' </summary>
    Private Sub CarregarProdutosDaPlanilha()
        Try
            lblStatus.Text = "📋 Conectando com planilha de produtos..."
            Application.DoEvents()

            ' Tentar abrir planilha de produtos existente
            AbrirPlanilhaProdutos()

            If planilhaProdutos Is Nothing Then
                ' Se não existe, criar planilha de exemplo
                CriarPlanilhaExemplo()
            End If

            ' Carregar produtos da planilha
            LerProdutosDaPlanilha()

            lblStatus.Text = $"✅ {produtosEncontrados.Count} produtos carregados. Digite para pesquisar..."

        Catch ex As Exception
            lblStatus.Text = "❌ Erro ao carregar planilha. Usando dados de exemplo."
            CriarProdutosExemplo()
        End Try
    End Sub

    ''' <summary>
    ''' Abre planilha de produtos existente
    ''' </summary>
    Private Sub AbrirPlanilhaProdutos()
        Try
            excelApp = New Application()
            excelApp.Visible = False
            excelApp.DisplayAlerts = False

            ' Tentar abrir planilha de produtos
            Dim caminhoArquivo As String = System.IO.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "Produtos_Madeireira.xlsx")

            If System.IO.File.Exists(caminhoArquivo) Then
                Dim workbook = excelApp.Workbooks.Open(caminhoArquivo)
                planilhaProdutos = workbook.Worksheets(1)
            End If

        Catch ex As Exception
            ' Se falhar, planilhaProdutos ficará Nothing
            Console.WriteLine("Não foi possível abrir planilha existente: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Cria planilha de exemplo com produtos típicos de madeireira
    ''' </summary>
    Private Sub CriarPlanilhaExemplo()
        Try
            If excelApp Is Nothing Then
                excelApp = New Application()
                excelApp.Visible = False
                excelApp.DisplayAlerts = False
            End If

            Dim workbook = excelApp.Workbooks.Add()
            planilhaProdutos = workbook.ActiveSheet
            planilhaProdutos.Name = "Produtos"

            ' Cabeçalho
            planilhaProdutos.Cells(1, 1).Value = "Código"
            planilhaProdutos.Cells(1, 2).Value = "Descrição"
            planilhaProdutos.Cells(1, 3).Value = "Material"
            planilhaProdutos.Cells(1, 4).Value = "Unidade"
            planilhaProdutos.Cells(1, 5).Value = "Preço"

            ' Produtos de exemplo
            Dim produtosExemplo = {
                {"MAD001", "Tábua de Pinus 2x4x3m", "Pinus", "UN", 25.00},
                {"MAD002", "Ripão 3x3x3m", "Eucalipto", "UN", 15.00},
                {"MAD003", "Compensado 18mm", "Compensado", "M²", 45.00},
                {"MAD004", "Viga 6x12x4m", "Peroba", "UN", 120.00},
                {"MAD005", "Caibro 5x7x3m", "Pinus", "UN", 18.50},
                {"MAD006", "Tábua Aparelhada 2x20x3m", "Cedrinho", "UN", 35.00},
                {"MAD007", "Sarrafo 2x3x3m", "Pinus", "UN", 8.50},
                {"MAD008", "Laminado 10mm", "Compensado", "M²", 28.00},
                {"MAD009", "Pontalete 7x7x3m", "Eucalipto", "UN", 22.00},
                {"MAD010", "Deck de Madeira", "Cumaru", "M²", 85.00}
            }

            For i = 0 To produtosExemplo.Length - 1
                planilhaProdutos.Cells(i + 2, 1).Value = produtosExemplo(i)(0)
                planilhaProdutos.Cells(i + 2, 2).Value = produtosExemplo(i)(1)
                planilhaProdutos.Cells(i + 2, 3).Value = produtosExemplo(i)(2)
                planilhaProdutos.Cells(i + 2, 4).Value = produtosExemplo(i)(3)
                planilhaProdutos.Cells(i + 2, 5).Value = produtosExemplo(i)(4)
            Next

            ' Salvar planilha
            Dim caminhoSalvar As String = System.IO.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "Produtos_Madeireira.xlsx")
            workbook.SaveAs(caminhoSalvar)

        Catch ex As Exception
            Console.WriteLine("Erro ao criar planilha de exemplo: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Lê produtos da planilha Excel
    ''' </summary>
    Private Sub LerProdutosDaPlanilha()
        If planilhaProdutos Is Nothing Then Return

        produtosEncontrados.Clear()

        Try
            Dim linha As Integer = 2 ' Começar da linha 2 (após cabeçalho)
            
            While Not String.IsNullOrEmpty(planilhaProdutos.Cells(linha, 1).Value?.ToString())
                Dim produto As New ProdutoTalao()
                produto.Codigo = planilhaProdutos.Cells(linha, 1).Value?.ToString()
                produto.Descricao = planilhaProdutos.Cells(linha, 2).Value?.ToString()
                produto.Material = planilhaProdutos.Cells(linha, 3).Value?.ToString()
                produto.Unidade = planilhaProdutos.Cells(linha, 4).Value?.ToString()
                
                ' Preço real
                Dim precoReal As Double
                If Double.TryParse(planilhaProdutos.Cells(linha, 5).Value?.ToString(), precoReal) Then
                    produto.PrecoUnitario = precoReal
                    produto.PrecoVisual = precoReal * 1000 ' Multiplicador visual
                End If

                produtosEncontrados.Add(produto)
                linha += 1
            End While

        Catch ex As Exception
            Console.WriteLine("Erro ao ler produtos: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Cria produtos de exemplo quando não há planilha
    ''' </summary>
    Private Sub CriarProdutosExemplo()
        produtosEncontrados.Clear()

        Dim exemplos = {
            New ProdutoTalao() With {.Codigo = "MAD001", .Descricao = "Tábua de Pinus 2x4x3m", .Material = "Pinus", .Unidade = "UN", .PrecoUnitario = 25.0, .PrecoVisual = 25000},
            New ProdutoTalao() With {.Codigo = "MAD002", .Descricao = "Ripão 3x3x3m", .Material = "Eucalipto", .Unidade = "UN", .PrecoUnitario = 15.0, .PrecoVisual = 15000},
            New ProdutoTalao() With {.Codigo = "MAD003", .Descricao = "Compensado 18mm", .Material = "Compensado", .Unidade = "M²", .PrecoUnitario = 45.0, .PrecoVisual = 45000},
            New ProdutoTalao() With {.Codigo = "MAD004", .Descricao = "Viga 6x12x4m", .Material = "Peroba", .Unidade = "UN", .PrecoUnitario = 120.0, .PrecoVisual = 120000},
            New ProdutoTalao() With {.Codigo = "MAD005", .Descricao = "Caibro 5x7x3m", .Material = "Pinus", .Unidade = "UN", .PrecoUnitario = 18.5, .PrecoVisual = 18500}
        }

        produtosEncontrados.AddRange(exemplos)
    End Sub

    ''' <summary>
    ''' Pesquisa em tempo real conforme o usuário digita
    ''' </summary>
    Private Sub PesquisarEmTempoReal(sender As Object, e As EventArgs)
        Dim termo = txtPesquisa.Text.ToLower().Trim()

        If String.IsNullOrEmpty(termo) Then
            ' Mostrar todos os produtos
            PreencherGrid(produtosEncontrados)
            lblStatus.Text = $"📋 {produtosEncontrados.Count} produtos disponíveis"
        Else
            ' Filtrar produtos
            Dim resultados = produtosEncontrados.Where(Function(p)
                p.Codigo.ToLower().Contains(termo) OrElse
                p.Descricao.ToLower().Contains(termo) OrElse
                p.Material.ToLower().Contains(termo)
            ).ToList()

            PreencherGrid(resultados)
            lblStatus.Text = $"🔍 {resultados.Count} produtos encontrados para '{termo}'"
        End If
    End Sub

    ''' <summary>
    ''' Preenche o grid com os produtos
    ''' </summary>
    Private Sub PreencherGrid(produtos As List(Of ProdutoTalao))
        dgvResultados.Rows.Clear()

        For Each produto In produtos
            dgvResultados.Rows.Add(
                produto.Codigo,
                produto.Descricao,
                produto.Material,
                produto.Unidade,
                produto.PrecoUnitario.ToString("C2"),
                produto.PrecoVisual.ToString("N0")
            )
        Next
    End Sub

    ''' <summary>
    ''' Atualiza estado do botão selecionar
    ''' </summary>
    Private Sub AtualizarBotaoSelecionar(sender As Object, e As EventArgs)
        btnSelecionar.Enabled = dgvResultados.SelectedRows.Count > 0
    End Sub

    ''' <summary>
    ''' Seleciona produto com duplo clique
    ''' </summary>
    Private Sub SelecionarProdutoComDuploClick(sender As Object, e As EventArgs)
        If dgvResultados.SelectedRows.Count > 0 Then
            SelecionarProduto()
        End If
    End Sub

    ''' <summary>
    ''' Seleciona o produto e fecha o formulário
    ''' </summary>
    Private Sub btnSelecionar_Click(sender As Object, e As EventArgs) Handles btnSelecionar.Click
        SelecionarProduto()
    End Sub

    ''' <summary>
    ''' Processa a seleção do produto
    ''' </summary>
    Private Sub SelecionarProduto()
        If dgvResultados.SelectedRows.Count = 0 Then Return

        Dim row = dgvResultados.SelectedRows(0)
        ProdutoSelecionado = New ProdutoTalao() With {
            .Codigo = row.Cells("Codigo").Value.ToString(),
            .Descricao = row.Cells("Descricao").Value.ToString(),
            .Material = row.Cells("Material").Value.ToString(),
            .Unidade = row.Cells("Unidade").Value.ToString(),
            .PrecoUnitario = Double.Parse(row.Cells("PrecoReal").Value.ToString().Replace("R$", "").Replace(",", ".").Trim()),
            .PrecoVisual = Double.Parse(row.Cells("PrecoVisual").Value.ToString().Replace(".", ""))
        }

        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    ''' <summary>
    ''' Cancela a seleção
    ''' </summary>
    Private Sub btnCancelar_Click(sender As Object, e As EventArgs) Handles btnCancelar.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub

    ''' <summary>
    ''' Limpeza ao fechar o formulário
    ''' </summary>
    Protected Overrides Sub OnFormClosed(e As FormClosedEventArgs)
        Try
            If excelApp IsNot Nothing Then
                excelApp.Quit()
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp)
                excelApp = Nothing
            End If
        Catch
            ' Ignorar erros de limpeza
        Finally
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try

        MyBase.OnFormClosed(e)
    End Sub
End Class

''' <summary>
''' Extensão da classe ProdutoTalao para incluir campos de pesquisa
''' </summary>
Partial Public Class ProdutoTalao
    Public Property Codigo As String = ""
    Public Property Material As String = ""
    Public Property PrecoVisual As Double = 0
End Class