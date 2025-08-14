Imports System.Windows.Forms
Imports System.Drawing

''' <summary>
''' Sistema de relatórios integrado para o PDV
''' Gera relatórios de vendas, clientes e produtos com gráficos
''' </summary>
Public Class ReportsManager
    Private _database As DatabaseManager
    
    Public Sub New()
        _database = DatabaseManager.Instance
    End Sub
    
    ''' <summary>
    ''' Gera relatório de vendas por período
    ''' </summary>
    Public Function GerarRelatorioVendas(dataInicio As Date, dataFim As Date) As RelatorioVendas
        Try
            Dim relatorio As New RelatorioVendas() With {
                .DataInicio = dataInicio,
                .DataFim = dataFim,
                .DataGeracao = Date.Now
            }
            
            ' TODO: Buscar vendas reais do banco
            ' Por enquanto, dados de exemplo
            relatorio.TotalVendas = 15420.50
            relatorio.QuantidadeVendas = 45
            relatorio.TicketMedio = relatorio.TotalVendas / relatorio.QuantidadeVendas
            relatorio.VendasPorDia = GerarVendasPorDia(dataInicio, dataFim)
            relatorio.ProdutosMaisVendidos = GerarProdutosMaisVendidos()
            relatorio.VendasPorFormaPagamento = GerarVendasPorFormaPagamento()
            
            Return relatorio
        Catch ex As Exception
            Console.WriteLine($"Erro ao gerar relatório de vendas: {ex.Message}")
            Return Nothing
        End Try
    End Function
    
    ''' <summary>
    ''' Gera relatório de clientes
    ''' </summary>
    Public Function GerarRelatorioClientes() As RelatorioClientes
        Try
            Dim relatorio As New RelatorioClientes() With {
                .DataGeracao = Date.Now
            }
            
            ' TODO: Buscar dados reais do banco
            relatorio.TotalClientes = 127
            relatorio.ClientesAtivos = 98
            relatorio.ClientesInativos = 29
            relatorio.NovosCadastrosUltimos30Dias = 12
            relatorio.ClientesPorCidade = GerarClientesPorCidade()
            relatorio.TopClientesPorCompras = GerarTopClientes()
            
            Return relatorio
        Catch ex As Exception
            Console.WriteLine($"Erro ao gerar relatório de clientes: {ex.Message}")
            Return Nothing
        End Try
    End Function
    
    ''' <summary>
    ''' Gera dados de vendas por dia
    ''' </summary>
    Private Function GerarVendasPorDia(dataInicio As Date, dataFim As Date) As Dictionary(Of Date, Decimal)
        Dim vendas As New Dictionary(Of Date, Decimal)()
        Dim random As New Random()
        
        Dim dataAtual = dataInicio
        While dataAtual <= dataFim
            vendas.Add(dataAtual, random.Next(200, 1500))
            dataAtual = dataAtual.AddDays(1)
        End While
        
        Return vendas
    End Function
    
    ''' <summary>
    ''' Gera dados de produtos mais vendidos
    ''' </summary>
    Private Function GerarProdutosMaisVendidos() As List(Of ProdutoVendido)
        Return New List(Of ProdutoVendido) From {
            New ProdutoVendido() With {.Nome = "Tábua de Pinus 2x4m", .Quantidade = 127, .Valor = 3175.0},
            New ProdutoVendido() With {.Nome = "Ripão 3x3x3m", .Quantidade = 89, .Valor = 1335.0},
            New ProdutoVendido() With {.Nome = "Compensado 18mm", .Quantidade = 56, .Valor = 2520.0},
            New ProdutoVendido() With {.Nome = "Viga de Eucalipto 6x12", .Quantidade = 34, .Valor = 2890.0},
            New ProdutoVendido() With {.Nome = "Caibro 5x7x3m", .Quantidade = 78, .Valor = 1404.0}
        }
    End Function
    
    ''' <summary>
    ''' Gera dados de vendas por forma de pagamento
    ''' </summary>
    Private Function GerarVendasPorFormaPagamento() As Dictionary(Of String, Decimal)
        Return New Dictionary(Of String, Decimal) From {
            {"À Vista", 8540.30},
            {"Cartão Débito", 3210.75},
            {"Cartão Crédito", 2850.45},
            {"PIX", 1920.60},
            {"Fiado", 898.40}
        }
    End Function
    
    ''' <summary>
    ''' Gera dados de clientes por cidade
    ''' </summary>
    Private Function GerarClientesPorCidade() As Dictionary(Of String, Integer)
        Return New Dictionary(Of String, Integer) From {
            {"Paulista", 45},
            {"Recife", 38},
            {"Olinda", 24},
            {"Jaboatão", 15},
            {"Camaragibe", 5}
        }
    End Function
    
    ''' <summary>
    ''' Gera top clientes por compras
    ''' </summary>
    Private Function GerarTopClientes() As List(Of ClienteCompras)
        Return New List(Of ClienteCompras) From {
            New ClienteCompras() With {.Nome = "Construtora ABC Ltda", .TotalCompras = 45800.75, .QuantidadeCompras = 12},
            New ClienteCompras() With {.Nome = "João Silva", .TotalCompras = 8940.30, .QuantidadeCompras = 8},
            New ClienteCompras() With {.Nome = "Maria Oliveira", .TotalCompras = 6750.80, .QuantidadeCompras = 6},
            New ClienteCompras() With {.Nome = "Pedro Santos", .TotalCompras = 4320.45, .QuantidadeCompras = 5},
            New ClienteCompras() With {.Nome = "Ana Costa", .TotalCompras = 3840.90, .QuantidadeCompras = 4}
        }
    End Function
End Class

''' <summary>
''' Estrutura do relatório de vendas
''' </summary>
Public Class RelatorioVendas
    Public Property DataInicio As Date
    Public Property DataFim As Date
    Public Property DataGeracao As Date
    Public Property TotalVendas As Decimal
    Public Property QuantidadeVendas As Integer
    Public Property TicketMedio As Decimal
    Public Property VendasPorDia As Dictionary(Of Date, Decimal)
    Public Property ProdutosMaisVendidos As List(Of ProdutoVendido)
    Public Property VendasPorFormaPagamento As Dictionary(Of String, Decimal)
    
    Public Sub New()
        VendasPorDia = New Dictionary(Of Date, Decimal)()
        ProdutosMaisVendidos = New List(Of ProdutoVendido)()
        VendasPorFormaPagamento = New Dictionary(Of String, Decimal)()
    End Sub
End Class

''' <summary>
''' Estrutura do relatório de clientes
''' </summary>
Public Class RelatorioClientes
    Public Property DataGeracao As Date
    Public Property TotalClientes As Integer
    Public Property ClientesAtivos As Integer
    Public Property ClientesInativos As Integer
    Public Property NovosCadastrosUltimos30Dias As Integer
    Public Property ClientesPorCidade As Dictionary(Of String, Integer)
    Public Property TopClientesPorCompras As List(Of ClienteCompras)
    
    Public Sub New()
        ClientesPorCidade = New Dictionary(Of String, Integer)()
        TopClientesPorCompras = New List(Of ClienteCompras)()
    End Sub
End Class

''' <summary>
''' Estrutura para produto vendido
''' </summary>
Public Class ProdutoVendido
    Public Property Nome As String
    Public Property Quantidade As Integer
    Public Property Valor As Decimal
End Class

''' <summary>
''' Estrutura para cliente com compras
''' </summary>
Public Class ClienteCompras
    Public Property Nome As String
    Public Property TotalCompras As Decimal
    Public Property QuantidadeCompras As Integer
End Class

''' <summary>
''' Formulário de relatórios
''' </summary>
Public Class FormRelatorios
    Inherits Form
    
    Private WithEvents tabControl As TabControl
    Private WithEvents tabVendas As TabPage
    Private WithEvents tabClientes As TabPage
    Private WithEvents tabProdutos As TabPage
    Private WithEvents btnGerarRelatorio As Button
    Private WithEvents btnExportar As Button
    Private WithEvents btnFechar As Button
    Private WithEvents dtpDataInicio As DateTimePicker
    Private WithEvents dtpDataFim As DateTimePicker
    Private WithEvents rtbRelatorio As RichTextBox
    
    Private _reportsManager As ReportsManager
    
    Public Sub New()
        InitializeComponent()
        _reportsManager = New ReportsManager()
        ConfigurarInterface()
    End Sub
    
    Private Sub InitializeComponent()
        Me.Text = "Relatórios do Sistema"
        Me.Size = New Size(900, 700)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.BackColor = Color.WhiteSmoke
        
        ' Tab Control
        tabControl = New TabControl() With {
            .Dock = DockStyle.Fill,
            .Font = New Font("Segoe UI", 10)
        }
        
        ' Aba Vendas
        tabVendas = New TabPage("📊 Vendas") With {
            .BackColor = Color.White,
            .Padding = New Padding(10)
        }
        
        ' Aba Clientes
        tabClientes = New TabPage("👥 Clientes") With {
            .BackColor = Color.White,
            .Padding = New Padding(10)
        }
        
        ' Aba Produtos
        tabProdutos = New TabPage("📦 Produtos") With {
            .BackColor = Color.White,
            .Padding = New Padding(10)
        }
        
        tabControl.TabPages.AddRange({tabVendas, tabClientes, tabProdutos})
        
        ' Painel de controles na aba vendas
        Dim pnlControles As New Panel() With {
            .Dock = DockStyle.Top,
            .Height = 60,
            .BackColor = Color.LightGray,
            .Padding = New Padding(10)
        }
        
        Dim lblDataInicio As New Label() With {
            .Text = "Data Início:",
            .Location = New Point(10, 10),
            .AutoSize = True
        }
        
        dtpDataInicio = New DateTimePicker() With {
            .Location = New Point(10, 30),
            .Size = New Size(120, 25),
            .Value = Date.Today.AddDays(-30)
        }
        
        Dim lblDataFim As New Label() With {
            .Text = "Data Fim:",
            .Location = New Point(140, 10),
            .AutoSize = True
        }
        
        dtpDataFim = New DateTimePicker() With {
            .Location = New Point(140, 30),
            .Size = New Size(120, 25),
            .Value = Date.Today
        }
        
        btnGerarRelatorio = New Button() With {
            .Text = "📊 Gerar",
            .Location = New Point(270, 30),
            .Size = New Size(80, 25),
            .BackColor = Color.Green,
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
        
        btnExportar = New Button() With {
            .Text = "💾 Exportar",
            .Location = New Point(360, 30),
            .Size = New Size(80, 25),
            .BackColor = Color.Blue,
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
        
        pnlControles.Controls.AddRange({lblDataInicio, dtpDataInicio, lblDataFim, dtpDataFim, btnGerarRelatorio, btnExportar})
        
        ' Área de relatório
        rtbRelatorio = New RichTextBox() With {
            .Dock = DockStyle.Fill,
            .Font = New Font("Courier New", 9),
            .ReadOnly = True,
            .BackColor = Color.White
        }
        
        tabVendas.Controls.AddRange({pnlControles, rtbRelatorio})
        
        ' Botão fechar
        Dim pnlBotoes As New Panel() With {
            .Dock = DockStyle.Bottom,
            .Height = 50,
            .BackColor = Color.LightGray,
            .Padding = New Padding(10)
        }
        
        btnFechar = New Button() With {
            .Text = "❌ Fechar",
            .Location = New Point(800, 10),
            .Size = New Size(80, 30),
            .BackColor = Color.Gray,
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
        
        pnlBotoes.Controls.Add(btnFechar)
        
        Me.Controls.AddRange({tabControl, pnlBotoes})
    End Sub
    
    Private Sub ConfigurarInterface()
        ' Configurar abas de clientes e produtos
        ConfigurarAbaClientes()
        ConfigurarAbaProdutos()
    End Sub
    
    Private Sub ConfigurarAbaClientes()
        Dim lblClientes As New Label() With {
            .Text = "📈 RELATÓRIO DE CLIENTES" & Environment.NewLine & Environment.NewLine &
                   "• Total de clientes cadastrados" & Environment.NewLine &
                   "• Clientes ativos vs inativos" & Environment.NewLine &
                   "• Novos cadastros no período" & Environment.NewLine &
                   "• Distribuição por cidade" & Environment.NewLine &
                   "• Top clientes por volume de compras",
            .Font = New Font("Segoe UI", 12),
            .Location = New Point(20, 20),
            .AutoSize = True
        }
        
        Dim btnRelatorioClientes As New Button() With {
            .Text = "📊 Gerar Relatório de Clientes",
            .Location = New Point(20, 200),
            .Size = New Size(200, 40),
            .BackColor = Color.Orange,
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold)
        }
        
        AddHandler btnRelatorioClientes.Click, AddressOf GerarRelatorioClientes
        
        tabClientes.Controls.AddRange({lblClientes, btnRelatorioClientes})
    End Sub
    
    Private Sub ConfigurarAbaProdutos()
        Dim lblProdutos As New Label() With {
            .Text = "📦 RELATÓRIO DE PRODUTOS" & Environment.NewLine & Environment.NewLine &
                   "• Produtos mais vendidos" & Environment.NewLine &
                   "• Análise de estoque" & Environment.NewLine &
                   "• Produtos com baixo giro" & Environment.NewLine &
                   "• Margem de lucro por produto" & Environment.NewLine &
                   "• Relatório de movimentações",
            .Font = New Font("Segoe UI", 12),
            .Location = New Point(20, 20),
            .AutoSize = True
        }
        
        Dim btnRelatorioProdutos As New Button() With {
            .Text = "📊 Em Desenvolvimento",
            .Location = New Point(20, 200),
            .Size = New Size(200, 40),
            .BackColor = Color.Gray,
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .Enabled = False
        }
        
        tabProdutos.Controls.AddRange({lblProdutos, btnRelatorioProdutos})
    End Sub
    
    Private Sub btnGerarRelatorio_Click(sender As Object, e As EventArgs) Handles btnGerarRelatorio.Click
        Try
            Dim relatorio = _reportsManager.GerarRelatorioVendas(dtpDataInicio.Value, dtpDataFim.Value)
            
            If relatorio IsNot Nothing Then
                ExibirRelatorioVendas(relatorio)
            Else
                MessageBox.Show("Erro ao gerar relatório.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            
        Catch ex As Exception
            MessageBox.Show($"Erro ao gerar relatório: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    Private Sub GerarRelatorioClientes(sender As Object, e As EventArgs)
        Try
            Dim relatorio = _reportsManager.GerarRelatorioClientes()
            
            If relatorio IsNot Nothing Then
                ExibirRelatorioClientes(relatorio)
            Else
                MessageBox.Show("Erro ao gerar relatório de clientes.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            
        Catch ex As Exception
            MessageBox.Show($"Erro ao gerar relatório de clientes: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    Private Sub ExibirRelatorioVendas(relatorio As RelatorioVendas)
        rtbRelatorio.Clear()
        
        Dim sb As New System.Text.StringBuilder()
        sb.AppendLine("═══════════════════════════════════════════════════════════")
        sb.AppendLine($"              RELATÓRIO DE VENDAS")
        sb.AppendLine($"      Período: {relatorio.DataInicio:dd/MM/yyyy} a {relatorio.DataFim:dd/MM/yyyy}")
        sb.AppendLine($"      Gerado em: {relatorio.DataGeracao:dd/MM/yyyy HH:mm:ss}")
        sb.AppendLine("═══════════════════════════════════════════════════════════")
        sb.AppendLine()
        
        ' Resumo geral
        sb.AppendLine("📊 RESUMO GERAL:")
        sb.AppendLine($"   Total de Vendas: {relatorio.TotalVendas:C2}")
        sb.AppendLine($"   Quantidade de Vendas: {relatorio.QuantidadeVendas}")
        sb.AppendLine($"   Ticket Médio: {relatorio.TicketMedio:C2}")
        sb.AppendLine()
        
        ' Produtos mais vendidos
        sb.AppendLine("🏆 PRODUTOS MAIS VENDIDOS:")
        For Each produto In relatorio.ProdutosMaisVendidos.Take(5)
            sb.AppendLine($"   {produto.Nome.PadRight(25)} | Qtd: {produto.Quantidade.ToString().PadLeft(4)} | Total: {produto.Valor:C2}")
        Next
        sb.AppendLine()
        
        ' Vendas por forma de pagamento
        sb.AppendLine("💳 VENDAS POR FORMA DE PAGAMENTO:")
        For Each pagamento In relatorio.VendasPorFormaPagamento
            Dim percentual = (pagamento.Value / relatorio.TotalVendas) * 100
            sb.AppendLine($"   {pagamento.Key.PadRight(15)} | {pagamento.Value:C2} ({percentual:F1}%)")
        Next
        sb.AppendLine()
        
        ' Vendas por dia (últimos 7 dias)
        sb.AppendLine("📅 VENDAS DIÁRIAS (Últimos 7 dias):")
        For Each venda In relatorio.VendasPorDia.OrderByDescending(Function(v) v.Key).Take(7)
            sb.AppendLine($"   {venda.Key:dd/MM/yyyy} | {venda.Value:C2}")
        Next
        
        rtbRelatorio.Text = sb.ToString()
    End Sub
    
    Private Sub ExibirRelatorioClientes(relatorio As RelatorioClientes)
        ' Mudar para aba de clientes e exibir relatório
        tabControl.SelectedTab = tabClientes
        
        Dim rtbClientes As New RichTextBox() With {
            .Dock = DockStyle.Bottom,
            .Height = 400,
            .Font = New Font("Courier New", 9),
            .ReadOnly = True,
            .BackColor = Color.White
        }
        
        Dim sb As New System.Text.StringBuilder()
        sb.AppendLine("═══════════════════════════════════════════════════════════")
        sb.AppendLine($"              RELATÓRIO DE CLIENTES")
        sb.AppendLine($"      Gerado em: {relatorio.DataGeracao:dd/MM/yyyy HH:mm:ss}")
        sb.AppendLine("═══════════════════════════════════════════════════════════")
        sb.AppendLine()
        
        sb.AppendLine("📊 RESUMO GERAL:")
        sb.AppendLine($"   Total de Clientes: {relatorio.TotalClientes}")
        sb.AppendLine($"   Clientes Ativos: {relatorio.ClientesAtivos} ({(relatorio.ClientesAtivos / relatorio.TotalClientes * 100):F1}%)")
        sb.AppendLine($"   Clientes Inativos: {relatorio.ClientesInativos} ({(relatorio.ClientesInativos / relatorio.TotalClientes * 100):F1}%)")
        sb.AppendLine($"   Novos Cadastros (30 dias): {relatorio.NovosCadastrosUltimos30Dias}")
        sb.AppendLine()
        
        sb.AppendLine("🏙️ CLIENTES POR CIDADE:")
        For Each cidade In relatorio.ClientesPorCidade.OrderByDescending(Function(c) c.Value)
            Dim percentual = (cidade.Value / relatorio.TotalClientes) * 100
            sb.AppendLine($"   {cidade.Key.PadRight(15)} | {cidade.Value.ToString().PadLeft(3)} clientes ({percentual:F1}%)")
        Next
        sb.AppendLine()
        
        sb.AppendLine("💰 TOP CLIENTES POR COMPRAS:")
        For Each cliente In relatorio.TopClientesPorCompras
            sb.AppendLine($"   {cliente.Nome.PadRight(25)} | {cliente.TotalCompras:C2} ({cliente.QuantidadeCompras} compras)")
        Next
        
        rtbClientes.Text = sb.ToString()
        tabClientes.Controls.Add(rtbClientes)
    End Sub
    
    Private Sub btnExportar_Click(sender As Object, e As EventArgs) Handles btnExportar.Click
        Try
            Using saveDialog As New SaveFileDialog()
                saveDialog.Filter = "Arquivo de Texto (*.txt)|*.txt|Arquivo RTF (*.rtf)|*.rtf"
                saveDialog.FileName = $"Relatorio_Vendas_{Date.Now:yyyyMMdd}.txt"
                
                If saveDialog.ShowDialog() = DialogResult.OK Then
                    System.IO.File.WriteAllText(saveDialog.FileName, rtbRelatorio.Text)
                    MessageBox.Show("Relatório exportado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End Using
        Catch ex As Exception
            MessageBox.Show($"Erro ao exportar relatório: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    Private Sub btnFechar_Click(sender As Object, e As EventArgs) Handles btnFechar.Click
        Me.Close()
    End Sub
End Class