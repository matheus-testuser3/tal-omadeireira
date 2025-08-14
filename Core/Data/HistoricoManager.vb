Imports System.IO
Imports System.Xml.Serialization

''' <summary>
''' Gerenciador de histórico de vendas
''' Especializado em consultas e relatórios históricos
''' </summary>
Public Class HistoricoManager
    Private Shared ReadOnly _instance As New Lazy(Of HistoricoManager)(Function() New HistoricoManager())
    Private ReadOnly _logger As Logger
    Private ReadOnly _config As ConfigManager
    Private ReadOnly _dataManager As DataManager
    
    ''' <summary>
    ''' Instância singleton do HistoricoManager
    ''' </summary>
    Public Shared ReadOnly Property Instance As HistoricoManager
        Get
            Return _instance.Value
        End Get
    End Property
    
    ''' <summary>
    ''' Construtor privado
    ''' </summary>
    Private Sub New()
        _logger = Logger.Instance
        _config = ConfigManager.Instance
        _dataManager = DataManager.Instance
    End Sub
    
    ''' <summary>
    ''' Obtém vendas do dia atual
    ''' </summary>
    Public Function ObterVendasHoje() As List(Of Venda)
        Dim hoje = DateTime.Today
        Return _dataManager.ObterVendasPorPeriodo(hoje, hoje.AddDays(1).AddSeconds(-1))
    End Function
    
    ''' <summary>
    ''' Obtém vendas da semana atual
    ''' </summary>
    Public Function ObterVendasSemana() As List(Of Venda)
        Dim hoje = DateTime.Today
        Dim inicioSemana = hoje.AddDays(-(hoje.DayOfWeek - DayOfWeek.Monday))
        Dim fimSemana = inicioSemana.AddDays(6).AddHours(23).AddMinutes(59).AddSeconds(59)
        
        Return _dataManager.ObterVendasPorPeriodo(inicioSemana, fimSemana)
    End Function
    
    ''' <summary>
    ''' Obtém vendas do mês atual
    ''' </summary>
    Public Function ObterVendasMes() As List(Of Venda)
        Dim hoje = DateTime.Today
        Dim inicioMes = New DateTime(hoje.Year, hoje.Month, 1)
        Dim fimMes = inicioMes.AddMonths(1).AddSeconds(-1)
        
        Return _dataManager.ObterVendasPorPeriodo(inicioMes, fimMes)
    End Function
    
    ''' <summary>
    ''' Obtém vendas do ano atual
    ''' </summary>
    Public Function ObterVendasAno() As List(Of Venda)
        Dim hoje = DateTime.Today
        Dim inicioAno = New DateTime(hoje.Year, 1, 1)
        Dim fimAno = inicioAno.AddYears(1).AddSeconds(-1)
        
        Return _dataManager.ObterVendasPorPeriodo(inicioAno, fimAno)
    End Function
    
    ''' <summary>
    ''' Busca vendas por critério de pesquisa
    ''' </summary>
    Public Function BuscarVendas(criterio As CriterioBusca) As List(Of Venda)
        Try
            Dim vendas = _dataManager.ObterVendasPorPeriodo(criterio.DataInicio, criterio.DataFim)
            
            ' Filtrar por cliente se especificado
            If Not String.IsNullOrWhiteSpace(criterio.NomeCliente) Then
                vendas = vendas.Where(Function(v) v.Cliente?.Nome.ToLower().Contains(criterio.NomeCliente.ToLower())).ToList()
            End If
            
            ' Filtrar por vendedor se especificado
            If Not String.IsNullOrWhiteSpace(criterio.Vendedor) Then
                vendas = vendas.Where(Function(v) v.Vendedor.ToLower().Contains(criterio.Vendedor.ToLower())).ToList()
            End If
            
            ' Filtrar por número do talão se especificado
            If Not String.IsNullOrWhiteSpace(criterio.NumeroTalao) Then
                vendas = vendas.Where(Function(v) v.NumeroTalao.Contains(criterio.NumeroTalao)).ToList()
            End If
            
            ' Filtrar por valor mínimo se especificado
            If criterio.ValorMinimo.HasValue Then
                vendas = vendas.Where(Function(v) v.ValorTotal >= criterio.ValorMinimo.Value).ToList()
            End If
            
            ' Filtrar por valor máximo se especificado
            If criterio.ValorMaximo.HasValue Then
                vendas = vendas.Where(Function(v) v.ValorTotal <= criterio.ValorMaximo.Value).ToList()
            End If
            
            Return vendas
            
        Catch ex As Exception
            _logger.Error("Erro ao buscar vendas", ex)
            Return New List(Of Venda)()
        End Try
    End Function
    
    ''' <summary>
    ''' Gera relatório de vendas por período
    ''' </summary>
    Public Function GerarRelatorioVendas(dataInicio As DateTime, dataFim As DateTime) As RelatorioVendas
        Try
            Dim vendas = _dataManager.ObterVendasPorPeriodo(dataInicio, dataFim)
            
            Dim relatorio = New RelatorioVendas() With {
                .DataInicio = dataInicio,
                .DataFim = dataFim,
                .DataGeracao = DateTime.Now,
                .QuantidadeVendas = vendas.Count,
                .ValorTotal = vendas.Sum(Function(v) v.ValorTotal),
                .ValorMedio = If(vendas.Count > 0, vendas.Average(Function(v) v.ValorTotal), 0),
                .MaiorVenda = If(vendas.Count > 0, vendas.Max(Function(v) v.ValorTotal), 0),
                .MenorVenda = If(vendas.Count > 0, vendas.Min(Function(v) v.ValorTotal), 0)
            }
            
            ' Produtos mais vendidos
            relatorio.ProdutosMaisVendidos = ObterProdutosMaisVendidos(vendas, 10)
            
            ' Clientes mais frequentes
            relatorio.ClientesMaisFrequentes = ObterClientesMaisFrequentes(vendas, 10)
            
            ' Vendedores com mais vendas
            relatorio.VendedoresMaisAtivos = ObterVendedoresMaisAtivos(vendas, 10)
            
            ' Vendas por dia
            relatorio.VendasPorDia = ObterVendasPorDia(vendas)
            
            _logger.Info($"Relatório de vendas gerado para período {dataInicio:dd/MM/yyyy} a {dataFim:dd/MM/yyyy}")
            
            Return relatorio
            
        Catch ex As Exception
            _logger.Error("Erro ao gerar relatório de vendas", ex)
            Return New RelatorioVendas()
        End Try
    End Function
    
    ''' <summary>
    ''' Obtém produtos mais vendidos
    ''' </summary>
    Private Function ObterProdutosMaisVendidos(vendas As List(Of Venda), top As Integer) As List(Of ProdutoVendido)
        Try
            Return vendas.SelectMany(Function(v) v.Itens) _
                       .GroupBy(Function(i) i.Produto.Descricao) _
                       .Select(Function(g) New ProdutoVendido() With {
                           .Descricao = g.Key,
                           .QuantidadeVendida = g.Sum(Function(i) i.Quantidade),
                           .ValorTotal = g.Sum(Function(i) i.ValorTotal)
                       }) _
                       .OrderByDescending(Function(p) p.QuantidadeVendida) _
                       .Take(top) _
                       .ToList()
        Catch
            Return New List(Of ProdutoVendido)()
        End Try
    End Function
    
    ''' <summary>
    ''' Obtém clientes mais frequentes
    ''' </summary>
    Private Function ObterClientesMaisFrequentes(vendas As List(Of Venda), top As Integer) As List(Of ClienteFrequente)
        Try
            Return vendas.GroupBy(Function(v) v.Cliente.Nome) _
                       .Select(Function(g) New ClienteFrequente() With {
                           .Nome = g.Key,
                           .QuantidadeCompras = g.Count(),
                           .ValorTotal = g.Sum(Function(v) v.ValorTotal)
                       }) _
                       .OrderByDescending(Function(c) c.QuantidadeCompras) _
                       .Take(top) _
                       .ToList()
        Catch
            Return New List(Of ClienteFrequente)()
        End Try
    End Function
    
    ''' <summary>
    ''' Obtém vendedores mais ativos
    ''' </summary>
    Private Function ObterVendedoresMaisAtivos(vendas As List(Of Venda), top As Integer) As List(Of VendedorAtivo)
        Try
            Return vendas.GroupBy(Function(v) v.Vendedor) _
                       .Select(Function(g) New VendedorAtivo() With {
                           .Nome = g.Key,
                           .QuantidadeVendas = g.Count(),
                           .ValorTotal = g.Sum(Function(v) v.ValorTotal)
                       }) _
                       .OrderByDescending(Function(v) v.QuantidadeVendas) _
                       .Take(top) _
                       .ToList()
        Catch
            Return New List(Of VendedorAtivo)()
        End Try
    End Function
    
    ''' <summary>
    ''' Obtém vendas por dia
    ''' </summary>
    Private Function ObterVendasPorDia(vendas As List(Of Venda)) As List(Of VendaPorDia)
        Try
            Return vendas.GroupBy(Function(v) v.DataVenda.Date) _
                       .Select(Function(g) New VendaPorDia() With {
                           .Data = g.Key,
                           .QuantidadeVendas = g.Count(),
                           .ValorTotal = g.Sum(Function(v) v.ValorTotal)
                       }) _
                       .OrderBy(Function(v) v.Data) _
                       .ToList()
        Catch
            Return New List(Of VendaPorDia)()
        End Try
    End Function
    
    ''' <summary>
    ''' Exporta relatório para XML
    ''' </summary>
    Public Function ExportarRelatorio(relatorio As RelatorioVendas, caminhoArquivo As String) As Boolean
        Try
            Dim serializer = New XmlSerializer(GetType(RelatorioVendas))
            Using writer = New FileStream(caminhoArquivo, FileMode.Create)
                serializer.Serialize(writer, relatorio)
            End Using
            
            _logger.Info($"Relatório exportado para {caminhoArquivo}")
            Return True
            
        Catch ex As Exception
            _logger.Error("Erro ao exportar relatório", ex)
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' Gera relatório de produtos em baixo estoque
    ''' </summary>
    Public Function GerarRelatorioProdutosBaixoEstoque() As List(Of Produto)
        Try
            Return _dataManager.ObterProdutosAtivos() _
                             .Where(Function(p) p.EstoqueBaixo()) _
                             .OrderBy(Function(p) p.EstoqueAtual) _
                             .ToList()
        Catch ex As Exception
            _logger.Error("Erro ao gerar relatório de baixo estoque", ex)
            Return New List(Of Produto)()
        End Try
    End Function
    
    ''' <summary>
    ''' Limpa histórico antigo baseado na configuração
    ''' </summary>
    Public Function LimparHistoricoAntigo() As Boolean
        Try
            Dim diasManter = _config.ManterHistoricoDias
            Dim dataLimite = DateTime.Now.AddDays(-diasManter)
            
            Dim vendasAntigas = _dataManager.ObterVendasPorPeriodo(DateTime.MinValue, dataLimite)
            Dim countRemovidas = 0
            
            ' Note: Esta implementação seria expandida para realmente remover as vendas antigas
            ' Por ora, apenas registra o que seria removido
            
            _logger.Info($"Limpeza de histórico: {vendasAntigas.Count} vendas seriam removidas (anteriores a {dataLimite:dd/MM/yyyy})")
            
            Return True
            
        Catch ex As Exception
            _logger.Error("Erro ao limpar histórico antigo", ex)
            Return False
        End Try
    End Function
End Class

#Region "Classes de Relatório"

''' <summary>
''' Critério de busca para vendas
''' </summary>
Public Class CriterioBusca
    Public Property DataInicio As DateTime
    Public Property DataFim As DateTime
    Public Property NomeCliente As String
    Public Property Vendedor As String
    Public Property NumeroTalao As String
    Public Property ValorMinimo As Decimal?
    Public Property ValorMaximo As Decimal?
End Class

''' <summary>
''' Relatório de vendas
''' </summary>
Public Class RelatorioVendas
    Public Property DataInicio As DateTime
    Public Property DataFim As DateTime
    Public Property DataGeracao As DateTime
    Public Property QuantidadeVendas As Integer
    Public Property ValorTotal As Decimal
    Public Property ValorMedio As Decimal
    Public Property MaiorVenda As Decimal
    Public Property MenorVenda As Decimal
    Public Property ProdutosMaisVendidos As List(Of ProdutoVendido)
    Public Property ClientesMaisFrequentes As List(Of ClienteFrequente)
    Public Property VendedoresMaisAtivos As List(Of VendedorAtivo)
    Public Property VendasPorDia As List(Of VendaPorDia)
    
    Public Sub New()
        ProdutosMaisVendidos = New List(Of ProdutoVendido)()
        ClientesMaisFrequentes = New List(Of ClienteFrequente)()
        VendedoresMaisAtivos = New List(Of VendedorAtivo)()
        VendasPorDia = New List(Of VendaPorDia)()
    End Sub
End Class

''' <summary>
''' Produto vendido no relatório
''' </summary>
Public Class ProdutoVendido
    Public Property Descricao As String
    Public Property QuantidadeVendida As Decimal
    Public Property ValorTotal As Decimal
End Class

''' <summary>
''' Cliente frequente no relatório
''' </summary>
Public Class ClienteFrequente
    Public Property Nome As String
    Public Property QuantidadeCompras As Integer
    Public Property ValorTotal As Decimal
End Class

''' <summary>
''' Vendedor ativo no relatório
''' </summary>
Public Class VendedorAtivo
    Public Property Nome As String
    Public Property QuantidadeVendas As Integer
    Public Property ValorTotal As Decimal
End Class

''' <summary>
''' Venda por dia no relatório
''' </summary>
Public Class VendaPorDia
    Public Property Data As DateTime
    Public Property QuantidadeVendas As Integer
    Public Property ValorTotal As Decimal
End Class

#End Region