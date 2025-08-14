''' <summary>
''' Serviço principal para gerenciamento de vendas
''' Centraliza todas as operações relacionadas a vendas
''' </summary>
Public Class VendaService
    Private ReadOnly _dataManager As DataManager
    Private ReadOnly _excelService As ExcelService
    Private ReadOnly _logger As Logger
    
    ''' <summary>
    ''' Construtor
    ''' </summary>
    Public Sub New()
        _dataManager = DataManager.Instance
        _excelService = New ExcelService()
        _logger = Logger.Instance
    End Sub
    
    ''' <summary>
    ''' Processa uma venda completa
    ''' </summary>
    Public Function ProcessarVenda(venda As Venda) As Boolean
        Try
            _logger.Info($"Iniciando processamento da venda {venda.NumeroTalao}")
            
            ' Validar venda
            If Not ValidarVenda(venda) Then
                _logger.Warning($"Venda {venda.NumeroTalao} inválida")
                Return False
            End If
            
            ' Salvar venda no histórico
            _dataManager.SalvarVenda(venda)
            
            ' Gerar talão no Excel
            Dim talaoGerado = _excelService.GerarTalao(venda)
            
            If talaoGerado Then
                venda.Status = StatusVenda.Finalizada
                _dataManager.AtualizarVenda(venda)
                
                ' Log de auditoria
                _logger.Audit("VENDA_FINALIZADA", 
                    $"Talão: {venda.NumeroTalao}, Cliente: {venda.Cliente.Nome}, Valor: {venda.ValorTotal:C}",
                    venda.Vendedor)
                
                Return True
            Else
                _logger.Error($"Erro ao gerar talão para venda {venda.NumeroTalao}")
                Return False
            End If
            
        Catch ex As Exception
            _logger.Error($"Erro ao processar venda {venda.NumeroTalao}", ex)
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' Valida uma venda antes do processamento
    ''' </summary>
    Private Function ValidarVenda(venda As Venda) As Boolean
        ' Validação básica
        If Not venda.IsValid() Then
            Return False
        End If
        
        ' Validar cliente
        If venda.Cliente Is Nothing OrElse Not venda.Cliente.IsValid() Then
            Return False
        End If
        
        ' Validar itens
        If venda.Itens.Count = 0 Then
            Return False
        End If
        
        For Each item In venda.Itens
            If Not item.IsValid() Then
                Return False
            End If
        Next
        
        ' Validar valores
        If venda.ValorTotal <= 0 Then
            Return False
        End If
        
        Return True
    End Function
    
    ''' <summary>
    ''' Obtém histórico de vendas por período
    ''' </summary>
    Public Function ObterVendasPorPeriodo(dataInicio As DateTime, dataFim As DateTime) As List(Of Venda)
        Try
            Return _dataManager.ObterVendasPorPeriodo(dataInicio, dataFim)
        Catch ex As Exception
            _logger.Error("Erro ao obter vendas por período", ex)
            Return New List(Of Venda)()
        End Try
    End Function
    
    ''' <summary>
    ''' Obtém vendas por cliente
    ''' </summary>
    Public Function ObterVendasPorCliente(clienteId As Integer) As List(Of Venda)
        Try
            Return _dataManager.ObterVendasPorCliente(clienteId)
        Catch ex As Exception
            _logger.Error($"Erro ao obter vendas do cliente {clienteId}", ex)
            Return New List(Of Venda)()
        End Try
    End Function
    
    ''' <summary>
    ''' Busca venda por número do talão
    ''' </summary>
    Public Function BuscarVendaPorTalao(numeroTalao As String) As Venda
        Try
            Return _dataManager.BuscarVendaPorTalao(numeroTalao)
        Catch ex As Exception
            _logger.Error($"Erro ao buscar venda pelo talão {numeroTalao}", ex)
            Return Nothing
        End Try
    End Function
    
    ''' <summary>
    ''' Cancela uma venda
    ''' </summary>
    Public Function CancelarVenda(numeroTalao As String, motivo As String) As Boolean
        Try
            Dim venda = _dataManager.BuscarVendaPorTalao(numeroTalao)
            If venda Is Nothing Then
                Return False
            End If
            
            venda.Status = StatusVenda.Cancelada
            venda.Observacoes = $"Cancelada: {motivo}"
            
            _dataManager.AtualizarVenda(venda)
            
            _logger.Audit("VENDA_CANCELADA", 
                $"Talão: {numeroTalao}, Motivo: {motivo}",
                "Sistema")
            
            Return True
        Catch ex As Exception
            _logger.Error($"Erro ao cancelar venda {numeroTalao}", ex)
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' Reimprimir talão
    ''' </summary>
    Public Function ReimprimirTalao(numeroTalao As String) As Boolean
        Try
            Dim venda = _dataManager.BuscarVendaPorTalao(numeroTalao)
            If venda Is Nothing Then
                Return False
            End If
            
            Dim resultado = _excelService.GerarTalao(venda, True) ' true para reimpressão
            
            If resultado Then
                _logger.Audit("TALAO_REIMPRESSO", 
                    $"Talão: {numeroTalao}",
                    "Sistema")
            End If
            
            Return resultado
        Catch ex As Exception
            _logger.Error($"Erro ao reimprimir talão {numeroTalao}", ex)
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' Obtém estatísticas de vendas
    ''' </summary>
    Public Function ObterEstatisticas(periodo As DateTime) As VendaEstatisticas
        Try
            Dim vendas = _dataManager.ObterVendasPorPeriodo(periodo, DateTime.Now)
            
            Dim stats = New VendaEstatisticas() With {
                .TotalVendas = vendas.Count,
                .ValorTotal = vendas.Sum(Function(v) v.ValorTotal),
                .TicketMedio = If(vendas.Count > 0, vendas.Average(Function(v) v.ValorTotal), 0),
                .ProdutoMaisVendido = ObterProdutoMaisVendido(vendas),
                .ClienteMaisFrequente = ObterClienteMaisFrequente(vendas)
            }
            
            Return stats
        Catch ex As Exception
            _logger.Error("Erro ao obter estatísticas", ex)
            Return New VendaEstatisticas()
        End Try
    End Function
    
    ''' <summary>
    ''' Obtém produto mais vendido
    ''' </summary>
    Private Function ObterProdutoMaisVendido(vendas As List(Of Venda)) As String
        Try
            Dim produtos = vendas.SelectMany(Function(v) v.Itens) _
                              .GroupBy(Function(i) i.Produto.Descricao) _
                              .OrderByDescending(Function(g) g.Sum(Function(i) i.Quantidade)) _
                              .FirstOrDefault()
            
            Return If(produtos IsNot Nothing, produtos.Key, "N/A")
        Catch
            Return "N/A"
        End Try
    End Function
    
    ''' <summary>
    ''' Obtém cliente mais frequente
    ''' </summary>
    Private Function ObterClienteMaisFrequente(vendas As List(Of Venda)) As String
        Try
            Dim cliente = vendas.GroupBy(Function(v) v.Cliente.Nome) _
                              .OrderByDescending(Function(g) g.Count()) _
                              .FirstOrDefault()
            
            Return If(cliente IsNot Nothing, cliente.Key, "N/A")
        Catch
            Return "N/A"
        End Try
    End Function
End Class

''' <summary>
''' Estatísticas de vendas
''' </summary>
Public Class VendaEstatisticas
    Public Property TotalVendas As Integer
    Public Property ValorTotal As Decimal
    Public Property TicketMedio As Decimal
    Public Property ProdutoMaisVendido As String
    Public Property ClienteMaisFrequente As String
End Class