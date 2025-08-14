Imports System.IO
Imports System.Xml.Serialization
Imports System.Collections.Concurrent

''' <summary>
''' Gerenciador centralizado de dados
''' Gerencia persistência e cache de dados do sistema
''' </summary>
Public Class DataManager
    Private Shared ReadOnly _instance As New Lazy(Of DataManager)(Function() New DataManager())
    Private ReadOnly _logger As Logger
    Private ReadOnly _config As ConfigManager
    Private ReadOnly _dataPath As String
    
    ' Cache de dados em memória
    Private ReadOnly _vendas As ConcurrentDictionary(Of String, Venda)
    Private ReadOnly _clientes As ConcurrentDictionary(Of Integer, Cliente)
    Private ReadOnly _produtos As ConcurrentDictionary(Of Integer, Produto)
    
    ''' <summary>
    ''' Instância singleton do DataManager
    ''' </summary>
    Public Shared ReadOnly Property Instance As DataManager
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
        _dataPath = Path.Combine(Application.StartupPath, "Data")
        
        ' Inicializar cache
        _vendas = New ConcurrentDictionary(Of String, Venda)()
        _clientes = New ConcurrentDictionary(Of Integer, Cliente)()
        _produtos = New ConcurrentDictionary(Of Integer, Produto)()
        
        ' Criar diretório de dados
        If Not Directory.Exists(_dataPath) Then
            Directory.CreateDirectory(_dataPath)
        End If
        
        ' Carregar dados existentes
        CarregarDados()
    End Sub
    
    #Region "Gestão de Vendas"
    
    ''' <summary>
    ''' Salva uma venda no sistema
    ''' </summary>
    Public Function SalvarVenda(venda As Venda) As Boolean
        Try
            ' Validar venda
            If Not venda.IsValid() Then
                _logger.Warning($"Tentativa de salvar venda inválida: {venda.NumeroTalao}")
                Return False
            End If
            
            ' Adicionar/atualizar no cache
            _vendas.AddOrUpdate(venda.NumeroTalao, venda, Function(key, oldValue) venda)
            
            ' Persistir no arquivo
            PersistirVendas()
            
            _logger.Info($"Venda salva: {venda.NumeroTalao}")
            Return True
            
        Catch ex As Exception
            _logger.Error($"Erro ao salvar venda {venda.NumeroTalao}", ex)
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' Atualiza uma venda existente
    ''' </summary>
    Public Function AtualizarVenda(venda As Venda) As Boolean
        Try
            If _vendas.ContainsKey(venda.NumeroTalao) Then
                _vendas(venda.NumeroTalao) = venda
                PersistirVendas()
                _logger.Info($"Venda atualizada: {venda.NumeroTalao}")
                Return True
            Else
                _logger.Warning($"Tentativa de atualizar venda inexistente: {venda.NumeroTalao}")
                Return False
            End If
            
        Catch ex As Exception
            _logger.Error($"Erro ao atualizar venda {venda.NumeroTalao}", ex)
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' Busca venda por número do talão
    ''' </summary>
    Public Function BuscarVendaPorTalao(numeroTalao As String) As Venda
        Try
            If _vendas.ContainsKey(numeroTalao) Then
                Return _vendas(numeroTalao)
            End If
            Return Nothing
        Catch ex As Exception
            _logger.Error($"Erro ao buscar venda {numeroTalao}", ex)
            Return Nothing
        End Try
    End Function
    
    ''' <summary>
    ''' Obtém vendas por período
    ''' </summary>
    Public Function ObterVendasPorPeriodo(dataInicio As DateTime, dataFim As DateTime) As List(Of Venda)
        Try
            Return _vendas.Values.Where(Function(v) v.DataVenda >= dataInicio AndAlso v.DataVenda <= dataFim).ToList()
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
            Return _vendas.Values.Where(Function(v) v.Cliente?.Id = clienteId).ToList()
        Catch ex As Exception
            _logger.Error($"Erro ao obter vendas do cliente {clienteId}", ex)
            Return New List(Of Venda)()
        End Try
    End Function
    
    ''' <summary>
    ''' Obtém todas as vendas
    ''' </summary>
    Public Function ObterTodasVendas() As List(Of Venda)
        Try
            Return _vendas.Values.ToList()
        Catch ex As Exception
            _logger.Error("Erro ao obter todas as vendas", ex)
            Return New List(Of Venda)()
        End Try
    End Function
    
    #End Region
    
    #Region "Gestão de Clientes"
    
    ''' <summary>
    ''' Salva um cliente no sistema
    ''' </summary>
    Public Function SalvarCliente(cliente As Cliente) As Boolean
        Try
            If Not cliente.IsValid() Then
                _logger.Warning($"Tentativa de salvar cliente inválido: {cliente.Nome}")
                Return False
            End If
            
            ' Gerar ID se necessário
            If cliente.Id = 0 Then
                cliente.Id = ObterProximoIdCliente()
            End If
            
            _clientes.AddOrUpdate(cliente.Id, cliente, Function(key, oldValue) cliente)
            PersistirClientes()
            
            _logger.Info($"Cliente salvo: {cliente.Nome}")
            Return True
            
        Catch ex As Exception
            _logger.Error($"Erro ao salvar cliente {cliente.Nome}", ex)
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' Busca cliente por ID
    ''' </summary>
    Public Function BuscarClientePorId(id As Integer) As Cliente
        Try
            If _clientes.ContainsKey(id) Then
                Return _clientes(id)
            End If
            Return Nothing
        Catch ex As Exception
            _logger.Error($"Erro ao buscar cliente {id}", ex)
            Return Nothing
        End Try
    End Function
    
    ''' <summary>
    ''' Busca clientes por nome
    ''' </summary>
    Public Function BuscarClientesPorNome(nome As String) As List(Of Cliente)
        Try
            Return _clientes.Values.Where(Function(c) c.Nome.ToLower().Contains(nome.ToLower())).ToList()
        Catch ex As Exception
            _logger.Error($"Erro ao buscar clientes por nome {nome}", ex)
            Return New List(Of Cliente)()
        End Try
    End Function
    
    ''' <summary>
    ''' Obtém todos os clientes
    ''' </summary>
    Public Function ObterTodosClientes() As List(Of Cliente)
        Try
            Return _clientes.Values.ToList()
        Catch ex As Exception
            _logger.Error("Erro ao obter todos os clientes", ex)
            Return New List(Of Cliente)()
        End Try
    End Function
    
    ''' <summary>
    ''' Obtém próximo ID de cliente
    ''' </summary>
    Private Function ObterProximoIdCliente() As Integer
        Return If(_clientes.Any(), _clientes.Keys.Max() + 1, 1)
    End Function
    
    #End Region
    
    #Region "Gestão de Produtos"
    
    ''' <summary>
    ''' Salva um produto no sistema
    ''' </summary>
    Public Function SalvarProduto(produto As Produto) As Boolean
        Try
            If Not produto.IsValid() Then
                _logger.Warning($"Tentativa de salvar produto inválido: {produto.Descricao}")
                Return False
            End If
            
            ' Gerar ID se necessário
            If produto.Id = 0 Then
                produto.Id = ObterProximoIdProduto()
            End If
            
            _produtos.AddOrUpdate(produto.Id, produto, Function(key, oldValue) produto)
            PersistirProdutos()
            
            _logger.Info($"Produto salvo: {produto.Descricao}")
            Return True
            
        Catch ex As Exception
            _logger.Error($"Erro ao salvar produto {produto.Descricao}", ex)
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' Busca produto por código
    ''' </summary>
    Public Function BuscarProdutoPorCodigo(codigo As String) As Produto
        Try
            Return _produtos.Values.FirstOrDefault(Function(p) p.Codigo.Equals(codigo, StringComparison.OrdinalIgnoreCase))
        Catch ex As Exception
            _logger.Error($"Erro ao buscar produto {codigo}", ex)
            Return Nothing
        End Try
    End Function
    
    ''' <summary>
    ''' Busca produtos por descrição
    ''' </summary>
    Public Function BuscarProdutosPorDescricao(descricao As String) As List(Of Produto)
        Try
            Return _produtos.Values.Where(Function(p) p.Descricao.ToLower().Contains(descricao.ToLower()) AndAlso p.Ativo).ToList()
        Catch ex As Exception
            _logger.Error($"Erro ao buscar produtos por descrição {descricao}", ex)
            Return New List(Of Produto)()
        End Try
    End Function
    
    ''' <summary>
    ''' Obtém todos os produtos ativos
    ''' </summary>
    Public Function ObterProdutosAtivos() As List(Of Produto)
        Try
            Return _produtos.Values.Where(Function(p) p.Ativo).ToList()
        Catch ex As Exception
            _logger.Error("Erro ao obter produtos ativos", ex)
            Return New List(Of Produto)()
        End Try
    End Function
    
    ''' <summary>
    ''' Obtém próximo ID de produto
    ''' </summary>
    Private Function ObterProximoIdProduto() As Integer
        Return If(_produtos.Any(), _produtos.Keys.Max() + 1, 1)
    End Function
    
    #End Region
    
    #Region "Persistência de Dados"
    
    ''' <summary>
    ''' Carrega dados dos arquivos
    ''' </summary>
    Private Sub CarregarDados()
        Try
            CarregarVendas()
            CarregarClientes()
            CarregarProdutos()
            _logger.Info("Dados carregados com sucesso")
        Catch ex As Exception
            _logger.Error("Erro ao carregar dados", ex)
        End Try
    End Sub
    
    ''' <summary>
    ''' Carrega vendas do arquivo
    ''' </summary>
    Private Sub CarregarVendas()
        Try
            Dim vendasPath = Path.Combine(_dataPath, "vendas.xml")
            If File.Exists(vendasPath) Then
                Dim vendas = DeserializarObjeto(Of List(Of Venda))(vendasPath)
                If vendas IsNot Nothing Then
                    For Each venda In vendas
                        _vendas.TryAdd(venda.NumeroTalao, venda)
                    Next
                End If
            End If
        Catch ex As Exception
            _logger.Error("Erro ao carregar vendas", ex)
        End Try
    End Sub
    
    ''' <summary>
    ''' Carrega clientes do arquivo
    ''' </summary>
    Private Sub CarregarClientes()
        Try
            Dim clientesPath = Path.Combine(_dataPath, "clientes.xml")
            If File.Exists(clientesPath) Then
                Dim clientes = DeserializarObjeto(Of List(Of Cliente))(clientesPath)
                If clientes IsNot Nothing Then
                    For Each cliente In clientes
                        _clientes.TryAdd(cliente.Id, cliente)
                    Next
                End If
            End If
        Catch ex As Exception
            _logger.Error("Erro ao carregar clientes", ex)
        End Try
    End Sub
    
    ''' <summary>
    ''' Carrega produtos do arquivo
    ''' </summary>
    Private Sub CarregarProdutos()
        Try
            Dim produtosPath = Path.Combine(_dataPath, "produtos.xml")
            If File.Exists(produtosPath) Then
                Dim produtos = DeserializarObjeto(Of List(Of Produto))(produtosPath)
                If produtos IsNot Nothing Then
                    For Each produto In produtos
                        _produtos.TryAdd(produto.Id, produto)
                    Next
                End If
            Else
                ' Carregar produtos padrão se não existir arquivo
                CarregarProdutosPadrao()
            End If
        Catch ex As Exception
            _logger.Error("Erro ao carregar produtos", ex)
        End Try
    End Sub
    
    ''' <summary>
    ''' Carrega produtos padrão para madeireira
    ''' </summary>
    Private Sub CarregarProdutosPadrao()
        Try
            Dim produtosPadrao = New List(Of Produto) From {
                New Produto("001", "Tábua de Pinus 2x4m", "UN", 25.0D) With {.Id = 1},
                New Produto("002", "Ripão 3x3x3m", "UN", 15.0D) With {.Id = 2},
                New Produto("003", "Compensado 18mm", "M²", 45.0D) With {.Id = 3},
                New Produto("004", "Caibro 5x6x3m", "UN", 12.0D) With {.Id = 4},
                New Produto("005", "Viga 6x12x4m", "UN", 35.0D) With {.Id = 5}
            }
            
            For Each produto In produtosPadrao
                _produtos.TryAdd(produto.Id, produto)
            Next
            
            PersistirProdutos()
            _logger.Info("Produtos padrão carregados")
            
        Catch ex As Exception
            _logger.Error("Erro ao carregar produtos padrão", ex)
        End Try
    End Sub
    
    ''' <summary>
    ''' Persiste vendas no arquivo
    ''' </summary>
    Private Sub PersistirVendas()
        Try
            Dim vendasPath = Path.Combine(_dataPath, "vendas.xml")
            SerializarObjeto(_vendas.Values.ToList(), vendasPath)
        Catch ex As Exception
            _logger.Error("Erro ao persistir vendas", ex)
        End Try
    End Sub
    
    ''' <summary>
    ''' Persiste clientes no arquivo
    ''' </summary>
    Private Sub PersistirClientes()
        Try
            Dim clientesPath = Path.Combine(_dataPath, "clientes.xml")
            SerializarObjeto(_clientes.Values.ToList(), clientesPath)
        Catch ex As Exception
            _logger.Error("Erro ao persistir clientes", ex)
        End Try
    End Sub
    
    ''' <summary>
    ''' Persiste produtos no arquivo
    ''' </summary>
    Private Sub PersistirProdutos()
        Try
            Dim produtosPath = Path.Combine(_dataPath, "produtos.xml")
            SerializarObjeto(_produtos.Values.ToList(), produtosPath)
        Catch ex As Exception
            _logger.Error("Erro ao persistir produtos", ex)
        End Try
    End Sub
    
    ''' <summary>
    ''' Serializa objeto para XML
    ''' </summary>
    Private Sub SerializarObjeto(Of T)(obj As T, filePath As String)
        Dim serializer = New XmlSerializer(GetType(T))
        Using writer = New FileStream(filePath, FileMode.Create)
            serializer.Serialize(writer, obj)
        End Using
    End Sub
    
    ''' <summary>
    ''' Deserializa objeto do XML
    ''' </summary>
    Private Function DeserializarObjeto(Of T)(filePath As String) As T
        Dim serializer = New XmlSerializer(GetType(T))
        Using reader = New FileStream(filePath, FileMode.Open)
            Return CType(serializer.Deserialize(reader), T)
        End Using
    End Function
    
    #End Region
    
    ''' <summary>
    ''' Limpa cache de dados
    ''' </summary>
    Public Sub LimparCache()
        Try
            _vendas.Clear()
            _clientes.Clear()
            _produtos.Clear()
            CarregarDados()
            _logger.Info("Cache limpo e dados recarregados")
        Catch ex As Exception
            _logger.Error("Erro ao limpar cache", ex)
        End Try
    End Sub
    
    ''' <summary>
    ''' Obtém estatísticas do cache
    ''' </summary>
    Public Function ObterEstatisticasCache() As String
        Return $"Vendas: {_vendas.Count}, Clientes: {_clientes.Count}, Produtos: {_produtos.Count}"
    End Function
End Class