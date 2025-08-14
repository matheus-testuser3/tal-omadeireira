Imports System.Collections.Concurrent
Imports System.Threading

''' <summary>
''' Sistema de cache thread-safe para melhorar performance
''' </summary>
Public Class CacheManager
    
    Private Shared _instance As CacheManager
    Private Shared ReadOnly _lockObject As New Object()
    
    Private ReadOnly _cache As New ConcurrentDictionary(Of String, CacheItem)()
    Private ReadOnly _logger As LoggingSystem = LoggingSystem.Instance
    Private ReadOnly _config As EnhancedConfigurationManager = EnhancedConfigurationManager.Instance
    Private ReadOnly _cleanupTimer As Timer
    
    ''' <summary>
    ''' Singleton instance
    ''' </summary>
    Public Shared ReadOnly Property Instance As CacheManager
        Get
            If _instance Is Nothing Then
                SyncLock _lockObject
                    If _instance Is Nothing Then
                        _instance = New CacheManager()
                    End If
                End SyncLock
            End If
            Return _instance
        End Get
    End Property
    
    ''' <summary>
    ''' Construtor privado
    ''' </summary>
    Private Sub New()
        ' Timer para limpeza automática a cada 5 minutos
        _cleanupTimer = New Timer(AddressOf CleanupExpiredItems, Nothing, TimeSpan.FromMinutes(5), TimeSpan.FromMinutes(5))
        _logger.LogInfo("CacheManager", "Sistema de cache inicializado")
    End Sub
    
    ''' <summary>
    ''' Adiciona item ao cache
    ''' </summary>
    Public Sub Set(Of T)(key As String, value As T, Optional expirationMinutes As Integer? = Nothing)
        If Not _config.CacheEnabled Then Return
        
        Try
            Dim expiration = If(expirationMinutes, _config.CacheExpirationMinutes)
            Dim expiryTime = Date.Now.AddMinutes(expiration)
            
            Dim cacheItem = New CacheItem() With {
                .Value = value,
                .ExpiryTime = expiryTime,
                .CreatedTime = Date.Now,
                .AccessCount = 0
            }
            
            _cache.AddOrUpdate(key, cacheItem, Function(k, existing) cacheItem)
            
            _logger.LogDebug("CacheManager", $"Item adicionado ao cache: {key} (expira em {expiration} min)")
            
        Catch ex As Exception
            _logger.LogError("CacheManager", $"Erro ao adicionar item ao cache: {key}", ex)
        End Try
    End Sub
    
    ''' <summary>
    ''' Obtém item do cache
    ''' </summary>
    Public Function Get(Of T)(key As String) As T
        If Not _config.CacheEnabled Then Return Nothing
        
        Try
            Dim cacheItem As CacheItem = Nothing
            If _cache.TryGetValue(key, cacheItem) Then
                
                ' Verificar se expirou
                If Date.Now > cacheItem.ExpiryTime Then
                    Remove(key)
                    _logger.LogDebug("CacheManager", $"Item expirado removido do cache: {key}")
                    Return Nothing
                End If
                
                ' Atualizar estatísticas de acesso
                Interlocked.Increment(cacheItem.AccessCount)
                cacheItem.LastAccessTime = Date.Now
                
                _logger.LogDebug("CacheManager", $"Item encontrado no cache: {key}")
                Return DirectCast(cacheItem.Value, T)
            End If
            
            Return Nothing
            
        Catch ex As Exception
            _logger.LogError("CacheManager", $"Erro ao obter item do cache: {key}", ex)
            Return Nothing
        End Try
    End Function
    
    ''' <summary>
    ''' Obtém item do cache ou executa função para criá-lo
    ''' </summary>
    Public Function GetOrSet(Of T)(key As String, factory As Func(Of T), Optional expirationMinutes As Integer? = Nothing) As T
        Dim cachedValue = Get(Of T)(key)
        
        If cachedValue IsNot Nothing Then
            Return cachedValue
        End If
        
        ' Item não está no cache, criar
        Try
            Dim newValue = factory()
            Set(key, newValue, expirationMinutes)
            Return newValue
            
        Catch ex As Exception
            _logger.LogError("CacheManager", $"Erro ao executar factory para chave: {key}", ex)
            Throw
        End Try
    End Function
    
    ''' <summary>
    ''' Remove item do cache
    ''' </summary>
    Public Sub Remove(key As String)
        Try
            Dim removed As CacheItem = Nothing
            If _cache.TryRemove(key, removed) Then
                _logger.LogDebug("CacheManager", $"Item removido do cache: {key}")
            End If
            
        Catch ex As Exception
            _logger.LogError("CacheManager", $"Erro ao remover item do cache: {key}", ex)
        End Try
    End Sub
    
    ''' <summary>
    ''' Verifica se item existe no cache e não expirou
    ''' </summary>
    Public Function Contains(key As String) As Boolean
        If Not _config.CacheEnabled Then Return False
        
        Try
            Dim cacheItem As CacheItem = Nothing
            If _cache.TryGetValue(key, cacheItem) Then
                If Date.Now <= cacheItem.ExpiryTime Then
                    Return True
                Else
                    Remove(key) ' Remove item expirado
                End If
            End If
            
            Return False
            
        Catch ex As Exception
            _logger.LogError("CacheManager", $"Erro ao verificar existência no cache: {key}", ex)
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' Limpa todo o cache
    ''' </summary>
    Public Sub Clear()
        Try
            Dim count = _cache.Count
            _cache.Clear()
            _logger.LogInfo("CacheManager", $"Cache limpo: {count} itens removidos")
            
        Catch ex As Exception
            _logger.LogError("CacheManager", "Erro ao limpar cache", ex)
        End Try
    End Sub
    
    ''' <summary>
    ''' Obtém estatísticas do cache
    ''' </summary>
    Public Function GetStatistics() As CacheStatistics
        Try
            Dim stats = New CacheStatistics()
            
            stats.TotalItems = _cache.Count
            stats.IsEnabled = _config.CacheEnabled
            
            If _cache.Count > 0 Then
                Dim items = _cache.Values.ToList()
                stats.ExpiredItems = items.Count(Function(i) Date.Now > i.ExpiryTime)
                stats.TotalAccesses = items.Sum(Function(i) i.AccessCount)
                stats.AverageAccessCount = If(items.Count > 0, stats.TotalAccesses / items.Count, 0)
                
                If items.Any() Then
                    stats.OldestItemAge = Date.Now.Subtract(items.Min(Function(i) i.CreatedTime))
                    stats.NewestItemAge = Date.Now.Subtract(items.Max(Function(i) i.CreatedTime))
                End If
            End If
            
            Return stats
            
        Catch ex As Exception
            _logger.LogError("CacheManager", "Erro ao obter estatísticas do cache", ex)
            Return New CacheStatistics()
        End Try
    End Function
    
    ''' <summary>
    ''' Remove itens expirados (chamado pelo timer)
    ''' </summary>
    Private Sub CleanupExpiredItems(state As Object)
        Try
            Dim now = Date.Now
            Dim expiredKeys = New List(Of String)()
            
            For Each kvp In _cache
                If now > kvp.Value.ExpiryTime Then
                    expiredKeys.Add(kvp.Key)
                End If
            Next
            
            For Each key In expiredKeys
                Remove(key)
            Next
            
            If expiredKeys.Count > 0 Then
                _logger.LogDebug("CacheManager", $"Limpeza automática: {expiredKeys.Count} itens expirados removidos")
            End If
            
        Catch ex As Exception
            _logger.LogError("CacheManager", "Erro na limpeza automática do cache", ex)
        End Try
    End Sub
    
    ''' <summary>
    ''' Libera recursos do cache
    ''' </summary>
    Public Sub Dispose()
        Try
            _cleanupTimer?.Dispose()
            Clear()
            _logger.LogInfo("CacheManager", "Cache manager finalizado")
            
        Catch ex As Exception
            _logger.LogError("CacheManager", "Erro ao finalizar cache manager", ex)
        End Try
    End Sub
End Class

''' <summary>
''' Item individual do cache
''' </summary>
Friend Class CacheItem
    Public Property Value As Object
    Public Property ExpiryTime As Date
    Public Property CreatedTime As Date
    Public Property LastAccessTime As Date
    Public Property AccessCount As Integer
End Class

''' <summary>
''' Estatísticas do cache
''' </summary>
Public Class CacheStatistics
    Public Property TotalItems As Integer
    Public Property ExpiredItems As Integer
    Public Property TotalAccesses As Long
    Public Property AverageAccessCount As Double
    Public Property IsEnabled As Boolean
    Public Property OldestItemAge As TimeSpan
    Public Property NewestItemAge As TimeSpan
    
    Public Overrides Function ToString() As String
        Return $"Cache: {TotalItems} itens, {ExpiredItems} expirados, {TotalAccesses} acessos totais, Média: {AverageAccessCount:F1} acessos/item"
    End Function
End Class

''' <summary>
''' Extensões para facilitar uso do cache
''' </summary>
Public Module CacheExtensions
    
    ''' <summary>
    ''' Cache para clientes
    ''' </summary>
    <Extension>
    Public Function CacheClientes(repo As ClienteRepository) As List(Of Cliente)
        Return CacheManager.Instance.GetOrSet("todos_clientes", 
            Function() repo.GetAll(), 
            30) ' Cache por 30 minutos
    End Function
    
    ''' <summary>
    ''' Cache para produtos
    ''' </summary>
    <Extension>
    Public Function CacheProdutos(repo As ProdutoRepository) As List(Of Produto)
        Return CacheManager.Instance.GetOrSet("todos_produtos", 
            Function() repo.GetAll(), 
            60) ' Cache por 1 hora
    End Function
    
    ''' <summary>
    ''' Cache para busca de produtos
    ''' </summary>
    <Extension>
    Public Function CacheBuscarProdutos(repo As ProdutoRepository, termo As String) As List(Of Produto)
        Dim cacheKey = $"busca_produtos_{termo.ToLower().Trim()}"
        Return CacheManager.Instance.GetOrSet(cacheKey, 
            Function() repo.BuscarPorDescricaoOuCodigo(termo), 
            15) ' Cache por 15 minutos
    End Function
    
    ''' <summary>
    ''' Cache para configurações
    ''' </summary>
    <Extension>
    Public Function CacheConfiguracao(config As EnhancedConfigurationManager, key As String, defaultValue As String) As String
        Dim cacheKey = $"config_{key}"
        Return CacheManager.Instance.GetOrSet(cacheKey, 
            Function() System.Configuration.ConfigurationManager.AppSettings(key) Or defaultValue, 
            120) ' Cache por 2 horas
    End Function
End Module