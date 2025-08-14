Imports System.Configuration
Imports System.IO

''' <summary>
''' Sistema centralizado de gerenciamento de configurações
''' Fornece acesso thread-safe e tipado às configurações do sistema
''' </summary>
Public Class ConfigurationManager
    
    Private Shared _instance As ConfigurationManager
    Private Shared ReadOnly _lockObject As New Object()
    Private ReadOnly _logger As LoggingSystem = LoggingSystem.Instance
    Private ReadOnly _configCache As New Dictionary(Of String, Object)()
    Private _lastConfigCheck As Date = Date.MinValue
    
    ''' <summary>
    ''' Singleton instance
    ''' </summary>
    Public Shared ReadOnly Property Instance As ConfigurationManager
        Get
            If _instance Is Nothing Then
                SyncLock _lockObject
                    If _instance Is Nothing Then
                        _instance = New ConfigurationManager()
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
        LoadConfigurations()
        _logger.LogInfo("ConfigurationManager", "Sistema de configurações inicializado")
    End Sub
    
    #Region "Configurações da Madeireira"
    
    Public ReadOnly Property NomeMadeireira As String
        Get
            Return GetConfigValue("NomeMadeireira", "Madeireira Maria Luiza")
        End Get
    End Property
    
    Public ReadOnly Property EnderecoMadeireira As String
        Get
            Return GetConfigValue("EnderecoMadeireira", "Rua Principal, 123 - Centro")
        End Get
    End Property
    
    Public ReadOnly Property CidadeMadeireira As String
        Get
            Return GetConfigValue("CidadeMadeireira", "Paulista/PE")
        End Get
    End Property
    
    Public ReadOnly Property CEPMadeireira As String
        Get
            Return GetConfigValue("CEPMadeireira", "53401-445")
        End Get
    End Property
    
    Public ReadOnly Property TelefoneMadeireira As String
        Get
            Return GetConfigValue("TelefoneMadeireira", "(81) 3436-1234")
        End Get
    End Property
    
    Public ReadOnly Property CNPJMadeireira As String
        Get
            Return GetConfigValue("CNPJMadeireira", "12.345.678/0001-90")
        End Get
    End Property
    
    #End Region
    
    #Region "Configurações do Sistema"
    
    Public ReadOnly Property VendedorPadrao As String
        Get
            Return GetConfigValue("VendedorPadrao", "Sistema")
        End Get
    End Property
    
    Public ReadOnly Property ExcelVisivel As Boolean
        Get
            Return GetConfigValue("ExcelVisivel", False)
        End Get
    End Property
    
    Public ReadOnly Property SalvarTalaoTemporario As Boolean
        Get
            Return GetConfigValue("SalvarTalaoTemporario", False)
        End Get
    End Property
    
    Public ReadOnly Property UsarBancoAccess As Boolean
        Get
            Return GetConfigValue("UsarBancoAccess", False)
        End Get
    End Property
    
    Public ReadOnly Property CaminhoBackup As String
        Get
            Return GetConfigValue("CaminhoBackup", "C:\Backup\PDV\")
        End Get
    End Property
    
    Public ReadOnly Property ConexaoBanco As String
        Get
            Return GetConfigValue("ConexaoBanco", "")
        End Get
    End Property
    
    Public ReadOnly Property TimeoutExcel As Integer
        Get
            Return GetConfigValue("TimeoutExcel", 30000)
        End Get
    End Property
    
    Public ReadOnly Property LogLevel As String
        Get
            Return GetConfigValue("LogLevel", "Info")
        End Get
    End Property
    
    Public ReadOnly Property MaxLogFiles As Integer
        Get
            Return GetConfigValue("MaxLogFiles", 30)
        End Get
    End Property
    
    #End Region
    
    #Region "Configurações Avançadas"
    
    Public ReadOnly Property EnablePerformanceMonitoring As Boolean
        Get
            Return GetConfigValue("EnablePerformanceMonitoring", False)
        End Get
    End Property
    
    Public ReadOnly Property CacheEnabled As Boolean
        Get
            Return GetConfigValue("CacheEnabled", True)
        End Get
    End Property
    
    Public ReadOnly Property CacheExpirationMinutes As Integer
        Get
            Return GetConfigValue("CacheExpirationMinutes", 60)
        End Get
    End Property
    
    Public ReadOnly Property AutoBackupEnabled As Boolean
        Get
            Return GetConfigValue("AutoBackupEnabled", True)
        End Get
    End Property
    
    Public ReadOnly Property AutoBackupIntervalHours As Integer
        Get
            Return GetConfigValue("AutoBackupIntervalHours", 24)
        End Get
    End Property
    
    #End Region
    
    #Region "Métodos Internos"
    
    ''' <summary>
    ''' Obtém valor de configuração com cache e tipo específico
    ''' </summary>
    Private Function GetConfigValue(Of T)(key As String, defaultValue As T) As T
        Try
            ' Verificar cache primeiro
            If _configCache.ContainsKey(key) Then
                Return DirectCast(_configCache(key), T)
            End If
            
            ' Recarregar configurações se necessário
            If Date.Now.Subtract(_lastConfigCheck).TotalMinutes > 5 Then
                RefreshConfigurations()
            End If
            
            ' Buscar no app.config
            Dim configValue = System.Configuration.ConfigurationManager.AppSettings(key)
            
            If String.IsNullOrEmpty(configValue) Then
                _configCache(key) = defaultValue
                Return defaultValue
            End If
            
            ' Converter para o tipo apropriado
            Dim convertedValue As T = ConvertToType(Of T)(configValue, defaultValue)
            _configCache(key) = convertedValue
            
            Return convertedValue
            
        Catch ex As Exception
            _logger.LogWarning("ConfigurationManager", $"Erro ao obter configuração '{key}', usando valor padrão", ex)
            Return defaultValue
        End Try
    End Function
    
    ''' <summary>
    ''' Converte string para tipo específico
    ''' </summary>
    Private Function ConvertToType(Of T)(value As String, defaultValue As T) As T
        Try
            Dim targetType = GetType(T)
            
            If targetType Is GetType(String) Then
                Return DirectCast(DirectCast(value, Object), T)
            ElseIf targetType Is GetType(Boolean) Then
                Return DirectCast(DirectCast(Boolean.Parse(value), Object), T)
            ElseIf targetType Is GetType(Integer) Then
                Return DirectCast(DirectCast(Integer.Parse(value), Object), T)
            ElseIf targetType Is GetType(Decimal) Then
                Return DirectCast(DirectCast(Decimal.Parse(value), Object), T)
            ElseIf targetType Is GetType(Double) Then
                Return DirectCast(DirectCast(Double.Parse(value), Object), T)
            Else
                ' Para outros tipos, tentar conversão genérica
                Return DirectCast(Convert.ChangeType(value, targetType), T)
            End If
            
        Catch
            Return defaultValue
        End Try
    End Function
    
    ''' <summary>
    ''' Carrega todas as configurações no cache
    ''' </summary>
    Private Sub LoadConfigurations()
        Try
            SyncLock _lockObject
                _configCache.Clear()
                _lastConfigCheck = Date.Now
                
                ' Pré-carregar configurações críticas
                Dim criticalSettings = {
                    "NomeMadeireira", "VendedorPadrao", "ExcelVisivel",
                    "TimeoutExcel", "LogLevel", "CacheEnabled"
                }
                
                For Each setting In criticalSettings
                    ' Força carregamento no cache
                    GetConfigValue(setting, "")
                Next
                
                _logger.LogInfo("ConfigurationManager", $"Configurações carregadas: {_configCache.Count} itens")
            End SyncLock
            
        Catch ex As Exception
            _logger.LogError("ConfigurationManager", "Erro ao carregar configurações", ex)
        End Try
    End Sub
    
    ''' <summary>
    ''' Atualiza configurações se necessário
    ''' </summary>
    Private Sub RefreshConfigurations()
        Try
            SyncLock _lockObject
                ' Verificar se app.config foi modificado
                Dim configFile = AppDomain.CurrentDomain.SetupInformation.ConfigurationFile
                If File.Exists(configFile) Then
                    Dim lastWrite = File.GetLastWriteTime(configFile)
                    If lastWrite > _lastConfigCheck Then
                        _logger.LogInfo("ConfigurationManager", "Arquivo de configuração modificado, recarregando...")
                        LoadConfigurations()
                        Return
                    End If
                End If
                
                _lastConfigCheck = Date.Now
            End SyncLock
            
        Catch ex As Exception
            _logger.LogWarning("ConfigurationManager", "Erro ao verificar modificações de configuração", ex)
        End Try
    End Sub
    
    #End Region
    
    #Region "Métodos Públicos"
    
    ''' <summary>
    ''' Força recarregamento de todas as configurações
    ''' </summary>
    Public Sub ReloadConfigurations()
        _logger.LogInfo("ConfigurationManager", "Recarregamento manual de configurações solicitado")
        LoadConfigurations()
    End Sub
    
    ''' <summary>
    ''' Obtém todas as configurações para debug
    ''' </summary>
    Public Function GetAllConfigurations() As Dictionary(Of String, Object)
        SyncLock _lockObject
            Return New Dictionary(Of String, Object)(_configCache)
        End SyncLock
    End Function
    
    ''' <summary>
    ''' Valida configurações críticas
    ''' </summary>
    Public Function ValidateConfigurations() As ValidationResult
        Try
            Dim errors = New List(Of String)()
            
            ' Validar configurações obrigatórias
            If String.IsNullOrWhiteSpace(NomeMadeireira) Then
                errors.Add("Nome da madeireira não configurado")
            End If
            
            If TimeoutExcel <= 0 Then
                errors.Add("Timeout do Excel deve ser maior que zero")
            End If
            
            If CacheExpirationMinutes <= 0 Then
                errors.Add("Tempo de expiração do cache deve ser maior que zero")
            End If
            
            ' Validar caminhos
            Try
                If Not String.IsNullOrEmpty(CaminhoBackup) AndAlso Not Directory.Exists(Path.GetDirectoryName(CaminhoBackup)) Then
                    errors.Add($"Diretório pai do caminho de backup não existe: {CaminhoBackup}")
                End If
            Catch ex As Exception
                errors.Add($"Caminho de backup inválido: {ex.Message}")
            End Try
            
            If errors.Count > 0 Then
                Return New ValidationResult(False, String.Join("; ", errors))
            End If
            
            Return New ValidationResult(True, "Todas as configurações são válidas")
            
        Catch ex As Exception
            _logger.LogError("ConfigurationManager", "Erro ao validar configurações", ex)
            Return New ValidationResult(False, $"Erro na validação: {ex.Message}")
        End Try
    End Function
    
    ''' <summary>
    ''' Cria backup das configurações atuais
    ''' </summary>
    Public Function BackupConfigurations() As String
        Try
            Dim backupFile = Path.Combine(CaminhoBackup, $"config_backup_{Date.Now:yyyyMMdd_HHmmss}.txt")
            
            ' Criar diretório se não existir
            Directory.CreateDirectory(Path.GetDirectoryName(backupFile))
            
            Dim content = New List(Of String)()
            content.Add($"# Backup de Configurações - {Date.Now}")
            content.Add($"# Sistema PDV - Madeireira Maria Luiza")
            content.Add("")
            
            For Each kvp In GetAllConfigurations()
                content.Add($"{kvp.Key}={kvp.Value}")
            Next
            
            File.WriteAllLines(backupFile, content)
            
            _logger.LogInfo("ConfigurationManager", $"Backup de configurações criado: {backupFile}")
            Return backupFile
            
        Catch ex As Exception
            _logger.LogError("ConfigurationManager", "Erro ao criar backup de configurações", ex)
            Throw
        End Try
    End Function
    
    #End Region
    
    #Region "Métodos Públicos Adicionais"
    
    ''' <summary>
    ''' Obtém valor de configuração genérico com valor padrão (método público)
    ''' </summary>
    Public Function GetConfigValuePublic(Of T)(key As String, defaultValue As T) As T
        Return GetConfigValue(key, defaultValue)
    End Function
    
    #End Region
End Class