Imports System.Configuration
Imports System.IO
Imports System.Xml.Serialization

''' <summary>
''' Gerenciador centralizado de configurações do sistema
''' Gerencia configurações do App.config e configurações customizadas
''' </summary>
Public Class ConfigManager
    Private Shared ReadOnly _instance As New Lazy(Of ConfigManager)(Function() New ConfigManager())
    Private ReadOnly _customConfigPath As String
    Private _customSettings As Dictionary(Of String, String)
    
    ''' <summary>
    ''' Instância singleton do gerenciador de configurações
    ''' </summary>
    Public Shared ReadOnly Property Instance As ConfigManager
        Get
            Return _instance.Value
        End Get
    End Property
    
    ''' <summary>
    ''' Construtor privado
    ''' </summary>
    Private Sub New()
        _customConfigPath = Path.Combine(Application.StartupPath, "Config", "CustomSettings.xml")
        _customSettings = New Dictionary(Of String, String)()
        LoadCustomSettings()
    End Sub
    
    #Region "Configurações da Madeireira"
    
    Public ReadOnly Property NomeMadeireira As String
        Get
            Return GetAppSetting("NomeMadeireira", "Madeireira Maria Luiza")
        End Get
    End Property
    
    Public ReadOnly Property EnderecoMadeireira As String
        Get
            Return GetAppSetting("EnderecoMadeireira", "Rua Principal, 123 - Centro")
        End Get
    End Property
    
    Public ReadOnly Property CidadeMadeireira As String
        Get
            Return GetAppSetting("CidadeMadeireira", "Paulista/PE")
        End Get
    End Property
    
    Public ReadOnly Property CEPMadeireira As String
        Get
            Return GetAppSetting("CEPMadeireira", "53401-445")
        End Get
    End Property
    
    Public ReadOnly Property TelefoneMadeireira As String
        Get
            Return GetAppSetting("TelefoneMadeireira", "(81) 3436-1234")
        End Get
    End Property
    
    Public ReadOnly Property CNPJMadeireira As String
        Get
            Return GetAppSetting("CNPJMadeireira", "12.345.678/0001-90")
        End Get
    End Property
    
    #End Region
    
    #Region "Configurações do Sistema"
    
    Public ReadOnly Property VendedorPadrao As String
        Get
            Return GetAppSetting("VendedorPadrao", "Sistema")
        End Get
    End Property
    
    Public ReadOnly Property ExcelVisivel As Boolean
        Get
            Return Boolean.Parse(GetAppSetting("ExcelVisivel", "false"))
        End Get
    End Property
    
    Public ReadOnly Property SalvarTalaoTemporario As Boolean
        Get
            Return Boolean.Parse(GetAppSetting("SalvarTalaoTemporario", "false"))
        End Get
    End Property
    
    Public ReadOnly Property BackupAutomatico As Boolean
        Get
            Return Boolean.Parse(GetCustomSetting("BackupAutomatico", "true"))
        End Get
    End Property
    
    Public ReadOnly Property IntervaloBacKupHoras As Integer
        Get
            Return Integer.Parse(GetCustomSetting("IntervaloBacKupHoras", "24"))
        End Get
    End Property
    
    Public ReadOnly Property ManterHistoricoDias As Integer
        Get
            Return Integer.Parse(GetCustomSetting("ManterHistoricoDias", "365"))
        End Get
    End Property
    
    #End Region
    
    #Region "Configurações de Performance"
    
    Public ReadOnly Property TimeoutExcelSegundos As Integer
        Get
            Return Integer.Parse(GetCustomSetting("TimeoutExcelSegundos", "30"))
        End Get
    End Property
    
    Public ReadOnly Property CacheSize As Integer
        Get
            Return Integer.Parse(GetCustomSetting("CacheSize", "100"))
        End Get
    End Property
    
    Public ReadOnly Property LogLevel As String
        Get
            Return GetCustomSetting("LogLevel", "INFO")
        End Get
    End Property
    
    #End Region
    
    ''' <summary>
    ''' Obtém configuração do App.config
    ''' </summary>
    Private Function GetAppSetting(key As String, defaultValue As String) As String
        Try
            Dim value = ConfigurationManager.AppSettings(key)
            Return If(String.IsNullOrWhiteSpace(value), defaultValue, value)
        Catch
            Logger.Instance.Warning($"Erro ao ler configuração {key}, usando valor padrão")
            Return defaultValue
        End Try
    End Function
    
    ''' <summary>
    ''' Obtém configuração customizada
    ''' </summary>
    Private Function GetCustomSetting(key As String, defaultValue As String) As String
        If _customSettings.ContainsKey(key) Then
            Return _customSettings(key)
        End If
        Return defaultValue
    End Function
    
    ''' <summary>
    ''' Define configuração customizada
    ''' </summary>
    Public Sub SetCustomSetting(key As String, value As String)
        _customSettings(key) = value
        SaveCustomSettings()
    End Sub
    
    ''' <summary>
    ''' Carrega configurações customizadas do arquivo XML
    ''' </summary>
    Private Sub LoadCustomSettings()
        Try
            If File.Exists(_customConfigPath) Then
                Dim serializer = New XmlSerializer(GetType(Dictionary(Of String, String)))
                Using reader = New FileStream(_customConfigPath, FileMode.Open)
                    _customSettings = CType(serializer.Deserialize(reader), Dictionary(Of String, String))
                End Using
            End If
        Catch ex As Exception
            Logger.Instance.Error("Erro ao carregar configurações customizadas", ex)
            _customSettings = New Dictionary(Of String, String)()
        End Try
    End Sub
    
    ''' <summary>
    ''' Salva configurações customizadas no arquivo XML
    ''' </summary>
    Private Sub SaveCustomSettings()
        Try
            Dim directory = Path.GetDirectoryName(_customConfigPath)
            If Not Directory.Exists(directory) Then
                Directory.CreateDirectory(directory)
            End If
            
            Dim serializer = New XmlSerializer(GetType(Dictionary(Of String, String)))
            Using writer = New FileStream(_customConfigPath, FileMode.Create)
                serializer.Serialize(writer, _customSettings)
            End Using
        Catch ex As Exception
            Logger.Instance.Error("Erro ao salvar configurações customizadas", ex)
        End Try
    End Sub
    
    ''' <summary>
    ''' Recarrega todas as configurações
    ''' </summary>
    Public Sub ReloadSettings()
        ConfigurationManager.RefreshSection("appSettings")
        LoadCustomSettings()
    End Sub
    
    ''' <summary>
    ''' Obtém todas as configurações como dicionário
    ''' </summary>
    Public Function GetAllSettings() As Dictionary(Of String, String)
        Dim settings = New Dictionary(Of String, String)()
        
        ' Adicionar configurações do App.config
        For Each key As String In ConfigurationManager.AppSettings.AllKeys
            settings(key) = ConfigurationManager.AppSettings(key)
        Next
        
        ' Adicionar configurações customizadas
        For Each kvp In _customSettings
            settings($"Custom_{kvp.Key}") = kvp.Value
        Next
        
        Return settings
    End Function
End Class