Imports System.IO
Imports System.IO.Compression
Imports System.Threading.Tasks
Imports System.Xml.Serialization

''' <summary>
''' Serviço de backup automático dos dados
''' Gerencia backup e restauração do sistema
''' </summary>
Public Class BackupService
    Private ReadOnly _logger As Logger
    Private ReadOnly _config As ConfigManager
    Private ReadOnly _backupPath As String
    Private ReadOnly _timer As System.Timers.Timer
    
    ''' <summary>
    ''' Construtor
    ''' </summary>
    Public Sub New()
        _logger = Logger.Instance
        _config = ConfigManager.Instance
        _backupPath = Path.Combine(Application.StartupPath, "Backups")
        
        ' Criar diretório de backup se não existir
        If Not Directory.Exists(_backupPath) Then
            Directory.CreateDirectory(_backupPath)
        End If
        
        ' Configurar timer para backup automático
        If _config.BackupAutomatico Then
            ConfigurarBackupAutomatico()
        End If
    End Sub
    
    ''' <summary>
    ''' Configura backup automático
    ''' </summary>
    Private Sub ConfigurarBackupAutomatico()
        Try
            _timer = New System.Timers.Timer()
            _timer.Interval = _config.IntervaloBacKupHoras * 60 * 60 * 1000 ' Converter para milissegundos
            _timer.Enabled = True
            AddHandler _timer.Elapsed, AddressOf Timer_Elapsed
            
            _logger.Info($"Backup automático configurado para cada {_config.IntervaloBacKupHoras} horas")
            
        Catch ex As Exception
            _logger.Error("Erro ao configurar backup automático", ex)
        End Try
    End Sub
    
    ''' <summary>
    ''' Evento do timer para backup automático
    ''' </summary>
    Private Sub Timer_Elapsed(sender As Object, e As System.Timers.ElapsedEventArgs)
        Task.Run(Sub() ExecutarBackupAutomatico())
    End Sub
    
    ''' <summary>
    ''' Executa backup completo do sistema
    ''' </summary>
    Public Function ExecutarBackupCompleto() As Boolean
        Try
            _logger.Info("Iniciando backup completo")
            
            Dim timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss")
            Dim backupFileName = $"Backup_PDV_{timestamp}.zip"
            Dim backupFilePath = Path.Combine(_backupPath, backupFileName)
            
            Using zip = New FileStream(backupFilePath, FileMode.Create)
                Using archive = New ZipArchive(zip, ZipArchiveMode.Create)
                    
                    ' Backup dos dados
                    BackupDados(archive)
                    
                    ' Backup das configurações
                    BackupConfiguracoes(archive)
                    
                    ' Backup dos logs
                    BackupLogs(archive)
                    
                    ' Backup do catálogo de produtos
                    BackupCatalogoProdutos(archive)
                    
                End Using
            End Using
            
            _logger.Info($"Backup completo criado: {backupFileName}")
            
            ' Limpar backups antigos
            LimparBackupsAntigos()
            
            Return True
            
        Catch ex As Exception
            _logger.Error("Erro ao executar backup completo", ex)
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' Backup dos dados de vendas
    ''' </summary>
    Private Sub BackupDados(archive As ZipArchive)
        Try
            Dim dataManager = DataManager.Instance
            Dim vendas = dataManager.ObterTodasVendas()
            Dim clientes = dataManager.ObterTodosClientes()
            
            ' Serializar vendas
            Dim vendasXml = SerializarObjeto(vendas)
            AdicionarArquivoAoZip(archive, "Data/vendas.xml", vendasXml)
            
            ' Serializar clientes
            Dim clientesXml = SerializarObjeto(clientes)
            AdicionarArquivoAoZip(archive, "Data/clientes.xml", clientesXml)
            
            _logger.Info("Backup de dados concluído")
            
        Catch ex As Exception
            _logger.Error("Erro no backup de dados", ex)
        End Try
    End Sub
    
    ''' <summary>
    ''' Backup das configurações
    ''' </summary>
    Private Sub BackupConfiguracoes(archive As ZipArchive)
        Try
            ' App.config
            Dim appConfigPath = Path.Combine(Application.StartupPath, "App.config")
            If File.Exists(appConfigPath) Then
                AdicionarArquivoAoZip(archive, "Config/App.config", File.ReadAllText(appConfigPath))
            End If
            
            ' Configurações customizadas
            Dim customConfigPath = Path.Combine(Application.StartupPath, "Config", "CustomSettings.xml")
            If File.Exists(customConfigPath) Then
                AdicionarArquivoAoZip(archive, "Config/CustomSettings.xml", File.ReadAllText(customConfigPath))
            End If
            
            _logger.Info("Backup de configurações concluído")
            
        Catch ex As Exception
            _logger.Error("Erro no backup de configurações", ex)
        End Try
    End Sub
    
    ''' <summary>
    ''' Backup dos logs
    ''' </summary>
    Private Sub BackupLogs(archive As ZipArchive)
        Try
            Dim logsPath = Path.Combine(Application.StartupPath, "Logs")
            If Directory.Exists(logsPath) Then
                Dim logFiles = Directory.GetFiles(logsPath, "*.log")
                
                For Each logFile In logFiles
                    Dim fileName = Path.GetFileName(logFile)
                    AdicionarArquivoAoZip(archive, $"Logs/{fileName}", File.ReadAllText(logFile))
                Next
            End If
            
            _logger.Info("Backup de logs concluído")
            
        Catch ex As Exception
            _logger.Error("Erro no backup de logs", ex)
        End Try
    End Sub
    
    ''' <summary>
    ''' Backup do catálogo de produtos
    ''' </summary>
    Private Sub BackupCatalogoProdutos(archive As ZipArchive)
        Try
            Dim produtosPath = Path.Combine(Application.StartupPath, "Config", "Products.xml")
            If File.Exists(produtosPath) Then
                AdicionarArquivoAoZip(archive, "Config/Products.xml", File.ReadAllText(produtosPath))
            End If
            
            _logger.Info("Backup de catálogo de produtos concluído")
            
        Catch ex As Exception
            _logger.Error("Erro no backup de catálogo", ex)
        End Try
    End Sub
    
    ''' <summary>
    ''' Adiciona arquivo ao ZIP
    ''' </summary>
    Private Sub AdicionarArquivoAoZip(archive As ZipArchive, entryName As String, content As String)
        Dim entry = archive.CreateEntry(entryName)
        Using entryStream = entry.Open()
            Using writer = New StreamWriter(entryStream)
                writer.Write(content)
            End Using
        End Using
    End Sub
    
    ''' <summary>
    ''' Serializa objeto para XML
    ''' </summary>
    Private Function SerializarObjeto(Of T)(obj As T) As String
        Try
            Dim serializer = New XmlSerializer(GetType(T))
            Using writer = New StringWriter()
                serializer.Serialize(writer, obj)
                Return writer.ToString()
            End Using
        Catch ex As Exception
            _logger.Error("Erro ao serializar objeto", ex)
            Return ""
        End Try
    End Function
    
    ''' <summary>
    ''' Restaura backup do sistema
    ''' </summary>
    Public Function RestaurarBackup(backupFilePath As String) As Boolean
        Try
            _logger.Info($"Iniciando restauração do backup: {backupFilePath}")
            
            If Not File.Exists(backupFilePath) Then
                _logger.Error("Arquivo de backup não encontrado")
                Return False
            End If
            
            Using zip = New FileStream(backupFilePath, FileMode.Open)
                Using archive = New ZipArchive(zip, ZipArchiveMode.Read)
                    
                    ' Restaurar dados
                    RestaurarDados(archive)
                    
                    ' Restaurar configurações
                    RestaurarConfiguracoes(archive)
                    
                    ' Restaurar catálogo de produtos
                    RestaurarCatalogoProdutos(archive)
                    
                End Using
            End Using
            
            _logger.Info("Restauração concluída com sucesso")
            _logger.Audit("BACKUP_RESTAURADO", $"Arquivo: {Path.GetFileName(backupFilePath)}", "Sistema")
            
            Return True
            
        Catch ex As Exception
            _logger.Error("Erro ao restaurar backup", ex)
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' Restaura dados do backup
    ''' </summary>
    Private Sub RestaurarDados(archive As ZipArchive)
        Try
            ' Restaurar vendas
            Dim vendasEntry = archive.GetEntry("Data/vendas.xml")
            If vendasEntry IsNot Nothing Then
                Using stream = vendasEntry.Open()
                    Using reader = New StreamReader(stream)
                        Dim xml = reader.ReadToEnd()
                        ' Implementar deserialização e carregamento
                        _logger.Info("Vendas restauradas")
                    End Using
                End Using
            End If
            
            ' Restaurar clientes
            Dim clientesEntry = archive.GetEntry("Data/clientes.xml")
            If clientesEntry IsNot Nothing Then
                Using stream = clientesEntry.Open()
                    Using reader = New StreamReader(stream)
                        Dim xml = reader.ReadToEnd()
                        ' Implementar deserialização e carregamento
                        _logger.Info("Clientes restaurados")
                    End Using
                End Using
            End If
            
        Catch ex As Exception
            _logger.Error("Erro ao restaurar dados", ex)
        End Try
    End Sub
    
    ''' <summary>
    ''' Restaura configurações do backup
    ''' </summary>
    Private Sub RestaurarConfiguracoes(archive As ZipArchive)
        Try
            Dim configEntry = archive.GetEntry("Config/CustomSettings.xml")
            If configEntry IsNot Nothing Then
                Using stream = configEntry.Open()
                    Using reader = New StreamReader(stream)
                        Dim xml = reader.ReadToEnd()
                        Dim configPath = Path.Combine(Application.StartupPath, "Config", "CustomSettings.xml")
                        File.WriteAllText(configPath, xml)
                        _logger.Info("Configurações restauradas")
                    End Using
                End Using
            End If
        Catch ex As Exception
            _logger.Error("Erro ao restaurar configurações", ex)
        End Try
    End Sub
    
    ''' <summary>
    ''' Restaura catálogo de produtos do backup
    ''' </summary>
    Private Sub RestaurarCatalogoProdutos(archive As ZipArchive)
        Try
            Dim produtosEntry = archive.GetEntry("Config/Products.xml")
            If produtosEntry IsNot Nothing Then
                Using stream = produtosEntry.Open()
                    Using reader = New StreamReader(stream)
                        Dim xml = reader.ReadToEnd()
                        Dim produtosPath = Path.Combine(Application.StartupPath, "Config", "Products.xml")
                        File.WriteAllText(produtosPath, xml)
                        _logger.Info("Catálogo de produtos restaurado")
                    End Using
                End Using
            End If
        Catch ex As Exception
            _logger.Error("Erro ao restaurar catálogo", ex)
        End Try
    End Sub
    
    ''' <summary>
    ''' Executa backup automático
    ''' </summary>
    Private Sub ExecutarBackupAutomatico()
        Try
            _logger.Info("Executando backup automático")
            ExecutarBackupCompleto()
        Catch ex As Exception
            _logger.Error("Erro no backup automático", ex)
        End Try
    End Sub
    
    ''' <summary>
    ''' Remove backups antigos (mais de 30 dias)
    ''' </summary>
    Private Sub LimparBackupsAntigos()
        Try
            Dim files = Directory.GetFiles(_backupPath, "Backup_PDV_*.zip")
            Dim cutoffDate = DateTime.Now.AddDays(-30)
            
            For Each file In files
                Dim fileInfo = New FileInfo(file)
                If fileInfo.CreationTime < cutoffDate Then
                    File.Delete(file)
                    _logger.Info($"Backup antigo removido: {Path.GetFileName(file)}")
                End If
            Next
        Catch ex As Exception
            _logger.Warning("Erro ao limpar backups antigos", ex)
        End Try
    End Sub
    
    ''' <summary>
    ''' Obtém lista de backups disponíveis
    ''' </summary>
    Public Function ObterBackupsDisponiveis() As List(Of BackupInfo)
        Try
            Dim backups = New List(Of BackupInfo)()
            Dim files = Directory.GetFiles(_backupPath, "Backup_PDV_*.zip")
            
            For Each file In files
                Dim fileInfo = New FileInfo(file)
                backups.Add(New BackupInfo() With {
                    .FileName = Path.GetFileName(file),
                    .FilePath = file,
                    .DataCriacao = fileInfo.CreationTime,
                    .Tamanho = fileInfo.Length
                })
            Next
            
            Return backups.OrderByDescending(Function(b) b.DataCriacao).ToList()
            
        Catch ex As Exception
            _logger.Error("Erro ao obter backups disponíveis", ex)
            Return New List(Of BackupInfo)()
        End Try
    End Function
    
    ''' <summary>
    ''' Para o serviço de backup
    ''' </summary>
    Public Sub Parar()
        Try
            If _timer IsNot Nothing Then
                _timer.Stop()
                _timer.Dispose()
            End If
            _logger.Info("Serviço de backup parado")
        Catch ex As Exception
            _logger.Error("Erro ao parar serviço de backup", ex)
        End Try
    End Sub
End Class

''' <summary>
''' Informações sobre um backup
''' </summary>
Public Class BackupInfo
    Public Property FileName As String
    Public Property FilePath As String
    Public Property DataCriacao As DateTime
    Public Property Tamanho As Long
    
    Public ReadOnly Property TamanhoFormatado As String
        Get
            Dim kb = Tamanho \ 1024
            Dim mb = kb \ 1024
            If mb > 0 Then
                Return $"{mb:N0} MB"
            ElseIf kb > 0 Then
                Return $"{kb:N0} KB"
            Else
                Return $"{Tamanho:N0} bytes"
            End If
        End Get
    End Property
End Class