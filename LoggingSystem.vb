Imports System.IO
Imports System.Configuration

''' <summary>
''' Sistema de logging centralizado para o Sistema PDV
''' Gerencia logs de operações, erros e eventos do sistema
''' </summary>
Public Class LoggingSystem
    
    Private Shared _instance As LoggingSystem
    Private ReadOnly _logDirectory As String
    Private ReadOnly _logFileName As String
    Private ReadOnly _lockObject As New Object()
    
    ' Níveis de log
    Public Enum LogLevel
        Info
        Warning
        [Error]
        Debug
        Critical
    End Enum
    
    ''' <summary>
    ''' Singleton instance do sistema de logging
    ''' </summary>
    Public Shared ReadOnly Property Instance As LoggingSystem
        Get
            If _instance Is Nothing Then
                _instance = New LoggingSystem()
            End If
            Return _instance
        End Get
    End Property
    
    ''' <summary>
    ''' Construtor privado para padrão Singleton
    ''' </summary>
    Private Sub New()
        Try
            ' Configurar diretório de logs
            _logDirectory = Path.Combine(
                If(ConfigurationManager.AppSettings("CaminhoBackup"), Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)),
                "PDV_Logs"
            )
            
            ' Criar diretório se não existir
            If Not Directory.Exists(_logDirectory) Then
                Directory.CreateDirectory(_logDirectory)
            End If
            
            ' Nome do arquivo de log com data
            _logFileName = Path.Combine(_logDirectory, $"PDV_Log_{Date.Now:yyyy-MM-dd}.log")
            
            ' Log inicial do sistema
            WriteLog(LogLevel.Info, "Sistema de Logging", "Sistema de logging inicializado com sucesso")
            
        Catch ex As Exception
            ' Em caso de erro, usar diretório temporário
            _logDirectory = Path.GetTempPath()
            _logFileName = Path.Combine(_logDirectory, $"PDV_Log_{Date.Now:yyyy-MM-dd}.log")
        End Try
    End Sub
    
    ''' <summary>
    ''' Escreve uma entrada de log
    ''' </summary>
    Public Sub WriteLog(level As LogLevel, modulo As String, mensagem As String, Optional ex As Exception = Nothing)
        Try
            SyncLock _lockObject
                Dim logEntry = BuildLogEntry(level, modulo, mensagem, ex)
                
                ' Escrever no arquivo
                File.AppendAllText(_logFileName, logEntry & Environment.NewLine)
                
                ' Também escrever no console para debug
                If level >= LogLevel.Warning Then
                    Console.WriteLine(logEntry)
                End If
                
                ' Limpar logs antigos (manter apenas 30 dias)
                CleanOldLogs()
            End SyncLock
            
        Catch logEx As Exception
            ' Em caso de falha no logging, pelo menos tentar console
            Console.WriteLine($"[LOGGING ERROR] {logEx.Message}")
            Console.WriteLine($"[ORIGINAL] {level} - {modulo}: {mensagem}")
        End Try
    End Sub
    
    ''' <summary>
    ''' Log de informação
    ''' </summary>
    Public Sub LogInfo(modulo As String, mensagem As String)
        WriteLog(LogLevel.Info, modulo, mensagem)
    End Sub
    
    ''' <summary>
    ''' Log de aviso
    ''' </summary>
    Public Sub LogWarning(modulo As String, mensagem As String)
        WriteLog(LogLevel.Warning, modulo, mensagem)
    End Sub
    
    ''' <summary>
    ''' Log de erro
    ''' </summary>
    Public Sub LogError(modulo As String, mensagem As String, Optional ex As Exception = Nothing)
        WriteLog(LogLevel.Error, modulo, mensagem, ex)
    End Sub
    
    ''' <summary>
    ''' Log de debug
    ''' </summary>
    Public Sub LogDebug(modulo As String, mensagem As String)
        WriteLog(LogLevel.Debug, modulo, mensagem)
    End Sub
    
    ''' <summary>
    ''' Log crítico
    ''' </summary>
    Public Sub LogCritical(modulo As String, mensagem As String, Optional ex As Exception = Nothing)
        WriteLog(LogLevel.Critical, modulo, mensagem, ex)
    End Sub
    
    ''' <summary>
    ''' Constrói entrada formatada de log
    ''' </summary>
    Private Function BuildLogEntry(level As LogLevel, modulo As String, mensagem As String, ex As Exception) As String
        Dim timestamp = Date.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")
        Dim levelStr = level.ToString().ToUpper().PadRight(8)
        Dim moduloStr = modulo.PadRight(20)
        
        Dim entry = $"[{timestamp}] [{levelStr}] [{moduloStr}] {mensagem}"
        
        ' Adicionar detalhes da exceção se presente
        If ex IsNot Nothing Then
            entry &= Environment.NewLine & $"    Exception: {ex.GetType().Name}: {ex.Message}"
            If Not String.IsNullOrEmpty(ex.StackTrace) Then
                entry &= Environment.NewLine & $"    StackTrace: {ex.StackTrace}"
            End If
            If ex.InnerException IsNot Nothing Then
                entry &= Environment.NewLine & $"    InnerException: {ex.InnerException.Message}"
            End If
        End If
        
        Return entry
    End Function
    
    ''' <summary>
    ''' Remove logs antigos para evitar acúmulo excessivo
    ''' </summary>
    Private Sub CleanOldLogs()
        Try
            ' Executar limpeza apenas uma vez por dia
            Dim lastCleanFile = Path.Combine(_logDirectory, "last_clean.txt")
            Dim lastClean = Date.MinValue
            
            If File.Exists(lastCleanFile) Then
                Date.TryParse(File.ReadAllText(lastCleanFile), lastClean)
            End If
            
            If Date.Now.Date > lastClean.Date Then
                Dim cutoffDate = Date.Now.AddDays(-30)
                Dim logFiles = Directory.GetFiles(_logDirectory, "PDV_Log_*.log")
                
                For Each logFile In logFiles
                    Dim fileInfo = New FileInfo(logFile)
                    If fileInfo.CreationTime < cutoffDate Then
                        Try
                            File.Delete(logFile)
                        Catch
                            ' Ignorar erro se não conseguir deletar
                        End Try
                    End If
                Next
                
                ' Marcar limpeza realizada
                File.WriteAllText(lastCleanFile, Date.Now.ToString())
            End If
            
        Catch
            ' Ignorar erros na limpeza
        End Try
    End Sub
    
    ''' <summary>
    ''' Obtém estatísticas de logs do dia atual
    ''' </summary>
    Public Function GetTodayStats() As Dictionary(Of LogLevel, Integer)
        Dim stats = New Dictionary(Of LogLevel, Integer)()
        
        Try
            If File.Exists(_logFileName) Then
                Dim lines = File.ReadAllLines(_logFileName)
                
                For Each level As LogLevel In [Enum].GetValues(GetType(LogLevel))
                    stats(level) = lines.Count(Function(line) line.Contains($"[{level.ToString().ToUpper().PadRight(8)}]"))
                Next
            End If
        Catch
            ' Em caso de erro, retornar estatísticas vazias
        End Try
        
        Return stats
    End Function
    
    ''' <summary>
    ''' Obtém últimas N entradas de log
    ''' </summary>
    Public Function GetRecentLogs(count As Integer) As List(Of String)
        Dim recentLogs = New List(Of String)()
        
        Try
            If File.Exists(_logFileName) Then
                Dim lines = File.ReadAllLines(_logFileName)
                Dim startIndex = Math.Max(0, lines.Length - count)
                
                For i = startIndex To lines.Length - 1
                    recentLogs.Add(lines(i))
                Next
            End If
        Catch
            ' Em caso de erro, retornar lista vazia
        End Try
        
        Return recentLogs
    End Function
End Class