Imports System.IO
Imports System.Configuration

''' <summary>
''' Sistema de logging estruturado para o PDV
''' Gerencia logs de operações, erros e auditoria
''' </summary>
Public Class Logger
    Private Shared ReadOnly _instance As New Lazy(Of Logger)(Function() New Logger())
    Private ReadOnly _logPath As String
    Private ReadOnly _lockObject As New Object()
    
    ''' <summary>
    ''' Instância singleton do logger
    ''' </summary>
    Public Shared ReadOnly Property Instance As Logger
        Get
            Return _instance.Value
        End Get
    End Property
    
    ''' <summary>
    ''' Construtor privado para padrão singleton
    ''' </summary>
    Private Sub New()
        _logPath = Path.Combine(Application.StartupPath, "Logs")
        If Not Directory.Exists(_logPath) Then
            Directory.CreateDirectory(_logPath)
        End If
    End Sub
    
    ''' <summary>
    ''' Registra uma informação
    ''' </summary>
    Public Sub Info(message As String, Optional category As String = "INFO")
        WriteLog(LogLevel.Info, message, category)
    End Sub
    
    ''' <summary>
    ''' Registra um aviso
    ''' </summary>
    Public Sub Warning(message As String, Optional category As String = "WARNING")
        WriteLog(LogLevel.Warning, message, category)
    End Sub
    
    ''' <summary>
    ''' Registra um erro
    ''' </summary>
    Public Sub Error(message As String, Optional ex As Exception = Nothing, Optional category As String = "ERROR")
        Dim fullMessage = message
        If ex IsNot Nothing Then
            fullMessage &= $" | Exception: {ex.Message} | StackTrace: {ex.StackTrace}"
        End If
        WriteLog(LogLevel.Error, fullMessage, category)
    End Sub
    
    ''' <summary>
    ''' Registra um erro crítico
    ''' </summary>
    Public Sub Critical(message As String, Optional ex As Exception = Nothing, Optional category As String = "CRITICAL")
        Dim fullMessage = message
        If ex IsNot Nothing Then
            fullMessage &= $" | Exception: {ex.Message} | StackTrace: {ex.StackTrace}"
        End If
        WriteLog(LogLevel.Critical, fullMessage, category)
    End Sub
    
    ''' <summary>
    ''' Registra auditoria de operações
    ''' </summary>
    Public Sub Audit(operation As String, details As String, Optional user As String = "System")
        Dim message = $"User: {user} | Operation: {operation} | Details: {details}"
        WriteLog(LogLevel.Audit, message, "AUDIT")
    End Sub
    
    ''' <summary>
    ''' Escreve log no arquivo
    ''' </summary>
    Private Sub WriteLog(level As LogLevel, message As String, category As String)
        Try
            SyncLock _lockObject
                Dim fileName = $"PDV_{DateTime.Now:yyyyMMdd}.log"
                Dim filePath = Path.Combine(_logPath, fileName)
                
                Dim logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} | {level.ToString().ToUpper()} | {category} | {message}{Environment.NewLine}"
                
                File.AppendAllText(filePath, logEntry)
                
                ' Manter apenas logs dos últimos 30 dias
                CleanupOldLogs()
            End SyncLock
        Catch ex As Exception
            ' Evitar loop infinito de erro no logger
            Console.WriteLine($"Erro no logger: {ex.Message}")
        End Try
    End Sub
    
    ''' <summary>
    ''' Remove logs antigos (mais de 30 dias)
    ''' </summary>
    Private Sub CleanupOldLogs()
        Try
            Dim files = Directory.GetFiles(_logPath, "PDV_*.log")
            Dim cutoffDate = DateTime.Now.AddDays(-30)
            
            For Each file In files
                Dim fileInfo = New FileInfo(file)
                If fileInfo.CreationTime < cutoffDate Then
                    File.Delete(file)
                End If
            Next
        Catch
            ' Ignorar erros de limpeza
        End Try
    End Sub
    
    ''' <summary>
    ''' Obtém logs do dia atual
    ''' </summary>
    Public Function GetTodayLogs() As String
        Try
            Dim fileName = $"PDV_{DateTime.Now:yyyyMMdd}.log"
            Dim filePath = Path.Combine(_logPath, fileName)
            
            If File.Exists(filePath) Then
                Return File.ReadAllText(filePath)
            End If
        Catch ex As Exception
            Error("Erro ao ler logs", ex)
        End Try
        
        Return String.Empty
    End Function
End Class

''' <summary>
''' Níveis de log
''' </summary>
Public Enum LogLevel
    Info = 1
    Warning = 2
    Error = 3
    Critical = 4
    Audit = 5
End Enum