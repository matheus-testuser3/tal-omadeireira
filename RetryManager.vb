Imports System.Threading

''' <summary>
''' Sistema de retry para operações críticas
''' Implementa políticas de retry com backoff exponencial
''' </summary>
Public Class RetryManager
    
    Private Shared ReadOnly _logger As LoggingSystem = LoggingSystem.Instance
    
    ''' <summary>
    ''' Política de retry padrão
    ''' </summary>
    Public Shared ReadOnly Property DefaultPolicy As RetryPolicy
        Get
            Return New RetryPolicy() With {
                .MaxRetries = 3,
                .InitialDelay = TimeSpan.FromMilliseconds(500),
                .MaxDelay = TimeSpan.FromSeconds(10),
                .UseExponentialBackoff = True,
                .BackoffMultiplier = 2.0
            }
        End Get
    End Property
    
    ''' <summary>
    ''' Política para operações Excel
    ''' </summary>
    Public Shared ReadOnly Property ExcelPolicy As RetryPolicy
        Get
            Return New RetryPolicy() With {
                .MaxRetries = 5,
                .InitialDelay = TimeSpan.FromSeconds(1),
                .MaxDelay = TimeSpan.FromSeconds(30),
                .UseExponentialBackoff = True,
                .BackoffMultiplier = 1.5
            }
        End Get
    End Property
    
    ''' <summary>
    ''' Política para operações de rede
    ''' </summary>
    Public Shared ReadOnly Property NetworkPolicy As RetryPolicy
        Get
            Return New RetryPolicy() With {
                .MaxRetries = 3,
                .InitialDelay = TimeSpan.FromSeconds(2),
                .MaxDelay = TimeSpan.FromSeconds(15),
                .UseExponentialBackoff = True,
                .BackoffMultiplier = 2.0
            }
        End Get
    End Property
    
    ''' <summary>
    ''' Executa operação com retry
    ''' </summary>
    Public Shared Function Execute(Of T)(operation As Func(Of T), policy As RetryPolicy, operationName As String) As T
        Dim lastException As Exception = Nothing
        
        For attempt = 0 To policy.MaxRetries
            Try
                If attempt > 0 Then
                    _logger.LogInfo("RetryManager", $"Tentativa {attempt + 1}/{policy.MaxRetries + 1} para {operationName}")
                End If
                
                Return operation()
                
            Catch ex As Exception
                lastException = ex
                
                If attempt = policy.MaxRetries Then
                    _logger.LogError("RetryManager", $"Operação {operationName} falhou após {policy.MaxRetries + 1} tentativas", ex)
                    Throw New RetryException($"Operação {operationName} falhou após {policy.MaxRetries + 1} tentativas", ex)
                End If
                
                ' Verificar se é uma exceção que não deve ser tentada novamente
                If Not ShouldRetry(ex) Then
                    _logger.LogWarning("RetryManager", $"Operação {operationName} não será repetida devido ao tipo de erro: {ex.GetType().Name}")
                    Throw
                End If
                
                Dim delay = CalculateDelay(attempt, policy)
                _logger.LogWarning("RetryManager", $"Operação {operationName} falhou (tentativa {attempt + 1}), tentando novamente em {delay.TotalSeconds:F1}s: {ex.Message}")
                
                Thread.Sleep(delay)
            End Try
        Next
        
        ' Nunca deve chegar aqui, mas por segurança
        Throw lastException
    End Function
    
    ''' <summary>
    ''' Executa operação sem retorno com retry
    ''' </summary>
    Public Shared Sub Execute(operation As Action, policy As RetryPolicy, operationName As String)
        Execute(Function()
                    operation()
                    Return True
                End Function, policy, operationName)
    End Sub
    
    ''' <summary>
    ''' Executa operação async com retry
    ''' </summary>
    Public Shared Async Function ExecuteAsync(Of T)(operation As Func(Of Task(Of T)), policy As RetryPolicy, operationName As String) As Task(Of T)
        Dim lastException As Exception = Nothing
        
        For attempt = 0 To policy.MaxRetries
            Try
                If attempt > 0 Then
                    _logger.LogInfo("RetryManager", $"Tentativa async {attempt + 1}/{policy.MaxRetries + 1} para {operationName}")
                End If
                
                Return Await operation()
                
            Catch ex As Exception
                lastException = ex
                
                If attempt = policy.MaxRetries Then
                    _logger.LogError("RetryManager", $"Operação async {operationName} falhou após {policy.MaxRetries + 1} tentativas", ex)
                    Throw New RetryException($"Operação async {operationName} falhou após {policy.MaxRetries + 1} tentativas", ex)
                End If
                
                If Not ShouldRetry(ex) Then
                    _logger.LogWarning("RetryManager", $"Operação async {operationName} não será repetida devido ao tipo de erro: {ex.GetType().Name}")
                    Throw
                End If
                
                Dim delay = CalculateDelay(attempt, policy)
                _logger.LogWarning("RetryManager", $"Operação async {operationName} falhou (tentativa {attempt + 1}), tentando novamente em {delay.TotalSeconds:F1}s: {ex.Message}")
                
                Await Task.Delay(delay)
            End Try
        Next
        
        Throw lastException
    End Function
    
    ''' <summary>
    ''' Calcula delay baseado na política
    ''' </summary>
    Private Shared Function CalculateDelay(attempt As Integer, policy As RetryPolicy) As TimeSpan
        If Not policy.UseExponentialBackoff Then
            Return policy.InitialDelay
        End If
        
        Dim delay = TimeSpan.FromMilliseconds(policy.InitialDelay.TotalMilliseconds * Math.Pow(policy.BackoffMultiplier, attempt))
        
        ' Aplicar jitter para evitar thundering herd
        If policy.UseJitter Then
            Dim jitter = New Random().NextDouble() * 0.1 + 0.9 ' ±10% jitter
            delay = TimeSpan.FromMilliseconds(delay.TotalMilliseconds * jitter)
        End If
        
        ' Limitar ao delay máximo
        If delay > policy.MaxDelay Then
            delay = policy.MaxDelay
        End If
        
        Return delay
    End Function
    
    ''' <summary>
    ''' Determina se a exceção deve ser tentada novamente
    ''' </summary>
    Private Shared Function ShouldRetry(ex As Exception) As Boolean
        ' Não tentar novamente para ArgumentException, ArgumentNullException, etc.
        If TypeOf ex Is ArgumentException OrElse 
           TypeOf ex Is ArgumentNullException OrElse 
           TypeOf ex Is InvalidOperationException OrElse 
           TypeOf ex Is NotSupportedException Then
            Return False
        End If
        
        ' Não tentar novamente para SecurityException, UnauthorizedException
        If ex.GetType().Name.Contains("Security") OrElse 
           ex.GetType().Name.Contains("Unauthorized") Then
            Return False
        End If
        
        ' Por padrão, tentar novamente
        Return True
    End Function
End Class

''' <summary>
''' Política de retry
''' </summary>
Public Class RetryPolicy
    Public Property MaxRetries As Integer = 3
    Public Property InitialDelay As TimeSpan = TimeSpan.FromMilliseconds(500)
    Public Property MaxDelay As TimeSpan = TimeSpan.FromSeconds(10)
    Public Property UseExponentialBackoff As Boolean = True
    Public Property BackoffMultiplier As Double = 2.0
    Public Property UseJitter As Boolean = True
End Class

''' <summary>
''' Exceção específica para falhas de retry
''' </summary>
Public Class RetryException
    Inherits Exception
    
    Public Sub New(message As String)
        MyBase.New(message)
    End Sub
    
    Public Sub New(message As String, innerException As Exception)
        MyBase.New(message, innerException)
    End Sub
End Class

''' <summary>
''' Extensões para facilitar uso do RetryManager
''' </summary>
Public Module RetryExtensions
    
    ''' <summary>
    ''' Extensão para ExcelAutomation com retry
    ''' </summary>
    <Extension>
    Public Sub ProcessarTalaoCompletoComRetry(automation As ExcelAutomation, dados As DadosTalao)
        RetryManager.Execute(
            Sub() automation.ProcessarTalaoCompleto(dados),
            RetryManager.ExcelPolicy,
            "ProcessarTalaoCompleto"
        )
    End Sub
    
    ''' <summary>
    ''' Extensão para consulta de CEP com retry
    ''' </summary>
    <Extension>
    Public Function ConsultarCEPComRetry(apiManager As ExternalAPIManager, cep As String) As CEPResult
        Return RetryManager.Execute(
            Function() apiManager.ConsultarCEP(cep),
            RetryManager.NetworkPolicy,
            $"ConsultarCEP({cep})"
        )
    End Function
    
    ''' <summary>
    ''' Extensão para operações de banco com retry
    ''' </summary>
    <Extension>
    Public Function GetAllComRetry(Of T)(repository As IRepository(Of T)) As List(Of T)
        Return RetryManager.Execute(
            Function() repository.GetAll(),
            RetryManager.DefaultPolicy,
            $"Repository.GetAll<{GetType(T).Name}>"
        )
    End Function
    
    ''' <summary>
    ''' Extensão para salvar com retry
    ''' </summary>
    <Extension>
    Public Function AddComRetry(Of T)(repository As IRepository(Of T), entity As T) As Integer
        Return RetryManager.Execute(
            Function() repository.Add(entity),
            RetryManager.DefaultPolicy,
            $"Repository.Add<{GetType(T).Name}>"
        )
    End Function
End Module

''' <summary>
''' Circuit Breaker para prevenir cascata de falhas
''' </summary>
Public Class CircuitBreaker
    
    Private ReadOnly _maxFailures As Integer
    Private ReadOnly _timeout As TimeSpan
    Private ReadOnly _logger As LoggingSystem = LoggingSystem.Instance
    
    Private _failureCount As Integer = 0
    Private _lastFailureTime As Date = Date.MinValue
    Private _state As CircuitBreakerState = CircuitBreakerState.Closed
    
    Public Sub New(maxFailures As Integer, timeout As TimeSpan)
        _maxFailures = maxFailures
        _timeout = timeout
    End Sub
    
    ''' <summary>
    ''' Executa operação através do circuit breaker
    ''' </summary>
    Public Function Execute(Of T)(operation As Func(Of T), operationName As String) As T
        Select Case _state
            Case CircuitBreakerState.Open
                ' Circuit aberto - verificar se é hora de tentar novamente
                If Date.Now.Subtract(_lastFailureTime) > _timeout Then
                    _state = CircuitBreakerState.HalfOpen
                    _logger.LogInfo("CircuitBreaker", $"Circuit breaker {operationName} mudou para Half-Open")
                Else
                    Throw New CircuitBreakerOpenException($"Circuit breaker aberto para {operationName}")
                End If
                
            Case CircuitBreakerState.HalfOpen
                ' Estado de teste - uma falha vai abrir novamente
                
            Case CircuitBreakerState.Closed
                ' Normal
        End Select
        
        Try
            Dim result = operation()
            
            ' Sucesso - resetar contador e fechar circuit
            If _state = CircuitBreakerState.HalfOpen Then
                _logger.LogInfo("CircuitBreaker", $"Circuit breaker {operationName} fechado após sucesso")
            End If
            
            _failureCount = 0
            _state = CircuitBreakerState.Closed
            
            Return result
            
        Catch ex As Exception
            _failureCount += 1
            _lastFailureTime = Date.Now
            
            If _failureCount >= _maxFailures Then
                _state = CircuitBreakerState.Open
                _logger.LogWarning("CircuitBreaker", $"Circuit breaker {operationName} aberto após {_failureCount} falhas")
            End If
            
            Throw
        End Try
    End Function
    
    ''' <summary>
    ''' Estado atual do circuit breaker
    ''' </summary>
    Public ReadOnly Property State As CircuitBreakerState
        Get
            Return _state
        End Get
    End Property
    
    ''' <summary>
    ''' Contador de falhas atual
    ''' </summary>
    Public ReadOnly Property FailureCount As Integer
        Get
            Return _failureCount
        End Get
    End Property
End Class

''' <summary>
''' Estados do Circuit Breaker
''' </summary>
Public Enum CircuitBreakerState
    Closed    ' Funcionando normalmente
    Open      ' Circuito aberto, rejeitando chamadas
    HalfOpen  ' Testando se o serviço voltou
End Enum

''' <summary>
''' Exceção quando circuit breaker está aberto
''' </summary>
Public Class CircuitBreakerOpenException
    Inherits Exception
    
    Public Sub New(message As String)
        MyBase.New(message)
    End Sub
End Class