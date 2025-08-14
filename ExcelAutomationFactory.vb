Imports System.Management

''' <summary>
''' Factory para criação de componentes de automação Excel
''' Implementa padrão Factory Method para melhor organização e testabilidade
''' </summary>
Public Class ExcelAutomationFactory
    
    Private Shared ReadOnly _logger As LoggingSystem = LoggingSystem.Instance
    
    ''' <summary>
    ''' Tipos de automação Excel disponíveis
    ''' </summary>
    Public Enum AutomationType
        StandardPDV
        ReportGenerator
        TemplateProcessor
        BatchProcessor
    End Enum
    
    ''' <summary>
    ''' Cria instância de automação Excel baseada no tipo
    ''' </summary>
    Public Shared Function CreateAutomation(tipo As AutomationType) As IExcelAutomation
        _logger.LogInfo("ExcelAutomationFactory", $"Criando automação tipo: {tipo}")
        
        Try
            Select Case tipo
                Case AutomationType.StandardPDV
                    Return New StandardPDVAutomation()
                Case AutomationType.ReportGenerator
                    Return New ReportGeneratorAutomation()
                Case AutomationType.TemplateProcessor
                    Return New TemplateProcessorAutomation()
                Case AutomationType.BatchProcessor
                    Return New BatchProcessorAutomation()
                Case Else
                    Throw New ArgumentException($"Tipo de automação não suportado: {tipo}")
            End Select
            
        Catch ex As Exception
            _logger.LogError("ExcelAutomationFactory", "Erro ao criar automação", ex)
            Throw
        End Try
    End Function
    
    ''' <summary>
    ''' Valida pré-requisitos para automação Excel
    ''' </summary>
    Public Shared Function ValidatePrerequisites() As ValidationResult
        Try
            ' Verificar instalação do Excel
            If Not IsExcelInstalled() Then
                Return New ValidationResult(False, "Microsoft Excel não está instalado")
            End If
            
            ' Verificar versão do Excel
            Dim version = GetExcelVersion()
            If version < New Version("12.0") Then ' Excel 2007+
                Return New ValidationResult(False, "Versão do Excel muito antiga. Mínimo: Excel 2007")
            End If
            
            ' Verificar permissões VBA
            If Not CanAccessVBA() Then
                Return New ValidationResult(False, "Acesso VBA não configurado. Habilite 'Confiar no acesso ao projeto VBA'")
            End If
            
            ' Verificar .NET Framework
            If Not IsDotNetFrameworkAvailable() Then
                Return New ValidationResult(False, ".NET Framework 4.7.2+ não encontrado")
            End If
            
            Return New ValidationResult(True, "Todos os pré-requisitos atendidos")
            
        Catch ex As Exception
            _logger.LogError("ExcelAutomationFactory", "Erro ao validar pré-requisitos", ex)
            Return New ValidationResult(False, $"Erro na validação: {ex.Message}")
        End Try
    End Function
    
    ''' <summary>
    ''' Verifica se Excel está instalado
    ''' </summary>
    Private Shared Function IsExcelInstalled() As Boolean
        Try
            Dim excelType = Type.GetTypeFromProgID("Excel.Application")
            Return excelType IsNot Nothing
        Catch
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' Obtém versão do Excel instalado
    ''' </summary>
    Private Shared Function GetExcelVersion() As Version
        Try
            Using excelApp = CreateObject("Excel.Application")
                Dim versionString = excelApp.Version.ToString()
                Return New Version(versionString)
            End Using
        Catch
            Return New Version("0.0")
        End Try
    End Function
    
    ''' <summary>
    ''' Verifica se pode acessar VBA
    ''' </summary>
    Private Shared Function CanAccessVBA() As Boolean
        Try
            Using excelApp = CreateObject("Excel.Application")
                excelApp.Visible = False
                Using workbook = excelApp.Workbooks.Add()
                    ' Tentar acessar projeto VBA
                    Dim projectName = workbook.VBProject.Name
                    Return Not String.IsNullOrEmpty(projectName)
                End Using
            End Using
        Catch
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' Verifica se .NET Framework está disponível
    ''' </summary>
    Private Shared Function IsDotNetFrameworkAvailable() As Boolean
        Try
            Return Environment.Version >= New Version("4.7.2")
        Catch
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' Obtém configurações recomendadas baseadas no ambiente
    ''' </summary>
    Public Shared Function GetRecommendedSettings() As Dictionary(Of String, Object)
        Dim settings = New Dictionary(Of String, Object)()
        
        Try
            ' Configurações baseadas na capacidade do sistema
            Dim totalRAM = GetTotalSystemMemory()
            Dim cpuCores = Environment.ProcessorCount
            
            ' Timeout baseado em recursos
            settings("DefaultTimeout") = If(totalRAM > 8GB, 60000, 120000) ' ms
            
            ' Visibilidade do Excel baseada em debug
            settings("ExcelVisible") = System.Diagnostics.Debugger.IsAttached
            
            ' Processamento paralelo baseado em CPU
            settings("EnableParallelProcessing") = cpuCores > 2
            
            ' Cache baseado em memória
            settings("EnableCaching") = totalRAM > 4GB
            
            ' Configurações de performance
            settings("ScreenUpdating") = False
            settings("DisplayAlerts") = False
            settings("EnableEvents") = False
            settings("Calculation") = "Manual"
            
            _logger.LogInfo("ExcelAutomationFactory", $"Configurações recomendadas geradas para sistema com {totalRAM}GB RAM e {cpuCores} cores")
            
        Catch ex As Exception
            _logger.LogWarning("ExcelAutomationFactory", "Erro ao gerar configurações, usando padrões", ex)
            
            ' Configurações padrão seguras
            settings("DefaultTimeout") = 90000
            settings("ExcelVisible") = False
            settings("EnableParallelProcessing") = False
            settings("EnableCaching") = False
        End Try
        
        Return settings
    End Function
    
    ''' <summary>
    ''' Obtém memória total do sistema (aproximada)
    ''' </summary>
    Private Shared Function GetTotalSystemMemory() As Long
        Try
            Dim searcher = New Management.ManagementObjectSearcher("SELECT TotalPhysicalMemory FROM Win32_ComputerSystem")
            For Each obj In searcher.Get()
                Return CLng(obj("TotalPhysicalMemory")) \ (1024 * 1024 * 1024) ' Convert to GB
            Next
        Catch
            ' Fallback para 4GB se não conseguir detectar
            Return 4
        End Try
        Return 4
    End Function
End Class

''' <summary>
''' Interface comum para automação Excel
''' </summary>
Public Interface IExcelAutomation
    Inherits IDisposable
    
    Sub ProcessarTalaoCompleto(dados As DadosTalao)
    Function ValidarDados(dados As DadosTalao) As ValidationResult
    ReadOnly Property IsInitialized As Boolean
    ReadOnly Property LastError As String
End Interface

''' <summary>
''' Implementação padrão para automação PDV
''' </summary>
Public Class StandardPDVAutomation
    Implements IExcelAutomation
    
    Private _excelAutomation As ExcelAutomation
    Private _disposed As Boolean = False
    
    Public Sub New()
        _excelAutomation = New ExcelAutomation()
    End Sub
    
    Public ReadOnly Property IsInitialized As Boolean Implements IExcelAutomation.IsInitialized
        Get
            Return _excelAutomation IsNot Nothing
        End Get
    End Property
    
    Public ReadOnly Property LastError As String Implements IExcelAutomation.LastError
        Get
            Return _lastError
        End Get
    End Property
    
    Private _lastError As String = ""
    
    Public Sub ProcessarTalaoCompleto(dados As DadosTalao) Implements IExcelAutomation.ProcessarTalaoCompleto
        Try
            _excelAutomation.ProcessarTalaoCompleto(dados)
        Catch ex As Exception
            _lastError = ex.Message
            Throw
        End Try
    End Sub
    
    Public Function ValidarDados(dados As DadosTalao) As ValidationResult Implements IExcelAutomation.ValidarDados
        Try
            ' Validação específica para PDV
            If dados Is Nothing Then
                Return New ValidationResult(False, "Dados não fornecidos")
            End If
            
            If String.IsNullOrWhiteSpace(dados.NomeCliente) Then
                Return New ValidationResult(False, "Nome do cliente é obrigatório")
            End If
            
            If dados.Produtos Is Nothing OrElse dados.Produtos.Count = 0 Then
                Return New ValidationResult(False, "Pelo menos um produto deve ser informado")
            End If
            
            ' Validar produtos individualmente
            For Each produto In dados.Produtos
                If String.IsNullOrWhiteSpace(produto.Descricao) Then
                    Return New ValidationResult(False, "Descrição do produto é obrigatória")
                End If
                
                If produto.Quantidade <= 0 Then
                    Return New ValidationResult(False, "Quantidade deve ser maior que zero")
                End If
                
                If produto.PrecoUnitario <= 0 Then
                    Return New ValidationResult(False, "Preço unitário deve ser maior que zero")
                End If
            Next
            
            Return New ValidationResult(True)
            
        Catch ex As Exception
            _lastError = ex.Message
            Return New ValidationResult(False, $"Erro na validação: {ex.Message}")
        End Try
    End Function
    
    Public Sub Dispose() Implements IDisposable.Dispose
        If Not _disposed Then
            _excelAutomation?.Dispose()
            _excelAutomation = Nothing
            _disposed = True
        End If
    End Sub
End Class

''' <summary>
''' Placeholder para outras implementações específicas
''' </summary>
Public Class ReportGeneratorAutomation
    Implements IExcelAutomation
    
    Public ReadOnly Property IsInitialized As Boolean Implements IExcelAutomation.IsInitialized
        Get
            Return False ' TODO: Implementar
        End Get
    End Property
    
    Public ReadOnly Property LastError As String Implements IExcelAutomation.LastError
        Get
            Return "Não implementado"
        End Get
    End Property
    
    Public Sub ProcessarTalaoCompleto(dados As DadosTalao) Implements IExcelAutomation.ProcessarTalaoCompleto
        Throw New NotImplementedException("ReportGeneratorAutomation não implementado")
    End Sub
    
    Public Function ValidarDados(dados As DadosTalao) As ValidationResult Implements IExcelAutomation.ValidarDados
        Return New ValidationResult(False, "Não implementado")
    End Function
    
    Public Sub Dispose() Implements IDisposable.Dispose
        ' TODO: Implementar dispose
    End Sub
End Class

Public Class TemplateProcessorAutomation
    Implements IExcelAutomation
    
    Public ReadOnly Property IsInitialized As Boolean Implements IExcelAutomation.IsInitialized
        Get
            Return False ' TODO: Implementar
        End Get
    End Property
    
    Public ReadOnly Property LastError As String Implements IExcelAutomation.LastError
        Get
            Return "Não implementado"
        End Get
    End Property
    
    Public Sub ProcessarTalaoCompleto(dados As DadosTalao) Implements IExcelAutomation.ProcessarTalaoCompleto
        Throw New NotImplementedException("TemplateProcessorAutomation não implementado")
    End Sub
    
    Public Function ValidarDados(dados As DadosTalao) As ValidationResult Implements IExcelAutomation.ValidarDados
        Return New ValidationResult(False, "Não implementado")
    End Function
    
    Public Sub Dispose() Implements IDisposable.Dispose
        ' TODO: Implementar dispose
    End Sub
End Class

Public Class BatchProcessorAutomation
    Implements IExcelAutomation
    
    Public ReadOnly Property IsInitialized As Boolean Implements IExcelAutomation.IsInitialized
        Get
            Return False ' TODO: Implementar
        End Get
    End Property
    
    Public ReadOnly Property LastError As String Implements IExcelAutomation.LastError
        Get
            Return "Não implementado"
        End Get
    End Property
    
    Public Sub ProcessarTalaoCompleto(dados As DadosTalao) Implements IExcelAutomation.ProcessarTalaoCompleto
        Throw New NotImplementedException("BatchProcessorAutomation não implementado")
    End Sub
    
    Public Function ValidarDados(dados As DadosTalao) As ValidationResult Implements IExcelAutomation.ValidarDados
        Return New ValidationResult(False, "Não implementado")
    End Function
    
    Public Sub Dispose() Implements IDisposable.Dispose
        ' TODO: Implementar dispose
    End Sub
End Class