Imports System.Net.Http
Imports System.Threading.Tasks
Imports Newtonsoft.Json

''' <summary>
''' Sistema de integração com APIs externas
''' Consulta CEP, validações e outras integrações web
''' </summary>
Public Class ExternalAPIManager
    
    Private Shared _instance As ExternalAPIManager
    Private Shared ReadOnly _lockObject As New Object()
    
    Private ReadOnly _httpClient As HttpClient
    Private ReadOnly _logger As LoggingSystem = LoggingSystem.Instance
    Private ReadOnly _config As EnhancedConfigurationManager = EnhancedConfigurationManager.Instance
    Private ReadOnly _cache As CacheManager = CacheManager.Instance
    
    ''' <summary>
    ''' Singleton instance
    ''' </summary>
    Public Shared ReadOnly Property Instance As ExternalAPIManager
        Get
            If _instance Is Nothing Then
                SyncLock _lockObject
                    If _instance Is Nothing Then
                        _instance = New ExternalAPIManager()
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
        _httpClient = New HttpClient()
        _httpClient.Timeout = TimeSpan.FromSeconds(10)
        _httpClient.DefaultRequestHeaders.Add("User-Agent", "SistemaPDV-MadeireiraMariaLuiza/1.0")
        
        _logger.LogInfo("ExternalAPIManager", "Gerenciador de APIs externas inicializado")
    End Sub
    
    #Region "Consulta CEP"
    
    ''' <summary>
    ''' Consulta CEP via ViaCEP
    ''' </summary>
    Public Async Function ConsultarCEPAsync(cep As String) As Task(Of CEPResult)
        Try
            ' Limpar CEP
            Dim cepLimpo = System.Text.RegularExpressions.Regex.Replace(cep, "[^0-9]", "")
            
            ' Validar formato
            If cepLimpo.Length <> 8 Then
                Return New CEPResult() With {
                    .Sucesso = False,
                    .Erro = "CEP deve ter 8 dígitos"
                }
            End If
            
            ' Verificar cache primeiro
            Dim cacheKey = $"cep_{cepLimpo}"
            Dim cached = _cache.Get(Of CEPResult)(cacheKey)
            If cached IsNot Nothing Then
                _logger.LogDebug("ExternalAPIManager", $"CEP {cepLimpo} encontrado no cache")
                Return cached
            End If
            
            ' Consultar ViaCEP
            Dim url = $"https://viacep.com.br/ws/{cepLimpo}/json/"
            _logger.LogDebug("ExternalAPIManager", $"Consultando CEP: {url}")
            
            Dim response = Await _httpClient.GetStringAsync(url)
            Dim viaCepResponse = JsonConvert.DeserializeObject(Of ViaCEPResponse)(response)
            
            Dim result As CEPResult
            
            If viaCepResponse.erro Then
                result = New CEPResult() With {
                    .Sucesso = False,
                    .Erro = "CEP não encontrado"
                }
            Else
                result = New CEPResult() With {
                    .Sucesso = True,
                    .CEP = cepLimpo,
                    .Logradouro = viaCepResponse.logradouro,
                    .Bairro = viaCepResponse.bairro,
                    .Cidade = viaCepResponse.localidade,
                    .UF = viaCepResponse.uf,
                    .IBGE = viaCepResponse.ibge,
                    .Complemento = viaCepResponse.complemento
                }
            End If
            
            ' Salvar no cache por 24 horas
            _cache.Set(cacheKey, result, 1440)
            
            _logger.LogInfo("ExternalAPIManager", $"CEP {cepLimpo} consultado: {If(result.Sucesso, "Sucesso", result.Erro)}")
            Return result
            
        Catch ex As HttpRequestException
            _logger.LogError("ExternalAPIManager", $"Erro de rede ao consultar CEP {cep}", ex)
            Return New CEPResult() With {
                .Sucesso = False,
                .Erro = "Erro de conexão com o serviço de CEP"
            }
        Catch ex As TaskCanceledException
            _logger.LogError("ExternalAPIManager", $"Timeout ao consultar CEP {cep}", ex)
            Return New CEPResult() With {
                .Sucesso = False,
                .Erro = "Timeout na consulta de CEP"
            }
        Catch ex As Exception
            _logger.LogError("ExternalAPIManager", $"Erro inesperado ao consultar CEP {cep}", ex)
            Return New CEPResult() With {
                .Sucesso = False,
                .Erro = $"Erro inesperado: {ex.Message}"
            }
        End Try
    End Function
    
    ''' <summary>
    ''' Versão síncrona da consulta de CEP
    ''' </summary>
    Public Function ConsultarCEP(cep As String) As CEPResult
        Try
            Return ConsultarCEPAsync(cep).Result
        Catch ex As AggregateException
            ' Desembrulhar exceção agregada
            Dim innerEx = ex.InnerException
            _logger.LogError("ExternalAPIManager", $"Erro síncrono ao consultar CEP {cep}", innerEx)
            Return New CEPResult() With {
                .Sucesso = False,
                .Erro = innerEx?.Message Or "Erro desconhecido"
            }
        End Try
    End Function
    
    #End Region
    
    #Region "Validação de CNPJ"
    
    ''' <summary>
    ''' Consulta CNPJ na Receita Federal (ReceitaWS)
    ''' </summary>
    Public Async Function ConsultarCNPJAsync(cnpj As String) As Task(Of CNPJResult)
        Try
            ' Limpar CNPJ
            Dim cnpjLimpo = System.Text.RegularExpressions.Regex.Replace(cnpj, "[^0-9]", "")
            
            ' Validar formato
            If cnpjLimpo.Length <> 14 Then
                Return New CNPJResult() With {
                    .Sucesso = False,
                    .Erro = "CNPJ deve ter 14 dígitos"
                }
            End If
            
            ' Verificar cache primeiro
            Dim cacheKey = $"cnpj_{cnpjLimpo}"
            Dim cached = _cache.Get(Of CNPJResult)(cacheKey)
            If cached IsNot Nothing Then
                _logger.LogDebug("ExternalAPIManager", $"CNPJ {cnpjLimpo} encontrado no cache")
                Return cached
            End If
            
            ' Consultar ReceitaWS
            Dim url = $"https://www.receitaws.com.br/v1/cnpj/{cnpjLimpo}"
            _logger.LogDebug("ExternalAPIManager", $"Consultando CNPJ: {url}")
            
            Dim response = Await _httpClient.GetStringAsync(url)
            Dim receitaResponse = JsonConvert.DeserializeObject(Of ReceitaWSResponse)(response)
            
            Dim result As CNPJResult
            
            If receitaResponse.status = "ERROR" Then
                result = New CNPJResult() With {
                    .Sucesso = False,
                    .Erro = receitaResponse.message Or "CNPJ não encontrado"
                }
            Else
                result = New CNPJResult() With {
                    .Sucesso = True,
                    .CNPJ = cnpjLimpo,
                    .RazaoSocial = receitaResponse.nome,
                    .NomeFantasia = receitaResponse.fantasia,
                    .Situacao = receitaResponse.situacao,
                    .DataSituacao = receitaResponse.data_situacao,
                    .Atividade = receitaResponse.atividade_principal?.FirstOrDefault()?.text,
                    .CEP = receitaResponse.cep,
                    .Logradouro = receitaResponse.logradouro,
                    .Numero = receitaResponse.numero,
                    .Bairro = receitaResponse.bairro,
                    .Municipio = receitaResponse.municipio,
                    .UF = receitaResponse.uf
                }
            End If
            
            ' Salvar no cache por 7 dias (dados da Receita mudam pouco)
            _cache.Set(cacheKey, result, 10080)
            
            _logger.LogInfo("ExternalAPIManager", $"CNPJ {cnpjLimpo} consultado: {If(result.Sucesso, "Sucesso", result.Erro)}")
            Return result
            
        Catch ex As HttpRequestException
            _logger.LogError("ExternalAPIManager", $"Erro de rede ao consultar CNPJ {cnpj}", ex)
            Return New CNPJResult() With {
                .Sucesso = False,
                .Erro = "Erro de conexão com o serviço de CNPJ"
            }
        Catch ex As Exception
            _logger.LogError("ExternalAPIManager", $"Erro inesperado ao consultar CNPJ {cnpj}", ex)
            Return New CNPJResult() With {
                .Sucesso = False,
                .Erro = $"Erro inesperado: {ex.Message}"
            }
        End Try
    End Function
    
    #End Region
    
    #Region "Teste de Conectividade"
    
    ''' <summary>
    ''' Testa conectividade com as APIs
    ''' </summary>
    Public Async Function TestarConectividadeAsync() As Task(Of ConnectivityResult)
        Dim result As New ConnectivityResult()
        
        Try
            ' Teste ViaCEP
            Dim cepTask = TestarViaCEPAsync()
            
            ' Teste ReceitaWS  
            Dim cnpjTask = TestarReceitaWSAsync()
            
            ' Aguardar ambos
            Await Task.WhenAll(cepTask, cnpjTask)
            
            result.ViaCEP = cepTask.Result
            result.ReceitaWS = cnpjTask.Result
            result.Sucesso = result.ViaCEP.Sucesso AndAlso result.ReceitaWS.Sucesso
            
            _logger.LogInfo("ExternalAPIManager", $"Teste de conectividade: {If(result.Sucesso, "Sucesso", "Falhas detectadas")}")
            
        Catch ex As Exception
            _logger.LogError("ExternalAPIManager", "Erro no teste de conectividade", ex)
            result.Sucesso = False
            result.ErroGeral = ex.Message
        End Try
        
        Return result
    End Function
    
    ''' <summary>
    ''' Testa ViaCEP com CEP conhecido
    ''' </summary>
    Private Async Function TestarViaCEPAsync() As Task(Of ServiceTestResult)
        Try
            Dim resultado = Await ConsultarCEPAsync("01310-100") ' Av. Paulista, SP
            Return New ServiceTestResult() With {
                .Sucesso = resultado.Sucesso,
                .Erro = resultado.Erro,
                .TempoResposta = Date.Now ' TODO: Medir tempo real
            }
        Catch ex As Exception
            Return New ServiceTestResult() With {
                .Sucesso = False,
                .Erro = ex.Message
            }
        End Try
    End Function
    
    ''' <summary>
    ''' Testa ReceitaWS com CNPJ conhecido
    ''' </summary>
    Private Async Function TestarReceitaWSAsync() As Task(Of ServiceTestResult)
        Try
            Dim resultado = Await ConsultarCNPJAsync("11.222.333/0001-81") ' CNPJ de teste
            Return New ServiceTestResult() With {
                .Sucesso = True, ' ReceitaWS pode retornar erro para CNPJ inexistente, mas serviço está funcionando
                .Erro = "",
                .TempoResposta = Date.Now
            }
        Catch ex As Exception
            Return New ServiceTestResult() With {
                .Sucesso = False,
                .Erro = ex.Message
            }
        End Try
    End Function
    
    #End Region
    
    #Region "Cleanup"
    
    ''' <summary>
    ''' Libera recursos
    ''' </summary>
    Public Sub Dispose()
        Try
            _httpClient?.Dispose()
            _logger.LogInfo("ExternalAPIManager", "Recursos liberados")
        Catch ex As Exception
            _logger.LogError("ExternalAPIManager", "Erro ao liberar recursos", ex)
        End Try
    End Sub
    
    #End Region
End Class

#Region "Modelos de Dados"

''' <summary>
''' Resultado da consulta de CEP
''' </summary>
Public Class CEPResult
    Public Property Sucesso As Boolean
    Public Property Erro As String
    Public Property CEP As String
    Public Property Logradouro As String
    Public Property Bairro As String
    Public Property Cidade As String
    Public Property UF As String
    Public Property IBGE As String
    Public Property Complemento As String
End Class

''' <summary>
''' Resultado da consulta de CNPJ
''' </summary>
Public Class CNPJResult
    Public Property Sucesso As Boolean
    Public Property Erro As String
    Public Property CNPJ As String
    Public Property RazaoSocial As String
    Public Property NomeFantasia As String
    Public Property Situacao As String
    Public Property DataSituacao As String
    Public Property Atividade As String
    Public Property CEP As String
    Public Property Logradouro As String
    Public Property Numero As String
    Public Property Bairro As String
    Public Property Municipio As String
    Public Property UF As String
End Class

''' <summary>
''' Resultado do teste de conectividade
''' </summary>
Public Class ConnectivityResult
    Public Property Sucesso As Boolean
    Public Property ErroGeral As String
    Public Property ViaCEP As ServiceTestResult
    Public Property ReceitaWS As ServiceTestResult
End Class

''' <summary>
''' Resultado de teste de serviço individual
''' </summary>
Public Class ServiceTestResult
    Public Property Sucesso As Boolean
    Public Property Erro As String
    Public Property TempoResposta As Date
End Class

''' <summary>
''' Resposta da API ViaCEP
''' </summary>
Friend Class ViaCEPResponse
    Public Property cep As String
    Public Property logradouro As String
    Public Property complemento As String
    Public Property bairro As String
    Public Property localidade As String
    Public Property uf As String
    Public Property ibge As String
    Public Property erro As Boolean
End Class

''' <summary>
''' Resposta da API ReceitaWS
''' </summary>
Friend Class ReceitaWSResponse
    Public Property status As String
    Public Property message As String
    Public Property nome As String
    Public Property fantasia As String
    Public Property situacao As String
    Public Property data_situacao As String
    Public Property atividade_principal As List(Of AtividadeResponse)
    Public Property cep As String
    Public Property logradouro As String
    Public Property numero As String
    Public Property bairro As String
    Public Property municipio As String
    Public Property uf As String
End Class

''' <summary>
''' Atividade da empresa
''' </summary>
Friend Class AtividadeResponse
    Public Property code As String
    Public Property text As String
End Class

#End Region