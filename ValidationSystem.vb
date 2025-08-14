Imports System.Text.RegularExpressions
Imports System.Globalization

''' <summary>
''' Sistema centralizado de validação de dados
''' Fornece validações consistentes para todo o sistema
''' </summary>
Public Class ValidationSystem
    
    Private Shared ReadOnly _logger As LoggingSystem = LoggingSystem.Instance
    
    #Region "Validações de Texto"
    
    ''' <summary>
    ''' Valida se uma string não está vazia ou nula
    ''' </summary>
    Public Shared Function ValidateRequired(value As String, fieldName As String) As ValidationResult
        If String.IsNullOrWhiteSpace(value) Then
            Return New ValidationResult(False, $"O campo '{fieldName}' é obrigatório")
        End If
        Return New ValidationResult(True)
    End Function
    
    ''' <summary>
    ''' Valida tamanho mínimo e máximo de string
    ''' </summary>
    Public Shared Function ValidateStringLength(value As String, fieldName As String, 
                                               Optional minLength As Integer = 0, 
                                               Optional maxLength As Integer = Integer.MaxValue) As ValidationResult
        If String.IsNullOrEmpty(value) Then
            If minLength > 0 Then
                Return New ValidationResult(False, $"O campo '{fieldName}' é obrigatório")
            End If
            Return New ValidationResult(True)
        End If
        
        If value.Length < minLength Then
            Return New ValidationResult(False, $"O campo '{fieldName}' deve ter pelo menos {minLength} caracteres")
        End If
        
        If value.Length > maxLength Then
            Return New ValidationResult(False, $"O campo '{fieldName}' deve ter no máximo {maxLength} caracteres")
        End If
        
        Return New ValidationResult(True)
    End Function
    
    #End Region
    
    #Region "Validações de Documento"
    
    ''' <summary>
    ''' Valida formato de CPF
    ''' </summary>
    Public Shared Function ValidateCPF(cpf As String) As ValidationResult
        If String.IsNullOrWhiteSpace(cpf) Then
            Return New ValidationResult(False, "CPF não informado")
        End If
        
        ' Remove formatação
        cpf = Regex.Replace(cpf, "[^0-9]", "")
        
        ' Verifica se tem 11 dígitos
        If cpf.Length <> 11 Then
            Return New ValidationResult(False, "CPF deve ter 11 dígitos")
        End If
        
        ' Verifica CPFs inválidos conhecidos
        If cpf = "00000000000" Or cpf = "11111111111" Or cpf = "22222222222" Or
           cpf = "33333333333" Or cpf = "44444444444" Or cpf = "55555555555" Or
           cpf = "66666666666" Or cpf = "77777777777" Or cpf = "88888888888" Or
           cpf = "99999999999" Then
            Return New ValidationResult(False, "CPF inválido")
        End If
        
        ' Validação do algoritmo do CPF
        Try
            Dim digito1 = CalcularDigitoCPF(cpf.Substring(0, 9))
            Dim digito2 = CalcularDigitoCPF(cpf.Substring(0, 9) & digito1)
            
            If cpf.EndsWith(digito1.ToString() & digito2.ToString()) Then
                Return New ValidationResult(True)
            Else
                Return New ValidationResult(False, "CPF inválido")
            End If
            
        Catch ex As Exception
            _logger.LogError("ValidationSystem", "Erro ao validar CPF", ex)
            Return New ValidationResult(False, "Erro na validação do CPF")
        End Try
    End Function
    
    ''' <summary>
    ''' Valida formato de CNPJ
    ''' </summary>
    Public Shared Function ValidateCNPJ(cnpj As String) As ValidationResult
        If String.IsNullOrWhiteSpace(cnpj) Then
            Return New ValidationResult(False, "CNPJ não informado")
        End If
        
        ' Remove formatação
        cnpj = Regex.Replace(cnpj, "[^0-9]", "")
        
        ' Verifica se tem 14 dígitos
        If cnpj.Length <> 14 Then
            Return New ValidationResult(False, "CNPJ deve ter 14 dígitos")
        End If
        
        ' Verifica CNPJs inválidos conhecidos
        If cnpj = "00000000000000" Or cnpj = "11111111111111" Then
            Return New ValidationResult(False, "CNPJ inválido")
        End If
        
        ' Validação do algoritmo do CNPJ
        Try
            Dim digito1 = CalcularDigitoCNPJ(cnpj.Substring(0, 12), {5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2})
            Dim digito2 = CalcularDigitoCNPJ(cnpj.Substring(0, 12) & digito1, {6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2})
            
            If cnpj.EndsWith(digito1.ToString() & digito2.ToString()) Then
                Return New ValidationResult(True)
            Else
                Return New ValidationResult(False, "CNPJ inválido")
            End If
            
        Catch ex As Exception
            _logger.LogError("ValidationSystem", "Erro ao validar CNPJ", ex)
            Return New ValidationResult(False, "Erro na validação do CNPJ")
        End Try
    End Function
    
    #End Region
    
    #Region "Validações de Contato"
    
    ''' <summary>
    ''' Valida formato de email
    ''' </summary>
    Public Shared Function ValidateEmail(email As String) As ValidationResult
        If String.IsNullOrWhiteSpace(email) Then
            Return New ValidationResult(False, "Email não informado")
        End If
        
        Try
            Dim emailRegex = New Regex("^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$")
            If emailRegex.IsMatch(email) Then
                Return New ValidationResult(True)
            Else
                Return New ValidationResult(False, "Formato de email inválido")
            End If
            
        Catch ex As Exception
            _logger.LogError("ValidationSystem", "Erro ao validar email", ex)
            Return New ValidationResult(False, "Erro na validação do email")
        End Try
    End Function
    
    ''' <summary>
    ''' Valida formato de telefone brasileiro
    ''' </summary>
    Public Shared Function ValidatePhone(telefone As String) As ValidationResult
        If String.IsNullOrWhiteSpace(telefone) Then
            Return New ValidationResult(False, "Telefone não informado")
        End If
        
        ' Remove formatação
        Dim telefoneNumeros = Regex.Replace(telefone, "[^0-9]", "")
        
        ' Valida formato brasileiro (10 ou 11 dígitos)
        If telefoneNumeros.Length < 10 Or telefoneNumeros.Length > 11 Then
            Return New ValidationResult(False, "Telefone deve ter 10 ou 11 dígitos")
        End If
        
        ' Validações básicas de formato
        If telefoneNumeros.Length = 11 AndAlso Not telefoneNumeros.Substring(2, 1) = "9" Then
            Return New ValidationResult(False, "Celular deve começar com 9 após o DDD")
        End If
        
        Return New ValidationResult(True)
    End Function
    
    ''' <summary>
    ''' Valida formato de CEP
    ''' </summary>
    Public Shared Function ValidateCEP(cep As String) As ValidationResult
        If String.IsNullOrWhiteSpace(cep) Then
            Return New ValidationResult(False, "CEP não informado")
        End If
        
        ' Remove formatação
        Dim cepNumeros = Regex.Replace(cep, "[^0-9]", "")
        
        If cepNumeros.Length <> 8 Then
            Return New ValidationResult(False, "CEP deve ter 8 dígitos")
        End If
        
        ' Verifica se não é um CEP inválido conhecido
        If cepNumeros = "00000000" Then
            Return New ValidationResult(False, "CEP inválido")
        End If
        
        Return New ValidationResult(True)
    End Function
    
    #End Region
    
    #Region "Validações Numéricas"
    
    ''' <summary>
    ''' Valida valores decimais
    ''' </summary>
    Public Shared Function ValidateDecimal(value As String, fieldName As String, 
                                         Optional minValue As Decimal = Decimal.MinValue,
                                         Optional maxValue As Decimal = Decimal.MaxValue) As ValidationResult
        If String.IsNullOrWhiteSpace(value) Then
            Return New ValidationResult(False, $"O campo '{fieldName}' é obrigatório")
        End If
        
        Dim decimalValue As Decimal
        If Not Decimal.TryParse(value, NumberStyles.Number, CultureInfo.CurrentCulture, decimalValue) Then
            Return New ValidationResult(False, $"O campo '{fieldName}' deve ser um número válido")
        End If
        
        If decimalValue < minValue Then
            Return New ValidationResult(False, $"O campo '{fieldName}' deve ser maior ou igual a {minValue}")
        End If
        
        If decimalValue > maxValue Then
            Return New ValidationResult(False, $"O campo '{fieldName}' deve ser menor ou igual a {maxValue}")
        End If
        
        Return New ValidationResult(True)
    End Function
    
    ''' <summary>
    ''' Valida valores inteiros
    ''' </summary>
    Public Shared Function ValidateInteger(value As String, fieldName As String,
                                         Optional minValue As Integer = Integer.MinValue,
                                         Optional maxValue As Integer = Integer.MaxValue) As ValidationResult
        If String.IsNullOrWhiteSpace(value) Then
            Return New ValidationResult(False, $"O campo '{fieldName}' é obrigatório")
        End If
        
        Dim intValue As Integer
        If Not Integer.TryParse(value, intValue) Then
            Return New ValidationResult(False, $"O campo '{fieldName}' deve ser um número inteiro válido")
        End If
        
        If intValue < minValue Then
            Return New ValidationResult(False, $"O campo '{fieldName}' deve ser maior ou igual a {minValue}")
        End If
        
        If intValue > maxValue Then
            Return New ValidationResult(False, $"O campo '{fieldName}' deve ser menor ou igual a {maxValue}")
        End If
        
        Return New ValidationResult(True)
    End Function
    
    #End Region
    
    #Region "Métodos Auxiliares"
    
    ''' <summary>
    ''' Calcula dígito verificador do CPF
    ''' </summary>
    Private Shared Function CalcularDigitoCPF(cpf As String) As Integer
        Dim soma = 0
        Dim peso = cpf.Length + 1
        
        For i = 0 To cpf.Length - 1
            soma += Integer.Parse(cpf(i)) * peso
            peso -= 1
        Next
        
        Dim resto = soma Mod 11
        Return If(resto < 2, 0, 11 - resto)
    End Function
    
    ''' <summary>
    ''' Calcula dígito verificador do CNPJ
    ''' </summary>
    Private Shared Function CalcularDigitoCNPJ(cnpj As String, pesos As Integer()) As Integer
        Dim soma = 0
        
        For i = 0 To cnpj.Length - 1
            soma += Integer.Parse(cnpj(i)) * pesos(i)
        Next
        
        Dim resto = soma Mod 11
        Return If(resto < 2, 0, 11 - resto)
    End Function
    
    #End Region
End Class

''' <summary>
''' Resultado de uma validação
''' </summary>
Public Class ValidationResult
    Public Property IsValid As Boolean
    Public Property ErrorMessage As String
    
    Public Sub New(isValid As Boolean, Optional errorMessage As String = "")
        Me.IsValid = isValid
        Me.ErrorMessage = errorMessage
    End Sub
    
    Public Overrides Function ToString() As String
        Return If(IsValid, "Válido", $"Inválido: {ErrorMessage}")
    End Function
End Class