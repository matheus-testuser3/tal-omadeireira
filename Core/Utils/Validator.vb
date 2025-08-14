Imports System.Text.RegularExpressions
Imports System.ComponentModel.DataAnnotations

''' <summary>
''' Classe utilitária para validação de dados
''' Centraliza todas as validações do sistema
''' </summary>
Public Class Validator
    
    ''' <summary>
    ''' Valida CPF/CNPJ
    ''' </summary>
    Public Shared Function ValidarDocumento(documento As String) As Boolean
        If String.IsNullOrWhiteSpace(documento) Then Return False
        
        ' Remove caracteres especiais
        documento = Regex.Replace(documento, "[^0-9]", "")
        
        ' Valida CPF (11 dígitos)
        If documento.Length = 11 Then
            Return ValidarCPF(documento)
        End If
        
        ' Valida CNPJ (14 dígitos)  
        If documento.Length = 14 Then
            Return ValidarCNPJ(documento)
        End If
        
        Return False
    End Function
    
    ''' <summary>
    ''' Valida CPF
    ''' </summary>
    Private Shared Function ValidarCPF(cpf As String) As Boolean
        ' Verifica se todos os dígitos são iguais
        If cpf.All(Function(c) c = cpf(0)) Then Return False
        
        ' Calcula o primeiro dígito verificador
        Dim soma = 0
        For i = 0 To 8
            soma += Integer.Parse(cpf(i).ToString()) * (10 - i)
        Next
        Dim resto = soma Mod 11
        Dim digito1 = If(resto < 2, 0, 11 - resto)
        
        ' Verifica o primeiro dígito
        If Integer.Parse(cpf(9).ToString()) <> digito1 Then Return False
        
        ' Calcula o segundo dígito verificador
        soma = 0
        For i = 0 To 9
            soma += Integer.Parse(cpf(i).ToString()) * (11 - i)
        Next
        resto = soma Mod 11
        Dim digito2 = If(resto < 2, 0, 11 - resto)
        
        ' Verifica o segundo dígito
        Return Integer.Parse(cpf(10).ToString()) = digito2
    End Function
    
    ''' <summary>
    ''' Valida CNPJ
    ''' </summary>
    Private Shared Function ValidarCNPJ(cnpj As String) As Boolean
        ' Verifica se todos os dígitos são iguais
        If cnpj.All(Function(c) c = cnpj(0)) Then Return False
        
        Dim multiplicadores1() As Integer = {5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2}
        Dim multiplicadores2() As Integer = {6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2}
        
        ' Calcula o primeiro dígito verificador
        Dim soma = 0
        For i = 0 To 11
            soma += Integer.Parse(cnpj(i).ToString()) * multiplicadores1(i)
        Next
        Dim resto = soma Mod 11
        Dim digito1 = If(resto < 2, 0, 11 - resto)
        
        ' Verifica o primeiro dígito
        If Integer.Parse(cnpj(12).ToString()) <> digito1 Then Return False
        
        ' Calcula o segundo dígito verificador
        soma = 0
        For i = 0 To 12
            soma += Integer.Parse(cnpj(i).ToString()) * multiplicadores2(i)
        Next
        resto = soma Mod 11
        Dim digito2 = If(resto < 2, 0, 11 - resto)
        
        ' Verifica o segundo dígito
        Return Integer.Parse(cnpj(13).ToString()) = digito2
    End Function
    
    ''' <summary>
    ''' Valida CEP
    ''' </summary>
    Public Shared Function ValidarCEP(cep As String) As Boolean
        If String.IsNullOrWhiteSpace(cep) Then Return False
        Return Regex.IsMatch(cep, "^\d{5}-?\d{3}$")
    End Function
    
    ''' <summary>
    ''' Valida telefone
    ''' </summary>
    Public Shared Function ValidarTelefone(telefone As String) As Boolean
        If String.IsNullOrWhiteSpace(telefone) Then Return False
        Return Regex.IsMatch(telefone, "^\(\d{2}\)\s?\d{4,5}-?\d{4}$")
    End Function
    
    ''' <summary>
    ''' Valida email
    ''' </summary>
    Public Shared Function ValidarEmail(email As String) As Boolean
        If String.IsNullOrWhiteSpace(email) Then Return False
        Try
            Dim addr = New System.Net.Mail.MailAddress(email)
            Return addr.Address = email
        Catch
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' Valida valor monetário
    ''' </summary>
    Public Shared Function ValidarValorMonetario(valor As String) As Tuple(Of Boolean, Decimal)
        If String.IsNullOrWhiteSpace(valor) Then Return New Tuple(Of Boolean, Decimal)(False, 0)
        
        ' Remove símbolos de moeda e espaços
        valor = valor.Replace("R$", "").Replace("$", "").Trim()
        
        Dim resultado As Decimal
        Dim isValid = Decimal.TryParse(valor, resultado) AndAlso resultado >= 0
        
        Return New Tuple(Of Boolean, Decimal)(isValid, resultado)
    End Function
    
    ''' <summary>
    ''' Valida quantidade
    ''' </summary>
    Public Shared Function ValidarQuantidade(quantidade As String) As Tuple(Of Boolean, Decimal)
        If String.IsNullOrWhiteSpace(quantidade) Then Return New Tuple(Of Boolean, Decimal)(False, 0)
        
        Dim resultado As Decimal
        Dim isValid = Decimal.TryParse(quantidade, resultado) AndAlso resultado > 0
        
        Return New Tuple(Of Boolean, Decimal)(isValid, resultado)
    End Function
    
    ''' <summary>
    ''' Valida objeto usando Data Annotations
    ''' </summary>
    Public Shared Function ValidarObjeto(obj As Object) As List(Of ValidationResult)
        Dim context = New ValidationContext(obj)
        Dim results = New List(Of ValidationResult)()
        Validator.TryValidateObject(obj, context, results, True)
        Return results
    End Function
    
    ''' <summary>
    ''' Formata CEP
    ''' </summary>
    Public Shared Function FormatarCEP(cep As String) As String
        If String.IsNullOrWhiteSpace(cep) Then Return ""
        cep = Regex.Replace(cep, "[^0-9]", "")
        If cep.Length = 8 Then
            Return $"{cep.Substring(0, 5)}-{cep.Substring(5, 3)}"
        End If
        Return cep
    End Function
    
    ''' <summary>
    ''' Formata telefone
    ''' </summary>
    Public Shared Function FormatarTelefone(telefone As String) As String
        If String.IsNullOrWhiteSpace(telefone) Then Return ""
        telefone = Regex.Replace(telefone, "[^0-9]", "")
        
        If telefone.Length = 10 Then
            Return $"({telefone.Substring(0, 2)}) {telefone.Substring(2, 4)}-{telefone.Substring(6, 4)}"
        ElseIf telefone.Length = 11 Then
            Return $"({telefone.Substring(0, 2)}) {telefone.Substring(2, 5)}-{telefone.Substring(7, 4)}"
        End If
        
        Return telefone
    End Function
End Class