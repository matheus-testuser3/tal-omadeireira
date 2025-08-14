Imports System.ComponentModel.DataAnnotations

''' <summary>
''' Modelo de dados para Cliente
''' Representa as informações de um cliente da madeireira
''' </summary>
Public Class Cliente
    Public Property Id As Integer
    
    <Required(ErrorMessage:="Nome do cliente é obrigatório")>
    <StringLength(100, ErrorMessage:="Nome deve ter no máximo 100 caracteres")>
    Public Property Nome As String
    
    <StringLength(200, ErrorMessage:="Endereço deve ter no máximo 200 caracteres")>
    Public Property Endereco As String
    
    <RegularExpression("^\d{5}-?\d{3}$", ErrorMessage:="CEP deve estar no formato 00000-000")>
    Public Property CEP As String
    
    <StringLength(100, ErrorMessage:="Cidade deve ter no máximo 100 caracteres")>
    Public Property Cidade As String
    
    <RegularExpression("^\(\d{2}\)\s?\d{4,5}-?\d{4}$", ErrorMessage:="Telefone deve estar no formato (00) 00000-0000")>
    Public Property Telefone As String
    
    Public Property DataCadastro As DateTime
    Public Property UltimaCompra As DateTime?
    Public Property TotalCompras As Decimal
    
    ''' <summary>
    ''' Construtor padrão
    ''' </summary>
    Public Sub New()
        DataCadastro = DateTime.Now
        TotalCompras = 0
    End Sub
    
    ''' <summary>
    ''' Construtor com parâmetros principais
    ''' </summary>
    Public Sub New(nome As String, endereco As String, telefone As String)
        Me.New()
        Me.Nome = nome
        Me.Endereco = endereco
        Me.Telefone = telefone
    End Sub
    
    ''' <summary>
    ''' Valida se os dados do cliente estão corretos
    ''' </summary>
    Public Function IsValid() As Boolean
        Return Not String.IsNullOrWhiteSpace(Nome) AndAlso Nome.Length <= 100
    End Function
    
    ''' <summary>
    ''' Retorna representação em string do cliente
    ''' </summary>
    Public Overrides Function ToString() As String
        Return $"{Nome} - {Telefone}"
    End Function
End Class