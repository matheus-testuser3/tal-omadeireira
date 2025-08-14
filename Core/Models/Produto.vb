Imports System.ComponentModel.DataAnnotations

''' <summary>
''' Modelo de dados para Produto
''' Representa um produto vendido na madeireira
''' </summary>
Public Class Produto
    Public Property Id As Integer
    
    <Required(ErrorMessage:="Código do produto é obrigatório")>
    <StringLength(20, ErrorMessage:="Código deve ter no máximo 20 caracteres")>
    Public Property Codigo As String
    
    <Required(ErrorMessage:="Descrição do produto é obrigatória")>
    <StringLength(200, ErrorMessage:="Descrição deve ter no máximo 200 caracteres")>
    Public Property Descricao As String
    
    <Required(ErrorMessage:="Unidade de medida é obrigatória")>
    <StringLength(10, ErrorMessage:="Unidade deve ter no máximo 10 caracteres")>
    Public Property Unidade As String
    
    <Range(0.01, Double.MaxValue, ErrorMessage:="Preço deve ser maior que zero")>
    Public Property PrecoUnitario As Decimal
    
    Public Property Categoria As String
    Public Property EstoqueAtual As Decimal
    Public Property EstoqueMinimo As Decimal
    Public Property DataCadastro As DateTime
    Public Property Ativo As Boolean
    
    ''' <summary>
    ''' Construtor padrão
    ''' </summary>
    Public Sub New()
        DataCadastro = DateTime.Now
        Ativo = True
        EstoqueAtual = 0
        EstoqueMinimo = 0
    End Sub
    
    ''' <summary>
    ''' Construtor com parâmetros principais
    ''' </summary>
    Public Sub New(codigo As String, descricao As String, unidade As String, preco As Decimal)
        Me.New()
        Me.Codigo = codigo
        Me.Descricao = descricao
        Me.Unidade = unidade
        Me.PrecoUnitario = preco
    End Sub
    
    ''' <summary>
    ''' Valida se os dados do produto estão corretos
    ''' </summary>
    Public Function IsValid() As Boolean
        Return Not String.IsNullOrWhiteSpace(Codigo) AndAlso 
               Not String.IsNullOrWhiteSpace(Descricao) AndAlso
               Not String.IsNullOrWhiteSpace(Unidade) AndAlso
               PrecoUnitario > 0
    End Function
    
    ''' <summary>
    ''' Verifica se o produto está com estoque baixo
    ''' </summary>
    Public Function EstoqueBaixo() As Boolean
        Return EstoqueAtual <= EstoqueMinimo
    End Function
    
    ''' <summary>
    ''' Retorna representação em string do produto
    ''' </summary>
    Public Overrides Function ToString() As String
        Return $"{Codigo} - {Descricao} ({Unidade}) - {PrecoUnitario:C}"
    End Function
End Class