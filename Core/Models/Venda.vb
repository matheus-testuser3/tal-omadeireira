Imports System.ComponentModel.DataAnnotations

''' <summary>
''' Modelo de dados para Venda
''' Representa uma venda completa com itens
''' </summary>
Public Class Venda
    Public Property Id As Integer
    Public Property NumeroTalao As String
    Public Property DataVenda As DateTime
    Public Property Cliente As Cliente
    Public Property Vendedor As String
    Public Property FormaPagamento As String
    Public Property Itens As List(Of ItemVenda)
    Public Property ValorTotal As Decimal
    Public Property Observacoes As String
    Public Property Status As StatusVenda
    
    ''' <summary>
    ''' Construtor padrão
    ''' </summary>
    Public Sub New()
        DataVenda = DateTime.Now
        Itens = New List(Of ItemVenda)()
        Status = StatusVenda.Ativa
        NumeroTalao = GerarNumeroTalao()
    End Sub
    
    ''' <summary>
    ''' Construtor com cliente e vendedor
    ''' </summary>
    Public Sub New(cliente As Cliente, vendedor As String)
        Me.New()
        Me.Cliente = cliente
        Me.Vendedor = vendedor
    End Sub
    
    ''' <summary>
    ''' Adiciona um item à venda
    ''' </summary>
    Public Sub AdicionarItem(produto As Produto, quantidade As Decimal)
        Dim item As New ItemVenda(produto, quantidade)
        Itens.Add(item)
        CalcularTotal()
    End Sub
    
    ''' <summary>
    ''' Remove um item da venda
    ''' </summary>
    Public Sub RemoverItem(item As ItemVenda)
        Itens.Remove(item)
        CalcularTotal()
    End Sub
    
    ''' <summary>
    ''' Calcula o valor total da venda
    ''' </summary>
    Public Sub CalcularTotal()
        ValorTotal = Itens.Sum(Function(i) i.ValorTotal)
    End Sub
    
    ''' <summary>
    ''' Valida se a venda está completa
    ''' </summary>
    Public Function IsValid() As Boolean
        Return Cliente IsNot Nothing AndAlso
               Not String.IsNullOrWhiteSpace(Vendedor) AndAlso
               Itens.Count > 0 AndAlso
               Itens.All(Function(i) i.IsValid())
    End Function
    
    ''' <summary>
    ''' Gera número sequencial do talão
    ''' </summary>
    Private Function GerarNumeroTalao() As String
        ' Por enquanto usa timestamp, depois será sequencial do banco
        Return DateTime.Now.ToString("yyyyMMddHHmmss")
    End Function
    
    ''' <summary>
    ''' Retorna representação em string da venda
    ''' </summary>
    Public Overrides Function ToString() As String
        Return $"Talão {NumeroTalao} - {Cliente?.Nome} - {ValorTotal:C}"
    End Function
End Class

''' <summary>
''' Representa um item individual da venda
''' </summary>
Public Class ItemVenda
    Public Property Id As Integer
    Public Property Produto As Produto
    Public Property Quantidade As Decimal
    Public Property PrecoUnitario As Decimal
    Public Property ValorTotal As Decimal
    
    ''' <summary>
    ''' Construtor padrão
    ''' </summary>
    Public Sub New()
    End Sub
    
    ''' <summary>
    ''' Construtor com produto e quantidade
    ''' </summary>
    Public Sub New(produto As Produto, quantidade As Decimal)
        Me.Produto = produto
        Me.Quantidade = quantidade
        Me.PrecoUnitario = produto.PrecoUnitario
        Me.ValorTotal = quantidade * PrecoUnitario
    End Sub
    
    ''' <summary>
    ''' Valida o item da venda
    ''' </summary>
    Public Function IsValid() As Boolean
        Return Produto IsNot Nothing AndAlso
               Quantidade > 0 AndAlso
               PrecoUnitario > 0
    End Function
    
    ''' <summary>
    ''' Recalcula o valor total do item
    ''' </summary>
    Public Sub CalcularTotal()
        ValorTotal = Quantidade * PrecoUnitario
    End Sub
    
    ''' <summary>
    ''' Retorna representação em string do item
    ''' </summary>
    Public Overrides Function ToString() As String
        Return $"{Produto?.Descricao} - {Quantidade} {Produto?.Unidade} x {PrecoUnitario:C} = {ValorTotal:C}"
    End Function
End Class

''' <summary>
''' Status possíveis para uma venda
''' </summary>
Public Enum StatusVenda
    Ativa = 1
    Finalizada = 2
    Cancelada = 3
End Enum