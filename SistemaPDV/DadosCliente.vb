' Classe para armazenar dados do cliente
Public Class DadosCliente
    Public Property Nome As String
    Public Property Endereco As String
    Public Property Cidade As String
    Public Property CEP As String
    Public Property Produtos As String
    Public Property ValorTotal As String
    Public Property FormaPagamento As String
    Public Property Vendedor As String

    Public Sub New()
        ' Valores padrão
        Cidade = "Paulista"
        CEP = "55431-165"
        Vendedor = "matheus-testuser3"
        FormaPagamento = "Dinheiro"
    End Sub

    Public Function ValidarDados() As Boolean
        Return Not String.IsNullOrWhiteSpace(Nome) AndAlso
               Not String.IsNullOrWhiteSpace(Produtos) AndAlso
               Not String.IsNullOrWhiteSpace(ValorTotal) AndAlso
               ValorTotal <> "0,00"
    End Function

    Public Overrides Function ToString() As String
        Return $"Cliente: {Nome}, Valor: R$ {ValorTotal}"
    End Function
End Class