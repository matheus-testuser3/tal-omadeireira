Imports System.ComponentModel

''' <summary>
''' Modelos de dados para o Sistema PDV - Madeireira Maria Luiza
''' Classes que representam entidades do negócio
''' </summary>

#Region "Classe Cliente"
''' <summary>
''' Representa um cliente da madeireira
''' </summary>
Public Class Cliente
    Public Property ID As Integer
    Public Property Nome As String
    Public Property Endereco As String
    Public Property CEP As String
    Public Property Cidade As String
    Public Property UF As String
    Public Property Telefone As String
    Public Property Email As String
    Public Property CPF_CNPJ As String
    Public Property DataCadastro As Date
    Public Property Ativo As Boolean
    Public Property Observacoes As String
    
    Public Sub New()
        ID = 0
        Nome = ""
        Endereco = ""
        CEP = ""
        Cidade = ""
        UF = ""
        Telefone = ""
        Email = ""
        CPF_CNPJ = ""
        DataCadastro = Date.Now
        Ativo = True
        Observacoes = ""
    End Sub
    
    Public Overrides Function ToString() As String
        Return $"{Nome} - {Cidade}/{UF}"
    End Function
End Class
#End Region

#Region "Classe Produto"
''' <summary>
''' Representa um produto da madeireira
''' </summary>
Public Class Produto
    Public Property ID As Integer
    Public Property Codigo As String
    Public Property Descricao As String
    Public Property Secao As String
    Public Property Unidade As String
    Public Property PrecoVenda As Decimal
    Public Property PrecoCusto As Decimal
    Public Property EstoqueAtual As Double
    Public Property EstoqueMinimo As Double
    Public Property Ativo As Boolean
    Public Property DataCadastro As Date
    Public Property Observacoes As String
    
    Public Sub New()
        ID = 0
        Codigo = ""
        Descricao = ""
        Secao = ""
        Unidade = "UN"
        PrecoVenda = 0
        PrecoCusto = 0
        EstoqueAtual = 0
        EstoqueMinimo = 0
        Ativo = True
        DataCadastro = Date.Now
        Observacoes = ""
    End Sub
    
    Public ReadOnly Property MargemLucro As Decimal
        Get
            If PrecoCusto > 0 Then
                Return ((PrecoVenda - PrecoCusto) / PrecoCusto) * 100
            End If
            Return 0
        End Get
    End Property
    
    Public Overrides Function ToString() As String
        Return $"{Codigo} - {Descricao}"
    End Function
End Class
#End Region

#Region "Classe ItemVenda"
''' <summary>
''' Representa um item individual de uma venda
''' </summary>
Public Class ItemVenda
    Public Property ID As Integer
    Public Property VendaID As Integer
    Public Property ProdutoID As Integer
    Public Property Produto As Produto
    Public Property Quantidade As Double
    Public Property PrecoUnitario As Decimal
    Public Property Desconto As Decimal
    Public Property Observacoes As String
    
    Public Sub New()
        ID = 0
        VendaID = 0
        ProdutoID = 0
        Produto = New Produto()
        Quantidade = 1
        PrecoUnitario = 0
        Desconto = 0
        Observacoes = ""
    End Sub
    
    Public ReadOnly Property Subtotal As Decimal
        Get
            Return (PrecoUnitario * Quantidade) - Desconto
        End Get
    End Property
    
    Public Overrides Function ToString() As String
        Return $"{Quantidade} {Produto.Unidade} - {Produto.Descricao}"
    End Function
End Class
#End Region

#Region "Classe Venda"
''' <summary>
''' Representa uma venda completa
''' </summary>
Public Class Venda
    Public Property ID As Integer
    Public Property NumeroTalao As String
    Public Property ClienteID As Integer
    Public Property Cliente As Cliente
    Public Property DataVenda As Date
    Public Property FormaPagamento As String
    Public Property VendedorID As Integer
    Public Property VendedorNome As String
    Public Property Itens As List(Of ItemVenda)
    Public Property DescontoGeral As Decimal
    Public Property Frete As Decimal
    Public Property Status As String
    Public Property Observacoes As String
    
    Public Sub New()
        ID = 0
        NumeroTalao = GerarNumeroTalao()
        ClienteID = 0
        Cliente = New Cliente()
        DataVenda = Date.Now
        FormaPagamento = "À Vista"
        VendedorID = 0
        VendedorNome = ""
        Itens = New List(Of ItemVenda)()
        DescontoGeral = 0
        Frete = 0
        Status = "Pendente"
        Observacoes = ""
    End Sub
    
    Public ReadOnly Property SubtotalItens As Decimal
        Get
            Return Itens.Sum(Function(i) i.Subtotal)
        End Get
    End Property
    
    Public ReadOnly Property TotalGeral As Decimal
        Get
            Return SubtotalItens - DescontoGeral + Frete
        End Get
    End Property
    
    Public ReadOnly Property QuantidadeItens As Integer
        Get
            Return Itens.Count
        End Get
    End Property
    
    Private Function GerarNumeroTalao() As String
        Return Date.Now.ToString("yyyyMMddHHmmss")
    End Function
    
    Public Overrides Function ToString() As String
        Return $"Talão {NumeroTalao} - {Cliente.Nome} - R$ {TotalGeral:F2}"
    End Function
End Class
#End Region

#Region "Classe Vendedor"
''' <summary>
''' Representa um vendedor do sistema
''' </summary>
Public Class Vendedor
    Public Property ID As Integer
    Public Property Nome As String
    Public Property Usuario As String
    Public Property Email As String
    Public Property Telefone As String
    Public Property Comissao As Decimal
    Public Property Ativo As Boolean
    Public Property DataCadastro As Date
    
    Public Sub New()
        ID = 0
        Nome = ""
        Usuario = ""
        Email = ""
        Telefone = ""
        Comissao = 0
        Ativo = True
        DataCadastro = Date.Now
    End Sub
    
    Public Overrides Function ToString() As String
        Return Nome
    End Function
End Class
#End Region

#Region "Classe Configuracao"
''' <summary>
''' Configurações gerais do sistema
''' </summary>
Public Class ConfiguracaoSistema
    Public Property NomeMadeireira As String
    Public Property EnderecoMadeireira As String
    Public Property CidadeMadeireira As String
    Public Property CEPMadeireira As String
    Public Property TelefoneMadeireira As String
    Public Property CNPJMadeireira As String
    Public Property VendedorPadrao As String
    Public Property ConexaoBanco As String
    Public Property UsarBancoAccess As Boolean
    Public Property CaminhoBackup As String
    
    Public Sub New()
        CarregarConfiguracoes()
    End Sub
    
    Private Sub CarregarConfiguracoes()
        Try
            NomeMadeireira = If(ConfigurationManager.AppSettings("NomeMadeireira"), "Madeireira")
            EnderecoMadeireira = If(ConfigurationManager.AppSettings("EnderecoMadeireira"), "")
            CidadeMadeireira = If(ConfigurationManager.AppSettings("CidadeMadeireira"), "")
            CEPMadeireira = If(ConfigurationManager.AppSettings("CEPMadeireira"), "")
            TelefoneMadeireira = If(ConfigurationManager.AppSettings("TelefoneMadeireira"), "")
            CNPJMadeireira = If(ConfigurationManager.AppSettings("CNPJMadeireira"), "")
            VendedorPadrao = If(ConfigurationManager.AppSettings("VendedorPadrao"), "Sistema")
            ConexaoBanco = If(ConfigurationManager.AppSettings("ConexaoBanco"), "")
            UsarBancoAccess = Boolean.Parse(If(ConfigurationManager.AppSettings("UsarBancoAccess"), "false"))
            CaminhoBackup = If(ConfigurationManager.AppSettings("CaminhoBackup"), "C:\Backup\PDV\")
        Catch ex As Exception
            ' Valores padrão em caso de erro
            NomeMadeireira = "Madeireira Maria Luiza"
            VendedorPadrao = "Sistema"
            UsarBancoAccess = False
        End Try
    End Sub
End Class
#End Region

#Region "Classes para Talão"
''' <summary>
''' Dados específicos para geração de talão
''' </summary>
Public Class DadosTalao
    Public Property NomeCliente As String
    Public Property EnderecoCliente As String
    Public Property CEP As String
    Public Property Cidade As String
    Public Property Telefone As String
    Public Property FormaPagamento As String
    Public Property Vendedor As String
    Public Property NumeroTalao As String
    Public Property DataVenda As Date
    Public Property Produtos As List(Of ProdutoTalao)
    Public Property TotalGeral As Decimal
    Public Property Desconto As Decimal
    Public Property Observacoes As String
    
    Public Sub New()
        NomeCliente = ""
        EnderecoCliente = ""
        CEP = ""
        Cidade = ""
        Telefone = ""
        FormaPagamento = "À Vista"
        Vendedor = ""
        NumeroTalao = ""
        DataVenda = Date.Now
        Produtos = New List(Of ProdutoTalao)()
        TotalGeral = 0
        Desconto = 0
        Observacoes = ""
    End Sub
End Class

''' <summary>
''' Produto específico para talão
''' </summary>
Public Class ProdutoTalao
    Public Property Descricao As String
    Public Property Quantidade As Double
    Public Property Unidade As String
    Public Property PrecoUnitario As Decimal
    Public Property PrecoTotal As Decimal
    Public Property Observacoes As String
    
    Public Sub New()
        Descricao = ""
        Quantidade = 0
        Unidade = "UN"
        PrecoUnitario = 0
        PrecoTotal = 0
        Observacoes = ""
    End Sub
    
    Public ReadOnly Property SubTotal As Decimal
        Get
            Return PrecoUnitario * Quantidade
        End Get
    End Property
End Class
#End Region