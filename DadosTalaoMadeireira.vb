''' <summary>
''' Estruturas de dados específicas para backup e restauração de talões - Madeireira Maria Luiza
''' Data/Hora: 2025-08-14 11:16:26 UTC
''' Usuário: matheus-testuser3
''' Sistema de Backup e Restauração de Talões
''' </summary>

Imports Newtonsoft.Json
Imports System.ComponentModel

''' <summary>
''' Classe principal para dados de talão específicos da madeireira
''' Inclui propriedades calculadas e serialização JSON para backup
''' </summary>
<JsonObject(MemberSerialization.OptIn)>
Public Class DadosTalaoMadeireira
    
    ' === INFORMAÇÕES DO TALÃO ===
    <JsonProperty("numeroTalao")>
    Public Property NumeroTalao As String
    
    <JsonProperty("dataEmissao")>
    Public Property DataEmissao As DateTime
    
    <JsonProperty("dataVencimento")>
    Public Property DataVencimento As DateTime?
    
    ' === DADOS DO CLIENTE ===
    <JsonProperty("nomeCliente")>
    Public Property NomeCliente As String
    
    <JsonProperty("enderecoCliente")>
    Public Property EnderecoCliente As String
    
    <JsonProperty("cep")>
    Public Property CEP As String
    
    <JsonProperty("cidade")>
    Public Property Cidade As String
    
    <JsonProperty("telefone")>
    Public Property Telefone As String
    
    <JsonProperty("cpfCnpj")>
    Public Property CPF_CNPJ As String
    
    ' === PRODUTOS ESPECÍFICOS DE MADEIREIRA ===
    <JsonProperty("produtos")>
    Public Property Produtos As List(Of ProdutoTalaoMadeireira)
    
    ' === INFORMAÇÕES COMERCIAIS ===
    <JsonProperty("formaPagamento")>
    Public Property FormaPagamento As String
    
    <JsonProperty("vendedor")>
    Public Property Vendedor As String
    
    <JsonProperty("observacoes")>
    Public Property Observacoes As String
    
    <JsonProperty("desconto")>
    Public Property Desconto As Decimal
    
    <JsonProperty("tipoDesconto")>
    Public Property TipoDesconto As String ' "PERCENTUAL" ou "VALOR"
    
    ' === METADADOS DE BACKUP ===
    <JsonProperty("origemBackup")>
    Public Property OrigemBackup As String ' Caminho do arquivo original
    
    <JsonProperty("formatoDetectado")>
    Public Property FormatoDetectado As String ' "MADEIREIRA" ou "GENERICO"
    
    <JsonProperty("dataImportacao")>
    Public Property DataImportacao As DateTime
    
    <JsonProperty("usuarioImportacao")>
    Public Property UsuarioImportacao As String
    
    ' === PROPRIEDADES CALCULADAS ===
    <JsonIgnore>
    Public ReadOnly Property SubTotal As Decimal
        Get
            If Produtos Is Nothing Then Return 0
            Return Produtos.Sum(Function(p) p.ValorTotal)
        End Get
    End Property
    
    <JsonIgnore>
    Public ReadOnly Property ValorDesconto As Decimal
        Get
            If TipoDesconto = "PERCENTUAL" Then
                Return SubTotal * (Desconto / 100)
            Else
                Return Desconto
            End If
        End Get
    End Property
    
    <JsonIgnore>
    Public ReadOnly Property ValorTotal As Decimal
        Get
            Return SubTotal - ValorDesconto
        End Get
    End Property
    
    <JsonIgnore>
    Public ReadOnly Property QuantidadeTotalProdutos As Integer
        Get
            If Produtos Is Nothing Then Return 0
            Return Produtos.Count
        End Get
    End Property
    
    <JsonIgnore>
    Public ReadOnly Property ResumoDescricao As String
        Get
            Return $"Talão #{NumeroTalao} - {NomeCliente} - {ValorTotal:C2}"
        End Get
    End Property
    
    ' === CONSTRUTOR ===
    Public Sub New()
        Produtos = New List(Of ProdutoTalaoMadeireira)()
        DataEmissao = DateTime.Now
        UsuarioImportacao = "matheus-testuser3"
        DataImportacao = DateTime.UtcNow
        TipoDesconto = "VALOR"
        Desconto = 0
    End Sub
    
    ' === MÉTODOS DE VALIDAÇÃO ===
    Public Function ValidarDados() As List(Of String)
        Dim erros As New List(Of String)()
        
        If String.IsNullOrWhiteSpace(NomeCliente) Then
            erros.Add("Nome do cliente é obrigatório")
        End If
        
        If String.IsNullOrWhiteSpace(NumeroTalao) Then
            erros.Add("Número do talão é obrigatório")
        End If
        
        If Produtos Is Nothing OrElse Produtos.Count = 0 Then
            erros.Add("Pelo menos um produto deve ser informado")
        End If
        
        If Not String.IsNullOrWhiteSpace(CEP) AndAlso Not System.Text.RegularExpressions.Regex.IsMatch(CEP, "^\d{5}-?\d{3}$") Then
            erros.Add("Formato de CEP inválido")
        End If
        
        Return erros
    End Function
End Class

''' <summary>
''' Classe para produtos específicos de madeireira com unidades de medida adequadas
''' </summary>
<JsonObject(MemberSerialization.OptIn)>
Public Class ProdutoTalaoMadeireira
    
    ' === INFORMAÇÕES BÁSICAS ===
    <JsonProperty("descricao")>
    Public Property Descricao As String
    
    <JsonProperty("quantidade")>
    Public Property Quantidade As Decimal
    
    <JsonProperty("unidade")>
    Public Property Unidade As String ' m³, m², m, pc, kg, ton
    
    <JsonProperty("precoUnitario")>
    Public Property PrecoUnitario As Decimal
    
    ' === ESPECIFICAÇÕES DE MADEIRA ===
    <JsonProperty("tipoMadeira")>
    Public Property TipoMadeira As String ' massaranduba, ipê, peroba, pinus
    
    <JsonProperty("dimensoes")>
    Public Property Dimensoes As String ' Ex: "6x6cm", "4x12cm", "2x30cm"
    
    <JsonProperty("comprimento")>
    Public Property Comprimento As String ' Ex: "3m", "4m", "5m", "6m"
    
    <JsonProperty("categoria")>
    Public Property Categoria As String ' barrotes, cabros, tábuas, vigas
    
    <JsonProperty("tratamento")>
    Public Property Tratamento As String ' autoclavado, natural, impregnado
    
    <JsonProperty("qualidade")>
    Public Property Qualidade As String ' 1ª, 2ª, 3ª, construção
    
    ' === CÁLCULOS ===
    <JsonIgnore>
    Public ReadOnly Property ValorTotal As Decimal
        Get
            Return Quantidade * PrecoUnitario
        End Get
    End Property
    
    <JsonIgnore>
    Public ReadOnly Property DescricaoCompleta As String
        Get
            Dim partes As New List(Of String)()
            
            If Not String.IsNullOrWhiteSpace(Categoria) Then partes.Add(Categoria)
            If Not String.IsNullOrWhiteSpace(TipoMadeira) Then partes.Add(TipoMadeira)
            If Not String.IsNullOrWhiteSpace(Dimensoes) Then partes.Add(Dimensoes)
            If Not String.IsNullOrWhiteSpace(Comprimento) Then partes.Add(Comprimento)
            If Not String.IsNullOrWhiteSpace(Qualidade) Then partes.Add($"({Qualidade})")
            
            Dim descCompleta = String.Join(" ", partes)
            
            If String.IsNullOrWhiteSpace(descCompleta) Then
                Return Descricao
            Else
                Return If(String.IsNullOrWhiteSpace(Descricao), descCompleta, $"{Descricao} - {descCompleta}")
            End If
        End Get
    End Property
    
    ' === CONSTRUTOR ===
    Public Sub New()
        Quantidade = 1
        PrecoUnitario = 0
        Unidade = "UN"
    End Sub
    
    Public Sub New(descricao As String, quantidade As Decimal, unidade As String, precoUnitario As Decimal)
        Me.Descricao = descricao
        Me.Quantidade = quantidade
        Me.Unidade = unidade
        Me.PrecoUnitario = precoUnitario
    End Sub
    
    ' === MÉTODOS DE VALIDAÇÃO ===
    Public Function ValidarProduto() As List(Of String)
        Dim erros As New List(Of String)()
        
        If String.IsNullOrWhiteSpace(Descricao) AndAlso String.IsNullOrWhiteSpace(Categoria) Then
            erros.Add("Descrição ou categoria do produto é obrigatória")
        End If
        
        If Quantidade <= 0 Then
            erros.Add("Quantidade deve ser maior que zero")
        End If
        
        If PrecoUnitario < 0 Then
            erros.Add("Preço unitário não pode ser negativo")
        End If
        
        If String.IsNullOrWhiteSpace(Unidade) Then
            erros.Add("Unidade de medida é obrigatória")
        End If
        
        Return erros
    End Function
End Class

''' <summary>
''' Enumeração para tipos de formato de backup detectados
''' </summary>
Public Enum TipoFormatoBackup
    Madeireira
    Generico
    Desconhecido
End Enum

''' <summary>
''' Enumeração para unidades de medida específicas da madeireira
''' </summary>
Public Enum UnidadeMedidaMadeireira
    MetroCubico ' m³
    MetroQuadrado ' m²
    MetroLinear ' m
    Pecas ' pc
    Quilogramas ' kg
    Toneladas ' ton
    Unidades ' un
End Enum

''' <summary>
''' Classe para configurações de backup específicas da madeireira
''' </summary>
Public Class ConfiguracaoBackupMadeireira
    Public Property CaminhoBackupsImportados As String
    Public Property CaminhoTaloesGerados As String
    Public Property CaminhoBackupJSON As String
    Public Property FormatoDataBackup As String
    Public Property PrefixoArquivoBackup As String
    Public Property ManterHistoricoBackups As Boolean
    Public Property DiasRetencaoBackups As Integer
    
    Public Sub New()
        CaminhoBackupsImportados = "Backups"
        CaminhoTaloesGerados = "Taloes"
        CaminhoBackupJSON = "BackupJSON"
        FormatoDataBackup = "yyyy-MM-dd_HH-mm-ss"
        PrefixoArquivoBackup = "backup_talao_"
        ManterHistoricoBackups = True
        DiasRetencaoBackups = 90
    End Sub
End Class