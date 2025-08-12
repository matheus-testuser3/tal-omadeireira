''' <summary>
''' Módulo VBA para geração de talões - Sistema principal
''' Este código será injetado dinamicamente no Excel
''' </summary>
Public Class ModuloTalao

    ''' <summary>
    ''' Retorna o código VBA completo para geração de talões
    ''' </summary>
    Public Function ObterCodigoVBA() As String
        Return "
' ===== MÓDULO TALÃO - SISTEMA DE MAPEAMENTO =====
' Módulo responsável por mapeamento de células em vez de impressão
' Madeireira Maria Luiza - Sistema PDV Integrado

Option Explicit

' Variáveis globais para mapeamento
Dim TalaoWorksheet As Worksheet
Dim DadosCliente As String
Dim DadosProdutos As String
Dim NumeroTalao As String
Dim MapaCelulas As Object

' ===== FUNÇÃO PRINCIPAL =====
Public Sub ProcessarTalaoCompleto()
    ' Função principal que mapeia dados em células específicas
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Inicializar mapeamento
    InicializarMapeamento
    
    ' Coletar dados se não foram fornecidos externamente
    If DadosCliente = """" Then
        ColetarDadosReais
    End If
    
    ' Mapear dados nas células
    MapearDadosNasCelulas
    
    ' Aplicar formatação inteligente
    AplicarFormatacaoInteligente
    
    ' Configurar visualização
    ConfigurarVisualizacao
    
    ' Finalizar
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox ""Dados mapeados com sucesso na planilha!"", vbInformation, ""Sistema PDV""
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox ""Erro ao mapear dados: "" & Err.Description, vbCritical, ""Erro""
End Sub

' ===== INICIALIZAÇÃO DO MAPEAMENTO =====
Private Sub InicializarMapeamento()
    ' Configurar planilha ativa
    Set TalaoWorksheet = ActiveSheet
    
    ' Gerar número do talão se não existir
    If NumeroTalao = """" Then
        NumeroTalao = Format(Now, ""yyyymmddhhmmss"")
    End If
    
    ' Criar dicionário de mapeamento de células
    Set MapaCelulas = CreateObject(""Scripting.Dictionary"")
    DefinirMapeamentoCelulas
    
    ' Configurar formatação básica
    With TalaoWorksheet
        .PageSetup.PaperSize = xlPaperA4
        .PageSetup.Orientation = xlPortrait
        .PageSetup.LeftMargin = Application.InchesToPoints(0.5)
        .PageSetup.RightMargin = Application.InchesToPoints(0.5)
        .PageSetup.TopMargin = Application.InchesToPoints(0.5)
        .PageSetup.BottomMargin = Application.InchesToPoints(0.5)
    End With
End Sub

' ===== DEFINIR MAPEAMENTO DE CÉLULAS =====
Private Sub DefinirMapeamentoCelulas()
    ' Dados da empresa
    MapaCelulas.Add ""NOME_EMPRESA"", ""A1""
    MapaCelulas.Add ""ENDERECO_EMPRESA"", ""A2""
    MapaCelulas.Add ""CIDADE_EMPRESA"", ""A3""
    MapaCelulas.Add ""TELEFONE_EMPRESA"", ""A4""
    MapaCelulas.Add ""CNPJ_EMPRESA"", ""A5""
    
    ' Dados do talão
    MapaCelulas.Add ""NUMERO_TALAO"", ""F7""
    MapaCelulas.Add ""DATA_TALAO"", ""A8""
    
    ' Dados do cliente
    MapaCelulas.Add ""NOME_CLIENTE"", ""B10""
    MapaCelulas.Add ""ENDERECO_CLIENTE"", ""B11""
    MapaCelulas.Add ""CIDADE_CEP_CLIENTE"", ""B12""
    MapaCelulas.Add ""TELEFONE_CLIENTE"", ""B13""
    
    ' Totais e informações finais
    MapaCelulas.Add ""FORMA_PAGAMENTO"", ""B29""
    MapaCelulas.Add ""VENDEDOR"", ""B30""
    MapaCelulas.Add ""TOTAL_GERAL"", ""E27""
    MapaCelulas.Add ""ASSINATURA_CLIENTE"", ""A32""
End Sub

' ===== COLETA DE DADOS =====
Private Sub ColetarDadosReais()
    ' Esta função coleta dados quando não fornecidos externamente
    ' Dados padrão para teste do sistema de mapeamento
    
    DadosCliente = ""João Silva - TESTE MAPEAMENTO|Rua das Árvores, 123|55431-165|Paulista/PE|(81) 9876-5432""
    DadosProdutos = ""Tábua Pinus 2x4m|5|UN|25.00|125.00|25000^^Ripão 3x3x3m|10|UN|15.00|150.00|15000^^Compensado 18mm|2|M²|45.00|90.00|45000""
End Sub

' ===== MAPEAMENTO DE DADOS =====
Public Sub MapearDadosNasCelulas()
    On Error GoTo ErrorHandler
    
    ' Criar estrutura básica
    CriarEstruturaMapeada
    
    ' Mapear dados da empresa
    MapearDadosEmpresa
    
    ' Mapear dados do talão
    MapearDadosTalao
    
    ' Mapear dados do cliente
    MapearDadosCliente
    
    ' Mapear produtos
    MapearProdutos
    
    ' Mapear totais e informações finais
    MapearTotaisFinais
    
    ' Criar segunda via mapeada
    CriarSegundaViaMapeada
    
    Exit Sub
    
ErrorHandler:
    MsgBox ""Erro ao mapear dados: "" & Err.Description, vbCritical, ""Erro""
End Sub

' ===== ESTRUTURA MAPEADA =====
Private Sub CriarEstruturaMapeada()
    With TalaoWorksheet
        ' Configurar larguras das colunas
        .Columns(""A"").ColumnWidth = 35  ' Descrição
        .Columns(""B"").ColumnWidth = 8   ' Quantidade  
        .Columns(""C"").ColumnWidth = 6   ' Unidade
        .Columns(""D"").ColumnWidth = 12  ' Preço unitário
        .Columns(""E"").ColumnWidth = 12  ' Total
        .Columns(""F"").ColumnWidth = 15  ' Extra
        
        ' Labels dos campos
        .Cells(10, 1).Value = ""CLIENTE:""
        .Cells(11, 1).Value = ""ENDEREÇO:""
        .Cells(12, 1).Value = ""CIDADE/CEP:""
        .Cells(13, 1).Value = ""TELEFONE:""
        .Cells(29, 1).Value = ""FORMA DE PAGAMENTO:""
        .Cells(30, 1).Value = ""VENDEDOR:""
        .Cells(27, 4).Value = ""TOTAL GERAL:""
        
        ' Cabeçalho da tabela de produtos
        .Cells(15, 1).Value = ""DESCRIÇÃO""
        .Cells(15, 2).Value = ""QTD""
        .Cells(15, 3).Value = ""UN""
        .Cells(15, 4).Value = ""PREÇO UNIT.""
        .Cells(15, 5).Value = ""TOTAL""
        .Cells(15, 6).Value = ""VISUAL""
    End With
End Sub

' ===== MAPEAR DADOS DA EMPRESA =====
Private Sub MapearDadosEmpresa()
    EscreverCelulaMapeada ""NOME_EMPRESA"", ""MADEIREIRA MARIA LUIZA""
    EscreverCelulaMapeada ""ENDERECO_EMPRESA"", ""Rua Principal, 123 - Centro""
    EscreverCelulaMapeada ""CIDADE_EMPRESA"", ""Paulista/PE - CEP: 53401-445""
    EscreverCelulaMapeada ""TELEFONE_EMPRESA"", ""Tel: (81) 3436-1234""
    EscreverCelulaMapeada ""CNPJ_EMPRESA"", ""CNPJ: 12.345.678/0001-90""
End Sub

' ===== MAPEAR DADOS DO TALÃO =====
Private Sub MapearDadosTalao()
    TalaoWorksheet.Cells(7, 1).Value = ""TALÃO DE VENDA Nº:""
    EscreverCelulaMapeada ""NUMERO_TALAO"", NumeroTalao
    EscreverCelulaMapeada ""DATA_TALAO"", ""Data: "" & Format(Now, ""dd/mm/yyyy hh:mm"")
End Sub

' ===== MAPEAR DADOS DO CLIENTE =====
Private Sub MapearDadosCliente()
    Dim ClienteArray As Variant
    
    If DadosCliente <> """" Then
        ClienteArray = Split(DadosCliente, ""|"")
    Else
        ClienteArray = Array(""Cliente Teste"", ""Endereço Teste"", ""12345-678"", ""Cidade/UF"", ""(11) 1234-5678"")
    End If
    
    EscreverCelulaMapeada ""NOME_CLIENTE"", ClienteArray(0)
    EscreverCelulaMapeada ""ENDERECO_CLIENTE"", ClienteArray(1)
    EscreverCelulaMapeada ""CIDADE_CEP_CLIENTE"", ClienteArray(3) & "" - CEP: "" & ClienteArray(2)
    EscreverCelulaMapeada ""TELEFONE_CLIENTE"", ClienteArray(4)
End Sub

' ===== MAPEAR PRODUTOS =====
Private Sub MapearProdutos()
    Dim LinhaAtual As Integer
    Dim ProdutosArray As Variant
    Dim ProdutoAtual As Variant
    Dim TotalGeral As Double
    
    LinhaAtual = 16
    TotalGeral = 0
    
    If DadosProdutos <> """" Then
        ProdutosArray = Split(DadosProdutos, ""^^"")
        
        Dim i As Integer
        For i = 0 To UBound(ProdutosArray)
            ProdutoAtual = Split(ProdutosArray(i), ""|"")
            
            With TalaoWorksheet
                .Cells(LinhaAtual, 1).Value = ProdutoAtual(0) ' Descrição
                .Cells(LinhaAtual, 2).Value = CDbl(ProdutoAtual(1)) ' Quantidade
                .Cells(LinhaAtual, 3).Value = ProdutoAtual(2) ' Unidade
                .Cells(LinhaAtual, 4).Value = CDbl(ProdutoAtual(3)) ' Preço unitário
                .Cells(LinhaAtual, 5).Value = CDbl(ProdutoAtual(4)) ' Total
                
                ' Preço visual (multiplicador 1000)
                If UBound(ProdutoAtual) >= 5 Then
                    .Cells(LinhaAtual, 6).Value = CDbl(ProdutoAtual(5)) ' Preço visual
                Else
                    .Cells(LinhaAtual, 6).Value = CDbl(ProdutoAtual(3)) * 1000
                End If
                
                ' Formatar valores
                .Cells(LinhaAtual, 4).NumberFormat = ""R$ #,##0.00""
                .Cells(LinhaAtual, 5).NumberFormat = ""R$ #,##0.00""
                .Cells(LinhaAtual, 6).NumberFormat = ""#,##0""
                
                ' Bordas
                .Range(""A"" & LinhaAtual & "":F"" & LinhaAtual).Borders.LineStyle = xlContinuous
            End With
            
            TotalGeral = TotalGeral + CDbl(ProdutoAtual(4))
            LinhaAtual = LinhaAtual + 1
        Next i
    End If
    
    ' Mapear total geral
    EscreverCelulaMapeada ""TOTAL_GERAL"", TotalGeral
    With TalaoWorksheet.Range(MapaCelulas(""TOTAL_GERAL""))
        .NumberFormat = ""R$ #,##0.00""
        .Font.Bold = True
    End With
End Sub

' ===== MAPEAR TOTAIS E FINAIS =====
Private Sub MapearTotaisFinais()
    EscreverCelulaMapeada ""FORMA_PAGAMENTO"", ""DINHEIRO""
    EscreverCelulaMapeada ""VENDEDOR"", ""matheus-testuser3""
    EscreverCelulaMapeada ""ASSINATURA_CLIENTE"", ""CLIENTE: _________________________________""
    
    TalaoWorksheet.Cells(33, 1).Value = ""           (NOME E ASSINATURA)""
End Sub

' ===== SEGUNDA VIA MAPEADA =====
Private Sub CriarSegundaViaMapeada()
    Dim LinhaSegundaVia As Integer
    LinhaSegundaVia = 36
    
    With TalaoWorksheet
        ' Separador
        .Cells(LinhaSegundaVia, 1).Value = ""✂️ --- CORTE AQUI - SEGUNDA VIA --- ✂️""
        .Range(""A"" & LinhaSegundaVia & "":F"" & LinhaSegundaVia).Merge
        .Cells(LinhaSegundaVia, 1).HorizontalAlignment = xlCenter
        .Cells(LinhaSegundaVia, 1).Font.Bold = True
        
        LinhaSegundaVia = LinhaSegundaVia + 2
        
        ' Dados resumidos mapeados
        .Cells(LinhaSegundaVia, 1).Value = ""MADEIREIRA MARIA LUIZA""
        .Range(""A"" & LinhaSegundaVia & "":F"" & LinhaSegundaVia).Merge
        .Cells(LinhaSegundaVia, 1).Font.Size = 14
        .Cells(LinhaSegundaVia, 1).Font.Bold = True
        .Cells(LinhaSegundaVia, 1).HorizontalAlignment = xlCenter
        
        LinhaSegundaVia = LinhaSegundaVia + 1
        .Cells(LinhaSegundaVia, 1).Value = ""TALÃO Nº: "" & NumeroTalao & "" - "" & Format(Now, ""dd/mm/yyyy hh:mm"")
        
        LinhaSegundaVia = LinhaSegundaVia + 1
        Dim ClienteNome As String
        If DadosCliente <> """" Then
            ClienteNome = Split(DadosCliente, ""|"")(0)
        Else
            ClienteNome = ""Cliente Teste""
        End If
        .Cells(LinhaSegundaVia, 1).Value = ""CLIENTE: "" & ClienteNome
        
        LinhaSegundaVia = LinhaSegundaVia + 1
        .Cells(LinhaSegundaVia, 1).Value = ""TOTAL: [TOTAL_RESUMIDO]""
        .Cells(LinhaSegundaVia, 1).Font.Bold = True
        
        LinhaSegundaVia = LinhaSegundaVia + 1
        .Cells(LinhaSegundaVia, 1).Value = ""PAGAMENTO: DINHEIRO""
        
        LinhaSegundaVia = LinhaSegundaVia + 1
        .Cells(LinhaSegundaVia, 1).Value = ""VENDEDOR: matheus-testuser3""
    End With
End Sub

' ===== FORMATAÇÃO INTELIGENTE =====
Public Sub AplicarFormatacaoInteligente()
    On Error GoTo ErrorHandler
    
    With TalaoWorksheet
        ' Formatação do cabeçalho da empresa
        With .Range(""A1:F1"")
            .Merge
            .Font.Size = 18
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
        
        ' Formatação do título do talão
        With .Cells(7, 1)
            .Font.Size = 14
            .Font.Bold = True
        End With
        
        ' Formatação dos labels
        With .Range(""A10:A13,A29:A30"")
            .Font.Bold = True
        End With
        
        With .Cells(27, 4)
            .Font.Bold = True
            .HorizontalAlignment = xlRight
        End With
        
        ' Formatação da tabela de produtos
        With .Range(""A15:F15"")
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Interior.Color = RGB(230, 230, 230)
        End With
        
        ' Bordas da tabela
        .Range(""A15:F27"").Borders.LineStyle = xlContinuous
        
        ' Formatação especial para total geral
        With .Range(""D27:E27"")
            .Font.Bold = True
            .Borders.LineStyle = xlContinuous
            .Borders(xlEdgeTop).Weight = xlThick
        End With
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox ""Erro ao aplicar formatação: "" & Err.Description, vbCritical, ""Erro""
End Sub

' ===== CONFIGURAÇÃO DE VISUALIZAÇÃO =====
Public Sub ConfigurarVisualizacao()
    On Error GoTo ErrorHandler
    
    With TalaoWorksheet
        With .PageSetup
            .PrintArea = ""A1:F"" & .UsedRange.Rows.Count
            .FitToPagesWide = 1
            .FitToPagesTall = False
            .CenterHorizontally = True
            .PrintTitleRows = ""1:15""
        End With
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox ""Erro ao configurar visualização: "" & Err.Description, vbCritical, ""Erro""
End Sub

' ===== FUNÇÕES AUXILIARES =====

' Função para escrever em célula mapeada
Private Sub EscreverCelulaMapeada(Chave As String, Valor As Variant)
    If MapaCelulas.Exists(Chave) Then
        TalaoWorksheet.Range(MapaCelulas(Chave)).Value = Valor
    End If
End Sub

' Função para definir dados do cliente externamente
Public Sub DefinirDadosCliente(Nome As String, Endereco As String, CEP As String, Cidade As String, Telefone As String)
    DadosCliente = Nome & ""|"" & Endereco & ""|"" & CEP & ""|"" & Cidade & ""|"" & Telefone
End Sub

' Função para adicionar produto
Public Sub AdicionarProduto(Descricao As String, Quantidade As Double, Unidade As String, PrecoUnit As Double, PrecoVisual As Double)
    Dim Total As Double
    Total = Quantidade * PrecoUnit
    
    If DadosProdutos = """" Then
        DadosProdutos = Descricao & ""|"" & Quantidade & ""|"" & Unidade & ""|"" & PrecoUnit & ""|"" & Total & ""|"" & PrecoVisual
    Else
        DadosProdutos = DadosProdutos & ""^^"" & Descricao & ""|"" & Quantidade & ""|"" & Unidade & ""|"" & PrecoUnit & ""|"" & Total & ""|"" & PrecoVisual
    End If
End Sub

' Função para limpar dados
Public Sub LimparDados()
    DadosCliente = """"
    DadosProdutos = """"
    NumeroTalao = """"
End Sub

' Função para definir número do talão
Public Sub DefinirNumeroTalao(Numero As String)
    NumeroTalao = Numero
End Sub

' Função para obter informações do mapeamento
Public Function ObterInfoMapeamento() As String
    Dim Info As String
    Info = ""=== MAPEAMENTO DE CÉLULAS ==="""
    
    Dim Chave As Variant
    For Each Chave In MapaCelulas.Keys
        Info = Info & vbCrLf & Chave & "": "" & MapaCelulas(Chave)
    Next Chave
    
    ObterInfoMapeamento = Info
End Function

"
    End Function
End Class