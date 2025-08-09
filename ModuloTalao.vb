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
' ===== MÓDULO TALÃO - SISTEMA PRINCIPAL =====
' Módulo responsável por toda a lógica de geração de talões
' Madeireira Maria Luiza - Sistema PDV Integrado

Option Explicit

' Variáveis globais
Dim TalaoWorksheet As Worksheet
Dim DadosCliente As String
Dim DadosProdutos As String
Dim NumeroTalao As String

' ===== FUNÇÃO PRINCIPAL =====
Public Sub ProcessarTalaoCompleto()
    ' Função principal que processa todo o talão
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Inicializar processamento
    InicializarTalao
    
    ' Coletar dados se não foram fornecidos externamente
    If DadosCliente = """" Then
        ColetarDadosReais
    End If
    
    ' Gerar o talão completo
    GerarTalaoCompleto
    
    ' Configurar para impressão
    ConfigurarImpressaoCompleta
    
    ' Finalizar
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox ""Talão gerado com sucesso!"", vbInformation, ""Sistema PDV""
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox ""Erro ao processar talão: "" & Err.Description, vbCritical, ""Erro""
End Sub

' ===== INICIALIZAÇÃO =====
Private Sub InicializarTalao()
    ' Configurar planilha ativa
    Set TalaoWorksheet = ActiveSheet
    
    ' Gerar número do talão se não existir
    If NumeroTalao = """" Then
        NumeroTalao = Format(Now, ""yyyymmddhhmmss"")
    End If
    
    ' Limpar planilha
    TalaoWorksheet.Cells.Clear
    
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

' ===== COLETA DE DADOS =====
Private Sub ColetarDadosReais()
    ' Esta função coleta dados quando não fornecidos externamente
    ' Normalmente os dados vêm do VB.NET, mas esta é uma função de backup
    
    DadosCliente = ""João Silva - TESTE|Rua das Árvores, 123|55431-165|Paulista/PE|(81) 9876-5432""
    DadosProdutos = ""Tábua Pinus 2x4m|5|UN|25.00|125.00^^Ripão 3x3x3m|10|UN|15.00|150.00^^Compensado 18mm|2|M²|45.00|90.00""
End Sub

' ===== GERAÇÃO DO TALÃO =====
Public Sub GerarTalaoCompleto()
    On Error GoTo ErrorHandler
    
    ' Criar cabeçalho da empresa
    CriarCabecalhoEmpresa
    
    ' Criar dados do talão
    CriarDadosTalao
    
    ' Criar dados do cliente
    CriarDadosCliente
    
    ' Criar tabela de produtos
    CriarTabelaProdutos
    
    ' Criar totais e rodapé
    CriarTotaisRodape
    
    ' Criar segunda via
    CriarSegundaVia
    
    Exit Sub
    
ErrorHandler:
    MsgBox ""Erro ao gerar talão: "" & Err.Description, vbCritical, ""Erro""
End Sub

' ===== CABEÇALHO DA EMPRESA =====
Private Sub CriarCabecalhoEmpresa()
    With TalaoWorksheet
        ' Nome da empresa
        .Cells(1, 1).Value = ""MADEIREIRA MARIA LUIZA""
        .Cells(1, 1).Font.Size = 18
        .Cells(1, 1).Font.Bold = True
        .Range(""A1:G1"").Merge
        .Cells(1, 1).HorizontalAlignment = xlCenter
        
        ' Endereço
        .Cells(2, 1).Value = ""Rua Principal, 123 - Centro""
        .Cells(2, 1).HorizontalAlignment = xlCenter
        .Range(""A2:G2"").Merge
        
        ' Cidade e telefone
        .Cells(3, 1).Value = ""Paulista/PE - CEP: 53401-445 - Tel: (81) 3436-1234""
        .Cells(3, 1).HorizontalAlignment = xlCenter
        .Range(""A3:G3"").Merge
        
        ' CNPJ
        .Cells(4, 1).Value = ""CNPJ: 12.345.678/0001-90""
        .Cells(4, 1).HorizontalAlignment = xlCenter
        .Range(""A4:G4"").Merge
        
        ' Linha separadora
        .Range(""A5:G5"").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range(""A5:G5"").Borders(xlEdgeBottom).Weight = xlThick
    End With
End Sub

' ===== DADOS DO TALÃO =====
Private Sub CriarDadosTalao()
    With TalaoWorksheet
        ' Título do talão
        .Cells(7, 1).Value = ""TALÃO DE VENDA Nº: "" & NumeroTalao
        .Cells(7, 1).Font.Size = 14
        .Cells(7, 1).Font.Bold = True
        
        ' Data e hora
        .Cells(8, 1).Value = ""Data: "" & Format(Now, ""dd/mm/yyyy hh:mm"")
        .Cells(8, 1).Font.Bold = True
        
        ' Espaço
        .Cells(9, 1).Value = """"
    End With
End Sub

' ===== DADOS DO CLIENTE =====
Private Sub CriarDadosCliente()
    Dim ClienteArray As Variant
    
    If DadosCliente <> """" Then
        ClienteArray = Split(DadosCliente, ""|"")
    Else
        ' Dados padrão se não fornecidos
        ClienteArray = Array(""Cliente Teste"", ""Endereço Teste"", ""12345-678"", ""Cidade/UF"", ""(11) 1234-5678"")
    End If
    
    With TalaoWorksheet
        .Cells(10, 1).Value = ""CLIENTE:""
        .Cells(10, 1).Font.Bold = True
        .Cells(10, 2).Value = ClienteArray(0)
        
        .Cells(11, 1).Value = ""ENDEREÇO:""
        .Cells(11, 1).Font.Bold = True
        .Cells(11, 2).Value = ClienteArray(1)
        
        .Cells(12, 1).Value = ""CIDADE/CEP:""
        .Cells(12, 1).Font.Bold = True
        .Cells(12, 2).Value = ClienteArray(3) & "" - CEP: "" & ClienteArray(2)
        
        .Cells(13, 1).Value = ""TELEFONE:""
        .Cells(13, 1).Font.Bold = True
        .Cells(13, 2).Value = ClienteArray(4)
        
        ' Espaço
        .Cells(14, 1).Value = """"
    End With
End Sub

' ===== TABELA DE PRODUTOS =====
Private Sub CriarTabelaProdutos()
    Dim LinhaAtual As Integer
    Dim ProdutosArray As Variant
    Dim ProdutoAtual As Variant
    Dim TotalGeral As Double
    
    LinhaAtual = 15
    TotalGeral = 0
    
    With TalaoWorksheet
        ' Cabeçalho da tabela
        .Cells(LinhaAtual, 1).Value = ""DESCRIÇÃO""
        .Cells(LinhaAtual, 2).Value = ""QTD""
        .Cells(LinhaAtual, 3).Value = ""UN""
        .Cells(LinhaAtual, 4).Value = ""PREÇO UNIT.""
        .Cells(LinhaAtual, 5).Value = ""TOTAL""
        
        ' Formatar cabeçalho
        With .Range(""A"" & LinhaAtual & "":E"" & LinhaAtual)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Interior.Color = RGB(230, 230, 230)
        End With
        
        LinhaAtual = LinhaAtual + 1
        
        ' Processar produtos
        If DadosProdutos <> """" Then
            ProdutosArray = Split(DadosProdutos, ""^^"")
            
            Dim i As Integer
            For i = 0 To UBound(ProdutosArray)
                ProdutoAtual = Split(ProdutosArray(i), ""|"")
                
                .Cells(LinhaAtual, 1).Value = ProdutoAtual(0) ' Descrição
                .Cells(LinhaAtual, 2).Value = CDbl(ProdutoAtual(1)) ' Quantidade
                .Cells(LinhaAtual, 3).Value = ProdutoAtual(2) ' Unidade
                .Cells(LinhaAtual, 4).Value = CDbl(ProdutoAtual(3)) ' Preço unitário
                .Cells(LinhaAtual, 5).Value = CDbl(ProdutoAtual(4)) ' Total
                
                ' Formatar valores
                .Cells(LinhaAtual, 4).NumberFormat = ""R$ #,##0.00""
                .Cells(LinhaAtual, 5).NumberFormat = ""R$ #,##0.00""
                
                ' Bordas
                .Range(""A"" & LinhaAtual & "":E"" & LinhaAtual).Borders.LineStyle = xlContinuous
                
                TotalGeral = TotalGeral + CDbl(ProdutoAtual(4))
                LinhaAtual = LinhaAtual + 1
            Next i
        Else
            ' Produto exemplo se não fornecidos
            .Cells(LinhaAtual, 1).Value = ""Produto Exemplo""
            .Cells(LinhaAtual, 2).Value = 1
            .Cells(LinhaAtual, 3).Value = ""UN""
            .Cells(LinhaAtual, 4).Value = 10
            .Cells(LinhaAtual, 5).Value = 10
            .Cells(LinhaAtual, 4).NumberFormat = ""R$ #,##0.00""
            .Cells(LinhaAtual, 5).NumberFormat = ""R$ #,##0.00""
            .Range(""A"" & LinhaAtual & "":E"" & LinhaAtual).Borders.LineStyle = xlContinuous
            TotalGeral = 10
            LinhaAtual = LinhaAtual + 1
        End If
        
        ' Total geral
        LinhaAtual = LinhaAtual + 1
        .Cells(LinhaAtual, 4).Value = ""TOTAL GERAL:""
        .Cells(LinhaAtual, 4).Font.Bold = True
        .Cells(LinhaAtual, 4).HorizontalAlignment = xlRight
        .Cells(LinhaAtual, 5).Value = TotalGeral
        .Cells(LinhaAtual, 5).NumberFormat = ""R$ #,##0.00""
        .Cells(LinhaAtual, 5).Font.Bold = True
        .Range(""D"" & LinhaAtual & "":E"" & LinhaAtual).Borders.LineStyle = xlContinuous
        .Range(""D"" & LinhaAtual & "":E"" & LinhaAtual).Borders(xlEdgeTop).Weight = xlThick
    End With
End Sub

' ===== TOTAIS E RODAPÉ =====
Private Sub CriarTotaisRodape()
    Dim LinhaAtual As Integer
    LinhaAtual = TalaoWorksheet.UsedRange.Rows.Count + 2
    
    With TalaoWorksheet
        ' Forma de pagamento
        .Cells(LinhaAtual, 1).Value = ""FORMA DE PAGAMENTO: DINHEIRO""
        .Cells(LinhaAtual, 1).Font.Bold = True
        LinhaAtual = LinhaAtual + 1
        
        ' Vendedor
        .Cells(LinhaAtual, 1).Value = ""VENDEDOR: matheus-testuser3""
        .Cells(LinhaAtual, 1).Font.Bold = True
        LinhaAtual = LinhaAtual + 2
        
        ' Assinatura do cliente
        .Cells(LinhaAtual, 1).Value = ""CLIENTE: _________________________________""
        .Cells(LinhaAtual, 1).Font.Bold = True
        LinhaAtual = LinhaAtual + 1
        
        .Cells(LinhaAtual, 1).Value = ""           (NOME E ASSINATURA)""
        .Cells(LinhaAtual, 1).Font.Size = 8
    End With
End Sub

' ===== SEGUNDA VIA =====
Private Sub CriarSegundaVia()
    Dim LinhaInicial As Integer
    LinhaInicial = TalaoWorksheet.UsedRange.Rows.Count + 4
    
    With TalaoWorksheet
        ' Separador
        .Cells(LinhaInicial, 1).Value = ""✂️ --- CORTE AQUI - SEGUNDA VIA --- ✂️""
        .Cells(LinhaInicial, 1).HorizontalAlignment = xlCenter
        .Cells(LinhaInicial, 1).Font.Bold = True
        .Range(""A"" & LinhaInicial & "":E"" & LinhaInicial).Merge
        
        LinhaInicial = LinhaInicial + 2
        
        ' Cabeçalho resumido
        .Cells(LinhaInicial, 1).Value = ""MADEIREIRA MARIA LUIZA""
        .Cells(LinhaInicial, 1).Font.Size = 14
        .Cells(LinhaInicial, 1).Font.Bold = True
        .Cells(LinhaInicial, 1).HorizontalAlignment = xlCenter
        .Range(""A"" & LinhaInicial & "":E"" & LinhaInicial).Merge
        
        LinhaInicial = LinhaInicial + 1
        
        ' Dados resumidos
        .Cells(LinhaInicial, 1).Value = ""TALÃO Nº: "" & NumeroTalao & "" - "" & Format(Now, ""dd/mm/yyyy hh:mm"")
        .Cells(LinhaInicial, 1).Font.Bold = True
        
        LinhaInicial = LinhaInicial + 1
        
        ' Cliente resumido
        If DadosCliente <> """" Then
            Dim ClienteArray As Variant
            ClienteArray = Split(DadosCliente, ""|"")
            .Cells(LinhaInicial, 1).Value = ""CLIENTE: "" & ClienteArray(0)
        Else
            .Cells(LinhaInicial, 1).Value = ""CLIENTE: Cliente Teste""
        End If
        
        LinhaInicial = LinhaInicial + 1
        
        ' Total resumido (calcular novamente)
        Dim TotalResumido As Double
        TotalResumido = 0
        If DadosProdutos <> """" Then
            Dim ProdutosArray As Variant
            Dim ProdutoAtual As Variant
            ProdutosArray = Split(DadosProdutos, ""^^"")
            
            Dim i As Integer
            For i = 0 To UBound(ProdutosArray)
                ProdutoAtual = Split(ProdutosArray(i), ""|"")
                TotalResumido = TotalResumido + CDbl(ProdutoAtual(4))
            Next i
        Else
            TotalResumido = 10
        End If
        
        .Cells(LinhaInicial, 1).Value = ""TOTAL: "" & Format(TotalResumido, ""R$ #,##0.00"")
        .Cells(LinhaInicial, 1).Font.Bold = True
        .Cells(LinhaInicial, 1).Font.Size = 12
        
        LinhaInicial = LinhaInicial + 1
        .Cells(LinhaInicial, 1).Value = ""PAGAMENTO: DINHEIRO""
        LinhaInicial = LinhaInicial + 1
        .Cells(LinhaInicial, 1).Value = ""VENDEDOR: matheus-testuser3""
    End With
End Sub

' ===== CONFIGURAÇÃO DE IMPRESSÃO =====
Public Sub ConfigurarImpressaoCompleta()
    On Error GoTo ErrorHandler
    
    With TalaoWorksheet
        ' Ajustar colunas
        .Columns(""A"").ColumnWidth = 35  ' Descrição
        .Columns(""B"").ColumnWidth = 8   ' Quantidade  
        .Columns(""C"").ColumnWidth = 6   ' Unidade
        .Columns(""D"").ColumnWidth = 12  ' Preço unitário
        .Columns(""E"").ColumnWidth = 12  ' Total
        
        ' Configurar página
        With .PageSetup
            .PrintArea = ""A1:E"" & .UsedRange.Rows.Count
            .FitToPagesWide = 1
            .FitToPagesTall = False
            .CenterHorizontally = True
            .CenterVertically = False
            .PrintTitleRows = ""1:15""
        End With
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox ""Erro ao configurar impressão: "" & Err.Description, vbCritical, ""Erro""
End Sub

' ===== FUNÇÕES AUXILIARES =====

' Função para definir dados do cliente externamente
Public Sub DefinirDadosCliente(Nome As String, Endereco As String, CEP As String, Cidade As String, Telefone As String)
    DadosCliente = Nome & ""|"" & Endereco & ""|"" & CEP & ""|"" & Cidade & ""|"" & Telefone
End Sub

' Função para adicionar produto
Public Sub AdicionarProduto(Descricao As String, Quantidade As Double, Unidade As String, PrecoUnit As Double)
    Dim Total As Double
    Total = Quantidade * PrecoUnit
    
    If DadosProdutos = """" Then
        DadosProdutos = Descricao & ""|"" & Quantidade & ""|"" & Unidade & ""|"" & PrecoUnit & ""|"" & Total
    Else
        DadosProdutos = DadosProdutos & ""^^"" & Descricao & ""|"" & Quantidade & ""|"" & Unidade & ""|"" & PrecoUnit & ""|"" & Total
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

"
    End Function
End Class