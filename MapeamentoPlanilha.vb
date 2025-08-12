Imports Microsoft.Office.Interop.Excel
Imports System.Configuration

''' <summary>
''' Sistema de mapeamento de planilha - Escreve dados em células específicas
''' Substitui o sistema de impressão por escrita inteligente em Excel
''' </summary>
Public Class MapeamentoPlanilha

    ' Aplicação Excel e objetos
    Private xlApp As Application
    Private xlWorkbook As Workbook
    Private xlWorksheet As Worksheet

    ' Mapeamento de células
    Private ReadOnly mapaCelulas As Dictionary(Of String, String)

    ' Status do sistema
    Public Property StatusProcessamento As String = "AGUARDANDO"

    ''' <summary>
    ''' Construtor - Inicializa o mapeamento de células
    ''' </summary>
    Public Sub New()
        ' Inicializar mapeamento de células
        mapaCelulas = New Dictionary(Of String, String) From {
            {"NOME_EMPRESA", "A1"},
            {"ENDERECO_EMPRESA", "A2"},
            {"CIDADE_EMPRESA", "A3"},
            {"TELEFONE_EMPRESA", "A4"},
            {"CNPJ_EMPRESA", "A5"},
            {"NUMERO_TALAO", "F7"},
            {"DATA_TALAO", "A8"},
            {"NOME_CLIENTE", "B10"},
            {"ENDERECO_CLIENTE", "B11"},
            {"CIDADE_CEP_CLIENTE", "B12"},
            {"TELEFONE_CLIENTE", "B13"},
            {"FORMA_PAGAMENTO", "B29"},
            {"VENDEDOR", "B30"},
            {"ASSINATURA_CLIENTE", "A32"},
            {"TOTAL_GERAL", "E27"}
        }
    End Sub

    ''' <summary>
    ''' Escreve dados na planilha mapeada (função principal)
    ''' </summary>
    ''' <param name="dados">Dados do talão a serem escritos</param>
    Public Sub EscreverNaPlanilhaMapeada(dados As DadosTalao)
        Try
            StatusProcessamento = "INICIANDO"
            
            ' Abrir Excel e configurar planilha
            AbrirExcelEConfigurar()
            
            ' Criar template inteligente
            CriarTemplateInteligente()
            
            ' Escrever dados da empresa
            EscreverDadosEmpresa()
            
            ' Escrever dados do talão
            EscreverDadosTalao(dados)
            
            ' Escrever dados do cliente
            EscreverDadosCliente(dados)
            
            ' Escrever produtos com formatação
            EscreverProdutos(dados)
            
            ' Escrever totais e informações finais
            EscreverTotaisEFinalizacao(dados)
            
            ' Aplicar formatação automática
            AplicarFormatacaoAutomatica()
            
            ' Configurar para visualização/impressão
            ConfigurarVisualizacao()
            
            StatusProcessamento = "CONCLUIDO"
            
        Catch ex As Exception
            StatusProcessamento = $"ERRO: {ex.Message}"
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' Abre Excel e configura planilha base
    ''' </summary>
    Private Sub AbrirExcelEConfigurar()
        StatusProcessamento = "ABRINDO_EXCEL"
        
        xlApp = New Application()
        xlApp.Visible = Boolean.Parse(ConfigurationManager.AppSettings("ExcelVisivel") OrElse "false")
        xlApp.DisplayAlerts = False
        xlApp.ScreenUpdating = False
        
        xlWorkbook = xlApp.Workbooks.Add()
        xlWorksheet = xlWorkbook.ActiveSheet
        xlWorksheet.Name = "Talao_" & Date.Now.ToString("HHmmss")
        
        ' Configurar página
        With xlWorksheet.PageSetup
            .PaperSize = XlPaperSize.xlPaperA4
            .Orientation = XlPageOrientation.xlPortrait
            .LeftMargin = xlApp.InchesToPoints(0.5)
            .RightMargin = xlApp.InchesToPoints(0.5)
            .TopMargin = xlApp.InchesToPoints(0.5)
            .BottomMargin = xlApp.InchesToPoints(0.5)
        End With
    End Sub

    ''' <summary>
    ''' Cria template inteligente baseado no mapeamento
    ''' </summary>
    Private Sub CriarTemplateInteligente()
        StatusProcessamento = "CRIANDO_TEMPLATE"
        
        ' Configurar larguras das colunas
        xlWorksheet.Columns("A").ColumnWidth = 35  ' Descrição
        xlWorksheet.Columns("B").ColumnWidth = 8   ' Quantidade
        xlWorksheet.Columns("C").ColumnWidth = 6   ' Unidade
        xlWorksheet.Columns("D").ColumnWidth = 12  ' Preço unitário
        xlWorksheet.Columns("E").ColumnWidth = 12  ' Total
        xlWorksheet.Columns("F").ColumnWidth = 15  ' Extra
        
        ' Estrutura básica do talão
        CriarEstruturaTalao()
    End Sub

    ''' <summary>
    ''' Cria estrutura básica do talão
    ''' </summary>
    Private Sub CriarEstruturaTalao()
        ' Cabeçalho da tabela de produtos
        EscreverCelula("A15", "DESCRIÇÃO")
        EscreverCelula("B15", "QTD")
        EscreverCelula("C15", "UN")
        EscreverCelula("D15", "PREÇO UNIT.")
        EscreverCelula("E15", "TOTAL")
        
        ' Labels dos campos
        EscreverCelula("A10", "CLIENTE:")
        EscreverCelula("A11", "ENDEREÇO:")
        EscreverCelula("A12", "CIDADE/CEP:")
        EscreverCelula("A13", "TELEFONE:")
        EscreverCelula("A29", "FORMA DE PAGAMENTO:")
        EscreverCelula("A30", "VENDEDOR:")
        EscreverCelula("D27", "TOTAL GERAL:")
    End Sub

    ''' <summary>
    ''' Escreve dados da empresa usando mapeamento
    ''' </summary>
    Private Sub EscreverDadosEmpresa()
        StatusProcessamento = "ESCREVENDO_EMPRESA"
        
        EscreverCelulaMapeada("NOME_EMPRESA", ConfigurationManager.AppSettings("NomeMadeireira"))
        EscreverCelulaMapeada("ENDERECO_EMPRESA", ConfigurationManager.AppSettings("EnderecoMadeireira"))
        EscreverCelulaMapeada("CIDADE_EMPRESA", ConfigurationManager.AppSettings("CidadeMadeireira") & 
                             " - CEP: " & ConfigurationManager.AppSettings("CEPMadeireira"))
        EscreverCelulaMapeada("TELEFONE_EMPRESA", "Telefone: " & ConfigurationManager.AppSettings("TelefoneMadeireira"))
        EscreverCelulaMapeada("CNPJ_EMPRESA", "CNPJ: " & ConfigurationManager.AppSettings("CNPJMadeireira"))
    End Sub

    ''' <summary>
    ''' Escreve dados do talão
    ''' </summary>
    Private Sub EscreverDadosTalao(dados As DadosTalao)
        StatusProcessamento = "ESCREVENDO_TALAO"
        
        EscreverCelula("A7", "TALÃO DE VENDA Nº:")
        EscreverCelulaMapeada("NUMERO_TALAO", dados.NumeroTalao)
        EscreverCelulaMapeada("DATA_TALAO", "Data: " & dados.DataVenda.ToString("dd/MM/yyyy HH:mm"))
    End Sub

    ''' <summary>
    ''' Escreve dados do cliente usando mapeamento
    ''' </summary>
    Private Sub EscreverDadosCliente(dados As DadosTalao)
        StatusProcessamento = "ESCREVENDO_CLIENTE"
        
        EscreverCelulaMapeada("NOME_CLIENTE", dados.NomeCliente)
        EscreverCelulaMapeada("ENDERECO_CLIENTE", dados.EnderecoCliente)
        EscreverCelulaMapeada("CIDADE_CEP_CLIENTE", dados.Cidade & " - CEP: " & dados.CEP)
        EscreverCelulaMapeada("TELEFONE_CLIENTE", dados.Telefone)
    End Sub

    ''' <summary>
    ''' Escreve produtos com formatação visual inteligente
    ''' </summary>
    Private Sub EscreverProdutos(dados As DadosTalao)
        StatusProcessamento = "ESCREVENDO_PRODUTOS"
        
        Dim linha As Integer = 16
        Dim totalGeral As Double = 0
        
        For Each produto In dados.Produtos
            ' Escrever dados do produto
            EscreverCelula($"A{linha}", produto.Descricao)
            EscreverCelula($"B{linha}", produto.Quantidade)
            EscreverCelula($"C{linha}", produto.Unidade)
            EscreverCelula($"D{linha}", produto.PrecoUnitario)
            EscreverCelula($"E{linha}", produto.PrecoTotal)
            
            ' Aplicar formatação de valores monetários
            FormatarCelula($"D{linha}", "MOEDA")
            FormatarCelula($"E{linha}", "MOEDA")
            
            ' Aplicar formatação visual especial se há multiplicador
            Dim produtoCompleto = TryCast(produto, ProdutoTalao)
            If produtoCompleto IsNot Nothing AndAlso produtoCompleto.PrecoVisual > 0 Then
                ' Adicionar comentário com preço visual
                AdicionarComentario($"D{linha}", $"Preço Visual: {produtoCompleto.PrecoVisual:N0}")
            End If
            
            totalGeral += produto.PrecoTotal
            linha += 1
        Next
        
        ' Escrever total geral
        EscreverCelulaMapeada("TOTAL_GERAL", totalGeral)
        FormatarCelula(mapaCelulas("TOTAL_GERAL"), "MOEDA")
    End Sub

    ''' <summary>
    ''' Escreve totais e informações de finalização
    ''' </summary>
    Private Sub EscreverTotaisEFinalizacao(dados As DadosTalao)
        StatusProcessamento = "ESCREVENDO_FINALIZACAO"
        
        EscreverCelulaMapeada("FORMA_PAGAMENTO", dados.FormaPagamento)
        EscreverCelulaMapeada("VENDEDOR", dados.Vendedor)
        EscreverCelulaMapeada("ASSINATURA_CLIENTE", "CLIENTE: _________________________________")
        
        ' Adicionar observação sobre assinatura
        EscreverCelula("A33", "           (NOME E ASSINATURA)")
        
        ' Criar segunda via resumida
        CriarSegundaViaResumida(dados)
    End Sub

    ''' <summary>
    ''' Cria segunda via resumida
    ''' </summary>
    Private Sub CriarSegundaViaResumida(dados As DadosTalao)
        Dim linhaSeparador As Integer = 36
        
        ' Separador
        EscreverCelula($"A{linhaSeparador}", "✂️ --- CORTE AQUI - SEGUNDA VIA --- ✂️")
        MesclarCelulas($"A{linhaSeparador}:E{linhaSeparador}")
        
        ' Dados resumidos da segunda via
        linhaSeparador += 2
        EscreverCelula($"A{linhaSeparador}", ConfigurationManager.AppSettings("NomeMadeireira"))
        MesclarCelulas($"A{linhaSeparador}:E{linhaSeparador}")
        
        linhaSeparador += 1
        EscreverCelula($"A{linhaSeparador}", $"TALÃO Nº: {dados.NumeroTalao} - {dados.DataVenda:dd/MM/yyyy HH:mm}")
        
        linhaSeparador += 1
        EscreverCelula($"A{linhaSeparador}", $"CLIENTE: {dados.NomeCliente}")
        
        linhaSeparador += 1
        Dim totalGeral = dados.Produtos.Sum(Function(p) p.PrecoTotal)
        EscreverCelula($"A{linhaSeparador}", $"TOTAL: {totalGeral:C}")
        
        linhaSeparador += 1
        EscreverCelula($"A{linhaSeparador}", $"PAGAMENTO: {dados.FormaPagamento}")
        
        linhaSeparador += 1
        EscreverCelula($"A{linhaSeparador}", $"VENDEDOR: {dados.Vendedor}")
    End Sub

    ''' <summary>
    ''' Aplica formatação automática à planilha
    ''' </summary>
    Private Sub AplicarFormatacaoAutomatica()
        StatusProcessamento = "APLICANDO_FORMATACAO"
        
        ' Formatação do cabeçalho da empresa
        FormatarCelula("A1", "TITULO_EMPRESA")
        MesclarCelulas("A1:F1")
        
        ' Formatação do título do talão
        FormatarCelula("A7", "TITULO_TALAO")
        
        ' Formatação dos labels
        FormatarCelulas("A10:A13", "LABEL")
        FormatarCelulas("A29:A30", "LABEL")
        FormatarCelula("D27", "LABEL")
        
        ' Formatação da tabela de produtos
        FormatarCelulas("A15:E15", "CABECALHO_TABELA")
        
        ' Bordas da tabela
        Dim ultimaLinhaProduto = 16 + dados.Produtos.Count - 1
        AdicionarBordas($"A15:E{ultimaLinhaProduto}")
        
        ' Formatação especial para total geral
        FormatarCelula("D27:E27", "TOTAL_GERAL")
    End Sub

    ''' <summary>
    ''' Configura planilha para visualização e impressão
    ''' </summary>
    Private Sub ConfigurarVisualizacao()
        StatusProcessamento = "CONFIGURANDO_VISUALIZACAO"
        
        With xlWorksheet.PageSetup
            .PrintArea = "A1:F" & (xlWorksheet.UsedRange.Rows.Count + 2).ToString()
            .FitToPagesWide = 1
            .FitToPagesTall = False
            .CenterHorizontally = True
            .PrintTitleRows = "1:15"
        End With
        
        ' Habilitar quebras de página automáticas
        xlWorksheet.DisplayPageBreaks = True
    End Sub

    ''' <summary>
    ''' Escreve valor em célula específica
    ''' </summary>
    Private Sub EscreverCelula(endereco As String, valor As Object)
        Try
            xlWorksheet.Range(endereco).Value = valor
        Catch ex As Exception
            Console.WriteLine($"Erro ao escrever célula {endereco}: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Escreve valor usando mapeamento de célula
    ''' </summary>
    Private Sub EscreverCelulaMapeada(chave As String, valor As Object)
        If mapaCelulas.ContainsKey(chave) Then
            EscreverCelula(mapaCelulas(chave), valor)
        End If
    End Sub

    ''' <summary>
    ''' Aplica formatação específica a uma célula
    ''' </summary>
    Private Sub FormatarCelula(endereco As String, tipoFormatacao As String)
        Try
            Dim range = xlWorksheet.Range(endereco)
            
            Select Case tipoFormatacao.ToUpper()
                Case "TITULO_EMPRESA"
                    range.Font.Size = 18
                    range.Font.Bold = True
                    range.HorizontalAlignment = XlHAlign.xlHAlignCenter
                    
                Case "TITULO_TALAO"
                    range.Font.Size = 14
                    range.Font.Bold = True
                    
                Case "LABEL"
                    range.Font.Bold = True
                    
                Case "CABECALHO_TABELA"
                    range.Font.Bold = True
                    range.HorizontalAlignment = XlHAlign.xlHAlignCenter
                    range.Interior.Color = RGB(230, 230, 230)
                    range.Borders.LineStyle = XlLineStyle.xlContinuous
                    
                Case "MOEDA"
                    range.NumberFormat = "R$ #,##0.00"
                    
                Case "TOTAL_GERAL"
                    range.Font.Bold = True
                    range.Borders.LineStyle = XlLineStyle.xlContinuous
                    range.Borders(XlBordersIndex.xlEdgeTop).Weight = XlBorderWeight.xlThick
                    
            End Select
            
        Catch ex As Exception
            Console.WriteLine($"Erro ao formatar célula {endereco}: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Aplica formatação a um range de células
    ''' </summary>
    Private Sub FormatarCelulas(endereco As String, tipoFormatacao As String)
        FormatarCelula(endereco, tipoFormatacao)
    End Sub

    ''' <summary>
    ''' Mescla células
    ''' </summary>
    Private Sub MesclarCelulas(endereco As String)
        Try
            xlWorksheet.Range(endereco).Merge()
        Catch ex As Exception
            Console.WriteLine($"Erro ao mesclar células {endereco}: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Adiciona bordas a um range
    ''' </summary>
    Private Sub AdicionarBordas(endereco As String)
        Try
            xlWorksheet.Range(endereco).Borders.LineStyle = XlLineStyle.xlContinuous
        Catch ex As Exception
            Console.WriteLine($"Erro ao adicionar bordas {endereco}: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Adiciona comentário a uma célula
    ''' </summary>
    Private Sub AdicionarComentario(endereco As String, texto As String)
        Try
            Dim cell = xlWorksheet.Range(endereco)
            If cell.Comment Is Nothing Then
                cell.AddComment(texto)
            Else
                cell.Comment.Text(texto)
            End If
        Catch ex As Exception
            Console.WriteLine($"Erro ao adicionar comentário {endereco}: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Salva planilha mapeada se configurado
    ''' </summary>
    Public Sub SalvarSePreciso()
        Try
            If Boolean.Parse(ConfigurationManager.AppSettings("SalvarTalaoTemporario") OrElse "false") Then
                Dim nomeArquivo As String = "Talao_Mapeado_" & Date.Now.ToString("yyyyMMdd_HHmmss") & ".xlsx"
                Dim caminho As String = System.IO.Path.Combine(System.IO.Path.GetTempPath(), nomeArquivo)
                xlWorkbook.SaveAs(caminho)
                StatusProcessamento = $"SALVO: {caminho}"
            End If
        Catch ex As Exception
            Console.WriteLine($"Erro ao salvar: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Exibe planilha para visualização
    ''' </summary>
    Public Sub ExibirParaVisualizacao()
        Try
            xlApp.Visible = True
            xlApp.ScreenUpdating = True
            xlWorksheet.Activate()
            xlWorksheet.Range("A1").Select()
            StatusProcessamento = "EXIBINDO"
        Catch ex As Exception
            Console.WriteLine($"Erro ao exibir: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Fecha Excel e libera recursos
    ''' </summary>
    Public Sub FecharExcel()
        Try
            If xlWorkbook IsNot Nothing Then
                xlWorkbook.Close(SaveChanges:=False)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbook)
            End If

            If xlApp IsNot Nothing Then
                xlApp.Quit()
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp)
            End If

        Catch ex As Exception
            Console.WriteLine($"Erro ao fechar Excel: {ex.Message}")
        Finally
            xlWorksheet = Nothing
            xlWorkbook = Nothing
            xlApp = Nothing
            GC.Collect()
            GC.WaitForPendingFinalizers()
            StatusProcessamento = "FECHADO"
        End Try
    End Sub

    ''' <summary>
    ''' Obtém informações do mapeamento atual
    ''' </summary>
    Public Function ObterInfoMapeamento() As String
        Dim sb As New System.Text.StringBuilder()
        sb.AppendLine("=== MAPEAMENTO DE CÉLULAS ===")
        
        For Each kvp In mapaCelulas
            sb.AppendLine($"{kvp.Key}: {kvp.Value}")
        Next
        
        Return sb.ToString()
    End Function

    ''' <summary>
    ''' Permite adicionar novo mapeamento de célula
    ''' </summary>
    Public Sub AdicionarMapeamento(chave As String, endereco As String)
        mapaCelulas(chave) = endereco
    End Sub

    ''' <summary>
    ''' Remove mapeamento de célula
    ''' </summary>
    Public Sub RemoverMapeamento(chave As String)
        If mapaCelulas.ContainsKey(chave) Then
            mapaCelulas.Remove(chave)
        End If
    End Sub
End Class