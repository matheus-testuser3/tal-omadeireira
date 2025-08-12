Imports Microsoft.Office.Interop.Excel
Imports System.Configuration

''' <summary>
''' Classe responsável pela automação do Excel e integração com VBA
''' Gerencia abertura/fechamento automático do Excel e execução de macros
''' </summary>
Public Class ExcelAutomation
    Private xlApp As Application
    Private xlWorkbook As Workbook
    Private xlWorksheet As Worksheet
    Private vbaInjected As Boolean = False

    ''' <summary>
    ''' Processa o talão completo com sistema de mapeamento Excel
    ''' </summary>
    Public Sub ProcessarTalaoCompleto(dados As DadosTalao)
        Try
            ' Usar novo sistema de mapeamento em vez de impressão
            EscreverNaPlanilhaMapeada(dados)

        Finally
            ' Fechar Excel se necessário
            FecharExcel()
        End Try
    End Sub

    ''' <summary>
    ''' Escreve dados na planilha mapeada (substitui impressão)
    ''' </summary>
    Public Sub EscreverNaPlanilhaMapeada(dados As DadosTalao)
        Try
            ' Usar sistema de mapeamento de planilha
            Dim mapeamento As New MapeamentoPlanilha()
            mapeamento.EscreverNaPlanilhaMapeada(dados)
            
            ' Verificar se deve salvar ou exibir
            Dim exibirPlanilha = Boolean.Parse(ConfigurationManager.AppSettings("ExcelVisivel") OrElse "false")
            
            If exibirPlanilha Then
                mapeamento.ExibirParaVisualizacao()
            Else
                mapeamento.SalvarSePreciso()
                mapeamento.FecharExcel()
            End If

        Catch ex As Exception
            Throw New Exception("Erro no sistema de mapeamento: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Abre Excel em modo background
    ''' </summary>
    Private Sub AbrirExcel()
        xlApp = New Application()
        xlApp.Visible = Boolean.Parse(ConfigurationManager.AppSettings("ExcelVisivel") OrElse "false")
        xlApp.DisplayAlerts = False
        xlApp.ScreenUpdating = False

        ' Criar nova pasta de trabalho
        xlWorkbook = xlApp.Workbooks.Add()
        xlWorksheet = xlWorkbook.ActiveSheet
        xlWorksheet.Name = "Talão_" & Date.Now.ToString("HHmmss")
    End Sub

    ''' <summary>
    ''' Cria planilha temporária com configurações básicas
    ''' </summary>
    Private Sub CriarPlanilhaTemporaria()
        ' Configurar página
        With xlWorksheet.PageSetup
            .PaperSize = XlPaperSize.xlPaperA4
            .Orientation = XlPageOrientation.xlPortrait
            .LeftMargin = xlApp.InchesToPoints(0.5)
            .RightMargin = xlApp.InchesToPoints(0.5)
            .TopMargin = xlApp.InchesToPoints(0.5)
            .BottomMargin = xlApp.InchesToPoints(0.5)
            .HeaderMargin = xlApp.InchesToPoints(0.3)
            .FooterMargin = xlApp.InchesToPoints(0.3)
        End With
    End Sub

    ''' <summary>
    ''' Injeta módulos VBA na planilha
    ''' </summary>
    Private Sub InjetarModulosVBA()
        Try
            ' Obter código VBA dos módulos
            Dim moduloTalao As New ModuloTalao()
            Dim moduloTemplate As New ModuloTemplate()
            Dim moduloIntegracao As New ModuloIntegracao()

            ' Adicionar módulos VBA ao projeto
            Dim vbaModule1 = xlWorkbook.VBProject.VBComponents.Add(ComponentType:=1) ' Módulo padrão
            vbaModule1.Name = "ModuloTalao"
            vbaModule1.CodeModule.AddFromString(moduloTalao.ObterCodigoVBA())

            Dim vbaModule2 = xlWorkbook.VBProject.VBComponents.Add(ComponentType:=1)
            vbaModule2.Name = "ModuloTemplate"
            vbaModule2.CodeModule.AddFromString(moduloTemplate.ObterCodigoVBA())

            Dim vbaModule3 = xlWorkbook.VBProject.VBComponents.Add(ComponentType:=1)
            vbaModule3.Name = "ModuloIntegracao"
            vbaModule3.CodeModule.AddFromString(moduloIntegracao.ObterCodigoVBA())

            vbaInjected = True

        Catch ex As Exception
            ' Se não conseguir injetar VBA, usar método direto
            Console.WriteLine("Aviso: VBA não injetado, usando método direto. " & ex.Message)
            vbaInjected = False
        End Try
    End Sub

    ''' <summary>
    ''' Cria template profissional do talão
    ''' </summary>
    Private Sub CriarTemplate()
        If vbaInjected Then
            ' Executar via VBA se disponível
            Try
                xlApp.Run("ModuloTemplate.CriarTemplateAutomatico")
                Return
            Catch ex As Exception
                Console.WriteLine("Erro VBA, usando método direto: " & ex.Message)
            End Try
        End If

        ' Método direto caso VBA não funcione
        CriarTemplateDireto()
    End Sub

    ''' <summary>
    ''' Cria template diretamente via .NET (fallback)
    ''' </summary>
    Private Sub CriarTemplateDireto()
        ' Título da madeireira
        xlWorksheet.Cells(1, 1).Value = ConfigurationManager.AppSettings("NomeMadeireira")
        xlWorksheet.Cells(1, 1).Font.Size = 16
        xlWorksheet.Cells(1, 1).Font.Bold = True
        xlWorksheet.Range("A1:G1").Merge()
        xlWorksheet.Cells(1, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter

        ' Dados da empresa
        xlWorksheet.Cells(2, 1).Value = ConfigurationManager.AppSettings("EnderecoMadeireira")
        xlWorksheet.Cells(3, 1).Value = ConfigurationManager.AppSettings("CidadeMadeireira") & " - CEP: " & ConfigurationManager.AppSettings("CEPMadeireira")
        xlWorksheet.Cells(4, 1).Value = "Telefone: " & ConfigurationManager.AppSettings("TelefoneMadeireira")
        xlWorksheet.Cells(5, 1).Value = "CNPJ: " & ConfigurationManager.AppSettings("CNPJMadeireira")

        ' Título do talão
        xlWorksheet.Cells(7, 1).Value = "TALÃO DE VENDA Nº:"
        xlWorksheet.Cells(7, 1).Font.Size = 14
        xlWorksheet.Cells(7, 1).Font.Bold = True

        ' Data
        xlWorksheet.Cells(8, 1).Value = "Data: " & Date.Now.ToString("dd/MM/yyyy HH:mm")

        ' Dados do cliente (cabeçalho)
        xlWorksheet.Cells(10, 1).Value = "CLIENTE:"
        xlWorksheet.Cells(10, 1).Font.Bold = True
        xlWorksheet.Cells(11, 1).Value = "ENDEREÇO:"
        xlWorksheet.Cells(11, 1).Font.Bold = True
        xlWorksheet.Cells(12, 1).Value = "CIDADE/CEP:"
        xlWorksheet.Cells(12, 1).Font.Bold = True
        xlWorksheet.Cells(13, 1).Value = "TELEFONE:"
        xlWorksheet.Cells(13, 1).Font.Bold = True

        ' Cabeçalho dos produtos
        xlWorksheet.Cells(15, 1).Value = "DESCRIÇÃO"
        xlWorksheet.Cells(15, 2).Value = "QTD"
        xlWorksheet.Cells(15, 3).Value = "UN"
        xlWorksheet.Cells(15, 4).Value = "PREÇO UNIT."
        xlWorksheet.Cells(15, 5).Value = "TOTAL"

        ' Formatar cabeçalho dos produtos
        Dim headerRange = xlWorksheet.Range("A15:E15")
        headerRange.Font.Bold = True
        headerRange.Borders.LineStyle = XlLineStyle.xlContinuous
        headerRange.Interior.Color = RGB(230, 230, 230)
        headerRange.HorizontalAlignment = XlHAlign.xlHAlignCenter

        ' Ajustar larguras das colunas
        xlWorksheet.Columns("A").ColumnWidth = 35 ' Descrição
        xlWorksheet.Columns("B").ColumnWidth = 8  ' Quantidade
        xlWorksheet.Columns("C").ColumnWidth = 6  ' Unidade
        xlWorksheet.Columns("D").ColumnWidth = 12 ' Preço unitário
        xlWorksheet.Columns("E").ColumnWidth = 12 ' Total
    End Sub

    ''' <summary>
    ''' Preenche dados do cliente e produtos
    ''' </summary>
    Private Sub PreencherDados(dados As DadosTalao)
        ' Número do talão
        xlWorksheet.Cells(7, 5).Value = dados.NumeroTalao

        ' Dados do cliente
        xlWorksheet.Cells(10, 2).Value = dados.NomeCliente
        xlWorksheet.Cells(11, 2).Value = dados.EnderecoCliente
        xlWorksheet.Cells(12, 2).Value = dados.Cidade & " - CEP: " & dados.CEP
        xlWorksheet.Cells(13, 2).Value = dados.Telefone

        ' Produtos
        Dim linha As Integer = 16
        Dim totalGeral As Double = 0

        For Each produto In dados.Produtos
            xlWorksheet.Cells(linha, 1).Value = produto.Descricao
            xlWorksheet.Cells(linha, 2).Value = produto.Quantidade
            xlWorksheet.Cells(linha, 3).Value = produto.Unidade
            xlWorksheet.Cells(linha, 4).Value = produto.PrecoUnitario
            xlWorksheet.Cells(linha, 5).Value = produto.PrecoTotal

            ' Formatar valores monetários
            xlWorksheet.Cells(linha, 4).NumberFormat = "R$ #,##0.00"
            xlWorksheet.Cells(linha, 5).NumberFormat = "R$ #,##0.00"

            ' Adicionar bordas
            xlWorksheet.Range("A" & linha & ":E" & linha).Borders.LineStyle = XlLineStyle.xlContinuous

            totalGeral += produto.PrecoTotal
            linha += 1
        Next

        ' Total geral
        linha += 1
        xlWorksheet.Cells(linha, 4).Value = "TOTAL GERAL:"
        xlWorksheet.Cells(linha, 4).Font.Bold = True
        xlWorksheet.Cells(linha, 5).Value = totalGeral
        xlWorksheet.Cells(linha, 5).NumberFormat = "R$ #,##0.00"
        xlWorksheet.Cells(linha, 5).Font.Bold = True

        ' Forma de pagamento
        linha += 2
        xlWorksheet.Cells(linha, 1).Value = "FORMA DE PAGAMENTO: " & dados.FormaPagamento
        xlWorksheet.Cells(linha, 1).Font.Bold = True

        ' Vendedor
        linha += 1
        xlWorksheet.Cells(linha, 1).Value = "VENDEDOR: " & dados.Vendedor

        ' Linha para assinatura do cliente
        linha += 3
        xlWorksheet.Cells(linha, 1).Value = "CLIENTE: ________________________________"
        xlWorksheet.Cells(linha, 1).Font.Bold = True

        ' Segundo talão (cópia)
        CriarSegundoTalao(dados, linha + 3)
    End Sub

    ''' <summary>
    ''' Cria segunda via do talão na mesma página
    ''' </summary>
    Private Sub CriarSegundoTalao(dados As DadosTalao, linhaInicial As Integer)
        ' Adicionar linha separadora
        xlWorksheet.Cells(linhaInicial, 1).Value = "--- SEGUNDA VIA ---"
        xlWorksheet.Cells(linhaInicial, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
        xlWorksheet.Cells(linhaInicial, 1).Font.Bold = True
        xlWorksheet.Range("A" & linhaInicial & ":E" & linhaInicial).Merge()

        linhaInicial += 2

        ' Repetir estrutura do primeiro talão de forma resumida
        xlWorksheet.Cells(linhaInicial, 1).Value = ConfigurationManager.AppSettings("NomeMadeireira")
        xlWorksheet.Cells(linhaInicial, 1).Font.Size = 14
        xlWorksheet.Cells(linhaInicial, 1).Font.Bold = True

        xlWorksheet.Cells(linhaInicial + 1, 1).Value = "TALÃO Nº: " & dados.NumeroTalao & " - " & Date.Now.ToString("dd/MM/yyyy")
        xlWorksheet.Cells(linhaInicial + 2, 1).Value = "CLIENTE: " & dados.NomeCliente

        ' Total resumido
        Dim totalGeral As Double = dados.Produtos.Sum(Function(p) p.PrecoTotal)
        xlWorksheet.Cells(linhaInicial + 3, 1).Value = "TOTAL: " & totalGeral.ToString("C")
        xlWorksheet.Cells(linhaInicial + 3, 1).Font.Bold = True
        xlWorksheet.Cells(linhaInicial + 4, 1).Value = "PAGAMENTO: " & dados.FormaPagamento
        xlWorksheet.Cells(linhaInicial + 5, 1).Value = "VENDEDOR: " & dados.Vendedor
    End Sub

    ''' <summary>
    ''' Configura as opções de impressão
    ''' </summary>
    Private Sub ConfigurarImpressao()
        With xlWorksheet.PageSetup
            .PrintArea = "A1:E" & (xlWorksheet.UsedRange.Rows.Count + 2).ToString()
            .FitToPagesWide = 1
            .FitToPagesTall = 1
            .CenterHorizontally = True
            .PrintTitleRows = "1:15" ' Repetir cabeçalho em todas as páginas
        End With
    End Sub

    ''' <summary>
    ''' Executa a impressão do talão
    ''' </summary>
    Private Sub ImprimirTalao()
        Try
            If vbaInjected Then
                ' Tentar via VBA primeiro
                xlApp.Run("ModuloTalao.ConfigurarImpressaoCompleta")
            End If

            ' Imprimir direto
            xlWorksheet.PrintOut(Copies:=1, Preview:=False)

            ' Salvar temporariamente se configurado
            If Boolean.Parse(ConfigurationManager.AppSettings("SalvarTalaoTemporario") OrElse "false") Then
                Dim nomeArquivo As String = "Talao_" & Date.Now.ToString("yyyyMMdd_HHmmss") & ".xlsx"
                xlWorkbook.SaveAs(System.IO.Path.GetTempPath() & nomeArquivo)
            End If

        Catch ex As Exception
            Throw New Exception("Erro ao imprimir talão: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Fecha Excel e libera recursos
    ''' </summary>
    Private Sub FecharExcel()
        Try
            If xlWorkbook IsNot Nothing Then
                xlWorkbook.Close(SaveChanges:=False)
                ReleaseComObject(xlWorkbook)
            End If

            If xlApp IsNot Nothing Then
                xlApp.Quit()
                ReleaseComObject(xlApp)
            End If

        Catch ex As Exception
            ' Ignorar erros ao fechar
            Console.WriteLine("Aviso ao fechar Excel: " & ex.Message)
        Finally
            xlWorksheet = Nothing
            xlWorkbook = Nothing
            xlApp = Nothing
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Sub

    ''' <summary>
    ''' Libera objetos COM adequadamente
    ''' </summary>
    Private Sub ReleaseComObject(obj As Object)
        Try
            If obj IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            End If
        Catch
            ' Ignorar erros de liberação
        End Try
    End Sub

    ''' <summary>
    ''' Função para testar automação Excel com sistema de mapeamento
    ''' </summary>
    Public Sub TestarAutomacao()
        Dim dadosTeste As New DadosTalao()
        dadosTeste.NomeCliente = "Cliente Teste - Mapeamento"
        dadosTeste.EnderecoCliente = "Rua Teste, 123"
        dadosTeste.CEP = "12345-678"
        dadosTeste.Cidade = "Cidade Teste/UF"
        dadosTeste.Telefone = "(11) 1234-5678"
        dadosTeste.FormaPagamento = "Dinheiro"
        dadosTeste.Vendedor = "Vendedor Teste"

        Dim produto As New ProdutoTalao()
        produto.Codigo = "TEST001"
        produto.Descricao = "Produto Teste - Mapeamento"
        produto.Quantidade = 1
        produto.Unidade = "UN"
        produto.PrecoUnitario = 10.0
        produto.PrecoTotal = 10.0
        produto.PrecoVisual = 10000
        dadosTeste.Produtos.Add(produto)

        ' Usar novo sistema de mapeamento
        EscreverNaPlanilhaMapeada(dadosTeste)
    End Sub
End Class