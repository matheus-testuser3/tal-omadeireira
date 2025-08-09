' ====================================================================
' MODULOTALAO_VBA.VB - Módulo de Integração Excel/VBA
' Sistema completo para geração de talão automático
' Este arquivo contém o código-fonte completo do módulo VBA
' ====================================================================

Imports System
Imports Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices
Imports System.IO
Imports System.Windows.Forms

Public Class ModuloTalaoVBA
    Private excelApp As Application
    Private workbook As Workbook
    Private worksheet As Worksheet

    Public Sub ProcessarTalaoCompleto(dados As DadosCliente)
        Try
            ' Inicializar Excel
            InicializarExcel()

            ' Criar template automático
            CriarTemplateAutomatico()

            ' Preencher dados
            PreencherDados(dados)

            ' Gerar talão duplo
            GerarTalaoDuplo()

            ' Configurar e imprimir
            ConfigurarImpressao()
            ImprimirTalao()

        Finally
            ' Limpar recursos
            LimparRecursos()
        End Try
    End Sub

    Private Sub InicializarExcel()
        Try
            excelApp = New Application()
            excelApp.Visible = False
            excelApp.DisplayAlerts = False
            workbook = excelApp.Workbooks.Add()
            worksheet = CType(workbook.ActiveSheet, Worksheet)
        Catch ex As Exception
            Throw New Exception($"Erro ao inicializar Excel: {ex.Message}")
        End Try
    End Sub

    Private Sub CriarTemplateAutomatico()
        Try
            ' Configurar página
            With worksheet.PageSetup
                .Orientation = XlPageOrientation.xlLandscape
                .PaperSize = XlPaperSize.xlPaperA4
                .LeftMargin = 20
                .RightMargin = 20
                .TopMargin = 20
                .BottomMargin = 20
            End With

            ' Criar estrutura do talão lado esquerdo (Cliente)
            CriarEstruturaTalao(1) ' Coluna A (Cliente)
            
            ' Criar estrutura do talão lado direito (Vendedor)
            CriarEstruturaTalao(10) ' Coluna J (Vendedor)

            ' Adicionar linha divisória
            worksheet.Range("I:I").ColumnWidth = 2
            worksheet.Range("I1:I30").Interior.Color = RGB(200, 200, 200)

        Catch ex As Exception
            Throw New Exception($"Erro ao criar template: {ex.Message}")
        End Try
    End Sub

    Private Sub CriarEstruturaTalao(colunaInicial As Integer)
        Dim coluna As String = Chr(64 + colunaInicial) ' Converter número para letra

        ' Cabeçalho da empresa
        worksheet.Range($"{coluna}1:{Chr(64 + colunaInicial + 7)}3").Merge()
        worksheet.Range($"{coluna}1").Value = "MADEIREIRA MARIA LUIZA"
        worksheet.Range($"{coluna}1").Font.Size = 16
        worksheet.Range($"{coluna}1").Font.Bold = True
        worksheet.Range($"{coluna}1").HorizontalAlignment = XlHAlign.xlHAlignCenter

        ' Endereço
        worksheet.Range($"{coluna}4:{Chr(64 + colunaInicial + 7)}4").Merge()
        worksheet.Range($"{coluna}4").Value = "Av. Dr. Olíncio Guerreiro Leite - 631-Paq Amadeu-Paulista-PE-55431-165"
        worksheet.Range($"{coluna}4").Font.Size = 10
        worksheet.Range($"{coluna}4").HorizontalAlignment = XlHAlign.xlHAlignCenter

        ' Telefone
        worksheet.Range($"{coluna}5:{Chr(64 + colunaInicial + 7)}5").Merge()
        worksheet.Range($"{coluna}5").Value = "Telefone: (81) 98570-1522"
        worksheet.Range($"{coluna}5").Font.Size = 10
        worksheet.Range($"{coluna}5").HorizontalAlignment = XlHAlign.xlHAlignCenter

        ' CNPJ
        worksheet.Range($"{coluna}6:{Chr(64 + colunaInicial + 7)}6").Merge()
        worksheet.Range($"{coluna}6").Value = "CNPJ: 48.905.025/001-61"
        worksheet.Range($"{coluna}6").Font.Size = 10
        worksheet.Range($"{coluna}6").HorizontalAlignment = XlHAlign.xlHAlignCenter

        ' Linha separadora
        worksheet.Range($"{coluna}7:{Chr(64 + colunaInicial + 7)}7").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        ' Campos do cliente
        worksheet.Range($"{coluna}9").Value = "Cliente:"
        worksheet.Range($"{coluna}9").Font.Bold = True
        worksheet.Range($"{Chr(64 + colunaInicial + 1)}9:{Chr(64 + colunaInicial + 7)}9").Merge()

        worksheet.Range($"{coluna}11").Value = "Endereço:"
        worksheet.Range($"{coluna}11").Font.Bold = True
        worksheet.Range($"{Chr(64 + colunaInicial + 1)}11:{Chr(64 + colunaInicial + 7)}11").Merge()

        worksheet.Range($"{coluna}13").Value = "Cidade:"
        worksheet.Range($"{coluna}13").Font.Bold = True
        worksheet.Range($"{Chr(64 + colunaInicial + 1)}13:{Chr(64 + colunaInicial + 3)}13").Merge()

        worksheet.Range($"{Chr(64 + colunaInicial + 4)}13").Value = "CEP:"
        worksheet.Range($"{Chr(64 + colunaInicial + 4)}13").Font.Bold = True
        worksheet.Range($"{Chr(64 + colunaInicial + 5)}13:{Chr(64 + colunaInicial + 7)}13").Merge()

        ' Produtos/Serviços
        worksheet.Range($"{coluna}15").Value = "Produtos/Serviços:"
        worksheet.Range($"{coluna}15").Font.Bold = True
        worksheet.Range($"{Chr(64 + colunaInicial + 1)}15:{Chr(64 + colunaInicial + 7)}19").Merge()

        ' Valor Total
        worksheet.Range($"{coluna}21").Value = "Valor Total:"
        worksheet.Range($"{coluna}21").Font.Bold = True
        worksheet.Range($"{Chr(64 + colunaInicial + 1)}21:{Chr(64 + colunaInicial + 3)}21").Merge()

        ' Forma de Pagamento
        worksheet.Range($"{Chr(64 + colunaInicial + 4)}21").Value = "Forma Pgto:"
        worksheet.Range($"{Chr(64 + colunaInicial + 4)}21").Font.Bold = True
        worksheet.Range($"{Chr(64 + colunaInicial + 5)}21:{Chr(64 + colunaInicial + 7)}21").Merge()

        ' Vendedor
        worksheet.Range($"{coluna}23").Value = "Vendedor:"
        worksheet.Range($"{coluna}23").Font.Bold = True
        worksheet.Range($"{Chr(64 + colunaInicial + 1)}23:{Chr(64 + colunaInicial + 7)}23").Merge()

        ' Data
        worksheet.Range($"{coluna}25").Value = "Data:"
        worksheet.Range($"{coluna}25").Font.Bold = True
        worksheet.Range($"{Chr(64 + colunaInicial + 1)}25:{Chr(64 + colunaInicial + 3)}25").Merge()
        worksheet.Range($"{Chr(64 + colunaInicial + 1)}25").Value = DateTime.Now.ToString("dd/MM/yyyy")

        ' Assinatura
        worksheet.Range($"{coluna}27").Value = "Assinatura:"
        worksheet.Range($"{coluna}27").Font.Bold = True
        worksheet.Range($"{Chr(64 + colunaInicial + 1)}27:{Chr(64 + colunaInicial + 7)}27").Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous

        ' Rodapé
        worksheet.Range($"{coluna}29:{Chr(64 + colunaInicial + 7)}29").Merge()
        worksheet.Range($"{coluna}29").Value = "WhatsApp: (81) 98570-1522 | Instagram: @madeireiramaria"
        worksheet.Range($"{coluna}29").Font.Size = 8
        worksheet.Range($"{coluna}29").HorizontalAlignment = XlHAlign.xlHAlignCenter

        ' Aplicar bordas
        worksheet.Range($"{coluna}1:{Chr(64 + colunaInicial + 7)}29").Borders.LineStyle = XlLineStyle.xlContinuous
    End Sub

    Private Sub PreencherDados(dados As DadosCliente)
        Try
            ' Preencher lado esquerdo (Cliente)
            PreencherLadoTalao(1, dados)

            ' Preencher lado direito (Vendedor)
            PreencherLadoTalao(10, dados)

        Catch ex As Exception
            Throw New Exception($"Erro ao preencher dados: {ex.Message}")
        End Try
    End Sub

    Private Sub PreencherLadoTalao(colunaInicial As Integer, dados As DadosCliente)
        Dim colunaValor As String = Chr(64 + colunaInicial + 1)

        ' Cliente
        worksheet.Range($"{colunaValor}9").Value = dados.Nome

        ' Endereço
        worksheet.Range($"{colunaValor}11").Value = dados.Endereco

        ' Cidade
        worksheet.Range($"{colunaValor}13").Value = dados.Cidade

        ' CEP
        worksheet.Range($"{Chr(64 + colunaInicial + 5)}13").Value = dados.CEP

        ' Produtos
        worksheet.Range($"{colunaValor}15").Value = dados.Produtos

        ' Valor Total
        worksheet.Range($"{colunaValor}21").Value = $"R$ {dados.ValorTotal}"

        ' Forma de Pagamento
        worksheet.Range($"{Chr(64 + colunaInicial + 5)}21").Value = dados.FormaPagamento

        ' Vendedor
        worksheet.Range($"{colunaValor}23").Value = dados.Vendedor
    End Sub

    Private Sub GerarTalaoDuplo()
        ' Ajustar largura das colunas
        worksheet.Columns("A:H").ColumnWidth = 12
        worksheet.Columns("J:Q").ColumnWidth = 12

        ' Ajustar altura das linhas
        worksheet.Rows("15:19").RowHeight = 20
    End Sub

    Private Sub ConfigurarImpressao()
        Try
            With worksheet.PageSetup
                .PrintArea = "A1:Q29"
                .FitToPagesWide = 1
                .FitToPagesTall = 1
                .CenterHorizontally = True
                .CenterVertically = True
            End With
        Catch ex As Exception
            Throw New Exception($"Erro ao configurar impressão: {ex.Message}")
        End Try
    End Sub

    Private Sub ImprimirTalao()
        Try
            ' Verificar se existe impressora padrão
            If excelApp.ActivePrinter IsNot Nothing Then
                worksheet.PrintOut()
            Else
                ' Mostrar diálogo de impressão
                worksheet.PrintPreview()
            End If
        Catch ex As Exception
            Throw New Exception($"Erro ao imprimir: {ex.Message}")
        End Try
    End Sub

    Private Sub LimparRecursos()
        Try
            If worksheet IsNot Nothing Then
                Marshal.ReleaseComObject(worksheet)
                worksheet = Nothing
            End If

            If workbook IsNot Nothing Then
                workbook.Close(False)
                Marshal.ReleaseComObject(workbook)
                workbook = Nothing
            End If

            If excelApp IsNot Nothing Then
                excelApp.Quit()
                Marshal.ReleaseComObject(excelApp)
                excelApp = Nothing
            End If

            ' Forçar coleta de lixo
            GC.Collect()
            GC.WaitForPendingFinalizers()

        Catch ex As Exception
            ' Log erro mas não propagar
            Console.WriteLine($"Erro ao limpar recursos: {ex.Message}")
        End Try
    End Sub

    ' Função auxiliar para RGB
    Private Function RGB(r As Integer, g As Integer, b As Integer) As Integer
        Return r + (g * 256) + (b * 256 * 256)
    End Function
End Class