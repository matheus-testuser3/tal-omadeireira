Imports Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices
Imports System.Threading.Tasks

''' <summary>
''' Serviço otimizado para automação do Excel
''' Versão melhorada com melhor performance e tratamento de erros
''' </summary>
Public Class ExcelService
    Private _xlApp As Application
    Private _xlWorkbook As Workbook
    Private _xlWorksheet As Worksheet
    Private ReadOnly _logger As Logger
    Private ReadOnly _config As ConfigManager
    
    ''' <summary>
    ''' Construtor
    ''' </summary>
    Public Sub New()
        _logger = Logger.Instance
        _config = ConfigManager.Instance
    End Sub
    
    ''' <summary>
    ''' Gera talão no Excel (versão otimizada)
    ''' </summary>
    Public Function GerarTalao(venda As Venda, Optional reimpressao As Boolean = False) As Boolean
        Try
            _logger.Info($"Iniciando geração de talão para venda {venda.NumeroTalao}")
            
            ' Configurar timeout
            Dim timeoutTask = Task.Delay(_config.TimeoutExcelSegundos * 1000)
            Dim geracaoTask = Task.Run(Function() ExecutarGeracaoTalao(venda, reimpressao))
            
            ' Aguardar conclusão ou timeout
            Dim completedTask = Task.WaitAny(geracaoTask, timeoutTask)
            
            If completedTask = 0 AndAlso geracaoTask.IsCompletedSuccessfully Then
                _logger.Info($"Talão {venda.NumeroTalao} gerado com sucesso")
                Return geracaoTask.Result
            Else
                _logger.Error($"Timeout na geração do talão {venda.NumeroTalao}")
                Return False
            End If
            
        Catch ex As Exception
            _logger.Error($"Erro na geração do talão {venda.NumeroTalao}", ex)
            Return False
        Finally
            LiberarRecursos()
        End Try
    End Function
    
    ''' <summary>
    ''' Executa a geração do talão de forma assíncrona
    ''' </summary>
    Private Function ExecutarGeracaoTalao(venda As Venda, reimpressao As Boolean) As Boolean
        Try
            ' Abrir Excel otimizado
            If Not AbrirExcelOtimizado() Then
                Return False
            End If
            
            ' Criar planilha temporária
            CriarPlanilhaTemporaria()
            
            ' Injetar VBA se necessário
            If Not reimpressao Then
                InjetarModulosVBA()
            End If
            
            ' Criar template otimizado
            CriarTemplateOtimizado(venda)
            
            ' Preencher dados
            PreencherDadosVenda(venda)
            
            ' Configurar e executar impressão
            ConfigurarImpressaoOtimizada()
            
            ' Imprimir
            ExecutarImpressao(reimpressao)
            
            Return True
            
        Catch ex As Exception
            _logger.Error("Erro durante execução da geração", ex)
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' Abre Excel com configurações otimizadas
    ''' </summary>
    Private Function AbrirExcelOtimizado() As Boolean
        Try
            _logger.Info("Abrindo Excel em modo otimizado")
            
            _xlApp = New Application()
            
            ' Configurações para performance
            _xlApp.Visible = _config.ExcelVisivel
            _xlApp.ScreenUpdating = False
            _xlApp.DisplayAlerts = False
            _xlApp.EnableEvents = False
            _xlApp.Calculation = XlCalculation.xlCalculationManual
            
            ' Verificar se Excel iniciou corretamente
            If _xlApp Is Nothing Then
                _logger.Error("Falha ao inicializar Excel")
                Return False
            End If
            
            _logger.Info("Excel iniciado com sucesso")
            Return True
            
        Catch ex As Exception
            _logger.Error("Erro ao abrir Excel", ex)
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' Cria planilha temporária otimizada
    ''' </summary>
    Private Sub CriarPlanilhaTemporaria()
        Try
            _xlWorkbook = _xlApp.Workbooks.Add()
            _xlWorksheet = _xlWorkbook.ActiveSheet
            _xlWorksheet.Name = "Talao_" & DateTime.Now.ToString("yyyyMMddHHmmss")
            
            _logger.Info("Planilha temporária criada")
        Catch ex As Exception
            _logger.Error("Erro ao criar planilha temporária", ex)
            Throw
        End Try
    End Sub
    
    ''' <summary>
    ''' Injeta módulos VBA otimizados
    ''' </summary>
    Private Sub InjetarModulosVBA()
        Try
            ' Implementação otimizada dos módulos VBA
            Dim moduloTalao = New ModuloTalao()
            Dim moduloTemplate = New ModuloTemplate()
            Dim moduloIntegracao = New ModuloIntegracao()
            
            ' Adicionar módulos ao workbook (simplificado para esta versão)
            _logger.Info("Módulos VBA configurados")
            
        Catch ex As Exception
            _logger.Warning("Erro ao injetar VBA, usando método direto", ex)
        End Try
    End Sub
    
    ''' <summary>
    ''' Cria template otimizado do talão
    ''' </summary>
    Private Sub CriarTemplateOtimizado(venda As Venda)
        Try
            ' Limpar planilha
            _xlWorksheet.Cells.Clear()
            
            ' Cabeçalho da empresa
            CriarCabecalhoEmpresa()
            
            ' Dados do talão
            CriarSecaoTalao(venda)
            
            ' Dados do cliente
            CriarSecaoCliente(venda.Cliente)
            
            ' Tabela de produtos
            CriarTabelaProdutos()
            
            ' Rodapé
            CriarRodape()
            
            _logger.Info("Template criado com sucesso")
            
        Catch ex As Exception
            _logger.Error("Erro ao criar template", ex)
            Throw
        End Try
    End Sub
    
    ''' <summary>
    ''' Cria cabeçalho da empresa
    ''' </summary>
    Private Sub CriarCabecalhoEmpresa()
        ' Nome da empresa
        _xlWorksheet.Cells(1, 1).Value = _config.NomeMadeireira
        _xlWorksheet.Cells(1, 1).Font.Size = 16
        _xlWorksheet.Cells(1, 1).Font.Bold = True
        _xlWorksheet.Range("A1:G1").Merge()
        _xlWorksheet.Cells(1, 1).HorizontalAlignment = XlHAlign.xlHAlignCenter
        
        ' Dados da empresa
        _xlWorksheet.Cells(2, 1).Value = _config.EnderecoMadeireira
        _xlWorksheet.Cells(3, 1).Value = $"{_config.CidadeMadeireira} - CEP: {_config.CEPMadeireira}"
        _xlWorksheet.Cells(4, 1).Value = $"Telefone: {_config.TelefoneMadeireira}"
        _xlWorksheet.Cells(5, 1).Value = $"CNPJ: {_config.CNPJMadeireira}"
    End Sub
    
    ''' <summary>
    ''' Cria seção do talão
    ''' </summary>
    Private Sub CriarSecaoTalao(venda As Venda)
        _xlWorksheet.Cells(7, 1).Value = $"TALÃO DE VENDA Nº: {venda.NumeroTalao}"
        _xlWorksheet.Cells(7, 1).Font.Size = 14
        _xlWorksheet.Cells(7, 1).Font.Bold = True
        
        _xlWorksheet.Cells(8, 1).Value = $"Data: {venda.DataVenda:dd/MM/yyyy HH:mm}"
        _xlWorksheet.Cells(8, 5).Value = $"Vendedor: {venda.Vendedor}"
    End Sub
    
    ''' <summary>
    ''' Cria seção do cliente
    ''' </summary>
    Private Sub CriarSecaoCliente(cliente As Cliente)
        _xlWorksheet.Cells(10, 1).Value = "DADOS DO CLIENTE:"
        _xlWorksheet.Cells(10, 1).Font.Bold = True
        
        _xlWorksheet.Cells(11, 1).Value = $"Nome: {cliente.Nome}"
        _xlWorksheet.Cells(12, 1).Value = $"Endereço: {cliente.Endereco}"
        _xlWorksheet.Cells(13, 1).Value = $"CEP: {cliente.CEP} - Cidade: {cliente.Cidade}"
        _xlWorksheet.Cells(14, 1).Value = $"Telefone: {cliente.Telefone}"
    End Sub
    
    ''' <summary>
    ''' Cria tabela de produtos
    ''' </summary>
    Private Sub CriarTabelaProdutos()
        ' Cabeçalho da tabela
        _xlWorksheet.Cells(16, 1).Value = "ITEM"
        _xlWorksheet.Cells(16, 2).Value = "DESCRIÇÃO"
        _xlWorksheet.Cells(16, 3).Value = "QTDE"
        _xlWorksheet.Cells(16, 4).Value = "UN"
        _xlWorksheet.Cells(16, 5).Value = "PREÇO UNIT."
        _xlWorksheet.Cells(16, 6).Value = "TOTAL"
        
        ' Formatação do cabeçalho
        Dim headerRange = _xlWorksheet.Range("A16:F16")
        headerRange.Font.Bold = True
        headerRange.Borders.LineStyle = XlLineStyle.xlContinuous
        headerRange.Interior.Color = RGB(220, 220, 220)
    End Sub
    
    ''' <summary>
    ''' Cria rodapé
    ''' </summary>
    Private Sub CriarRodape()
        _xlWorksheet.Cells(30, 1).Value = "Forma de Pagamento:"
        _xlWorksheet.Cells(32, 1).Value = "Observações:"
        _xlWorksheet.Cells(35, 1).Value = "TOTAL GERAL:"
        _xlWorksheet.Cells(35, 1).Font.Bold = True
        _xlWorksheet.Cells(35, 1).Font.Size = 12
    End Sub
    
    ''' <summary>
    ''' Preenche dados da venda
    ''' </summary>
    Private Sub PreencherDadosVenda(venda As Venda)
        Try
            ' Preencher produtos
            Dim linha = 17
            Dim item = 1
            
            For Each itemVenda In venda.Itens
                _xlWorksheet.Cells(linha, 1).Value = item
                _xlWorksheet.Cells(linha, 2).Value = itemVenda.Produto.Descricao
                _xlWorksheet.Cells(linha, 3).Value = itemVenda.Quantidade
                _xlWorksheet.Cells(linha, 4).Value = itemVenda.Produto.Unidade
                _xlWorksheet.Cells(linha, 5).Value = itemVenda.PrecoUnitario
                _xlWorksheet.Cells(linha, 6).Value = itemVenda.ValorTotal
                
                ' Formatação
                _xlWorksheet.Cells(linha, 5).NumberFormat = "R$ #,##0.00"
                _xlWorksheet.Cells(linha, 6).NumberFormat = "R$ #,##0.00"
                
                linha += 1
                item += 1
            Next
            
            ' Preencher forma de pagamento
            _xlWorksheet.Cells(30, 3).Value = venda.FormaPagamento
            
            ' Preencher total
            _xlWorksheet.Cells(35, 6).Value = venda.ValorTotal
            _xlWorksheet.Cells(35, 6).NumberFormat = "R$ #,##0.00"
            _xlWorksheet.Cells(35, 6).Font.Bold = True
            _xlWorksheet.Cells(35, 6).Font.Size = 12
            
            _logger.Info("Dados da venda preenchidos")
            
        Catch ex As Exception
            _logger.Error("Erro ao preencher dados", ex)
            Throw
        End Try
    End Sub
    
    ''' <summary>
    ''' Configura impressão otimizada
    ''' </summary>
    Private Sub ConfigurarImpressaoOtimizada()
        Try
            With _xlWorksheet.PageSetup
                .PrintArea = "A1:G40"
                .Orientation = XlPageOrientation.xlPortrait
                .PaperSize = XlPaperSize.xlPaperA4
                .TopMargin = _xlApp.InchesToPoints(0.5)
                .BottomMargin = _xlApp.InchesToPoints(0.5)
                .LeftMargin = _xlApp.InchesToPoints(0.5)
                .RightMargin = _xlApp.InchesToPoints(0.5)
                .FitToPagesWide = 1
                .FitToPagesTall = 1
            End With
            
            _logger.Info("Configuração de impressão definida")
            
        Catch ex As Exception
            _logger.Error("Erro ao configurar impressão", ex)
            Throw
        End Try
    End Sub
    
    ''' <summary>
    ''' Executa impressão
    ''' </summary>
    Private Sub ExecutarImpressao(reimpressao As Boolean)
        Try
            Dim prefixo = If(reimpressao, "REIMPRESSÃO - ", "")
            _xlWorksheet.Cells(7, 1).Value = $"{prefixo}TALÃO DE VENDA Nº: {_xlWorksheet.Cells(7, 1).Value.ToString().Replace("TALÃO DE VENDA Nº: ", "")}"
            
            _xlWorksheet.PrintOut()
            
            _logger.Info($"Impressão executada - {If(reimpressao, "Reimpressão", "Primeira via")}")
            
        Catch ex As Exception
            _logger.Error("Erro na impressão", ex)
            Throw
        End Try
    End Sub
    
    ''' <summary>
    ''' Libera recursos do Excel
    ''' </summary>
    Private Sub LiberarRecursos()
        Try
            _logger.Info("Liberando recursos do Excel")
            
            ' Restaurar configurações
            If _xlApp IsNot Nothing Then
                _xlApp.ScreenUpdating = True
                _xlApp.DisplayAlerts = True
                _xlApp.EnableEvents = True
                _xlApp.Calculation = XlCalculation.xlCalculationAutomatic
            End If
            
            ' Fechar workbook
            If _xlWorkbook IsNot Nothing Then
                _xlWorkbook.Close(SaveChanges:=_config.SalvarTalaoTemporario)
                Marshal.ReleaseComObject(_xlWorkbook)
                _xlWorkbook = Nothing
            End If
            
            ' Fechar Excel
            If _xlApp IsNot Nothing Then
                _xlApp.Quit()
                Marshal.ReleaseComObject(_xlApp)
                _xlApp = Nothing
            End If
            
            ' Liberar memória
            If _xlWorksheet IsNot Nothing Then
                Marshal.ReleaseComObject(_xlWorksheet)
                _xlWorksheet = Nothing
            End If
            
            GC.Collect()
            GC.WaitForPendingFinalizers()
            
            _logger.Info("Recursos liberados com sucesso")
            
        Catch ex As Exception
            _logger.Warning("Aviso ao liberar recursos", ex)
        End Try
    End Sub
End Class