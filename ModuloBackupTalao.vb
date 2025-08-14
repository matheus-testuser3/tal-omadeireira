''' <summary>
''' Módulo principal para backup e restauração de talões - Madeireira Maria Luiza
''' Data/Hora: 2025-08-14 11:16:26 UTC
''' Usuário: matheus-testuser3
''' Sistema de Backup e Restauração de Talões
''' </summary>

Imports Microsoft.Office.Interop.Excel
Imports Newtonsoft.Json
Imports System.IO
Imports System.Configuration
Imports System.Globalization

''' <summary>
''' Classe principal para importação de planilhas de backup e geração de talões
''' Detecta formato automático e processa dados específicos de madeireira
''' </summary>
Public Class ModuloBackupTalao
    
    ' === CONFIGURAÇÕES ===
    Private ReadOnly config As ConfiguracaoBackupMadeireira
    Private ReadOnly logDebug As Boolean = True
    
    ' === ESTADO DA IMPORTAÇÃO ===
    Private ultimaImportacao As List(Of DadosTalaoMadeireira)
    Private ultimoArquivoImportado As String
    
    ''' <summary>
    ''' Construtor que carrega configurações do App.config
    ''' </summary>
    Public Sub New()
        config = New ConfiguracaoBackupMadeireira()
        
        ' Carregar configurações do App.config se disponíveis
        Try
            Dim caminhoBackups = ConfigurationManager.AppSettings("CaminhoBackupsImportados")
            If Not String.IsNullOrEmpty(caminhoBackups) Then config.CaminhoBackupsImportados = caminhoBackups
            
            Dim caminhoTaloes = ConfigurationManager.AppSettings("CaminhoTaloesGerados")
            If Not String.IsNullOrEmpty(caminhoTaloes) Then config.CaminhoTaloesGerados = caminhoTaloes
            
            Dim caminhoJSON = ConfigurationManager.AppSettings("CaminhoBackupJSON")
            If Not String.IsNullOrEmpty(caminhoJSON) Then config.CaminhoBackupJSON = caminhoJSON
            
            Dim formatoData = ConfigurationManager.AppSettings("FormatoDataBackup")
            If Not String.IsNullOrEmpty(formatoData) Then config.FormatoDataBackup = formatoData
            
            Dim prefixo = ConfigurationManager.AppSettings("PrefixoArquivoBackup")
            If Not String.IsNullOrEmpty(prefixo) Then config.PrefixoArquivoBackup = prefixo
            
            Boolean.TryParse(ConfigurationManager.AppSettings("ManterHistoricoBackups"), config.ManterHistoricoBackups)
            Integer.TryParse(ConfigurationManager.AppSettings("DiasRetencaoBackups"), config.DiasRetencaoBackups)
            
        Catch ex As Exception
            LogDebug($"Erro ao carregar configurações: {ex.Message} - usando padrões")
        End Try
    End Sub
    
    ''' <summary>
    ''' Importa planilhas Excel de backup e detecta formato automaticamente
    ''' </summary>
    Public Function ImportarBackupExcel(caminhoArquivo As String) As List(Of DadosTalaoMadeireira)
        Try
            LogDebug($"=== INÍCIO IMPORTAÇÃO BACKUP ===")
            LogDebug($"Arquivo: {caminhoArquivo}")
            LogDebug($"Data/Hora: {DateTime.UtcNow:yyyy-MM-dd HH:mm:ss} UTC")
            LogDebug($"Usuário: matheus-testuser3")
            
            ' Verificar se arquivo existe
            If Not File.Exists(caminhoArquivo) Then
                Throw New FileNotFoundException($"Arquivo não encontrado: {caminhoArquivo}")
            End If
            
            ' Criar diretórios se não existirem
            CriarDiretoriosBackup()
            
            Dim taloes As New List(Of DadosTalaoMadeireira)()
            Dim xlApp As Application = Nothing
            Dim xlWorkbook As Workbook = Nothing
            
            Try
                ' Abrir Excel em background
                xlApp = New Application()
                xlApp.Visible = False
                xlApp.DisplayAlerts = False
                xlApp.ScreenUpdating = False
                
                LogDebug("Excel aberto em background para importação")
                
                ' Abrir planilha de backup
                xlWorkbook = xlApp.Workbooks.Open(caminhoArquivo)
                LogDebug($"Planilha aberta: {xlWorkbook.Name}")
                
                ' Detectar formato da planilha
                Dim formatoDetectado = DetectarFormatoBackup(xlWorkbook)
                LogDebug($"Formato detectado: {formatoDetectado}")
                
                ' Processar de acordo com o formato
                Select Case formatoDetectado
                    Case TipoFormatoBackup.Madeireira
                        taloes = ProcessarFormatoMadeireira(xlWorkbook)
                    Case TipoFormatoBackup.Generico
                        taloes = ProcessarFormatoGenerico(xlWorkbook)
                    Case Else
                        Throw New InvalidOperationException("Formato de planilha não reconhecido")
                End Select
                
                ' Salvar backup local em JSON
                SalvarBackupJSON(taloes, caminhoArquivo)
                
                LogDebug($"Importação concluída: {taloes.Count} talões processados")
                
                ' Armazenar para uso posterior
                ultimaImportacao = taloes
                ultimoArquivoImportado = caminhoArquivo
                
                Return taloes
                
            Finally
                ' Cleanup Excel
                If xlWorkbook IsNot Nothing Then
                    xlWorkbook.Close(False)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbook)
                End If
                
                If xlApp IsNot Nothing Then
                    xlApp.Quit()
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp)
                End If
                
                LogDebug("Recursos Excel liberados")
            End Try
            
        Catch ex As Exception
            LogDebug($"ERRO na importação: {ex.Message}")
            Throw New InvalidOperationException($"Erro ao importar backup: {ex.Message}", ex)
        End Try
    End Function
    
    ''' <summary>
    ''' Detecta automaticamente o formato da planilha de backup
    ''' </summary>
    Private Function DetectarFormatoBackup(xlWorkbook As Workbook) As TipoFormatoBackup
        Try
            LogDebug("=== DETECÇÃO DE FORMATO ===")
            
            Dim worksheet As Worksheet = xlWorkbook.Sheets(1)
            
            ' Analisar cabeçalhos da primeira linha
            Dim cabecalhos As New List(Of String)()
            For col As Integer = 1 To 20 ' Verificar até 20 colunas
                Dim valor = worksheet.Cells(1, col).Value
                If valor IsNot Nothing Then
                    cabecalhos.Add(valor.ToString().ToUpper().Trim())
                Else
                    Exit For
                End If
            Next
            
            LogDebug($"Cabeçalhos encontrados: {String.Join(", ", cabecalhos)}")
            
            ' Palavras-chave específicas da madeireira
            Dim chavesEspecificasMadeireira = {
                "TIPO_MADEIRA", "CATEGORIA", "DIMENSOES", "COMPRIMENTO",
                "TRATAMENTO", "QUALIDADE", "M³", "M²", "BARROTE",
                "CABRO", "TABUA", "VIGA", "MASSARANDUBA", "IPÊ",
                "PEROBA", "PINUS", "AUTOCLAVADO"
            }
            
            ' Palavras-chave genéricas de talão
            Dim chavesGenericasTalao = {
                "CLIENTE", "PRODUTO", "QUANTIDADE", "PRECO", "TOTAL",
                "TALAO", "NUMERO", "DATA", "VENDEDOR"
            }
            
            Dim pontosMadeireira = 0
            Dim pontosGenerico = 0
            
            ' Contar correspondências
            For Each cabecalho In cabecalhos
                If chavesEspecificasMadeireira.Any(Function(k) cabecalho.Contains(k)) Then
                    pontosMadeireira += 1
                End If
                If chavesGenericasTalao.Any(Function(k) cabecalho.Contains(k)) Then
                    pontosGenerico += 1
                End If
            Next
            
            LogDebug($"Pontos Madeireira: {pontosMadeireira}, Pontos Genérico: {pontosGenerico}")
            
            ' Decidir formato baseado na pontuação
            If pontosMadeireira >= 2 Then
                Return TipoFormatoBackup.Madeireira
            ElseIf pontosGenerico >= 3 Then
                Return TipoFormatoBackup.Generico
            Else
                Return TipoFormatoBackup.Desconhecido
            End If
            
        Catch ex As Exception
            LogDebug($"Erro na detecção de formato: {ex.Message}")
            Return TipoFormatoBackup.Desconhecido
        End Try
    End Function
    
    ''' <summary>
    ''' Processa planilha no formato específico da madeireira
    ''' </summary>
    Private Function ProcessarFormatoMadeireira(xlWorkbook As Workbook) As List(Of DadosTalaoMadeireira)
        LogDebug("=== PROCESSAMENTO FORMATO MADEIREIRA ===")
        
        Dim taloes As New List(Of DadosTalaoMadeireira)()
        Dim worksheet As Worksheet = xlWorkbook.Sheets(1)
        
        ' Mapear colunas específicas da madeireira
        Dim mapeamentoColunas = MapearColunasEspecificas(worksheet)
        LogDebug($"Colunas mapeadas: {mapeamentoColunas.Count}")
        
        ' Processar dados linha por linha
        Dim linha = 2 ' Começar após cabeçalho
        Dim talaoAtual As DadosTalaoMadeireira = Nothing
        
        Do While Not String.IsNullOrEmpty(worksheet.Cells(linha, 1).Value?.ToString())
            
            ' Verificar se é início de novo talão
            Dim numeroTalao = ObterValorCelula(worksheet, linha, mapeamentoColunas, "NUMERO_TALAO")
            
            If Not String.IsNullOrEmpty(numeroTalao) Then
                ' Finalizar talão anterior se existir
                If talaoAtual IsNot Nothing Then
                    taloes.Add(talaoAtual)
                End If
                
                ' Criar novo talão
                talaoAtual = New DadosTalaoMadeireira()
                talaoAtual.NumeroTalao = numeroTalao
                talaoAtual.FormatoDetectado = "MADEIREIRA"
                talaoAtual.OrigemBackup = xlWorkbook.FullName
                
                ' Processar dados do cliente
                ProcessarDadosClienteMadeireira(worksheet, linha, mapeamentoColunas, talaoAtual)
            End If
            
            ' Processar produto atual
            If talaoAtual IsNot Nothing Then
                Dim produto = ProcessarProdutoMadeireira(worksheet, linha, mapeamentoColunas)
                If produto IsNot Nothing Then
                    talaoAtual.Produtos.Add(produto)
                End If
            End If
            
            linha += 1
        Loop
        
        ' Adicionar último talão
        If talaoAtual IsNot Nothing Then
            taloes.Add(talaoAtual)
        End If
        
        LogDebug($"Processamento concluído: {taloes.Count} talões, {taloes.Sum(Function(t) t.Produtos.Count)} produtos")
        
        Return taloes
    End Function
    
    ''' <summary>
    ''' Processa planilha no formato genérico
    ''' </summary>
    Private Function ProcessarFormatoGenerico(xlWorkbook As Workbook) As List(Of DadosTalaoMadeireira)
        LogDebug("=== PROCESSAMENTO FORMATO GENÉRICO ===")
        
        Dim taloes As New List(Of DadosTalaoMadeireira)()
        Dim worksheet As Worksheet = xlWorkbook.Sheets(1)
        
        ' Detectar colunas automaticamente
        Dim mapeamentoColunas = DetectarColunasAutomaticamente(worksheet)
        LogDebug($"Colunas detectadas: {mapeamentoColunas.Count}")
        
        ' Processar dados de forma inteligente
        Dim linha = 2
        Dim talaoAtual As DadosTalaoMadeireira = Nothing
        
        Do While Not String.IsNullOrEmpty(worksheet.Cells(linha, 1).Value?.ToString())
            
            ' Tentar extrair número do talão
            Dim numeroTalao = ExtrairNumeroTalaoGenerico(worksheet, linha, mapeamentoColunas)
            
            If Not String.IsNullOrEmpty(numeroTalao) Then
                If talaoAtual IsNot Nothing Then
                    taloes.Add(talaoAtual)
                End If
                
                talaoAtual = New DadosTalaoMadeireira()
                talaoAtual.NumeroTalao = numeroTalao
                talaoAtual.FormatoDetectado = "GENERICO"
                talaoAtual.OrigemBackup = xlWorkbook.FullName
                
                ProcessarDadosClienteGenerico(worksheet, linha, mapeamentoColunas, talaoAtual)
            End If
            
            If talaoAtual IsNot Nothing Then
                Dim produto = ProcessarProdutoGenerico(worksheet, linha, mapeamentoColunas)
                If produto IsNot Nothing Then
                    talaoAtual.Produtos.Add(produto)
                End If
            End If
            
            linha += 1
        Loop
        
        If talaoAtual IsNot Nothing Then
            taloes.Add(talaoAtual)
        End If
        
        LogDebug($"Processamento genérico concluído: {taloes.Count} talões")
        
        Return taloes
    End Function
    
    ''' <summary>
    ''' Gera novo talão formatado a partir dos dados importados
    ''' </summary>
    Public Function GerarTalaoFormatado(talao As DadosTalaoMadeireira) As String
        Try
            LogDebug($"=== GERAÇÃO TALÃO FORMATADO ===")
            LogDebug($"Talão: {talao.NumeroTalao}")
            
            ' Converter para formato compatível com ExcelAutomation
            Dim dadosCompativel = ConverterParaFormatoCompativel(talao)
            
            ' Usar ExcelAutomation existente
            Dim excel As New ExcelAutomation()
            excel.ProcessarTalaoCompleto(dadosCompativel)
            
            ' Gerar nome do arquivo
            Dim nomeArquivo = $"Talao_{talao.NumeroTalao}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
            Dim caminhoCompleto = Path.Combine(config.CaminhoTaloesGerados, nomeArquivo)
            
            LogDebug($"Talão gerado: {caminhoCompleto}")
            
            Return caminhoCompleto
            
        Catch ex As Exception
            LogDebug($"ERRO na geração: {ex.Message}")
            Throw New InvalidOperationException($"Erro ao gerar talão: {ex.Message}", ex)
        End Try
    End Function
    
    ' === MÉTODOS AUXILIARES ===
    
    Private Sub CriarDiretoriosBackup()
        For Each caminho In {config.CaminhoBackupsImportados, config.CaminhoTaloesGerados, config.CaminhoBackupJSON}
            If Not Directory.Exists(caminho) Then
                Directory.CreateDirectory(caminho)
                LogDebug($"Diretório criado: {caminho}")
            End If
        Next
    End Sub
    
    Private Sub SalvarBackupJSON(taloes As List(Of DadosTalaoMadeireira), arquivoOrigem As String)
        Try
            Dim nomeBackup = $"{config.PrefixoArquivoBackup}{DateTime.Now.ToString(config.FormatoDataBackup)}.json"
            Dim caminhoBackup = Path.Combine(config.CaminhoBackupJSON, nomeBackup)
            
            Dim json = JsonConvert.SerializeObject(taloes, Formatting.Indented)
            File.WriteAllText(caminhoBackup, json)
            
            LogDebug($"Backup JSON salvo: {caminhoBackup}")
            
        Catch ex As Exception
            LogDebug($"Erro ao salvar backup JSON: {ex.Message}")
        End Try
    End Sub
    
    Private Function ConverterParaFormatoCompativel(talao As DadosTalaoMadeireira) As DadosTalao
        Dim dadosCompativel As New DadosTalao()
        
        ' Dados básicos
        dadosCompativel.NumeroTalao = talao.NumeroTalao
        dadosCompativel.NomeCliente = talao.NomeCliente
        dadosCompativel.EnderecoCliente = talao.EnderecoCliente
        dadosCompativel.CEP = talao.CEP
        dadosCompativel.Cidade = talao.Cidade
        dadosCompativel.Telefone = talao.Telefone
        dadosCompativel.FormaPagamento = talao.FormaPagamento
        dadosCompativel.Vendedor = talao.Vendedor
        dadosCompativel.DataVenda = talao.DataEmissao
        
        ' Converter produtos
        For Each produtoMadeireira In talao.Produtos
            Dim produto As New ProdutoTalao()
            produto.Descricao = produtoMadeireira.DescricaoCompleta
            produto.Quantidade = produtoMadeireira.Quantidade
            produto.Unidade = produtoMadeireira.Unidade
            produto.PrecoUnitario = produtoMadeireira.PrecoUnitario
            produto.PrecoTotal = produtoMadeireira.ValorTotal
            
            dadosCompativel.Produtos.Add(produto)
        Next
        
        Return dadosCompativel
    End Function
    
    Private Sub LogDebug(mensagem As String)
        If logDebug Then
            Debug.WriteLine($"[BACKUP-TALAO] {DateTime.UtcNow:HH:mm:ss.fff} - {mensagem}")
        End If
    End Sub
    
    ' === MÉTODOS DE PROCESSAMENTO (IMPLEMENTAÇÃO SIMPLIFICADA) ===
    
    Private Function MapearColunasEspecificas(worksheet As Worksheet) As Dictionary(Of String, Integer)
        ' Implementação simplificada - mapear colunas conhecidas
        Return New Dictionary(Of String, Integer)() From {
            {"NUMERO_TALAO", 1},
            {"CLIENTE", 2},
            {"PRODUTO", 3},
            {"QUANTIDADE", 4},
            {"UNIDADE", 5},
            {"PRECO", 6}
        }
    End Function
    
    Private Function DetectarColunasAutomaticamente(worksheet As Worksheet) As Dictionary(Of String, Integer)
        ' Implementação simplificada - detectar colunas automaticamente
        Return MapearColunasEspecificas(worksheet)
    End Function
    
    Private Function ObterValorCelula(worksheet As Worksheet, linha As Integer, mapeamento As Dictionary(Of String, Integer), chave As String) As String
        If mapeamento.ContainsKey(chave) Then
            Return worksheet.Cells(linha, mapeamento(chave)).Value?.ToString() ?? ""
        End If
        Return ""
    End Function
    
    Private Sub ProcessarDadosClienteMadeireira(worksheet As Worksheet, linha As Integer, mapeamento As Dictionary(Of String, Integer), talao As DadosTalaoMadeireira)
        talao.NomeCliente = ObterValorCelula(worksheet, linha, mapeamento, "CLIENTE")
        ' Implementar demais campos conforme necessário
    End Sub
    
    Private Function ProcessarProdutoMadeireira(worksheet As Worksheet, linha As Integer, mapeamento As Dictionary(Of String, Integer)) As ProdutoTalaoMadeireira
        Dim produto As New ProdutoTalaoMadeireira()
        produto.Descricao = ObterValorCelula(worksheet, linha, mapeamento, "PRODUTO")
        
        If Decimal.TryParse(ObterValorCelula(worksheet, linha, mapeamento, "QUANTIDADE"), produto.Quantidade) AndAlso
           Decimal.TryParse(ObterValorCelula(worksheet, linha, mapeamento, "PRECO"), produto.PrecoUnitario) Then
            produto.Unidade = ObterValorCelula(worksheet, linha, mapeamento, "UNIDADE")
            Return produto
        End If
        
        Return Nothing
    End Function
    
    Private Function ExtrairNumeroTalaoGenerico(worksheet As Worksheet, linha As Integer, mapeamento As Dictionary(Of String, Integer)) As String
        Return ObterValorCelula(worksheet, linha, mapeamento, "NUMERO_TALAO")
    End Function
    
    Private Sub ProcessarDadosClienteGenerico(worksheet As Worksheet, linha As Integer, mapeamento As Dictionary(Of String, Integer), talao As DadosTalaoMadeireira)
        ProcessarDadosClienteMadeireira(worksheet, linha, mapeamento, talao)
    End Sub
    
    Private Function ProcessarProdutoGenerico(worksheet As Worksheet, linha As Integer, mapeamento As Dictionary(Of String, Integer)) As ProdutoTalaoMadeireira
        Return ProcessarProdutoMadeireira(worksheet, linha, mapeamento)
    End Function
    
    ' === PROPRIEDADES PÚBLICAS ===
    
    Public ReadOnly Property UltimaImportacao As List(Of DadosTalaoMadeireira)
        Get
            Return ultimaImportacao
        End Get
    End Property
    
    Public ReadOnly Property UltimoArquivoImportado As String
        Get
            Return ultimoArquivoImportado
        End Get
    End Property
    
End Class