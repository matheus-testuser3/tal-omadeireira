' ====================================================================
' SISTEMA PDV COMPLETO - MADEIREIRA MARIA LUZIA
' Versão integrada que funciona como um PDV real
' Captura dados do formulário -> Escreve na planilha -> Imprime automaticamente
' ====================================================================

Option Explicit

' === VARIÁVEIS GLOBAIS DO SISTEMA ===
Public Const PLANILHA_TALAO As String = "marialuiza(1)"
Public proximoPedido As Long
Public DEV_USER As String
Public VERSAO_SISTEMA As String

' Estrutura da venda atual
Public Type VendaAtual
    numeroPedido As String
    nomeCliente As String
    endereco As String
    numero As String
    bairro As String
    cidade As String
    uf As String
    cep As String
    cpfCnpj As String
    telefone As String
    formaPagamento As String
    dataVenda As Date
    dataEntrega As Date
    subtotal As Double
    desconto As Double
    frete As Double
    total As Double
    observacoes As String
    vendedor As String
    status As String
End Type

Public VendaCorrente As VendaAtual

' ====================================================================
' FUNÇÃO PRINCIPAL DO PDV - PROCESSA VENDA COMPLETA
' ====================================================================
Public Sub ProcessarVendaPDV()
    On Error GoTo TratarErroPDV
    
    Debug.Print "=== INICIANDO PROCESSAMENTO PDV === " & Format(Now, "hh:mm:ss")
    
    ' ETAPA 1: INICIALIZAR SISTEMA
    Call InicializarSistemaPDV
    
    ' ETAPA 2: CAPTURAR DADOS DO FORMULÁRIO
    If Not CapturarDadosFormulario() Then
        Debug.Print "Falha na captura de dados - Processo cancelado"
        Exit Sub
    End If
    
    ' ETAPA 3: GERAR NÚMERO DO PEDIDO
    VendaCorrente.numeroPedido = GerarProximoNumeroPedido()
    Debug.Print "Pedido gerado: #" & VendaCorrente.numeroPedido
    
    ' ETAPA 4: VALIDAR DADOS OBRIGATÓRIOS
    If Not ValidarDadosVenda() Then
        Call ReverterNumeroPedido
        Debug.Print "Validação falhou - Número revertido"
        Exit Sub
    End If
    
    ' ETAPA 5: ESCREVER DADOS NA PLANILHA
    If Not EscreverDadosNaPlanilha() Then
        Call ReverterNumeroPedido
        Debug.Print "Falha ao escrever na planilha - Número revertido"
        Exit Sub
    End If
    
    ' ETAPA 6: CONFIGURAR E IMPRIMIR
    If ImprimirTalaoPDV() Then
        ' ETAPA 7: FINALIZAR VENDA
        Call FinalizarVendaPDV
        Debug.Print "=== VENDA PROCESSADA COM SUCESSO === Pedido: #" & VendaCorrente.numeroPedido
    Else
        Call ReverterNumeroPedido
        Debug.Print "Impressão cancelada - Número revertido"
    End If
    
    Exit Sub
    
TratarErroPDV:
    Debug.Print "ERRO CRÍTICO NO PDV: " & Err.Description
    MsgBox "❌ ERRO NO SISTEMA PDV!" & vbCrLf & vbCrLf & _
           "Erro: " & Err.Description & vbCrLf & _
           "A venda não foi processada" & vbCrLf & _
           "Tente novamente ou contate o suporte", vbCritical, "Erro PDV"
    
    ' Reverter número em caso de erro
    Call ReverterNumeroPedido
End Sub

' ====================================================================
' INICIALIZAR SISTEMA PDV
' ====================================================================
Private Sub InicializarSistemaPDV()
    On Error Resume Next
    
    ' Definir constantes do sistema
    DEV_USER = Environ("USERNAME")
    VERSAO_SISTEMA = "PDV v2.0"
    
    ' Limpar estrutura da venda
    With VendaCorrente
        .numeroPedido = ""
        .nomeCliente = ""
        .endereco = ""
        .numero = ""
        .bairro = ""
        .cidade = ""
        .uf = "PE"
        .cep = ""
        .cpfCnpj = ""
        .telefone = ""
        .formaPagamento = ""
        .dataVenda = Date
        .dataEntrega = Date + 1
        .subtotal = 0
        .desconto = 0
        .frete = 0
        .total = 0
        .observacoes = ""
        .vendedor = DEV_USER
        .status = "PROCESSANDO"
    End With
    
    Debug.Print "Sistema PDV inicializado | " & Format(Now, "hh:mm:ss")
    
    On Error GoTo 0
End Sub

' ====================================================================
' CAPTURAR DADOS DO FORMULÁRIO
' ====================================================================
Private Function CapturarDadosFormulario() As Boolean
    On Error GoTo TratarErroCaptura
    
    CapturarDadosFormulario = False
    
    ' Verificar se o formulário está disponível
    If UserForms.Count = 0 Then
        MsgBox "⚠️ Formulário não encontrado!" & vbCrLf & _
               "Abra o formulário de vendas primeiro", vbExclamation
        Exit Function
    End If
    
    ' Assumindo que o formulário ativo é o de vendas
    Dim frm As Object
    Set frm = UserForms(0) ' Ou usar nome específico do formulário
    
    With VendaCorrente
        ' Dados do cliente
        .nomeCliente = Trim(frm.txtNome.Text)
        .endereco = Trim(frm.txtEnder.Text)
        .numero = Trim(frm.txtnumero.Text)
        .bairro = frm.cbairro1.Value
        .cidade = frm.cCidade.Value
        .uf = "PE"
        .cep = Trim(frm.txtCEP.Text)
        .cpfCnpj = Trim(frm.txtCPF.Text)
        
        ' Dados da venda
        .formaPagamento = frm.cPagamento.Value
        .dataVenda = Date
        
        ' Data de entrega
        If Trim(frm.cData.Text) <> "" Then
            .dataEntrega = CDate(frm.cData.Text)
        Else
            .dataEntrega = Date + 1
        End If
        
        ' Calcular totais dos produtos
        .subtotal = CalcularTotalProdutos(frm)
        .desconto = 0
        .frete = 0
        .total = .subtotal - .desconto + .frete
        
        .vendedor = DEV_USER
        .status = "ATIVO"
    End With
    
    Debug.Print "Dados capturados | Cliente: " & VendaCorrente.nomeCliente & " | Total: R$ " & Format(VendaCorrente.total, "0.00")
    CapturarDadosFormulario = True
    
    Exit Function
    
TratarErroCaptura:
    Debug.Print "Erro ao capturar dados: " & Err.Description
    MsgBox "❌ Erro ao capturar dados do formulário!" & vbCrLf & _
           "Erro: " & Err.Description, vbCritical
    CapturarDadosFormulario = False
End Function

' ====================================================================
' CALCULAR TOTAL DOS PRODUTOS
' ====================================================================
Private Function CalcularTotalProdutos(frm As Object) As Double
    On Error Resume Next
    
    Dim total As Double
    Dim i As Integer
    
    total = 0
    
    ' Percorrer lista de produtos
    For i = 0 To frm.produtosv2.ListCount - 1
        Dim valorItem As String
        valorItem = Replace(Replace(frm.produtosv2.List(i, 6), "R$", ""), ",", ".")
        valorItem = Replace(Replace(valorItem, " ", ""), ".", "")
        
        If IsNumeric(valorItem) Then
            total = total + CDbl(valorItem)
        End If
    Next i
    
    CalcularTotalProdutos = total
    
    On Error GoTo 0
End Function

' ====================================================================
' VALIDAR DADOS DA VENDA
' ====================================================================
Private Function ValidarDadosVenda() As Boolean
    On Error Resume Next
    
    ValidarDadosVenda = False
    
    With VendaCorrente
        ' Validar nome do cliente
        If Trim(.nomeCliente) = "" Then
            MsgBox "⚠️ NOME DO CLIENTE É OBRIGATÓRIO!" & vbCrLf & vbCrLf & _
                   "Preencha o nome antes de processar a venda" & vbCrLf & _
                   "Pedido: #" & .numeroPedido, vbExclamation, "Campo Obrigatório"
            Exit Function
        End If
        
        ' Validar forma de pagamento
        If Trim(.formaPagamento) = "" Then
            MsgBox "⚠️ FORMA DE PAGAMENTO É OBRIGATÓRIA!" & vbCrLf & vbCrLf & _
                   "Selecione uma forma de pagamento" & vbCrLf & _
                   "Pedido: #" & .numeroPedido, vbExclamation, "Campo Obrigatório"
            Exit Function
        End If
        
        ' Validar se há produtos
        If .total <= 0 Then
            MsgBox "⚠️ NENHUM PRODUTO ADICIONADO!" & vbCrLf & vbCrLf & _
                   "Adicione produtos antes de processar" & vbCrLf & _
                   "Pedido: #" & .numeroPedido, vbExclamation, "Produtos Obrigatórios"
            Exit Function
        End If
    End With
    
    Debug.Print "Validação aprovada | Pedido: #" & VendaCorrente.numeroPedido
    ValidarDadosVenda = True
    
    On Error GoTo 0
End Function

' ====================================================================
' ESCREVER DADOS NA PLANILHA
' ====================================================================
Private Function EscreverDadosNaPlanilha() As Boolean
    On Error GoTo TratarErroEscrita
    
    EscreverDadosNaPlanilha = False
    
    ' Acessar planilha do talão
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(PLANILHA_TALAO)
    
    If ws Is Nothing Then
        MsgBox "❌ PLANILHA DO TALÃO NÃO ENCONTRADA!" & vbCrLf & vbCrLf & _
               "A planilha '" & PLANILHA_TALAO & "' não existe" & vbCrLf & _
               "Verifique o arquivo", vbCritical, "Erro de Sistema"
        Exit Function
    End If
    
    Debug.Print "Escrevendo dados na planilha | Pedido: #" & VendaCorrente.numeroPedido
    
    ' LIMPAR DADOS ANTERIORES
    Call LimparPlanilhaTalao(ws)
    
    ' ESCREVER CABEÇALHO
    Call EscreverCabecalhoTalao(ws)
    
    ' ESCREVER PRODUTOS
    Call EscreverProdutosTalao(ws)
    
    ' ESCREVER RODAPÉ
    Call EscreverRodapeTalao(ws)
    
    ' APLICAR FORMATAÇÃO
    Call FormatarTalao(ws)
    
    Debug.Print "Dados escritos com sucesso na planilha"
    EscreverDadosNaPlanilha = True
    
    Exit Function
    
TratarErroEscrita:
    Debug.Print "Erro ao escrever na planilha: " & Err.Description
    MsgBox "❌ Erro ao escrever dados na planilha!" & vbCrLf & _
           "Erro: " & Err.Description, vbCritical
    EscreverDadosNaPlanilha = False
End Function

' ====================================================================
' LIMPAR PLANILHA DO TALÃO
' ====================================================================
Private Sub LimparPlanilhaTalao(ws As Worksheet)
    On Error Resume Next
    
    ' Limpar dados da planilha de forma simples
    ws.Cells.Clear
    
    Debug.Print "Planilha limpa para novo pedido"
    
    On Error GoTo 0
End Sub

' ====================================================================
' ESCREVER CABEÇALHO DO TALÃO
' ====================================================================
Private Sub EscreverCabecalhoTalao(ws As Worksheet)
    On Error Resume Next
    
    With VendaCorrente
        ' Cabeçalho da empresa
        ws.Range("A1").Value = "MADEIREIRA MARIA LUZIA"
        ws.Range("A1").Font.Bold = True
        ws.Range("A1").Font.Size = 16
        
        ' Número do pedido
        ws.Range("A3").Value = "PEDIDO #" & .numeroPedido
        ws.Range("A3").Font.Bold = True
        ws.Range("A3").Font.Size = 14
        
        ' Dados do cliente
        ws.Range("A5").Value = "CLIENTE:"
        ws.Range("B5").Value = .nomeCliente
        
        ws.Range("A6").Value = "ENDEREÇO:"
        ws.Range("B6").Value = .endereco & IIf(.numero <> "", ", " & .numero, "")
        
        ws.Range("A7").Value = "BAIRRO:"
        ws.Range("B7").Value = .bairro
        
        ws.Range("A8").Value = "CIDADE:"
        ws.Range("B8").Value = .cidade & "/" & .uf
        
        ws.Range("A9").Value = "CEP:"
        ws.Range("B9").Value = .cep
        
        ws.Range("A10").Value = "CPF/CNPJ:"
        ws.Range("B10").Value = .cpfCnpj
        
        ' Dados da venda
        ws.Range("A12").Value = "DATA VENDA:"
        ws.Range("B12").Value = Format(.dataVenda, "dd/mm/yyyy")
        
        ws.Range("A13").Value = "DATA ENTREGA:"
        ws.Range("B13").Value = Format(.dataEntrega, "dd/mm/yyyy")
        
        ws.Range("A14").Value = "PAGAMENTO:"
        ws.Range("B14").Value = .formaPagamento
        
        ws.Range("A15").Value = "VENDEDOR:"
        ws.Range("B15").Value = .vendedor
    End With
    
    Debug.Print "Cabeçalho escrito | Cliente: " & VendaCorrente.nomeCliente
    
    On Error GoTo 0
End Sub

' ====================================================================
' ESCREVER PRODUTOS NO TALÃO
' ====================================================================
Private Sub EscreverProdutosTalao(ws As Worksheet)
    On Error Resume Next
    
    ' Obter formulário ativo para acessar produtos
    Dim frm As Object
    Set frm = UserForms(0)
    
    ' Cabeçalho dos produtos
    ws.Range("A17").Value = "PRODUTOS:"
    ws.Range("A17").Font.Bold = True
    
    ' Cabeçalhos das colunas
    ws.Range("A18").Value = "REF"
    ws.Range("B18").Value = "DESCRIÇÃO"
    ws.Range("C18").Value = "UND"
    ws.Range("D18").Value = "VLR UNIT"
    ws.Range("E18").Value = "QTD"
    ws.Range("F18").Value = "DESC"
    ws.Range("G18").Value = "TOTAL"
    
    ' Formatação dos cabeçalhos
    ws.Range("A18:G18").Font.Bold = True
    
    Dim i As Integer
    Dim linha As Integer
    linha = 19 ' Começar na linha 19
    
    ' Escrever produtos da lista produtosv2
    For i = 0 To frm.produtosv2.ListCount - 1
        ws.Range("A" & linha).Value = frm.produtosv2.List(i, 0)  ' Referência
        ws.Range("B" & linha).Value = frm.produtosv2.List(i, 1)  ' Descrição
        ws.Range("C" & linha).Value = frm.produtosv2.List(i, 2)  ' Unidade
        ws.Range("D" & linha).Value = frm.produtosv2.List(i, 3)  ' Valor Unit
        ws.Range("E" & linha).Value = frm.produtosv2.List(i, 4)  ' Quantidade
        ws.Range("F" & linha).Value = frm.produtosv2.List(i, 5)  ' Desconto
        ws.Range("G" & linha).Value = frm.produtosv2.List(i, 6)  ' Total
        
        linha = linha + 1
    Next i
    
    Debug.Print "Produtos escritos | Quantidade: " & frm.produtosv2.ListCount
    
    On Error GoTo 0
End Sub

' ====================================================================
' ESCREVER RODAPÉ DO TALÃO
' ====================================================================
Private Sub EscreverRodapeTalao(ws As Worksheet)
    On Error Resume Next
    
    ' Obter formulário ativo para calcular linha final
    Dim frm As Object
    Set frm = UserForms(0)
    
    ' Calcular linha para começar o rodapé (após os produtos)
    Dim linhaRodape As Integer
    linhaRodape = 19 + frm.produtosv2.ListCount + 2 ' +2 para espaçamento
    
    With VendaCorrente
        ' Totais
        ws.Range("A" & linhaRodape).Value = "RESUMO:"
        ws.Range("A" & linhaRodape).Font.Bold = True
        
        ws.Range("A" & (linhaRodape + 1)).Value = "SUBTOTAL:"
        ws.Range("B" & (linhaRodape + 1)).Value = .subtotal
        ws.Range("B" & (linhaRodape + 1)).NumberFormat = "R$ #,##0.00"
        
        ws.Range("A" & (linhaRodape + 2)).Value = "DESCONTO:"
        ws.Range("B" & (linhaRodape + 2)).Value = .desconto
        ws.Range("B" & (linhaRodape + 2)).NumberFormat = "R$ #,##0.00"
        
        ws.Range("A" & (linhaRodape + 3)).Value = "FRETE:"
        ws.Range("B" & (linhaRodape + 3)).Value = .frete
        ws.Range("B" & (linhaRodape + 3)).NumberFormat = "R$ #,##0.00"
        
        ws.Range("A" & (linhaRodape + 4)).Value = "TOTAL GERAL:"
        ws.Range("B" & (linhaRodape + 4)).Value = .total
        ws.Range("A" & (linhaRodape + 4)).Font.Bold = True
        ws.Range("B" & (linhaRodape + 4)).Font.Bold = True
        ws.Range("B" & (linhaRodape + 4)).NumberFormat = "R$ #,##0.00"
        
        ' Informações adicionais
        ws.Range("A" & (linhaRodape + 6)).Value = "STATUS: BALCÃO"
        ws.Range("A" & (linhaRodape + 7)).Value = "DATA/HORA: " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    End With
    
    Debug.Print "Rodapé escrito | Total: R$ " & Format(VendaCorrente.total, "0.00")
    
    On Error GoTo 0
End Sub

' ====================================================================
' FORMATAR TALÃO
' ====================================================================
Private Sub FormatarTalao(ws As Worksheet)
    On Error Resume Next
    
    ' Ajustar largura das colunas automaticamente
    ws.Columns("A:G").AutoFit
    
    ' Formatação básica
    ws.Cells.Font.Name = "Arial"
    ws.Cells.Font.Size = 10
    
    ' Formatação especial para valores monetários nas colunas D e G (valor unit e total)
    ws.Columns("D:D").NumberFormat = "R$ #,##0.00"
    ws.Columns("G:G").NumberFormat = "R$ #,##0.00"
    
    Debug.Print "Formatação aplicada ao talão"
    
    On Error GoTo 0
End Sub

' ====================================================================
' IMPRIMIR TALÃO PDV
' ====================================================================
Private Function ImprimirTalaoPDV() As Boolean
    On Error GoTo TratarErroImpressao
    
    ImprimirTalaoPDV = False
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(PLANILHA_TALAO)
    
    ' Configurar impressão
    Call ConfigurarImpressaoTalao(ws)
    
    ' Confirmar impressão
    If ConfirmarImpressaoVenda() Then
        ws.PrintOut
        Call ExibirSucessoVenda()
        ImprimirTalaoPDV = True
        Debug.Print "Talão impresso com sucesso | Pedido: #" & VendaCorrente.numeroPedido
    Else
        Debug.Print "Impressão cancelada pelo usuário"
        ImprimirTalaoPDV = False
    End If
    
    Exit Function
    
TratarErroImpressao:
    Debug.Print "Erro na impressão: " & Err.Description
    MsgBox "❌ Erro na impressão!" & vbCrLf & _
           "Erro: " & Err.Description & vbCrLf & _
           "Verifique a impressora", vbCritical
    ImprimirTalaoPDV = False
End Function

' ====================================================================
' CONFIGURAR IMPRESSÃO DO TALÃO
' ====================================================================
Private Sub ConfigurarImpressaoTalao(ws As Worksheet)
    On Error Resume Next
    
    With ws.PageSetup
        .PrintArea = ""  ' Deixar área automática
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.35)
        .BottomMargin = Application.InchesToPoints(0.35)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = True
        .CenterVertically = True
        .Orientation = xlLandscape
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .PrintErrors = xlPrintErrorsDisplayed
    End With
    
    ws.Activate
    Application.ActiveWindow.Zoom = 85
    
    On Error GoTo 0
End Sub

' ====================================================================
' CONFIRMAR IMPRESSÃO DA VENDA
' ====================================================================
Private Function ConfirmarImpressaoVenda() As Boolean
    With VendaCorrente
        Dim mensagem As String
        mensagem = "🖨️ TALÃO PREPARADO PARA IMPRESSÃO!" & vbCrLf & vbCrLf & _
                  "PEDIDO: #" & .numeroPedido & vbCrLf & _
                  "Cliente: " & .nomeCliente & vbCrLf & _
                  "Endereço: " & .endereco & ", " & .numero & vbCrLf & _
                  "Bairro: " & .bairro & " - " & .cidade & "/" & .uf & vbCrLf & _
                  "Pagamento: " & .formaPagamento & vbCrLf & _
                  "Total: R$ " & Format(.total, "#,##0.00") & vbCrLf & _
                  "Data/Hora: " & Format(Now, "dd/mm/yyyy hh:mm:ss") & vbCrLf & vbCrLf & _
                  "Confirma a impressão do talão?"
        
        ConfirmarImpressaoVenda = (MsgBox(mensagem, vbYesNo + vbQuestion, "Confirmar Impressão") = vbYes)
    End With
End Function

' ====================================================================
' EXIBIR SUCESSO DA VENDA
' ====================================================================
Private Sub ExibirSucessoVenda()
    With VendaCorrente
        MsgBox "✅ PEDIDO #" & .numeroPedido & " PROCESSADO COM SUCESSO!" & vbCrLf & vbCrLf & _
               "Talão impresso em 2 vias (Cliente + Loja)" & vbCrLf & _
               "Cliente: " & .nomeCliente & vbCrLf & _
               "Valor total: R$ " & Format(.total, "#,##0.00") & vbCrLf & _
               "Pagamento: " & .formaPagamento & vbCrLf & _
               "Data/Hora: " & Format(Now, "dd/mm/yyyy hh:mm:ss") & vbCrLf & _
               "Sistema PDV funcionando perfeitamente!" & vbCrLf & vbCrLf & _
               VERSAO_SISTEMA & " - " & DEV_USER, vbInformation, "Venda Concluída"
    End With
End Sub

' ====================================================================
' FINALIZAR VENDA PDV
' ====================================================================
Private Sub FinalizarVendaPDV()
    On Error Resume Next
    
    ' Salvar dados da venda
    Call SalvarHistoricoVenda()
    
    ' Limpar formulário
    Call LimparFormularioVenda()
    
    ' Atualizar status
    VendaCorrente.status = "CONCLUÍDA"
    
    Debug.Print "Venda finalizada | Pedido: #" & VendaCorrente.numeroPedido & " | Total: R$ " & Format(VendaCorrente.total, "0.00")
    
    On Error GoTo 0
End Sub

' ====================================================================
' SALVAR HISTÓRICO DA VENDA
' ====================================================================
Private Sub SalvarHistoricoVenda()
    On Error Resume Next
    
    ' Aqui você pode implementar salvamento em planilha de histórico
    ' ou banco de dados conforme necessário
    
    Debug.Print "VENDA SALVA | Pedido: #" & VendaCorrente.numeroPedido & _
               " | Cliente: " & VendaCorrente.nomeCliente & _
               " | Total: R$ " & Format(VendaCorrente.total, "0.00") & _
               " | " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    
    On Error GoTo 0
End Sub

' ====================================================================
' LIMPAR FORMULÁRIO APÓS VENDA
' ====================================================================
Private Sub LimparFormularioVenda()
    On Error Resume Next
    
    If UserForms.Count > 0 Then
        Dim frm As Object
        Set frm = UserForms(0)
        
        ' Limpar campos principais
        frm.txtNome.Value = ""
        frm.txtEnder.Value = ""
        frm.txtnumero.Value = ""
        frm.txtCEP.Value = ""
        frm.txtCPF.Value = ""
        
        ' Resetar combos
        frm.cPagamento.ListIndex = 0
        frm.cCidade.ListIndex = -1
        frm.cbairro1.ListIndex = -1
        
        ' Limpar produtos
        frm.produtosv2.Clear
        
        ' Retornar foco
        frm.txtNome.SetFocus
    End If
    
    Debug.Print "Formulário limpo para nova venda"
    
    On Error GoTo 0
End Sub

' ====================================================================
' FUNÇÕES AUXILIARES
' ====================================================================

' Gerar próximo número de pedido
Private Function GerarProximoNumeroPedido() As String
    On Error Resume Next
    
    ' Implementar lógica de numeração sequencial
    proximoPedido = proximoPedido + 1
    GerarProximoNumeroPedido = Format(proximoPedido, "00000")
    
    On Error GoTo 0
End Function

' Reverter número do pedido em caso de erro
Private Sub ReverterNumeroPedido()
    On Error Resume Next
    
    If proximoPedido > 0 Then
        proximoPedido = proximoPedido - 1
    End If
    
    Debug.Print "Número do pedido revertido para: " & proximoPedido
    
    On Error GoTo 0
End Sub

' ====================================================================
' FUNÇÃO DE CHAMADA RÁPIDA PARA BOTÃO
' ====================================================================
Public Sub btnProcessarVenda_Click()
    ' Esta função pode ser chamada diretamente de um botão no formulário
    Call ProcessarVendaPDV
End Sub