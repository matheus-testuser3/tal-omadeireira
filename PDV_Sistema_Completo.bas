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
    
    ' Percorrer lista de produtos usando produtosv2
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
    
    ' ESCREVER DADOS DIRETAMENTE NA PLANILHA EXISTENTE
    Call EscreverDadosCliente(ws)
    Call EscreverDadosProdutos(ws)
    Call EscreverDadosVenda(ws)
    
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
' ESCREVER DADOS DO CLIENTE NA PLANILHA
' ====================================================================
Private Sub EscreverDadosCliente(ws As Worksheet)
    On Error Resume Next
    
    With VendaCorrente
        ' Escrever dados do cliente nas células específicas da planilha
        ws.Range("B6").Value = "PEDIDO #" & .numeroPedido
        ws.Range("B7").Value = .nomeCliente
        ws.Range("B8").Value = .endereco & IIf(.numero <> "", ", " & .numero, "")
        ws.Range("F8").Value = .bairro
        ws.Range("B9").Value = .cpfCnpj
        ws.Range("E9").Value = .cidade
        ws.Range("G9").Value = .uf
        ws.Range("H9").Value = .cep
        
        ' Duplicar dados para via do cliente (lado direito)
        ws.Range("M6").Value = "PEDIDO #" & .numeroPedido
        ws.Range("M7").Value = .nomeCliente
        ws.Range("M8").Value = .endereco & IIf(.numero <> "", ", " & .numero, "")
        ws.Range("Q8").Value = .bairro
        ws.Range("M9").Value = .cpfCnpj
        ws.Range("P9").Value = .cidade
        ws.Range("R9").Value = .uf
        ws.Range("T9").Value = .cep
    End With
    
    Debug.Print "Dados do cliente escritos na planilha"
    
    On Error GoTo 0
End Sub

' ====================================================================
' ESCREVER DADOS DOS PRODUTOS NA PLANILHA
' ====================================================================
Private Sub EscreverDadosProdutos(ws As Worksheet)
    On Error Resume Next
    
    ' Obter formulário ativo para acessar produtos
    Dim frm As Object
    Set frm = UserForms(0)
    
    Dim i As Integer
    Dim linhaAtual As Integer
    Dim totalProdutos As Integer
    
    ' Limpar área de produtos primeiro
    ws.Range("B11:H21").ClearContents  ' Lado esquerdo
    ws.Range("M11:T21").ClearContents  ' Lado direito
    
    totalProdutos = frm.produtosv2.ListCount
    
    ' Escrever produtos na planilha (máximo 10 produtos)
    For i = 0 To IIf(totalProdutos > 10, 9, totalProdutos - 1)
        linhaAtual = 11 + i
        
        ' LADO ESQUERDO (Via da loja)
        ws.Range("B" & linhaAtual).Value = frm.produtosv2.List(i, 0)  ' Referência
        ws.Range("C" & linhaAtual).Value = frm.produtosv2.List(i, 1)  ' Descrição
        ws.Range("D" & linhaAtual).Value = frm.produtosv2.List(i, 2)  ' Unidade
        ws.Range("E" & linhaAtual).Value = frm.produtosv2.List(i, 3)  ' Valor Unit
        ws.Range("F" & linhaAtual).Value = frm.produtosv2.List(i, 4)  ' Quantidade
        ws.Range("G" & linhaAtual).Value = frm.produtosv2.List(i, 5)  ' Desconto
        ws.Range("H" & linhaAtual).Value = frm.produtosv2.List(i, 6)  ' Total
        
        ' LADO DIREITO (Via do cliente)
        ws.Range("M" & linhaAtual).Value = frm.produtosv2.List(i, 0)  ' Referência
        ws.Range("N" & linhaAtual).Value = frm.produtosv2.List(i, 1)  ' Descrição
        ws.Range("O" & linhaAtual).Value = frm.produtosv2.List(i, 2)  ' Unidade
        ws.Range("P" & linhaAtual).Value = frm.produtosv2.List(i, 3)  ' Valor Unit
        ws.Range("Q" & linhaAtual).Value = frm.produtosv2.List(i, 4)  ' Quantidade
        ws.Range("R" & linhaAtual).Value = frm.produtosv2.List(i, 5)  ' Desconto
        ws.Range("S" & linhaAtual).Value = frm.produtosv2.List(i, 6)  ' Total
    Next i
    
    Debug.Print "Produtos escritos na planilha | Quantidade: " & totalProdutos
    
    On Error GoTo 0
End Sub

' ====================================================================
' ESCREVER DADOS DA VENDA NA PLANILHA
' ====================================================================
Private Sub EscreverDadosVenda(ws As Worksheet)
    On Error Resume Next
    
    With VendaCorrente
        ' Escrever informações da venda nas células específicas
        ws.Range("B22").Value = "FORMA PAGAMENTO:"
        ws.Range("C22").Value = .formaPagamento
        
        ws.Range("B23").Value = "DATA VENDA:"
        ws.Range("C23").Value = Format(.dataVenda, "dd/mm/yyyy")
        
        ws.Range("B24").Value = "DATA ENTREGA:"
        ws.Range("C24").Value = Format(.dataEntrega, "dd/mm/yyyy")
        
        ws.Range("B25").Value = "TOTAL GERAL:"
        ws.Range("C25").Value = .total
        ws.Range("C25").NumberFormat = "R$ #,##0.00"
        
        ' Duplicar informações para via do cliente
        ws.Range("M22").Value = "FORMA PAGAMENTO:"
        ws.Range("N22").Value = .formaPagamento
        
        ws.Range("M23").Value = "DATA VENDA:"
        ws.Range("N23").Value = Format(.dataVenda, "dd/mm/yyyy")
        
        ws.Range("M24").Value = "DATA ENTREGA:"
        ws.Range("N24").Value = Format(.dataEntrega, "dd/mm/yyyy")
        
        ws.Range("M25").Value = "TOTAL GERAL:"
        ws.Range("N25").Value = .total
        ws.Range("N25").NumberFormat = "R$ #,##0.00"
        
        ' Informações adicionais
        ws.Range("B26").Value = "VENDEDOR: " & .vendedor
        ws.Range("B27").Value = "PROCESSADO: " & Format(Now, "dd/mm/yyyy hh:mm:ss")
        
        ws.Range("M26").Value = "VENDEDOR: " & .vendedor
        ws.Range("M27").Value = "PROCESSADO: " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    End With
    
    Debug.Print "Dados da venda escritos na planilha"
    
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
        
        ' Limpar produtos (usando produtosv2)
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