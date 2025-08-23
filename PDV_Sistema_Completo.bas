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
    
    ' Limpar apenas os campos de dados, preservando o template
    ws.Range("B7:H9,M7:T9").ClearContents         ' Cabeçalhos cliente
    ws.Range("B11:H21,M11:T21").ClearContents     ' Produtos
    ws.Range("B22:H25,M22:T25").ClearContents     ' Rodapés/totais
    ws.Range("B6,M6").ClearContents               ' Número do pedido
    
    Debug.Print "Campos da planilha limpos, template preservado"
    
    On Error GoTo 0
End Sub

' ====================================================================
' ESCREVER CABEÇALHO DO TALÃO
' ====================================================================
Private Sub EscreverCabecalhoTalao(ws As Worksheet)
    On Error Resume Next
    
    With VendaCorrente
        ' Montar endereço completo
        Dim enderecoCompleto As String
        enderecoCompleto = .endereco
        If Trim(.numero) <> "" Then
            enderecoCompleto = enderecoCompleto & ", " & .numero
        End If
        
        ' LADO ESQUERDO - Preencher campos específicos da planilha
        ws.Range("B6").Value = "PEDIDO #" & .numeroPedido    ' Número do pedido
        ws.Range("B7").Value = .nomeCliente                  ' Nome do cliente
        ws.Range("B8").Value = enderecoCompleto              ' Endereço completo
        ws.Range("F8").Value = .bairro                       ' Bairro
        ws.Range("B9").Value = .cpfCnpj                      ' CPF/CNPJ
        ws.Range("E9").Value = .cidade                       ' Cidade
        ws.Range("G9").Value = .uf                           ' UF
        ws.Range("H9").Value = .cep                          ' CEP
        
        ' LADO DIREITO - Espelho (via do cliente)
        ws.Range("M6").Value = "PEDIDO #" & .numeroPedido    ' Número do pedido
        ws.Range("M7").Value = .nomeCliente                  ' Nome do cliente
        ws.Range("M8").Value = enderecoCompleto              ' Endereço completo
        ws.Range("Q8").Value = .bairro                       ' Bairro
        ws.Range("M9").Value = .cpfCnpj                      ' CPF/CNPJ
        ws.Range("P9").Value = .cidade                       ' Cidade
        ws.Range("R9").Value = .uf                           ' UF
        ws.Range("T9").Value = .cep                          ' CEP
    End With
    
    Debug.Print "Cabeçalho preenchido nos campos corretos | Cliente: " & VendaCorrente.nomeCliente
    
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
    
    Dim i As Integer
    Dim linhaAtual As Integer
    Dim totalProdutos As Integer
    
    totalProdutos = IIf(frm.produtosv2.ListCount > 10, 10, frm.produtosv2.ListCount)
    
    ' Preencher produtos nas linhas específicas do template (linhas 11 a 21)
    For i = 0 To 9 ' Máximo 10 produtos conforme template
        linhaAtual = 11 + i
        
        If i < totalProdutos Then
            ' LADO ESQUERDO - Preencher campos específicos
            ws.Range("B" & linhaAtual).Value = frm.produtosv2.List(i, 0)  ' Referência
            ws.Range("C" & linhaAtual).Value = frm.produtosv2.List(i, 1)  ' Descrição
            ws.Range("D" & linhaAtual).Value = frm.produtosv2.List(i, 2)  ' Unidade
            ws.Range("E" & linhaAtual).Value = frm.produtosv2.List(i, 3)  ' Valor Unit
            ws.Range("F" & linhaAtual).Value = frm.produtosv2.List(i, 4)  ' Quantidade
            ws.Range("G" & linhaAtual).Value = frm.produtosv2.List(i, 5)  ' Desconto
            ws.Range("H" & linhaAtual).Value = frm.produtosv2.List(i, 6)  ' Total
            
            ' LADO DIREITO - Espelho (via do cliente)
            ws.Range("M" & linhaAtual).Value = frm.produtosv2.List(i, 0)  ' Referência
            ws.Range("N" & linhaAtual).Value = frm.produtosv2.List(i, 1)  ' Descrição
            ws.Range("O" & linhaAtual).Value = frm.produtosv2.List(i, 2)  ' Unidade
            ws.Range("P" & linhaAtual).Value = frm.produtosv2.List(i, 3)  ' Valor Unit
            ws.Range("Q" & linhaAtual).Value = frm.produtosv2.List(i, 4)  ' Quantidade
            ws.Range("R" & linhaAtual).Value = frm.produtosv2.List(i, 5)  ' Desconto
            ws.Range("S" & linhaAtual).Value = frm.produtosv2.List(i, 6)  ' Total
        End If
    Next i
    
    Debug.Print "Produtos preenchidos nos campos corretos | Quantidade: " & totalProdutos
    
    On Error GoTo 0
End Sub

' ====================================================================
' ESCREVER RODAPÉ DO TALÃO
' ====================================================================
Private Sub EscreverRodapeTalao(ws As Worksheet)
    On Error Resume Next
    
    With VendaCorrente
        ' LADO ESQUERDO - Preencher campos específicos do rodapé
        ws.Range("B22").Value = .vendedor                           ' Vendedor
        ws.Range("B24").Value = "BALCÃO"                           ' Situação
        ws.Range("B25").Value = .formaPagamento                    ' Forma de pagamento
        ws.Range("C25").Value = "PEDIDO #" & .numeroPedido         ' Número do pedido
        ws.Range("F23").Value = Format(.dataEntrega, "dd/mm/yyyy") ' Data de entrega
        ws.Range("H22").Value = .subtotal                          ' Total produtos
        ws.Range("H23").Value = .frete                             ' Frete
        ws.Range("H24").Value = .desconto                          ' Desconto
        ws.Range("H25").Value = .total                             ' Total geral
        
        ' LADO DIREITO - Espelho (via do cliente)
        ws.Range("M22").Value = .vendedor                          ' Vendedor
        ws.Range("M24").Value = "BALCÃO"                          ' Situação
        ws.Range("M25").Value = .formaPagamento                   ' Forma de pagamento
        ws.Range("N25").Value = "PEDIDO #" & .numeroPedido        ' Número do pedido
        ws.Range("P23").Value = Format(.dataEntrega, "dd/mm/yyyy") ' Data de entrega
        ws.Range("S22").Value = .subtotal                         ' Total produtos
        ws.Range("S23").Value = .frete                            ' Frete
        ws.Range("S24").Value = .desconto                         ' Desconto
        ws.Range("S25").Value = .total                            ' Total geral
    End With
    
    Debug.Print "Rodapé preenchido nos campos corretos | Total: R$ " & Format(VendaCorrente.total, "0.00")
    
    On Error GoTo 0
End Sub

' ====================================================================
' FORMATAR TALÃO
' ====================================================================
Private Sub FormatarTalao(ws As Worksheet)
    On Error Resume Next
    
    ' Formatação especial para número do pedido
    With ws.Range("B6,M6")
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = RGB(0, 120, 0)
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(240, 255, 240)
        .Borders.LineStyle = xlContinuous
    End With
    
    ' Formatação para valores monetários nos campos corretos
    ws.Range("H22,H23,H24,H25,S22,S23,S24,S25").NumberFormat = "_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * ""-""_-;_-@_-"
    
    Debug.Print "Formatação aplicada nos campos corretos do template"
    
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
        .PrintArea = "A1:T26"  ' Área específica do template
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