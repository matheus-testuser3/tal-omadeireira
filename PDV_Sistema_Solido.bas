' ====================================================================
' SISTEMA PDV EXCEL - VERSÃO SÓLIDA E ADAPTADA
' Desenvolvido para VBA Excel com tratamento robusto de erros
' ====================================================================

Option Explicit

' ====================================================================
' DECLARAÇÕES GLOBAIS E CONSTANTES
' ====================================================================
Public Const PLANILHA_TALAO As String = "marialuiza(1)"
Public Const VERSAO_SISTEMA As String = "PDV Excel v2.5"
Public Const COR_BRANCO As Long = RGB(255, 255, 255)
Public Const COR_VERDE_CLARO As Long = RGB(240, 255, 240)
Public Const COR_VERMELHO_CLARO As Long = RGB(255, 240, 240)

' Variáveis globais
Public proximoPedido As Long
Public DEV_USER As String

' Estrutura para dados da venda
Public Type VendaInfo
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
    quantidadeProdutos As Integer
End Type

Public VendaCorrente As VendaInfo

' ====================================================================
' FUNÇÃO PRINCIPAL - VERSÃO SÓLIDA E ADAPTADA
' ====================================================================
Public Sub ProcessarVendaPDVSolido()
    On Error GoTo TratarErroPrincipal
    
    Dim inicioProcesso As Double
    inicioProcesso = Timer
    
    Debug.Print String(80, "=")
    Debug.Print "INICIANDO SISTEMA PDV SÓLIDO | " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Debug.Print String(80, "=")
    
    ' ETAPA 1: INICIALIZAR SISTEMA
    If Not InicializarSistemaPDVSolido() Then
        Debug.Print "❌ Falha na inicialização do sistema"
        Exit Sub
    End If
    
    ' ETAPA 2: VALIDAR FORMULÁRIO ATIVO
    Dim frmAtivo As Object
    Set frmAtivo = ObterFormularioAtivo()
    If frmAtivo Is Nothing Then
        MsgBox "❌ Nenhum formulário de vendas encontrado!" & vbCrLf & _
               "Abra o formulário de vendas primeiro.", vbCritical, "Sistema PDV"
        Exit Sub
    End If
    
    ' ETAPA 3: CAPTURAR E VALIDAR DADOS
    If Not CapturarDadosFormularioSolido(frmAtivo) Then
        Debug.Print "❌ Falha na captura de dados"
        Exit Sub
    End If
    
    ' ETAPA 4: GERAR NÚMERO DO PEDIDO
    VendaCorrente.numeroPedido = GerarProximoNumeroPedidoSolido()
    Debug.Print "✅ Pedido gerado: #" & VendaCorrente.numeroPedido
    
    ' ETAPA 5: VALIDAR DADOS COMPLETOS
    If Not ValidarDadosVendaSolido() Then
        Call ReverterNumeroPedidoSolido
        Debug.Print "❌ Validação falhou - Número revertido"
        Exit Sub
    End If
    
    ' ETAPA 6: PROCESSAR PLANILHA
    If Not ProcessarPlanilhaTalaoSolido() Then
        Call ReverterNumeroPedidoSolido
        Debug.Print "❌ Falha no processamento da planilha"
        Exit Sub
    End If
    
    ' ETAPA 7: CONFIRMAR E IMPRIMIR
    If ConfirmarImpressaoSolido() Then
        If ExecutarImpressaoSolido() Then
            ' ETAPA 8: FINALIZAR VENDA
            Call FinalizarVendaSolido(frmAtivo)
            
            Dim tempoTotal As Double
            tempoTotal = Timer - inicioProcesso
            
            Debug.Print String(80, "=")
            Debug.Print "✅ VENDA PROCESSADA COM SUCESSO!"
            Debug.Print "Pedido: #" & VendaCorrente.numeroPedido
            Debug.Print "Cliente: " & VendaCorrente.nomeCliente
            Debug.Print "Total: R$ " & Format(VendaCorrente.total, "#,##0.00")
            Debug.Print "Tempo de processamento: " & Format(tempoTotal, "0.00") & "s"
            Debug.Print String(80, "=")
        Else
            Call ReverterNumeroPedidoSolido
            Debug.Print "❌ Impressão falhou - Número revertido"
        End If
    Else
        Call ReverterNumeroPedidoSolido
        Debug.Print "❌ Impressão cancelada - Número revertido"
    End If
    
    Exit Sub
    
TratarErroPrincipal:
    Debug.Print "❌ ERRO CRÍTICO NO SISTEMA PDV: " & Err.Description
    
    MsgBox "❌ ERRO CRÍTICO NO SISTEMA PDV!" & vbCrLf & vbCrLf & _
           "Erro: " & Err.Description & vbCrLf & _
           "Número: " & Err.Number & vbCrLf & _
           "Local: ProcessarVendaPDVSolido" & vbCrLf & vbCrLf & _
           "O sistema será reinicializado." & vbCrLf & _
           "Contate o suporte se o erro persistir.", vbCritical, "Erro Crítico PDV"
    
    ' Tentar reverter número em caso de erro
    Call ReverterNumeroPedidoSolido
    
    ' Reinicializar sistema
    Call InicializarSistemaPDVSolido
End Sub

' ====================================================================
' INICIALIZAÇÃO DO SISTEMA - VERSÃO SÓLIDA
' ====================================================================
Private Function InicializarSistemaPDVSolido() As Boolean
    On Error GoTo TratarErroInicializacao
    
    InicializarSistemaPDVSolido = False
    
    Debug.Print "🔄 Inicializando Sistema PDV Sólido..."
    
    ' Definir usuário do sistema
    DEV_USER = Environ("USERNAME")
    If DEV_USER = "" Then DEV_USER = "USUARIO_SISTEMA"
    
    ' Inicializar contador de pedidos se necessário
    If proximoPedido = 0 Then
        proximoPedido = ObterUltimoNumeroPedido()
    End If
    
    ' Limpar estrutura da venda
    Call LimparEstruturaVenda
    
    ' Verificar se Excel está respondendo
    If Not TestarExcelDisponivel() Then
        MsgBox "❌ Excel não está respondendo adequadamente!" & vbCrLf & _
               "Reinicie o Excel e tente novamente.", vbCritical
        Exit Function
    End If
    
    ' Verificar se planilha do talão existe
    If Not PlanilhaExiste(PLANILHA_TALAO) Then
        MsgBox "❌ Planilha do talão não encontrada!" & vbCrLf & _
               "Planilha necessária: '" & PLANILHA_TALAO & "'" & vbCrLf & _
               "Verifique se o arquivo está correto.", vbCritical
        Exit Function
    End If
    
    Debug.Print "✅ Sistema inicializado com sucesso"
    Debug.Print "👤 Usuário: " & DEV_USER
    Debug.Print "📝 Próximo pedido: #" & Format(proximoPedido + 1, "00000")
    Debug.Print "📊 Planilha: " & PLANILHA_TALAO
    
    InicializarSistemaPDVSolido = True
    Exit Function
    
TratarErroInicializacao:
    Debug.Print "❌ Erro na inicialização: " & Err.Description
    MsgBox "❌ Erro na inicialização do sistema!" & vbCrLf & _
           "Erro: " & Err.Description, vbCritical
    InicializarSistemaPDVSolido = False
End Function

' ====================================================================
' OBTER FORMULÁRIO ATIVO - VERSÃO SÓLIDA
' ====================================================================
Private Function ObterFormularioAtivo() As Object
    On Error GoTo TratarErroFormulario
    
    Set ObterFormularioAtivo = Nothing
    
    ' Verificar se há UserForms carregados
    If UserForms.Count = 0 Then
        Debug.Print "❌ Nenhum formulário carregado"
        Exit Function
    End If
    
    ' Tentar encontrar o formulário de vendas
    Dim i As Integer
    For i = 0 To UserForms.Count - 1
        Dim frm As Object
        Set frm = UserForms(i)
        
        ' Verificar se o formulário tem os controles necessários
        If FormularioValido(frm) Then
            Set ObterFormularioAtivo = frm
            Debug.Print "✅ Formulário válido encontrado: " & frm.Name
            Exit Function
        End If
    Next i
    
    Debug.Print "❌ Nenhum formulário válido encontrado"
    Exit Function
    
TratarErroFormulario:
    Debug.Print "❌ Erro ao obter formulário: " & Err.Description
    Set ObterFormularioAtivo = Nothing
End Function

' ====================================================================
' VALIDAR FORMULÁRIO - VERSÃO SÓLIDA
' ====================================================================
Private Function FormularioValido(frm As Object) As Boolean
    On Error GoTo TratarErroValidacao
    
    FormularioValido = False
    
    ' Lista de controles obrigatórios
    Dim controlesObrigatorios As Variant
    controlesObrigatorios = Array("txtNome", "txtEnder", "cPagamento", "produtosv1")
    
    ' Verificar se todos os controles existem
    Dim i As Integer
    For i = 0 To UBound(controlesObrigatorios)
        If Not ControleExiste(frm, controlesObrigatorios(i)) Then
            Debug.Print "❌ Controle não encontrado: " & controlesObrigatorios(i)
            Exit Function
        End If
    Next i
    
    FormularioValido = True
    Exit Function
    
TratarErroValidacao:
    Debug.Print "❌ Erro na validação do formulário: " & Err.Description
    FormularioValido = False
End Function

' ====================================================================
' CAPTURAR DADOS DO FORMULÁRIO - VERSÃO SÓLIDA
' ====================================================================
Private Function CapturarDadosFormularioSolido(frm As Object) As Boolean
    On Error GoTo TratarErroCaptura
    
    CapturarDadosFormularioSolido = False
    
    Debug.Print "🔄 Capturando dados do formulário..."
    
    With VendaCorrente
        ' Dados básicos do cliente
        .nomeCliente = LimparTexto(ObterValorControle(frm, "txtNome"))
        .endereco = LimparTexto(ObterValorControle(frm, "txtEnder"))
        .numero = LimparTexto(ObterValorControle(frm, "txtnumero"))
        .bairro = ObterValorControle(frm, "cbairro1")
        .cidade = ObterValorControle(frm, "cCidade")
        .uf = "PE"
        .cep = LimparTexto(ObterValorControle(frm, "txtCEP"))
        .cpfCnpj = LimparTexto(ObterValorControle(frm, "txtCPF"))
        
        ' Dados da venda
        .formaPagamento = ObterValorControle(frm, "cPagamento")
        .dataVenda = Date
        
        ' Data de entrega
        Dim dataEntregaTexto As String
        dataEntregaTexto = LimparTexto(ObterValorControle(frm, "cData"))
        If dataEntregaTexto <> "" And IsDate(dataEntregaTexto) Then
            .dataEntrega = CDate(dataEntregaTexto)
        Else
            .dataEntrega = Date + 1
        End If
        
        ' Calcular totais
        .quantidadeProdutos = ObterQuantidadeProdutos(frm)
        .subtotal = CalcularTotalProdutosSolido(frm)
        .desconto = 0
        .frete = 0
        .total = .subtotal - .desconto + .frete
        
        .vendedor = DEV_USER
        .status = "PROCESSANDO"
    End With
    
    Debug.Print "✅ Dados capturados:"
    Debug.Print "   Cliente: " & VendaCorrente.nomeCliente
    Debug.Print "   Produtos: " & VendaCorrente.quantidadeProdutos
    Debug.Print "   Total: R$ " & Format(VendaCorrente.total, "#,##0.00")
    Debug.Print "   Pagamento: " & VendaCorrente.formaPagamento
    
    CapturarDadosFormularioSolido = True
    Exit Function
    
TratarErroCaptura:
    Debug.Print "❌ Erro ao capturar dados: " & Err.Description
    MsgBox "❌ Erro ao capturar dados do formulário!" & vbCrLf & _
           "Erro: " & Err.Description & vbCrLf & _
           "Verifique se todos os campos estão preenchidos corretamente.", vbCritical
    CapturarDadosFormularioSolido = False
End Function

' ====================================================================
' VALIDAR DADOS DA VENDA - VERSÃO SÓLIDA
' ====================================================================
Private Function ValidarDadosVendaSolido() As Boolean
    On Error Resume Next
    
    ValidarDadosVendaSolido = False
    
    Debug.Print "🔄 Validando dados da venda..."
    
    With VendaCorrente
        ' Validar nome do cliente
        If Trim(.nomeCliente) = "" Then
            MsgBox "⚠️ NOME DO CLIENTE É OBRIGATÓRIO!" & vbCrLf & vbCrLf & _
                   "Preencha o nome do cliente antes de processar" & vbCrLf & _
                   "Pedido que seria gerado: #" & .numeroPedido, vbExclamation, "Campo Obrigatório"
            Exit Function
        End If
        
        ' Validar forma de pagamento
        If Trim(.formaPagamento) = "" Then
            MsgBox "⚠️ FORMA DE PAGAMENTO É OBRIGATÓRIA!" & vbCrLf & vbCrLf & _
                   "Selecione uma forma de pagamento" & vbCrLf & _
                   "Pedido: #" & .numeroPedido, vbExclamation, "Campo Obrigatório"
            Exit Function
        End If
        
        ' Validar produtos
        If .quantidadeProdutos = 0 Then
            MsgBox "⚠️ NENHUM PRODUTO ADICIONADO!" & vbCrLf & vbCrLf & _
                   "Adicione pelo menos um produto" & vbCrLf & _
                   "Pedido: #" & .numeroPedido, vbExclamation, "Produtos Obrigatórios"
            Exit Function
        End If
        
        ' Validar total
        If .total <= 0 Then
            MsgBox "⚠️ VALOR TOTAL INVÁLIDO!" & vbCrLf & vbCrLf & _
                   "O valor total deve ser maior que zero" & vbCrLf & _
                   "Total atual: R$ " & Format(.total, "#,##0.00"), vbExclamation, "Valor Inválido"
            Exit Function
        End If
        
        ' Validar data de entrega
        If .dataEntrega < .dataVenda Then
            MsgBox "⚠️ DATA DE ENTREGA INVÁLIDA!" & vbCrLf & vbCrLf & _
                   "A data de entrega não pode ser anterior à venda" & vbCrLf & _
                   "Data da venda: " & Format(.dataVenda, "dd/mm/yyyy") & vbCrLf & _
                   "Data de entrega: " & Format(.dataEntrega, "dd/mm/yyyy"), vbExclamation, "Data Inválida"
            Exit Function
        End If
    End With
    
    Debug.Print "✅ Dados validados com sucesso"
    ValidarDadosVendaSolido = True
End Function

' ====================================================================
' PROCESSAR PLANILHA DO TALÃO - VERSÃO SÓLIDA
' ====================================================================
Private Function ProcessarPlanilhaTalaoSolido() As Boolean
    On Error GoTo TratarErroProcessamento
    
    ProcessarPlanilhaTalaoSolido = False
    
    Debug.Print "🔄 Processando planilha do talão..."
    
    ' Obter referência da planilha
    Dim ws As Worksheet
    Set ws = ObterPlanilhaTalao()
    If ws Is Nothing Then Exit Function
    
    ' Desabilitar atualizações para melhor performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Processar dados na planilha
    Call LimparAreaTalao(ws)
    Call EscreverCabecalhoTalao(ws)
    Call EscreverProdutosTalao(ws)
    Call EscreverRodapeTalao(ws)
    Call FormatarTalaoSolido(ws)
    
    ' Reabilitar atualizações
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    Debug.Print "✅ Planilha processada com sucesso"
    ProcessarPlanilhaTalaoSolido = True
    Exit Function
    
TratarErroProcessamento:
    ' Reabilitar atualizações em caso de erro
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    Debug.Print "❌ Erro no processamento da planilha: " & Err.Description
    MsgBox "❌ Erro ao processar planilha do talão!" & vbCrLf & _
           "Erro: " & Err.Description, vbCritical
    ProcessarPlanilhaTalaoSolido = False
End Function

' ====================================================================
' FUNÇÕES AUXILIARES - VERSÃO SÓLIDA
' ====================================================================

' Limpar estrutura da venda
Private Sub LimparEstruturaVenda()
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
        .status = ""
        .quantidadeProdutos = 0
    End With
End Sub

' Verificar se Excel está disponível
Private Function TestarExcelDisponivel() As Boolean
    On Error GoTo TratarErroTeste
    
    TestarExcelDisponivel = False
    
    ' Testar operações básicas do Excel
    Dim teste As String
    teste = Application.Name
    
    If Application.Workbooks.Count = 0 Then
        Exit Function
    End If
    
    TestarExcelDisponivel = True
    Exit Function
    
TratarErroTeste:
    TestarExcelDisponivel = False
End Function

' Verificar se planilha existe
Private Function PlanilhaExiste(nomePlanilha As String) As Boolean
    On Error GoTo TratarErroPlanilha
    
    PlanilhaExiste = False
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(nomePlanilha)
    
    If Not ws Is Nothing Then
        PlanilhaExiste = True
    End If
    
    Exit Function
    
TratarErroPlanilha:
    PlanilhaExiste = False
End Function

' Verificar se controle existe no formulário
Private Function ControleExiste(frm As Object, nomeControle As String) As Boolean
    On Error GoTo TratarErroControle
    
    ControleExiste = False
    
    Dim ctrl As Object
    Set ctrl = frm.Controls(nomeControle)
    
    If Not ctrl Is Nothing Then
        ControleExiste = True
    End If
    
    Exit Function
    
TratarErroControle:
    ControleExiste = False
End Function

' Obter valor de um controle com segurança
Private Function ObterValorControle(frm As Object, nomeControle As String) As String
    On Error GoTo TratarErroValor
    
    ObterValorControle = ""
    
    If ControleExiste(frm, nomeControle) Then
        Dim ctrl As Object
        Set ctrl = frm.Controls(nomeControle)
        
        ' Diferentes tipos de controles
        If TypeName(ctrl) = "TextBox" Then
            ObterValorControle = ctrl.Text
        ElseIf TypeName(ctrl) = "ComboBox" Then
            ObterValorControle = ctrl.Value
        ElseIf TypeName(ctrl) = "ListBox" Then
            ObterValorControle = ctrl.Value
        Else
            ObterValorControle = ctrl.Value
        End If
    End If
    
    Exit Function
    
TratarErroValor:
    ObterValorControle = ""
End Function

' Limpar texto removendo espaços extras
Private Function LimparTexto(texto As String) As String
    LimparTexto = Trim(Replace(texto, "  ", " "))
End Function

' Obter quantidade de produtos
Private Function ObterQuantidadeProdutos(frm As Object) As Integer
    On Error Resume Next
    
    ObterQuantidadeProdutos = 0
    
    If ControleExiste(frm, "produtosv1") Then
        ObterQuantidadeProdutos = frm.Controls("produtosv1").ListCount
    End If
End Function

' Calcular total dos produtos
Private Function CalcularTotalProdutosSolido(frm As Object) As Double
    On Error Resume Next
    
    CalcularTotalProdutosSolido = 0
    
    If Not ControleExiste(frm, "produtosv1") Then Exit Function
    
    Dim listaProdutos As Object
    Set listaProdutos = frm.Controls("produtosv1")
    
    Dim i As Integer
    Dim total As Double
    total = 0
    
    For i = 0 To listaProdutos.ListCount - 1
        Dim valorTexto As String
        valorTexto = listaProdutos.List(i, 6) ' Coluna do total
        
        ' Limpar formatação monetária
        valorTexto = Replace(valorTexto, "R$", "")
        valorTexto = Replace(valorTexto, " ", "")
        valorTexto = Replace(valorTexto, ".", "")
        valorTexto = Replace(valorTexto, ",", ".")
        
        If IsNumeric(valorTexto) Then
            total = total + CDbl(valorTexto)
        End If
    Next i
    
    CalcularTotalProdutosSolido = total
End Function

' Obter planilha do talão
Private Function ObterPlanilhaTalao() As Worksheet
    On Error GoTo TratarErroPlanilhaTalao
    
    Set ObterPlanilhaTalao = Nothing
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(PLANILHA_TALAO)
    
    If ws Is Nothing Then
        MsgBox "❌ PLANILHA DO TALÃO NÃO ENCONTRADA!" & vbCrLf & _
               "Planilha: '" & PLANILHA_TALAO & "'" & vbCrLf & _
               "Verifique se o nome está correto.", vbCritical
        Exit Function
    End If
    
    Set ObterPlanilhaTalao = ws
    Exit Function
    
TratarErroPlanilhaTalao:
    Set ObterPlanilhaTalao = Nothing
End Function

' Gerar próximo número de pedido
Private Function GerarProximoNumeroPedidoSolido() As String
    On Error Resume Next
    
    proximoPedido = proximoPedido + 1
    GerarProximoNumeroPedidoSolido = Format(proximoPedido, "00000")
    
    Debug.Print "📝 Número gerado: #" & GerarProximoNumeroPedidoSolido
End Function

' Reverter número do pedido
Private Sub ReverterNumeroPedidoSolido()
    On Error Resume Next
    
    If proximoPedido > 0 Then
        proximoPedido = proximoPedido - 1
        Debug.Print "🔄 Número revertido para: #" & Format(proximoPedido, "00000")
    End If
End Sub

' Obter último número de pedido (implementação básica)
Private Function ObterUltimoNumeroPedido() As Long
    On Error Resume Next
    
    ' Implementação básica - pode ser melhorada conforme necessidade
    ObterUltimoNumeroPedido = 1000
    
    ' Aqui você poderia implementar lógica para ler de planilha de histórico
    ' ou arquivo de configuração
End Function

' ====================================================================
' FUNÇÕES DE PROCESSAMENTO DA PLANILHA
' ====================================================================

' Limpar área do talão
Private Sub LimparAreaTalao(ws As Worksheet)
    On Error Resume Next
    
    ' Limpar áreas específicas preservando o template
    ws.Range("B7:H9,M7:T9").ClearContents         ' Cabeçalhos
    ws.Range("B11:H21,M11:T21").ClearContents     ' Produtos
    ws.Range("B22:H25,M22:T25").ClearContents     ' Rodapés
End Sub

' Escrever cabeçalho do talão
Private Sub EscreverCabecalhoTalao(ws As Worksheet)
    On Error Resume Next
    
    With VendaCorrente
        ' Número do pedido (destaque)
        ws.Range("B6").Value = "PEDIDO #" & .numeroPedido
        ws.Range("M6").Value = "PEDIDO #" & .numeroPedido
        
        ' Dados do cliente - lado esquerdo
        ws.Range("B7").Value = .nomeCliente
        ws.Range("B8").Value = .endereco & IIf(.numero <> "", ", " & .numero, "")
        ws.Range("F8").Value = .bairro
        ws.Range("B9").Value = .cpfCnpj
        ws.Range("E9").Value = .cidade
        ws.Range("G9").Value = .uf
        ws.Range("H9").Value = .cep
        
        ' Dados do cliente - lado direito (espelho)
        ws.Range("M7").Value = .nomeCliente
        ws.Range("M8").Value = .endereco & IIf(.numero <> "", ", " & .numero, "")
        ws.Range("Q8").Value = .bairro
        ws.Range("M9").Value = .cpfCnpj
        ws.Range("P9").Value = .cidade
        ws.Range("R9").Value = .uf
        ws.Range("T9").Value = .cep
    End With
End Sub

' Escrever produtos do talão
Private Sub EscreverProdutosTalao(ws As Worksheet)
    On Error Resume Next
    
    ' Obter formulário para acessar produtos
    Dim frm As Object
    Set frm = ObterFormularioAtivo()
    If frm Is Nothing Then Exit Sub
    
    If Not ControleExiste(frm, "produtosv1") Then Exit Sub
    
    Dim listaProdutos As Object
    Set listaProdutos = frm.Controls("produtosv1")
    
    Dim i As Integer
    Dim linhaAtual As Integer
    Dim totalProdutos As Integer
    
    totalProdutos = IIf(listaProdutos.ListCount > 10, 10, listaProdutos.ListCount)
    
    For i = 0 To totalProdutos - 1
        linhaAtual = 11 + i
        
        ' LADO ESQUERDO (Via da loja)
        ws.Range("B" & linhaAtual).Value = listaProdutos.List(i, 0)  ' Referência
        ws.Range("C" & linhaAtual).Value = listaProdutos.List(i, 1)  ' Descrição
        ws.Range("D" & linhaAtual).Value = listaProdutos.List(i, 2)  ' Unidade
        ws.Range("E" & linhaAtual).Value = listaProdutos.List(i, 3)  ' Valor Unit
        ws.Range("F" & linhaAtual).Value = listaProdutos.List(i, 4)  ' Quantidade
        ws.Range("G" & linhaAtual).Value = listaProdutos.List(i, 5)  ' Desconto
        ws.Range("H" & linhaAtual).Value = listaProdutos.List(i, 6)  ' Total
        
        ' LADO DIREITO (Via do cliente)
        ws.Range("M" & linhaAtual).Value = listaProdutos.List(i, 0)  ' Referência
        ws.Range("N" & linhaAtual).Value = listaProdutos.List(i, 1)  ' Descrição
        ws.Range("O" & linhaAtual).Value = listaProdutos.List(i, 2)  ' Unidade
        ws.Range("P" & linhaAtual).Value = listaProdutos.List(i, 3)  ' Valor Unit
        ws.Range("Q" & linhaAtual).Value = listaProdutos.List(i, 4)  ' Quantidade
        ws.Range("R" & linhaAtual).Value = listaProdutos.List(i, 5)  ' Desconto
        ws.Range("S" & linhaAtual).Value = listaProdutos.List(i, 6)  ' Total
    Next i
End Sub

' Escrever rodapé do talão
Private Sub EscreverRodapeTalao(ws As Worksheet)
    On Error Resume Next
    
    With VendaCorrente
        ' LADO ESQUERDO - Informações da venda
        ws.Range("B22").Value = .vendedor                           ' Vendedor
        ws.Range("B24").Value = "BALCÃO"                           ' Situação
        ws.Range("B25").Value = .formaPagamento                    ' Forma de pagamento
        ws.Range("C25").Value = "PEDIDO #" & .numeroPedido         ' Número do pedido
        ws.Range("F23").Value = Format(.dataEntrega, "dd/mm/yyyy") ' Data de entrega
        ws.Range("H22").Value = .subtotal                          ' Total produtos
        ws.Range("H23").Value = .frete                             ' Frete
        ws.Range("H24").Value = .desconto                          ' Desconto
        ws.Range("H25").Value = .total                             ' Total geral
        
        ' LADO DIREITO - Espelho
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
End Sub

' Formatar talão
Private Sub FormatarTalaoSolido(ws As Worksheet)
    On Error Resume Next
    
    ' Formatação do número do pedido (destaque)
    With ws.Range("B6,M6")
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = RGB(0, 120, 0)
        .HorizontalAlignment = xlCenter
        .Interior.Color = COR_VERDE_CLARO
        .Borders.LineStyle = xlContinuous
    End With
    
    ' Formatação geral do conteúdo
    With ws.Range("B7:T25")
        .WrapText = False
        .ShrinkToFit = False
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Size = 10
    End With
    
    ' Ajustar largura das colunas
    ws.Columns("B:B").ColumnWidth = 15    ' Referência
    ws.Columns("C:C").ColumnWidth = 30    ' Descrição
    ws.Columns("D:H").ColumnWidth = 12    ' Dados numéricos
    ws.Columns("M:M").ColumnWidth = 15    ' Referência direita
    ws.Columns("N:N").ColumnWidth = 30    ' Descrição direita
    ws.Columns("O:S").ColumnWidth = 12    ' Dados numéricos direita
    
    ' Formatação especial para valores monetários
    ws.Range("H22,H25,S22,S25").NumberFormat = "_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * ""-""_-;_-@_-"
End Sub

' ====================================================================
' FUNÇÕES DE IMPRESSÃO - VERSÃO SÓLIDA
' ====================================================================

' Confirmar impressão
Private Function ConfirmarImpressaoSolido() As Boolean
    With VendaCorrente
        Dim mensagem As String
        mensagem = "🖨️ TALÃO PREPARADO PARA IMPRESSÃO!" & vbCrLf & vbCrLf & _
                  "PEDIDO: #" & .numeroPedido & vbCrLf & _
                  "Cliente: " & .nomeCliente & vbCrLf & _
                  "Endereço: " & .endereco & IIf(.numero <> "", ", " & .numero, "") & vbCrLf & _
                  "Bairro: " & .bairro & " - " & .cidade & "/" & .uf & vbCrLf & _
                  "Pagamento: " & .formaPagamento & vbCrLf & _
                  "Total: R$ " & Format(.total, "#,##0.00") & vbCrLf & _
                  "Produtos: " & .quantidadeProdutos & " itens" & vbCrLf & _
                  "Data/Hora: " & Format(Now, "dd/mm/yyyy hh:mm:ss") & vbCrLf & vbCrLf & _
                  "Confirma a impressão do talão?"
        
        ConfirmarImpressaoSolido = (MsgBox(mensagem, vbYesNo + vbQuestion, "Confirmar Impressão") = vbYes)
        
        If ConfirmarImpressaoSolido Then
            Debug.Print "✅ Usuário confirmou impressão do pedido #" & .numeroPedido
        Else
            Debug.Print "❌ Usuário cancelou impressão do pedido #" & .numeroPedido
        End If
    End With
End Function

' Executar impressão
Private Function ExecutarImpressaoSolido() As Boolean
    On Error GoTo TratarErroImpressao
    
    ExecutarImpressaoSolido = False
    
    Debug.Print "🖨️ Executando impressão..."
    
    Dim ws As Worksheet
    Set ws = ObterPlanilhaTalao()
    If ws Is Nothing Then Exit Function
    
    ' Configurar impressão
    Call ConfigurarImpressaoSolido(ws)
    
    ' Executar impressão
    ws.PrintOut
    
    ' Exibir sucesso
    Call ExibirSucessoImpressaoSolido()
    
    Debug.Print "✅ Impressão executada com sucesso"
    ExecutarImpressaoSolido = True
    Exit Function
    
TratarErroImpressao:
    Debug.Print "❌ Erro na impressão: " & Err.Description
    MsgBox "❌ Erro na impressão!" & vbCrLf & _
           "Erro: " & Err.Description & vbCrLf & _
           "Verifique se a impressora está conectada e funcionando.", vbCritical
    ExecutarImpressaoSolido = False
End Function

' Configurar impressão
Private Sub ConfigurarImpressaoSolido(ws As Worksheet)
    On Error Resume Next
    
    With ws.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
        .PrintArea = "A1:T26"
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
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
    
    ' Ativar planilha e ajustar zoom
    ws.Activate
    If Not Application.ActiveWindow Is Nothing Then
        Application.ActiveWindow.Zoom = 85
    End If
End Sub

' Exibir sucesso da impressão
Private Sub ExibirSucessoImpressaoSolido()
    With VendaCorrente
        MsgBox "✅ PEDIDO #" & .numeroPedido & " IMPRESSO COM SUCESSO!" & vbCrLf & vbCrLf & _
               "Talão impresso em 2 vias (Cliente + Loja)" & vbCrLf & _
               "Cliente: " & .nomeCliente & vbCrLf & _
               "Valor total: R$ " & Format(.total, "#,##0.00") & vbCrLf & _
               "Próximo pedido será: #" & Format(proximoPedido + 1, "00000") & vbCrLf & _
               "Total de impressões hoje: " & proximoPedido & vbCrLf & _
               "Horário: " & Format(Now, "dd/mm/yyyy hh:mm:ss") & vbCrLf & _
               "Sistema funcionando perfeitamente!" & vbCrLf & vbCrLf & _
               VERSAO_SISTEMA & " - " & DEV_USER, vbInformation, "Impressão Concluída"
    End With
End Sub

' ====================================================================
' FINALIZAÇÃO DA VENDA - VERSÃO SÓLIDA
' ====================================================================

' Finalizar venda
Private Sub FinalizarVendaSolido(frm As Object)
    On Error Resume Next
    
    Debug.Print "🔄 Finalizando venda..."
    
    ' Salvar dados da venda
    Call SalvarHistoricoVendaSolido()
    
    ' Limpar formulário
    Call LimparFormularioSolido(frm)
    
    ' Atualizar status
    VendaCorrente.status = "CONCLUÍDA"
    
    Debug.Print "✅ Venda finalizada com sucesso"
End Sub

' Salvar histórico da venda
Private Sub SalvarHistoricoVendaSolido()
    On Error Resume Next
    
    ' Log detalhado da venda
    Debug.Print "💾 VENDA SALVA | " & _
               "Pedido: #" & VendaCorrente.numeroPedido & " | " & _
               "Cliente: " & VendaCorrente.nomeCliente & " | " & _
               "Total: R$ " & Format(VendaCorrente.total, "0.00") & " | " & _
               "Pagamento: " & VendaCorrente.formaPagamento & " | " & _
               "Produtos: " & VendaCorrente.quantidadeProdutos & " | " & _
               Format(Now, "dd/mm/yyyy hh:mm:ss") & " | " & _
               DEV_USER
    
    ' Aqui você pode implementar salvamento em planilha de histórico
    ' Call SalvarEmPlanilhaHistorico()
End Sub

' Limpar formulário
Private Sub LimparFormularioSolido(frm As Object)
    On Error Resume Next
    
    Debug.Print "🧹 Limpando formulário para nova venda..."
    
    ' Limpar campos de texto
    If ControleExiste(frm, "txtNome") Then frm.Controls("txtNome").Value = ""
    If ControleExiste(frm, "txtEnder") Then frm.Controls("txtEnder").Value = ""
    If ControleExiste(frm, "txtnumero") Then frm.Controls("txtnumero").Value = ""
    If ControleExiste(frm, "txtCEP") Then
        frm.Controls("txtCEP").Value = ""
        frm.Controls("txtCEP").BackColor = COR_BRANCO
    End If
    If ControleExiste(frm, "txtCPF") Then frm.Controls("txtCPF").Value = ""
    
    ' Resetar combos
    If ControleExiste(frm, "cPagamento") Then frm.Controls("cPagamento").ListIndex = 0
    If ControleExiste(frm, "cCidade") Then frm.Controls("cCidade").ListIndex = -1
    If ControleExiste(frm, "cbairro1") Then frm.Controls("cbairro1").ListIndex = -1
    
    ' Limpar lista de produtos
    If ControleExiste(frm, "produtosv1") Then frm.Controls("produtosv1").Clear
    
    ' Retornar foco para o primeiro campo
    If ControleExiste(frm, "txtNome") Then frm.Controls("txtNome").SetFocus
    
    Debug.Print "✅ Formulário limpo, pronto para nova venda"
End Sub

' ====================================================================
' FUNÇÃO PARA BOTÃO - INTERFACE SIMPLIFICADA
' ====================================================================
Public Sub btnProcessarVendaSolido_Click()
    ' Esta função pode ser chamada diretamente de um botão no formulário
    Call ProcessarVendaPDVSolido
End Sub

' ====================================================================
' FUNÇÃO DE COMPATIBILIDADE COM CÓDIGO ORIGINAL
' ====================================================================
Public Sub btnImprimirUltraSimples_Click()
    ' Manter compatibilidade com o código original
    Call ProcessarVendaPDVSolido
End Sub