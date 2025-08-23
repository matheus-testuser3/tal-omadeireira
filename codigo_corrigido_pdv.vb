' === CÓDIGO CORRIGIDO - INTEGRAÇÃO lblTotal COM lstSelecionados ===

Private Const QTD_MINIMA As Double = 0.01
Private Const QTD_MAXIMA As Double = 999999.99

' === FUNÇÃO PRINCIPAL CORRIGIDA - TextoParaMoeda ===
Private Function TextoParaMoeda(texto As String) As Currency
    On Error GoTo TratarErro
    
    Dim textoLimpo As String
    textoLimpo = Trim(texto)
    
    ' Remover símbolos monetários e espaços
    textoLimpo = Replace(textoLimpo, "R$", "")
    textoLimpo = Replace(textoLimpo, " ", "")
    
    ' Se estiver vazio, retornar zero
    If textoLimpo = "" Then
        TextoParaMoeda = 0
        Exit Function
    End If
    
    ' Converter vírgula para ponto (padrão brasileiro para inglês)
    ' Mas apenas se não houver ponto já presente
    If InStr(textoLimpo, ".") > 0 And InStr(textoLimpo, ",") > 0 Then
        ' Formato brasileiro: 1.234,56 -> 1234.56
        textoLimpo = Replace(textoLimpo, ".", "")
        textoLimpo = Replace(textoLimpo, ",", ".")
    ElseIf InStr(textoLimpo, ",") > 0 And InStr(textoLimpo, ".") = 0 Then
        ' Apenas vírgula: 1234,56 -> 1234.56
        textoLimpo = Replace(textoLimpo, ",", ".")
    End If
    
    ' Converter para Currency
    If IsNumeric(textoLimpo) Then
        TextoParaMoeda = CCur(textoLimpo)
    Else
        TextoParaMoeda = 0
    End If
    
    Exit Function
    
TratarErro:
    TextoParaMoeda = 0
End Function

' === FUNÇÃO AtualizarResumo COMPLETAMENTE REFEITA ===
Private Sub AtualizarResumo()
    On Error GoTo TratarErro
    
    Dim subtotal As Currency
    Dim totalDescontos As Currency
    Dim totalFinal As Currency
    Dim i As Long
    
    ' Resetar totais
    subtotal = 0
    totalDescontos = 0
    totalFinal = 0
    
    ' Verificar se há itens na lista
    If Me.lstSelecionados.ListCount = 0 Then
        GoTo AtualizarLabels
    End If
    
    ' Percorrer cada item da lista
    For i = 0 To Me.lstSelecionados.ListCount - 1
        Dim precoUnitario As Currency
        Dim quantidade As Double
        Dim desconto As Currency
        Dim totalItem As Currency
        
        ' 1. Obter preço unitário (coluna 3)
        precoUnitario = TextoParaMoeda(CStr(Me.lstSelecionados.List(i, 3)))
        
        ' 2. Obter quantidade (coluna 4) - formato "1.000 UN"
        Dim qtdTexto As String
        qtdTexto = CStr(Me.lstSelecionados.List(i, 4))
        
        ' Extrair apenas o número da quantidade
        If InStr(qtdTexto, " ") > 0 Then
            qtdTexto = Trim(Left(qtdTexto, InStr(qtdTexto, " ") - 1))
        End If
        
        ' Converter quantidade para número
        If IsNumeric(Replace(qtdTexto, ",", ".")) Then
            quantidade = CDbl(Replace(qtdTexto, ",", "."))
        Else
            quantidade = 1
        End If
        
        ' 3. Obter desconto (coluna 5)
        desconto = TextoParaMoeda(CStr(Me.lstSelecionados.List(i, 5)))
        
        ' 4. Calcular subtotal do item
        Dim subtotalItem As Currency
        subtotalItem = precoUnitario * quantidade
        subtotal = subtotal + subtotalItem
        
        ' 5. Somar descontos
        totalDescontos = totalDescontos + desconto
        
        ' 6. Calcular total do item (com desconto) e verificar se está correto na lista
        totalItem = subtotalItem - desconto
        If totalItem < 0 Then totalItem = 0
        
        ' 7. CORREÇÃO CRÍTICA: Atualizar o valor total na lista se estiver incorreto
        Dim valorAtualNaLista As Currency
        valorAtualNaLista = TextoParaMoeda(CStr(Me.lstSelecionados.List(i, 6)))
        
        If Abs(valorAtualNaLista - totalItem) > 0.01 Then
            Me.lstSelecionados.List(i, 6) = Format(totalItem, "R$ #,##0.00")
        End If
        
        ' 8. Somar ao total final
        totalFinal = totalFinal + totalItem
    Next i
    
AtualizarLabels:
    ' Atualizar os labels com proteção contra erros
    On Error Resume Next
    
    ' Label de subtotal (se existir)
    If Not Me.lblsubTotal Is Nothing Then
        Me.lblsubTotal.Caption = "Subtotal: " & Format(subtotal, "R$ #,##0.00")
    End If
    
    ' Label de total de descontos (se existir)
    If Not Me.lblTotalDescontos Is Nothing Then
        Me.lblTotalDescontos.Caption = "Descontos: " & Format(totalDescontos, "R$ #,##0.00")
    End If
    
    ' Label principal de total - CORREÇÃO CRÍTICA
    If Not Me.lblTotal Is Nothing Then
        Me.lblTotal.Caption = "Total: " & Format(totalFinal, "R$ #,##0.00")
        ' Forçar atualização visual
        Me.lblTotal.Refresh
    End If
    
    ' Label de total de itens (se existir)
    If Not Me.lblTotalItens Is Nothing Then
        Me.lblTotalItens.Caption = "Itens: " & Me.lstSelecionados.ListCount
    End If
    
    On Error GoTo 0
    Exit Sub
    
TratarErro:
    Debug.Print "Erro em AtualizarResumo: " & Err.Description & " - Linha: " & Erl
    
    ' Em caso de erro, definir valores seguros
    On Error Resume Next
    If Not Me.lblTotal Is Nothing Then
        Me.lblTotal.Caption = "Total: R$ 0,00"
    End If
    If Not Me.lblTotalItens Is Nothing Then
        Me.lblTotalItens.Caption = "Itens: 0"
    End If
    On Error GoTo 0
End Sub

' === EVENTOS CORRIGIDOS PARA GARANTIR SINCRONIZAÇÃO ===

Private Sub lstSelecionados_Click()
    ' Atualizar resumo sempre que a seleção mudar
    Call AtualizarResumo
End Sub

Private Sub lstSelecionados_Change()
    ' Evento principal para mudanças na lista
    Call AtualizarResumo
End Sub

Private Sub lstSelecionados_AfterUpdate()
    ' Garantir atualização após qualquer modificação
    Call AtualizarResumo
End Sub

' === CORREÇÃO NO btnAdicionar_Click ===
Private Sub btnAdicionar_Click()
    If Me.lstProdutos.ListIndex < 0 Then
        MsgBox "Selecione um produto.", vbExclamation
        Exit Sub
    End If

    Dim i As Long, idProduto As String
    idProduto = Me.lstProdutos.List(Me.lstProdutos.ListIndex, 0)

    ' Verificar se produto já foi adicionado
    For i = 0 To Me.lstSelecionados.ListCount - 1
        If CStr(Me.lstSelecionados.List(i, 0)) = idProduto Then
            MsgBox "Produto já adicionado.", vbInformation
            Exit Sub
        End If
    Next i

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Produtos")

    Dim precoUnitario As Currency
    Dim unidadeProduto As String
    Dim linhaAtual As Long

    ' Buscar preço e unidade corretos da planilha
    For linhaAtual = 2 To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        If CStr(ws.Cells(linhaAtual, 1).Value) = idProduto Then
            precoUnitario = ws.Cells(linhaAtual, 6).Value ' Preço de Venda da Coluna F
            unidadeProduto = CStr(ws.Cells(linhaAtual, 3).Value) ' Unidade da Coluna C
            Exit For
        End If
    Next linhaAtual

    ' Se não encontrou a unidade na planilha, usar a da lista como fallback
    If unidadeProduto = "" Then
        unidadeProduto = Me.lstProdutos.List(Me.lstProdutos.ListIndex, 2)
    End If

    Dim qtd As Double
    qtd = CDbl(Me.txtQuantidade.Value)

    Dim totalItem As Currency
    totalItem = precoUnitario * qtd

    ' Adicionar item com formatação correta
    With Me.lstSelecionados
        .AddItem
        .List(.ListCount - 1, 0) = idProduto
        .List(.ListCount - 1, 1) = Me.lstProdutos.List(Me.lstProdutos.ListIndex, 1)
        .List(.ListCount - 1, 2) = unidadeProduto
        .List(.ListCount - 1, 3) = Format(precoUnitario, "R$ #,##0.00")
        .List(.ListCount - 1, 4) = Format(qtd, "0.000") & " " & unidadeProduto
        .List(.ListCount - 1, 5) = "R$ 0,00"  ' Desconto inicial = zero
        .List(.ListCount - 1, 6) = Format(totalItem, "R$ #,##0.00")
    End With

    ' CORREÇÃO CRÍTICA: Forçar atualização imediata
    Call AtualizarResumo
    Me.txtQuantidade.Value = 1
    
    ' Forçar refresh visual
    Me.Repaint
End Sub

' === CORREÇÃO NO btnRemover_Click ===
Private Sub btnRemover_Click()
    If Me.lstSelecionados.ListIndex >= 0 Then
        Me.lstSelecionados.RemoveItem Me.lstSelecionados.ListIndex
        ' CORREÇÃO CRÍTICA: Atualizar resumo imediatamente após remoção
        Call AtualizarResumo
        Me.Repaint
    Else
        MsgBox "Selecione um item para remover.", vbExclamation
    End If
End Sub

' === CORREÇÃO NO AplicarDescontoItem ===
Private Sub AplicarDescontoItem(indiceItem As Long, valorDesconto As Double, tipoDesconto As String)
    Dim precoUnitario As Currency
    Dim quantidade As Double
    Dim descontoAplicado As Currency
    Dim novoTotal As Currency

    precoUnitario = TextoParaMoeda(CStr(Me.lstSelecionados.List(indiceItem, 3)))
    
    ' Obter quantidade corretamente
    Dim qtdTexto As String
    qtdTexto = CStr(Me.lstSelecionados.List(indiceItem, 4))
    If InStr(qtdTexto, " ") > 0 Then
        qtdTexto = Trim(Left(qtdTexto, InStr(qtdTexto, " ") - 1))
    End If
    quantidade = CDbl(Replace(qtdTexto, ",", "."))

    If tipoDesconto = "percentual" Then
        descontoAplicado = (precoUnitario * quantidade) * (valorDesconto / 100)
    Else
        descontoAplicado = valorDesconto
    End If

    novoTotal = (precoUnitario * quantidade) - descontoAplicado

    If novoTotal < 0 Then
        novoTotal = 0
        descontoAplicado = precoUnitario * quantidade
    End If

    ' Atualizar os valores na lista
    Me.lstSelecionados.List(indiceItem, 5) = Format(descontoAplicado, "R$ #,##0.00")
    Me.lstSelecionados.List(indiceItem, 6) = Format(novoTotal, "R$ #,##0.00")

    ' CORREÇÃO CRÍTICA: Forçar atualização imediata
    Call AtualizarResumo
    Me.Repaint

    If tipoDesconto = "percentual" Then
        MsgBox "Desconto de " & valorDesconto & "% aplicado!" & vbCrLf & _
               "Desconto: " & Format(descontoAplicado, "R$ #,##0.00"), vbInformation
    Else
        MsgBox "Desconto de " & Format(valorDesconto, "R$ #,##0.00") & " aplicado!", vbInformation
    End If
End Sub

' === CORREÇÃO NO btnRemoverDesconto_Click ===
Private Sub btnRemoverDesconto_Click()
    If Me.lstSelecionados.ListIndex < 0 Then
        MsgBox "Selecione um produto para remover desconto.", vbExclamation
        Exit Sub
    End If

    Dim indiceItem As Long
    indiceItem = Me.lstSelecionados.ListIndex

    Dim precoUnitario As Currency
    Dim quantidade As Double
    Dim totalSemDesconto As Currency

    precoUnitario = TextoParaMoeda(CStr(Me.lstSelecionados.List(indiceItem, 3)))
    
    ' Obter quantidade corretamente
    Dim qtdTexto As String
    qtdTexto = CStr(Me.lstSelecionados.List(indiceItem, 4))
    If InStr(qtdTexto, " ") > 0 Then
        qtdTexto = Trim(Left(qtdTexto, InStr(qtdTexto, " ") - 1))
    End If
    quantidade = CDbl(Replace(qtdTexto, ",", "."))
    
    totalSemDesconto = precoUnitario * quantidade

    ' Remover desconto e recalcular
    Me.lstSelecionados.List(indiceItem, 5) = "R$ 0,00"
    Me.lstSelecionados.List(indiceItem, 6) = Format(totalSemDesconto, "R$ #,##0.00")

    ' CORREÇÃO CRÍTICA: Forçar atualização imediata
    Call AtualizarResumo
    Me.Repaint
    
    MsgBox "Desconto removido com sucesso!", vbInformation
End Sub

' === RESTO DO CÓDIGO MANTIDO (btnAplicarDesconto_Click, btnEnviarParaProdutos_Click, etc.) ===

Private Sub btnAplicarDesconto_Click()
    If Me.lstSelecionados.ListIndex < 0 Then
        MsgBox "Selecione um produto para aplicar desconto.", vbExclamation
        Exit Sub
    End If

    Dim valorDesconto As Double
    Dim tipoDesconto As String

    If Me.optDescontoPercentual.Value = True Then
        tipoDesconto = "percentual"
        valorDesconto = CDbl(Me.txtDesconto.Value)
        
        If valorDesconto < 0 Or valorDesconto > 100 Then
            MsgBox "Desconto percentual deve estar entre 0% e 100%.", vbExclamation
            Exit Sub
        End If
    Else
        tipoDesconto = "valor"
        valorDesconto = TextoParaMoeda(Me.txtDesconto.Value)
        
        If valorDesconto < 0 Then
            MsgBox "Valor de desconto não pode ser negativo.", vbExclamation
            Exit Sub
        End If
    End If

    Call AplicarDescontoItem(Me.lstSelecionados.ListIndex, valorDesconto, tipoDesconto)
    Me.txtDesconto.Value = ""
End Sub

Private Sub btnEnviarParaProdutos_Click()
    On Error GoTo TratarErro

    If Me.lstSelecionados.ListCount = 0 Then
        MsgBox "Nenhum produto selecionado!", vbExclamation
        Exit Sub
    End If

    ' Forçar atualização final antes do envio
    Call AtualizarResumo

    ' Calcular totais
    Dim TotalItensLocal As Long
    Dim ValorTotalLocal As Currency
    Dim i As Long

    TotalItensLocal = Me.lstSelecionados.ListCount
    ValorTotalLocal = 0

    ' Calcular valor total usando a mesma lógica do AtualizarResumo
    For i = 0 To Me.lstSelecionados.ListCount - 1
        ValorTotalLocal = ValorTotalLocal + TextoParaMoeda(CStr(Me.lstSelecionados.List(i, 6)))
    Next i

    ' Obter data/hora atual e usuário
    Dim dataHoraAtual As String
    Dim usuarioAtual As String
    dataHoraAtual = Format(Now, "yyyy-mm-dd hh:mm:ss")
    usuarioAtual = "Matheus-TestUser1"

    ' Confirmar envio
    Dim resposta As VbMsgBoxResult
    resposta = MsgBox("Enviar " & TotalItensLocal & " produto(s) para frmPDVPrincipal?" & vbCrLf & vbCrLf & _
                     "Valor Total: " & Format(ValorTotalLocal, "R$ #,##0.00") & vbCrLf & vbCrLf & _
                     "Destino: frmPDVPrincipal.produtosv2" & vbCrLf & _
                     "Usuário: " & usuarioAtual & vbCrLf & _
                     "Data/Hora: " & dataHoraAtual, vbYesNo + vbQuestion, "Confirmar Envio")

    If resposta = vbYes Then
        ' Tentar conectar com frmPDVPrincipal
        Dim formPrincipal As Object
        Dim produtosv2 As Object
        Dim conectouSucesso As Boolean
        conectouSucesso = False
        
        ' Primeira tentativa: buscar na coleção UserForms
        Dim frm As Object
        On Error Resume Next
        For Each frm In VBA.UserForms
            If TypeName(frm) = "frmPDVPrincipal" Then
                Set formPrincipal = frm
                Exit For
            End If
        Next frm
        
        ' Segunda tentativa: acesso direto se não encontrou
        If formPrincipal Is Nothing Then
            Set formPrincipal = frmPDVPrincipal
        End If
        
        ' Tentar acessar o controle produtosv2
        If Not formPrincipal Is Nothing Then
            Set produtosv2 = formPrincipal.Controls("produtosv2")
            If produtosv2 Is Nothing Then
                ' Fallback para produtosv1 se produtosv2 não existir
                Set produtosv2 = formPrincipal.Controls("produtosv1")
            End If
            
            If Not produtosv2 Is Nothing Then
                conectouSucesso = True
            End If
        End If
        On Error GoTo TratarErro
        
        If conectouSucesso Then
            ' SUCESSO - Configurar e enviar dados
            Dim enviados As Long
            enviados = 0
            
            ' Configurar a listbox de destino
            On Error Resume Next
            With produtosv2
                .ColumnCount = 7
                .ColumnWidths = "60;140;30;60;40;80;60"
            End With
            On Error GoTo TratarErro
            
            ' Enviar cada item individualmente com verificação
            For i = 0 To Me.lstSelecionados.ListCount - 1
                On Error Resume Next
                
                ' Adicionar nova linha
                produtosv2.AddItem
                
                ' Verificar se a linha foi criada e popular dados
                If produtosv2.ListCount > 0 Then
                    Dim ultimaLinha As Long
                    ultimaLinha = produtosv2.ListCount - 1
                    
                    ' Popular dados coluna por coluna
                    produtosv2.List(ultimaLinha, 0) = CStr(Me.lstSelecionados.List(i, 0)) ' Referencia
                    produtosv2.List(ultimaLinha, 1) = CStr(Me.lstSelecionados.List(i, 1)) ' Descrição
                    produtosv2.List(ultimaLinha, 2) = CStr(Me.lstSelecionados.List(i, 2)) ' Uni
                    produtosv2.List(ultimaLinha, 3) = CStr(Me.lstSelecionados.List(i, 3)) ' Valor
                    produtosv2.List(ultimaLinha, 4) = CStr(Me.lstSelecionados.List(i, 4)) ' Quant.
                    produtosv2.List(ultimaLinha, 5) = CStr(Me.lstSelecionados.List(i, 5)) ' Desc.
                    produtosv2.List(ultimaLinha, 6) = CStr(Me.lstSelecionados.List(i, 6)) ' Valor Total
                    
                    enviados = enviados + 1
                End If
                
                On Error GoTo TratarErro
            Next i
            
            ' SUCESSO!
            Dim nomeDestino As String
            nomeDestino = IIf(produtosv2.Name = "produtosv2", "produtosv2", "produtosv1")
            
            MsgBox "PRODUTOS ENVIADOS COM SUCESSO!" & vbCrLf & vbCrLf & _
                   "Conectado com: frmPDVPrincipal" & vbCrLf & _
                   "Produtos enviados: " & enviados & " de " & TotalItensLocal & vbCrLf & _
                   "Valor total: " & Format(ValorTotalLocal, "R$ #,##0.00") & vbCrLf & _
                   "ListBox " & nomeDestino & " atualizada" & vbCrLf & vbCrLf & _
                   usuarioAtual & " | " & dataHoraAtual, vbInformation, "Envio Concluído!"
            
            Unload Me
        Else
            MsgBox "ERRO: Não foi possível conectar com frmPDVPrincipal" & vbCrLf & vbCrLf & _
                   "Verifique se o formulário principal está aberto.", vbExclamation, "Erro na Conexão"
        End If
    End If

    Exit Sub

TratarErro:
    MsgBox "ERRO: " & Err.Description & vbCrLf & vbCrLf & _
           usuarioAtual & " | " & Format(Now, "yyyy-mm-dd hh:mm:ss"), vbExclamation, "Erro no Envio"
End Sub

' === EVENTOS DE QUANTIDADE COM ATUALIZAÇÃO ===
Private Sub btnMais_Click()
    Dim valorAtual As Double
    Dim unidadeAtual As String
    
    valorAtual = CDbl(Me.txtQuantidade.Value)
    
    ' Verificar unidade do produto selecionado para incremento inteligente
    If Me.lstProdutos.ListIndex >= 0 Then
        unidadeAtual = UCase(Trim(CStr(Me.lstProdutos.List(Me.lstProdutos.ListIndex, 2))))
    Else
        unidadeAtual = "UN"
    End If
    
    ' Incremento baseado na unidade específica da madeireira
    Dim incremento As Double
    Select Case unidadeAtual
        Case "CM"
            If valorAtual < 10 Then
                incremento = 1      ' 1 cm para valores pequenos
            ElseIf valorAtual < 100 Then
                incremento = 5      ' 5 cm para valores médios
            Else
                incremento = 10     ' 10 cm para valores grandes
            End If
            
        Case "M"
            If valorAtual < 1 Then
                incremento = 0.1    ' 10 cm
            ElseIf valorAtual < 10 Then
                incremento = 0.5    ' 50 cm
            Else
                incremento = 1      ' 1 metro
            End If
            
        Case "M²", "M2"
            If valorAtual < 1 Then
                incremento = 0.25   ' 0,25 m²
            ElseIf valorAtual < 10 Then
                incremento = 0.5    ' 0,5 m²
            Else
                incremento = 1      ' 1 m²
            End If
            
        Case "M³", "M3"
            If valorAtual < 1 Then
                incremento = 0.1    ' 0,1 m³
            Else
                incremento = 0.5    ' 0,5 m³
            End If
            
        Case "PÇ", "PC", "UN", "PEÇA"
            incremento = 1          ' 1 peça
            
        Case "KG"
            If valorAtual < 1 Then
                incremento = 0.1    ' 100g
            ElseIf valorAtual < 10 Then
                incremento = 0.5    ' 500g
            Else
                incremento = 1      ' 1 kg
            End If
            
        Case "L", "LT"
            If valorAtual < 1 Then
                incremento = 0.1    ' 100ml
            ElseIf valorAtual < 10 Then
                incremento = 0.5    ' 500ml
            Else
                incremento = 1      ' 1 litro
            End If
            
        Case Else ' Outras unidades
            If valorAtual < 1 Then
                incremento = 0.1
            ElseIf valorAtual < 10 Then
                incremento = 0.5
            Else
                incremento = 1
            End If
    End Select
    
    If (valorAtual + incremento) <= QTD_MAXIMA Then
        Me.txtQuantidade.Value = Round(valorAtual + incremento, 2)
    End If
End Sub

Private Sub btnMenos_Click()
    Dim valorAtual As Double
    Dim unidadeAtual As String
    
    valorAtual = CDbl(Me.txtQuantidade.Value)
    
    ' Verificar unidade do produto selecionado para decremento inteligente
    If Me.lstProdutos.ListIndex >= 0 Then
        unidadeAtual = UCase(Trim(CStr(Me.lstProdutos.List(Me.lstProdutos.ListIndex, 2))))
    Else
        unidadeAtual = "UN"
    End If
    
    ' Decremento baseado na unidade específica da madeireira
    Dim decremento As Double
    Select Case unidadeAtual
        Case "CM"
            If valorAtual <= 10 Then
                decremento = 1      ' 1 cm para valores pequenos
            ElseIf valorAtual <= 100 Then
                decremento = 5      ' 5 cm para valores médios
            Else
                decremento = 10     ' 10 cm para valores grandes
            End If
            
        Case "M"
            If valorAtual <= 1 Then
                decremento = 0.1    ' 10 cm
            ElseIf valorAtual <= 10 Then
                decremento = 0.5    ' 50 cm
            Else
                decremento = 1      ' 1 metro
            End If
            
        Case "M²", "M2"
            If valorAtual <= 1 Then
                decremento = 0.25   ' 0,25 m²
            ElseIf valorAtual <= 10 Then
                decremento = 0.5    ' 0,5 m²
            Else
                decremento = 1      ' 1 m²
            End If
            
        Case "M³", "M3"
            If valorAtual <= 1 Then
                decremento = 0.1    ' 0,1 m³
            Else
                decremento = 0.5    ' 0,5 m³
            End If
            
        Case "PÇ", "PC", "UN", "PEÇA"
            decremento = 1          ' 1 peça
            
        Case "KG"
            If valorAtual <= 1 Then
                decremento = 0.1    ' 100g
            ElseIf valorAtual <= 10 Then
                decremento = 0.5    ' 500g
            Else
                decremento = 1      ' 1 kg
            End If
            
        Case "L", "LT"
            If valorAtual <= 1 Then
                decremento = 0.1    ' 100ml
            ElseIf valorAtual <= 10 Then
                decremento = 0.5    ' 500ml
            Else
                decremento = 1      ' 1 litro
            End If
            
        Case Else ' Outras unidades
            If valorAtual <= 1 Then
                decremento = 0.1
            ElseIf valorAtual <= 10 Then
                decremento = 0.5
            Else
                decremento = 1
            End If
    End Select
    
    If (valorAtual - decremento) >= QTD_MINIMA Then
        Me.txtQuantidade.Value = Round(valorAtual - decremento, 2)
    End If
End Sub

' === OUTROS EVENTOS E MÉTODOS MANTIDOS ===

Private Sub optDescontoPercentual_Click()
    On Error Resume Next
    If Me.optDescontoPercentual.Value = True Then
        Me.lblDesconto.Caption = "Desconto (%):"
        Me.txtDesconto.Value = ""
    End If
    On Error GoTo 0
End Sub

Private Sub optDescontoValor_Click()
    On Error Resume Next
    If Me.optDescontoValor.Value = True Then
        Me.lblDesconto.Caption = "Desconto (R$):"
        Me.txtDesconto.Value = ""
    End If
    On Error GoTo 0
End Sub

Private Sub lblTotal_Click()
    ' Forçar atualização quando o usuário clicar no total
    Call AtualizarResumo
End Sub

Private Sub lstProdutos_Click()
    ' Evento mantido para compatibilidade
End Sub

Private Sub txtPesquisa_Change()
    Dim termoBusca As String
    termoBusca = UCase(Trim(Me.txtPesquisa.Value))
    If Len(termoBusca) >= 2 Then
        Call PesquisarProdutos(termoBusca)
    ElseIf Len(termoBusca) = 0 Then
        Call CarregarTodosProdutos
    End If
End Sub

Private Sub btnLimparPesquisa_Click()
    Me.txtPesquisa.Value = ""
    Me.txtPesquisa.SetFocus
End Sub

Private Sub CarregarTodosProdutos()
    On Error GoTo TratarErro
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Produtos")
    Dim ultimaLinha As Long
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Me.lstProdutos.Clear

    Dim i As Long
    For i = 2 To ultimaLinha
        If Trim(CStr(ws.Cells(i, 1).Value)) <> "" Then
            With Me.lstProdutos
                .AddItem
                .List(.ListCount - 1, 0) = ws.Cells(i, 1).Value
                .List(.ListCount - 1, 1) = ws.Cells(i, 2).Value
                .List(.ListCount - 1, 2) = ws.Cells(i, 3).Value
                .List(.ListCount - 1, 3) = ws.Cells(i, 4).Value
                ' Usar coluna F (6) para preço de venda
                .List(.ListCount - 1, 4) = Format(ws.Cells(i, 6).Value, "R$ #,##0.00")
            End With
        End If
    Next i
    
    Exit Sub

TratarErro:
    MsgBox "Erro ao carregar produtos: " & Err.Description, vbCritical
End Sub

Private Sub PesquisarProdutos(termo As String)
    On Error GoTo TratarErro
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Produtos")
    Me.lstProdutos.Clear

    Dim i As Long
    For i = 2 To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        If UCase(CStr(ws.Cells(i, 1).Value)) Like "*" & termo & "*" Or _
           UCase(CStr(ws.Cells(i, 2).Value)) Like "*" & termo & "*" Then
            With Me.lstProdutos
                .AddItem
                .List(.ListCount - 1, 0) = ws.Cells(i, 1).Value
                .List(.ListCount - 1, 1) = ws.Cells(i, 2).Value
                .List(.ListCount - 1, 2) = ws.Cells(i, 3).Value
                .List(.ListCount - 1, 3) = ws.Cells(i, 4).Value
                ' Usar coluna F (6) para preço de venda
                .List(.ListCount - 1, 4) = Format(ws.Cells(i, 6).Value, "R$ #,##0.00")
            End With
        End If
    Next i
    
    Exit Sub

TratarErro:
    MsgBox "Erro na pesquisa: " & Err.Description, vbCritical
End Sub

Private Sub UserForm_Initialize()
    On Error GoTo TratarErro
    
    Me.txtQuantidade.Value = 1
    Call ConfigurarListas

    ' Configurar opções padrão de desconto
    On Error Resume Next
    ' Me.optDescontoPercentual.Value = True
    ' Me.lblDesconto.Caption = "Desconto (%):"
    On Error GoTo 0

    Call CarregarTodosProdutos
    Call AtualizarResumo
    Me.txtPesquisa.SetFocus
    
    Exit Sub

TratarErro:
    MsgBox "Erro ao iniciar: " & Err.Description, vbCritical
End Sub

Private Sub ConfigurarListas()
    With lstProdutos
        .ColumnCount = 5
        .ColumnWidths = "60;250;100;80;60"
    End With
    
    With lstSelecionados
        .ColumnCount = 7
        .ColumnWidths = "60;140;30;60;40;80;60"
    End With
End Sub