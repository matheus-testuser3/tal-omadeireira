Imports System.ComponentModel

''' <summary>
''' Sistema de cálculos para o PDV - Madeireira Maria Luiza
''' Gerencia cálculos de totais, descontos, frete e impostos
''' </summary>
Public Class CalculadoraPDV
    
    #Region "Propriedades"
    
    ''' <summary>
    ''' Lista de itens da venda
    ''' </summary>
    Public Property Itens As List(Of ItemVenda)
    
    ''' <summary>
    ''' Desconto geral em valor
    ''' </summary>
    Public Property DescontoGeral As Decimal
    
    ''' <summary>
    ''' Desconto geral em percentual
    ''' </summary>
    Public Property DescontoPercentual As Decimal
    
    ''' <summary>
    ''' Valor do frete
    ''' </summary>
    Public Property Frete As Decimal
    
    ''' <summary>
    ''' Taxa de acréscimo (cartão, etc.)
    ''' </summary>
    Public Property TaxaAcrescimo As Decimal
    
    ''' <summary>
    ''' Percentual de comissão do vendedor
    ''' </summary>
    Public Property ComissaoVendedor As Decimal
    
    #End Region
    
    #Region "Construtor"
    
    Public Sub New()
        Itens = New List(Of ItemVenda)()
        DescontoGeral = 0
        DescontoPercentual = 0
        Frete = 0
        TaxaAcrescimo = 0
        ComissaoVendedor = 0
    End Sub
    
    Public Sub New(itens As List(Of ItemVenda))
        Me.New()
        Me.Itens = itens
    End Sub
    
    #End Region
    
    #Region "Cálculos de Itens"
    
    ''' <summary>
    ''' Calcula subtotal de um item específico
    ''' </summary>
    Public Function CalcularSubtotalItem(item As ItemVenda) As Decimal
        If item Is Nothing Then Return 0
        
        Dim subtotal = item.Quantidade * item.PrecoUnitario
        Return subtotal - item.Desconto
    End Function
    
    ''' <summary>
    ''' Calcula subtotal de todos os itens (sem descontos gerais)
    ''' </summary>
    Public Function CalcularSubtotalItens() As Decimal
        Return Itens.Sum(Function(item) CalcularSubtotalItem(item))
    End Function
    
    ''' <summary>
    ''' Calcula quantidade total de itens
    ''' </summary>
    Public Function CalcularQuantidadeTotalItens() As Integer
        Return Itens.Count
    End Function
    
    ''' <summary>
    ''' Calcula peso/volume total (se aplicável)
    ''' </summary>
    Public Function CalcularQuantidadeTotalUnidades() As Double
        Return Itens.Sum(Function(item) item.Quantidade)
    End Function
    
    #End Region
    
    #Region "Cálculos de Descontos"
    
    ''' <summary>
    ''' Calcula desconto total em valor
    ''' </summary>
    Public Function CalcularDescontoTotal() As Decimal
        Dim descontoItens = Itens.Sum(Function(item) item.Desconto)
        Dim descontoGeralValor = DescontoGeral
        
        ' Se há desconto percentual, calcular sobre o subtotal
        If DescontoPercentual > 0 Then
            Dim subtotal = CalcularSubtotalItens()
            descontoGeralValor += (subtotal * DescontoPercentual / 100)
        End If
        
        Return descontoItens + descontoGeralValor
    End Function
    
    ''' <summary>
    ''' Calcula percentual de desconto total sobre o subtotal
    ''' </summary>
    Public Function CalcularPercentualDescontoTotal() As Decimal
        Dim subtotal = CalcularSubtotalItens()
        If subtotal = 0 Then Return 0
        
        Dim descontoTotal = CalcularDescontoTotal()
        Return (descontoTotal / subtotal) * 100
    End Function
    
    ''' <summary>
    ''' Aplica desconto percentual a todos os itens
    ''' </summary>
    Public Sub AplicarDescontoPercentualItens(percentual As Decimal)
        For Each item In Itens
            Dim descontoItem = (item.Quantidade * item.PrecoUnitario) * (percentual / 100)
            item.Desconto += descontoItem
        Next
    End Sub
    
    #End Region
    
    #Region "Cálculos de Totais"
    
    ''' <summary>
    ''' Calcula total líquido (subtotal - descontos + frete + acréscimos)
    ''' </summary>
    Public Function CalcularTotalLiquido() As Decimal
        Dim subtotal = CalcularSubtotalItens()
        Dim desconto = CalcularDescontoTotal()
        Dim acrescimo = CalcularAcrescimo()
        
        Return subtotal - desconto + Frete + acrescimo
    End Function
    
    ''' <summary>
    ''' Calcula acréscimo total (taxas de cartão, etc.)
    ''' </summary>
    Public Function CalcularAcrescimo() As Decimal
        If TaxaAcrescimo <= 0 Then Return 0
        
        Dim subtotal = CalcularSubtotalItens() - CalcularDescontoTotal()
        Return (subtotal * TaxaAcrescimo / 100)
    End Function
    
    ''' <summary>
    ''' Calcula valor da comissão do vendedor
    ''' </summary>
    Public Function CalcularComissaoVendedor() As Decimal
        If ComissaoVendedor <= 0 Then Return 0
        
        Dim totalLiquido = CalcularTotalLiquido()
        Return (totalLiquido * ComissaoVendedor / 100)
    End Function
    
    #End Region
    
    #Region "Utilitários de Formatação"
    
    ''' <summary>
    ''' Formata valor monetário para exibição
    ''' </summary>
    Public Shared Function FormatarMoeda(valor As Decimal) As String
        Return valor.ToString("C2", System.Globalization.CultureInfo.GetCultureInfo("pt-BR"))
    End Function
    
    ''' <summary>
    ''' Formata percentual para exibição
    ''' </summary>
    Public Shared Function FormatarPercentual(valor As Decimal) As String
        Return valor.ToString("N2") & "%"
    End Function
    
    ''' <summary>
    ''' Formata quantidade para exibição
    ''' </summary>
    Public Shared Function FormatarQuantidade(valor As Double) As String
        Return valor.ToString("N3")
    End Function
    
    #End Region
    
    #Region "Validações"
    
    ''' <summary>
    ''' Valida se os cálculos estão consistentes
    ''' </summary>
    Public Function ValidarCalculos() As List(Of String)
        Dim erros As New List(Of String)()
        
        ' Validar itens
        If Itens Is Nothing OrElse Itens.Count = 0 Then
            erros.Add("Nenhum item adicionado à venda")
        End If
        
        ' Validar valores negativos
        For Each item In Itens
            If item.Quantidade <= 0 Then
                erros.Add($"Quantidade inválida para item {item.Produto.Descricao}")
            End If
            
            If item.PrecoUnitario < 0 Then
                erros.Add($"Preço negativo para item {item.Produto.Descricao}")
            End If
            
            If item.Desconto < 0 Then
                erros.Add($"Desconto negativo para item {item.Produto.Descricao}")
            End If
        Next
        
        ' Validar descontos
        If DescontoPercentual < 0 OrElse DescontoPercentual > 100 Then
            erros.Add("Desconto percentual deve estar entre 0% e 100%")
        End If
        
        If DescontoGeral < 0 Then
            erros.Add("Desconto geral não pode ser negativo")
        End If
        
        ' Validar se desconto não é maior que subtotal
        Dim subtotal = CalcularSubtotalItens()
        Dim desconto = CalcularDescontoTotal()
        If desconto > subtotal Then
            erros.Add("Desconto total não pode ser maior que o subtotal")
        End If
        
        ' Validar frete
        If Frete < 0 Then
            erros.Add("Frete não pode ser negativo")
        End If
        
        ' Validar total final
        Dim total = CalcularTotalLiquido()
        If total <= 0 Then
            erros.Add("Total da venda deve ser positivo")
        End If
        
        Return erros
    End Function
    
    ''' <summary>
    ''' Verifica se os cálculos são válidos
    ''' </summary>
    Public Function CalculosValidos() As Boolean
        Return ValidarCalculos().Count = 0
    End Function
    
    #End Region
    
    #Region "Relatório de Cálculos"
    
    ''' <summary>
    ''' Gera resumo detalhado dos cálculos
    ''' </summary>
    Public Function GerarResumoCalculos() As String
        Dim sb As New System.Text.StringBuilder()
        
        sb.AppendLine("=== RESUMO DE CÁLCULOS ===")
        sb.AppendLine()
        sb.AppendLine($"Quantidade de itens: {CalcularQuantidadeTotalItens()}")
        sb.AppendLine($"Quantidade total: {FormatarQuantidade(CalcularQuantidadeTotalUnidades())}")
        sb.AppendLine()
        sb.AppendLine($"Subtotal dos itens: {FormatarMoeda(CalcularSubtotalItens())}")
        sb.AppendLine($"Desconto dos itens: {FormatarMoeda(Itens.Sum(Function(i) i.Desconto))}")
        
        If DescontoGeral > 0 Then
            sb.AppendLine($"Desconto geral: {FormatarMoeda(DescontoGeral)}")
        End If
        
        If DescontoPercentual > 0 Then
            sb.AppendLine($"Desconto percentual: {FormatarPercentual(DescontoPercentual)}")
        End If
        
        sb.AppendLine($"Desconto total: {FormatarMoeda(CalcularDescontoTotal())}")
        
        If Frete > 0 Then
            sb.AppendLine($"Frete: {FormatarMoeda(Frete)}")
        End If
        
        If TaxaAcrescimo > 0 Then
            sb.AppendLine($"Taxa de acréscimo ({FormatarPercentual(TaxaAcrescimo)}): {FormatarMoeda(CalcularAcrescimo())}")
        End If
        
        sb.AppendLine()
        sb.AppendLine($"TOTAL LÍQUIDO: {FormatarMoeda(CalcularTotalLiquido())}")
        
        If ComissaoVendedor > 0 Then
            sb.AppendLine($"Comissão vendedor ({FormatarPercentual(ComissaoVendedor)}): {FormatarMoeda(CalcularComissaoVendedor())}")
        End If
        
        Return sb.ToString()
    End Function
    
    #End Region
End Class

''' <summary>
''' Classe para cálculos específicos de madeireira
''' </summary>
Public Class CalculadoraMadeireira
    Inherits CalculadoraPDV
    
    ''' <summary>
    ''' Calcula metragem total de madeira
    ''' </summary>
    Public Function CalcularMetragemTotal() As Double
        Return Itens.Where(Function(i) i.Produto.Unidade.ToUpper().Contains("M")).
                     Sum(Function(i) i.Quantidade)
    End Function
    
    ''' <summary>
    ''' Calcula peso estimado baseado em densidade da madeira
    ''' </summary>
    Public Function CalcularPesoEstimado(Optional densidadeDefault As Double = 0.6) As Double
        Dim peso As Double = 0
        
        For Each item In Itens
            ' Estimar peso baseado na unidade e quantidade
            Select Case item.Produto.Unidade.ToUpper()
                Case "M³"
                    peso += item.Quantidade * densidadeDefault * 1000 ' kg
                Case "M²"
                    peso += item.Quantidade * 0.018 * densidadeDefault * 1000 ' 18mm padrão
                Case "M"
                    peso += item.Quantidade * 0.1 * densidadeDefault ' estimativa linear
                Case Else
                    peso += item.Quantidade * 5 ' peso estimado por unidade
            End Select
        Next
        
        Return peso
    End Function
    
    ''' <summary>
    ''' Calcula frete baseado em peso e distância
    ''' </summary>
    Public Function CalcularFreteEstimado(distanciaKm As Double, Optional tarifaPorKg As Decimal = 0.5) As Decimal
        Dim peso = CalcularPesoEstimado()
        Dim fretePeso = peso * tarifaPorKg
        Dim freteDistancia = distanciaKm * 2 ' R$ 2,00 por km
        
        Return Math.Max(fretePeso, freteDistancia)
    End Function
    
    ''' <summary>
    ''' Aplica desconto progressivo baseado na quantidade
    ''' </summary>
    Public Sub AplicarDescontoProgressivo()
        Dim total = CalcularTotalLiquido()
        
        Dim desconto As Decimal = 0
        
        If total >= 1000 Then
            desconto = 5 ' 5% acima de R$ 1.000
        End If
        
        If total >= 2000 Then
            desconto = 7 ' 7% acima de R$ 2.000
        End If
        
        If total >= 5000 Then
            desconto = 10 ' 10% acima de R$ 5.000
        End If
        
        If desconto > 0 Then
            DescontoPercentual = Math.Max(DescontoPercentual, desconto)
        End If
    End Sub
End Class

''' <summary>
''' Utilitários para cálculos em tempo real
''' </summary>
Public Module CalculosUtilities
    
    ''' <summary>
    ''' Atualiza totais em controles de interface
    ''' </summary>
    Public Sub AtualizarTotaisInterface(calculadora As CalculadoraPDV,
                                      lblSubtotal As Label,
                                      lblDesconto As Label,
                                      lblTotal As Label,
                                      Optional lblFrete As Label = Nothing,
                                      Optional lblQuantidade As Label = Nothing)
        Try
            lblSubtotal.Text = CalculadoraPDV.FormatarMoeda(calculadora.CalcularSubtotalItens())
            lblDesconto.Text = CalculadoraPDV.FormatarMoeda(calculadora.CalcularDescontoTotal())
            lblTotal.Text = CalculadoraPDV.FormatarMoeda(calculadora.CalcularTotalLiquido())
            
            If lblFrete IsNot Nothing Then
                lblFrete.Text = CalculadoraPDV.FormatarMoeda(calculadora.Frete)
            End If
            
            If lblQuantidade IsNot Nothing Then
                lblQuantidade.Text = calculadora.CalcularQuantidadeTotalItens().ToString() & " itens"
            End If
            
        Catch ex As Exception
            Console.WriteLine($"Erro ao atualizar interface: {ex.Message}")
        End Try
    End Sub
    
    ''' <summary>
    ''' Configura eventos de mudança automática para cálculos em tempo real
    ''' </summary>
    Public Sub ConfigurarCalculoAutomatico(dgvItens As DataGridView, calculadora As CalculadoraPDV)
        AddHandler dgvItens.CellValueChanged, Sub(sender, e)
            Try
                ' Recalcular quando valores mudam
                If e.ColumnIndex >= 0 AndAlso e.RowIndex >= 0 Then
                    Dim row = dgvItens.Rows(e.RowIndex)
                    If row.Tag IsNot Nothing AndAlso TypeOf row.Tag Is ItemVenda Then
                        Dim item = CType(row.Tag, ItemVenda)
                        
                        ' Atualizar valores do item baseado na grid
                        If dgvItens.Columns(e.ColumnIndex).Name.Contains("Quantidade") Then
                            Double.TryParse(row.Cells("Quantidade").Value?.ToString(), item.Quantidade)
                        ElseIf dgvItens.Columns(e.ColumnIndex).Name.Contains("Preco") Then
                            Decimal.TryParse(row.Cells("PrecoUnitario").Value?.ToString(), item.PrecoUnitario)
                        ElseIf dgvItens.Columns(e.ColumnIndex).Name.Contains("Desconto") Then
                            Decimal.TryParse(row.Cells("Desconto").Value?.ToString(), item.Desconto)
                        End If
                        
                        ' Atualizar subtotal na grid
                        row.Cells("Subtotal").Value = calculadora.CalcularSubtotalItem(item)
                    End If
                End If
            Catch ex As Exception
                Console.WriteLine($"Erro no cálculo automático: {ex.Message}")
            End Try
        End Sub
    End Sub
    
    ''' <summary>
    ''' Valida entrada numérica em TextBox
    ''' </summary>
    Public Sub ValidarEntradaNumerica(textBox As TextBox, Optional decimais As Boolean = True)
        AddHandler textBox.KeyPress, Sub(sender, e)
            If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) Then
                If decimais AndAlso (e.KeyChar = "."c OrElse e.KeyChar = ","c) Then
                    ' Permitir apenas um separador decimal
                    If textBox.Text.Contains(".") OrElse textBox.Text.Contains(",") Then
                        e.Handled = True
                    End If
                Else
                    e.Handled = True
                End If
            End If
        End Sub
    End Sub
End Module