''' <summary>
''' Módulo VBA para criação automática de templates
''' Responsável por criar layouts profissionais dinamicamente
''' </summary>
Public Class ModuloTemplate

    ''' <summary>
    ''' Retorna o código VBA para criação de templates
    ''' </summary>
    Public Function ObterCodigoVBA() As String
        Return "
' ===== MÓDULO TEMPLATE - CRIAÇÃO AUTOMÁTICA =====
' Módulo responsável pela criação de templates profissionais
' Madeireira Maria Luiza - Sistema PDV Integrado

Option Explicit

' ===== FUNÇÃO PRINCIPAL =====
Public Sub CriarTemplateAutomatico()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Limpar planilha
    ActiveSheet.Cells.Clear
    
    ' Criar template básico
    CriarLayoutBasico
    
    ' Definir formatação profissional
    DefinirFormatacaoProfissional
    
    ' Configurar layout duplo (primeira e segunda via)
    ConfigurarLayoutDuplo
    
    ' Adicionar elementos visuais
    AdicionarElementosVisuais
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox ""Erro ao criar template: "" & Err.Description, vbCritical, ""Erro""
End Sub

' ===== LAYOUT BÁSICO =====
Private Sub CriarLayoutBasico()
    With ActiveSheet
        ' Configurar página
        With .PageSetup
            .PaperSize = xlPaperA4
            .Orientation = xlPortrait
            .LeftMargin = Application.InchesToPoints(0.5)
            .RightMargin = Application.InchesToPoints(0.5)
            .TopMargin = Application.InchesToPoints(0.5)
            .BottomMargin = Application.InchesToPoints(0.5)
            .HeaderMargin = Application.InchesToPoints(0.3)
            .FooterMargin = Application.InchesToPoints(0.3)
        End With
        
        ' Definir larguras das colunas
        .Columns(""A"").ColumnWidth = 35  ' Descrição
        .Columns(""B"").ColumnWidth = 8   ' Quantidade
        .Columns(""C"").ColumnWidth = 6   ' Unidade
        .Columns(""D"").ColumnWidth = 12  ' Preço unitário
        .Columns(""E"").ColumnWidth = 12  ' Total
        .Columns(""F"").ColumnWidth = 2   ' Espaçamento
        .Columns(""G"").ColumnWidth = 15  ' Extra
        
        ' Cabeçalho da empresa
        .Cells(1, 1).Value = ""[NOME_EMPRESA]""
        .Cells(2, 1).Value = ""[ENDERECO_EMPRESA]""
        .Cells(3, 1).Value = ""[CIDADE_EMPRESA]""
        .Cells(4, 1).Value = ""[TELEFONE_EMPRESA]""
        .Cells(5, 1).Value = ""[CNPJ_EMPRESA]""
        
        ' Dados do talão
        .Cells(7, 1).Value = ""TALÃO DE VENDA Nº: [NUMERO_TALAO]""
        .Cells(8, 1).Value = ""Data: [DATA_HORA]""
        
        ' Dados do cliente
        .Cells(10, 1).Value = ""CLIENTE: [NOME_CLIENTE]""
        .Cells(11, 1).Value = ""ENDEREÇO: [ENDERECO_CLIENTE]""
        .Cells(12, 1).Value = ""CIDADE/CEP: [CIDADE_CEP_CLIENTE]""
        .Cells(13, 1).Value = ""TELEFONE: [TELEFONE_CLIENTE]""
        
        ' Cabeçalho da tabela de produtos
        .Cells(15, 1).Value = ""DESCRIÇÃO""
        .Cells(15, 2).Value = ""QTD""
        .Cells(15, 3).Value = ""UN""
        .Cells(15, 4).Value = ""PREÇO UNIT.""
        .Cells(15, 5).Value = ""TOTAL""
        
        ' Área para produtos (16-25)
        Dim i As Integer
        For i = 16 To 25
            .Cells(i, 1).Value = ""[PRODUTO_"" & (i - 15) & ""]""
        Next i
        
        ' Totais
        .Cells(27, 4).Value = ""TOTAL GERAL:""
        .Cells(27, 5).Value = ""[TOTAL_GERAL]""
        
        ' Forma de pagamento e vendedor
        .Cells(29, 1).Value = ""FORMA DE PAGAMENTO: [FORMA_PAGAMENTO]""
        .Cells(30, 1).Value = ""VENDEDOR: [VENDEDOR]""
        
        ' Assinatura
        .Cells(32, 1).Value = ""CLIENTE: _________________________________""
        .Cells(33, 1).Value = ""           (NOME E ASSINATURA)""
    End With
End Sub

' ===== FORMATAÇÃO PROFISSIONAL =====
Public Sub DefinirFormatacaoProfissional()
    With ActiveSheet
        ' Cabeçalho da empresa
        With .Range(""A1:G1"")
            .Merge
            .Font.Size = 18
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Font.Name = ""Arial""
        End With
        
        With .Range(""A2:G2"")
            .Merge
            .HorizontalAlignment = xlCenter
            .Font.Size = 10
        End With
        
        With .Range(""A3:G3"")
            .Merge
            .HorizontalAlignment = xlCenter
            .Font.Size = 10
        End With
        
        With .Range(""A4:G4"")
            .Merge
            .HorizontalAlignment = xlCenter
            .Font.Size = 10
        End With
        
        With .Range(""A5:G5"")
            .Merge
            .HorizontalAlignment = xlCenter
            .Font.Size = 10
            .Font.Bold = True
        End With
        
        ' Linha separadora após cabeçalho
        With .Range(""A6:G6"")
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Weight = xlThick
        End With
        
        ' Título do talão
        With .Cells(7, 1)
            .Font.Size = 14
            .Font.Bold = True
        End With
        
        With .Cells(8, 1)
            .Font.Bold = True
        End With
        
        ' Dados do cliente - formatação
        With .Range(""A10:A13"")
            .Font.Bold = True
        End With
        
        ' Cabeçalho da tabela de produtos
        With .Range(""A15:E15"")
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Interior.Color = RGB(230, 230, 230)
        End With
        
        ' Bordas da tabela de produtos
        With .Range(""A15:E25"")
            .Borders.LineStyle = xlContinuous
        End With
        
        ' Total geral
        With .Range(""D27:E27"")
            .Font.Bold = True
            .Borders.LineStyle = xlContinuous
            .Borders(xlEdgeTop).Weight = xlThick
        End With
        
        With .Cells(27, 4)
            .HorizontalAlignment = xlRight
        End With
        
        ' Forma de pagamento e vendedor
        With .Range(""A29:A30"")
            .Font.Bold = True
        End With
        
        ' Assinatura
        With .Cells(32, 1)
            .Font.Bold = True
        End With
        
        With .Cells(33, 1)
            .Font.Size = 8
        End With
    End With
End Sub

' ===== LAYOUT DUPLO =====
Public Sub ConfigurarLayoutDuplo()
    Dim LinhaSegundaVia As Integer
    LinhaSegundaVia = 36
    
    With ActiveSheet
        ' Separador
        .Cells(LinhaSegundaVia, 1).Value = ""✂️ --- CORTE AQUI - SEGUNDA VIA --- ✂️""
        With .Range(""A"" & LinhaSegundaVia & "":E"" & LinhaSegundaVia)
            .Merge
            .HorizontalAlignment = xlCenter
            .Font.Bold = True
            .Font.Size = 10
        End With
        
        LinhaSegundaVia = LinhaSegundaVia + 2
        
        ' Cabeçalho resumido da segunda via
        .Cells(LinhaSegundaVia, 1).Value = ""[NOME_EMPRESA]""
        With .Range(""A"" & LinhaSegundaVia & "":E"" & LinhaSegundaVia)
            .Merge
            .Font.Size = 14
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
        
        LinhaSegundaVia = LinhaSegundaVia + 1
        
        ' Dados resumidos da segunda via
        .Cells(LinhaSegundaVia, 1).Value = ""TALÃO Nº: [NUMERO_TALAO] - [DATA_HORA]""
        .Cells(LinhaSegundaVia, 1).Font.Bold = True
        
        LinhaSegundaVia = LinhaSegundaVia + 1
        .Cells(LinhaSegundaVia, 1).Value = ""CLIENTE: [NOME_CLIENTE]""
        
        LinhaSegundaVia = LinhaSegundaVia + 1
        .Cells(LinhaSegundaVia, 1).Value = ""TOTAL: [TOTAL_GERAL]""
        .Cells(LinhaSegundaVia, 1).Font.Bold = True
        .Cells(LinhaSegundaVia, 1).Font.Size = 12
        
        LinhaSegundaVia = LinhaSegundaVia + 1
        .Cells(LinhaSegundaVia, 1).Value = ""PAGAMENTO: [FORMA_PAGAMENTO]""
        
        LinhaSegundaVia = LinhaSegundaVia + 1
        .Cells(LinhaSegundaVia, 1).Value = ""VENDEDOR: [VENDEDOR]""
        
        LinhaSegundaVia = LinhaSegundaVia + 2
        .Cells(LinhaSegundaVia, 1).Value = ""CLIENTE: ________________________""
        .Cells(LinhaSegundaVia, 1).Font.Bold = True
    End With
End Sub

' ===== ELEMENTOS VISUAIS =====
Private Sub AdicionarElementosVisuais()
    With ActiveSheet
        ' Adicionar bordas decorativas
        AdicionarBordasDecorativas
        
        ' Adicionar área para logomarca
        AdicionarAreaLogomarca
        
        ' Adicionar rodapé profissional
        AdicionarRodapeProfissional
    End With
End Sub

' ===== BORDAS DECORATIVAS =====
Private Sub AdicionarBordasDecorativas()
    With ActiveSheet
        ' Borda superior
        With .Range(""A1:G1"")
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Weight = xlThick
        End With
        
        ' Bordas laterais do cabeçalho
        With .Range(""A1:A5"")
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeLeft).Weight = xlMedium
        End With
        
        With .Range(""G1:G5"")
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeRight).Weight = xlMedium
        End With
        
        ' Borda ao redor da tabela principal
        With .Range(""A15:E27"")
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeLeft).Weight = xlMedium
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeRight).Weight = xlMedium
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Weight = xlMedium
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Weight = xlMedium
        End With
    End With
End Sub

' ===== ÁREA PARA LOGOMARCA =====
Private Sub AdicionarAreaLogomarca()
    With ActiveSheet
        ' Criar área reservada para logo no canto superior direito
        .Cells(1, 6).Value = ""[LOGO]""
        With .Range(""F1:G5"")
            .Merge
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Interior.Color = RGB(245, 245, 245)
            .Font.Size = 8
        End With
    End With
End Sub

' ===== RODAPÉ PROFISSIONAL =====
Private Sub AdicionarRodapeProfissional()
    Dim UltimaLinha As Integer
    UltimaLinha = 45
    
    With ActiveSheet
        ' Linha separadora
        With .Range(""A"" & (UltimaLinha - 2) & "":G"" & (UltimaLinha - 2))
            .Borders(xlEdgeTop).LineStyle = xlContinuous
        End With
        
        ' Rodapé
        .Cells(UltimaLinha, 1).Value = ""Sistema PDV - Madeireira Maria Luiza © 2024""
        With .Range(""A"" & UltimaLinha & "":G"" & UltimaLinha)
            .Merge
            .HorizontalAlignment = xlCenter
            .Font.Size = 8
            .Font.Italic = True
        End With
    End With
End Sub

' ===== FUNÇÕES DE PERSONALIZAÇÃO =====

' Aplicar tema de cores
Public Sub AplicarTema(CorPrimaria As Long, CorSecundaria As Long)
    With ActiveSheet
        ' Cabeçalho da tabela
        .Range(""A15:E15"").Interior.Color = CorPrimaria
        
        ' Área do logo
        .Range(""F1:G5"").Interior.Color = CorSecundaria
        
        ' Total geral
        .Range(""D27:E27"").Interior.Color = CorSecundaria
    End With
End Sub

' Definir fonte personalizada
Public Sub DefinirFontePersonalizada(NomeFonte As String, TamanhoBase As Integer)
    With ActiveSheet
        ' Aplicar fonte a toda a planilha
        .Cells.Font.Name = NomeFonte
        .Cells.Font.Size = TamanhoBase
        
        ' Ajustar tamanhos específicos
        .Cells(1, 1).Font.Size = TamanhoBase + 8  ' Título empresa
        .Cells(7, 1).Font.Size = TamanhoBase + 4  ' Título talão
        .Range(""A15:E15"").Font.Size = TamanhoBase ' Cabeçalho tabela
    End With
End Sub

' Configurar papel personalizado
Public Sub ConfigurarPapelPersonalizado(Largura As Double, Altura As Double)
    With ActiveSheet.PageSetup
        .PaperSize = xlPaperUser
        .PrintArea = ""A1:G45""
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With
End Sub

' Substituir placeholders
Public Sub SubstituirPlaceholders(DadosEmpresa As String, DadosCliente As String, DadosTalao As String)
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Esta função seria chamada para substituir os placeholders [NOME_EMPRESA], etc.
    ' pelos dados reais fornecidos pelo VB.NET
    
    ' Exemplo de substituição
    ws.Cells.Replace ""[NOME_EMPRESA]"", ""MADEIREIRA MARIA LUIZA""
    ws.Cells.Replace ""[ENDERECO_EMPRESA]"", ""Rua Principal, 123 - Centro""
    ws.Cells.Replace ""[CIDADE_EMPRESA]"", ""Paulista/PE - CEP: 53401-445""
    ws.Cells.Replace ""[TELEFONE_EMPRESA]"", ""Tel: (81) 3436-1234""
    ws.Cells.Replace ""[CNPJ_EMPRESA]"", ""CNPJ: 12.345.678/0001-90""
End Sub

' Criar template específico para produto
Public Sub CriarTemplateProdutoEspecifico(TipoProduto As String)
    Select Case TipoProduto
        Case ""MADEIRA""
            CriarTemplateMadeira
        Case ""MATERIAL""
            CriarTemplateMaterial
        Case ""FERRAMENTA""
            CriarTemplateFerramentas
        Case Else
            CriarTemplateGenerico
    End Select
End Sub

' Template específico para madeiras
Private Sub CriarTemplateMadeira()
    With ActiveSheet
        ' Adicionar campos específicos para madeira
        .Cells(14, 1).Value = ""(Especificar: tipo, dimensões, qualidade)""
        .Cells(14, 1).Font.Size = 8
        .Cells(14, 1).Font.Italic = True
    End With
End Sub

' Template genérico
Private Sub CriarTemplateGenerico()
    ' Usar configurações padrão
    CriarLayoutBasico
End Sub

"
    End Function
End Class