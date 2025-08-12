Imports System.Windows.Forms
Imports System.Drawing

''' <summary>
''' Sistema universal de redimensionamento para diferentes resoluções
''' Adapta automaticamente a interface para 1366x768 e outras resoluções
''' </summary>
Public Class SistemaRedimensionamento

    ' Resolução base para cálculos (1366x768)
    Private Shared ReadOnly RESOLUCAO_BASE_WIDTH As Integer = 1366
    Private Shared ReadOnly RESOLUCAO_BASE_HEIGHT As Integer = 768

    ' Fatores de escala calculados
    Private Shared fatorEscalaX As Single = 1.0F
    Private Shared fatorEscalaY As Single = 1.0F

    ' Resolução atual do sistema
    Private Shared resolucaoAtual As Size

    ''' <summary>
    ''' Inicializa o sistema de redimensionamento
    ''' </summary>
    Shared Sub New()
        CalcularFatoresEscala()
    End Sub

    ''' <summary>
    ''' Calcula os fatores de escala baseados na resolução atual
    ''' </summary>
    Private Shared Sub CalcularFatoresEscala()
        Try
            ' Obter resolução atual da tela primária
            resolucaoAtual = Screen.PrimaryScreen.Bounds.Size

            ' Calcular fatores de escala
            fatorEscalaX = CSng(resolucaoAtual.Width) / RESOLUCAO_BASE_WIDTH
            fatorEscalaY = CSng(resolucaoAtual.Height) / RESOLUCAO_BASE_HEIGHT

            ' Limitar fatores para evitar interfaces muito grandes ou pequenas
            fatorEscalaX = Math.Max(0.75F, Math.Min(2.0F, fatorEscalaX))
            fatorEscalaY = Math.Max(0.75F, Math.Min(2.0F, fatorEscalaY))

        Catch ex As Exception
            ' Em caso de erro, usar fatores padrão
            fatorEscalaX = 1.0F
            fatorEscalaY = 1.0F
        End Try
    End Sub

    ''' <summary>
    ''' Adapta um formulário para a resolução atual
    ''' </summary>
    ''' <param name="form">Formulário a ser adaptado</param>
    Public Shared Sub AdaptarFormulario(form As Form)
        Try
            ' Calcular novo tamanho do formulário
            Dim novoWidth As Integer = CInt(form.Width * fatorEscalaX)
            Dim novoHeight As Integer = CInt(form.Height * fatorEscalaY)

            ' Aplicar novo tamanho
            form.Size = New Size(novoWidth, novoHeight)

            ' Adaptar fonte se necessário
            AdaptarFonte(form)

            ' Adaptar controles recursivamente
            AdaptarControles(form.Controls)

            ' Centralizar formulário
            form.StartPosition = FormStartPosition.CenterScreen

        Catch ex As Exception
            Console.WriteLine($"Erro ao adaptar formulário: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Adapta controles recursivamente
    ''' </summary>
    ''' <param name="controles">Coleção de controles</param>
    Private Shared Sub AdaptarControles(controles As Control.ControlCollection)
        For Each ctrl As Control In controles
            Try
                ' Adaptar posição
                ctrl.Location = New Point(
                    CInt(ctrl.Location.X * fatorEscalaX),
                    CInt(ctrl.Location.Y * fatorEscalaY)
                )

                ' Adaptar tamanho
                ctrl.Size = New Size(
                    CInt(ctrl.Size.Width * fatorEscalaX),
                    CInt(ctrl.Size.Height * fatorEscalaY)
                )

                ' Adaptar fonte
                AdaptarFonte(ctrl)

                ' Adaptações específicas por tipo de controle
                AdaptarControleEspecifico(ctrl)

                ' Processar controles filhos recursivamente
                If ctrl.HasChildren Then
                    AdaptarControles(ctrl.Controls)
                End If

            Catch ex As Exception
                ' Continuar mesmo se houver erro em um controle específico
                Console.WriteLine($"Erro ao adaptar controle {ctrl.Name}: {ex.Message}")
            End Try
        Next
    End Sub

    ''' <summary>
    ''' Adapta fonte de um controle
    ''' </summary>
    ''' <param name="ctrl">Controle a ser adaptado</param>
    Private Shared Sub AdaptarFonte(ctrl As Control)
        Try
            If ctrl.Font IsNot Nothing Then
                Dim novoTamanho As Single = ctrl.Font.Size * Math.Min(fatorEscalaX, fatorEscalaY)
                
                ' Limitar tamanho da fonte
                novoTamanho = Math.Max(7.0F, Math.Min(20.0F, novoTamanho))
                
                ctrl.Font = New Font(ctrl.Font.FontFamily, novoTamanho, ctrl.Font.Style)
            End If
        Catch ex As Exception
            ' Ignorar erros de fonte
        End Try
    End Sub

    ''' <summary>
    ''' Adaptações específicas por tipo de controle
    ''' </summary>
    ''' <param name="ctrl">Controle a ser adaptado</param>
    Private Shared Sub AdaptarControleEspecifico(ctrl As Control)
        Try
            Select Case ctrl.GetType()
                Case GetType(DataGridView)
                    AdaptarDataGridView(CType(ctrl, DataGridView))
                
                Case GetType(Button)
                    AdaptarBotao(CType(ctrl, Button))
                
                Case GetType(TextBox)
                    AdaptarTextBox(CType(ctrl, TextBox))
                
                Case GetType(ComboBox)
                    AdaptarComboBox(CType(ctrl, ComboBox))
                
                Case GetType(Panel)
                    AdaptarPanel(CType(ctrl, Panel))
                
                Case GetType(GroupBox)
                    AdaptarGroupBox(CType(ctrl, GroupBox))
            End Select
        Catch ex As Exception
            ' Ignorar erros específicos
        End Try
    End Sub

    ''' <summary>
    ''' Adapta DataGridView específico
    ''' </summary>
    Private Shared Sub AdaptarDataGridView(dgv As DataGridView)
        Try
            ' Adaptar altura das linhas
            dgv.RowTemplate.Height = CInt(dgv.RowTemplate.Height * fatorEscalaY)
            
            ' Adaptar largura das colunas proporcionalmente
            For Each coluna As DataGridViewColumn In dgv.Columns
                coluna.Width = CInt(coluna.Width * fatorEscalaX)
            Next
            
            ' Adaptar tamanho da fonte das células
            If dgv.DefaultCellStyle.Font IsNot Nothing Then
                Dim novoTamanho As Single = dgv.DefaultCellStyle.Font.Size * Math.Min(fatorEscalaX, fatorEscalaY)
                novoTamanho = Math.Max(7.0F, Math.Min(16.0F, novoTamanho))
                dgv.DefaultCellStyle.Font = New Font(dgv.DefaultCellStyle.Font.FontFamily, novoTamanho)
            End If
            
        Catch ex As Exception
            Console.WriteLine($"Erro ao adaptar DataGridView: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Adapta botões
    ''' </summary>
    Private Shared Sub AdaptarBotao(btn As Button)
        Try
            ' Aumentar padding interno se necessário
            If fatorEscalaX > 1.2 OrElse fatorEscalaY > 1.2 Then
                btn.Padding = New Padding(
                    CInt(btn.Padding.Left * fatorEscalaX),
                    CInt(btn.Padding.Top * fatorEscalaY),
                    CInt(btn.Padding.Right * fatorEscalaX),
                    CInt(btn.Padding.Bottom * fatorEscalaY)
                )
            End If
        Catch ex As Exception
            ' Ignorar erros
        End Try
    End Sub

    ''' <summary>
    ''' Adapta TextBox
    ''' </summary>
    Private Shared Sub AdaptarTextBox(txt As TextBox)
        Try
            ' Garantir altura mínima para legibilidade
            If txt.Height < 20 * fatorEscalaY Then
                txt.Height = CInt(20 * fatorEscalaY)
            End If
        Catch ex As Exception
            ' Ignorar erros
        End Try
    End Sub

    ''' <summary>
    ''' Adapta ComboBox
    ''' </summary>
    Private Shared Sub AdaptarComboBox(cmb As ComboBox)
        Try
            ' Ajustar altura do dropdown
            cmb.DropDownHeight = CInt(cmb.DropDownHeight * fatorEscalaY)
        Catch ex As Exception
            ' Ignorar erros
        End Try
    End Sub

    ''' <summary>
    ''' Adapta Panel
    ''' </summary>
    Private Shared Sub AdaptarPanel(pnl As Panel)
        Try
            ' Adaptar bordas se houver
            If pnl.BorderStyle <> BorderStyle.None Then
                ' Panel não tem propriedades de borda ajustáveis
                ' Mas podemos ajustar padding interno
                pnl.Padding = New Padding(
                    CInt(pnl.Padding.Left * fatorEscalaX),
                    CInt(pnl.Padding.Top * fatorEscalaY),
                    CInt(pnl.Padding.Right * fatorEscalaX),
                    CInt(pnl.Padding.Bottom * fatorEscalaY)
                )
            End If
        Catch ex As Exception
            ' Ignorar erros
        End Try
    End Sub

    ''' <summary>
    ''' Adapta GroupBox
    ''' </summary>
    Private Shared Sub AdaptarGroupBox(grp As GroupBox)
        Try
            ' Ajustar padding para o texto do cabeçalho
            grp.Padding = New Padding(
                CInt(grp.Padding.Left * fatorEscalaX),
                CInt(grp.Padding.Top * fatorEscalaY),
                CInt(grp.Padding.Right * fatorEscalaX),
                CInt(grp.Padding.Bottom * fatorEscalaY)
            )
        Catch ex As Exception
            ' Ignorar erros
        End Try
    End Sub

    ''' <summary>
    ''' Obtém as informações da resolução atual
    ''' </summary>
    ''' <returns>String com informações da resolução</returns>
    Public Shared Function ObterInfoResolucao() As String
        Return $"Resolução: {resolucaoAtual.Width}x{resolucaoAtual.Height} | " &
               $"Escala: {fatorEscalaX:F2}x{fatorEscalaY:F2} | " &
               $"Base: {RESOLUCAO_BASE_WIDTH}x{RESOLUCAO_BASE_HEIGHT}"
    End Function

    ''' <summary>
    ''' Verifica se a resolução atual precisa de adaptação
    ''' </summary>
    ''' <returns>True se precisar de adaptação</returns>
    Public Shared Function PrecisaAdaptacao() As Boolean
        Return Math.Abs(fatorEscalaX - 1.0F) > 0.1F OrElse Math.Abs(fatorEscalaY - 1.0F) > 0.1F
    End Function

    ''' <summary>
    ''' Calcula tamanho adaptado para um valor específico
    ''' </summary>
    ''' <param name="valor">Valor original</param>
    ''' <param name="direcao">X para horizontal, Y para vertical</param>
    ''' <returns>Valor adaptado</returns>
    Public Shared Function CalcularTamanhoAdaptado(valor As Integer, direcao As String) As Integer
        Select Case direcao.ToUpper()
            Case "X"
                Return CInt(valor * fatorEscalaX)
            Case "Y"
                Return CInt(valor * fatorEscalaY)
            Case Else
                Return CInt(valor * Math.Min(fatorEscalaX, fatorEscalaY))
        End Select
    End Function

    ''' <summary>
    ''' Aplica configurações responsivas para resolução específica
    ''' </summary>
    ''' <param name="form">Formulário a ser configurado</param>
    Public Shared Sub ConfigurarResponsivo(form As Form)
        Try
            ' Configurar âncoras para redimensionamento
            ConfigurarAncoras(form.Controls)
            
            ' Configurar MinimumSize e MaximumSize
            form.MinimumSize = New Size(
                CInt(800 * fatorEscalaX),
                CInt(600 * fatorEscalaY)
            )
            
            ' Se a tela for muito pequena, maximizar
            If resolucaoAtual.Width <= 1024 OrElse resolucaoAtual.Height <= 768 Then
                form.WindowState = FormWindowState.Maximized
            End If
            
        Catch ex As Exception
            Console.WriteLine($"Erro ao configurar responsivo: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Configura âncoras para controles principais
    ''' </summary>
    Private Shared Sub ConfigurarAncoras(controles As Control.ControlCollection)
        For Each ctrl As Control In controles
            Try
                Select Case ctrl.GetType()
                    Case GetType(DataGridView)
                        ' DataGridView deve expandir com o formulário
                        ctrl.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
                    
                    Case GetType(Panel)
                        ' Painéis principais também expandem
                        If ctrl.Dock = DockStyle.None Then
                            ctrl.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
                        End If
                    
                    Case GetType(Button)
                        ' Botões geralmente ficam ancorados à direita/bottom
                        If ctrl.Location.X > (ctrl.Parent?.Width / 2) Then
                            ctrl.Anchor = AnchorStyles.Top Or AnchorStyles.Right
                        End If
                End Select

                ' Processar controles filhos
                If ctrl.HasChildren Then
                    ConfigurarAncoras(ctrl.Controls)
                End If

            Catch ex As Exception
                ' Continuar mesmo com erros
            End Try
        Next
    End Sub

    ''' <summary>
    ''' Detecta e reporta características da tela
    ''' </summary>
    ''' <returns>Relatório das características da tela</returns>
    Public Shared Function DetectarCaracteristicasTela() As String
        Try
            Dim sb As New System.Text.StringBuilder()
            
            sb.AppendLine($"=== CARACTERÍSTICAS DA TELA ===")
            sb.AppendLine($"Resolução Primária: {Screen.PrimaryScreen.Bounds.Width}x{Screen.PrimaryScreen.Bounds.Height}")
            sb.AppendLine($"Área de Trabalho: {Screen.PrimaryScreen.WorkingArea.Width}x{Screen.PrimaryScreen.WorkingArea.Height}")
            sb.AppendLine($"DPI: {Screen.PrimaryScreen.Bounds.Width / (Screen.PrimaryScreen.Bounds.Width / 96):F1}")
            sb.AppendLine($"Número de Telas: {Screen.AllScreens.Length}")
            
            sb.AppendLine($"")
            sb.AppendLine($"=== ADAPTAÇÃO CALCULADA ===")
            sb.AppendLine($"Resolução Base: {RESOLUCAO_BASE_WIDTH}x{RESOLUCAO_BASE_HEIGHT}")
            sb.AppendLine($"Fator Escala X: {fatorEscalaX:F3}")
            sb.AppendLine($"Fator Escala Y: {fatorEscalaY:F3}")
            sb.AppendLine($"Precisa Adaptação: {If(PrecisaAdaptacao(), "SIM", "NÃO")}")
            
            Return sb.ToString()
            
        Catch ex As Exception
            Return $"Erro ao detectar características: {ex.Message}"
        End Try
    End Function

    ''' <summary>
    ''' Força recálculo dos fatores de escala (para mudanças de resolução)
    ''' </summary>
    Public Shared Sub RecalcularEscala()
        CalcularFatoresEscala()
    End Sub
End Class