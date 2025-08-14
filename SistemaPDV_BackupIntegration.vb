''' <summary>
''' Integração do sistema de backup com a interface principal - Madeireira Maria Luiza
''' Data/Hora: 2025-08-14 11:16:26 UTC
''' Usuário: matheus-testuser3
''' Sistema de Backup e Restauração de Talões
''' </summary>

Imports System.Windows.Forms
Imports System.Drawing
Imports System.IO
Imports System.Configuration

''' <summary>
''' Classe para integração do sistema de backup com o MainForm existente
''' Adiciona botões e eventos para importar backup e gerar talões
''' </summary>
Public Class SistemaPDV_BackupIntegration
    
    ' === REFERÊNCIAS ===
    Private ReadOnly mainForm As MainForm
    Private ReadOnly moduloBackup As New ModuloBackupTalao()
    
    ' === CONTROLES ADICIONADOS ===
    Private WithEvents btnImportarBackup As Button
    Private WithEvents btnGerarDeBackup As Button
    Private WithEvents lblBackupStatus As Label
    
    ' === ESTADO ===
    Private ultimosArquivosImportados As List(Of DadosTalaoMadeireira)
    
    ''' <summary>
    ''' Construtor da integração
    ''' </summary>
    Public Sub New(formularioPrincipal As MainForm)
        mainForm = formularioPrincipal
        InicializarIntegracao()
    End Sub
    
    ''' <summary>
    ''' Inicializa a integração adicionando controles ao formulário principal
    ''' </summary>
    Private Sub InicializarIntegracao()
        Try
            LogDebug("=== INICIALIZANDO INTEGRAÇÃO BACKUP ===")
            LogDebug($"Data/Hora: {DateTime.UtcNow:yyyy-MM-dd HH:mm:ss} UTC")
            LogDebug($"Usuário: matheus-testuser3")
            
            ' Encontrar o painel lateral do MainForm
            Dim pnlSidebar = EncontrarPainelSidebar()
            If pnlSidebar Is Nothing Then
                Throw New InvalidOperationException("Painel lateral não encontrado no formulário principal")
            End If
            
            ' Adicionar seção de backup
            AdicionarSecaoBackup(pnlSidebar)
            
            LogDebug("Integração de backup inicializada com sucesso")
            
        Catch ex As Exception
            LogDebug($"ERRO na inicialização da integração: {ex.Message}")
            MessageBox.Show($"Erro ao inicializar sistema de backup:{vbCrLf}{ex.Message}",
                          "Erro de Integração", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub
    
    ''' <summary>
    ''' Encontra o painel lateral no MainForm
    ''' </summary>
    Private Function EncontrarPainelSidebar() As Panel
        For Each control As Control In mainForm.Controls
            If TypeOf control Is Panel AndAlso control.Dock = DockStyle.Left Then
                Return CType(control, Panel)
            End If
        Next
        Return Nothing
    End Function
    
    ''' <summary>
    ''' Adiciona a seção de backup ao painel lateral
    ''' </summary>
    Private Sub AdicionarSecaoBackup(pnlSidebar As Panel)
        ' Calcular posição dos novos botões
        Dim ultimoBotao = EncontrarUltimoBotao(pnlSidebar)
        Dim posicaoY = If(ultimoBotao IsNot Nothing, ultimoBotao.Bottom + 20, 200)
        
        ' Separador visual
        Dim separador As New Panel()
        separador.Size = New Size(pnlSidebar.Width - 40, 2)
        separador.Location = New Point(20, posicaoY)
        separador.BackColor = Color.FromArgb(200, 200, 200)
        pnlSidebar.Controls.Add(separador)
        posicaoY += 25
        
        ' Label da seção
        Dim lblSecao As New Label()
        lblSecao.Text = "🗂️ SISTEMA DE BACKUP"
        lblSecao.Font = New Font("Segoe UI", 10.0F, FontStyle.Bold)
        lblSecao.ForeColor = Color.FromArgb(100, 100, 100)
        lblSecao.Location = New Point(20, posicaoY)
        lblSecao.Size = New Size(pnlSidebar.Width - 40, 25)
        pnlSidebar.Controls.Add(lblSecao)
        posicaoY += 35
        
        ' Botão Importar Backup
        btnImportarBackup = New Button()
        btnImportarBackup.Text = "📁 Importar Backup"
        btnImportarBackup.Size = New Size(pnlSidebar.Width - 40, 50)
        btnImportarBackup.Location = New Point(20, posicaoY)
        btnImportarBackup.BackColor = Color.FromArgb(70, 130, 180) ' Azul aço
        btnImportarBackup.ForeColor = Color.White
        btnImportarBackup.Font = New Font("Segoe UI", 10.0F, FontStyle.Bold)
        btnImportarBackup.FlatStyle = FlatStyle.Flat
        btnImportarBackup.FlatAppearance.BorderSize = 0
        btnImportarBackup.Cursor = Cursors.Hand
        pnlSidebar.Controls.Add(btnImportarBackup)
        posicaoY += 60
        
        ' Botão Gerar de Backup
        btnGerarDeBackup = New Button()
        btnGerarDeBackup.Text = "📋 Gerar de Backup"
        btnGerarDeBackup.Size = New Size(pnlSidebar.Width - 40, 50)
        btnGerarDeBackup.Location = New Point(20, posicaoY)
        btnGerarDeBackup.BackColor = Color.FromArgb(34, 139, 34) ' Verde madeira
        btnGerarDeBackup.ForeColor = Color.White
        btnGerarDeBackup.Font = New Font("Segoe UI", 10.0F, FontStyle.Bold)
        btnGerarDeBackup.FlatStyle = FlatStyle.Flat
        btnGerarDeBackup.FlatAppearance.BorderSize = 0
        btnGerarDeBackup.Cursor = Cursors.Hand
        btnGerarDeBackup.Enabled = False ' Inicialmente desabilitado
        pnlSidebar.Controls.Add(btnGerarDeBackup)
        posicaoY += 60
        
        ' Status do backup
        lblBackupStatus = New Label()
        lblBackupStatus.Text = "Aguardando importação..."
        lblBackupStatus.Font = New Font("Segoe UI", 8.5F, FontStyle.Italic)
        lblBackupStatus.ForeColor = Color.FromArgb(120, 120, 120)
        lblBackupStatus.Location = New Point(20, posicaoY)
        lblBackupStatus.Size = New Size(pnlSidebar.Width - 40, 40)
        lblBackupStatus.TextAlign = ContentAlignment.TopLeft
        pnlSidebar.Controls.Add(lblBackupStatus)
        
        ' Aplicar efeitos aos botões
        AplicarEfeitosVisuas()
        
        LogDebug($"Seção de backup adicionada na posição Y: {posicaoY - 150}")
    End Sub
    
    ''' <summary>
    ''' Encontra o último botão no painel para posicionar os novos
    ''' </summary>
    Private Function EncontrarUltimoBotao(painel As Panel) As Button
        Dim ultimoBotao As Button = Nothing
        Dim maiorY = 0
        
        For Each control As Control In painel.Controls
            If TypeOf control Is Button AndAlso control.Bottom > maiorY Then
                maiorY = control.Bottom
                ultimoBotao = CType(control, Button)
            End If
        Next
        
        Return ultimoBotao
    End Function
    
    ''' <summary>
    ''' Aplica efeitos visuais aos botões de backup
    ''' </summary>
    Private Sub AplicarEfeitosVisuas()
        ' Efeito hover para botão importar
        AddHandler btnImportarBackup.MouseEnter, Sub()
                                                     btnImportarBackup.BackColor = Color.FromArgb(85, 145, 195)
                                                 End Sub
        AddHandler btnImportarBackup.MouseLeave, Sub()
                                                     btnImportarBackup.BackColor = Color.FromArgb(70, 130, 180)
                                                 End Sub
        
        ' Efeito hover para botão gerar
        AddHandler btnGerarDeBackup.MouseEnter, Sub()
                                                   If btnGerarDeBackup.Enabled Then
                                                       btnGerarDeBackup.BackColor = Color.FromArgb(49, 154, 49)
                                                   End If
                                               End Sub
        AddHandler btnGerarDeBackup.MouseLeave, Sub()
                                                   If btnGerarDeBackup.Enabled Then
                                                       btnGerarDeBackup.BackColor = Color.FromArgb(34, 139, 34)
                                                   End If
                                               End Sub
    End Sub
    
    ' === EVENTOS DOS BOTÕES ===
    
    ''' <summary>
    ''' Evento do botão Importar Backup
    ''' </summary>
    Private Sub btnImportarBackup_Click(sender As Object, e As EventArgs) Handles btnImportarBackup.Click
        Try
            LogDebug("=== INÍCIO IMPORTAÇÃO BACKUP ===")
            
            ' Abrir diálogo para seleção do arquivo
            Using openFileDialog As New OpenFileDialog()
                openFileDialog.Title = "Selecionar Arquivo de Backup de Talões"
                openFileDialog.Filter = "Arquivos Excel (*.xlsx;*.xls)|*.xlsx;*.xls|Todos os arquivos (*.*)|*.*"
                openFileDialog.FilterIndex = 1
                openFileDialog.RestoreDirectory = True
                
                If openFileDialog.ShowDialog(mainForm) = DialogResult.OK Then
                    Dim caminhoArquivo = openFileDialog.FileName
                    
                    LogDebug($"Arquivo selecionado: {caminhoArquivo}")
                    
                    ' Mostrar indicador de progresso
                    AtualizarStatus("Importando backup...")
                    btnImportarBackup.Enabled = False
                    mainForm.Cursor = Cursors.WaitCursor
                    
                    Try
                        ' Importar backup
                        ultimosArquivosImportados = moduloBackup.ImportarBackupExcel(caminhoArquivo)
                        
                        ' Atualizar interface
                        btnGerarDeBackup.Enabled = (ultimosArquivosImportados.Count > 0)
                        AtualizarStatus($"✅ {ultimosArquivosImportados.Count} talões importados de:{vbCrLf}{Path.GetFileName(caminhoArquivo)}")
                        
                        LogDebug($"Importação concluída: {ultimosArquivosImportados.Count} talões")
                        
                        ' Mostrar resultado
                        MessageBox.Show($"Backup importado com sucesso!{vbCrLf}{vbCrLf}" &
                                      $"Total de talões encontrados: {ultimosArquivosImportados.Count}{vbCrLf}" &
                                      $"Arquivo: {Path.GetFileName(caminhoArquivo)}",
                                      "Importação Concluída", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        
                    Catch ex As Exception
                        LogDebug($"ERRO na importação: {ex.Message}")
                        AtualizarStatus("❌ Erro na importação")
                        
                        MessageBox.Show($"Erro ao importar backup:{vbCrLf}{vbCrLf}{ex.Message}",
                                      "Erro de Importação", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Finally
                        btnImportarBackup.Enabled = True
                        mainForm.Cursor = Cursors.Default
                    End Try
                End If
            End Using
            
        Catch ex As Exception
            LogDebug($"ERRO geral na importação: {ex.Message}")
            MessageBox.Show($"Erro inesperado:{vbCrLf}{ex.Message}",
                          "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    ''' <summary>
    ''' Evento do botão Gerar de Backup
    ''' </summary>
    Private Sub btnGerarDeBackup_Click(sender As Object, e As EventArgs) Handles btnGerarDeBackup.Click
        Try
            LogDebug("=== INÍCIO GERAÇÃO DE BACKUP ===")
            
            If ultimosArquivosImportados Is Nothing OrElse ultimosArquivosImportados.Count = 0 Then
                MessageBox.Show("Nenhum backup foi importado. Por favor, importe um arquivo de backup primeiro.",
                              "Backup Necessário", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            
            ' Abrir formulário de seleção
            Using formSelecao As New FormSelecaoTalaoBackup(ultimosArquivosImportados)
                If formSelecao.ShowDialog(mainForm) = DialogResult.OK Then
                    Dim talaoSelecionado = formSelecao.TalaoSelecionado
                    
                    If talaoSelecionado IsNot Nothing Then
                        LogDebug($"Talão selecionado: {talaoSelecionado.NumeroTalao}")
                        
                        ' Gerar talão formatado
                        AtualizarStatus("Gerando talão...")
                        btnGerarDeBackup.Enabled = False
                        mainForm.Cursor = Cursors.WaitCursor
                        
                        Try
                            Dim caminhoTalao = moduloBackup.GerarTalaoFormatado(talaoSelecionado)
                            
                            AtualizarStatus($"✅ Talão {talaoSelecionado.NumeroTalao} gerado")
                            
                            LogDebug($"Talão gerado: {caminhoTalao}")
                            
                            ' Perguntar se quer abrir o arquivo
                            Dim resultado = MessageBox.Show($"Talão gerado com sucesso!{vbCrLf}{vbCrLf}" &
                                                           $"Cliente: {talaoSelecionado.NomeCliente}{vbCrLf}" &
                                                           $"Valor: {talaoSelecionado.ValorTotal:C2}{vbCrLf}{vbCrLf}" &
                                                           $"Deseja abrir o arquivo agora?",
                                                           "Talão Gerado", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                            
                            If resultado = DialogResult.Yes AndAlso File.Exists(caminhoTalao) Then
                                Process.Start(caminhoTalao)
                            End If
                            
                        Catch ex As Exception
                            LogDebug($"ERRO na geração: {ex.Message}")
                            AtualizarStatus("❌ Erro na geração")
                            
                            MessageBox.Show($"Erro ao gerar talão:{vbCrLf}{vbCrLf}{ex.Message}",
                                          "Erro de Geração", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Finally
                            btnGerarDeBackup.Enabled = True
                            mainForm.Cursor = Cursors.Default
                        End Try
                    End If
                End If
            End Using
            
        Catch ex As Exception
            LogDebug($"ERRO geral na geração: {ex.Message}")
            MessageBox.Show($"Erro inesperado:{vbCrLf}{ex.Message}",
                          "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    ''' <summary>
    ''' Atualiza o status na interface
    ''' </summary>
    Private Sub AtualizarStatus(mensagem As String)
        If lblBackupStatus IsNot Nothing Then
            lblBackupStatus.Text = mensagem
            lblBackupStatus.Refresh()
        End If
        LogDebug($"Status atualizado: {mensagem}")
    End Sub
    
    ''' <summary>
    ''' Log de debug específico do sistema de backup
    ''' </summary>
    Private Sub LogDebug(mensagem As String)
        Debug.WriteLine($"[BACKUP-INTEGRATION] {DateTime.UtcNow:HH:mm:ss.fff} - {mensagem}")
    End Sub
    
    ''' <summary>
    ''' Método público para testar a integração
    ''' </summary>
    Public Sub TestarIntegracao()
        LogDebug("=== TESTE DE INTEGRAÇÃO ===")
        
        Try
            ' Verificar se os controles foram criados
            If btnImportarBackup Is Nothing OrElse btnGerarDeBackup Is Nothing Then
                Throw New InvalidOperationException("Controles de backup não foram inicializados")
            End If
            
            ' Verificar se os controles estão no formulário
            If Not mainForm.Controls.Contains(btnImportarBackup.Parent) Then
                Throw New InvalidOperationException("Controles de backup não estão no formulário principal")
            End If
            
            LogDebug("✅ Integração testada com sucesso")
            MessageBox.Show("Sistema de backup integrado e funcionando!",
                          "Teste de Integração", MessageBoxButtons.OK, MessageBoxIcon.Information)
            
        Catch ex As Exception
            LogDebug($"❌ ERRO no teste: {ex.Message}")
            MessageBox.Show($"Erro no teste de integração:{vbCrLf}{ex.Message}",
                          "Erro de Teste", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
End Class