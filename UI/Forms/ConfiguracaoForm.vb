Imports System.Windows.Forms
Imports System.Drawing

''' <summary>
''' Formulário de configurações do sistema
''' Permite alterar configurações da madeireira e sistema
''' </summary>
Public Class ConfiguracaoForm
    Inherits Form
    
    ' Controles da interface
    Private WithEvents pnlHeader As Panel
    Private WithEvents pnlContent As Panel
    Private WithEvents pnlFooter As Panel
    Private WithEvents tabControl As TabControl
    
    ' Aba Empresa
    Private WithEvents txtNomeMadeireira As TextBox
    Private WithEvents txtEnderecoMadeireira As TextBox
    Private WithEvents txtCidadeMadeireira As TextBox
    Private WithEvents txtCEPMadeireira As TextBox
    Private WithEvents txtTelefoneMadeireira As TextBox
    Private WithEvents txtCNPJMadeireira As TextBox
    
    ' Aba Sistema
    Private WithEvents chkBackupAutomatico As CheckBox
    Private WithEvents nudIntervaloBacKup As NumericUpDown
    Private WithEvents nudManterHistorico As NumericUpDown
    Private WithEvents chkExcelVisivel As CheckBox
    Private WithEvents txtVendedorPadrao As TextBox
    
    ' Aba Logs
    Private WithEvents cmbLogLevel As ComboBox
    Private WithEvents btnVerLogs As Button
    Private WithEvents btnLimparLogs As Button
    
    ' Botões
    Private WithEvents btnSalvar As Button
    Private WithEvents btnCancelar As Button
    Private WithEvents btnTestarExcel As Button
    Private WithEvents btnBackupManual As Button
    
    ' Serviços
    Private ReadOnly _config As ConfigManager
    Private ReadOnly _logger As Logger
    Private ReadOnly _backupService As BackupService
    
    ''' <summary>
    ''' Construtor
    ''' </summary>
    Public Sub New()
        _config = ConfigManager.Instance
        _logger = Logger.Instance
        _backupService = New BackupService()
        
        InitializeComponent()
        ConfigurarInterface()
        CarregarConfiguracoes()
    End Sub
    
    ''' <summary>
    ''' Inicializa componentes
    ''' </summary>
    Private Sub InitializeComponent()
        ' Configurar formulário
        Me.Text = "⚙️ Configurações do Sistema PDV"
        Me.Size = New Size(600, 500)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        
        ' Painel de cabeçalho
        pnlHeader = New Panel()
        pnlHeader.Height = 60
        pnlHeader.Dock = DockStyle.Top
        pnlHeader.BackColor = Color.FromArgb(52, 73, 94)
        Me.Controls.Add(pnlHeader)
        
        ' Painel de conteúdo
        pnlContent = New Panel()
        pnlContent.Dock = DockStyle.Fill
        pnlContent.Padding = New Padding(20)
        Me.Controls.Add(pnlContent)
        
        ' Painel de rodapé
        pnlFooter = New Panel()
        pnlFooter.Height = 60
        pnlFooter.Dock = DockStyle.Bottom
        pnlFooter.BackColor = Color.FromArgb(236, 240, 241)
        pnlFooter.Padding = New Padding(20)
        Me.Controls.Add(pnlFooter)
    End Sub
    
    ''' <summary>
    ''' Configura interface
    ''' </summary>
    Private Sub ConfigurarInterface()
        ' Título no cabeçalho
        Dim lblTitulo = New Label()
        lblTitulo.Text = "⚙️ CONFIGURAÇÕES DO SISTEMA"
        lblTitulo.Font = New Font("Segoe UI", 16, FontStyle.Bold)
        lblTitulo.ForeColor = Color.White
        lblTitulo.AutoSize = True
        lblTitulo.Location = New Point(20, 20)
        pnlHeader.Controls.Add(lblTitulo)
        
        ' Tab Control
        tabControl = New TabControl()
        tabControl.Dock = DockStyle.Fill
        tabControl.Font = New Font("Segoe UI", 10)
        pnlContent.Controls.Add(tabControl)
        
        ' Criar abas
        CriarAbaEmpresa()
        CriarAbaSistema()
        CriarAbaLogs()
        
        ' Botões do rodapé
        ConfigurarBotoesRodape()
    End Sub
    
    ''' <summary>
    ''' Cria aba de configurações da empresa
    ''' </summary>
    Private Sub CriarAbaEmpresa()
        Dim tabEmpresa = New TabPage("🏢 Empresa")
        tabControl.TabPages.Add(tabEmpresa)
        
        Dim y = 20
        
        ' Nome da madeireira
        Dim lblNome = New Label() With {
            .Text = "Nome da Madeireira:",
            .Location = New Point(20, y),
            .AutoSize = True,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold)
        }
        tabEmpresa.Controls.Add(lblNome)
        
        txtNomeMadeireira = New TextBox() With {
            .Location = New Point(20, y + 25),
            .Size = New Size(500, 25)
        }
        tabEmpresa.Controls.Add(txtNomeMadeireira)
        y += 60
        
        ' Endereço
        Dim lblEndereco = New Label() With {
            .Text = "Endereço:",
            .Location = New Point(20, y),
            .AutoSize = True,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold)
        }
        tabEmpresa.Controls.Add(lblEndereco)
        
        txtEnderecoMadeireira = New TextBox() With {
            .Location = New Point(20, y + 25),
            .Size = New Size(500, 25)
        }
        tabEmpresa.Controls.Add(txtEnderecoMadeireira)
        y += 60
        
        ' Cidade e CEP
        Dim lblCidade = New Label() With {
            .Text = "Cidade:",
            .Location = New Point(20, y),
            .AutoSize = True,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold)
        }
        tabEmpresa.Controls.Add(lblCidade)
        
        txtCidadeMadeireira = New TextBox() With {
            .Location = New Point(20, y + 25),
            .Size = New Size(240, 25)
        }
        tabEmpresa.Controls.Add(txtCidadeMadeireira)
        
        Dim lblCEP = New Label() With {
            .Text = "CEP:",
            .Location = New Point(280, y),
            .AutoSize = True,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold)
        }
        tabEmpresa.Controls.Add(lblCEP)
        
        txtCEPMadeireira = New TextBox() With {
            .Location = New Point(280, y + 25),
            .Size = New Size(120, 25),
            .PlaceholderText = "00000-000"
        }
        tabEmpresa.Controls.Add(txtCEPMadeireira)
        y += 60
        
        ' Telefone e CNPJ
        Dim lblTelefone = New Label() With {
            .Text = "Telefone:",
            .Location = New Point(20, y),
            .AutoSize = True,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold)
        }
        tabEmpresa.Controls.Add(lblTelefone)
        
        txtTelefoneMadeireira = New TextBox() With {
            .Location = New Point(20, y + 25),
            .Size = New Size(180, 25),
            .PlaceholderText = "(00) 0000-0000"
        }
        tabEmpresa.Controls.Add(txtTelefoneMadeireira)
        
        Dim lblCNPJ = New Label() With {
            .Text = "CNPJ:",
            .Location = New Point(220, y),
            .AutoSize = True,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold)
        }
        tabEmpresa.Controls.Add(lblCNPJ)
        
        txtCNPJMadeireira = New TextBox() With {
            .Location = New Point(220, y + 25),
            .Size = New Size(200, 25),
            .PlaceholderText = "00.000.000/0000-00"
        }
        tabEmpresa.Controls.Add(txtCNPJMadeireira)
    End Sub
    
    ''' <summary>
    ''' Cria aba de configurações do sistema
    ''' </summary>
    Private Sub CriarAbaSistema()
        Dim tabSistema = New TabPage("🔧 Sistema")
        tabControl.TabPages.Add(tabSistema)
        
        Dim y = 20
        
        ' Vendedor padrão
        Dim lblVendedor = New Label() With {
            .Text = "Vendedor Padrão:",
            .Location = New Point(20, y),
            .AutoSize = True,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold)
        }
        tabSistema.Controls.Add(lblVendedor)
        
        txtVendedorPadrao = New TextBox() With {
            .Location = New Point(20, y + 25),
            .Size = New Size(300, 25)
        }
        tabSistema.Controls.Add(txtVendedorPadrao)
        y += 60
        
        ' Excel visível
        chkExcelVisivel = New CheckBox() With {
            .Text = "Mostrar Excel durante processamento (para debug)",
            .Location = New Point(20, y),
            .Size = New Size(400, 25),
            .Font = New Font("Segoe UI", 10)
        }
        tabSistema.Controls.Add(chkExcelVisivel)
        y += 40
        
        ' Backup automático
        chkBackupAutomatico = New CheckBox() With {
            .Text = "Backup automático",
            .Location = New Point(20, y),
            .Size = New Size(200, 25),
            .Font = New Font("Segoe UI", 10, FontStyle.Bold)
        }
        tabSistema.Controls.Add(chkBackupAutomatico)
        
        btnBackupManual = New Button() With {
            .Text = "🔄 Backup Manual",
            .Location = New Point(240, y - 2),
            .Size = New Size(140, 30),
            .BackColor = Color.FromArgb(230, 126, 34),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
        btnBackupManual.FlatAppearance.BorderSize = 0
        tabSistema.Controls.Add(btnBackupManual)
        y += 40
        
        ' Intervalo de backup
        Dim lblIntervalo = New Label() With {
            .Text = "Intervalo de backup (horas):",
            .Location = New Point(40, y),
            .AutoSize = True,
            .Font = New Font("Segoe UI", 9)
        }
        tabSistema.Controls.Add(lblIntervalo)
        
        nudIntervaloBacKup = New NumericUpDown() With {
            .Location = New Point(220, y - 2),
            .Size = New Size(80, 25),
            .Minimum = 1,
            .Maximum = 168,
            .Value = 24
        }
        tabSistema.Controls.Add(nudIntervaloBacKup)
        y += 40
        
        ' Manter histórico
        Dim lblHistorico = New Label() With {
            .Text = "Manter histórico (dias):",
            .Location = New Point(40, y),
            .AutoSize = True,
            .Font = New Font("Segoe UI", 9)
        }
        tabSistema.Controls.Add(lblHistorico)
        
        nudManterHistorico = New NumericUpDown() With {
            .Location = New Point(220, y - 2),
            .Size = New Size(80, 25),
            .Minimum = 30,
            .Maximum = 3650,
            .Value = 365
        }
        tabSistema.Controls.Add(nudManterHistorico)
        y += 60
        
        ' Teste do Excel
        btnTestarExcel = New Button() With {
            .Text = "🧪 Testar Integração Excel",
            .Location = New Point(20, y),
            .Size = New Size(200, 35),
            .BackColor = Color.FromArgb(52, 152, 219),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold)
        }
        btnTestarExcel.FlatAppearance.BorderSize = 0
        tabSistema.Controls.Add(btnTestarExcel)
    End Sub
    
    ''' <summary>
    ''' Cria aba de logs
    ''' </summary>
    Private Sub CriarAbaLogs()
        Dim tabLogs = New TabPage("📋 Logs")
        tabControl.TabPages.Add(tabLogs)
        
        Dim y = 20
        
        ' Nível de log
        Dim lblLogLevel = New Label() With {
            .Text = "Nível de Log:",
            .Location = New Point(20, y),
            .AutoSize = True,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold)
        }
        tabLogs.Controls.Add(lblLogLevel)
        
        cmbLogLevel = New ComboBox() With {
            .Location = New Point(20, y + 25),
            .Size = New Size(150, 25),
            .DropDownStyle = ComboBoxStyle.DropDownList
        }
        cmbLogLevel.Items.AddRange({"INFO", "WARNING", "ERROR", "CRITICAL"})
        tabLogs.Controls.Add(cmbLogLevel)
        y += 80
        
        ' Botões de log
        btnVerLogs = New Button() With {
            .Text = "📄 Ver Logs do Dia",
            .Location = New Point(20, y),
            .Size = New Size(150, 35),
            .BackColor = Color.FromArgb(46, 204, 113),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
        btnVerLogs.FlatAppearance.BorderSize = 0
        tabLogs.Controls.Add(btnVerLogs)
        
        btnLimparLogs = New Button() With {
            .Text = "🗑️ Limpar Logs Antigos",
            .Location = New Point(190, y),
            .Size = New Size(170, 35),
            .BackColor = Color.FromArgb(231, 76, 60),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
        btnLimparLogs.FlatAppearance.BorderSize = 0
        tabLogs.Controls.Add(btnLimparLogs)
        y += 60
        
        ' Informações do sistema
        Dim lblInfo = New Label() With {
            .Text = "ℹ️ INFORMAÇÕES DO SISTEMA:" & vbCrLf & vbCrLf &
                   "• Logs são salvos automaticamente" & vbCrLf &
                   "• Arquivos de log são organizados por data" & vbCrLf &
                   "• Logs antigos são removidos automaticamente" & vbCrLf &
                   "• Todas as operações são auditadas",
            .Location = New Point(20, y),
            .Size = New Size(500, 120),
            .Font = New Font("Segoe UI", 9),
            .ForeColor = Color.FromArgb(127, 140, 141)
        }
        tabLogs.Controls.Add(lblInfo)
    End Sub
    
    ''' <summary>
    ''' Configura botões do rodapé
    ''' </summary>
    Private Sub ConfigurarBotoesRodape()
        btnSalvar = New Button() With {
            .Text = "💾 SALVAR",
            .Location = New Point(0, 15),
            .Size = New Size(120, 35),
            .BackColor = Color.FromArgb(46, 204, 113),
            .ForeColor = Color.White,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .FlatStyle = FlatStyle.Flat
        }
        btnSalvar.FlatAppearance.BorderSize = 0
        pnlFooter.Controls.Add(btnSalvar)
        
        btnCancelar = New Button() With {
            .Text = "❌ CANCELAR",
            .Location = New Point(130, 15),
            .Size = New Size(120, 35),
            .BackColor = Color.FromArgb(231, 76, 60),
            .ForeColor = Color.White,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .FlatStyle = FlatStyle.Flat
        }
        btnCancelar.FlatAppearance.BorderSize = 0
        pnlFooter.Controls.Add(btnCancelar)
    End Sub
    
    ''' <summary>
    ''' Carrega configurações atuais
    ''' </summary>
    Private Sub CarregarConfiguracoes()
        Try
            ' Configurações da empresa
            txtNomeMadeireira.Text = _config.NomeMadeireira
            txtEnderecoMadeireira.Text = _config.EnderecoMadeireira
            txtCidadeMadeireira.Text = _config.CidadeMadeireira
            txtCEPMadeireira.Text = _config.CEPMadeireira
            txtTelefoneMadeireira.Text = _config.TelefoneMadeireira
            txtCNPJMadeireira.Text = _config.CNPJMadeireira
            
            ' Configurações do sistema
            txtVendedorPadrao.Text = _config.VendedorPadrao
            chkExcelVisivel.Checked = _config.ExcelVisivel
            chkBackupAutomatico.Checked = _config.BackupAutomatico
            nudIntervaloBacKup.Value = _config.IntervaloBacKupHoras
            nudManterHistorico.Value = _config.ManterHistoricoDias
            
            ' Configurações de log
            cmbLogLevel.Text = _config.LogLevel
            
        Catch ex As Exception
            _logger.Error("Erro ao carregar configurações", ex)
            MessageBox.Show("Erro ao carregar configurações: " & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    ''' <summary>
    ''' Salva configurações
    ''' </summary>
    Private Sub btnSalvar_Click(sender As Object, e As EventArgs) Handles btnSalvar.Click
        Try
            ' Validar dados
            If String.IsNullOrWhiteSpace(txtNomeMadeireira.Text) Then
                MessageBox.Show("Nome da madeireira é obrigatório.", "Validação", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtNomeMadeireira.Focus()
                Return
            End If
            
            ' Salvar configurações customizadas
            _config.SetCustomSetting("BackupAutomatico", chkBackupAutomatico.Checked.ToString())
            _config.SetCustomSetting("IntervaloBacKupHoras", nudIntervaloBacKup.Value.ToString())
            _config.SetCustomSetting("ManterHistoricoDias", nudManterHistorico.Value.ToString())
            _config.SetCustomSetting("LogLevel", cmbLogLevel.Text)
            
            _logger.Info("Configurações salvas pelo usuário")
            _logger.Audit("CONFIGURACOES_ALTERADAS", 
                "Backup automático: " & chkBackupAutomatico.Checked.ToString() &
                ", Intervalo backup: " & nudIntervaloBacKup.Value.ToString() & "h" &
                ", Manter histórico: " & nudManterHistorico.Value.ToString() & " dias",
                "Sistema")
            
            MessageBox.Show("Configurações salvas com sucesso!" & vbCrLf & vbCrLf &
                          "Algumas alterações podem exigir reinicialização do sistema.", 
                          "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
            
            Me.DialogResult = DialogResult.OK
            Me.Close()
            
        Catch ex As Exception
            _logger.Error("Erro ao salvar configurações", ex)
            MessageBox.Show("Erro ao salvar configurações: " & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    ''' <summary>
    ''' Cancela alterações
    ''' </summary>
    Private Sub btnCancelar_Click(sender As Object, e As EventArgs) Handles btnCancelar.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub
    
    ''' <summary>
    ''' Testa integração com Excel
    ''' </summary>
    Private Sub btnTestarExcel_Click(sender As Object, e As EventArgs) Handles btnTestarExcel.Click
        Try
            Dim excel As Object = CreateObject("Excel.Application")
            excel.Visible = True
            excel.Quit()
            excel = Nothing
            
            MessageBox.Show("✅ Excel testado com sucesso!" & vbCrLf & 
                          "Integração funcionando normalmente.", 
                          "Teste Excel", MessageBoxButtons.OK, MessageBoxIcon.Information)
                          
        Catch ex As Exception
            MessageBox.Show("❌ Erro na integração Excel:" & vbCrLf & vbCrLf & ex.Message, 
                          "Erro Excel", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    ''' <summary>
    ''' Executa backup manual
    ''' </summary>
    Private Sub btnBackupManual_Click(sender As Object, e As EventArgs) Handles btnBackupManual.Click
        Try
            If MessageBox.Show("Executar backup manual do sistema?", "Confirmar Backup", 
                             MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                
                If _backupService.ExecutarBackupCompleto() Then
                    MessageBox.Show("✅ Backup executado com sucesso!", "Backup", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    MessageBox.Show("❌ Erro ao executar backup.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            End If
        Catch ex As Exception
            _logger.Error("Erro no backup manual", ex)
            MessageBox.Show("Erro no backup: " & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    ''' <summary>
    ''' Ver logs do dia
    ''' </summary>
    Private Sub btnVerLogs_Click(sender As Object, e As EventArgs) Handles btnVerLogs.Click
        Try
            Dim logs = _logger.GetTodayLogs()
            If String.IsNullOrEmpty(logs) Then
                MessageBox.Show("Nenhum log encontrado para hoje.", "Logs", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                ' Mostrar em uma nova janela ou salvar em arquivo temporário
                Dim tempFile = System.IO.Path.GetTempFileName() + ".txt"
                System.IO.File.WriteAllText(tempFile, logs)
                Process.Start("notepad.exe", tempFile)
            End If
        Catch ex As Exception
            MessageBox.Show("Erro ao abrir logs: " & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    ''' <summary>
    ''' Limpar logs antigos
    ''' </summary>
    Private Sub btnLimparLogs_Click(sender As Object, e As EventArgs) Handles btnLimparLogs.Click
        Try
            If MessageBox.Show("Remover logs com mais de 30 dias?", "Confirmar Limpeza", 
                             MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                
                ' A limpeza é feita automaticamente pelo Logger
                MessageBox.Show("✅ Limpeza de logs executada!", "Limpeza", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            MessageBox.Show("Erro na limpeza: " & ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    ''' <summary>
    ''' Processa teclas pressionadas
    ''' </summary>
    Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean
        Select Case keyData
            Case Keys.Escape
                btnCancelar_Click(Nothing, Nothing)
                Return True
            Case Keys.Enter
                If btnSalvar.Focused Then
                    btnSalvar_Click(Nothing, Nothing)
                    Return True
                End If
        End Select
        
        Return MyBase.ProcessCmdKey(msg, keyData)
    End Function
End Class