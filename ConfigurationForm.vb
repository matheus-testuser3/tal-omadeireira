Imports System.Windows.Forms
Imports System.Drawing

''' <summary>
''' Formulário para configurações do sistema
''' Interface visual para ajustar configurações sem editar App.config
''' </summary>
Public Class ConfigurationForm
    Inherits Form
    
    ' Controles da interface
    Private WithEvents tabControl As TabControl
    Private WithEvents tabEmpresa As TabPage
    Private WithEvents tabSistema As TabPage
    Private WithEvents tabAvancado As TabPage
    
    ' Campos da empresa
    Private WithEvents txtNomeEmpresa As TextBox
    Private WithEvents txtEnderecoEmpresa As TextBox
    Private WithEvents txtCidadeEmpresa As TextBox
    Private WithEvents txtCEPEmpresa As MaskedTextBox
    Private WithEvents txtTelefoneEmpresa As MaskedTextBox
    Private WithEvents txtCNPJEmpresa As MaskedTextBox
    Private WithEvents txtEmailEmpresa As TextBox
    
    ' Campos do sistema
    Private WithEvents txtVendedorPadrao As TextBox
    Private WithEvents chkExcelVisivel As CheckBox
    Private WithEvents chkSalvarTemporario As CheckBox
    Private WithEvents chkValidacaoTempoReal As CheckBox
    Private WithEvents chkAutoBackup As CheckBox
    Private WithEvents chkCacheEnabled As CheckBox
    Private WithEvents numTimeoutExcel As NumericUpDown
    Private WithEvents numCacheExpiration As NumericUpDown
    
    ' Botões
    Private WithEvents btnSalvar As Button
    Private WithEvents btnCancelar As Button
    Private WithEvents btnRestaurarPadrao As Button
    Private WithEvents btnTestar As Button
    
    ' Sistema
    Private ReadOnly _logger As LoggingSystem = LoggingSystem.Instance
    Private ReadOnly _config As EnhancedConfigurationManager = EnhancedConfigurationManager.Instance
    
    ''' <summary>
    ''' Construtor
    ''' </summary>
    Public Sub New()
        InitializeComponent()
        CarregarConfiguracoes()
        _logger.LogInfo("ConfigurationForm", "Formulário de configurações inicializado")
    End Sub
    
    ''' <summary>
    ''' Inicializa componentes da interface
    ''' </summary>
    Private Sub InitializeComponent()
        Me.Text = "⚙️ Configurações do Sistema"
        Me.Size = New Size(600, 500)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.BackColor = Color.FromArgb(236, 240, 241)
        Me.Font = New Font("Segoe UI", 9.0F)
        
        CriarTabControl()
        CriarTabEmpresa()
        CriarTabSistema()
        CriarTabAvancado()
        CriarBotoes()
    End Sub
    
    ''' <summary>
    ''' Cria controle de abas
    ''' </summary>
    Private Sub CriarTabControl()
        tabControl = New TabControl()
        tabControl.Size = New Size(560, 380)
        tabControl.Location = New Point(20, 20)
        tabControl.Font = New Font("Segoe UI", 9.0F)
        Me.Controls.Add(tabControl)
    End Sub
    
    ''' <summary>
    ''' Cria aba de configurações da empresa
    ''' </summary>
    Private Sub CriarTabEmpresa()
        tabEmpresa = New TabPage("🏪 Empresa")
        tabEmpresa.BackColor = Color.White
        tabControl.TabPages.Add(tabEmpresa)
        
        Dim y = 20
        
        ' Nome da empresa
        CriarCampoTexto(tabEmpresa, "Nome da Empresa:", txtNomeEmpresa, 20, y, 500)
        y += 40
        
        ' Endereço
        CriarCampoTexto(tabEmpresa, "Endereço:", txtEnderecoEmpresa, 20, y, 500)
        y += 40
        
        ' Cidade
        CriarCampoTexto(tabEmpresa, "Cidade:", txtCidadeEmpresa, 20, y, 300)
        y += 40
        
        ' CEP
        CriarCampoMascarado(tabEmpresa, "CEP:", txtCEPEmpresa, "00000-000", 20, y, 100)
        y += 40
        
        ' Telefone
        CriarCampoMascarado(tabEmpresa, "Telefone:", txtTelefoneEmpresa, "(00) 0000-0000", 20, y, 150)
        y += 40
        
        ' CNPJ
        CriarCampoMascarado(tabEmpresa, "CNPJ:", txtCNPJEmpresa, "00.000.000/0000-00", 20, y, 180)
        y += 40
        
        ' Email
        CriarCampoTexto(tabEmpresa, "Email:", txtEmailEmpresa, 20, y, 300)
    End Sub
    
    ''' <summary>
    ''' Cria aba de configurações do sistema
    ''' </summary>
    Private Sub CriarTabSistema()
        tabSistema = New TabPage("🔧 Sistema")
        tabSistema.BackColor = Color.White
        tabControl.TabPages.Add(tabSistema)
        
        Dim y = 20
        
        ' Vendedor padrão
        CriarCampoTexto(tabSistema, "Vendedor Padrão:", txtVendedorPadrao, 20, y, 200)
        y += 40
        
        ' Timeout Excel
        CriarCampoNumerico(tabSistema, "Timeout Excel (ms):", numTimeoutExcel, 20, y, 100, 5000, 120000)
        y += 40
        
        ' Expiração Cache
        CriarCampoNumerico(tabSistema, "Cache Expira (min):", numCacheExpiration, 20, y, 100, 5, 480)
        y += 60
        
        ' Checkboxes
        chkExcelVisivel = CriarCheckBox(tabSistema, "Excel visível durante processamento", 20, y)
        y += 30
        
        chkSalvarTemporario = CriarCheckBox(tabSistema, "Salvar talão temporário", 20, y)
        y += 30
        
        chkValidacaoTempoReal = CriarCheckBox(tabSistema, "Validação em tempo real", 20, y)
        y += 30
        
        chkAutoBackup = CriarCheckBox(tabSistema, "Backup automático", 20, y)
        y += 30
        
        chkCacheEnabled = CriarCheckBox(tabSistema, "Cache habilitado", 20, y)
    End Sub
    
    ''' <summary>
    ''' Cria aba de configurações avançadas
    ''' </summary>
    Private Sub CriarTabAvancado()
        tabAvancado = New TabPage("⚡ Avançado")
        tabAvancado.BackColor = Color.White
        tabControl.TabPages.Add(tabAvancado)
        
        ' Informações do sistema
        Dim lblInfo As New Label()
        lblInfo.Text = "ℹ️ Informações do Sistema"
        lblInfo.Font = New Font("Segoe UI", 12.0F, FontStyle.Bold)
        lblInfo.Location = New Point(20, 20)
        lblInfo.Size = New Size(300, 25)
        tabAvancado.Controls.Add(lblInfo)
        
        Dim txtInfo As New TextBox()
        txtInfo.Multiline = True
        txtInfo.ReadOnly = True
        txtInfo.ScrollBars = ScrollBars.Vertical
        txtInfo.Location = New Point(20, 50)
        txtInfo.Size = New Size(500, 200)
        txtInfo.Text = ObterInformacoesSistema()
        tabAvancado.Controls.Add(txtInfo)
        
        ' Botão de teste
        btnTestar = New Button()
        btnTestar.Text = "🧪 Testar Configurações"
        btnTestar.Location = New Point(20, 270)
        btnTestar.Size = New Size(150, 30)
        btnTestar.BackColor = Color.FromArgb(241, 196, 15)
        btnTestar.ForeColor = Color.White
        btnTestar.FlatStyle = FlatStyle.Flat
        btnTestar.FlatAppearance.BorderSize = 0
        tabAvancado.Controls.Add(btnTestar)
    End Sub
    
    ''' <summary>
    ''' Cria botões principais
    ''' </summary>
    Private Sub CriarBotoes()
        Dim pnlBotoes As New Panel()
        pnlBotoes.Size = New Size(560, 50)
        pnlBotoes.Location = New Point(20, 410)
        pnlBotoes.BackColor = Color.Transparent
        Me.Controls.Add(pnlBotoes)
        
        btnRestaurarPadrao = New Button()
        btnRestaurarPadrao.Text = "🔄 Restaurar Padrão"
        btnRestaurarPadrao.Location = New Point(0, 10)
        btnRestaurarPadrao.Size = New Size(120, 30)
        btnRestaurarPadrao.BackColor = Color.FromArgb(149, 165, 166)
        btnRestaurarPadrao.ForeColor = Color.White
        btnRestaurarPadrao.FlatStyle = FlatStyle.Flat
        btnRestaurarPadrao.FlatAppearance.BorderSize = 0
        pnlBotoes.Controls.Add(btnRestaurarPadrao)
        
        btnCancelar = New Button()
        btnCancelar.Text = "❌ Cancelar"
        btnCancelar.Location = New Point(320, 10)
        btnCancelar.Size = New Size(100, 30)
        btnCancelar.BackColor = Color.FromArgb(231, 76, 60)
        btnCancelar.ForeColor = Color.White
        btnCancelar.FlatStyle = FlatStyle.Flat
        btnCancelar.FlatAppearance.BorderSize = 0
        btnCancelar.DialogResult = DialogResult.Cancel
        pnlBotoes.Controls.Add(btnCancelar)
        
        btnSalvar = New Button()
        btnSalvar.Text = "💾 Salvar"
        btnSalvar.Location = New Point(440, 10)
        btnSalvar.Size = New Size(100, 30)
        btnSalvar.BackColor = Color.FromArgb(46, 204, 113)
        btnSalvar.ForeColor = Color.White
        btnSalvar.FlatStyle = FlatStyle.Flat
        btnSalvar.FlatAppearance.BorderSize = 0
        btnSalvar.DialogResult = DialogResult.OK
        pnlBotoes.Controls.Add(btnSalvar)
    End Sub
    
    #Region "Métodos Auxiliares"
    
    ''' <summary>
    ''' Cria campo de texto
    ''' </summary>
    Private Sub CriarCampoTexto(parent As Control, labelText As String, ByRef textBox As TextBox, x As Integer, y As Integer, width As Integer)
        Dim lbl As New Label()
        lbl.Text = labelText
        lbl.Location = New Point(x, y + 2)
        lbl.Size = New Size(120, 20)
        parent.Controls.Add(lbl)
        
        textBox = New TextBox()
        textBox.Location = New Point(x + 130, y)
        textBox.Size = New Size(width, 23)
        parent.Controls.Add(textBox)
    End Sub
    
    ''' <summary>
    ''' Cria campo mascarado
    ''' </summary>
    Private Sub CriarCampoMascarado(parent As Control, labelText As String, ByRef maskedTextBox As MaskedTextBox, mask As String, x As Integer, y As Integer, width As Integer)
        Dim lbl As New Label()
        lbl.Text = labelText
        lbl.Location = New Point(x, y + 2)
        lbl.Size = New Size(120, 20)
        parent.Controls.Add(lbl)
        
        maskedTextBox = New MaskedTextBox()
        maskedTextBox.Mask = mask
        maskedTextBox.Location = New Point(x + 130, y)
        maskedTextBox.Size = New Size(width, 23)
        parent.Controls.Add(maskedTextBox)
    End Sub
    
    ''' <summary>
    ''' Cria campo numérico
    ''' </summary>
    Private Sub CriarCampoNumerico(parent As Control, labelText As String, ByRef numericUpDown As NumericUpDown, x As Integer, y As Integer, width As Integer, min As Decimal, max As Decimal)
        Dim lbl As New Label()
        lbl.Text = labelText
        lbl.Location = New Point(x, y + 2)
        lbl.Size = New Size(120, 20)
        parent.Controls.Add(lbl)
        
        numericUpDown = New NumericUpDown()
        numericUpDown.Location = New Point(x + 130, y)
        numericUpDown.Size = New Size(width, 23)
        numericUpDown.Minimum = min
        numericUpDown.Maximum = max
        parent.Controls.Add(numericUpDown)
    End Sub
    
    ''' <summary>
    ''' Cria checkbox
    ''' </summary>
    Private Function CriarCheckBox(parent As Control, text As String, x As Integer, y As Integer) As CheckBox
        Dim chk As New CheckBox()
        chk.Text = text
        chk.Location = New Point(x, y)
        chk.Size = New Size(300, 20)
        chk.UseVisualStyleBackColor = True
        parent.Controls.Add(chk)
        Return chk
    End Function
    
    #End Region
    
    #Region "Carregamento e Salvamento"
    
    ''' <summary>
    ''' Carrega configurações atuais
    ''' </summary>
    Private Sub CarregarConfiguracoes()
        Try
            ' Dados da empresa
            txtNomeEmpresa.Text = _config.NomeMadeireira
            txtEnderecoEmpresa.Text = _config.EnderecoMadeireira
            txtCidadeEmpresa.Text = _config.CidadeMadeireira
            txtCEPEmpresa.Text = _config.CEPMadeireira
            txtTelefoneEmpresa.Text = _config.TelefoneMadeireira
            txtCNPJEmpresa.Text = _config.CNPJMadeireira
            txtEmailEmpresa.Text = _config.GetConfigValuePublic("EmailMadeireira", "")
            
            ' Configurações do sistema
            txtVendedorPadrao.Text = _config.VendedorPadrao
            numTimeoutExcel.Value = _config.TimeoutExcel
            numCacheExpiration.Value = _config.CacheExpirationMinutes
            
            chkExcelVisivel.Checked = _config.ExcelVisivel
            chkSalvarTemporario.Checked = _config.SalvarTalaoTemporario
            chkValidacaoTempoReal.Checked = _config.GetConfigValuePublic("ValidacaoTempoReal", True)
            chkAutoBackup.Checked = _config.AutoBackupEnabled
            chkCacheEnabled.Checked = _config.CacheEnabled
            
        Catch ex As Exception
            _logger.LogError("ConfigurationForm", "Erro ao carregar configurações", ex)
            MessageBox.Show($"Erro ao carregar configurações: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    ''' <summary>
    ''' Obtém informações do sistema
    ''' </summary>
    Private Function ObterInformacoesSistema() As String
        Try
            Dim info As New System.Text.StringBuilder()
            info.AppendLine($"Sistema Operacional: {Environment.OSVersion}")
            info.AppendLine($".NET Framework: {Environment.Version}")
            info.AppendLine($"Memória RAM: {GC.GetTotalMemory(False) / 1024 / 1024:N0} MB")
            info.AppendLine($"Processadores: {Environment.ProcessorCount}")
            info.AppendLine($"Usuário: {Environment.UserName}")
            info.AppendLine($"Máquina: {Environment.MachineName}")
            info.AppendLine()
            
            ' Validação de pré-requisitos
            Dim prereq = ExcelAutomationFactory.ValidatePrerequisites()
            info.AppendLine($"Pré-requisitos Excel: {If(prereq.IsValid, "✅ OK", "❌ " & prereq.ErrorMessage)}")
            
            ' Estatísticas de cache
            Dim cacheStats = CacheManager.Instance.GetStatistics()
            info.AppendLine($"Cache: {cacheStats}")
            
            Return info.ToString()
            
        Catch ex As Exception
            Return $"Erro ao obter informações: {ex.Message}"
        End Try
    End Function
    
    #End Region
    
    #Region "Eventos"
    
    ''' <summary>
    ''' Salvar configurações
    ''' </summary>
    Private Sub btnSalvar_Click(sender As Object, e As EventArgs) Handles btnSalvar.Click
        Try
            ' TODO: Implementar salvamento real no App.config
            ' Por enquanto, apenas mostrar mensagem
            MessageBox.Show("💡 Funcionalidade em desenvolvimento." & Environment.NewLine & 
                          "As configurações serão salvas em futuras versões.", 
                          "Em Desenvolvimento", MessageBoxButtons.OK, MessageBoxIcon.Information)
            
            _logger.LogInfo("ConfigurationForm", "Configurações salvas (simulado)")
            
        Catch ex As Exception
            _logger.LogError("ConfigurationForm", "Erro ao salvar configurações", ex)
            MessageBox.Show($"Erro ao salvar: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    ''' <summary>
    ''' Restaurar configurações padrão
    ''' </summary>
    Private Sub btnRestaurarPadrao_Click(sender As Object, e As EventArgs) Handles btnRestaurarPadrao.Click
        Dim resultado = MessageBox.Show("Isso irá restaurar todas as configurações para os valores padrão. Confirma?", 
                                       "Restaurar Padrão", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If resultado = DialogResult.Yes Then
            RestaurarConfiguracoesPadrao()
        End If
    End Sub
    
    ''' <summary>
    ''' Testar configurações
    ''' </summary>
    Private Sub btnTestar_Click(sender As Object, e As EventArgs) Handles btnTestar.Click
        Try
            ' Teste de validação de pré-requisitos
            Dim prereq = ExcelAutomationFactory.ValidatePrerequisites()
            
            ' Teste de validação de configurações
            Dim configValid = _config.ValidateConfigurations()
            
            Dim mensagem = $"🧪 Resultado dos Testes:{Environment.NewLine}{Environment.NewLine}" &
                          $"Pré-requisitos Excel: {If(prereq.IsValid, "✅ OK", "❌ " & prereq.ErrorMessage)}{Environment.NewLine}" &
                          $"Configurações: {If(configValid.IsValid, "✅ OK", "❌ " & configValid.ErrorMessage)}{Environment.NewLine}"
            
            MessageBox.Show(mensagem, "Teste de Configurações", MessageBoxButtons.OK, 
                          If(prereq.IsValid And configValid.IsValid, MessageBoxIcon.Information, MessageBoxIcon.Warning))
            
            _logger.LogInfo("ConfigurationForm", "Teste de configurações executado")
            
        Catch ex As Exception
            _logger.LogError("ConfigurationForm", "Erro ao testar configurações", ex)
            MessageBox.Show($"Erro no teste: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    ''' <summary>
    ''' Restaura configurações padrão
    ''' </summary>
    Private Sub RestaurarConfiguracoesPadrao()
        txtNomeEmpresa.Text = "Madeireira Maria Luiza"
        txtEnderecoEmpresa.Text = "Rua Principal, 123 - Centro"
        txtCidadeEmpresa.Text = "Paulista/PE"
        txtCEPEmpresa.Text = "53401-445"
        txtTelefoneEmpresa.Text = "(81) 3436-1234"
        txtCNPJEmpresa.Text = "12.345.678/0001-90"
        txtEmailEmpresa.Text = "contato@madeireiramaria.com.br"
        
        txtVendedorPadrao.Text = "Sistema"
        numTimeoutExcel.Value = 30000
        numCacheExpiration.Value = 60
        
        chkExcelVisivel.Checked = False
        chkSalvarTemporario.Checked = False
        chkValidacaoTempoReal.Checked = True
        chkAutoBackup.Checked = True
        chkCacheEnabled.Checked = True
    End Sub
    
    #End Region
End Class