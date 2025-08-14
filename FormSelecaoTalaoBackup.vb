''' <summary>
''' Formulário de seleção de talões importados do backup - Madeireira Maria Luiza
''' Data/Hora: 2025-08-14 11:16:26 UTC
''' Usuário: matheus-testuser3
''' Sistema de Backup e Restauração de Talões
''' </summary>

Imports System.Windows.Forms
Imports System.Drawing
Imports System.ComponentModel

''' <summary>
''' Interface para listar e selecionar talões importados de backup
''' DataGridView com informações detalhadas e seleção por duplo clique
''' </summary>
Public Class FormSelecaoTalaoBackup
    Inherits Form
    
    ' === CONTROLES DA INTERFACE ===
    Private WithEvents dgvTaloes As DataGridView
    Private WithEvents btnSelecionar As Button
    Private WithEvents btnAtualizar As Button
    Private WithEvents btnCancelar As Button
    Private WithEvents lblTitulo As Label
    Private WithEvents lblSubtitulo As Label
    Private WithEvents pnlHeader As Panel
    Private WithEvents pnlBotoes As Panel
    Private WithEvents pnlMain As Panel
    
    ' === DADOS ===
    Private taloes As List(Of DadosTalaoMadeireira)
    Private talaoSelecionado As DadosTalaoMadeireira
    
    ' === PROPRIEDADES ===
    Public ReadOnly Property TalaoSelecionado As DadosTalaoMadeireira
        Get
            Return talaoSelecionado
        End Get
    End Property
    
    ''' <summary>
    ''' Construtor do formulário de seleção
    ''' </summary>
    Public Sub New(taloesImportados As List(Of DadosTalaoMadeireira))
        taloes = taloesImportados
        InitializeComponent()
        ConfigurarInterface()
        CarregarDados()
    End Sub
    
    ''' <summary>
    ''' Inicializa os componentes da interface
    ''' </summary>
    Private Sub InitializeComponent()
        ' Configurações do formulário
        Me.Text = "Seleção de Talão - Madeireira Maria Luiza"
        Me.Size = New Size(1000, 700)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.BackColor = Color.WhiteSmoke
        Me.Font = New Font("Segoe UI", 9.0F, FontStyle.Regular)
        
        ' Painel de cabeçalho
        pnlHeader = New Panel()
        pnlHeader.Size = New Size(Me.ClientSize.Width, 80)
        pnlHeader.Dock = DockStyle.Top
        pnlHeader.BackColor = Color.FromArgb(34, 139, 34) ' Verde madeira
        
        ' Título
        lblTitulo = New Label()
        lblTitulo.Text = "Selecionar Talão para Geração"
        lblTitulo.Font = New Font("Segoe UI", 16.0F, FontStyle.Bold)
        lblTitulo.ForeColor = Color.White
        lblTitulo.Location = New Point(20, 15)
        lblTitulo.AutoSize = True
        
        ' Subtítulo
        lblSubtitulo = New Label()
        lblSubtitulo.Text = "Duplo clique no talão desejado ou selecione e clique em 'Selecionar'"
        lblSubtitulo.Font = New Font("Segoe UI", 10.0F, FontStyle.Regular)
        lblSubtitulo.ForeColor = Color.LightGray
        lblSubtitulo.Location = New Point(20, 45)
        lblSubtitulo.AutoSize = True
        
        pnlHeader.Controls.AddRange({lblTitulo, lblSubtitulo})
        
        ' Painel principal
        pnlMain = New Panel()
        pnlMain.Dock = DockStyle.Fill
        pnlMain.Padding = New Padding(20)
        pnlMain.BackColor = Color.White
        
        ' DataGridView
        dgvTaloes = New DataGridView()
        dgvTaloes.Dock = DockStyle.Fill
        dgvTaloes.AllowUserToAddRows = False
        dgvTaloes.AllowUserToDeleteRows = False
        dgvTaloes.ReadOnly = True
        dgvTaloes.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvTaloes.MultiSelect = False
        dgvTaloes.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgvTaloes.RowHeadersVisible = False
        dgvTaloes.BackgroundColor = Color.White
        dgvTaloes.BorderStyle = BorderStyle.None
        dgvTaloes.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal
        dgvTaloes.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None
        dgvTaloes.EnableHeadersVisualStyles = False
        
        ' Estilo do cabeçalho
        dgvTaloes.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(34, 139, 34)
        dgvTaloes.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        dgvTaloes.ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI", 9.0F, FontStyle.Bold)
        dgvTaloes.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgvTaloes.ColumnHeadersHeight = 35
        
        ' Estilo das linhas
        dgvTaloes.DefaultCellStyle.BackColor = Color.White
        dgvTaloes.DefaultCellStyle.ForeColor = Color.Black
        dgvTaloes.DefaultCellStyle.SelectionBackColor = Color.FromArgb(144, 238, 144) ' Verde claro
        dgvTaloes.DefaultCellStyle.SelectionForeColor = Color.Black
        dgvTaloes.DefaultCellStyle.Font = New Font("Segoe UI", 9.0F)
        dgvTaloes.DefaultCellStyle.Padding = New Padding(5)
        dgvTaloes.RowTemplate.Height = 30
        
        ' Estilo das linhas alternadas
        dgvTaloes.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(248, 248, 248)
        
        pnlMain.Controls.Add(dgvTaloes)
        
        ' Painel de botões
        pnlBotoes = New Panel()
        pnlBotoes.Size = New Size(Me.ClientSize.Width, 70)
        pnlBotoes.Dock = DockStyle.Bottom
        pnlBotoes.BackColor = Color.FromArgb(245, 245, 245)
        pnlBotoes.Padding = New Padding(20, 15, 20, 15)
        
        ' Botão Selecionar
        btnSelecionar = New Button()
        btnSelecionar.Text = "Selecionar Talão"
        btnSelecionar.Size = New Size(130, 40)
        btnSelecionar.Location = New Point(pnlBotoes.Width - 410, 15)
        btnSelecionar.BackColor = Color.FromArgb(34, 139, 34)
        btnSelecionar.ForeColor = Color.White
        btnSelecionar.Font = New Font("Segoe UI", 9.0F, FontStyle.Bold)
        btnSelecionar.FlatStyle = FlatStyle.Flat
        btnSelecionar.FlatAppearance.BorderSize = 0
        btnSelecionar.Cursor = Cursors.Hand
        btnSelecionar.Enabled = False
        
        ' Botão Atualizar
        btnAtualizar = New Button()
        btnAtualizar.Text = "Atualizar"
        btnAtualizar.Size = New Size(100, 40)
        btnAtualizar.Location = New Point(pnlBotoes.Width - 270, 15)
        btnAtualizar.BackColor = Color.FromArgb(70, 130, 180) ' Azul aço
        btnAtualizar.ForeColor = Color.White
        btnAtualizar.Font = New Font("Segoe UI", 9.0F, FontStyle.Bold)
        btnAtualizar.FlatStyle = FlatStyle.Flat
        btnAtualizar.FlatAppearance.BorderSize = 0
        btnAtualizar.Cursor = Cursors.Hand
        
        ' Botão Cancelar
        btnCancelar = New Button()
        btnCancelar.Text = "Cancelar"
        btnCancelar.Size = New Size(100, 40)
        btnCancelar.Location = New Point(pnlBotoes.Width - 160, 15)
        btnCancelar.BackColor = Color.FromArgb(220, 53, 69) ' Vermelho
        btnCancelar.ForeColor = Color.White
        btnCancelar.Font = New Font("Segoe UI", 9.0F, FontStyle.Bold)
        btnCancelar.FlatStyle = FlatStyle.Flat
        btnCancelar.FlatAppearance.BorderSize = 0
        btnCancelar.Cursor = Cursors.Hand
        
        pnlBotoes.Controls.AddRange({btnSelecionar, btnAtualizar, btnCancelar})
        
        ' Adicionar controles ao formulário
        Me.Controls.AddRange({pnlMain, pnlHeader, pnlBotoes})
    End Sub
    
    ''' <summary>
    ''' Configura eventos e comportamentos da interface
    ''' </summary>
    Private Sub ConfigurarInterface()
        ' Eventos de redimensionamento
        AddHandler Me.Resize, AddressOf FormSelecaoTalaoBackup_Resize
        
        ' Aplicar efeitos visuais aos botões
        AplicarEfeitosBotoes()
    End Sub
    
    ''' <summary>
    ''' Carrega os dados dos talões na grid
    ''' </summary>
    Private Sub CarregarDados()
        Try
            Debug.WriteLine($"[SELECAO-BACKUP] Carregando {taloes.Count} talões na grid")
            
            ' Configurar colunas da grid
            dgvTaloes.Columns.Clear()
            
            ' Coluna Número do Talão
            Dim colNumero As New DataGridViewTextBoxColumn()
            colNumero.Name = "Numero"
            colNumero.HeaderText = "Nº Talão"
            colNumero.Width = 100
            colNumero.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvTaloes.Columns.Add(colNumero)
            
            ' Coluna Cliente
            Dim colCliente As New DataGridViewTextBoxColumn()
            colCliente.Name = "Cliente"
            colCliente.HeaderText = "Cliente"
            colCliente.Width = 200
            dgvTaloes.Columns.Add(colCliente)
            
            ' Coluna Data Emissão
            Dim colData As New DataGridViewTextBoxColumn()
            colData.Name = "DataEmissao"
            colData.HeaderText = "Data Emissão"
            colData.Width = 120
            colData.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            colData.DefaultCellStyle.Format = "dd/MM/yyyy"
            dgvTaloes.Columns.Add(colData)
            
            ' Coluna Quantidade de Produtos
            Dim colQtdProdutos As New DataGridViewTextBoxColumn()
            colQtdProdutos.Name = "QtdProdutos"
            colQtdProdutos.HeaderText = "Produtos"
            colQtdProdutos.Width = 80
            colQtdProdutos.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvTaloes.Columns.Add(colQtdProdutos)
            
            ' Coluna Valor Total
            Dim colTotal As New DataGridViewTextBoxColumn()
            colTotal.Name = "ValorTotal"
            colTotal.HeaderText = "Valor Total"
            colTotal.Width = 120
            colTotal.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            colTotal.DefaultCellStyle.Format = "C2"
            dgvTaloes.Columns.Add(colTotal)
            
            ' Coluna Formato Detectado
            Dim colFormato As New DataGridViewTextBoxColumn()
            colFormato.Name = "Formato"
            colFormato.HeaderText = "Formato"
            colFormato.Width = 100
            colFormato.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvTaloes.Columns.Add(colFormato)
            
            ' Coluna Vendedor
            Dim colVendedor As New DataGridViewTextBoxColumn()
            colVendedor.Name = "Vendedor"
            colVendedor.HeaderText = "Vendedor"
            colVendedor.Width = 120
            dgvTaloes.Columns.Add(colVendedor)
            
            ' Carregar dados
            For Each talao In taloes
                Dim linha As DataGridViewRow = dgvTaloes.Rows(dgvTaloes.Rows.Add())
                linha.Tag = talao ' Armazenar referência ao objeto
                
                linha.Cells("Numero").Value = talao.NumeroTalao
                linha.Cells("Cliente").Value = talao.NomeCliente
                linha.Cells("DataEmissao").Value = talao.DataEmissao
                linha.Cells("QtdProdutos").Value = talao.QuantidadeTotalProdutos
                linha.Cells("ValorTotal").Value = talao.ValorTotal
                linha.Cells("Formato").Value = talao.FormatoDetectado
                linha.Cells("Vendedor").Value = talao.Vendedor
            Next
            
            ' Atualizar subtítulo com contagem
            lblSubtitulo.Text = $"Total de {taloes.Count} talões importados - Duplo clique para selecionar"
            
            Debug.WriteLine($"[SELECAO-BACKUP] Dados carregados com sucesso na grid")
            
        Catch ex As Exception
            Debug.WriteLine($"[SELECAO-BACKUP] ERRO ao carregar dados: {ex.Message}")
            MessageBox.Show($"Erro ao carregar dados dos talões:{vbCrLf}{ex.Message}",
                          "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    ''' <summary>
    ''' Aplica efeitos visuais aos botões
    ''' </summary>
    Private Sub AplicarEfeitosBotoes()
        For Each btn As Button In {btnSelecionar, btnAtualizar, btnCancelar}
            AddHandler btn.MouseEnter, Sub(sender, e)
                                          Dim button = CType(sender, Button)
                                          button.BackColor = AjustarBrilho(button.BackColor, 0.2F)
                                      End Sub
            
            AddHandler btn.MouseLeave, Sub(sender, e)
                                          Dim button = CType(sender, Button)
                                          Select Case button.Name
                                              Case "btnSelecionar"
                                                  button.BackColor = Color.FromArgb(34, 139, 34)
                                              Case "btnAtualizar"
                                                  button.BackColor = Color.FromArgb(70, 130, 180)
                                              Case "btnCancelar"
                                                  button.BackColor = Color.FromArgb(220, 53, 69)
                                          End Select
                                      End Sub
        Next
    End Sub
    
    ''' <summary>
    ''' Ajusta o brilho de uma cor
    ''' </summary>
    Private Function AjustarBrilho(cor As Color, fator As Single) As Color
        Dim r = Math.Min(255, CInt(cor.R * (1 + fator)))
        Dim g = Math.Min(255, CInt(cor.G * (1 + fator)))
        Dim b = Math.Min(255, CInt(cor.B * (1 + fator)))
        Return Color.FromArgb(r, g, b)
    End Function
    
    ' === EVENTOS ===
    
    ''' <summary>
    ''' Evento de seleção alterada na grid
    ''' </summary>
    Private Sub dgvTaloes_SelectionChanged(sender As Object, e As EventArgs) Handles dgvTaloes.SelectionChanged
        btnSelecionar.Enabled = (dgvTaloes.SelectedRows.Count > 0)
    End Sub
    
    ''' <summary>
    ''' Evento de duplo clique na grid - seleção direta
    ''' </summary>
    Private Sub dgvTaloes_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvTaloes.CellDoubleClick
        If e.RowIndex >= 0 Then
            SelecionarTalao()
        End If
    End Sub
    
    ''' <summary>
    ''' Evento do botão Selecionar
    ''' </summary>
    Private Sub btnSelecionar_Click(sender As Object, e As EventArgs) Handles btnSelecionar.Click
        SelecionarTalao()
    End Sub
    
    ''' <summary>
    ''' Evento do botão Atualizar
    ''' </summary>
    Private Sub btnAtualizar_Click(sender As Object, e As EventArgs) Handles btnAtualizar.Click
        CarregarDados()
        MessageBox.Show("Dados atualizados com sucesso!", "Atualização", 
                       MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
    
    ''' <summary>
    ''' Evento do botão Cancelar
    ''' </summary>
    Private Sub btnCancelar_Click(sender As Object, e As EventArgs) Handles btnCancelar.Click
        talaoSelecionado = Nothing
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub
    
    ''' <summary>
    ''' Evento de redimensionamento do formulário
    ''' </summary>
    Private Sub FormSelecaoTalaoBackup_Resize(sender As Object, e As EventArgs)
        If pnlBotoes IsNot Nothing Then
            ' Reposicionar botões
            btnCancelar.Location = New Point(pnlBotoes.Width - 140, 15)
            btnAtualizar.Location = New Point(pnlBotoes.Width - 250, 15)
            btnSelecionar.Location = New Point(pnlBotoes.Width - 390, 15)
        End If
    End Sub
    
    ''' <summary>
    ''' Realiza a seleção do talão
    ''' </summary>
    Private Sub SelecionarTalao()
        Try
            If dgvTaloes.SelectedRows.Count > 0 Then
                talaoSelecionado = CType(dgvTaloes.SelectedRows(0).Tag, DadosTalaoMadeireira)
                
                Debug.WriteLine($"[SELECAO-BACKUP] Talão selecionado: {talaoSelecionado.NumeroTalao}")
                
                Me.DialogResult = DialogResult.OK
                Me.Close()
            Else
                MessageBox.Show("Por favor, selecione um talão.", "Seleção Necessária",
                              MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
            
        Catch ex As Exception
            Debug.WriteLine($"[SELECAO-BACKUP] ERRO na seleção: {ex.Message}")
            MessageBox.Show($"Erro ao selecionar talão:{vbCrLf}{ex.Message}",
                          "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
End Class