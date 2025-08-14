Imports System.Windows.Forms
Imports System.Drawing

''' <summary>
''' Formulário de relatórios e consultas
''' Interface para visualização de histórico de vendas e relatórios
''' </summary>
Public Class RelatoriosForm
    Inherits Form
    
    ' Controles da interface
    Private WithEvents pnlHeader As Panel
    Private WithEvents pnlFilters As Panel
    Private WithEvents pnlContent As Panel
    Private WithEvents pnlFooter As Panel
    
    ' Filtros
    Private WithEvents dtpDataInicio As DateTimePicker
    Private WithEvents dtpDataFim As DateTimePicker
    Private WithEvents txtCliente As TextBox
    Private WithEvents txtVendedor As TextBox
    Private WithEvents btnFiltrar As Button
    Private WithEvents btnLimpar As Button
    
    ' Resultado
    Private WithEvents dgvVendas As DataGridView
    Private WithEvents lblTotalVendas As Label
    Private WithEvents lblValorTotal As Label
    Private WithEvents lblTicketMedio As Label
    
    ' Botões de ação
    Private WithEvents btnReimprimir As Button
    Private WithEvents btnExportar As Button
    Private WithEvents btnFechar As Button
    
    ' Serviços
    Private ReadOnly _historicoManager As HistoricoManager
    Private ReadOnly _vendaService As VendaService
    Private ReadOnly _logger As Logger
    
    ''' <summary>
    ''' Construtor
    ''' </summary>
    Public Sub New()
        _historicoManager = HistoricoManager.Instance
        _vendaService = New VendaService()
        _logger = Logger.Instance
        
        InitializeComponent()
        ConfigurarInterface()
        CarregarDadosIniciais()
    End Sub
    
    ''' <summary>
    ''' Inicializa componentes
    ''' </summary>
    Private Sub InitializeComponent()
        ' Configurar formulário
        Me.Text = "📊 Relatórios e Consultas - Sistema PDV"
        Me.Size = New Size(1200, 800)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.Sizable
        Me.MinimumSize = New Size(1000, 600)
        
        ' Painel de cabeçalho
        pnlHeader = New Panel()
        pnlHeader.Height = 80
        pnlHeader.Dock = DockStyle.Top
        pnlHeader.BackColor = Color.FromArgb(52, 73, 94)
        Me.Controls.Add(pnlHeader)
        
        ' Painel de filtros
        pnlFilters = New Panel()
        pnlFilters.Height = 120
        pnlFilters.Dock = DockStyle.Top
        pnlFilters.BackColor = Color.FromArgb(236, 240, 241)
        pnlFilters.Padding = New Padding(20)
        Me.Controls.Add(pnlFilters)
        
        ' Painel de conteúdo
        pnlContent = New Panel()
        pnlContent.Dock = DockStyle.Fill
        pnlContent.Padding = New Padding(20)
        Me.Controls.Add(pnlContent)
        
        ' Painel de rodapé
        pnlFooter = New Panel()
        pnlFooter.Height = 80
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
        lblTitulo.Text = "📊 RELATÓRIOS E CONSULTAS"
        lblTitulo.Font = New Font("Segoe UI", 18, FontStyle.Bold)
        lblTitulo.ForeColor = Color.White
        lblTitulo.AutoSize = True
        lblTitulo.Location = New Point(20, 25)
        pnlHeader.Controls.Add(lblTitulo)
        
        ' Filtros
        ConfigurarFiltros()
        
        ' Grid de resultados
        ConfigurarGrid()
        
        ' Estatísticas
        ConfigurarEstatisticas()
        
        ' Botões de ação
        ConfigurarBotoesAcao()
    End Sub
    
    ''' <summary>
    ''' Configura filtros
    ''' </summary>
    Private Sub ConfigurarFiltros()
        ' Data início
        Dim lblDataInicio = New Label()
        lblDataInicio.Text = "Data Início:"
        lblDataInicio.Location = New Point(0, 0)
        lblDataInicio.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        pnlFilters.Controls.Add(lblDataInicio)
        
        dtpDataInicio = New DateTimePicker()
        dtpDataInicio.Location = New Point(0, 25)
        dtpDataInicio.Size = New Size(150, 25)
        dtpDataInicio.Value = DateTime.Today.AddDays(-30)
        pnlFilters.Controls.Add(dtpDataInicio)
        
        ' Data fim
        Dim lblDataFim = New Label()
        lblDataFim.Text = "Data Fim:"
        lblDataFim.Location = New Point(170, 0)
        lblDataFim.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        pnlFilters.Controls.Add(lblDataFim)
        
        dtpDataFim = New DateTimePicker()
        dtpDataFim.Location = New Point(170, 25)
        dtpDataFim.Size = New Size(150, 25)
        dtpDataFim.Value = DateTime.Today
        pnlFilters.Controls.Add(dtpDataFim)
        
        ' Cliente
        Dim lblCliente = New Label()
        lblCliente.Text = "Cliente:"
        lblCliente.Location = New Point(340, 0)
        lblCliente.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        pnlFilters.Controls.Add(lblCliente)
        
        txtCliente = New TextBox()
        txtCliente.Location = New Point(340, 25)
        txtCliente.Size = New Size(200, 25)
        txtCliente.PlaceholderText = "Nome do cliente"
        pnlFilters.Controls.Add(txtCliente)
        
        ' Vendedor
        Dim lblVendedor = New Label()
        lblVendedor.Text = "Vendedor:"
        lblVendedor.Location = New Point(560, 0)
        lblVendedor.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        pnlFilters.Controls.Add(lblVendedor)
        
        txtVendedor = New TextBox()
        txtVendedor.Location = New Point(560, 25)
        txtVendedor.Size = New Size(200, 25)
        txtVendedor.PlaceholderText = "Nome do vendedor"
        pnlFilters.Controls.Add(txtVendedor)
        
        ' Botões de filtro
        btnFiltrar = New Button()
        btnFiltrar.Text = "🔍 FILTRAR"
        btnFiltrar.Location = New Point(0, 70)
        btnFiltrar.Size = New Size(120, 35)
        btnFiltrar.BackColor = Color.FromArgb(46, 204, 113)
        btnFiltrar.ForeColor = Color.White
        btnFiltrar.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        btnFiltrar.FlatStyle = FlatStyle.Flat
        btnFiltrar.FlatAppearance.BorderSize = 0
        pnlFilters.Controls.Add(btnFiltrar)
        
        btnLimpar = New Button()
        btnLimpar.Text = "🗑️ LIMPAR"
        btnLimpar.Location = New Point(130, 70)
        btnLimpar.Size = New Size(120, 35)
        btnLimpar.BackColor = Color.FromArgb(149, 165, 166)
        btnLimpar.ForeColor = Color.White
        btnLimpar.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        btnLimpar.FlatStyle = FlatStyle.Flat
        btnLimpar.FlatAppearance.BorderSize = 0
        pnlFilters.Controls.Add(btnLimpar)
    End Sub
    
    ''' <summary>
    ''' Configura grid de resultados
    ''' </summary>
    Private Sub ConfigurarGrid()
        dgvVendas = New DataGridView()
        dgvVendas.Dock = DockStyle.Fill
        dgvVendas.ReadOnly = True
        dgvVendas.AllowUserToAddRows = False
        dgvVendas.AllowUserToDeleteRows = False
        dgvVendas.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvVendas.MultiSelect = False
        dgvVendas.AutoGenerateColumns = False
        dgvVendas.BorderStyle = BorderStyle.None
        dgvVendas.BackgroundColor = Color.White
        dgvVendas.GridColor = Color.FromArgb(189, 195, 199)
        
        ' Configurar colunas
        dgvVendas.Columns.Add("NumeroTalao", "Talão")
        dgvVendas.Columns.Add("DataVenda", "Data")
        dgvVendas.Columns.Add("Cliente", "Cliente")
        dgvVendas.Columns.Add("Vendedor", "Vendedor")
        dgvVendas.Columns.Add("Produtos", "Produtos")
        dgvVendas.Columns.Add("ValorTotal", "Valor Total")
        dgvVendas.Columns.Add("FormaPagamento", "Pagamento")
        
        ' Configurar largura das colunas
        dgvVendas.Columns("NumeroTalao").Width = 120
        dgvVendas.Columns("DataVenda").Width = 120
        dgvVendas.Columns("Cliente").Width = 200
        dgvVendas.Columns("Vendedor").Width = 150
        dgvVendas.Columns("Produtos").Width = 80
        dgvVendas.Columns("ValorTotal").Width = 120
        dgvVendas.Columns("FormaPagamento").Width = 150
        
        ' Formatar coluna de valor
        dgvVendas.Columns("ValorTotal").DefaultCellStyle.Format = "C"
        dgvVendas.Columns("ValorTotal").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        
        ' Formatar coluna de data
        dgvVendas.Columns("DataVenda").DefaultCellStyle.Format = "dd/MM/yyyy HH:mm"
        
        ' Formatar coluna de produtos (centralized)
        dgvVendas.Columns("Produtos").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        
        pnlContent.Controls.Add(dgvVendas)
    End Sub
    
    ''' <summary>
    ''' Configura estatísticas
    ''' </summary>
    Private Sub ConfigurarEstatisticas()
        lblTotalVendas = New Label()
        lblTotalVendas.Text = "Total de Vendas: 0"
        lblTotalVendas.Location = New Point(0, 10)
        lblTotalVendas.Font = New Font("Segoe UI", 12, FontStyle.Bold)
        lblTotalVendas.ForeColor = Color.FromArgb(52, 73, 94)
        lblTotalVendas.AutoSize = True
        pnlFooter.Controls.Add(lblTotalVendas)
        
        lblValorTotal = New Label()
        lblValorTotal.Text = "Valor Total: R$ 0,00"
        lblValorTotal.Location = New Point(200, 10)
        lblValorTotal.Font = New Font("Segoe UI", 12, FontStyle.Bold)
        lblValorTotal.ForeColor = Color.FromArgb(46, 204, 113)
        lblValorTotal.AutoSize = True
        pnlFooter.Controls.Add(lblValorTotal)
        
        lblTicketMedio = New Label()
        lblTicketMedio.Text = "Ticket Médio: R$ 0,00"
        lblTicketMedio.Location = New Point(400, 10)
        lblTicketMedio.Font = New Font("Segoe UI", 12, FontStyle.Bold)
        lblTicketMedio.ForeColor = Color.FromArgb(155, 89, 182)
        lblTicketMedio.AutoSize = True
        pnlFooter.Controls.Add(lblTicketMedio)
    End Sub
    
    ''' <summary>
    ''' Configura botões de ação
    ''' </summary>
    Private Sub ConfigurarBotoesAcao()
        btnReimprimir = New Button()
        btnReimprimir.Text = "🖨️ REIMPRIMIR"
        btnReimprimir.Location = New Point(0, 40)
        btnReimprimir.Size = New Size(140, 35)
        btnReimprimir.BackColor = Color.FromArgb(52, 152, 219)
        btnReimprimir.ForeColor = Color.White
        btnReimprimir.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        btnReimprimir.FlatStyle = FlatStyle.Flat
        btnReimprimir.FlatAppearance.BorderSize = 0
        btnReimprimir.Enabled = False
        pnlFooter.Controls.Add(btnReimprimir)
        
        btnExportar = New Button()
        btnExportar.Text = "📄 EXPORTAR"
        btnExportar.Location = New Point(150, 40)
        btnExportar.Size = New Size(140, 35)
        btnExportar.BackColor = Color.FromArgb(230, 126, 34)
        btnExportar.ForeColor = Color.White
        btnExportar.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        btnExportar.FlatStyle = FlatStyle.Flat
        btnExportar.FlatAppearance.BorderSize = 0
        pnlFooter.Controls.Add(btnExportar)
        
        btnFechar = New Button()
        btnFechar.Text = "❌ FECHAR"
        btnFechar.Location = New Point(300, 40)
        btnFechar.Size = New Size(140, 35)
        btnFechar.BackColor = Color.FromArgb(231, 76, 60)
        btnFechar.ForeColor = Color.White
        btnFechar.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        btnFechar.FlatStyle = FlatStyle.Flat
        btnFechar.FlatAppearance.BorderSize = 0
        pnlFooter.Controls.Add(btnFechar)
    End Sub
    
    ''' <summary>
    ''' Carrega dados iniciais
    ''' </summary>
    Private Sub CarregarDadosIniciais()
        Try
            FiltrarVendas()
        Catch ex As Exception
            _logger.Error("Erro ao carregar dados iniciais do relatório", ex)
            MessageBox.Show("Erro ao carregar dados iniciais.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    ''' <summary>
    ''' Filtra vendas baseado nos critérios
    ''' </summary>
    Private Sub FiltrarVendas()
        Try
            Dim criterio = New CriterioBusca() With {
                .DataInicio = dtpDataInicio.Value.Date,
                .DataFim = dtpDataFim.Value.Date.AddDays(1).AddSeconds(-1),
                .NomeCliente = txtCliente.Text.Trim(),
                .Vendedor = txtVendedor.Text.Trim()
            }
            
            Dim vendas = _historicoManager.BuscarVendas(criterio)
            
            ' Limpar grid
            dgvVendas.Rows.Clear()
            
            ' Preencher grid
            For Each venda In vendas
                Dim row = dgvVendas.Rows.Add()
                dgvVendas.Rows(row).Cells("NumeroTalao").Value = venda.NumeroTalao
                dgvVendas.Rows(row).Cells("DataVenda").Value = venda.DataVenda
                dgvVendas.Rows(row).Cells("Cliente").Value = venda.Cliente.Nome
                dgvVendas.Rows(row).Cells("Vendedor").Value = venda.Vendedor
                dgvVendas.Rows(row).Cells("Produtos").Value = venda.Itens.Count
                dgvVendas.Rows(row).Cells("ValorTotal").Value = venda.ValorTotal
                dgvVendas.Rows(row).Cells("FormaPagamento").Value = venda.FormaPagamento
                dgvVendas.Rows(row).Tag = venda
            Next
            
            ' Atualizar estatísticas
            AtualizarEstatisticas(vendas)
            
        Catch ex As Exception
            _logger.Error("Erro ao filtrar vendas", ex)
            MessageBox.Show("Erro ao filtrar vendas.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    ''' <summary>
    ''' Atualiza estatísticas
    ''' </summary>
    Private Sub AtualizarEstatisticas(vendas As List(Of Venda))
        Try
            Dim totalVendas = vendas.Count
            Dim valorTotal = vendas.Sum(Function(v) v.ValorTotal)
            Dim ticketMedio = If(totalVendas > 0, valorTotal / totalVendas, 0)
            
            lblTotalVendas.Text = $"Total de Vendas: {totalVendas}"
            lblValorTotal.Text = $"Valor Total: {valorTotal:C}"
            lblTicketMedio.Text = $"Ticket Médio: {ticketMedio:C}"
            
        Catch ex As Exception
            _logger.Error("Erro ao atualizar estatísticas", ex)
        End Try
    End Sub
    
    ''' <summary>
    ''' Evento de filtrar
    ''' </summary>
    Private Sub btnFiltrar_Click(sender As Object, e As EventArgs) Handles btnFiltrar.Click
        FiltrarVendas()
    End Sub
    
    ''' <summary>
    ''' Evento de limpar filtros
    ''' </summary>
    Private Sub btnLimpar_Click(sender As Object, e As EventArgs) Handles btnLimpar.Click
        dtpDataInicio.Value = DateTime.Today.AddDays(-30)
        dtpDataFim.Value = DateTime.Today
        txtCliente.Text = ""
        txtVendedor.Text = ""
        FiltrarVendas()
    End Sub
    
    ''' <summary>
    ''' Evento de seleção no grid
    ''' </summary>
    Private Sub dgvVendas_SelectionChanged(sender As Object, e As EventArgs) Handles dgvVendas.SelectionChanged
        btnReimprimir.Enabled = dgvVendas.SelectedRows.Count > 0
    End Sub
    
    ''' <summary>
    ''' Evento de reimprimir talão
    ''' </summary>
    Private Sub btnReimprimir_Click(sender As Object, e As EventArgs) Handles btnReimprimir.Click
        Try
            If dgvVendas.SelectedRows.Count > 0 Then
                Dim venda = CType(dgvVendas.SelectedRows(0).Tag, Venda)
                
                If MessageBox.Show($"Confirma a reimpressão do talão {venda.NumeroTalao}?", 
                                 "Confirmar Reimpressão", 
                                 MessageBoxButtons.YesNo, 
                                 MessageBoxIcon.Question) = DialogResult.Yes Then
                    
                    If _vendaService.ReimprimirTalao(venda.NumeroTalao) Then
                        MessageBox.Show("Talão reimpresso com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Else
                        MessageBox.Show("Erro ao reimprimir talão.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                End If
            End If
        Catch ex As Exception
            _logger.Error("Erro ao reimprimir talão", ex)
            MessageBox.Show("Erro ao reimprimir talão.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    ''' <summary>
    ''' Evento de exportar relatório
    ''' </summary>
    Private Sub btnExportar_Click(sender As Object, e As EventArgs) Handles btnExportar.Click
        Try
            Dim relatorio = _historicoManager.GerarRelatorioVendas(dtpDataInicio.Value, dtpDataFim.Value)
            
            Dim saveDialog = New SaveFileDialog()
            saveDialog.Filter = "Arquivo XML|*.xml"
            saveDialog.FileName = $"Relatorio_Vendas_{DateTime.Now:yyyyMMdd_HHmmss}.xml"
            
            If saveDialog.ShowDialog() = DialogResult.OK Then
                If _historicoManager.ExportarRelatorio(relatorio, saveDialog.FileName) Then
                    MessageBox.Show($"Relatório exportado para:{vbCrLf}{saveDialog.FileName}", 
                                  "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    MessageBox.Show("Erro ao exportar relatório.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            End If
            
        Catch ex As Exception
            _logger.Error("Erro ao exportar relatório", ex)
            MessageBox.Show("Erro ao exportar relatório.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    ''' <summary>
    ''' Evento de fechar formulário
    ''' </summary>
    Private Sub btnFechar_Click(sender As Object, e As EventArgs) Handles btnFechar.Click
        Me.Close()
    End Sub
    
    ''' <summary>
    ''' Evento de tecla pressionada
    ''' </summary>
    Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean
        Select Case keyData
            Case Keys.Escape
                btnFechar_Click(Nothing, Nothing)
                Return True
            Case Keys.F5
                FiltrarVendas()
                Return True
            Case Keys.Enter
                If btnFiltrar.Focused Then
                    btnFiltrar_Click(Nothing, Nothing)
                    Return True
                End If
        End Select
        
        Return MyBase.ProcessCmdKey(msg, keyData)
    End Function
End Class