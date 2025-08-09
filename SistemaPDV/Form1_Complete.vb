' ====================================================================
' FORM1.VB - Interface Principal do Sistema PDV
' Sistema completo para Windows Forms
' Este arquivo contém o código-fonte completo da interface principal
' ====================================================================

Imports System
Imports System.Windows.Forms
Imports System.Drawing

Public Class Form1
    Inherits Form

    ' Controles principais
    Private sidePanel As Panel
    Private mainPanel As Panel
    Private dashboardPanel As Panel
    Private titleLabel As Label

    ' Botões do menu lateral
    Private btnPDV As Button
    Private btnProdutos As Button
    Private btnClientes As Button
    Private btnRelatorios As Button
    Private btnConfiguracao As Button

    ' Cards do dashboard
    Private cardVendas As Panel
    Private cardEstoque As Panel
    Private cardClientes As Panel
    Private cardFaturamento As Panel

    Public Sub New()
        InitializeComponent()
        SetupModernInterface()
    End Sub

    Private Sub InitializeComponent()
        ' Configurações básicas do form
        Me.Text = "Sistema PDV - Madeireira Maria Luiza"
        Me.Size = New Size(1200, 800)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.WindowState = FormWindowState.Maximized
        Me.BackColor = Color.FromArgb(248, 249, 250)
        Me.Font = New Font("Segoe UI", 9.0F, FontStyle.Regular)

        ' Inicializar controles
        InitializePanels()
        InitializeMenuButtons()
        InitializeDashboard()
    End Sub

    Private Sub InitializePanels()
        ' Painel lateral (menu)
        sidePanel = New Panel()
        sidePanel.Dock = DockStyle.Left
        sidePanel.Width = 250
        sidePanel.BackColor = Color.FromArgb(33, 37, 41)
        Me.Controls.Add(sidePanel)

        ' Título do sistema no painel lateral
        titleLabel = New Label()
        titleLabel.Text = "MADEIREIRA" & vbCrLf & "MARIA LUIZA"
        titleLabel.ForeColor = Color.White
        titleLabel.Font = New Font("Segoe UI", 14.0F, FontStyle.Bold)
        titleLabel.TextAlign = ContentAlignment.MiddleCenter
        titleLabel.Dock = DockStyle.Top
        titleLabel.Height = 80
        sidePanel.Controls.Add(titleLabel)

        ' Painel principal
        mainPanel = New Panel()
        mainPanel.Dock = DockStyle.Fill
        mainPanel.BackColor = Color.FromArgb(248, 249, 250)
        mainPanel.Padding = New Padding(20)
        Me.Controls.Add(mainPanel)

        ' Painel do dashboard
        dashboardPanel = New Panel()
        dashboardPanel.Dock = DockStyle.Fill
        dashboardPanel.BackColor = Color.Transparent
        mainPanel.Controls.Add(dashboardPanel)
    End Sub

    Private Sub InitializeMenuButtons()
        Dim buttonHeight As Integer = 60
        Dim buttonSpacing As Integer = 5
        Dim yPosition As Integer = 100

        ' Botão PDV / Caixa
        btnPDV = CreateMenuButton("🏪 PDV / CAIXA", yPosition)
        AddHandler btnPDV.Click, AddressOf BtnPDV_Click
        sidePanel.Controls.Add(btnPDV)
        yPosition += buttonHeight + buttonSpacing

        ' Botão Produtos
        btnProdutos = CreateMenuButton("📦 PRODUTOS", yPosition)
        sidePanel.Controls.Add(btnProdutos)
        yPosition += buttonHeight + buttonSpacing

        ' Botão Clientes
        btnClientes = CreateMenuButton("👥 CLIENTES", yPosition)
        sidePanel.Controls.Add(btnClientes)
        yPosition += buttonHeight + buttonSpacing

        ' Botão Relatórios
        btnRelatorios = CreateMenuButton("📊 RELATÓRIOS", yPosition)
        sidePanel.Controls.Add(btnRelatorios)
        yPosition += buttonHeight + buttonSpacing

        ' Botão Configuração
        btnConfiguracao = CreateMenuButton("⚙️ CONFIGURAÇÃO", yPosition)
        sidePanel.Controls.Add(btnConfiguracao)
    End Sub

    Private Function CreateMenuButton(text As String, yPos As Integer) As Button
        Dim btn As New Button()
        btn.Text = text
        btn.Size = New Size(230, 55)
        btn.Location = New Point(10, yPos)
        btn.FlatStyle = FlatStyle.Flat
        btn.FlatAppearance.BorderSize = 0
        btn.BackColor = Color.FromArgb(52, 58, 64)
        btn.ForeColor = Color.White
        btn.Font = New Font("Segoe UI", 10.0F, FontStyle.Bold)
        btn.TextAlign = ContentAlignment.MiddleLeft
        btn.Padding = New Padding(15, 0, 0, 0)
        btn.Cursor = Cursors.Hand

        ' Efeitos hover
        AddHandler btn.MouseEnter, Sub() btn.BackColor = Color.FromArgb(0, 123, 255)
        AddHandler btn.MouseLeave, Sub() btn.BackColor = Color.FromArgb(52, 58, 64)

        Return btn
    End Function

    Private Sub InitializeDashboard()
        ' Título do dashboard
        Dim dashTitle As New Label()
        dashTitle.Text = "Dashboard - Visão Geral"
        dashTitle.Font = New Font("Segoe UI", 18.0F, FontStyle.Bold)
        dashTitle.ForeColor = Color.FromArgb(33, 37, 41)
        dashTitle.Location = New Point(20, 20)
        dashTitle.Size = New Size(400, 40)
        dashboardPanel.Controls.Add(dashTitle)

        ' Container dos cards
        Dim cardContainer As New TableLayoutPanel()
        cardContainer.Location = New Point(20, 80)
        cardContainer.Size = New Size(800, 200)
        cardContainer.ColumnCount = 4
        cardContainer.RowCount = 1
        cardContainer.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 25))
        cardContainer.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 25))
        cardContainer.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 25))
        cardContainer.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 25))
        dashboardPanel.Controls.Add(cardContainer)

        ' Criar cards informativos
        cardVendas = CreateInfoCard("💰", "Vendas Hoje", "R$ 2.450,00", Color.FromArgb(40, 167, 69))
        cardEstoque = CreateInfoCard("📦", "Estoque", "1.240 itens", Color.FromArgb(255, 193, 7))
        cardClientes = CreateInfoCard("👥", "Clientes", "324", Color.FromArgb(0, 123, 255))
        cardFaturamento = CreateInfoCard("📊", "Faturamento", "R$ 18.930,00", Color.FromArgb(220, 53, 69))

        cardContainer.Controls.Add(cardVendas, 0, 0)
        cardContainer.Controls.Add(cardEstoque, 1, 0)
        cardContainer.Controls.Add(cardClientes, 2, 0)
        cardContainer.Controls.Add(cardFaturamento, 3, 0)

        ' Informações da empresa
        CreateCompanyInfo()
    End Sub

    Private Function CreateInfoCard(icon As String, title As String, value As String, color As Color) As Panel
        Dim card As New Panel()
        card.Size = New Size(180, 150)
        card.BackColor = Color.White
        card.Margin = New Padding(10)

        ' Borda colorida superior
        Dim topBorder As New Panel()
        topBorder.Dock = DockStyle.Top
        topBorder.Height = 4
        topBorder.BackColor = color
        card.Controls.Add(topBorder)

        ' Ícone
        Dim iconLabel As New Label()
        iconLabel.Text = icon
        iconLabel.Font = New Font("Segoe UI", 24.0F)
        iconLabel.Location = New Point(20, 20)
        iconLabel.Size = New Size(50, 40)
        card.Controls.Add(iconLabel)

        ' Título
        Dim titleLabel As New Label()
        titleLabel.Text = title
        titleLabel.Font = New Font("Segoe UI", 10.0F, FontStyle.Bold)
        titleLabel.ForeColor = Color.FromArgb(108, 117, 125)
        titleLabel.Location = New Point(20, 70)
        titleLabel.Size = New Size(140, 20)
        card.Controls.Add(titleLabel)

        ' Valor
        Dim valueLabel As New Label()
        valueLabel.Text = value
        valueLabel.Font = New Font("Segoe UI", 16.0F, FontStyle.Bold)
        valueLabel.ForeColor = color
        valueLabel.Location = New Point(20, 95)
        valueLabel.Size = New Size(140, 30)
        card.Controls.Add(valueLabel)

        Return card
    End Function

    Private Sub CreateCompanyInfo()
        Dim infoPanel As New Panel()
        infoPanel.Location = New Point(20, 320)
        infoPanel.Size = New Size(800, 200)
        infoPanel.BackColor = Color.White
        infoPanel.Padding = New Padding(20)
        dashboardPanel.Controls.Add(infoPanel)

        Dim infoTitle As New Label()
        infoTitle.Text = "Informações da Empresa"
        infoTitle.Font = New Font("Segoe UI", 14.0F, FontStyle.Bold)
        infoTitle.Location = New Point(20, 20)
        infoTitle.Size = New Size(300, 30)
        infoPanel.Controls.Add(infoTitle)

        Dim infoText As New Label()
        infoText.Text = "MADEIREIRA MARIA LUIZA" & vbCrLf &
                       "Av. Dr. Olíncio Guerreiro Leite - 631-Paq Amadeu-Paulista-PE-55431-165" & vbCrLf &
                       "Telefone: (81) 98570-1522" & vbCrLf &
                       "CNPJ: 48.905.025/001-61" & vbCrLf &
                       "WhatsApp e Instagram: @madeireiramaria"
        infoText.Font = New Font("Segoe UI", 10.0F)
        infoText.ForeColor = Color.FromArgb(108, 117, 125)
        infoText.Location = New Point(20, 60)
        infoText.Size = New Size(750, 120)
        infoPanel.Controls.Add(infoText)
    End Sub

    Private Sub SetupModernInterface()
        ' Aplicar sombras e efeitos visuais modernos
        Me.SetStyle(ControlStyles.AllPaintingInWmPaint Or
                   ControlStyles.UserPaint Or
                   ControlStyles.DoubleBuffer, True)
    End Sub

    Private Sub BtnPDV_Click(sender As Object, e As EventArgs)
        ' Abrir formulário de PDV
        Dim formPDV As New FormPDV()
        formPDV.ShowDialog()
    End Sub
End Class