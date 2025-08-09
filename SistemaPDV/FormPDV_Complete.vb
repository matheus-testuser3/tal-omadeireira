' ====================================================================
' FORMPDV.VB - Formulário de Entrada de Dados do PDV
' Sistema completo para Windows Forms
' Este arquivo contém o código-fonte completo do formulário PDV
' ====================================================================

Imports System
Imports System.Windows.Forms
Imports System.Drawing

Public Class FormPDV
    Inherits Form

    ' Controles de entrada de dados
    Private txtNomeCliente As TextBox
    Private txtEndereco As TextBox
    Private txtCidade As TextBox
    Private txtCEP As TextBox
    Private txtProdutos As TextBox
    Private txtValorTotal As TextBox
    Private cmbFormaPagamento As ComboBox
    Private txtVendedor As TextBox

    ' Botões
    Private btnGerarTalao As Button
    Private btnLimpar As Button
    Private btnFechar As Button

    ' Labels
    Private lblTitulo As Label

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub InitializeComponent()
        ' Configurações do formulário
        Me.Text = "PDV / Caixa - Gerar Talão"
        Me.Size = New Size(600, 700)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.BackColor = Color.FromArgb(248, 249, 250)
        Me.Font = New Font("Segoe UI", 9.0F)
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False

        CreateControls()
        LayoutControls()
    End Sub

    Private Sub CreateControls()
        ' Título
        lblTitulo = New Label()
        lblTitulo.Text = "Sistema PDV - Gerar Talão de Venda"
        lblTitulo.Font = New Font("Segoe UI", 16.0F, FontStyle.Bold)
        lblTitulo.ForeColor = Color.FromArgb(33, 37, 41)
        lblTitulo.TextAlign = ContentAlignment.MiddleCenter

        ' Campos de entrada
        txtNomeCliente = CreateTextBox("Digite o nome do cliente...")
        txtEndereco = CreateTextBox("Digite o endereço completo...")
        txtCidade = CreateTextBox("Paulista")
        txtCEP = CreateTextBox("55431-165")
        txtProdutos = CreateMultilineTextBox("Descreva os produtos/serviços...")
        txtValorTotal = CreateTextBox("0,00")
        txtVendedor = CreateTextBox("matheus-testuser3")

        ' ComboBox forma de pagamento
        cmbFormaPagamento = New ComboBox()
        cmbFormaPagamento.DropDownStyle = ComboBoxStyle.DropDownList
        cmbFormaPagamento.Items.AddRange({"Dinheiro", "Cartão Débito", "Cartão Crédito", "PIX", "Cheque", "Fiado"})
        cmbFormaPagamento.SelectedIndex = 0
        cmbFormaPagamento.Font = New Font("Segoe UI", 10.0F)

        ' Botões
        btnGerarTalao = CreateButton("🖨️ GERAR TALÃO", Color.FromArgb(40, 167, 69))
        btnLimpar = CreateButton("🗑️ LIMPAR", Color.FromArgb(255, 193, 7))
        btnFechar = CreateButton("❌ FECHAR", Color.FromArgb(220, 53, 69))

        ' Eventos dos botões
        AddHandler btnGerarTalao.Click, AddressOf BtnGerarTalao_Click
        AddHandler btnLimpar.Click, AddressOf BtnLimpar_Click
        AddHandler btnFechar.Click, AddressOf BtnFechar_Click
    End Sub

    Private Function CreateTextBox(placeholder As String) As TextBox
        Dim txt As New TextBox()
        txt.Font = New Font("Segoe UI", 10.0F)
        txt.ForeColor = Color.Gray
        txt.Text = placeholder
        txt.Size = New Size(400, 25)

        ' Eventos para placeholder
        AddHandler txt.Enter, Sub()
                                  If txt.Text = placeholder Then
                                      txt.Text = ""
                                      txt.ForeColor = Color.Black
                                  End If
                              End Sub

        AddHandler txt.Leave, Sub()
                                  If String.IsNullOrWhiteSpace(txt.Text) Then
                                      txt.Text = placeholder
                                      txt.ForeColor = Color.Gray
                                  End If
                              End Sub

        Return txt
    End Function

    Private Function CreateMultilineTextBox(placeholder As String) As TextBox
        Dim txt As TextBox = CreateTextBox(placeholder)
        txt.Multiline = True
        txt.Size = New Size(400, 80)
        txt.ScrollBars = ScrollBars.Vertical
        Return txt
    End Function

    Private Function CreateButton(text As String, color As Color) As Button
        Dim btn As New Button()
        btn.Text = text
        btn.Size = New Size(150, 40)
        btn.FlatStyle = FlatStyle.Flat
        btn.FlatAppearance.BorderSize = 0
        btn.BackColor = color
        btn.ForeColor = Color.White
        btn.Font = New Font("Segoe UI", 10.0F, FontStyle.Bold)
        btn.Cursor = Cursors.Hand

        ' Efeito hover
        AddHandler btn.MouseEnter, Sub() btn.BackColor = Color.FromArgb(color.R - 20, color.G - 20, color.B - 20)
        AddHandler btn.MouseLeave, Sub() btn.BackColor = color

        Return btn
    End Function

    Private Function CreateLabel(text As String) As Label
        Dim lbl As New Label()
        lbl.Text = text
        lbl.Font = New Font("Segoe UI", 10.0F, FontStyle.Bold)
        lbl.ForeColor = Color.FromArgb(33, 37, 41)
        lbl.Size = New Size(150, 25)
        Return lbl
    End Function

    Private Sub LayoutControls()
        Dim yPos As Integer = 30
        Dim spacing As Integer = 45
        Dim xLabel As Integer = 50
        Dim xControl As Integer = 50

        ' Título
        lblTitulo.Location = New Point(50, yPos)
        lblTitulo.Size = New Size(500, 30)
        Me.Controls.Add(lblTitulo)
        yPos += 60

        ' Nome do Cliente
        Me.Controls.Add(CreateLabel("Nome do Cliente:"))
        Me.Controls(Me.Controls.Count - 1).Location = New Point(xLabel, yPos)
        yPos += 25
        txtNomeCliente.Location = New Point(xControl, yPos)
        Me.Controls.Add(txtNomeCliente)
        yPos += spacing

        ' Endereço
        Me.Controls.Add(CreateLabel("Endereço:"))
        Me.Controls(Me.Controls.Count - 1).Location = New Point(xLabel, yPos)
        yPos += 25
        txtEndereco.Location = New Point(xControl, yPos)
        Me.Controls.Add(txtEndereco)
        yPos += spacing

        ' Cidade
        Me.Controls.Add(CreateLabel("Cidade:"))
        Me.Controls(Me.Controls.Count - 1).Location = New Point(xLabel, yPos)
        yPos += 25
        txtCidade.Location = New Point(xControl, yPos)
        Me.Controls.Add(txtCidade)
        yPos += spacing

        ' CEP
        Me.Controls.Add(CreateLabel("CEP:"))
        Me.Controls(Me.Controls.Count - 1).Location = New Point(xLabel, yPos)
        yPos += 25
        txtCEP.Location = New Point(xControl, yPos)
        Me.Controls.Add(txtCEP)
        yPos += spacing

        ' Produtos
        Me.Controls.Add(CreateLabel("Produtos/Serviços:"))
        Me.Controls(Me.Controls.Count - 1).Location = New Point(xLabel, yPos)
        yPos += 25
        txtProdutos.Location = New Point(xControl, yPos)
        Me.Controls.Add(txtProdutos)
        yPos += 100

        ' Valor Total
        Me.Controls.Add(CreateLabel("Valor Total (R$):"))
        Me.Controls(Me.Controls.Count - 1).Location = New Point(xLabel, yPos)
        yPos += 25
        txtValorTotal.Location = New Point(xControl, yPos)
        Me.Controls.Add(txtValorTotal)
        yPos += spacing

        ' Forma de Pagamento
        Me.Controls.Add(CreateLabel("Forma de Pagamento:"))
        Me.Controls(Me.Controls.Count - 1).Location = New Point(xLabel, yPos)
        yPos += 25
        cmbFormaPagamento.Location = New Point(xControl, yPos)
        cmbFormaPagamento.Size = New Size(400, 25)
        Me.Controls.Add(cmbFormaPagamento)
        yPos += spacing

        ' Vendedor
        Me.Controls.Add(CreateLabel("Vendedor:"))
        Me.Controls(Me.Controls.Count - 1).Location = New Point(xLabel, yPos)
        yPos += 25
        txtVendedor.Location = New Point(xControl, yPos)
        Me.Controls.Add(txtVendedor)
        yPos += 60

        ' Botões
        btnGerarTalao.Location = New Point(50, yPos)
        btnLimpar.Location = New Point(220, yPos)
        btnFechar.Location = New Point(390, yPos)

        Me.Controls.Add(btnGerarTalao)
        Me.Controls.Add(btnLimpar)
        Me.Controls.Add(btnFechar)
    End Sub

    Private Sub BtnGerarTalao_Click(sender As Object, e As EventArgs)
        ' Validar dados
        If Not ValidarDados() Then
            Return
        End If

        Try
            ' Criar instância do módulo VBA e processar talão
            Dim moduloVBA As New ModuloTalaoVBA()
            
            ' Coletar dados do formulário
            Dim dadosCliente As New DadosCliente()
            dadosCliente.Nome = GetTextoLimpo(txtNomeCliente)
            dadosCliente.Endereco = GetTextoLimpo(txtEndereco)
            dadosCliente.Cidade = GetTextoLimpo(txtCidade)
            dadosCliente.CEP = GetTextoLimpo(txtCEP)
            dadosCliente.Produtos = GetTextoLimpo(txtProdutos)
            dadosCliente.ValorTotal = GetTextoLimpo(txtValorTotal)
            dadosCliente.FormaPagamento = cmbFormaPagamento.SelectedItem.ToString()
            dadosCliente.Vendedor = GetTextoLimpo(txtVendedor)

            ' Chamar módulo VBA para gerar talão
            moduloVBA.ProcessarTalaoCompleto(dadosCliente)

            MessageBox.Show("Talão gerado e enviado para impressão com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)

            ' Limpar formulário para próxima venda
            LimparFormulario()

        Catch ex As Exception
            MessageBox.Show($"Erro ao gerar talão: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Function ValidarDados() As Boolean
        If String.IsNullOrWhiteSpace(GetTextoLimpo(txtNomeCliente)) Then
            MessageBox.Show("Por favor, informe o nome do cliente.", "Validação", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtNomeCliente.Focus()
            Return False
        End If

        If String.IsNullOrWhiteSpace(GetTextoLimpo(txtProdutos)) Then
            MessageBox.Show("Por favor, informe os produtos/serviços.", "Validação", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtProdutos.Focus()
            Return False
        End If

        If String.IsNullOrWhiteSpace(GetTextoLimpo(txtValorTotal)) OrElse GetTextoLimpo(txtValorTotal) = "0,00" Then
            MessageBox.Show("Por favor, informe o valor total.", "Validação", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtValorTotal.Focus()
            Return False
        End If

        Return True
    End Function

    Private Function GetTextoLimpo(textBox As TextBox) As String
        Dim texto As String = textBox.Text.Trim()
        ' Verificar se é placeholder
        If textBox.ForeColor = Color.Gray Then
            Return ""
        End If
        Return texto
    End Function

    Private Sub BtnLimpar_Click(sender As Object, e As EventArgs)
        LimparFormulario()
    End Sub

    Private Sub LimparFormulario()
        txtNomeCliente.Text = "Digite o nome do cliente..."
        txtNomeCliente.ForeColor = Color.Gray
        txtEndereco.Text = "Digite o endereço completo..."
        txtEndereco.ForeColor = Color.Gray
        txtCidade.Text = "Paulista"
        txtCEP.Text = "55431-165"
        txtProdutos.Text = "Descreva os produtos/serviços..."
        txtProdutos.ForeColor = Color.Gray
        txtValorTotal.Text = "0,00"
        cmbFormaPagamento.SelectedIndex = 0
        txtVendedor.Text = "matheus-testuser3"
        txtNomeCliente.Focus()
    End Sub

    Private Sub BtnFechar_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub
End Class