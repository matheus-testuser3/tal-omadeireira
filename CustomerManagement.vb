Imports System.Windows.Forms
Imports System.Drawing

''' <summary>
''' Sistema de gestão de clientes integrado
''' Formulário completo para CRUD de clientes com relatórios
''' </summary>
Public Class CustomerManager
    Private _clientes As List(Of Cliente)
    Private _database As DatabaseManager
    
    Public Sub New()
        _database = DatabaseManager.Instance
        _clientes = New List(Of Cliente)()
        CarregarClientes()
    End Sub
    
    ''' <summary>
    ''' Carrega todos os clientes do banco/Excel
    ''' </summary>
    Public Sub CarregarClientes()
        Try
            ' TODO: Implementar carregamento real do banco
            _clientes = ObterClientesPadrao()
        Catch ex As Exception
            Console.WriteLine($"Erro ao carregar clientes: {ex.Message}")
        End Try
    End Sub
    
    ''' <summary>
    ''' Busca clientes por termo
    ''' </summary>
    Public Function BuscarClientes(termo As String) As List(Of Cliente)
        If String.IsNullOrEmpty(termo) Then
            Return _clientes
        End If
        
        Return _clientes.Where(Function(c) 
            c.Nome.ToUpper().Contains(termo.ToUpper()) OrElse
            c.CPF_CNPJ.Contains(termo) OrElse
            c.Telefone.Contains(termo)
        ).ToList()
    End Function
    
    ''' <summary>
    ''' Adiciona novo cliente
    ''' </summary>
    Public Function AdicionarCliente(cliente As Cliente) As Boolean
        Try
            cliente.ID = _clientes.Count + 1
            cliente.DataCadastro = Date.Now
            _clientes.Add(cliente)
            Return True
        Catch ex As Exception
            Console.WriteLine($"Erro ao adicionar cliente: {ex.Message}")
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' Atualiza cliente existente
    ''' </summary>
    Public Function AtualizarCliente(cliente As Cliente) As Boolean
        Try
            Dim index = _clientes.FindIndex(Function(c) c.ID = cliente.ID)
            If index >= 0 Then
                _clientes(index) = cliente
                Return True
            End If
            Return False
        Catch ex As Exception
            Console.WriteLine($"Erro ao atualizar cliente: {ex.Message}")
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' Remove cliente
    ''' </summary>
    Public Function RemoverCliente(clienteId As Integer) As Boolean
        Try
            Return _clientes.RemoveAll(Function(c) c.ID = clienteId) > 0
        Catch ex As Exception
            Console.WriteLine($"Erro ao remover cliente: {ex.Message}")
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' Obtém cliente por ID
    ''' </summary>
    Public Function ObterClientePorId(id As Integer) As Cliente
        Return _clientes.FirstOrDefault(Function(c) c.ID = id)
    End Function
    
    ''' <summary>
    ''' Obtém clientes padrão para teste
    ''' </summary>
    Private Function ObterClientesPadrao() As List(Of Cliente)
        Return New List(Of Cliente) From {
            New Cliente() With {
                .ID = 1,
                .Nome = "João Silva",
                .Endereco = "Rua das Árvores, 123 - Centro",
                .CEP = "55431-165",
                .Cidade = "Paulista",
                .UF = "PE",
                .Telefone = "(81) 9876-5432",
                .CPF_CNPJ = "123.456.789-01",
                .Email = "joao@email.com",
                .DataCadastro = Date.Now.AddDays(-30),
                .Ativo = True
            },
            New Cliente() With {
                .ID = 2,
                .Nome = "Maria Oliveira",
                .Endereco = "Av. Principal, 456 - Jardim",
                .CEP = "55432-000",
                .Cidade = "Olinda",
                .UF = "PE",
                .Telefone = "(81) 9123-4567",
                .CPF_CNPJ = "987.654.321-02",
                .Email = "maria@email.com",
                .DataCadastro = Date.Now.AddDays(-15),
                .Ativo = True
            },
            New Cliente() With {
                .ID = 3,
                .Nome = "Construtora ABC Ltda",
                .Endereco = "Rua Industrial, 789 - Distrito",
                .CEP = "55433-111",
                .Cidade = "Recife",
                .UF = "PE",
                .Telefone = "(81) 3456-7890",
                .CPF_CNPJ = "12.345.678/0001-90",
                .Email = "contato@abcconstrutora.com",
                .DataCadastro = Date.Now.AddDays(-60),
                .Ativo = True
            }
        }
    End Function
    
    ''' <summary>
    ''' Propriedade para acessar todos os clientes
    ''' </summary>
    Public ReadOnly Property TodosClientes As List(Of Cliente)
        Get
            Return _clientes
        End Get
    End Property
End Class

''' <summary>
''' Formulário de gestão de clientes
''' </summary>
Public Class FormGestaoClientes
    Inherits Form
    
    Private WithEvents dgvClientes As DataGridView
    Private WithEvents txtBusca As TextBox
    Private WithEvents btnBuscar As Button
    Private WithEvents btnNovo As Button
    Private WithEvents btnEditar As Button
    Private WithEvents btnRemover As Button
    Private WithEvents btnRelatorios As Button
    Private WithEvents btnFechar As Button
    
    Private _manager As CustomerManager
    
    Public Sub New()
        InitializeComponent()
        _manager = New CustomerManager()
        ConfigurarInterface()
        CarregarDados()
    End Sub
    
    Private Sub InitializeComponent()
        Me.Text = "Gestão de Clientes"
        Me.Size = New Size(1000, 600)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.BackColor = Color.WhiteSmoke
        
        ' Painel de busca
        Dim pnlBusca As New Panel() With {
            .Dock = DockStyle.Top,
            .Height = 60,
            .BackColor = Color.White,
            .Padding = New Padding(10)
        }
        
        Dim lblBusca As New Label() With {
            .Text = "Buscar cliente:",
            .Location = New Point(10, 10),
            .AutoSize = True,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold)
        }
        
        txtBusca = New TextBox() With {
            .Location = New Point(10, 30),
            .Size = New Size(300, 25),
            .Font = New Font("Segoe UI", 10)
        }
        
        btnBuscar = New Button() With {
            .Text = "🔍 Buscar",
            .Location = New Point(320, 30),
            .Size = New Size(80, 25),
            .BackColor = Color.DodgerBlue,
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
        
        pnlBusca.Controls.AddRange({lblBusca, txtBusca, btnBuscar})
        
        ' Grid de clientes
        dgvClientes = New DataGridView() With {
            .Dock = DockStyle.Fill,
            .AllowUserToAddRows = False,
            .AllowUserToDeleteRows = False,
            .ReadOnly = True,
            .MultiSelect = False,
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            .BackgroundColor = Color.White,
            .BorderStyle = BorderStyle.None
        }
        
        ' Painel de botões
        Dim pnlBotoes As New Panel() With {
            .Dock = DockStyle.Bottom,
            .Height = 60,
            .BackColor = Color.LightGray,
            .Padding = New Padding(10)
        }
        
        btnNovo = New Button() With {
            .Text = "➕ Novo",
            .Location = New Point(10, 15),
            .Size = New Size(100, 30),
            .BackColor = Color.Green,
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold)
        }
        
        btnEditar = New Button() With {
            .Text = "✏️ Editar",
            .Location = New Point(120, 15),
            .Size = New Size(100, 30),
            .BackColor = Color.Orange,
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold)
        }
        
        btnRemover = New Button() With {
            .Text = "🗑️ Remover",
            .Location = New Point(230, 15),
            .Size = New Size(100, 30),
            .BackColor = Color.Red,
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold)
        }
        
        btnRelatorios = New Button() With {
            .Text = "📊 Relatórios",
            .Location = New Point(340, 15),
            .Size = New Size(120, 30),
            .BackColor = Color.Purple,
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold)
        }
        
        btnFechar = New Button() With {
            .Text = "❌ Fechar",
            .Location = New Point(880, 15),
            .Size = New Size(100, 30),
            .BackColor = Color.Gray,
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold)
        }
        
        pnlBotoes.Controls.AddRange({btnNovo, btnEditar, btnRemover, btnRelatorios, btnFechar})
        
        Me.Controls.AddRange({pnlBusca, dgvClientes, pnlBotoes})
    End Sub
    
    Private Sub ConfigurarInterface()
        ' Configurar grid
        dgvClientes.Columns.Add("ID", "ID")
        dgvClientes.Columns.Add("Nome", "Nome")
        dgvClientes.Columns.Add("CPF_CNPJ", "CPF/CNPJ")
        dgvClientes.Columns.Add("Telefone", "Telefone")
        dgvClientes.Columns.Add("Cidade", "Cidade")
        dgvClientes.Columns.Add("UF", "UF")
        dgvClientes.Columns.Add("DataCadastro", "Cadastro")
        dgvClientes.Columns.Add("Status", "Status")
        
        ' Configurar larguras
        dgvClientes.Columns("ID").Width = 50
        dgvClientes.Columns("Nome").Width = 250
        dgvClientes.Columns("CPF_CNPJ").Width = 150
        dgvClientes.Columns("Telefone").Width = 120
        dgvClientes.Columns("Cidade").Width = 120
        dgvClientes.Columns("UF").Width = 50
        dgvClientes.Columns("DataCadastro").Width = 100
        dgvClientes.Columns("Status").Width = 80
        
        ' Formatação
        dgvClientes.Columns("DataCadastro").DefaultCellStyle.Format = "dd/MM/yyyy"
    End Sub
    
    Private Sub CarregarDados()
        AtualizarGrid(_manager.TodosClientes)
    End Sub
    
    Private Sub AtualizarGrid(clientes As List(Of Cliente))
        dgvClientes.Rows.Clear()
        
        For Each cliente In clientes
            dgvClientes.Rows.Add(
                cliente.ID,
                cliente.Nome,
                cliente.CPF_CNPJ,
                cliente.Telefone,
                cliente.Cidade,
                cliente.UF,
                cliente.DataCadastro,
                If(cliente.Ativo, "Ativo", "Inativo")
            )
            
            ' Armazenar objeto cliente na tag da linha
            dgvClientes.Rows(dgvClientes.Rows.Count - 1).Tag = cliente
            
            ' Colorir linha se inativo
            If Not cliente.Ativo Then
                dgvClientes.Rows(dgvClientes.Rows.Count - 1).DefaultCellStyle.BackColor = Color.LightGray
            End If
        Next
    End Sub
    
    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click
        Dim termo = txtBusca.Text.Trim()
        Dim clientes = _manager.BuscarClientes(termo)
        AtualizarGrid(clientes)
    End Sub
    
    Private Sub btnNovo_Click(sender As Object, e As EventArgs) Handles btnNovo.Click
        Using form As New FormCadastroCliente()
            If form.ShowDialog() = DialogResult.OK Then
                _manager.AdicionarCliente(form.Cliente)
                CarregarDados()
            End If
        End Using
    End Sub
    
    Private Sub btnEditar_Click(sender As Object, e As EventArgs) Handles btnEditar.Click
        If dgvClientes.CurrentRow IsNot Nothing Then
            Dim cliente = CType(dgvClientes.CurrentRow.Tag, Cliente)
            Using form As New FormCadastroCliente(cliente)
                If form.ShowDialog() = DialogResult.OK Then
                    _manager.AtualizarCliente(form.Cliente)
                    CarregarDados()
                End If
            End Using
        Else
            MessageBox.Show("Selecione um cliente para editar.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub
    
    Private Sub btnRemover_Click(sender As Object, e As EventArgs) Handles btnRemover.Click
        If dgvClientes.CurrentRow IsNot Nothing Then
            Dim cliente = CType(dgvClientes.CurrentRow.Tag, Cliente)
            If MessageBox.Show($"Deseja remover o cliente '{cliente.Nome}'?", "Confirmar", 
                              MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                _manager.RemoverCliente(cliente.ID)
                CarregarDados()
            End If
        Else
            MessageBox.Show("Selecione um cliente para remover.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub
    
    Private Sub btnRelatorios_Click(sender As Object, e As EventArgs) Handles btnRelatorios.Click
        ' TODO: Implementar relatórios de clientes
        MessageBox.Show("Relatórios em desenvolvimento.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
    
    Private Sub btnFechar_Click(sender As Object, e As EventArgs) Handles btnFechar.Click
        Me.Close()
    End Sub
    
    Private Sub dgvClientes_DoubleClick(sender As Object, e As EventArgs) Handles dgvClientes.DoubleClick
        btnEditar_Click(sender, e)
    End Sub
    
    Private Sub txtBusca_TextChanged(sender As Object, e As EventArgs) Handles txtBusca.TextChanged
        ' Busca em tempo real
        If txtBusca.Text.Length >= 2 Then
            btnBuscar_Click(sender, e)
        ElseIf txtBusca.Text.Length = 0 Then
            CarregarDados()
        End If
    End Sub
End Class

''' <summary>
''' Formulário de cadastro/edição de cliente
''' </summary>
Public Class FormCadastroCliente
    Inherits Form
    
    Private WithEvents txtNome As TextBox
    Private WithEvents txtEndereco As TextBox
    Private WithEvents txtCEP As TextBox
    Private WithEvents txtCidade As TextBox
    Private WithEvents cmbUF As ComboBox
    Private WithEvents txtTelefone As TextBox
    Private WithEvents txtEmail As TextBox
    Private WithEvents txtCPF_CNPJ As TextBox
    Private WithEvents chkAtivo As CheckBox
    Private WithEvents txtObservacoes As TextBox
    Private WithEvents btnSalvar As Button
    Private WithEvents btnCancelar As Button
    Private WithEvents btnCalendario As Button
    
    Private _cliente As Cliente
    Private _editMode As Boolean
    
    Public Property Cliente As Cliente
        Get
            Return _cliente
        End Get
        Set(value As Cliente)
            _cliente = value
        End Set
    End Property
    
    Public Sub New(Optional cliente As Cliente = Nothing)
        InitializeComponent()
        
        If cliente IsNot Nothing Then
            _cliente = New Cliente() With {
                .ID = cliente.ID,
                .Nome = cliente.Nome,
                .Endereco = cliente.Endereco,
                .CEP = cliente.CEP,
                .Cidade = cliente.Cidade,
                .UF = cliente.UF,
                .Telefone = cliente.Telefone,
                .Email = cliente.Email,
                .CPF_CNPJ = cliente.CPF_CNPJ,
                .DataCadastro = cliente.DataCadastro,
                .Ativo = cliente.Ativo,
                .Observacoes = cliente.Observacoes
            }
            _editMode = True
            Me.Text = "Editar Cliente"
        Else
            _cliente = New Cliente()
            _editMode = False
            Me.Text = "Novo Cliente"
        End If
        
        PreencherCampos()
    End Sub
    
    Private Sub InitializeComponent()
        Me.Size = New Size(500, 600)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.BackColor = Color.White
        
        ' Criar controles
        Dim y As Integer = 20
        
        ' Nome
        Me.Controls.Add(New Label() With {.Text = "Nome:", .Location = New Point(20, y), .AutoSize = True})
        y += 20
        txtNome = New TextBox() With {.Location = New Point(20, y), .Size = New Size(440, 25)}
        Me.Controls.Add(txtNome)
        y += 40
        
        ' Endereço
        Me.Controls.Add(New Label() With {.Text = "Endereço:", .Location = New Point(20, y), .AutoSize = True})
        y += 20
        txtEndereco = New TextBox() With {.Location = New Point(20, y), .Size = New Size(440, 25)}
        Me.Controls.Add(txtEndereco)
        y += 40
        
        ' CEP e Cidade
        Me.Controls.Add(New Label() With {.Text = "CEP:", .Location = New Point(20, y), .AutoSize = True})
        Me.Controls.Add(New Label() With {.Text = "Cidade:", .Location = New Point(150, y), .AutoSize = True})
        y += 20
        txtCEP = New TextBox() With {.Location = New Point(20, y), .Size = New Size(120, 25)}
        txtCidade = New TextBox() With {.Location = New Point(150, y), .Size = New Size(200, 25)}
        Me.Controls.AddRange({txtCEP, txtCidade})
        y += 40
        
        ' UF e Telefone
        Me.Controls.Add(New Label() With {.Text = "UF:", .Location = New Point(20, y), .AutoSize = True})
        Me.Controls.Add(New Label() With {.Text = "Telefone:", .Location = New Point(100, y), .AutoSize = True})
        y += 20
        cmbUF = New ComboBox() With {.Location = New Point(20, y), .Size = New Size(70, 25), .DropDownStyle = ComboBoxStyle.DropDownList}
        txtTelefone = New TextBox() With {.Location = New Point(100, y), .Size = New Size(150, 25)}
        Me.Controls.AddRange({cmbUF, txtTelefone})
        y += 40
        
        ' Email e CPF/CNPJ
        Me.Controls.Add(New Label() With {.Text = "Email:", .Location = New Point(20, y), .AutoSize = True})
        y += 20
        txtEmail = New TextBox() With {.Location = New Point(20, y), .Size = New Size(440, 25)}
        Me.Controls.Add(txtEmail)
        y += 40
        
        Me.Controls.Add(New Label() With {.Text = "CPF/CNPJ:", .Location = New Point(20, y), .AutoSize = True})
        y += 20
        txtCPF_CNPJ = New TextBox() With {.Location = New Point(20, y), .Size = New Size(200, 25)}
        Me.Controls.Add(txtCPF_CNPJ)
        y += 40
        
        ' Ativo
        chkAtivo = New CheckBox() With {.Text = "Cliente Ativo", .Location = New Point(20, y), .AutoSize = True, .Checked = True}
        Me.Controls.Add(chkAtivo)
        y += 40
        
        ' Observações
        Me.Controls.Add(New Label() With {.Text = "Observações:", .Location = New Point(20, y), .AutoSize = True})
        y += 20
        txtObservacoes = New TextBox() With {.Location = New Point(20, y), .Size = New Size(440, 60), .Multiline = True, .ScrollBars = ScrollBars.Vertical}
        Me.Controls.Add(txtObservacoes)
        y += 80
        
        ' Botões
        btnSalvar = New Button() With {
            .Text = "💾 Salvar",
            .Location = New Point(280, y),
            .Size = New Size(90, 30),
            .BackColor = Color.Green,
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
        
        btnCancelar = New Button() With {
            .Text = "❌ Cancelar",
            .Location = New Point(380, y),
            .Size = New Size(90, 30),
            .BackColor = Color.Gray,
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
        
        Me.Controls.AddRange({btnSalvar, btnCancelar})
        
        ' Carregar UFs
        CarregarUFs()
    End Sub
    
    Private Sub CarregarUFs()
        cmbUF.Items.AddRange({"AC", "AL", "AP", "AM", "BA", "CE", "DF", "ES", "GO", "MA", "MT", "MS", "MG", "PA", "PB", "PR", "PE", "PI", "RJ", "RN", "RS", "RO", "RR", "SC", "SP", "SE", "TO"})
        cmbUF.SelectedItem = "PE"
    End Sub
    
    Private Sub PreencherCampos()
        If _cliente IsNot Nothing Then
            txtNome.Text = _cliente.Nome
            txtEndereco.Text = _cliente.Endereco
            txtCEP.Text = _cliente.CEP
            txtCidade.Text = _cliente.Cidade
            cmbUF.SelectedItem = _cliente.UF
            txtTelefone.Text = _cliente.Telefone
            txtEmail.Text = _cliente.Email
            txtCPF_CNPJ.Text = _cliente.CPF_CNPJ
            chkAtivo.Checked = _cliente.Ativo
            txtObservacoes.Text = _cliente.Observacoes
        End If
    End Sub
    
    Private Sub btnSalvar_Click(sender As Object, e As EventArgs) Handles btnSalvar.Click
        Try
            ' Validações
            If String.IsNullOrEmpty(txtNome.Text.Trim()) Then
                MessageBox.Show("Nome é obrigatório.", "Validação", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtNome.Focus()
                Return
            End If
            
            ' Atualizar objeto cliente
            _cliente.Nome = txtNome.Text.Trim()
            _cliente.Endereco = txtEndereco.Text.Trim()
            _cliente.CEP = txtCEP.Text.Trim()
            _cliente.Cidade = txtCidade.Text.Trim()
            _cliente.UF = cmbUF.SelectedItem?.ToString()
            _cliente.Telefone = txtTelefone.Text.Trim()
            _cliente.Email = txtEmail.Text.Trim()
            _cliente.CPF_CNPJ = txtCPF_CNPJ.Text.Trim()
            _cliente.Ativo = chkAtivo.Checked
            _cliente.Observacoes = txtObservacoes.Text.Trim()
            
            Me.DialogResult = DialogResult.OK
            Me.Close()
            
        Catch ex As Exception
            MessageBox.Show($"Erro ao salvar cliente: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    Private Sub btnCancelar_Click(sender As Object, e As EventArgs) Handles btnCancelar.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub
End Class