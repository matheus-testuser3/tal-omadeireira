Imports System.Data.OleDb
Imports System.Data
Imports System.IO

''' <summary>
''' Sistema de conexão inteligente com banco de dados
''' Fallback automático para planilhas Excel quando Access não disponível
''' </summary>
Public Class DatabaseManager
    Private Shared _instance As DatabaseManager
    Private _connectionString As String
    Private _useAccessDatabase As Boolean
    Private _excelFallbackPath As String
    Private _config As ConfiguracaoSistema
    
    ''' <summary>
    ''' Singleton instance
    ''' </summary>
    Public Shared ReadOnly Property Instance As DatabaseManager
        Get
            If _instance Is Nothing Then
                _instance = New DatabaseManager()
            End If
            Return _instance
        End Get
    End Property
    
    Private Sub New()
        _config = New ConfiguracaoSistema()
        InicializarConexao()
    End Sub
    
    ''' <summary>
    ''' Inicializa a conexão com banco de dados ou Excel
    ''' </summary>
    Private Sub InicializarConexao()
        Try
            ' Tentar conexão com Access primeiro
            If _config.UsarBancoAccess AndAlso Not String.IsNullOrEmpty(_config.ConexaoBanco) Then
                _connectionString = _config.ConexaoBanco
                If TestarConexaoAccess() Then
                    _useAccessDatabase = True
                    Console.WriteLine("Conectado ao banco Access com sucesso")
                    Return
                End If
            End If
            
            ' Fallback para Excel
            _useAccessDatabase = False
            _excelFallbackPath = Path.Combine(Application.StartupPath, "PDV_Data.xlsx")
            CriarPlanilhaFallback()
            Console.WriteLine("Usando fallback para planilha Excel")
            
        Catch ex As Exception
            Console.WriteLine($"Erro na inicialização do banco: {ex.Message}")
            _useAccessDatabase = False
        End Try
    End Sub
    
    ''' <summary>
    ''' Testa conexão com banco Access
    ''' </summary>
    Private Function TestarConexaoAccess() As Boolean
        Try
            Using conn As New OleDbConnection(_connectionString)
                conn.Open()
                conn.Close()
                Return True
            End Using
        Catch
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' Cria planilha Excel para fallback
    ''' </summary>
    Private Sub CriarPlanilhaFallback()
        Try
            If File.Exists(_excelFallbackPath) Then Return
            
            Dim xlApp As Object = CreateObject("Excel.Application")
            Dim xlWorkbook As Object = xlApp.Workbooks.Add()
            
            ' Criar planilhas
            CriarPlanilhaClientes(xlWorkbook)
            CriarPlanilhaProdutos(xlWorkbook)
            CriarPlanilhaVendas(xlWorkbook)
            CriarPlanilhaVendedores(xlWorkbook)
            
            xlWorkbook.SaveAs(_excelFallbackPath)
            xlWorkbook.Close()
            xlApp.Quit()
            
        Catch ex As Exception
            Console.WriteLine($"Erro ao criar planilha de fallback: {ex.Message}")
        End Try
    End Sub
    
    ''' <summary>
    ''' Cria planilha de clientes
    ''' </summary>
    Private Sub CriarPlanilhaClientes(workbook As Object)
        Dim ws As Object = workbook.Worksheets.Add()
        ws.Name = "Clientes"
        
        ' Cabeçalhos
        ws.Cells(1, 1).Value = "ID"
        ws.Cells(1, 2).Value = "Nome"
        ws.Cells(1, 3).Value = "Endereco"
        ws.Cells(1, 4).Value = "CEP"
        ws.Cells(1, 5).Value = "Cidade"
        ws.Cells(1, 6).Value = "UF"
        ws.Cells(1, 7).Value = "Telefone"
        ws.Cells(1, 8).Value = "Email"
        ws.Cells(1, 9).Value = "CPF_CNPJ"
        ws.Cells(1, 10).Value = "DataCadastro"
        ws.Cells(1, 11).Value = "Ativo"
        ws.Cells(1, 12).Value = "Observacoes"
        
        ' Formatação
        ws.Range("A1:L1").Font.Bold = True
        ws.Range("A1:L1").Interior.Color = RGB(200, 200, 200)
    End Sub
    
    ''' <summary>
    ''' Cria planilha de produtos
    ''' </summary>
    Private Sub CriarPlanilhaProdutos(workbook As Object)
        Dim ws As Object = workbook.Worksheets.Add()
        ws.Name = "Produtos"
        
        ' Cabeçalhos
        ws.Cells(1, 1).Value = "ID"
        ws.Cells(1, 2).Value = "Codigo"
        ws.Cells(1, 3).Value = "Descricao"
        ws.Cells(1, 4).Value = "Secao"
        ws.Cells(1, 5).Value = "Unidade"
        ws.Cells(1, 6).Value = "PrecoVenda"
        ws.Cells(1, 7).Value = "PrecoCusto"
        ws.Cells(1, 8).Value = "EstoqueAtual"
        ws.Cells(1, 9).Value = "EstoqueMinimo"
        ws.Cells(1, 10).Value = "Ativo"
        ws.Cells(1, 11).Value = "DataCadastro"
        ws.Cells(1, 12).Value = "Observacoes"
        
        ' Formatação
        ws.Range("A1:L1").Font.Bold = True
        ws.Range("A1:L1").Interior.Color = RGB(200, 200, 200)
        
        ' Dados iniciais
        AdicionarProdutosIniciais(ws)
    End Sub
    
    ''' <summary>
    ''' Adiciona produtos iniciais para teste
    ''' </summary>
    Private Sub AdicionarProdutosIniciais(ws As Object)
        Dim produtos() As Object = {
            {1, "TAB001", "Tábua de Pinus 2x4m", "Madeiras", "UN", 25.00, 18.00, 50, 10, True, Date.Now, ""},
            {2, "RIP001", "Ripão 3x3x3m", "Madeiras", "UN", 15.00, 10.00, 100, 20, True, Date.Now, ""},
            {3, "COM001", "Compensado 18mm", "Chapas", "M²", 45.00, 32.00, 25, 5, True, Date.Now, ""},
            {4, "VIG001", "Viga de Eucalipto 6x12", "Estruturas", "UN", 85.00, 60.00, 30, 5, True, Date.Now, ""},
            {5, "CAI001", "Caibro 5x7x3m", "Estruturas", "UN", 18.00, 12.00, 80, 15, True, Date.Now, ""}
        }
        
        For i = 0 To produtos.GetUpperBound(0)
            For j = 0 To produtos.GetUpperBound(1)
                ws.Cells(i + 2, j + 1).Value = produtos(i, j)
            Next
        Next
    End Sub
    
    ''' <summary>
    ''' Cria planilha de vendas
    ''' </summary>
    Private Sub CriarPlanilhaVendas(workbook As Object)
        Dim ws As Object = workbook.Worksheets.Add()
        ws.Name = "Vendas"
        
        ' Cabeçalhos
        ws.Cells(1, 1).Value = "ID"
        ws.Cells(1, 2).Value = "NumeroTalao"
        ws.Cells(1, 3).Value = "ClienteID"
        ws.Cells(1, 4).Value = "DataVenda"
        ws.Cells(1, 5).Value = "FormaPagamento"
        ws.Cells(1, 6).Value = "VendedorID"
        ws.Cells(1, 7).Value = "VendedorNome"
        ws.Cells(1, 8).Value = "DescontoGeral"
        ws.Cells(1, 9).Value = "Frete"
        ws.Cells(1, 10).Value = "Status"
        ws.Cells(1, 11).Value = "Observacoes"
        
        ' Formatação
        ws.Range("A1:K1").Font.Bold = True
        ws.Range("A1:K1").Interior.Color = RGB(200, 200, 200)
    End Sub
    
    ''' <summary>
    ''' Cria planilha de vendedores
    ''' </summary>
    Private Sub CriarPlanilhaVendedores(workbook As Object)
        Dim ws As Object = workbook.Worksheets.Add()
        ws.Name = "Vendedores"
        
        ' Cabeçalhos
        ws.Cells(1, 1).Value = "ID"
        ws.Cells(1, 2).Value = "Nome"
        ws.Cells(1, 3).Value = "Usuario"
        ws.Cells(1, 4).Value = "Email"
        ws.Cells(1, 5).Value = "Telefone"
        ws.Cells(1, 6).Value = "Comissao"
        ws.Cells(1, 7).Value = "Ativo"
        ws.Cells(1, 8).Value = "DataCadastro"
        
        ' Formatação
        ws.Range("A1:H1").Font.Bold = True
        ws.Range("A1:H1").Interior.Color = RGB(200, 200, 200)
        
        ' Vendedor padrão
        ws.Cells(2, 1).Value = 1
        ws.Cells(2, 2).Value = _config.VendedorPadrao
        ws.Cells(2, 3).Value = _config.VendedorPadrao.ToLower()
        ws.Cells(2, 4).Value = ""
        ws.Cells(2, 5).Value = ""
        ws.Cells(2, 6).Value = 0
        ws.Cells(2, 7).Value = True
        ws.Cells(2, 8).Value = Date.Now
    End Sub
    
    #Region "Métodos Públicos de Acesso a Dados"
    
    ''' <summary>
    ''' Busca produtos por termo
    ''' </summary>
    Public Function BuscarProdutos(termo As String, Optional secao As String = "") As List(Of Produto)
        If _useAccessDatabase Then
            Return BuscarProdutosAccess(termo, secao)
        Else
            Return BuscarProdutosExcel(termo, secao)
        End If
    End Function
    
    ''' <summary>
    ''' Busca todos os produtos
    ''' </summary>
    Public Function ObterTodosProdutos() As List(Of Produto)
        Return BuscarProdutos("")
    End Function
    
    ''' <summary>
    ''' Busca produtos no Excel
    ''' </summary>
    Private Function BuscarProdutosExcel(termo As String, secao As String) As List(Of Produto)
        Dim produtos As New List(Of Produto)()
        
        Try
            If Not File.Exists(_excelFallbackPath) Then
                CriarPlanilhaFallback()
            End If
            
            Dim xlApp As Object = CreateObject("Excel.Application")
            Dim xlWorkbook As Object = xlApp.Workbooks.Open(_excelFallbackPath)
            Dim xlWorksheet As Object = xlWorkbook.Worksheets("Produtos")
            
            Dim lastRow As Integer = xlWorksheet.Cells(xlWorksheet.Rows.Count, 1).End(-4162).Row ' xlUp = -4162
            
            For i = 2 To lastRow
                Dim produto As New Produto() With {
                    .ID = xlWorksheet.Cells(i, 1).Value,
                    .Codigo = xlWorksheet.Cells(i, 2).Value,
                    .Descricao = xlWorksheet.Cells(i, 3).Value,
                    .Secao = xlWorksheet.Cells(i, 4).Value,
                    .Unidade = xlWorksheet.Cells(i, 5).Value,
                    .PrecoVenda = xlWorksheet.Cells(i, 6).Value,
                    .PrecoCusto = xlWorksheet.Cells(i, 7).Value,
                    .EstoqueAtual = xlWorksheet.Cells(i, 8).Value,
                    .EstoqueMinimo = xlWorksheet.Cells(i, 9).Value,
                    .Ativo = xlWorksheet.Cells(i, 10).Value
                }
                
                ' Filtrar por termo e seção
                If (String.IsNullOrEmpty(termo) OrElse 
                    produto.Descricao.ToUpper().Contains(termo.ToUpper()) OrElse
                    produto.Codigo.ToUpper().Contains(termo.ToUpper())) AndAlso
                   (String.IsNullOrEmpty(secao) OrElse produto.Secao.Equals(secao, StringComparison.OrdinalIgnoreCase)) Then
                    produtos.Add(produto)
                End If
            Next
            
            xlWorkbook.Close(False)
            xlApp.Quit()
            
        Catch ex As Exception
            Console.WriteLine($"Erro ao buscar produtos no Excel: {ex.Message}")
        End Try
        
        Return produtos
    End Function
    
    ''' <summary>
    ''' Busca produtos no Access
    ''' </summary>
    Private Function BuscarProdutosAccess(termo As String, secao As String) As List(Of Produto)
        Dim produtos As New List(Of Produto)()
        
        Try
            Using conn As New OleDbConnection(_connectionString)
                conn.Open()
                
                Dim sql As String = "SELECT * FROM Produtos WHERE Ativo = True"
                If Not String.IsNullOrEmpty(termo) Then
                    sql &= " AND (Descricao LIKE @termo OR Codigo LIKE @termo)"
                End If
                If Not String.IsNullOrEmpty(secao) Then
                    sql &= " AND Secao = @secao"
                End If
                
                Using cmd As New OleDbCommand(sql, conn)
                    If Not String.IsNullOrEmpty(termo) Then
                        cmd.Parameters.AddWithValue("@termo", $"%{termo}%")
                    End If
                    If Not String.IsNullOrEmpty(secao) Then
                        cmd.Parameters.AddWithValue("@secao", secao)
                    End If
                    
                    Using reader As OleDbDataReader = cmd.ExecuteReader()
                        While reader.Read()
                            produtos.Add(New Produto() With {
                                .ID = reader("ID"),
                                .Codigo = reader("Codigo").ToString(),
                                .Descricao = reader("Descricao").ToString(),
                                .Secao = reader("Secao").ToString(),
                                .Unidade = reader("Unidade").ToString(),
                                .PrecoVenda = reader("PrecoVenda"),
                                .PrecoCusto = reader("PrecoCusto"),
                                .EstoqueAtual = reader("EstoqueAtual"),
                                .EstoqueMinimo = reader("EstoqueMinimo"),
                                .Ativo = reader("Ativo")
                            })
                        End While
                    End Using
                End Using
            End Using
            
        Catch ex As Exception
            Console.WriteLine($"Erro ao buscar produtos no Access: {ex.Message}")
        End Try
        
        Return produtos
    End Function
    
    ''' <summary>
    ''' Obtém todas as seções de produtos
    ''' </summary>
    Public Function ObterSecoesProdutos() As List(Of String)
        Dim secoes As New List(Of String)()
        
        Try
            Dim produtos = ObterTodosProdutos()
            secoes = produtos.Select(Function(p) p.Secao).Distinct().ToList()
        Catch ex As Exception
            Console.WriteLine($"Erro ao obter seções: {ex.Message}")
        End Try
        
        Return secoes
    End Function
    
    ''' <summary>
    ''' Salva uma venda
    ''' </summary>
    Public Function SalvarVenda(venda As Venda) As Boolean
        If _useAccessDatabase Then
            Return SalvarVendaAccess(venda)
        Else
            Return SalvarVendaExcel(venda)
        End If
    End Function
    
    ''' <summary>
    ''' Salva venda no Excel
    ''' </summary>
    Private Function SalvarVendaExcel(venda As Venda) As Boolean
        Try
            ' Implementar salvamento no Excel
            Return True
        Catch ex As Exception
            Console.WriteLine($"Erro ao salvar venda no Excel: {ex.Message}")
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' Salva venda no Access
    ''' </summary>
    Private Function SalvarVendaAccess(venda As Venda) As Boolean
        Try
            ' Implementar salvamento no Access
            Return True
        Catch ex As Exception
            Console.WriteLine($"Erro ao salvar venda no Access: {ex.Message}")
            Return False
        End Try
    End Function
    
    #End Region
    
    ''' <summary>
    ''' Verifica status da conexão
    ''' </summary>
    Public Function VerificarConexao() As String
        If _useAccessDatabase Then
            If TestarConexaoAccess() Then
                Return "Conectado ao banco Access"
            Else
                Return "Falha na conexão com Access"
            End If
        Else
            If File.Exists(_excelFallbackPath) Then
                Return "Usando planilha Excel como banco"
            Else
                Return "Arquivo Excel não encontrado"
            End If
        End If
    End Function
End Class