Imports System.Data
Imports System.Data.OleDb
Imports System.Linq.Expressions

''' <summary>
''' Interface para repositório genérico
''' </summary>
Public Interface IRepository(Of T)
    Function GetById(id As Integer) As T
    Function GetAll() As List(Of T)
    Function Find(predicate As Func(Of T, Boolean)) As List(Of T)
    Function Add(entity As T) As Integer
    Function Update(entity As T) As Boolean
    Function Delete(id As Integer) As Boolean
    Function Count() As Integer
    Function Exists(id As Integer) As Boolean
End Interface

''' <summary>
''' Repositório base genérico com implementação comum
''' </summary>
Public MustInherit Class BaseRepository(Of T As Class)
    Implements IRepository(Of T)
    
    Protected ReadOnly _logger As LoggingSystem = LoggingSystem.Instance
    Protected ReadOnly _config As EnhancedConfigurationManager = EnhancedConfigurationManager.Instance
    Protected ReadOnly _connectionString As String
    
    Protected Sub New()
        _connectionString = _config.ConexaoBanco
        If String.IsNullOrEmpty(_connectionString) Then
            _connectionString = GetDefaultConnectionString()
        End If
    End Sub
    
    ''' <summary>
    ''' String de conexão padrão se não configurada
    ''' </summary>
    Protected Overridable Function GetDefaultConnectionString() As String
        Dim dbPath = Path.Combine(_config.CaminhoBackup, "PDV_Database.accdb")
        Return $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbPath};Persist Security Info=False;"
    End Function
    
    ''' <summary>
    ''' Nome da tabela no banco de dados
    ''' </summary>
    Protected MustOverride ReadOnly Property TableName As String
    
    ''' <summary>
    ''' Mapeia DataRow para entidade
    ''' </summary>
    Protected MustOverride Function MapFromDataRow(row As DataRow) As T
    
    ''' <summary>
    ''' Mapeia entidade para parâmetros SQL
    ''' </summary>
    Protected MustOverride Function MapToParameters(entity As T) As Dictionary(Of String, Object)
    
    ''' <summary>
    ''' Obtém ID da entidade
    ''' </summary>
    Protected MustOverride Function GetEntityId(entity As T) As Integer
    
    #Region "Implementação IRepository"
    
    Public Overridable Function GetById(id As Integer) As T Implements IRepository(Of T).GetById
        Try
            Using connection = CreateConnection()
                connection.Open()
                
                Dim sql = $"SELECT * FROM {TableName} WHERE ID = @id"
                Using command = New OleDbCommand(sql, connection)
                    command.Parameters.AddWithValue("@id", id)
                    
                    Using adapter = New OleDbDataAdapter(command)
                        Dim table = New DataTable()
                        adapter.Fill(table)
                        
                        If table.Rows.Count > 0 Then
                            Return MapFromDataRow(table.Rows(0))
                        End If
                    End Using
                End Using
            End Using
            
            Return Nothing
            
        Catch ex As Exception
            _logger.LogError($"Repository<{GetType(T).Name}>", $"Erro ao buscar por ID {id}", ex)
            Throw
        End Try
    End Function
    
    Public Overridable Function GetAll() As List(Of T) Implements IRepository(Of T).GetAll
        Try
            Dim results = New List(Of T)()
            
            Using connection = CreateConnection()
                connection.Open()
                
                Dim sql = $"SELECT * FROM {TableName} ORDER BY ID"
                Using command = New OleDbCommand(sql, connection)
                    Using adapter = New OleDbDataAdapter(command)
                        Dim table = New DataTable()
                        adapter.Fill(table)
                        
                        For Each row As DataRow In table.Rows
                            results.Add(MapFromDataRow(row))
                        Next
                    End Using
                End Using
            End Using
            
            _logger.LogDebug($"Repository<{GetType(T).Name}>", $"Retornados {results.Count} registros")
            Return results
            
        Catch ex As Exception
            _logger.LogError($"Repository<{GetType(T).Name}>", "Erro ao buscar todos os registros", ex)
            Throw
        End Try
    End Function
    
    Public Overridable Function Find(predicate As Func(Of T, Boolean)) As List(Of T) Implements IRepository(Of T).Find
        Try
            ' Para implementação simples, buscar todos e filtrar em memória
            ' Em implementação avançada, converter predicate para SQL
            Dim allItems = GetAll()
            Return allItems.Where(predicate).ToList()
            
        Catch ex As Exception
            _logger.LogError($"Repository<{GetType(T).Name}>", "Erro na busca com filtro", ex)
            Throw
        End Try
    End Function
    
    Public Overridable Function Add(entity As T) As Integer Implements IRepository(Of T).Add
        Try
            Using connection = CreateConnection()
                connection.Open()
                
                Dim parameters = MapToParameters(entity)
                Dim fields = String.Join(", ", parameters.Keys)
                Dim values = String.Join(", ", parameters.Keys.Select(Function(k) "@" & k))
                
                Dim sql = $"INSERT INTO {TableName} ({fields}) VALUES ({values})"
                
                Using command = New OleDbCommand(sql, connection)
                    For Each param In parameters
                        command.Parameters.AddWithValue("@" & param.Key, If(param.Value, DBNull.Value))
                    Next
                    
                    command.ExecuteNonQuery()
                    
                    ' Obter ID gerado
                    command.CommandText = "SELECT @@IDENTITY"
                    Dim newId = Convert.ToInt32(command.ExecuteScalar())
                    
                    _logger.LogInfo($"Repository<{GetType(T).Name}>", $"Registro adicionado com ID {newId}")
                    Return newId
                End Using
            End Using
            
        Catch ex As Exception
            _logger.LogError($"Repository<{GetType(T).Name}>", "Erro ao adicionar registro", ex)
            Throw
        End Try
    End Function
    
    Public Overridable Function Update(entity As T) As Boolean Implements IRepository(Of T).Update
        Try
            Using connection = CreateConnection()
                connection.Open()
                
                Dim parameters = MapToParameters(entity)
                Dim sets = String.Join(", ", parameters.Keys.Select(Function(k) $"{k} = @{k}"))
                Dim entityId = GetEntityId(entity)
                
                Dim sql = $"UPDATE {TableName} SET {sets} WHERE ID = @id"
                
                Using command = New OleDbCommand(sql, connection)
                    For Each param In parameters
                        command.Parameters.AddWithValue("@" & param.Key, If(param.Value, DBNull.Value))
                    Next
                    command.Parameters.AddWithValue("@id", entityId)
                    
                    Dim rowsAffected = command.ExecuteNonQuery()
                    
                    _logger.LogInfo($"Repository<{GetType(T).Name}>", $"Registro ID {entityId} atualizado")
                    Return rowsAffected > 0
                End Using
            End Using
            
        Catch ex As Exception
            _logger.LogError($"Repository<{GetType(T).Name}>", "Erro ao atualizar registro", ex)
            Throw
        End Try
    End Function
    
    Public Overridable Function Delete(id As Integer) As Boolean Implements IRepository(Of T).Delete
        Try
            Using connection = CreateConnection()
                connection.Open()
                
                Dim sql = $"DELETE FROM {TableName} WHERE ID = @id"
                Using command = New OleDbCommand(sql, connection)
                    command.Parameters.AddWithValue("@id", id)
                    
                    Dim rowsAffected = command.ExecuteNonQuery()
                    
                    _logger.LogInfo($"Repository<{GetType(T).Name}>", $"Registro ID {id} removido")
                    Return rowsAffected > 0
                End Using
            End Using
            
        Catch ex As Exception
            _logger.LogError($"Repository<{GetType(T).Name}>", $"Erro ao remover registro ID {id}", ex)
            Throw
        End Try
    End Function
    
    Public Overridable Function Count() As Integer Implements IRepository(Of T).Count
        Try
            Using connection = CreateConnection()
                connection.Open()
                
                Dim sql = $"SELECT COUNT(*) FROM {TableName}"
                Using command = New OleDbCommand(sql, connection)
                    Return Convert.ToInt32(command.ExecuteScalar())
                End Using
            End Using
            
        Catch ex As Exception
            _logger.LogError($"Repository<{GetType(T).Name}>", "Erro ao contar registros", ex)
            Throw
        End Try
    End Function
    
    Public Overridable Function Exists(id As Integer) As Boolean Implements IRepository(Of T).Exists
        Try
            Using connection = CreateConnection()
                connection.Open()
                
                Dim sql = $"SELECT COUNT(*) FROM {TableName} WHERE ID = @id"
                Using command = New OleDbCommand(sql, connection)
                    command.Parameters.AddWithValue("@id", id)
                    Return Convert.ToInt32(command.ExecuteScalar()) > 0
                End Using
            End Using
            
        Catch ex As Exception
            _logger.LogError($"Repository<{GetType(T).Name}>", $"Erro ao verificar existência ID {id}", ex)
            Throw
        End Try
    End Function
    
    #End Region
    
    #Region "Métodos Auxiliares"
    
    ''' <summary>
    ''' Cria conexão com o banco de dados
    ''' </summary>
    Protected Function CreateConnection() As OleDbConnection
        Return New OleDbConnection(_connectionString)
    End Function
    
    ''' <summary>
    ''' Executa comando SQL customizado
    ''' </summary>
    Protected Function ExecuteCustomQuery(sql As String, parameters As Dictionary(Of String, Object)) As DataTable
        Try
            Using connection = CreateConnection()
                connection.Open()
                
                Using command = New OleDbCommand(sql, connection)
                    If parameters IsNot Nothing Then
                        For Each param In parameters
                            command.Parameters.AddWithValue("@" & param.Key, If(param.Value, DBNull.Value))
                        Next
                    End If
                    
                    Using adapter = New OleDbDataAdapter(command)
                        Dim table = New DataTable()
                        adapter.Fill(table)
                        Return table
                    End Using
                End Using
            End Using
            
        Catch ex As Exception
            _logger.LogError($"Repository<{GetType(T).Name}>", "Erro ao executar query customizada", ex)
            Throw
        End Try
    End Function
    
    ''' <summary>
    ''' Verifica se a tabela existe no banco
    ''' </summary>
    Protected Function TableExists() As Boolean
        Try
            Using connection = CreateConnection()
                connection.Open()
                
                Dim tables = connection.GetSchema("Tables")
                For Each row As DataRow In tables.Rows
                    If row("TABLE_NAME").ToString().Equals(TableName, StringComparison.OrdinalIgnoreCase) Then
                        Return True
                    End If
                Next
            End Using
            
            Return False
            
        Catch ex As Exception
            _logger.LogWarning($"Repository<{GetType(T).Name}>", "Erro ao verificar existência da tabela", ex)
            Return False
        End Try
    End Function
    
    #End Region
End Class

''' <summary>
''' Repositório específico para Cliente
''' </summary>
Public Class ClienteRepository
    Inherits BaseRepository(Of Cliente)
    
    Protected Overrides ReadOnly Property TableName As String
        Get
            Return "Clientes"
        End Get
    End Property
    
    Protected Overrides Function MapFromDataRow(row As DataRow) As Cliente
        Return New Cliente() With {
            .ID = Convert.ToInt32(row("ID")),
            .Nome = row("Nome").ToString(),
            .Endereco = row("Endereco").ToString(),
            .CEP = row("CEP").ToString(),
            .Cidade = row("Cidade").ToString(),
            .UF = row("UF").ToString(),
            .Telefone = row("Telefone").ToString(),
            .Email = row("Email").ToString(),
            .CPF_CNPJ = row("CPF_CNPJ").ToString(),
            .DataCadastro = Convert.ToDateTime(row("DataCadastro")),
            .Ativo = Convert.ToBoolean(row("Ativo")),
            .Observacoes = row("Observacoes").ToString()
        }
    End Function
    
    Protected Overrides Function MapToParameters(entity As Cliente) As Dictionary(Of String, Object)
        Return New Dictionary(Of String, Object) From {
            {"Nome", entity.Nome},
            {"Endereco", entity.Endereco},
            {"CEP", entity.CEP},
            {"Cidade", entity.Cidade},
            {"UF", entity.UF},
            {"Telefone", entity.Telefone},
            {"Email", entity.Email},
            {"CPF_CNPJ", entity.CPF_CNPJ},
            {"DataCadastro", entity.DataCadastro},
            {"Ativo", entity.Ativo},
            {"Observacoes", entity.Observacoes}
        }
    End Function
    
    Protected Overrides Function GetEntityId(entity As Cliente) As Integer
        Return entity.ID
    End Function
    
    ''' <summary>
    ''' Busca clientes por nome
    ''' </summary>
    Public Function BuscarPorNome(nome As String) As List(Of Cliente)
        Dim sql = "SELECT * FROM Clientes WHERE Nome LIKE @nome ORDER BY Nome"
        Dim parameters = New Dictionary(Of String, Object) From {{"nome", $"%{nome}%"}}
        
        Dim table = ExecuteCustomQuery(sql, parameters)
        Dim results = New List(Of Cliente)()
        
        For Each row As DataRow In table.Rows
            results.Add(MapFromDataRow(row))
        Next
        
        Return results
    End Function
End Class

''' <summary>
''' Repositório específico para Produto
''' </summary>
Public Class ProdutoRepository
    Inherits BaseRepository(Of Produto)
    
    Protected Overrides ReadOnly Property TableName As String
        Get
            Return "Produtos"
        End Get
    End Property
    
    Protected Overrides Function MapFromDataRow(row As DataRow) As Produto
        Return New Produto() With {
            .ID = Convert.ToInt32(row("ID")),
            .Codigo = row("Codigo").ToString(),
            .Descricao = row("Descricao").ToString(),
            .Secao = row("Secao").ToString(),
            .Unidade = row("Unidade").ToString(),
            .PrecoVenda = Convert.ToDecimal(row("PrecoVenda")),
            .PrecoCusto = Convert.ToDecimal(row("PrecoCusto")),
            .EstoqueAtual = Convert.ToDouble(row("EstoqueAtual")),
            .EstoqueMinimo = Convert.ToDouble(row("EstoqueMinimo")),
            .Ativo = Convert.ToBoolean(row("Ativo")),
            .DataCadastro = Convert.ToDateTime(row("DataCadastro")),
            .Observacoes = row("Observacoes").ToString()
        }
    End Function
    
    Protected Overrides Function MapToParameters(entity As Produto) As Dictionary(Of String, Object)
        Return New Dictionary(Of String, Object) From {
            {"Codigo", entity.Codigo},
            {"Descricao", entity.Descricao},
            {"Secao", entity.Secao},
            {"Unidade", entity.Unidade},
            {"PrecoVenda", entity.PrecoVenda},
            {"PrecoCusto", entity.PrecoCusto},
            {"EstoqueAtual", entity.EstoqueAtual},
            {"EstoqueMinimo", entity.EstoqueMinimo},
            {"Ativo", entity.Ativo},
            {"DataCadastro", entity.DataCadastro},
            {"Observacoes", entity.Observacoes}
        }
    End Function
    
    Protected Overrides Function GetEntityId(entity As Produto) As Integer
        Return entity.ID
    End Function
    
    ''' <summary>
    ''' Busca produtos por descrição ou código
    ''' </summary>
    Public Function BuscarPorDescricaoOuCodigo(termo As String) As List(Of Produto)
        Dim sql = "SELECT * FROM Produtos WHERE (Descricao LIKE @termo OR Codigo LIKE @termo) AND Ativo = True ORDER BY Descricao"
        Dim parameters = New Dictionary(Of String, Object) From {{"termo", $"%{termo}%"}}
        
        Dim table = ExecuteCustomQuery(sql, parameters)
        Dim results = New List(Of Produto)()
        
        For Each row As DataRow In table.Rows
            results.Add(MapFromDataRow(row))
        Next
        
        Return results
    End Function
    
    ''' <summary>
    ''' Busca produtos com estoque baixo
    ''' </summary>
    Public Function BuscarEstoqueBaixo() As List(Of Produto)
        Dim sql = "SELECT * FROM Produtos WHERE EstoqueAtual <= EstoqueMinimo AND Ativo = True ORDER BY EstoqueAtual"
        
        Dim table = ExecuteCustomQuery(sql, Nothing)
        Dim results = New List(Of Produto)()
        
        For Each row As DataRow In table.Rows
            results.Add(MapFromDataRow(row))
        Next
        
        Return results
    End Function
End Class