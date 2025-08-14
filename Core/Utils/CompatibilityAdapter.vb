''' <summary>
''' Adaptador para compatibilidade entre as classes antigas e novas
''' Permite transição gradual para nova arquitetura
''' </summary>
Public Class CompatibilityAdapter
    
    ''' <summary>
    ''' Converte DadosTalao para Venda (nova arquitetura)
    ''' </summary>
    Public Shared Function ConvertToVenda(dadosTalao As DadosTalao) As Venda
        Try
            ' Criar cliente
            Dim cliente = New Cliente(dadosTalao.NomeCliente, dadosTalao.EnderecoCliente, dadosTalao.Telefone) With {
                .CEP = dadosTalao.CEP,
                .Cidade = dadosTalao.Cidade
            }
            
            ' Criar venda
            Dim venda = New Venda(cliente, dadosTalao.Vendedor) With {
                .NumeroTalao = dadosTalao.NumeroTalao,
                .DataVenda = dadosTalao.DataVenda,
                .FormaPagamento = dadosTalao.FormaPagamento
            }
            
            ' Adicionar itens
            For Each produtoTalao In dadosTalao.Produtos
                Dim produto = New Produto() With {
                    .Descricao = produtoTalao.Descricao,
                    .Unidade = produtoTalao.Unidade,
                    .PrecoUnitario = CDec(produtoTalao.PrecoUnitario)
                }
                
                Dim itemVenda = New ItemVenda(produto, CDec(produtoTalao.Quantidade))
                venda.Itens.Add(itemVenda)
            Next
            
            ' Calcular total
            venda.CalcularTotal()
            
            Return venda
            
        Catch ex As Exception
            Logger.Instance.Error("Erro ao converter DadosTalao para Venda", ex)
            Return Nothing
        End Try
    End Function
    
    ''' <summary>
    ''' Converte Venda para DadosTalao (compatibilidade reversa)
    ''' </summary>
    Public Shared Function ConvertToDadosTalao(venda As Venda) As DadosTalao
        Try
            Dim dadosTalao = New DadosTalao() With {
                .NomeCliente = venda.Cliente.Nome,
                .EnderecoCliente = venda.Cliente.Endereco,
                .CEP = venda.Cliente.CEP,
                .Cidade = venda.Cliente.Cidade,
                .Telefone = venda.Cliente.Telefone,
                .Vendedor = venda.Vendedor,
                .FormaPagamento = venda.FormaPagamento,
                .DataVenda = venda.DataVenda,
                .NumeroTalao = venda.NumeroTalao
            }
            
            ' Converter itens
            For Each item In venda.Itens
                Dim produtoTalao = New ProdutoTalao() With {
                    .Descricao = item.Produto.Descricao,
                    .Quantidade = CDbl(item.Quantidade),
                    .Unidade = item.Produto.Unidade,
                    .PrecoUnitario = CDbl(item.PrecoUnitario),
                    .PrecoTotal = CDbl(item.ValorTotal)
                }
                
                dadosTalao.Produtos.Add(produtoTalao)
            Next
            
            Return dadosTalao
            
        Catch ex As Exception
            Logger.Instance.Error("Erro ao converter Venda para DadosTalao", ex)
            Return Nothing
        End Try
    End Function
    
    ''' <summary>
    ''' Converte ProdutoTalao para Produto
    ''' </summary>
    Public Shared Function ConvertToProduto(produtoTalao As ProdutoTalao) As Produto
        Return New Produto() With {
            .Descricao = produtoTalao.Descricao,
            .Unidade = produtoTalao.Unidade,
            .PrecoUnitario = CDec(produtoTalao.PrecoUnitario)
        }
    End Function
    
    ''' <summary>
    ''' Valida dados do talão usando novo sistema de validação
    ''' </summary>
    Public Shared Function ValidarDadosTalao(dadosTalao As DadosTalao) As List(Of String)
        Dim erros = New List(Of String)()
        
        Try
            ' Validar cliente
            If String.IsNullOrWhiteSpace(dadosTalao.NomeCliente) Then
                erros.Add("Nome do cliente é obrigatório")
            End If
            
            If Not String.IsNullOrWhiteSpace(dadosTalao.CEP) AndAlso Not Validator.ValidarCEP(dadosTalao.CEP) Then
                erros.Add("CEP inválido")
            End If
            
            If Not String.IsNullOrWhiteSpace(dadosTalao.Telefone) AndAlso Not Validator.ValidarTelefone(dadosTalao.Telefone) Then
                erros.Add("Telefone inválido")
            End If
            
            ' Validar produtos
            If dadosTalao.Produtos Is Nothing OrElse dadosTalao.Produtos.Count = 0 Then
                erros.Add("Pelo menos um produto deve ser adicionado")
            Else
                For i = 0 To dadosTalao.Produtos.Count - 1
                    Dim produto = dadosTalao.Produtos(i)
                    
                    If String.IsNullOrWhiteSpace(produto.Descricao) Then
                        erros.Add($"Produto {i + 1}: Descrição é obrigatória")
                    End If
                    
                    If produto.Quantidade <= 0 Then
                        erros.Add($"Produto {i + 1}: Quantidade deve ser maior que zero")
                    End If
                    
                    If produto.PrecoUnitario <= 0 Then
                        erros.Add($"Produto {i + 1}: Preço unitário deve ser maior que zero")
                    End If
                Next
            End If
            
            ' Validar vendedor
            If String.IsNullOrWhiteSpace(dadosTalao.Vendedor) Then
                erros.Add("Vendedor é obrigatório")
            End If
            
        Catch ex As Exception
            erros.Add($"Erro na validação: {ex.Message}")
            Logger.Instance.Error("Erro ao validar dados do talão", ex)
        End Try
        
        Return erros
    End Function
    
    ''' <summary>
    ''' Formata dados do cliente automaticamente
    ''' </summary>
    Public Shared Sub FormatarDadosCliente(dadosTalao As DadosTalao)
        Try
            ' Formatar CEP
            If Not String.IsNullOrWhiteSpace(dadosTalao.CEP) Then
                dadosTalao.CEP = Validator.FormatarCEP(dadosTalao.CEP)
            End If
            
            ' Formatar telefone
            If Not String.IsNullOrWhiteSpace(dadosTalao.Telefone) Then
                dadosTalao.Telefone = Validator.FormatarTelefone(dadosTalao.Telefone)
            End If
            
            ' Capitalizar nome
            If Not String.IsNullOrWhiteSpace(dadosTalao.NomeCliente) Then
                dadosTalao.NomeCliente = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(dadosTalao.NomeCliente.ToLower())
            End If
            
        Catch ex As Exception
            Logger.Instance.Warning("Erro ao formatar dados do cliente", ex)
        End Try
    End Sub
End Class