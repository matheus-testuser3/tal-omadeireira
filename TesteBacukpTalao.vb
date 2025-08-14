''' <summary>
''' Script de teste para o sistema de backup de talões
''' Data/Hora: 2025-08-14 11:16:26 UTC
''' Usuário: matheus-testuser3
''' </summary>

Imports System.IO

Module TesteBacukpTalao
    
    Sub Main()
        Console.WriteLine("=== TESTE DO SISTEMA DE BACKUP DE TALÕES ===")
        Console.WriteLine($"Data/Hora: {DateTime.UtcNow:yyyy-MM-dd HH:mm:ss} UTC")
        Console.WriteLine($"Usuário: matheus-testuser3")
        Console.WriteLine()
        
        Try
            ' Teste 1: Criar instância do módulo de backup
            Console.WriteLine("1. Testando criação do módulo de backup...")
            Dim moduloBackup As New ModuloBackupTalao()
            Console.WriteLine("✅ Módulo de backup criado com sucesso")
            
            ' Teste 2: Criar dados de teste
            Console.WriteLine("2. Testando classes de dados...")
            Dim talao As New DadosTalaoMadeireira()
            talao.NumeroTalao = "TESTE001"
            talao.NomeCliente = "Cliente Teste"
            talao.EnderecoCliente = "Rua Teste, 123"
            talao.CEP = "50000-000"
            talao.Cidade = "Recife/PE"
            talao.Telefone = "(81) 9999-8888"
            talao.Vendedor = "matheus-testuser3"
            talao.FormaPagamento = "Dinheiro"
            
            Dim produto As New ProdutoTalaoMadeireira()
            produto.Descricao = "Tábua Teste"
            produto.Categoria = "Tábuas"
            produto.TipoMadeira = "Pinus"
            produto.Dimensoes = "2x4cm"
            produto.Comprimento = "3m"
            produto.Quantidade = 10
            produto.Unidade = "pc"
            produto.PrecoUnitario = 25.50D
            
            talao.Produtos.Add(produto)
            
            Console.WriteLine($"✅ Talão teste criado: {talao.ResumoDescricao}")
            Console.WriteLine($"   Produto: {produto.DescricaoCompleta}")
            Console.WriteLine($"   Total: {talao.ValorTotal:C2}")
            
            ' Teste 3: Validação dos dados
            Console.WriteLine("3. Testando validação de dados...")
            Dim erros = talao.ValidarDados()
            If erros.Count = 0 Then
                Console.WriteLine("✅ Dados válidos")
            Else
                Console.WriteLine($"❌ Erros encontrados: {String.Join(", ", erros)}")
            End If
            
            ' Teste 4: Serialização JSON
            Console.WriteLine("4. Testando serialização JSON...")
            Dim json = Newtonsoft.Json.JsonConvert.SerializeObject(talao, Newtonsoft.Json.Formatting.Indented)
            Console.WriteLine($"✅ JSON gerado ({json.Length} caracteres)")
            
            ' Teste 5: Configuração
            Console.WriteLine("5. Testando configuração...")
            Dim config As New ConfiguracaoBackupMadeireira()
            Console.WriteLine($"✅ Configuração carregada:")
            Console.WriteLine($"   Backups: {config.CaminhoBackupsImportados}")
            Console.WriteLine($"   Talões: {config.CaminhoTaloesGerados}")
            Console.WriteLine($"   JSON: {config.CaminhoBackupJSON}")
            
            ' Teste 6: Verificar diretórios
            Console.WriteLine("6. Verificando diretórios...")
            Dim diretorios = {config.CaminhoBackupsImportados, config.CaminhoTaloesGerados, config.CaminhoBackupJSON}
            For Each dir In diretorios
                If Directory.Exists(dir) Then
                    Console.WriteLine($"✅ Diretório existe: {dir}")
                Else
                    Console.WriteLine($"⚠️ Diretório não existe: {dir}")
                End If
            Next
            
            Console.WriteLine()
            Console.WriteLine("=== TESTE CONCLUÍDO COM SUCESSO ===")
            
        Catch ex As Exception
            Console.WriteLine($"❌ ERRO no teste: {ex.Message}")
            Console.WriteLine($"Stack Trace: {ex.StackTrace}")
        End Try
        
        Console.WriteLine()
        Console.WriteLine("Pressione qualquer tecla para continuar...")
        Console.ReadKey()
    End Sub
    
End Module