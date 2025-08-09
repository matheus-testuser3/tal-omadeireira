# Manual do Usuário - Sistema PDV Madeireira Maria Luiza

## Introdução

O Sistema PDV da Madeireira Maria Luiza é uma solução completa para gerenciamento de vendas, desenvolvida especialmente para atender às necessidades da empresa. Este manual fornece instruções detalhadas sobre como utilizar todas as funcionalidades do sistema.

## Iniciando o Sistema

### 1. Abertura do Sistema
- Execute o arquivo `SistemaPDV.exe`
- A tela principal será exibida com o menu lateral e dashboard

### 2. Interface Principal
A interface principal contém:
- **Menu lateral** com opções de navegação
- **Dashboard** com informações em tempo real
- **Cards informativos** com dados de vendas, estoque, clientes e faturamento

## Realizando uma Venda

### Passo 1: Acessar o PDV
1. Clique no botão **"🏪 PDV / CAIXA"** no menu lateral
2. O formulário de entrada de dados será aberto

### Passo 2: Preencher Dados do Cliente
1. **Nome do Cliente** (obrigatório)
   - Digite o nome completo do cliente
   
2. **Endereço** (opcional)
   - Digite o endereço completo
   - Inclua rua, número, bairro
   
3. **Cidade** (pré-preenchido)
   - Valor padrão: "Paulista"
   - Pode ser alterado se necessário
   
4. **CEP** (pré-preenchido)
   - Valor padrão: "55431-165"
   - Pode ser alterado se necessário

### Passo 3: Informar Produtos/Serviços
1. **Produtos/Serviços** (obrigatório)
   - Digite a descrição detalhada dos itens vendidos
   - Exemplo: "Tábua de madeira 2x4x3m - 5 unidades"
   - Use múltiplas linhas se necessário

### Passo 4: Definir Valores e Pagamento
1. **Valor Total** (obrigatório)
   - Digite apenas o valor numérico
   - Exemplo: "125,50" (sem R$)
   
2. **Forma de Pagamento**
   - Selecione uma opção:
     - Dinheiro
     - Cartão Débito
     - Cartão Crédito
     - PIX
     - Cheque
     - Fiado

3. **Vendedor** (pré-preenchido)
   - Valor padrão: "matheus-testuser3"
   - Pode ser alterado conforme necessário

### Passo 5: Gerar o Talão
1. Clique no botão **"🖨️ GERAR TALÃO"**
2. O sistema validará os dados obrigatórios
3. Se algum campo obrigatório estiver vazio, será exibida uma mensagem
4. Com dados válidos, o sistema processará o talão

## Processo de Geração do Talão

### O que acontece automaticamente:
1. **Validação dos dados** informados
2. **Criação do template Excel** em tempo real
3. **Preenchimento automático** de todos os campos
4. **Formatação profissional** do talão
5. **Geração de via dupla** (cliente + vendedor)
6. **Configuração para impressão** em A4 paisagem
7. **Envio para impressora** ou visualização prévia

### Informações incluídas no talão:
- **Cabeçalho da empresa**
- **Dados do cliente**
- **Produtos/serviços vendidos**
- **Valores e forma de pagamento**
- **Data da venda**
- **Nome do vendedor**
- **Espaço para assinatura**
- **Contatos da empresa**

## Funcionalidades dos Botões

### Botão "🖨️ GERAR TALÃO"
- Valida os dados inseridos
- Gera e imprime o talão duplo
- Limpa o formulário automaticamente após sucesso

### Botão "🗑️ LIMPAR"
- Limpa todos os campos do formulário
- Restaura valores padrão (cidade, CEP, vendedor)
- Posiciona cursor no campo "Nome do Cliente"

### Botão "❌ FECHAR"
- Fecha o formulário de PDV
- Retorna à tela principal

## Validações do Sistema

### Campos Obrigatórios:
- ✅ **Nome do Cliente**: Não pode estar vazio
- ✅ **Produtos/Serviços**: Deve conter descrição
- ✅ **Valor Total**: Deve ser maior que 0,00

### Campos Opcionais:
- Endereço do cliente
- Cidade (com valor padrão)
- CEP (com valor padrão)

## Dicas de Uso

### Para melhor aproveitamento:
1. **Mantenha dados atualizados**: Cidade, CEP e vendedor padrão
2. **Seja específico**: Descreva produtos com detalhes (tamanho, quantidade)
3. **Confira valores**: Verifique valor total antes de gerar talão
4. **Use atalhos**: Tab para navegar entre campos

### Exemplos de preenchimento:

**Produtos/Serviços:**
```
- Tábua de madeira 2x4x3m - 10 unidades
- Compensado 15mm 1,22x2,44m - 3 chapas
- Serviço de corte personalizado
```

**Valor Total:**
```
Correto: 150,75
Incorreto: R$ 150,75 ou 150.75
```

## Resolução de Problemas

### Problema: "Excel não encontrado"
**Solução**: Instale o Microsoft Excel no computador

### Problema: "Erro ao imprimir"
**Solução**: 
- Verifique se a impressora está conectada
- Configure uma impressora padrão no Windows
- Teste a impressora com outro documento

### Problema: "Campos não preenchem"
**Solução**:
- Clique dentro do campo antes de digitar
- Aguarde o placeholder desaparecer
- Use Tab para navegar entre campos

### Problema: "Talão não formatou corretamente"
**Solução**:
- Verifique se o Excel está fechado antes de gerar novo talão
- Reinicie o sistema se necessário

## Informações Técnicas

### Requisitos:
- Windows 10/11
- Microsoft Excel 2016 ou superior
- Impressora configurada
- .NET Framework 4.7.2+

### Arquivos importantes:
- `SistemaPDV.exe` - Executável principal
- `App.config` - Configurações do sistema
- Pasta `Backup` - Backup automático (se habilitado)
- Pasta `Historico` - Histórico de vendas (se habilitado)

## Contato e Suporte

### Para suporte técnico:
- **Desenvolvedor**: matheus-testuser3
- **Empresa**: Madeireira Maria Luiza
- **Telefone**: (81) 98570-1522
- **WhatsApp**: (81) 98570-1522

### Para melhorias e sugestões:
- Utilize o sistema de issues no GitHub
- Entre em contato diretamente com o desenvolvedor

---

**Versão do Manual**: 1.0  
**Data de Atualização**: Dezembro 2024  
**Sistema**: PDV Madeireira Maria Luiza v1.0