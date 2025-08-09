# Especificação Técnica - Sistema PDV

## 📐 Arquitetura do Sistema

### Componentes Principais

#### 1. **SistemaPDV.vb** (MainForm)
- **Função:** Interface principal do sistema
- **Responsabilidades:**
  - Exibir menu lateral moderno
  - Gerenciar navegação entre funcionalidades
  - Controlar fluxo principal de geração de talões
  - Verificar instalação do Excel
- **Eventos principais:**
  - `btnGerarTalao_Click()` - Abre formulário de dados
  - `ProcessarTalao()` - Coordena geração do talão

#### 2. **FormPDV.vb** (Data Entry Form)
- **Função:** Coleta de dados do cliente e produtos
- **Controles principais:**
  - Campos de cliente (nome, endereço, CEP, cidade, telefone)
  - Grid de produtos com adicionar/remover
  - Forma de pagamento e vendedor
  - Botão de dados de teste
- **Validações:**
  - Campos obrigatórios
  - Formato de dados numéricos
  - Pelo menos um produto

#### 3. **ExcelAutomation.vb** (Core Engine)
- **Função:** Automação completa do Excel
- **Processo principal:**
  ```
  AbrirExcel() → CriarPlanilhaTemporaria() → InjetarModulosVBA() → 
  CriarTemplate() → PreencherDados() → ConfigurarImpressao() → 
  ImprimirTalao() → FecharExcel()
  ```
- **Recursos:**
  - Excel invisível em background
  - Gestão automática de recursos COM
  - Tratamento de erros robusto
  - Cleanup automático

#### 4. **ModuloTalao.vb** (VBA Core)
- **Função:** Código VBA para geração de talões
- **Principais funções VBA:**
  - `ProcessarTalaoCompleto()` - Função principal
  - `GerarTalaoCompleto()` - Layout completo
  - `CriarCabecalhoEmpresa()` - Header da empresa
  - `CriarTabelaProdutos()` - Tabela de produtos
  - `CriarSegundaVia()` - Segunda via do talão

#### 5. **ModuloTemplate.vb** (VBA Templates)
- **Função:** Criação automática de templates profissionais
- **Recursos:**
  - Layout duplo (primeira e segunda via)
  - Formatação profissional automática
  - Bordas e elementos visuais
  - Configuração de página otimizada

#### 6. **ModuloIntegracao.vb** (VBA Bridge)
- **Função:** Ponte de comunicação VB.NET ↔ VBA
- **Responsabilidades:**
  - Receber dados do VB.NET
  - Processar dados para formato VBA
  - Gerenciar status de processamento
  - Controlar planilha temporária

## 🔄 Fluxo de Dados

```
[Interface VB.NET] → [Validação de Dados] → [Excel Automation] → 
[VBA Injection] → [Template Creation] → [Data Population] → 
[Print Configuration] → [Print Execution] → [Cleanup] → [Success Message]
```

## 📊 Estruturas de Dados

### DadosTalao (VB.NET)
```vbnet
Public Class DadosTalao
    Public Property NomeCliente As String
    Public Property EnderecoCliente As String
    Public Property CEP As String
    Public Property Cidade As String
    Public Property Telefone As String
    Public Property Produtos As List(Of ProdutoTalao)
    Public Property FormaPagamento As String
    Public Property Vendedor As String
    Public Property DataVenda As Date
    Public Property NumeroTalao As String
End Class
```

### ProdutoTalao (VB.NET)
```vbnet
Public Class ProdutoTalao
    Public Property Descricao As String
    Public Property Quantidade As Double
    Public Property Unidade As String
    Public Property PrecoUnitario As Double
    Public Property PrecoTotal As Double
End Class
```

## 🛠️ Tecnologias e Dependências

### Framework
- **.NET Framework 4.7.2**
- **Windows Forms**
- **System.Configuration** - Para App.config

### Interoperabilidade
- **Microsoft.Office.Interop.Excel**
- **System.Runtime.InteropServices** - Gestão COM

### Recursos do Sistema
- **Impressora padrão do Windows**
- **Arquivos temporários** - Para planilhas temporárias
- **Registry/GAC** - Para assemblies do Office

## 🔧 Configurações

### App.config - Parâmetros Configuráveis
- `NomeMadeireira` - Nome da empresa
- `EnderecoMadeireira` - Endereço completo
- `TelefoneMadeireira` - Telefone de contato
- `CNPJMadeireira` - CNPJ da empresa
- `VendedorPadrao` - Vendedor padrão
- `ExcelVisivel` - Mostrar Excel durante processamento
- `SalvarTalaoTemporario` - Salvar talão em arquivo

## 🚫 Limitações e Requisitos

### Limitações
- Requer Microsoft Excel instalado
- Funciona apenas em Windows
- Dependente de .NET Framework
- Uma impressão por vez

### Requisitos de Permissão
- Acesso aos objetos VBA do Excel
- Permissão para criar/modificar arquivos temporários
- Acesso à impressora padrão
- Execução de código VBA (macro security)

## 🧪 Testes e Validação

### Dados de Teste Incluídos
- Cliente: "João Silva - TESTE"
- Produtos: Tábua Pinus, Ripão, Compensado
- Valores: Diversos preços de madeireira
- Vendedor: "matheus-testuser3"

### Cenários de Teste
1. **Teste básico:** Dados de teste → Geração → Impressão
2. **Teste de erro:** Excel fechado → Tratamento de erro
3. **Teste de validação:** Dados inválidos → Mensagens de erro
4. **Teste de recursos:** Múltiplas execuções → Cleanup de recursos

## 📈 Performance

### Otimizações Implementadas
- Excel em background (não visível)
- Screen updating desabilitado durante geração
- Gestão adequada de objetos COM
- Cleanup automático de recursos
- Cache de objetos quando possível

### Tempo Estimado de Execução
- Abertura do Excel: 2-5 segundos
- Geração do talão: 1-3 segundos
- Impressão: 2-10 segundos (dependente da impressora)
- Cleanup: 1-2 segundos
- **Total: 6-20 segundos**

---

**Desenvolvido por:** matheus-testuser3  
**Versão:** 1.0.0  
**Data:** 2024