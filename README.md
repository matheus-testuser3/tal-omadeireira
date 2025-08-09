# Sistema PDV Integrado - Madeireira Maria Luiza

![Status](https://img.shields.io/badge/Status-Desenvolvimento-yellow)
![Versão](https://img.shields.io/badge/Versão-1.0.0-blue)
![.NET Framework](https://img.shields.io/badge/.NET_Framework-4.7.2-purple)
![Excel](https://img.shields.io/badge/Excel-Required-green)

## 📋 Descrição do Projeto

Sistema completo de Ponto de Venda (PDV) desenvolvido em VB.NET que integra automaticamente com Microsoft Excel e VBA para geração profissional de talões de venda para madeireiras.

### 🎯 Objetivo Principal

Criar um sistema PDV completo que:
- ✅ **Abre Excel automaticamente** quando necessário
- ✅ **Cria planilha temporária** para geração do talão
- ✅ **Transfere módulos VBA** integrados no projeto VB.NET
- ✅ **Executa macros VBA** diretamente do VB.NET
- ✅ **Gera e imprime** talão automaticamente
- ✅ **Fecha Excel** após a operação

## 🏗️ Arquitetura da Solução

### **Interface Principal (VB.NET)**
- Interface moderna com menu lateral
- Formulário de entrada de dados (cliente, produtos, etc.)
- Botão "Gerar Talão" que chama o sistema VBA
- Integração Microsoft.Office.Interop.Excel
- Gerenciamento automático do Excel (abrir/fechar)

### **Módulos VBA Integrados**
Todos os módulos VBA são incorporados como código no projeto VB.NET:

#### **ModuloTalao.vb** - Sistema principal de talão
- `ProcessarTalaoCompleto()` - Função principal
- `GerarTalaoCompleto()` - Geração do layout
- `EscreverTalao()` - Preenchimento dos dados
- `ConfigurarImpressaoCompleta()` - Configuração de impressão
- `CriarSegundaVia()` - Geração da segunda via

#### **ModuloTemplate.vb** - Criação automática de template
- `CriarTemplateAutomatico()` - Template base
- `DefinirFormatacaoProfissional()` - Formatação visual
- `ConfigurarLayoutDuplo()` - Layout com primeira e segunda via
- `AdicionarElementosVisuais()` - Bordas e elementos gráficos

#### **ModuloIntegracao.vb** - Ponte VB.NET ↔ VBA
- `ReceberDadosDoVBNET()` - Interface de comunicação
- `ProcessarDadosColetados()` - Processamento de dados
- `RetornarStatusProcessamento()` - Status de execução
- `GerenciarPlanilhaTemporaria()` - Gestão da planilha temporária

### **ExcelAutomation.vb** - Controle do Excel
- `ProcessarTalaoCompleto()` - Função principal de automação
- `AbrirExcel()` - Abertura do Excel em background
- `InjetarModulosVBA()` - Injeção dinâmica dos módulos VBA
- `CriarTemplate()` - Criação do template
- `PreencherDados()` - Preenchimento com dados do cliente
- `ImprimirTalao()` - Impressão automática
- `FecharExcel()` - Fechamento e limpeza

## 🚀 Fluxo Automatizado

### **Passo 1: Usuário usa interface VB.NET**
```vbnet
' Usuario preenche dados no FormPDV.vb
- Nome do cliente
- Endereço, CEP, cidade
- Produtos e quantidades
- Forma de pagamento
```

### **Passo 2: VB.NET abre Excel automaticamente**
```vbnet
Dim xlApp As Excel.Application = New Excel.Application()
Dim xlWorkbook As Excel.Workbook = xlApp.Workbooks.Add()
xlApp.Visible = False ' Executar em background
```

### **Passo 3: VB.NET injeta módulos VBA**
```vbnet
' Adicionar todos os módulos VBA na planilha temporária
Dim vbaModule As Object = xlWorkbook.VBProject.VBComponents.Add(1)
vbaModule.CodeModule.AddFromString(codigoVBACompleto)
```

### **Passo 4: Execução automática**
```vbnet
' Chamar função principal do VBA passando dados
xlApp.Run("ProcessarTalaoCompleto", dadosColetados)
```

### **Passo 5: Impressão e finalização**
```vbnet
' VBA gera talão, imprime e retorna status
' VB.NET fecha Excel automaticamente
xlWorkbook.Close(False)
xlApp.Quit()
```

## 📁 Estrutura dos Arquivos

```
tal-omadeireira/
├── SistemaPDV.vb          # Interface principal VB.NET
├── FormPDV.vb             # Formulário entrada de dados  
├── ExcelAutomation.vb     # Automação do Excel
├── ModuloTalao.vb         # Sistema VBA de talão
├── ModuloTemplate.vb      # Template automático VBA
├── ModuloIntegracao.vb    # Ponte VB.NET ↔ VBA
├── SistemaPDV.vbproj      # Projeto VB.NET
├── App.config             # Configurações
├── README.md              # Esta documentação
└── Resources/
    └── LogoMadeireira.png # Logo para talão (futuro)
```

## 🛠️ Tecnologias Utilizadas

- **VB.NET** - Interface e controle principal
- **Microsoft.Office.Interop.Excel** - Integração com Excel
- **VBA** - Geração de talão (código integrado)
- **Windows Forms** - Interface gráfica
- **.NET Framework 4.7.2** - Base do sistema

## ⚙️ Requisitos do Sistema

### **Software Necessário**
- Windows 7 ou superior
- .NET Framework 4.7.2 ou superior
- Microsoft Excel 2010 ou superior
- Visual Studio 2017 ou superior (para desenvolvimento)

### **Permissões Necessárias**
- Acesso aos objetos VBA do Excel
- Permissão para criar arquivos temporários
- Acesso à impressora padrão

## 🚀 Como Usar

### **1. Executar o Sistema**
```
SistemaPDV.exe
```

### **2. Interface Principal**
- Abrir o programa VB.NET
- Ver interface moderna com menu lateral
- Clicar em "🧾 GERAR TALÃO"

### **3. Entrada de Dados**
- Preencher dados do cliente
- Adicionar produtos e quantidades
- Definir forma de pagamento
- Clicar em "✅ CONFIRMAR E GERAR TALÃO"

### **4. Processamento Automático**
- Sistema abre Excel automaticamente (invisível)
- Cria template na hora
- Gera talão duplo profissional
- Imprime automaticamente
- Fecha Excel
- Mostra mensagem de sucesso

## 🧪 Dados de Teste

Para facilitar a demonstração, o sistema inclui um botão "📝 Carregar Dados de Teste" que preenche automaticamente:

**Cliente:**
- Nome: João Silva - TESTE
- Endereço: Rua das Árvores, 123 - Centro
- CEP: 55431-165
- Cidade: Paulista/PE
- Telefone: (81) 9876-5432

**Produtos:**
- Tábua de Pinus 2x4m - 5 UN - R$ 25,00 = R$ 125,00
- Ripão 3x3x3m - 10 UN - R$ 15,00 = R$ 150,00
- Compensado 18mm - 2 M² - R$ 45,00 = R$ 90,00

**Vendedor:** matheus-testuser3

## ⚡ Vantagens da Solução

### ✅ **Para o Usuário Final**
- **Não precisa abrir Excel manualmente**
- **Não precisa ter planilhas salvas**
- **Não precisa conhecer VBA**
- **Interface moderna e simples**
- **Um clique para gerar talão**

### ✅ **Para o Desenvolvedor**
- **Código VBA preservado** e incorporado
- **Controle total** via VB.NET
- **Fácil manutenção** - módulos separados
- **Reutilização** - mesmo VBA em qualquer projeto
- **Backup automático** - código dentro do .exe

### ✅ **Para o Sistema**
- **Execução automática** - Excel abre/fecha sozinho
- **Template dinâmico** - criado por código
- **Memória otimizada** - Excel só aberto quando necessário
- **Erro mínimo** - processo controlado
- **Portabilidade** - funciona em qualquer máquina com Excel

## 🔧 Configurações

O arquivo `App.config` contém configurações personalizáveis:

```xml
<appSettings>
    <add key="NomeMadeireira" value="Madeireira Maria Luiza" />
    <add key="EnderecoMadeireira" value="Rua Principal, 123 - Centro" />
    <add key="CidadeMadeireira" value="Paulista/PE" />
    <add key="CEPMadeireira" value="53401-445" />
    <add key="TelefoneMadeireira" value="(81) 3436-1234" />
    <add key="CNPJMadeireira" value="12.345.678/0001-90" />
    <add key="VendedorPadrao" value="matheus-testuser3" />
    <add key="ExcelVisivel" value="false" />
    <add key="SalvarTalaoTemporario" value="false" />
</appSettings>
```

## 🐛 Solução de Problemas

### **Excel não encontrado**
- Verificar se Microsoft Excel está instalado
- Executar o sistema como Administrador
- Verificar se o Excel está atualizado

### **Erro de permissão VBA**
- Habilitar macros no Excel
- Verificar configurações de segurança
- Adicionar o sistema à lista de confiança

### **Erro de impressão**
- Verificar se há impressora configurada
- Testar impressão manual no Excel
- Verificar drivers da impressora

## 👨‍💻 Desenvolvimento

### **Como Compilar**
```bash
# Abrir o projeto no Visual Studio
# Compilar em modo Release
# O executável será gerado em bin/Release/
```

### **Como Contribuir**
1. Fork do repositório
2. Criar branch para feature
3. Implementar mudanças
4. Testar funcionamento
5. Submeter Pull Request

## 📞 Suporte

**Desenvolvedor:** matheus-testuser3  
**Email:** [Inserir email de contato]  
**GitHub:** https://github.com/matheus-testuser3/tal-omadeireira

## 📄 Licença

© 2024 - Todos os direitos reservados.

---

**🎯 RESULTADO FINAL:** O usuário simplesmente abre o programa, preenche os dados, clica em "Gerar Talão" e **TUDO É AUTOMÁTICO!** Não precisa mexer no Excel, não precisa ter planilhas, só usar a interface moderna!