# Sistema PDV Completo - Madeireira Maria Luiza

## Visão Geral
Sistema de Ponto de Venda (PDV) completo e moderno desenvolvido para a Madeireira Maria Luiza, integrando interface VB.NET com sistema de talão automatizado via VBA/Excel.

## Funcionalidades Principais

### 1. Interface Principal (Form1.vb)
- **Menu lateral moderno** com ícones e navegação intuitiva
- **Dashboard em tempo real** com cards informativos
- **Botões principais**: PDV/CAIXA, PRODUTOS, CLIENTES, RELATÓRIOS, CONFIGURAÇÃO
- **Interface responsiva** e profissional
- **Integração completa** com módulo VBA

### 2. Sistema de Talão Automatizado (FormPDV.vb)
- **Formulário de entrada** para dados do cliente
- **Campos obrigatórios**: Nome, Endereço, Cidade, CEP, Produtos, Valor Total
- **Validação automática** de dados
- **Integração direta** com módulo de impressão
- **Limpeza automática** do formulário após impressão

### 3. Geração de Talão VBA (ModuloTalaoVBA.vb)
- **Template automático** gerado em tempo real (sem planilha externa)
- **Talão duplo** (cliente + vendedor) lado a lado
- **Formatação profissional** com dados da empresa
- **Impressão automática** configurada para A4 landscape
- **Limpeza automática** de recursos Excel

## Estrutura do Projeto

```
SistemaPDV/
├── Program.vb              # Ponto de entrada da aplicação
├── Form1.vb                # Interface principal com menu e dashboard
├── FormPDV.vb              # Formulário de entrada de dados do PDV
├── ModuloTalaoVBA.vb       # Módulo de integração Excel/VBA
├── DadosCliente.vb         # Classe para dados do cliente
├── SistemaPDV.vbproj       # Arquivo de projeto
└── README.md               # Esta documentação
```

## Fluxo de Funcionamento

1. **Usuário abre o sistema** → Interface principal (Form1.vb)
2. **Clica em "PDV / CAIXA"** → Abre formulário de entrada (FormPDV.vb)
3. **Preenche dados do cliente**:
   - Nome do cliente
   - Endereço completo
   - Cidade (padrão: Paulista)
   - CEP (padrão: 55431-165)
   - Produtos/serviços
   - Valor total
   - Forma de pagamento
   - Vendedor
4. **Clica em "Gerar Talão"** → Sistema valida dados
5. **Processamento automático**:
   - Instancia ModuloTalaoVBA
   - Cria template Excel automaticamente
   - Preenche todos os campos
   - Gera talão duplo (2 vias lado a lado)
   - Configura impressão (A4 landscape)
   - Envia para impressora
6. **Limpa formulário** → Pronto para próxima venda

## Layout do Talão

### Cabeçalho
- **MADEIREIRA MARIA LUIZA**
- **Endereço**: Av. Dr. Olíncio Guerreiro Leite - 631-Paq Amadeu-Paulista-PE-55431-165
- **Telefone**: (81) 98570-1522
- **CNPJ**: 48.905.025/001-61

### Campos do Cliente
- Nome do cliente
- Endereço completo
- Cidade e CEP
- Produtos/serviços detalhados
- Valor total da compra
- Forma de pagamento
- Nome do vendedor
- Data da venda
- Campo para assinatura

### Rodapé
- WhatsApp: (81) 98570-1522
- Instagram: @madeireiramaria

### Formato
- **Talão duplo**: 2 vias idênticas lado a lado
- **Orientação**: Paisagem (A4 landscape)
- **Formatação**: Profissional com bordas e alinhamento
- **Separação**: Linha divisória para facilitar corte

## Dados da Empresa

**MADEIREIRA MARIA LUIZA**
- **Endereço**: Av. Dr. Olíncio Guerreiro Leite - 631-Paq Amadeu-Paulista-PE-55431-165
- **Telefone**: (81) 98570-1522
- **CNPJ**: 48.905.025/001-61
- **WhatsApp**: (81) 98570-1522
- **Instagram**: @madeireiramaria

## Requisitos Técnicos

### Desenvolvimento
- **Visual Studio** (VB.NET)
- **.NET Framework** 4.7.2 ou superior
- **Windows Forms Application**
- **Excel** instalado (para VBA)

### Produção
- **Windows** 10/11
- **Excel** 2016 ou superior
- **Impressora** configurada
- **.NET Framework** instalado

## Instalação e Configuração

1. **Clonar repositório**
```bash
git clone https://github.com/matheus-testuser3/tal-omadeireira.git
```

2. **Abrir projeto** no Visual Studio

3. **Configurar referências**:
   - Microsoft.Office.Interop.Excel
   - System.Windows.Forms
   - System.Drawing

4. **Compilar e executar**

## Funcionalidades Futuras

- [ ] **Sistema de clientes** com cadastro e histórico
- [ ] **Controle de estoque** básico
- [ ] **Relatórios de vendas** por período
- [ ] **Backup automático** de dados
- [ ] **Sistema de usuários** com permissões
- [ ] **Integração com código de barras**
- [ ] **Exportação para PDF**

## Suporte e Contato

Para suporte técnico ou melhorias:
- **Desenvolvedor**: matheus-testuser3
- **GitHub**: https://github.com/matheus-testuser3/tal-omadeireira
- **Empresa**: Madeireira Maria Luiza - (81) 98570-1522

## Licença

Sistema desenvolvido exclusivamente para a Madeireira Maria Luiza.
Todos os direitos reservados.