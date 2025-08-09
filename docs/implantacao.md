# Guia de Implantação - Sistema PDV Madeireira Maria Luiza

## Visão Geral
Este guia detalha como implantar o Sistema PDV completo em um ambiente Windows para produção.

## Estrutura do Projeto Completo

```
tal-omadeireira/
├── README.md                          # Documentação principal
├── SistemaPDV/                        # Diretório do projeto
│   ├── Program.vb                     # Ponto de entrada (demonstração)
│   ├── DadosCliente.vb               # Classe de dados
│   ├── SistemaPDV.vbproj             # Projeto console (demonstração)
│   ├── SistemaPDV_Windows.vbproj     # Projeto Windows Forms (produção)
│   ├── App.config                    # Configurações da aplicação
│   │
│   ├── Form1_Complete.vb             # Interface principal completa
│   ├── FormPDV_Complete.vb           # Formulário PDV completo
│   ├── ModuloTalaoVBA_Complete.vb    # Módulo VBA completo
│   │
│   └── bin/Debug/net8.0/             # Executáveis compilados
│
└── docs/
    ├── manual-usuario.md             # Manual do usuário
    └── implantacao.md               # Este guia
```

## Requisitos do Sistema

### Hardware Mínimo
- **Processador**: Intel Core i3 ou equivalente
- **Memória RAM**: 4 GB
- **Espaço em disco**: 500 MB livres
- **Impressora**: Qualquer impressora configurada no Windows

### Software Necessário
- **Sistema Operacional**: Windows 10/11 (64-bit)
- **Microsoft Excel**: 2016 ou superior
- **.NET Runtime**: 8.0 ou superior
- **Visual Studio**: 2022 (para desenvolvimento)

## Instalação para Produção

### Etapa 1: Preparar Ambiente Windows
1. **Instalar .NET Runtime 8.0**
   ```
   Download: https://dotnet.microsoft.com/download/dotnet/8.0
   Executar: dotnet-runtime-8.0.x-win-x64.exe
   ```

2. **Verificar Excel**
   - Abrir Excel e verificar se está funcionando
   - Habilitar macros se necessário
   - Configurar impressora padrão

### Etapa 2: Compilar Projeto Windows Forms
1. **Abrir Visual Studio 2022**
2. **Criar novo projeto VB.NET Windows Forms**
3. **Copiar arquivos completos:**
   - `Form1_Complete.vb` → `Form1.vb`
   - `FormPDV_Complete.vb` → `FormPDV.vb`
   - `ModuloTalaoVBA_Complete.vb` → `ModuloTalaoVBA.vb`
   - `DadosCliente.vb`
   - `App.config`

4. **Configurar projeto (.vbproj):**
   ```xml
   <Project Sdk="Microsoft.NET.Sdk">
     <PropertyGroup>
       <OutputType>WinExe</OutputType>
       <TargetFramework>net8.0-windows</TargetFramework>
       <UseWindowsForms>true</UseWindowsForms>
     </PropertyGroup>
     <ItemGroup>
       <PackageReference Include="Microsoft.Office.Interop.Excel" Version="15.0.4795.1001" />
     </ItemGroup>
   </Project>
   ```

5. **Compilar em modo Release:**
   ```
   Build → Configuration Manager → Release
   Build → Build Solution
   ```

### Etapa 3: Criar Instalador
1. **Usar ClickOnce Deployment**
   - Project → Properties → Publish
   - Configure URL de publicação
   - Definir pré-requisitos (.NET Runtime)

2. **Ou criar instalador MSI**
   - Adicionar projeto Setup ao solution
   - Configurar dependências
   - Gerar MSI

### Etapa 4: Configurar Aplicação
1. **Editar App.config** para ambiente:
   ```xml
   <add key="EmpresaNome" value="MADEIREIRA MARIA LUIZA" />
   <add key="DiretorioBackup" value="C:\SistemaPDV\Backup" />
   <add key="DiretorioHistorico" value="C:\SistemaPDV\Historico" />
   ```

2. **Criar diretórios necessários:**
   ```
   C:\SistemaPDV\
   C:\SistemaPDV\Backup\
   C:\SistemaPDV\Historico\
   ```

3. **Configurar permissões:**
   - Permissão de escrita em diretórios do sistema
   - Permissão para executar Excel
   - Permissão para impressão

## Teste de Funcionamento

### Teste 1: Interface Principal
1. Executar aplicativo
2. Verificar menu lateral carregado
3. Verificar dashboard com cards
4. Testar navegação entre telas

### Teste 2: Formulário PDV
1. Clicar em "PDV / CAIXA"
2. Preencher dados de teste:
   - Cliente: João Silva
   - Endereço: Rua Teste, 123
   - Cidade: Paulista
   - CEP: 55431-165
   - Produtos: Tábua de madeira
   - Valor: 25,00
3. Clicar "Gerar Talão"
4. Verificar abertura do Excel
5. Verificar formatação do talão
6. Testar impressão

### Teste 3: Integração VBA
1. Verificar criação automática de template
2. Verificar preenchimento de dados
3. Verificar talão duplo (2 vias)
4. Verificar configuração de impressão
5. Verificar limpeza de recursos

## Solução de Problemas

### Problema: "Excel não encontrado"
**Causa**: Microsoft Excel não instalado
**Solução**: 
- Instalar Microsoft Office ou Excel
- Verificar versão compatível (2016+)
- Registrar Office Interop

### Problema: "Erro de permissão"
**Causa**: Aplicativo sem permissões necessárias
**Solução**:
- Executar como administrador
- Configurar permissões de pasta
- Verificar antivírus

### Problema: "Impressora não funciona"
**Causa**: Impressora não configurada
**Solução**:
- Configurar impressora padrão
- Testar impressão em outro programa
- Verificar drivers de impressora

### Problema: "Template não cria"
**Causa**: Erro na integração VBA
**Solução**:
- Verificar Excel funcionando
- Habilitar macros se necessário
- Verificar .NET runtime

## Backup e Manutenção

### Backup Automático
- Sistema gera backup das vendas em `C:\SistemaPDV\Backup\`
- Backup diário automático (se configurado)
- Histórico de vendas em `C:\SistemaPDV\Historico\`

### Manutenção Preventiva
- **Semanal**: Verificar espaço em disco
- **Mensal**: Limpar arquivos temporários do Excel
- **Trimestral**: Backup completo do sistema
- **Anual**: Atualização de versão se disponível

## Atualização do Sistema

### Processo de Atualização
1. **Fazer backup** completo dos dados
2. **Fechar aplicativo** completamente
3. **Substituir executável** pela nova versão
4. **Verificar configurações** em App.config
5. **Testar funcionamento** completo
6. **Treinar usuários** se necessário

### Versionamento
- **v1.0**: Versão inicial completa
- **v1.1**: Melhorias de interface
- **v1.2**: Funcionalidades extras
- **v2.0**: Nova arquitetura

## Suporte Técnico

### Contatos
- **Desenvolvedor**: matheus-testuser3
- **Empresa**: Madeireira Maria Luiza
- **Telefone**: (81) 98570-1522
- **WhatsApp**: (81) 98570-1522
- **Email**: suporte@madeireiramaria.com.br

### Níveis de Suporte
1. **Nível 1**: Usuário final (manual do usuário)
2. **Nível 2**: TI local (este guia)
3. **Nível 3**: Desenvolvedor (suporte técnico)

### Logs do Sistema
- Logs de erro em: `C:\SistemaPDV\Logs\`
- Logs de transações: `C:\SistemaPDV\Historico\`
- Logs do Excel: Eventos do Windows

---

**Documento**: Guia de Implantação v1.0  
**Data**: Dezembro 2024  
**Autor**: matheus-testuser3  
**Sistema**: PDV Madeireira Maria Luiza