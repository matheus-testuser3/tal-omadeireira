# Sistema PDV - Guia de Instalação e Uso

## 🚀 Instalação Rápida

### Pré-requisitos
1. **Windows 7 ou superior**
2. **.NET Framework 4.7.2 ou superior**
3. **Microsoft Excel 2010 ou superior**

### Compilação
```bash
# Abrir no Visual Studio 2017 ou superior
# Compilar em modo Release
# Executável será gerado em: bin/Release/SistemaPDV.exe
```

## 📋 Como Usar

### 1. Primeira Execução
- Execute `SistemaPDV.exe`
- O sistema verificará se o Excel está instalado
- Interface principal será exibida

### 2. Geração de Talão
1. **Clique em "🧾 GERAR TALÃO"**
2. **Preencha os dados:**
   - Nome do cliente
   - Endereço completo
   - CEP e cidade
   - Telefone
3. **Adicione produtos:**
   - Descrição do produto
   - Quantidade
   - Unidade (UN, M, M², etc.)
   - Preço unitário
4. **Configure pagamento e vendedor**
5. **Clique em "✅ CONFIRMAR E GERAR TALÃO"**

### 3. Processo Automático
- ✅ Excel abre automaticamente em background
- ✅ Template profissional é criado
- ✅ Dados são preenchidos
- ✅ Talão é formatado e impresso
- ✅ Excel fecha automaticamente
- ✅ Mensagem de sucesso é exibida

## 🧪 Teste Rápido

Para testar o sistema rapidamente:
1. Clique em "📝 Carregar Dados de Teste"
2. Dados do cliente e produtos são preenchidos automaticamente
3. Clique em "✅ CONFIRMAR E GERAR TALÃO"
4. O sistema gerará um talão de teste

## ⚙️ Configurações

Edite o arquivo `App.config` para personalizar:

```xml
<add key="NomeMadeireira" value="SUA MADEIREIRA AQUI" />
<add key="EnderecoMadeireira" value="SEU ENDEREÇO" />
<add key="CidadeMadeireira" value="SUA CIDADE/UF" />
<add key="TelefoneMadeireira" value="SEU TELEFONE" />
<add key="CNPJMadeireira" value="SEU CNPJ" />
<add key="VendedorPadrao" value="NOME DO VENDEDOR" />
```

## 🐛 Solução de Problemas

### Excel não encontrado
- Instale Microsoft Excel
- Execute como Administrador
- Verifique se o Excel não está em uso

### Erro de permissão VBA
- Configure Excel para permitir macros
- Adicione o programa à lista de confiança
- Execute como Administrador

### Erro de impressão
- Configure uma impressora padrão
- Teste impressão manual no Excel
- Verifique drivers da impressora

## 📞 Suporte

**Desenvolvedor:** matheus-testuser3  
**GitHub:** https://github.com/matheus-testuser3/tal-omadeireira

---

**© 2024 - Sistema PDV Madeireira Maria Luiza**