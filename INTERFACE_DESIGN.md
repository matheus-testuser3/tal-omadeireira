# Interface do Sistema PDV - Visualização

## 🖥️ Tela Principal

```
╔════════════════════════════════════════════════════════════════════════════════════════╗
║                           Sistema PDV - Madeireira Maria Luiza                        ║
╠═══════════════════╤════════════════════════════════════════════════════════════════════╣
║                   │                                                                    ║
║     [LOGO]        │                    MADEIREIRA MARIA LUIZA                        ║
║                   │           Sistema de Ponto de Venda Integrado com                ║
║                   │              Geração Automática de Talões                        ║
║                   │                                                                    ║
║ 🧾 GERAR TALÃO   │   📋 INSTRUÇÕES DE USO:                                          ║
║                   │                                                                    ║
║ ⚙️ Configurações  │   1. Clique em 'GERAR TALÃO' para abrir formulário               ║
║                   │   2. Preencha dados do cliente e produtos                         ║
║ ℹ️ Sobre o        │   3. Sistema abrirá Excel automaticamente                        ║
║    Sistema        │   4. Talão será gerado e impresso automaticamente                ║
║                   │   5. Excel será fechado automaticamente                           ║
║                   │                                                                    ║
║                   │   ✅ Não precisa ter Excel aberto manualmente                     ║
║                   │   ✅ Não precisa ter planilhas salvas                             ║
║                   │   ✅ Todo o processo é automático!                                ║
║                   │                                                                    ║
║                   │   Rua Principal, 123 - Centro                                     ║
║                   │   📞 (81) 3436-1234                                               ║
║                   │   📋 CNPJ: 12.345.678/0001-90                                     ║
║                   │                                                                    ║
║ 🚪 Sair          │                                                                    ║
╚═══════════════════╧════════════════════════════════════════════════════════════════════╝
```

## 📝 Formulário de Entrada de Dados

```
╔════════════════════════════════════════════════════════════════════════════════════════╗
║                          Entrada de Dados - Talão de Venda                            ║
╠════════════════════════════════════════════════════════════════════════════════════════╣
║ 👤 DADOS DO CLIENTE                                           📝 Carregar Dados Teste ║
║                                                                                        ║
║ Nome do Cliente: [João Silva - TESTE                                    ]             ║
║ Endereço:       [Rua das Árvores, 123 - Centro                                    ]   ║
║ CEP:            [55431-165    ] Cidade: [Paulista/PE              ]                   ║
║ Telefone:       [(81) 9876-5432     ]                                                 ║
╠════════════════════════════════════════════════════════════════════════════════════════╣
║ 📦 PRODUTOS                                                                            ║
║                                                                                        ║
║ Descrição: [Tábua de Pinus 2x4m            ] Qtd:[5  ] Un:[UN] Preço:[25,00] [+][-]  ║
║                                                                                        ║
║ ┌────────────────────────────────────────────────────────────────────────────────────┐ ║
║ │ DESCRIÇÃO              │ QTD │ UN │ PREÇO UNIT. │ TOTAL      │                    │ ║
║ ├────────────────────────────────────────────────────────────────────────────────────┤ ║
║ │ Tábua de Pinus 2x4m    │  5  │ UN │ R$ 25,00    │ R$ 125,00  │                    │ ║
║ │ Ripão 3x3x3m           │ 10  │ UN │ R$ 15,00    │ R$ 150,00  │                    │ ║
║ │ Compensado 18mm        │  2  │ M² │ R$ 45,00    │ R$ 90,00   │                    │ ║
║ └────────────────────────────────────────────────────────────────────────────────────┘ ║
║                                                                                        ║
║ Forma de Pagamento: [Dinheiro     ▼] Vendedor: [matheus-testuser3        ]           ║
║                                                                                        ║
║                                              ✅ CONFIRMAR E GERAR TALÃO ❌ Cancelar   ║
╚════════════════════════════════════════════════════════════════════════════════════════╝
```

## 📄 Talão Gerado (Preview)

```
╔════════════════════════════════════════════════════════════════════════════════════════╗
║                           MADEIREIRA MARIA LUIZA                                      ║
║                         Rua Principal, 123 - Centro                                   ║
║                    Paulista/PE - CEP: 53401-445 - Tel: (81) 3436-1234                ║
║                            CNPJ: 12.345.678/0001-90                                   ║
║════════════════════════════════════════════════════════════════════════════════════════║
║                                                                                        ║
║ TALÃO DE VENDA Nº: 20241209155432                                                     ║
║ Data: 09/12/2024 15:54                                                                ║
║                                                                                        ║
║ CLIENTE:     João Silva - TESTE                                                       ║
║ ENDEREÇO:    Rua das Árvores, 123 - Centro                                           ║
║ CIDADE/CEP:  Paulista/PE - CEP: 55431-165                                            ║
║ TELEFONE:    (81) 9876-5432                                                           ║
║                                                                                        ║
║ ┌────────────────────────────────────────────────────────────────────────────────────┐ ║
║ │ DESCRIÇÃO              │ QTD │ UN │ PREÇO UNIT. │ TOTAL      │                    │ ║
║ ├────────────────────────────────────────────────────────────────────────────────────┤ ║
║ │ Tábua de Pinus 2x4m    │  5  │ UN │ R$ 25,00    │ R$ 125,00  │                    │ ║
║ │ Ripão 3x3x3m           │ 10  │ UN │ R$ 15,00    │ R$ 150,00  │                    │ ║
║ │ Compensado 18mm        │  2  │ M² │ R$ 45,00    │ R$ 90,00   │                    │ ║
║ ├────────────────────────────────────────────────────────────────────────────────────┤ ║
║ │                                        TOTAL GERAL: │ R$ 365,00  │                │ ║
║ └────────────────────────────────────────────────────────────────────────────────────┘ ║
║                                                                                        ║
║ FORMA DE PAGAMENTO: Dinheiro                                                          ║
║ VENDEDOR: matheus-testuser3                                                           ║
║                                                                                        ║
║ CLIENTE: _________________________________                                            ║
║            (NOME E ASSINATURA)                                                        ║
║                                                                                        ║
║ ✂️ --- CORTE AQUI - SEGUNDA VIA --- ✂️                                                ║
║                                                                                        ║
║                         MADEIREIRA MARIA LUIZA                                        ║
║ TALÃO Nº: 20241209155432 - 09/12/2024 15:54                                          ║
║ CLIENTE: João Silva - TESTE                                                           ║
║ TOTAL: R$ 365,00                                                                      ║
║ PAGAMENTO: Dinheiro                                                                    ║
║ VENDEDOR: matheus-testuser3                                                           ║
║                                                                                        ║
║ CLIENTE: ________________________                                                     ║
║                                                                                        ║
║ ──────────────────────────────────────────────────────────────────────────────────── ║
║               Sistema PDV - Madeireira Maria Luiza © 2024                            ║
╚════════════════════════════════════════════════════════════════════════════════════════╝
```

## 💫 Características Visuais

### Cores do Sistema
- **Verde Principal:** #2ECC71 (botão principal)
- **Azul Escuro:** #34495E (sidebar)
- **Cinza Claro:** #ECF0F1 (background)
- **Vermelho:** #E74C3C (botão sair/cancelar)
- **Amarelo:** #F1C40F (botão teste)

### Tipografia
- **Fonte Principal:** Segoe UI
- **Título Empresa:** 18pt, Bold
- **Títulos Seção:** 12pt, Bold
- **Texto Normal:** 10pt, Regular
- **Botões:** 10-12pt, Bold

### Layout
- **Sidebar:** 250px de largura
- **Formulário:** 900x700px, modal
- **Interface Principal:** 1200x800px
- **Talão:** A4, formatação profissional

---

**Esta visualização representa a interface moderna e profissional do Sistema PDV para Madeireira Maria Luiza, com foco na simplicidade e automação completa do processo de geração de talões.**