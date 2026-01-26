# 🗂️ Gerenciador de Planilhas

Sistema dinâmico para copiar dados de planilhas de origem para múltiplos destinos no Google Sheets.

## 🚀 Funcionalidades

✅ **Sistema Dinâmico**: Crie e edite configurações direto na interface  
✅ **JSON Configuration**: Todas as configurações ficam salvas em `planilhas_config.json`  
✅ **Granularidade**: Adicione destinos individualmente a qualquer momento
✅ **Resultados em Tabela**: Visualize o status de cada cópia de forma organizada
✅ **Proteção de Colunas**: Preserve colunas específicas no destino durante a cópia

## 🛠️ Como Usar

### 1. Iniciar o Programa

```bash
python main.py
```

### 2. Configurar uma Nova Operação (Origem)

1. Clique no botão azul **"➕ Nova Origem"**.
2. Dê um nome para a configuração (ex: "Relatório de Vendas").
3. Insira o **ID da Planilha de Origem** e o **Nome da Aba** de onde os dados virão.
4. Clique em **"💾 Criar Configuração"**.

### 3. Configurar Destinos

1. Selecione a configuração criada no menu.
2. Na área "Planilhas de Destino", clique em **"➕ Adicionar Planilha Destino"**.
3. Preencha o Nome, ID e Aba da planilha para onde os dados serão copiados.
4. **Opcional**: Indique as **colunas protegidas** (ex: `D, F, CU`) para que elas não sejam sobrescritas.
5. Repita para adicionar quantos destinos precisar.
6. Use o botão **"🗑️ Remover Destino"** se precisar excluir algum da lista.

### 4. Executar a Cópia

1. Verifique se a configuração e os destinos estão corretos.
2. Clique no botão verde **"▶️ INICIAR CÓPIA"**.
3. Acompanhe o progresso na barra verde e na tabela de resultados.

## 📁 Arquivos

- `main.py`: Launcher
- `src/app/interface.py`: Interface Gráfica
- `planilhas_config.json`: Banco de dados das configurações
