# Relatório Mensal — Projeto

Resumo rápido
---------------
Este projeto fornece uma pequena aplicação desktop (Tkinter + ttkbootstrap) para ler uma planilha Excel, aplicar filtros (mês, ano e tipo de recado) e gerar um relatório em texto (TXT). O objetivo é permitir gerar relatórios mensais de alunos/contratos a partir de um Excel.

Principais features implementadas
---------------------------------
- Interface gráfica (GUI) com abas
  - Aba "Gerar Relatório": seleção de arquivo Excel, filtros (mês, ano, tipo), escolha de destino do TXT e botão para gerar.
  - Aba "Relatório Gerado": visualizador de texto com opção de abrir no Bloco de Notas.

- Validação básica de entrada
  - Verifica se um arquivo Excel foi selecionado e cria um caminho padrão para saída quando necessário.

- Pipeline de processamento (separado em controller/service)
  - Controller (`controllers/relatorio_controller.py`): valida entradas e orquestra chamadas ao service.
  - Service (`services/relatorio_service.py`): lê a planilha (via `utils/excel_utils.py`), aplica filtros e monta o texto do relatório.

- Utilitários
  - `utils/excel_utils.py`: leitura e normalização da planilha, conversão de datas, criação de colunas auxiliares (`MES_NUM`, `ANO`, `MES_NOME`) e função para aplicar filtros.
  - `utils/file_utils.py`: funções para salvar o relatório em TXT (e função para PDF, disponível mas não utilizada por padrão).

- Documentação inline
  - Docstrings e comentários em arquivos principais para facilitar manutenção futura (explicação de entradas/saídas e erros possíveis).

- Gerador de planilha fictícia (script)
  - `generate_planilha_ficticia.py`: utilitário para gerar planilhas Excel fictícias (nomes/matrículas/datas/recados). Útil para testes e para subir exemplos sem dados reais.

Como rodar
----------
1. Crie um ambiente virtual e instale dependências:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

2. (Opcional) Gere uma planilha fictícia para testes:

```powershell
# padrão gera 1200 registros
python generate_planilha_ficticia.py
```

3. Execute a aplicação GUI:

```powershell
python main.py
```

Arquivos importantes
-------------------
- `main.py` — entrada da aplicação e montagem da GUI
- `controllers/relatorio_controller.py` — valida inputs e chama o service
- `services/relatorio_service.py` — lógica de geração do relatório
- `utils/excel_utils.py` — leitura/normalização de Excel e filtros
- `utils/file_utils.py` — salvar TXT/PDF
- `generate_planilha_ficticia.py` — script para gerar dados fictícios
- `data/` — pasta para planilhas de exemplo (ex.: `PLANILHA_EXEMPLO.xlsx`)
- `requirements.txt` — dependências do projeto

Observações e próximos passos sugeridos
--------------------------------------
- O fluxo atual gera somente o arquivo TXT por padrão. A geração de PDF foi removida do fluxo principal para evitar problemas com formatação e dependências.
- Melhorias recomendadas:
  - Adicionar testes automatizados (pytest) cobrindo leitura, filtros e geração de arquivo.
  - Implementar logging para facilitar depuração em ambientes headless.
  - Consertar/especificar versões das dependências em `requirements.txt` antes de subir para integração contínua.
  - Opcional: adicionar um `DOCUMENTATION.md` com exemplos de planilha (colunas esperadas) e um template para os usuários preencherem.

Licença
-------
Adicione um arquivo `LICENSE` conforme sua preferência (por exemplo MIT) antes de publicar o repositório.

Contato
-------
Se quiser, eu posso:
- adicionar `DOCUMENTATION.md` com um template de planilha;
- criar testes `pytest` básicos e um `Makefile`/`tasks.json` para facilitar execução;
- fixar versões no `requirements.txt`.
