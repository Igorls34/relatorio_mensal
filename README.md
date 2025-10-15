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
  - `generate_planilha_ficticia.py`: utilitário opcional para gerar planilhas Excel fictícias (nomes/matrículas/datas/recados). Use apenas se precisar de dados de exemplo — não é necessário para rodar o projeto.

# Relatório Mensal — Contexto e motivação

Visão geral
-----------
Este projeto nasceu de uma necessidade real de trabalho: lidar com um grande volume de alunos e prazos em um sistema que, por muito tempo, continha dados inconsistentes e mal preenchidos. A falta de revisão e limpeza do banco de dados tornava o acompanhamento dos prazos impreciso e trabalhoso.

Motivação (caso real)
---------------------
No meu trabalho eu precisava identificar com clareza quais alunos estavam com pagamentos ou procedimentos em atraso, quem precisava ser notificado e quais os prazos de vencimento. O sistema que usamos fornecia apenas uma planilha bruta como saída — com muitas informações redundantes ou incorretas — o que atrapalhava a análise.

Para resolver isso eu:

1. Fiz uma limpeza no banco de dados do meu trabalho, corrigindo informações incorretas e removendo registros inválidos para ter uma base confiável.
2. Trabalhei com o suporte do sistema para obter uma exportação mais completa, que devolvia uma planilha com as informações necessárias.
3. Percebi a necessidade de um relatório mais organizado e automatizado, gerado mensalmente, que me permitisse filtrar por mês, categoria e identificar rapidamente quem estava em atraso.

Solução implementada
---------------------
Criei este pequeno sistema para processar a planilha (exportada pelo sistema) e gerar um relatório mensal organizado. Hoje a versão atual do projeto:

- Lê uma planilha Excel com os dados brutos.
- Normaliza e valida as colunas principais (ex.: matrícula, nome, fim do contrato, recados).
- Aplica filtros por mês, ano e tipo de recado.
- Gera um arquivo TXT com o relatório organizado por mês, listando os alunos e suas informações relevantes.

Por que um TXT?
---------------
No primeiro passo preferi gerar um TXT simples e legível — fácil de abrir e revisar rapidamente. Futuramente a ideia é evoluir para saídas mais ricas (PDF/Excel/integração com outros sistemas) e relatórios mais detalhados.

Proteção de dados e planilha de exemplo
--------------------------------------
Para evitar qualquer risco de vazamento de dados sensíveis dos alunos, a planilha de exemplo presente no repositório (quando houver) contém apenas dados totalmente fictícios — nomes, matrículas e datas geradas aleatoriamente. Eu gerei essa planilha apenas como demonstração do formato e nunca inclui dados reais dos alunos neste projeto.

Como usar (resumido)
--------------------
1. Instale dependências:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

2. Execute a GUI:

```powershell
python main.py
```

3. Na aba "Gerar Relatório": selecione a planilha (ou use a planilha de exemplo fictícia), escolha filtros e gere o TXT. Na aba "Relatório Gerado" você pode visualizar o resultado.

Arquivos e responsabilidades
----------------------------
- `main.py` — interface gráfica e interação com o usuário.
- `controllers/relatorio_controller.py` — valida entradas e orquestra chamadas ao serviço.
- `services/relatorio_service.py` — lógica de leitura, filtragem, agrupamento e montagem do relatório.
- `utils/excel_utils.py` — funções de leitura/normalização de Excel e aplicação de filtros.
- `utils/file_utils.py` — funções para salvar o relatório (TXT, PDF).

Próximos passos planejados
-------------------------
- Tornar a saída mais rica (PDF formatado, Excel organizado) e adicionar opções de exportação.
- Adicionar logging e testes automáticos (pytest) para aumentar robustez.
- Implementar filtros mais avançados e dashboards para facilitar a gestão diária.

Por que compartilhei este projeto?
---------------------------------
Apesar de ser um sistema simples, separei este projeto para demonstrar como eu resolvi um problema real no trabalho: limpeza de dados, automação de relatórios e geração de saída operável para controle de prazos. O resultado foi útil para meu trabalho e acredito que a solução pode ser expandida para uso em outras instituições.

Contato e contribuições
-----------------------
Se quiser colaborar, sugerir melhorias ou adaptar o projeto para outro contexto, abra uma issue ou um pull request no repositório.

## 📄 Licença

Este projeto está licenciado sob os termos da [Licença MIT](./LICENSE).

Você é livre para usar, modificar e distribuir este código para fins pessoais ou comerciais, desde que mantenha o crédito ao autor original.

