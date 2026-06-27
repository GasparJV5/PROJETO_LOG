MVP Conferência NF-e x OC
1. Visão geral
Este projeto nasceu de uma necessidade operacional real: automatizar a conferência entre uma NF-e em XML e uma Ordem de Compra (OC), reduzindo trabalho manual, retrabalho, risco de erro humano e falta de rastreabilidade no processo de recebimento/agendamento.
A ideia inicial foi construir um fluxo em que o sistema recebesse arquivos de entrada, extraísse as informações principais, classificasse os dados em cabeçalho e itens, realizasse uma primeira conferência documental e gerasse saídas estruturadas para futura visualização em cards, SharePoint ou dashboard.
A primeira conquista técnica foi validar um MVP funcional capaz de:
ler XML de NF-e;
ler OC em Word/DOC;
extrair cabeçalho da nota fiscal;
extrair itens da nota fiscal;
extrair cabeçalho da OC;
extrair itens da OC;
comparar o cabeçalho;
gerar arquivos CSV e Excel de resultado;
registrar logs de processamento.
---
2. Problema de negócio
No processo atual, a conferência entre nota fiscal e ordem de compra pode exigir validações manuais como:
fornecedor correto;
CNPJ correto;
cliente/destinatário correto;
endereço/local de entrega;
valor total do pedido versus valor total da nota;
prazo de entrega;
vínculo documental entre OC e NF-e;
itens solicitados versus itens faturados.
Esse processo manual pode gerar desperdícios operacionais, como:
tempo gasto em conferências repetitivas;
retrabalho por divergência não identificada cedo;
risco de aceitar nota fiscal sem vínculo correto com a OC;
dificuldade de rastrear erros;
baixa visibilidade para colaboradores e gestão.
---
3. Objetivo do MVP
O objetivo do MVP é criar uma base automatizada para conferência documental entre NF-e XML e OC Word/DOC, com foco inicial em:
extração de dados;
separação entre cabeçalho e itens;
conferência de cabeçalho;
geração de saídas estruturadas;
preparação para futura comparação item a item;
preparação para cards em SharePoint ou outra interface.
A comparação item a item, fator de conversão, unidade, embalagem e regras avançadas ficam para a próxima fase.
---
4. Regra de negócio principal
A regra mais importante definida durante o desenvolvimento foi:
> **O Doc. Origem da OC deve constar na observação/dados adicionais da NF-e.**
Ou seja, o vínculo documental entre a OC e a NF-e não deve ser validado prioritariamente pelo número da OC, mas sim pelo campo Doc. Origem.
Quando o Doc. Origem da OC não é encontrado na observação da NF-e, o sistema classifica o cabeçalho como:
```text
REJEITADA_DOC_ORIGEM_NAO_REFERENCIADO
```
Essa regra reforça rastreabilidade, governança documental e controle operacional.
---
5. Arquitetura atual
A arquitetura atual do MVP trabalha com:
```text
XML da NF-e + DOC/Word da OC
```
Estrutura esperada
```text
mvp_conferencia/
│
├── entrada/
│   ├── xml/
│   │   └── .gitkeep
│   └── oc/
│       └── .gitkeep
│
├── resultado/
│   └── .gitkeep
│
├── saida/
│   └── .gitkeep
│
├── main.py
├── requirements.txt
├── .gitignore
└── README.md
```
Entrada
```text
entrada/xml/*.xml
entrada/oc/*.doc
```
Saída
```text
resultado/xml_cabecalho.csv
resultado/xml_itens.csv
resultado/oc_doc_cabecalho.csv
resultado/oc_doc_itens.csv
resultado/conferencia_cabecalho.csv
resultado/log_processamento.csv
resultado/resultado.xlsx
```
---
6. O que foi implementado
Bloco	Descrição	Status
Leitura XML	Leitura do arquivo XML da NF-e	Concluído
Cabeçalho XML	Extração de emitente, destinatário, CNPJ, endereço, datas, valores e observação	Concluído
Itens XML	Extração de código, descrição, unidade, quantidade, valor unitário, valor total, lote, validade e IPI	Concluído
Leitura DOC/Word	Leitura da OC em formato Word/DOC, com conversão via Microsoft Word/pywin32 quando necessário	Concluído
Cabeçalho OC	Extração de número da OC, pedido, Doc. Origem, fornecedor, CNPJ, cliente, prazo e valor total	Concluído
Itens OC	Extração de código, descrição, unidade, embalagem, quantidade, valor unitário, desconto, IPI e valor total	Concluído
Conferência cabeçalho	Comparação inicial entre XML e OC	Concluído
Log	Registro do processamento	Concluído
Excel/CSV	Geração de saídas estruturadas	Concluído
Comparação item a item	Comparação detalhada dos itens	Próxima fase
Fator/unidade/embalagem	Tratamento de unidade, embalagem e quantidade base	Próxima fase
Cards/SharePoint	Saída final para cards	Próxima fase
Front-end	Interface de uso	Futuro
Machine Learning	Apoio a identificação de padrões e fatores	Futuro opcional
---
7. Metodologias aplicadas
7.1 Lean Logistics
O projeto aplica princípios de Lean Logistics ao reduzir desperdícios no fluxo de conferência fiscal/logística.
Exemplos de desperdícios atacados:
conferência manual repetitiva;
retrabalho por erro documental;
espera por validação humana;
falta de rastreabilidade;
movimentação desnecessária de informação entre arquivos.
A automação cria um fluxo mais enxuto:
```text
Entrada de documentos → Extração → Conferência → Diagnóstico → Saída estruturada
```
---
7.2 Kaizen
O desenvolvimento seguiu uma abordagem Kaizen, com melhoria contínua incremental.
Em vez de tentar construir o sistema completo de uma vez, o projeto foi evoluído em pequenos ciclos:
entender o problema;
testar XML;
testar TXT;
testar XLS;
testar PDF;
testar DOC/Word;
escolher o melhor formato;
validar extração;
validar cabeçalho;
deixar itens para a próxima fase.
Essa abordagem evitou complexidade prematura e permitiu aprender com os dados reais.
---
7.3 BPM — Business Process Management
O projeto também aplica BPM ao transformar um processo operacional em um fluxo estruturado e monitorável.
O processo foi mapeado em etapas:
```text
Receber NF-e XML
Receber OC
Extrair cabeçalho
Extrair itens
Validar vínculo por Doc. Origem
Gerar diagnóstico
Disponibilizar resultado
```
A partir disso, o processo deixa de ser apenas uma conferência manual e passa a ser um fluxo com:
entradas definidas;
regras de negócio explícitas;
saídas auditáveis;
logs;
possibilidade de indicadores.
---
7.4 Mapeamento de processos
Durante o MVP, foram identificados os documentos e suas melhores funções:
Fonte	Melhor uso
XML NF-e	Documento fiscal estruturado com cabeçalho e itens da nota
DOC/Word OC	Melhor formato encontrado para cabeçalho e itens da OC
PDF OC	Bom para evidência visual/auditoria, mas mais difícil para extração de itens
XLS OC	Bom para itens tabulares, mas fraco para cabeçalho
TXT/Relatório	Alternativa possível, mas dependente do layout
A conclusão foi que o Word/DOC é o melhor formato atual para comparar com o XML, pois reúne cabeçalho e itens em uma fonte única.
---
7.5 Governança de dados e LGPD
O projeto lida com documentos que podem conter informações sensíveis empresariais, como:
CNPJ;
endereço;
fornecedor;
cliente;
valores;
dados fiscais;
itens comprados.
Por isso, o repositório foi preparado para não versionar arquivos reais.
O `.gitignore` deve impedir o envio de:
XMLs reais;
DOCs reais;
PDFs reais;
XLS/XLSX reais;
CSVs de resultado;
arquivos de debug;
planilhas geradas.
Apenas o código, documentação e estrutura vazia devem ser versionados.
---
8. Como executar
8.1 Instalar dependências
```bash
pip install -r requirements.txt
```
8.2 Colocar arquivos de teste localmente
```text
entrada/xml/nota.xml
entrada/oc/oc.doc
```
8.3 Executar
```bash
python main.py
```
8.4 Ver resultados
```text
resultado/xml_cabecalho.csv
resultado/xml_itens.csv
resultado/oc_doc_cabecalho.csv
resultado/oc_doc_itens.csv
resultado/conferencia_cabecalho.csv
resultado/log_processamento.csv
resultado/resultado.xlsx
```
---
9. Resultado conquistado hoje
A principal conquista foi validar um fluxo funcional em ambiente local:
```text
XML NF-e + DOC OC → Extração → Cabeçalho/Itens → Conferência cabeçalho → CSV/Excel/Log
```
O sistema conseguiu:
ler XML local;
ler DOC/Word local;
extrair cabeçalho da OC;
extrair itens da OC;
extrair cabeçalho e itens da NF-e;
aplicar regra de Doc. Origem;
gerar resultado estruturado;
registrar log;
versionar o código com Git sem subir dados sensíveis.
---
10. Próximas fases
10.1 Comparação item a item
Comparar:
código do item;
descrição;
unidade;
embalagem;
quantidade;
valor unitário;
IPI;
valor total.
10.2 Normalização e fator
Criar lógica para tratar:
caixa;
pacote;
unidade;
embalagem;
quantidade base;
fator de conversão.
10.3 Saída para cards
Criar arquivo consolidado:
```text
resultado_cards.csv
```
Possíveis campos:
```text
numero_nota
numero_oc
doc_origem
fornecedor
status_cabecalho
status_itens
status_geral
problemas
data_processamento
```
10.4 Front-end
Possibilidades:
Streamlit para protótipo local;
SharePoint List para visual oficial;
Power Automate para popular cards;
Power Apps para interface corporativa.
10.5 Machine Learning
Machine learning pode ser avaliado futuramente para apoiar:
identificação de padrões de unidade/fator;
similaridade de descrições;
sugestão de conversões;
aprendizado com histórico de divergências.
No momento, a prioridade é resolver por regras, normalização e tabela de fatores.
---
11. Status atual
```text
MVP Parte 1: Concluído
Próxima fase: Comparação item a item
```
---
12. Resumo executivo
Este MVP automatiza a primeira camada de conferência entre NF-e e OC, aplicando princípios de Lean Logistics, Kaizen, BPM, mapeamento de processos, governança de dados e melhoria contínua.
A entrega atual transforma um processo manual em um fluxo estruturado, auditável e evolutivo, criando a base para comparação de itens, cards operacionais e futuras integrações com ferramentas Microsoft.
