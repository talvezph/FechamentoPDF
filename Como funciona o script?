Atuo no setor Administrativo, na área de Logística, em uma empresa de médio porte que possui diversas operações em andamento.

Dentro desse contexto, desempenho várias funções em diferentes operações. Em especial, sou responsável integral por uma delas, a qual envolve um processo bastante exaustivo, que exige a análise minuciosa de diversos arquivos em PDF para a realização de fechamentos quinzenais e mensais.

Diante dessa demanda repetitiva e detalhista, tive a iniciativa de desenvolver um programa que facilitasse a extração desses dados e, se possível, automatizasse grande parte — ou até mesmo todos — os cálculos necessários para a consolidação dos relatórios e posterior apresentação à diretoria.

Foi então que iniciei meus estudos com a linguagem Python. Apesar dos desafios, está sendo uma experiência extremamente enriquecedora. Sempre tive interesse e facilidade com tecnologia e computação, e esse tem sido meu primeiro projeto na área de programação — já com aplicação prática imediata, uma vez que quinzenalmente preciso entregar esses dados.


Basicamente o programa faz o seguinte processo:

1. Visão Geral

O script script_fechamento.py foi aprimorado para automatizar o processo de cálculo de fechamento de motoristas, extraindo dados de PDFs e consolidando-os em uma planilha Excel. As principais melhorias focam em robustez, configurabilidade e facilidade de uso.

2. Melhorias Implementadas

2.1. Refatoração para Configuração Externa

Todos os caminhos de arquivos e valores fixos (como VALOR_ENTREGA e BONUS_DIARIO) foram movidos para um arquivo de configuração config.ini. Isso permite que o usuário ajuste facilmente esses parâmetros sem modificar o código-fonte do script.

Exemplo de config.ini:

Plain Text


[Paths]
pdfs_folder = pdfs
type_sheet = Tipo de Veiculos.xlsx
output_excel = Fechamento_Motoristas.xlsx

[Values]
delivery_value = 3.80
daily_bonus = 30.00


2.2. Implementação de Logging e Tratamento de Erros

A biblioteca logging do Python foi integrada para registrar o fluxo de execução do script, bem como avisos e erros. Isso facilita a depuração e o monitoramento do script em produção. Mensagens informativas, de aviso e de erro são agora exibidas no console e coletadas para um relatório de erros.

2.3. Validação de Dados Extraídos

Foram adicionadas validações para os dados extraídos dos PDFs e lidos da planilha de veículos. O script agora verifica a presença de colunas essenciais na planilha de veículos e valida os formatos de datas e valores extraídos dos PDFs, registrando avisos ou erros conforme necessário.

2.4. Adição de Parâmetros de Linha de Comando

O script agora aceita argumentos de linha de comando usando a biblioteca argparse. Isso permite que o usuário especifique a pasta de PDFs, a planilha de tipos de veículos e o nome do arquivo Excel de saída diretamente ao executar o script, sobrescrevendo as configurações do config.ini se fornecidos.

Exemplo de uso:

Bash


python3.11 script_fechamento.py --pdfs_folder /caminho/para/meus_pdfs --type_sheet minha_planilha.xlsx --output_excel meu_fechamento.xlsx


2.5. Criação de Relatório de Erros

Um relatório consolidado de todos os erros e avisos registrados durante a execução do script é gerado em um arquivo separado (error_report.log por padrão). Isso fornece um resumo claro de quaisquer problemas encontrados, ajudando na identificação e resolução de falhas.

2.6. Implementação de Testes de Unidade

Testes de unidade foram desenvolvidos usando o framework unittest para garantir a funcionalidade e a robustez das principais funções do script. Isso inclui testes para normalização de texto, busca de nomes aproximados, extração de dados de PDFs e cálculo de fechamento. Os testes ajudam a verificar se as alterações não introduziram novos bugs e se o script se comporta conforme o esperado em diferentes cenários.

3. Como Usar o Script

1.
Configuração Inicial:

•
Certifique-se de que o arquivo config.ini esteja presente no mesmo diretório do script e configurado com os caminhos e valores padrão desejados.

•
Coloque a planilha de tipos de veículos (ex: Tipo de Veiculos.xlsx) no caminho especificado no config.ini ou via argumento de linha de comando.

•
Crie uma pasta para os PDFs (ex: pdfs) e coloque os PDFs dos motoristas dentro dela.



2.
Execução:

•
Execute o script a partir do terminal:

•
Para sobrescrever as configurações do config.ini:



3.
Verificação de Resultados:

•
O arquivo Excel de saída (Fechamento_Motoristas.xlsx por padrão) será gerado no diretório do script.

•
Verifique o arquivo error_report.log (se gerado) para quaisquer avisos ou erros durante a execução.



4. Estrutura de Arquivos

Plain Text


fechamento_motoristas/
├── script_fechamento.py
├── config.ini
├── Tipo de Veiculos.xlsx
├── pdfs/
│   ├── motorista_1.pdf
│   └── motorista_2.pdf
├── test_fechamento.py
└── error_report.log (gerado após a execução)


Avaliação sobre o projeto.

Considerando todo o processo que o programa realiza até o momento (com melhorias ainda em andamento), ele já proporcionou uma economia significativa de tempo, reduzindo de dois a três dias de trabalho manual para apenas alguns minutos. Isso me permite dedicar mais tempo à análise dos dados e à parte estratégica do projeto, em vez de concentrar esforços apenas na extração manual das informações.
