# 🤖 Automatizador DIMOB: Adeus à Digitação Manual!

Este projeto é a solução definitiva para imobiliárias, administradoras e escritórios de contabilidade que enfrentam o desafio anual de preencher a DIMOB (Declaração de Informações sobre Atividades Imobiliárias).

Sabemos que este é um processo manual, propenso a erros e que consome um tempo precioso. O Automatizador DIMOB foi criado para transformar essa tarefa, lendo os dados diretamente de uma planilha Excel e preenchendo o programa da Receita Federal de forma rápida, precisa e segura.

## ✨ Funcionalidades Principais

- **Automação Completa**: Preenche tanto os dados fixos da sua empresa (CNPJ, Endereço) quanto os dados variáveis de cada contrato (locatário, datas, valores) a partir de uma planilha.
- **Leitura Direta do Excel**: Utiliza a biblioteca pandas para carregar todos os seus registros de uma só vez.
- **Digitação Segura**: Simula a digitação humana, inserindo dados caractere por caractere em campos sensíveis para evitar falhas de preenchimento comuns em softwares do governo.
- **Parada de Emergência**: A qualquer momento, a execução pode ser interrompida de forma segura simplesmente pressionando a tecla Esc.
- **Altamente Configurável**: Todas as coordenadas e valores fixos estão centralizados no início do código, facilitando a adaptação para a sua realidade.
- **Lógica Inteligente**: Lida com formatação de datas, conversão de valores numéricos (ponto para vírgula) e até a seleção de UF e Município em menus dropdown.

## ⚙️ Como Funciona

O robô segue um fluxo lógico e cuidadosamente cronometrado para garantir o sucesso do preenchimento:

### Pausa Inicial
Após a execução, o script aguarda 5 segundos. Este é o tempo que você tem para clicar na janela do programa da DIMOB, garantindo que ela esteja em foco.

### Leitura da Planilha
O arquivo Excel especificado é carregado na memória.

### Loop de Preenchimento
O script começa a percorrer cada linha da sua planilha, que representa um contrato/registro a ser declarado.

### Inserção de Dados
Para cada registro, ele:

- Preenche os dados fixos da sua empresa.
- Preenche os dados do locatário e do contrato, lidos da linha atual da planilha.
- Realiza os cliques necessários para selecionar UF e Município.
- Clica nos botões "OK" e "Incluir" para salvar o registro no programa da DIMOB.
- Pausa brevemente entre as ações para garantir que o programa da DIMOB processe cada comando.

### Monitoramento Contínuo
A cada novo registro, o script verifica se a tecla Esc foi pressionada, permitindo uma interrupção segura.

### Conclusão
Ao final de todas as linhas da planilha, uma mensagem de sucesso é exibida no terminal.

