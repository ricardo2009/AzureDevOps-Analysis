# Azure DevOps Analysis

Este repositório contém um script Python que analisa e consolida dados de projetos, repositórios, pipelines, commits e agentes do Azure DevOps. Os resultados são salvos em um arquivo Excel para fácil visualização e análise.

## Como usar

1. Clone este repositório.
2. Instale as dependências necessárias com `pip install -r requirements.txt`.
3. Atualize as variáveis `personal_access_token` e `organization_url` no script com suas próprias informações.
4. Execute o script com `python main.py`.

## Dependências

Este script depende das seguintes bibliotecas Python:

- azure-devops
- pandas
- openpyxl

## Resultados

Os resultados são salvos em um arquivo Excel chamado `analysis_result.xlsx`. Este arquivo contém várias planilhas com dados sobre projetos, repositórios, pipelines, commits e agentes.

## Contribuições

Contribuições são bem-vindas! Por favor, abra uma issue ou pull request se você tiver melhorias ou correções para sugerir.
