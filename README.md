# Mark - Sistema de Etiquetas Logísticas

O Mark é um projeto desenvolvido para automatizar a geração de etiquetas logísticas a partir de uma base de dados operacional.

O objetivo do projeto é reduzir retrabalho, padronizar etiquetas e agilizar o processo de separação e expedição.

## Objetivo

Automatizar a criação de etiquetas por carga, código de produto, descrição, transportadora e quantidade de volumes.

## Problema

Antes da automação, a geração de etiquetas era um processo manual, repetitivo e sujeito a erros, principalmente quando havia muitos itens, volumes e cargas diferentes.

## Solução

O sistema utiliza Excel e VBA para:

- Ler uma base de dados de cargas
- Agrupar itens por carga e código
- Calcular volumes totais
- Gerar etiquetas automaticamente
- Criar sequência de volumes, como 1/4, 2/4, 3/4 e 4/4
- Permitir visualização antes da impressão
- Padronizar o layout das etiquetas

## Tecnologias utilizadas

- Excel
- VBA
- Automação de processos
- Tratamento de dados
- Lógica operacional
- Processos logísticos

## Funcionalidades

- Geração automática de etiquetas
- Agrupamento por carga e código
- Cálculo automático de volumes
- Impressão automatizada
- Visualização antes da impressão
- Padronização das informações da etiqueta

## Regras de negócio

Exemplo de regras utilizadas no projeto:

- Bombona: 48 unidades por palete
- Caixa 24x1: 24 volumes por palete
- Caixa 12x500 / 12x1: 62 volumes por palete

## Estrutura do projeto

```text
screenshots/   Prints do sistema
sample-data/   Base fictícia para demonstração
vba/           Códigos VBA exportados
docs/          Documentação das regras de negócio
