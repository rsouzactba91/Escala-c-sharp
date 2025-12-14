# Sistema de Escala (WinForms)

Projeto desenvolvido em C# para gerenciar escalas de trabalho. A ideia principal é pegar uma tabela com a escala do mês inteiro e filtrar automaticamente para mostrar apenas a operação do dia (quem trabalha hoje).

## O que o sistema faz
1.  **Visão Mensal:** Permite lançar a escala de todos os funcionários para o mês.
2.  **Visão Diária:** O sistema lê a escala mensal e gera uma lista só com quem está escalado para o dia atual.
3.  **Clima:** Mostra a previsão do tempo de Curitiba/PR (via API) para ajudar no planejamento do dia.

## Tecnologias
* C# (Windows Forms)
* Integração com API externa (HG Weather)
* Manipulação de JSON

## Status
Em desenvolvimento. Atualmente trabalhando na lógica de conversão da grade mensal para diária.

---
