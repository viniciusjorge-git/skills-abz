---
name: formulario-reembolso
description: Gera o formulário padrão de reembolso/prestação de contas da Albatroz em Excel (.xlsx) a partir de invoices, recibos, notas e comprovantes enviados pelo usuário, com conversão PTAX para despesas em moeda estrangeira. Usar quando o usuário pedir um reembolso, prestação de contas, ou enviar comprovantes/invoices para reembolsar.
---

# Formulário de Reembolso (Albatroz)

Gera a planilha padrão de reembolso da Albatroz a partir dos comprovantes fornecidos (PDFs, fotos de recibos, invoices, faturas de cartão).

## Passo 1 — Coletar e ler os comprovantes

1. Receba os comprovantes por anexo no chat ou em uma pasta conectada. Leia todos antes de montar qualquer planilha.
2. PDFs com texto: extrair com `pdftotext -layout`. Fotos de recibos (jpeg/png): ler como imagem e transcrever. Faturas de cartão protegidas por senha: pedir a senha ao usuário e abrir com `qpdf --password=<senha> --decrypt`.
3. De cada comprovante, extrair: **data da compra**, **estabelecimento/fornecedor**, **valor**, **moeda**, **identificador** (nº da invoice, recibo, check, order id, folio) e, quando houver, a confirmação na fatura do cartão.
4. Se houver reembolsos anteriores disponíveis (planilhas na pasta ou no histórico), conferir item a item que nenhum comprovante novo já foi reembolsado — compare por número de invoice/recibo, data e valor. Itens já reembolsados ficam FORA da nova planilha; avise o usuário.

## Passo 2 — Perguntar os dados que variam por pessoa

Usar AskUserQuestion (ou perguntar em texto) ANTES de montar a planilha. O gestor muda de pessoa para pessoa — **nunca assumir**:

- **Nome do gestor imediato** (obrigatório perguntar sempre)
- Nome completo do solicitante e CPF (se não conhecidos na conversa)
- Dados para pagamento: banco e chave PIX
- Se a viagem/despesa tem um nome ou evento associado (ex.: "ITC Vegas 2026"), para o subtítulo

## Passo 3 — Converter despesas em moeda estrangeira

Regra da empresa: **cotação PTAX de venda do dia da compra**, do Banco Central do Brasil.

1. Buscar a PTAX na API SGS do BCB (série 1 = dólar venda):
   `https://api.bcb.gov.br/dados/serie/bcdata.sgs.1/dados?formato=json&dataInicial=DD/MM/AAAA&dataFinal=DD/MM/AAAA`
2. Compra em fim de semana/feriado sem cotação: usar a PTAX do dia útil anterior e anotar isso na descrição.
3. Despesa em dólar cobrada no **cartão de crédito brasileiro**: usar o valor em R$ da fatura + a linha de **IOF Transações Exterior** da própria fatura (lançar o IOF como linha separada). Nesse caso NÃO usar PTAX — o valor real em reais já existe.
4. Despesa paga em dólar (conta internacional, cartão em USD): converter por `US$ × PTAX`, sempre por **fórmula** na planilha, nunca valor pré-calculado.

## Passo 4 — Montar a planilha (openpyxl)

Arquivo: `Formulário_Reembolso_<assunto>_<AAAA-MM>.xlsx`. Fonte Arial em tudo. Usar fórmulas (SUM, multiplicações), nunca totais hardcoded.

**Cabeçalho:**
- Linhas 2–3 (merge A:J): título "FORMULÁRIO DE REEMBOLSO E/OU PRESTAÇÃO DE CONTAS" (Arial 14, negrito, centralizado)
- Linha 4 (merge A:J): subtítulo com o evento/viagem ou "(REFERENTE A UM ADIANTAMENTO REALIZADO)"
- Bloco de identificação (rótulos com fundo cinza D9D9D9 e bordas finas): NOME COMPLETO, CPF, UNIDADE, GESTOR IMEDIATO (valor perguntado no Passo 2); à direita, DADOS PARA PAGAMENTO com BANCO e CHAVE PIX
- VALOR ADIANTADO (0 se não houver) e DATA DA TRANSFERÊNCIA

**Tabela principal (cabeçalho na linha 16, colunas A–J, fundo cinza):**
A: DATA DA NOTA FISCAL E/OU USO EM KM · B: OBSERVAÇÕES (NOME DO CLIENTE E/OU EMPRESA) · C: KM RODADO · D: KM EM REAIS · E: ALIMENTAÇÃO · F: LOCOMOÇÃO (TÁXI E/OU APLICATIVOS) · G: ESTACIONAMENTO · H: PEDÁGIO · I: OUTROS NÃO INFORMADO · J: DESCRIÇÃO DE "OUTROS" / DETALHE

- Uma linha por despesa, valor na coluna da categoria certa (refeições → E; Uber/táxi → F; hospedagem, assinaturas, software e demais → I)
- Data como data real (formato DD/MM/YYYY); valores com formato `#,##0.00`
- Coluna J: descrição completa — o que é, nº da invoice/recibo, onde foi confirmado (fatura ou extrato) e, se convertido, `US$ X × PTAX Y`
- IOF de cartão: linha própria, logo abaixo da despesa, descrição "IOF s/ compra ... (Fatura <mês>)"

**Totais:**
- Linha "VALOR POR DESCRIÇÃO": `=SUM(...)` por coluna D a I + total geral em J
- Duas linhas abaixo: "TOTAL A REEMBOLSAR" (merge G:I, à direita) com `="R$" #,##0.00` apontando para o total

**Bloco MEMÓRIA DE CONVERSÃO US$ → R$** (abaixo do total, apenas se houver despesas em moeda estrangeira):
Colunas: DATA DA COMPRA · DESPESA · COMPROVANTE · VALOR (US$) · PTAX VENDA (R$/US$) · DATA DA COTAÇÃO · VALOR CONVERTIDO (R$) · FONTE DA COTAÇÃO
- US$ e PTAX em fonte azul (0000FF) = dados de entrada; célula da PTAX com fundo amarelo (FFFF00)
- VALOR CONVERTIDO por fórmula `=D*E`; FONTE: "Banco Central do Brasil - PTAX venda (SGS série 1)"
- Linha de totais US$ e R$; os valores da tabela principal devem **referenciar** estas células (ex.: `=G35`), para recálculo automático
- Nota abaixo: referência ao boletim PTAX/site do BCB e a data da consulta

**Rodapé:** linhas de assinatura (solicitante à esquerda, gestor à direita, borda superior) e legenda explicando azul/amarelo e a regra de conversão.

## Passo 5 — Verificar e entregar

1. Recalcular com LibreOffice headless (ou o script `recalc.py` da skill xlsx, se disponível) — zero erros de fórmula.
2. Reabrir com `data_only=True` e conferir: total geral = soma manual dos itens; cada conversão = US$ × PTAX.
3. Enviar o arquivo ao usuário e, se houver pasta conectada, salvar na pasta dos comprovantes.
4. Resumir em uma tabela curta: categorias, totais em US$ e R$, PTAX usadas — e apontar pendências (recibo sem gorjeta, fatura faltando, despesa suspeita de duplicidade).

## Regras fixas

- Nunca incluir despesa sem comprovante legível; listar as que ficaram de fora e por quê.
- Nunca inventar cotação: sempre da API do BCB, com a data usada registrada na planilha.
- Valores de entrada em azul, cotações em amarelo, todo o resto calculado por fórmula.
- Moeda padrão do formulário é o Real (R$); o total a reembolsar é sempre em R$.
