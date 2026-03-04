# Programa-Vendas-Excel

# Sistema de Vendas – Excel VBA

## Visão Geral

Sistema de ponto de venda (PDV) desenvolvido em Excel com VBA para registro de vendas, controle acumulado por forma de pagamento e consolidação diária.

Os dados são persistidos em arquivos `.txt` externos utilizando modo `Append`.

---

## Estrutura da Planilha

### Área de Vendas

* Coluna B: Produto
* Coluna C: Valor unitário
* Coluna D: Quantidade
* Coluna E: Sub-total
* B16: Nome do comprador
* E16: Total da venda

### Controle Diário

* C20: Total Dinheiro
* C21: Total Cartão
* C22: Total Geral

---

## Fluxo de Operação

1. Incremento de produtos via botões (macros de clique).
2. Inserção do nome do comprador em B16.
3. Registro da venda:

   * `Retângulo4_Clique` → Dinheiro
   * `Retângulo2_Clique` → Cartão
4. Execução de `SalvarValorEDataEmArquivo`.
5. Reset de quantidades e nome.
6. Fechamento do dia via `Retângulo1_Clique`.

---

## Persistência de Dados

### 1. Log Transacional

Arquivo:
`Vendas.txt`

Formato:

```
FormaPagamento,Valor,DataHora,NomeComprador
```

Exemplo:

```
Cartão,240.60,03/03/2026 20:45:12,Alan
```

Cada venda é registrada individualmente.

---

### 2. Consolidação Diária

Arquivo:
`vendas_dia.txt`

Formato:

```
Dinheiro,Valor,Data
Cartão,Valor,Data
Geral,Valor,Data
```

Exemplo:

```
Dinheiro,290.50,03/03/2026
Cartão,120.00,03/03/2026
Geral,410.50,03/03/2026
```

Gerado ao executar o fechamento do dia.

---

## Principais Macros

### Incremento de Produtos

Padrão utilizado:

```vb
Range("D[x]").Value = Range("D[x]").Value + 1
```

Cada botão incrementa a quantidade do respectivo produto.

---

### Registro de Venda

```vb
Sub SalvarValorEDataEmArquivo(formaPagamento As String)
```

Responsável por:

* Capturar total da venda (E16)
* Capturar nome do comprador (B16)
* Registrar data e hora
* Persistir dados no log

---

### Fechamento do Dia

Macro:

```
Retângulo1_Clique
```

Funções:

* Confirmação via `MsgBox`
* Registro do resumo diário
* Reset dos acumuladores

---

## Estrutura de Diretórios

Caminho padrão de armazenamento:

```
C:\Users\aland\Documents\Projetos Excel\Programa Vendas Excel\Dados\
```

Arquivos gerados:

* `Vendas.txt`
* `vendas_dia.txt`

---

## Observações Técnicas

* Persistência realizada via `Open ... For Append`.
* Separação entre log transacional e consolidação.
* Reset automático pós-venda.
* Estrutura modular com responsabilidade separada entre registro e fechamento.
