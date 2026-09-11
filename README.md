## 📚 Objetivo do projeto no GitHub


Este projeto faz parte do meu portfólio de **Excel, análise de dados, automação e modelagem de cenários**, demonstrando a aplicação prática de:

* Funções financeiras;
* Tratamento de dados;
* Integração com APIs;
* Estruturação de parâmetros;
* Automatização de cálculos;
* Modelagem de cenários;
* Organização de informações para tomada de decisão.








## 🎯 Objetivo do Simulador de Investimentos

O simulador foi desenvolvido para demonstrar, de forma simples e visual, como diferentes valores de aporte e períodos de investimento podem impactar a formação de patrimônio ao longo do tempo.

A planilha combina:

* Simulação de investimentos de longo prazo;
* Aportes iniciais e mensais;
* Taxa Selic obtida por API;
* Projeção de patrimônio acumulado;
* Estimativa de rendimento mensal;
* Sugestão de investimento baseada na renda mensal;
* Distribuição de aportes entre diferentes categorias de FIIs;
* Perfis de investimento **Conservador, Moderado e Agressivo**.



## 🚀 Funcionalidades

### Tesouro Selic

Permite informar:

* Aporte inicial;
* Aporte mensal;
* Período de investimento em anos.

A partir dessas informações, o simulador calcula:

* Sugestão de investimento mensal;
* Patrimônio acumulado;
* Rendimento mensal estimado;
* Informar até 4 projeções contando com o tempo em anos. O simulador inicial a sugestão com 2, 5, 10 e 15 anos.

### Fundos Imobiliários

O simulador permite selecionar um perfil de investimento e informar o valor do aporte mensal.

A partir do perfil selecionado, o valor é distribuído entre seis categorias:

* Papel;
* Tijolo;
* Híbridos;
* FOFs;
* Desenvolvimento;
* Hotelarias.

A distribuição percentual é armazenada na aba **Configuração** e recuperada automaticamente de acordo com o perfil selecionado.



## 🔌 Integração com API

A planilha utiliza uma API do **Banco Central do Brasil (BCB)** para obter a taxa utilizada na simulação.

Endpoint configurado na planilha:

```text
https://api.bcb.gov.br/dados/serie/bcdata.sgs.1178/dados/ultimos/1?formato=json
```

O retorno da API possui estrutura semelhante a:

```json
[
  {
    "data": "10/09/2026",
    "valor": "13.90"
  }
]
```

A planilha utiliza o valor retornado pela API para alimentar automaticamente o cálculo da taxa.



## 🧮 Cálculos

### 1. Sugestão de investimento

A planilha considera **30% do salário mensal** como sugestão de valor para investimento.

**Fórmula utilizada:**

```excel
=D13*30%
```

Onde:

* `D13` = salário mensal;
* `30%` = percentual sugerido para investimento.



### 2. Conversão da taxa obtida pela API

Como o retorno da API é recebido em formato JSON, a planilha utiliza uma sequência de substituições para extrair o valor numérico.

**Fórmula utilizada:**

```excel
=VALUE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(RIGHT(B11,8),"""",""),"}",""),"]",""),".",","))/100
```

Essa fórmula:

1. Extrai os últimos caracteres do retorno;
2. Remove aspas;
3. Remove `}`;
4. Remove `]`;
5. Substitui o ponto decimal por vírgula;
6. Converte o resultado em número;
7. Divide por `100` para obter o percentual em formato decimal.



### 3. Patrimônio acumulado

O patrimônio projetado utiliza a função financeira `PV` (Valor Presente) para calcular o valor futuro dos aportes mensais, somando o aporte inicial.

**Fórmula utilizada:**

```excel
=-PV((1+D21)^(1/12)-1,D20*12,D19)+D18
```

Onde:

* `D21` = taxa utilizada na simulação;
* `D20` = período de investimento em anos;
* `D19` = aporte mensal;
* `D18` = aporte inicial;
* `12` = quantidade de meses por ano.

A expressão:

```excel
(1+D21)^(1/12)-1
```

é utilizada para obter uma taxa mensal equivalente a partir da taxa informada.



### 4. Rendimento mensal estimado

O rendimento mensal é calculado aplicando a taxa mensal sobre o patrimônio acumulado.

**Fórmula utilizada:**

```excel
=D23*((1+D21)^(1/12)-1)
```

Onde:

* `D23` = patrimônio acumulado;
* `D21` = taxa utilizada na simulação.



### 5. Projeções de patrimônio

O simulador também calcula automaticamente o patrimônio estimado para diferentes períodos.

Os períodos utilizados são:

* 2 anos;
* 5 anos;
* 10 anos;
* 15 anos.

**Exemplo para 2 anos:**

```excel
=-PV((1+D21)^(1/12)-1,B27*12,D19)+D18
```

O mesmo conceito é aplicado aos demais períodos, alterando apenas a quantidade de anos.



### 6. Rendimento projetado

Para cada período, o rendimento mensal estimado é calculado sobre o patrimônio correspondente.

**Exemplo:**

```excel
=C27*((1+D21)^(1/12)-1)
```



## 🏢 Distribuição de FIIs

A distribuição dos aportes é baseada no perfil selecionado.

A planilha utiliza uma chave composta pelo perfil e pelo tipo de FII para localizar automaticamente o percentual correspondente na aba **Configuração**.

**Fórmula utilizada:**

```excel
=VLOOKUP($D$34&"-"&B38,Configuração!$B:$E,4,FALSE)
```

A chave é formada por:

```text
Perfil + Tipo de FII
```

Exemplo:

```text
Moderado-PAPEL
```



### 💰 Valor destinado a cada categoria

Após identificar o percentual recomendado, o valor do aporte é multiplicado pelo percentual correspondente.

**Fórmula utilizada:**

```excel
=$D$35*C38
```

Onde:

* `D35` = aporte mensal destinado aos FIIs;
* `C38` = percentual recomendado para a categoria.



## 📈 Perfis de investimento

A aba **Configuração** contém três perfis:

| Perfil         | Característica                                                       |
| -------------- | -------------------------------------------------------------------- |
| 🟢 Conservador | Maior concentração em ativos/categorias consideradas mais defensivas |
| 🟡 Moderado    | Distribuição intermediária entre as categorias                       |
| 🔴 Agressivo   | Maior exposição a categorias de maior potencial de risco/retorno     |

### Distribuição configurada

| Tipo de FII     | Conservador | Moderado | Agressivo |
| --------------- | ----------: | -------: | --------: |
| Papel           |         30% |      32% |       50% |
| Tijolo          |         50% |      35% |       10% |
| Híbridos        |         10% |       8% |        5% |
| FOFs            |         10% |       5% |        5% |
| Desenvolvimento |          0% |      10% |       20% |
| Hotelarias      |          0% |      10% |       10% |
| **Total**       |    **100%** | **100%** |  **100%** |



## 📁 Estrutura da planilha

O arquivo está organizado em três abas principais:

### `APP`

Interface principal do simulador.

Contém:

* Taxa Selic;
* Dados do investidor;
* Simulação do Tesouro Selic;
* Projeções;
* Distribuição de FIIs;
* Seleção de perfil.

### `Configuração`

Centraliza parâmetros utilizados pelo simulador, incluindo:

* Endpoint da API;
* Perfis de investimento;
* Categorias de FIIs;
* Percentuais de distribuição;
* Chaves utilizadas nas buscas.



## 🛠️ Tecnologias e recursos utilizados

* **WPS-Planilhas**
* Fórmulas financeiras do Excel
* `PV`
* `VLOOKUP`
* `VALUE`
* `SUBSTITUTE`
* `RIGHT`
* Integração com **API do Banco Central do Brasil**
* JSON
* Modelagem de cenários de investimento


## ▶️ Como utilizar

1. Abra o arquivo Excel.
2. Acesse a aba **APP**.
3. Informe seu salário mensal.
4. Informe o aporte inicial.
5. Informe o valor que pretende investir mensalmente.
6. Informe o período desejado para o investimento.
7. Consulte o patrimônio e rendimento projetados.
8. Para a simulação de FIIs, selecione o perfil de investimento.
9. Informe o valor do aporte mensal destinado aos FIIs.
10. Consulte a distribuição sugerida entre as categorias.



## ⚠️ Observações importantes

Este projeto possui **finalidade educacional e de demonstração de conceitos de Excel, automação e simulação financeira**.

Os resultados apresentados são projeções matemáticas baseadas nos parâmetros informados e **não representam garantia de rentabilidade ou recomendação de investimento**.

Rentabilidade passada ou taxas de referência não garantem resultados futuros.

A distribuição entre FIIs apresentada pelo simulador também não deve ser interpretada como recomendação individual de investimento.



## 📚 Objetivo do projeto no GitHub

Este projeto faz parte do meu portfólio de **Excel, análise de dados, automação e modelagem de cenários**, demonstrando a aplicação prática de:

* Funções financeiras;
* Tratamento de dados;
* Integração com APIs;
* Estruturação de parâmetros;
* Automatização de cálculos;
* Modelagem de cenários;
* Organização de informações para tomada de decisão.
