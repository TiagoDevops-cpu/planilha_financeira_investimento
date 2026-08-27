```python
import pandas as pd

file_path = 'Planilha_Financeira_investimento.xlsx'
xls = pd.ExcelFile(file_path)
print("Sheet names:", xls.sheet_names)

for sheet in xls.sheet_names:
    df = pd.read_excel(file_path, sheet_name=sheet)
    print(f"\n--- Sheet: {sheet} ---")
    print("Columns:", df.columns.tolist())
    print("Shape:", df.shape)
    print("Head:\n", df.head(3))


```

```text
Sheet names: ['Metta_Invest', 'Planilha1']

--- Sheet: Metta_Invest ---
Columns: ['Unnamed: 0', 'Unnamed: 1', 'Unnamed: 2', 'Unnamed: 3']
Shape: (40, 4)
Head:
    Unnamed: 0 Unnamed: 1 Unnamed: 2 Unnamed: 3
0         NaN        NaN        NaN        NaN
1         NaN        NaN        NaN        NaN
2         NaN        NaN        NaN        NaN

--- Sheet: Planilha1 ---
Columns: ['Unnamed: 0', 'Unnamed: 1', 'Unnamed: 2', 'Unnamed: 3', 'Unnamed: 4', 'Unnamed: 5', 'Unnamed: 6', 'Unnamed: 7', 'Unnamed: 8']
Shape: (19, 9)
Head:
            Unnamed: 0   Unnamed: 1   Unnamed: 2           Unnamed: 3  Unnamed: 4  Unnamed: 5  Unnamed: 6       Unnamed: 7 Unnamed: 8
0               CHAVE       PERFIL  TIPO DE FII  Percentual Sugerido         NaN         NaN         NaN              NaN        NaN
1  Conservador-TIJOLO  Conservador       TIJOLO                  0.5         NaN         NaN         NaN              NaN          %
2   Conservador-PAPEL  Conservador        PAPEL                  0.3         NaN         NaN         NaN  Moderado-TIJOLO        0.4


```

```python
df_metta = pd.read_excel(file_path, sheet_name='Metta_Invest')
print("Metta_Invest non-empty rows:")
print(df_metta.dropna(how='all'))

df_p1 = pd.read_excel(file_path, sheet_name='Planilha1')
print("\nPlanilha1:")
print(df_p1.dropna(how='all'))


```

```text
Metta_Invest non-empty rows:
    Unnamed: 0                      Unnamed: 1           Unnamed: 2     Unnamed: 3
8          NaN                   CONFIGURAÇÕES                  NaN            NaN
9          NaN                         Salário                  NaN           5000
10         NaN              Redimento Carteira                  NaN          0.006
11         NaN  Sugestão de Investimento (30%)                  NaN           1500
13         NaN             INVESTIMENTO MENSAL                  NaN            NaN
14         NaN        Quanto Investir por mês?                  NaN            500
15         NaN               Por Quantos Anos?                  NaN             10
16         NaN      Taxa de Rendimento mensal?                  NaN        0.01079
17         NaN         Patrimônio acumulado?\n                  NaN  121642.106265
18         NaN             Dividendos Mensais?                  NaN     729.852638
20         NaN                        CENÁRIOS                  NaN      Dividendo
21         2.0                Quanto em 2 Anos         13613.813649      81.682882
22         5.0                Quanto em 5 Anos         41888.456999     251.330742
23        10.0               Quanto em 10 Anos        121642.106265     729.852638
24        20.0               Quanto em 20 Anos        562599.200049      3375.5952
25        30.0               Quanto em 30 Anos       2161084.827502   12966.508965
29         NaN                          PERFIL             Moderado            NaN
30         NaN   VALOR A SER INVESTIDO POR MÊS                  500            NaN
32         NaN                     TIPO DE FII  Percentual Sugerido        Valores
33         NaN                          TIJOLO                  0.4            200
34         NaN                           PAPEL                 0.32            160
35         NaN                        HÍBRIDOS                 0.08             40
36         NaN                            FOFs                  0.1             50
37         NaN                 DESENVOLVIMENTO                 0.05             25
38         NaN                      HOTELARIAS                 0.05             25
39         NaN                             NaN                  NaN            500

Planilha1:
                     Unnamed: 0   Unnamed: 1       Unnamed: 2           Unnamed: 3  Unnamed: 4  Unnamed: 5  Unnamed: 6       Unnamed: 7 Unnamed: 8
0                         CHAVE       PERFIL      TIPO DE FII  Percentual Sugerido         NaN         NaN         NaN              NaN        NaN
1            Conservador-TIJOLO  Conservador           TIJOLO                  0.5         NaN         NaN         NaN              NaN          %
2             Conservador-PAPEL  Conservador            PAPEL                  0.3         NaN         NaN         NaN  Moderado-TIJOLO        0.4
3          Conservador-HÍBRIDOS  Conservador         HÍBRIDOS                  0.1         NaN         NaN         NaN              NaN        NaN
4              Conservador-FOFs  Conservador             FOFs                  0.1         NaN         NaN         NaN              NaN        NaN
5   Conservador-DESENVOLVIMENTO  Conservador  DESENVOLVIMENTO                    0         NaN         NaN         NaN              NaN        NaN
6        Conservador-HOTELARIAS  Conservador       HOTELARIAS                    0         NaN         NaN         NaN              NaN        NaN
7               Moderado-TIJOLO     Moderado           TIJOLO                  0.4         NaN         NaN         NaN              NaN        NaN
8                Moderado-PAPEL     Moderado            PAPEL                 0.32         NaN         NaN         NaN              NaN        NaN
9             Moderado-HÍBRIDOS     Moderado         HÍBRIDOS                 0.08         NaN         NaN         NaN              NaN        NaN
10                Moderado-FOFs     Moderado             FOFs                  0.1         NaN         NaN         NaN              NaN        NaN
11     Moderado-DESENVOLVIMENTO     Moderado  DESENVOLVIMENTO                 0.05         NaN         NaN         NaN              NaN        NaN
12          Moderado-HOTELARIAS     Moderado       HOTELARIAS                 0.05         NaN         NaN         NaN              NaN        NaN
13             Agressivo-TIJOLO    Agressivo           TIJOLO                  0.1         NaN         NaN         NaN              NaN        NaN
14              Agressivo-PAPEL    Agressivo            PAPEL                  0.5         NaN         NaN         NaN              NaN        NaN
15           Agressivo-HÍBRIDOS    Agressivo         HÍBRIDOS                 0.05         NaN         NaN         NaN              NaN        NaN
16               Agressivo-FOFs    Agressivo             FOFs                 0.05         NaN         NaN         NaN              NaN        NaN
17    Agressivo-DESENVOLVIMENTO    Agressivo  DESENVOLVIMENTO                  0.2         NaN         NaN         NaN              NaN        NaN
18         Agressivo-HOTELARIAS    Agressivo       HOTELARIAS                  0.1         NaN         NaN         NaN              NaN        NaN


```

Aqui está um modelo completo de **`README.md`** pronto para documentar o seu projeto no GitHub ou repositório de escolha.

---

# 📊 Planilha Financeira de Investimentos em FIIs

Uma ferramenta em Excel projetada para planejamento financeiro, projeção de acúmulo de patrimônio a longo prazo e alocação estratégica em **Fundos de Investimento Imobiliário (FIIs)** de acordo com o perfil do investidor.

---

## 🚀 Funcionalidades

* **Configurações de Aporte & Renda:** Cálculo da sugestão de investimento mensal com base no salário e rendimento da carteira.
* **Projeção de Patrimônio e Dividendos:** Simulação de juros compostos para projeção de patrimônio acumulado e dividendos mensais estimados.
* **Análise de Cenários:** Simulações automáticas para horizontes de tempo de 2, 5, 10, 20 e 30 anos.
* **Alocação por Perfil de Risco:** Distribuição percentual automatizada dos investimentos por categoria de FIIs de acordo com o perfil selecionado:
* **Conservador**
* **Moderado**
* **Agressivo**



---

## 📁 Estrutura do Arquivo

O arquivo `Planilha_Financeira_investimento.xlsx` está dividido em duas abas principais:

### 1. `Metta_Invest` (Painel Principal)

* **Configurações:** Entrada de salário, taxa de rendimento esperada e alocação recomendada (ex: 30% da renda).
* **Calculadora de Investimento Mensal:** Definição do aporte mensal, taxa de juros e prazos para estimativa de resultado final.
* **Tabela de Cenários:** Visão detalhada do patrimônio e dos dividendos acumulados ao longo dos anos.
* **Divisão por Perfil:** Distribuição do aporte mensal em reais (R$) entre os diferentes tipos de FIIs com base no perfil escolhido.

### 2. `Planilha1` (Matriz de Alocação / Procv)

* Tabela contendo a matriz de percentuais recomendados para cada perfil e tipo de fundo imobiliário:
* Tijolo
* Papel
* Híbridos
* FOFs (Fundos de Fundos)
* Desenvolvimento
* Hotelaria



---

## 📐 Estrutura de Alocação por Perfil

| Tipo de FII | Conservador | Moderado | Agressivo |
| --- | --- | --- | --- |
| **Tijolo** | 50% | 40% | 10% |
| **Papel** | 30% | 32% | 50% |
| **Híbridos** | 10% | 8% | 5% |
| **FOFs** | 10% | 10% | 5% |
| **Desenvolvimento** | 0% | 5% | 20% |
| **Hotelaria** | 0% | 5% | 10% |

---

## 🛠️ Como Usar

1. Baixe o arquivo `Planilha_Financeira_investimento.xlsx`.
2. Abra no **Microsoft Excel**, **Google Planilhas** ou software compatível.
3. Na aba **`Metta_Invest`**:
* Insira o valor do seu **Salário** e o valor do **Aporte Mensal**.
* Ajuste a **Taxa de Rendimento Mensal** estimada.
* Selecione seu **Perfil de Investidor** (Conservador, Moderado ou Agressivo) para ver o desdobramento do investimento por categoria de FII.
