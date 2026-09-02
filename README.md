<div align="center">

# 📱 Tabela de Preço — TIM Revendas

**Consulta de preços de aparelhos e acessórios por plano TIM**  
Atualizado semanalmente conforme tabela oficial do canal Revendas.

[![Status](https://img.shields.io/badge/status-ativo-brightgreen?style=flat-square)](https://hermes-ecaflip.github.io/Tabela-de-preco-da-TIM/)
[![Última atualização](https://img.shields.io/badge/tabela-02%2F09%2F2026-blue?style=flat-square)](https://hermes-ecaflip.github.io/Tabela-de-preco-da-TIM/)
[![Canal](https://img.shields.io/badge/canal-Revendas-orange?style=flat-square)](https://hermes-ecaflip.github.io/Tabela-de-preco-da-TIM/)
[![Licença](https://img.shields.io/badge/licen%C3%A7a-MIT-lightgrey?style=flat-square)](LICENSE)

</div>

---

## 📖 Sobre o Projeto

Ferramenta web para consulta rápida de preços de aparelhos e acessórios comercializados no canal **Revendas TIM**, com suporte a múltiplos planos (fidelizados e não fidelizados).

Desenvolvida para uso interno de vendedores e revendedores, permitindo consultar em segundos o valor de qualquer aparelho em qualquer plano disponível na tabela oficial.

---

## ✨ Funcionalidades

| Funcionalidade | Descrição |
|---|---|
| 🔍 **Busca em tempo real** | Pesquisa instantânea por nome do aparelho ou acessório |
| 📱 **276 itens catalogados** | Celulares, tablets, acessórios TIM e itens de loja |
| 💰 **Preços sem fidelização** | PRÉ (balcão), Base Faturamento, Controle e Pós não fidelizados |
| 📋 **Planos fidelizados** | Mobile (celular) e TIM Fibra (internet), com todos os subplanos |
| 🏷️ **Logos das marcas** | Apple, Samsung, Motorola, Xiaomi, Nokia, JBL, Sony, PlayStation, Xbox e mais |
| 📅 **Data de vigência** | Exibe data de atualização e validade de cada produto |
| 💡 **Comparativo de economia** | Mostra a diferença entre o preço de balcão e o plano selecionado |
| 📱 **100% responsivo** | Otimizado para celular, tablet, notebook e desktop |
| ⚡ **Boost Trade In** | Ativa/desativa o desconto de troca direto no card, quando o produto tem Boost |
| 📚 **Catálogo por plano** | Lista os produtos com o preço já aplicado ao plano do cliente, com filtros de marca, tipo e faixa de preço digitável |

---

## 🖥️ Demonstração

> Acesse pelo GitHub Pages:  
> **[https://hermes-ecaflip.github.io/Tabela-de-preco-da-TIM/](https://hermes-ecaflip.github.io/Tabela-de-preco-da-TIM/)**

---

## 🚀 Como Usar

### 1. Clone o repositório

```bash
git clone https://github.com/hermes-ecaflip/Tabela-de-preco-da-TIM.git
cd Tabela-de-preco-da-TIM
```

### 2. Abra no navegador

Basta abrir o arquivo `index.html` diretamente no navegador — não requer servidor ou dependências externas.

```bash
# macOS / Linux
open index.html

# Windows
start index.html
```

### 3. Pesquise um aparelho

Digite o nome do aparelho ou acessório no campo de busca, selecione o plano desejado e o preço será exibido automaticamente.

---

## 🔄 Como Atualizar a Tabela

A tabela é atualizada semanalmente pela TIM. Para atualizar o site:

1. Renomeie o novo arquivo `.xlsx` para **`tabela_precos.xlsx`**
2. Faça o upload no repositório (substituindo o anterior)
3. O GitHub Actions detecta a mudança e roda `processar_planilha.py` automaticamente
4. Em ~1 minuto, `script.js`, `index.html` e `README.md` são atualizados e o site entra no ar

> **Nota:** A data é lida do campo `DATA INÍCIO`, os planos são detectados automaticamente, e o **Boost Trade In** é lido da aba `TRADE IN` e casado por código — tudo sem edição manual.

---

## 🗂️ Estrutura do Projeto

```
Tabela-de-preco-da-TIM/
│
├── .github/
│   └── workflows/
│       └── atualizar_tabela.yml  # GitHub Actions — atualização automática
│
├── index.html              # HTML semântico (header / main / footer)
├── style.css               # Estilos organizados por seções comentadas
├── script.js               # Lógica JavaScript com comentários
├── processar_planilha.py   # Script Python de extração da planilha
├── tabela_precos.xlsx      # Planilha oficial TIM (substituir a cada semana)
├── README.md               # Este arquivo
└── LICENSE                 # Licença MIT
```

---

## 📦 Tecnologias Utilizadas

- **HTML5 semântico** — `<header>`, `<main>`, `<section>`, `<article>`, `<footer>`, `<details>`
- **CSS3** — variáveis, grid, animações, responsivo mobile-first
- **JavaScript vanilla** — sem frameworks ou dependências externas
- **Google Fonts** — Inter + DM Sans
- **Python + pandas + openpyxl** — extração automática dos dados do `.xlsx`
- **GitHub Actions** — CI/CD para publicação automática
- **GitHub Pages** — hospedagem estática gratuita

---

## 📋 Planos Suportados

<details>
<summary><strong>💰 Sem Fidelização</strong></summary>

- PRÉ (Balcão)
- Base Faturamento Retail
- Controle Não Fidelizado
- Pós Pago Não Fidelizado

</details>

<details>
<summary><strong>📱 Mobile / Celular (Fidelizado)</strong></summary>

- TIM Controle / Plus / Premium / Smart / Redes Sociais
- TIM Black / Plus / Premium
- TIM Black A (15GB) / B (20GB) / C (25GB) / C Ultra
- TIM Black Família / A One / Plus / Premium / C One / VIP

</details>

<details>
<summary><strong>🌐 Fibra / Internet</strong></summary>

- TIM Fibra 300M, 400M, 500M, 600M, 1GB, 2GB (diversas variações)

</details>



---

## ⚠️ Aviso Legal

Este projeto é de uso **interno e exclusivo para revendedores credenciados TIM**. Os preços exibidos são válidos para o canal Revendas conforme tabela oficial vigente. Não é uma ferramenta oficial da TIM S.A.

---

## 📄 Licença

Distribuído sob a licença MIT. Consulte o arquivo [LICENSE](LICENSE) para mais detalhes.

---

<div align="center">
  <sub>Mantido por revendedor credenciado TIM · Atualizado em 02/09/2026</sub>
</div>
