# 🔍 Pesquisa de Preços da Concorrência

> Aplicativo de desktop em Python que automatiza a pesquisa de preços de concorrentes. Ele lê uma planilha em que cada produto tem o link da página em cada loja, visita as páginas com o Playwright, extrai o preço, marca produtos indisponíveis e grava o resultado em uma nova aba datada, preservando o histórico das pesquisas anteriores.

[![Python](https://img.shields.io/badge/Python-3776AB?style=flat&logo=python&logoColor=white)](https://www.python.org)
[![Playwright](https://img.shields.io/badge/Playwright-2EAD33?style=flat)](https://playwright.dev/python/)
[![openpyxl](https://img.shields.io/badge/openpyxl-217346?style=flat)](https://openpyxl.readthedocs.io)
[![tkinter](https://img.shields.io/badge/tkinter-GUI-4B8BBE?style=flat)](https://docs.python.org/3/library/tkinter.html)

> ⚖️ **Uso do código.** Este projeto foi desenvolvido no contexto profissional do autor e é publicado, em versão sanitizada e com dados fictícios, exclusivamente para fins de portfólio. Todos os direitos reservados. Não é permitido o uso comercial sem autorização.

---

## 🎯 O problema

A pesquisa de preços era feita à mão: abrir o site de cada concorrente, procurar o produto, anotar o preço na planilha. Com dezenas de produtos e quatro lojas, o processo levava horas e era feito com pouca frequência. A automação reduziu a pesquisa a alguns cliques e passou a guardar o histórico de cada rodada.

## ⚙️ Como funciona

1. **A planilha é a fonte de verdade.** Coluna A tem o SKU, coluna B o descritivo e as colunas seguintes, uma por loja, têm o preço atual com um hyperlink para a página do produto. Para incluir um produto novo basta adicionar a linha com os links, nenhuma alteração no código.
2. **As lojas ficam em `lojas.json`.** Cada loja define como é reconhecida no cabeçalho, qual regra de extração de preço usar e se exige login.
3. **O programa detecta a aba mais recente** (nome no formato `DDMMAAAA`), visita cada link em um navegador headless e extrai o preço do texto visível da página com expressão regular para valores em reais.
4. **Produtos indisponíveis** são identificados pela ausência do botão de compra (configurável por loja) ou por avisos como "fora de estoque", e gravados como R$ 0,00.
5. **Lojas com login** usam o navegador do próprio usuário (perfil persistente do Edge). O programa abre a página de login, o usuário autentica e confirma na interface. Nenhuma credencial é armazenada.
6. **Ao salvar**, o programa cria uma nova aba com a data do dia copiando a anterior, inclusive a formatação condicional, e grava os preços novos. A aba anterior é preservada, o que permite comparar rodadas.

A interface mostra o andamento em tempo real: verde para preço atualizado, amarelo para preço igual, cinza para indisponível e vermelho para erro.

## ✨ O que o projeto demonstra

- Automação de navegador com Playwright (headless e perfil persistente do usuário)
- Extração de dados de páginas heterogêneas com regras configuráveis por loja
- Leitura de hyperlinks e escrita em Excel com openpyxl, com cópia de abas e formatação condicional
- Interface gráfica em tkinter com thread de trabalho, tabela de resultados e controle de parada
- Configuração separada do código (`lojas.json`), permitindo trocar as lojas sem programar
- Documentação de instalação e uso para a equipe de compras (`docs/`)

## 🚀 Como executar

> ⚠️ **Os arquivos deste repositório são exemplos.** As lojas em `lojas.exemplo.json` e os links da planilha de exemplo são fictícios e não abrem páginas reais. Para usar de verdade, é preciso criar o `lojas.json` com as lojas pesquisadas (cabeçalho, regra de preço e login, se houver) e usar a sua própria planilha, com os hyperlinks dos produtos. O programa não funciona só com os arquivos baixados.

Requisitos: Windows com Python 3.10+ e Microsoft Edge (apenas para lojas com login).

1. Execute `INSTALAR.bat` (instala `openpyxl` e `playwright`, baixa o Chromium e cria o `EXECUTAR.bat`).
2. Edite `lojas.json` com as lojas pesquisadas (o instalador cria o arquivo a partir de `lojas.exemplo.json`).
3. Abra o programa com `EXECUTAR.bat`, selecione a planilha e clique em Iniciar Busca.

Para testar o fluxo, use `exemplos/pesquisa_precos_exemplo.xlsx` (os links apontam para domínios de exemplo, então os produtos aparecerão como erro, o que é esperado).

> ℹ️ Respeite os termos de uso dos sites pesquisados. O programa acessa apenas páginas públicas de produto ou páginas em que o próprio usuário fez login.

## 📁 Estrutura

```
├── atualizar_precos.py     Aplicativo principal
├── lojas.exemplo.json      Modelo de configuração das lojas
├── INSTALAR.bat            Instala dependências e cria o EXECUTAR.bat
├── docs/                   Manual de instalação e uso
└── exemplos/               Planilha de exemplo com dados fictícios
```

---

## 👤 Autor

**Heberth Augusto**

[![LinkedIn](https://img.shields.io/badge/LinkedIn-0077B5?style=flat)](https://www.linkedin.com/in/heberth-augusto-ti/)
[![GitHub](https://img.shields.io/badge/GitHub-100000?style=flat&logo=github&logoColor=white)](https://github.com/heberthaugusto)

*Desenvolvido com IA generativa (Claude) como ferramenta de trabalho, com definição das regras, revisão do código e validação em uso real pelo autor.*
