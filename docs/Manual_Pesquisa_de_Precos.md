# Pesquisa de Preços, manual de instalação e uso

## 1. Instalação

1. **Instale o Python.** Acesse python.org/downloads e baixe a versão mais recente. Durante a instalação, marque obrigatoriamente "Add Python to PATH". Sem isso o instalador não funciona.
2. **Execute o `INSTALAR.bat`.** Coloque `INSTALAR.bat`, `atualizar_precos.py` e `lojas.exemplo.json` na mesma pasta e clique duas vezes no instalador. Ele instala as bibliotecas e o navegador necessários, cria o `lojas.json` e o `EXECUTAR.bat`.
3. **Configure as lojas.** Abra `lojas.json` e ajuste o nome de cada loja como aparece no cabeçalho da planilha, a regra de preço e, se houver login, a página de login.
4. **Google Drive para Desktop (opcional).** Para sincronizar a planilha com o Drive, instale o aplicativo e mantenha a planilha dentro da pasta do Drive no computador.

## 2. Como usar

1. **Feche o navegador (Edge) completamente** se alguma loja exigir login. O programa abre o Edge com o seu perfil, e isso falha se ele já estiver aberto.
2. **Abra o programa** com `EXECUTAR.bat`.
3. **Selecione a planilha.** O programa detecta a aba com a data mais recente e mostra quantos produtos encontrou.
4. **Inicie a busca.** Se houver loja com login, o navegador abre na página de login. Autentique normalmente e clique no botão laranja "Estou logado, CONTINUAR" na janela do programa. Não feche o navegador durante a pesquisa.
5. **Acompanhe os resultados** na tabela: verde para preço atualizado, amarelo para preço igual, cinza para indisponível (salvo como R$ 0,00) e vermelho para erro (também R$ 0,00).
6. **Salve a planilha.** O programa cria uma nova aba com a data do dia (ex.: `11032026`) copiando os preços atualizados. A aba anterior é preservada.

## 3. Adicionar novos produtos

Abra a planilha no Excel, na aba mais recente, e adicione uma linha com o SKU na coluna A e o descritivo na coluna B. Nas colunas das lojas, informe o preço atual e adicione o hyperlink da página do produto (botão direito na célula, Link, cole a URL). Na próxima execução o produto entra automaticamente na pesquisa.

> ℹ️ O programa identifica os produtos pelo hyperlink da célula. Células sem hyperlink são ignoradas.

## 4. Resumo rápido

| Passo | Ação |
|---|---|
| 1 | Fechar o navegador |
| 2 | Abrir o `EXECUTAR.bat` |
| 3 | Selecionar a planilha |
| 4 | Iniciar Busca |
| 5 | Fazer login (se necessário) e clicar em CONTINUAR |
| 6 | Aguardar a pesquisa terminar |
| 7 | Salvar Planilha |
