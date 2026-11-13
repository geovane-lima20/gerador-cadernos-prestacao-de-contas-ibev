# IBEV | Gerador de Cadernos de Comprovantes

Ferramenta simples para gerar um arquivo Word (`.docx`) com as **saídas mensais** da planilha de movimentação financeira da IBEV.

**Não precisa instalar nada, criar conta, rodar Python nem contratar um servidor.** O processamento do Excel e a criação do Word ocorrem no seu navegador. O site pode ser hospedado pelo GitHub Pages.

## Como usar

1. Abra o endereço da ferramenta (ou o arquivo `index.html`).
2. Selecione o arquivo `MOVIMENTO 2026.xlsx` no seu computador. Se estiver no OneDrive, obtenha a versão `.xlsx` atualizada antes.
3. Selecione o **mês**. A ferramenta reconhece as 12 abas, de janeiro a dezembro.
4. Marque as **categorias** que deseja incluir. Todas começam selecionadas.
5. Opcionalmente, marque **Adicionar uma folha de resumo geral ao final**.
6. Confira os valores e os avisos. Se houver divergência, o sistema pede uma confirmação explícita.
7. Clique em **Baixar caderno em Word (.docx)**.
8. Abra o Word e imprima **A4, paisagem**, sem escala adicional, preferencialmente conferindo a prévia de impressão.

## Publicar no GitHub Pages (sem linha de comando)

1. Crie um repositório no GitHub, por exemplo `ibev-caderno`.
2. Envie para a **raiz** do repositório os arquivos `index.html`, `styles.css`, `app.js` e `.nojekyll` (o `README.md` também pode ser enviado).
3. Entre em **Settings > Pages**.
4. Em **Build and deployment**, escolha **Deploy from a branch**.
5. Selecione a branch **main**, a pasta **/(root)** e clique em **Save**.
6. O endereço ficará no formato `https://SEU-USUARIO.github.io/ibev-caderno/` (conforme a configuração do seu repositório).

Fonte: [Documentação oficial GitHub Pages](https://docs.github.com/en/pages/getting-started-with-github-pages/configuring-a-publishing-source-for-your-github-pages-site).

> **Atenção:** GitHub Pages publica o site na internet. Não envie a planilha financeira, PDFs de comprovantes, arquivos Word gerados ou senhas para o repositório. O ZIP fornecido contém somente o código do site; a planilha de teste não está incluída.

## Regras reproduzidas do programa original

- Lê os blocos de **resumo das saídas** de cada aba mensal.
- Gera as categorias na ordem: **FOLHA DE PAGAMENTO / PESSOAL**, **ADMINISTRATIVO**, **PATRIMONIO**, **REDES**, **DESPESA ALIMENTAÇÃO**, **DESPESAS BANCÁRIAS**.
- Oculta os lançamentos de valor zero, preserva as descrições e calcula a linha **TOTAL**.
- Compara a soma de cada categoria ao subtotal registrado na planilha e sinaliza divergências.
- Reconhece o formato antigo (subtotais na coluna C) e o novo (linha explícita TOTAL).
- Reconhece a seção de despesas de janeiro de 2026, mesmo que ela esteja identificada como `RESUMO ENTRADA` na planilha.
- Considera registros cuja descrição coincida com o nome da categoria, desde que tenham um valor. Isso corrige uma omissão do código antigo.
- Cria folhas de **A4 na horizontal**, com a parte esquerda livre para colar/fixar o comprovante, linha de dobra pontilhada, título, tabela e subtotal.
- Quando uma categoria tem muitos itens, pode usar páginas de continuação para manter as linhas legíveis.

## Informações sobre os dados

- **Privacidade:** não há serviços externos, autenticação, banco de dados, envio de arquivos a terceiros ou bibliotecas baixadas de uma CDN. O arquivo fica apenas na memória do navegador enquanto a ferramenta está em uso.
- **Formato aceito:** `.xlsx` tradicional do Microsoft Excel, não protegido por senha. Não suporta `.xls` antigo ou `.xlsm`.
- **Conferência:** os números são lidos do valor armazenado no arquivo Excel, inclusive valores calculados de fórmulas já salvos. Por segurança, atualize e salve sua planilha antes de selecioná-la.
- **Compatibilidade:** recomendados Chrome, Edge ou Firefox atualizados, por utilizarem a API nativa de descompactação de arquivos ZIP do navegador.
- **Sem garantia contábil:** o sistema confere somatórios do resumo fornecido, mas não substitui a conciliação com os comprovantes e os lançamentos originais.

## Arquivos do projeto

- `index.html`: página e componentes da interface.
- `styles.css`: estilo visual adaptável a computador e celular.
- `app.js`: leitura do Excel, conferências, prévia e geração do Word.
- `.nojekyll`: configuração auxiliar do GitHub Pages.

## Alterar as categorias no futuro

A lista de seis categorias fica no início de `app.js`, em `CATEGORIES`. Para usar a ferramenta com uma planilha que tenha categorias diferentes, será necessário adaptar os nomes e possivelmente as regras de extração.

## Testes realizados

A versão 1.0 foi testada com o arquivo **MOVIMENTO 2026.xlsx** fornecido pela IBEV, inclusive os formatos de janeiro e de agosto. O Word gerado para agosto de 2026 foi aberto/renderizado, produzindo seis folhas A4 em paisagem. Os totais das seis categorias de agosto conferiram com a planilha, totalizando **R$ 12.176,30**.
