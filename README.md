# LootZ

Um balcão de inventário para parties de RPG: importe uma planilha, filtre os itens, decida o que vender e acompanhe a divisão do ouro em tempo real.

## Publicar no GitHub Pages

1. Crie um repositório no GitHub e envie todos os arquivos desta pasta.
2. No repositório, abra **Settings → Pages**.
3. Em **Build and deployment**, escolha **Deploy from a branch**.
4. Selecione a branch `main`, a pasta `/ (root)` e clique em **Save**.

O site é totalmente estático e não precisa de servidor ou etapa de build.

## Planilha

O inventário começa vazio. O botão **Carregar exemplo** abre o arquivo `inventario_venda_rpg_ordenado.xlsx` incluído no projeto, enquanto **Importar planilha** aceita outro `.xlsx`, `.xls` ou `.csv`. Para melhor compatibilidade, use cabeçalhos equivalentes a:

`Origem`, `Item`, `Qtd.`, `Valor real (gp)`, `Valor estimado (gp)`, `Categoria`, `Vender?`, `Valor base (gp)`, `Bônus aplicado` e `Observação`.

As alterações ficam salvas apenas no navegador do usuário. A planilha importada nunca é enviada a um servidor.
