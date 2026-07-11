# Como ligar as inscrições (planilha + e-mail de confirmação)

Siga uma vez só. Leva ~10 minutos e é **gratuito**. No fim, cada inscrição
cai numa planilha sua e a participante recebe um e-mail de confirmação.

---

## Passo 1 — Criar a planilha
1. Acesse https://sheets.google.com e crie uma **planilha em branco**.
2. Dê um nome, ex.: **Inscrições Mulheres Curadas**.

## Passo 2 — Colar o script
1. Dentro da planilha, no menu de cima, clique em **Extensões → Apps Script**.
2. Vai abrir uma nova aba com uma caixa de código. **Apague** o que estiver lá.
3. Abra o arquivo **`google-apps-script.gs`** (deste projeto), copie **todo** o
   conteúdo e cole na caixa.
4. Clique no ícone de **salvar** (disquete) ou `Ctrl + S`.

## Passo 3 — Publicar (Deploy)
1. No canto superior direito, clique em **Implantar → Nova implantação**.
2. Clique na engrenagem ⚙️ e escolha **App da Web**.
3. Preencha:
   - **Executar como:** Eu (seu e-mail)
   - **Quem pode acessar:** **Qualquer pessoa**  ← importante!
4. Clique em **Implantar**.
5. Vai pedir para **autorizar**. Clique em **Autorizar acesso**, escolha sua
   conta Google, e em "Google não verificou este app" clique em
   **Avançado → Acessar (nome do projeto)** e **Permitir**.
   > Isso é normal — é o seu próprio script pedindo permissão para escrever na
   > sua planilha e enviar e-mails pela sua conta.
6. No fim aparece uma **URL do app da Web** (começa com `https://script.google.com/...`).
   **Copie essa URL.**

## Passo 4 — Colar a URL na página
1. Abra o arquivo **`index.html`**.
2. Lá no início do `<script>` procure a linha:
   ```js
   var URL_PLANILHA = "COLE_A_URL_AQUI";
   ```
3. Substitua `COLE_A_URL_AQUI` pela URL que você copiou. Ex.:
   ```js
   var URL_PLANILHA = "https://script.google.com/macros/s/AKfy.../exec";
   ```
4. Salve o arquivo.

## Passo 5 — Testar
1. Publique a página (veja abaixo) ou abra o `index.html`.
2. Assista o vídeo até o fim → preencha o formulário → **Quero participar**.
3. Confira:
   - a linha nova apareceu na sua **planilha**;
   - o **e-mail de confirmação** chegou (olhe também spam/promoções).

---

## Como contabilizar as inscrições
Tudo fica na planilha, uma linha por pessoa, com data/hora. Você pode:
- **Contar total:** o número da última linha já é o total.
- **Filtrar por grupo** (SGroup, Solys, Convidada): menu **Dados → Criar filtro**.
- **Ver quantos por grupo:** numa célula vazia use
  `=CONTSE(E:E; "Solys")` (troque "Solys" pelo grupo desejado).

---

## Publicar a página na internet (grátis, opcional)
A página é um site estático — dá para publicar no **GitHub Pages**:
1. No GitHub, vá em **Settings → Pages**.
2. Em **Source**, escolha a branch (`main`) e a pasta **/ (root)** → **Save**.
3. Em ~1 minuto o GitHub mostra o link público (ex.:
   `https://solysprojetos.github.io/mulherescuradas/`).

> Como o arquivo se chama `index.html`, o link abre a página direto.

---

## Dúvidas comuns
- **Mudei o script depois. E agora?** Refaça o **Passo 3** escolhendo
  **Gerenciar implantações → Editar → Nova versão → Implantar** (a URL continua a mesma).
- **O e-mail não chega?** Verifique se no Passo 3 você autorizou o acesso.
  O Google permite muitos envios por dia na conta gratuita — suficiente para o evento.
- **Quero mudar o texto do e-mail?** Edite a função `enviarConfirmacao`
  no `google-apps-script.gs` e refaça o deploy (nova versão).
