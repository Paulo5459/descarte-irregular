# Estudo de Caso – Descarte Irregular • GCM APGO

Sistema de registro das visitas de campo da Inspetoria de Inteligência (05/10/2026 a 05/11/2026).

- **Aplicativo de campo** (`src/Index.html`): HTML, CSS e JS para celular. O agente registra as visitas com GPS, fotos com carimbo e envio para a planilha.
- **Script da planilha** (`src/Code.gs`): monta as abas, recebe as visitas, salva as fotos no Drive e exporta a planilha em Excel.
- **Configuração** (`src/appsscript.json`): fuso de São Paulo e App da Web executando como o dono, com acesso por qualquer pessoa que tenha o link e o PIN.

## Estrutura

```
descarte-irregular/
├── src/
│   ├── Code.gs            ← servidor (Google Apps Script)
│   ├── Index.html         ← aplicativo de campo
│   └── appsscript.json    ← manifesto do projeto
├── docs/                  ← documento do estudo de caso
├── .clasp.json.example    ← modelo de ligação com o Apps Script
├── package.json           ← atalhos do clasp
└── README.md
```

## Planilha no Drive

Pasta: **Estudo Descarte Irregular – GCM APGO**
Planilha: **Planilha – Estudo Descarte Irregular – GCM APGO**
https://docs.google.com/spreadsheets/d/1FkX_A4Z__UXCO6KyGp4Zwc3IzdCE1JtNzhSc39LuCGA/edit

## Opção A – Conectar pelo VS Code (clasp)

1. Instale o Node.js e, no terminal do VS Code, dentro desta pasta, rode: `npm install`
2. Ative a API do Apps Script uma única vez: https://script.google.com/home/usersettings → **Google Apps Script API: Ativada**
3. Faça login com a sua conta Google: `npx clasp login`
4. Abra a planilha → **Extensões → Apps Script** → **Configurações do projeto (engrenagem)** → copie o **ID do script**.
5. Copie `.clasp.json.example` para `.clasp.json` e cole o ID no lugar de `COLE_AQUI_O_ID_DO_SCRIPT`.
6. Envie os códigos: `npx clasp push` (confirme com **y** se perguntar sobre sobrescrever o manifesto).
7. No editor do Apps Script, selecione a função `configurarPlanilha` → **Executar** → autorize.
8. Implante: no editor, **Implantar → Nova implantação → App da Web**, ou use `npm run implantar`.
9. Copie a URL que termina em **/exec** e envie às equipes.

Depois de editar algo no VS Code: `npx clasp push`. Depois, no Apps Script, vá em **Implantar → Gerenciar implantações → editar → Nova versão**. Assim o link das equipes continua o mesmo.

## Opção B – Manual (copiar e colar)

1. Abra a planilha → **Extensões → Apps Script**.
2. Cole o conteúdo de `src/Code.gs` no arquivo `Código.gs`.
3. Crie o arquivo HTML **Index** (exatamente esse nome) e cole `src/Index.html`.
4. Execute `configurarPlanilha` e autorize.
5. **Implantar → Nova implantação → App da Web**. Em "Executar como", escolha **Eu**. Em "Quem pode acessar", escolha **Qualquer pessoa**.

## Acesso ao aplicativo

- **Agentes / matrícula:** texto livre.
- **PIN:** o padrão é `2026`. Troque antes de 05/10 pelo menu **Estudo Descarte → Alterar PIN de acesso do aplicativo**.
- O aplicativo só funciona pela URL **/exec**. Se você abrir o `Index.html` direto do computador, aparece o aviso de que ele foi aberto fora do Google.

## Abas da planilha

| Aba | Conteúdo |
|---|---|
| Indicadores | Andamento, tendências, área, limpeza, apreensões, visitas por equipe e por semana, ranking de câmeras |
| Consolidação | Volume S1–S4 por ponto, limpeza, apreensão e tendência automática |
| Visitas | Uma linha por visita enviada pelo aplicativo, com link das fotos |
| Escala | Escala diária: data, dia, equipe (Charlie → Delta → Alfa → Bravo), pontos e visitas realizadas |
| Pontos | Os 41 pontos do documento da SDU com a regional e o dia da visita |

## Onde ajustar

- Datas, equipes, dias e grupos de pontos: objeto `CFG` no início de `src/Code.gs`.
- Lista dos 41 pontos e regionais: constante `PONTOS` em `src/Code.gs`.
- Campos da ficha e visual: `src/Index.html`.
