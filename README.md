# Sistema de Monitoramento de Rotas 🚗

Sistema completo para monitoramento de rotas de veículos corporativos, com **Painel do Motorista** (mobile-first) e **Painel do Supervisor** (desktop), usando Google Apps Script como backend e Google Sheets como banco de dados.

---

## ✨ Funcionalidades

### Painel do Motorista
- Formulário de pré-viagem: Nome, Placa, Modelo e Observações de Avaria
- Tela de rota ativa com data/hora de início registrada automaticamente
- Rastreamento GPS periódico (intervalo configurável pelo supervisor)
- Campo para registrar ocorrências/imprevistos durante a viagem
- Botão de finalização com data/hora registrada automaticamente
- Estado persistido em `localStorage` (motorista pode recarregar a página sem perder a rota)

### Painel do Supervisor
- **Dashboard:** lista em tempo real de todos os veículos em rota (auto-refresh a cada 30 s)
  - Nome, placa, modelo, horário de início, observações
  - Expansão de detalhes com log GPS e link para Google Maps
  - Botão de "Forçar Finalização" para encerrar rotas remotamente
- **Veículos:** liberação/cadastro de veículos na frota, lista com status
- **Configurações:** define o intervalo de registro GPS (1 – 30 minutos)
- **Relatórios:** geração e download de CSV com todas as rotas (filtrável por período e placa), incluindo log GPS completo

---

## 🗂️ Estrutura dos Arquivos

| Arquivo | Descrição |
|---|---|
| `Code.gs` | Backend Google Apps Script |
| `Driver.html` | Painel do Motorista (HTML template do Apps Script) |
| `Supervisor.html` | Painel do Supervisor (HTML template do Apps Script) |
| `appsscript.json` | Manifest do projeto Apps Script |

---

## 🚀 Instalação e Configuração

### Pré-requisitos
- Conta Google (Google Workspace ou conta pessoal)
- Acesso ao Google Sheets e Google Apps Script

### Passo a Passo

#### 1. Criar a Planilha
1. Acesse [sheets.google.com](https://sheets.google.com) e crie uma nova planilha em branco.
2. Anote a **URL** da planilha (será usada para vincular ao Apps Script).

#### 2. Abrir o Editor do Apps Script
1. Na planilha criada, vá em **Extensões → Apps Script**.
2. O editor do Apps Script será aberto.

#### 3. Configurar o Projeto
1. **Apagar** o conteúdo padrão do arquivo `Código.gs`.
2. **Copiar** o conteúdo de [`Code.gs`](./Code.gs) e colar no editor.
3. Criar um novo arquivo HTML clicando em **+** → **HTML** e nomear como `Driver` (sem extensão).
   - Copiar o conteúdo de [`Driver.html`](./Driver.html) para este arquivo.
4. Criar outro arquivo HTML chamado `Supervisor` (sem extensão).
   - Copiar o conteúdo de [`Supervisor.html`](./Supervisor.html) para este arquivo.
5. Clique em **Arquivo → Propriedades do projeto** e defina o fuso horário como **America/Sao_Paulo**.

> ⚠️ Os arquivos HTML no Apps Script **não devem ter a extensão `.html`** — apenas o nome: `Driver` e `Supervisor`.

#### 4. Publicar como Web App
1. Clique em **Implantar → Nova implantação**.
2. Selecione o tipo **Aplicativo da Web**.
3. Configure:
   - **Executar como:** Eu (conta do proprietário)
   - **Quem tem acesso:** Qualquer pessoa (ou "Qualquer pessoa dentro da organização" para maior segurança)
4. Clique em **Implantar** e copie a **URL do Web App** gerada.

#### 5. Autorizar Permissões
Na primeira execução, o Google solicitará autorização para acessar a planilha. Siga as instruções na tela para conceder as permissões necessárias.

---

## 📱 Como Usar

### Motorista
Acesse a URL do Web App (ou com `?page=driver`):
```
https://script.google.com/macros/s/SEU_ID/exec
```
ou
```
https://script.google.com/macros/s/SEU_ID/exec?page=driver
```

### Supervisor
Acesse a URL do Web App com o parâmetro `?page=supervisor`:
```
https://script.google.com/macros/s/SEU_ID/exec?page=supervisor
```

---

## 🗃️ Estrutura do Banco de Dados (Google Sheets)

O sistema cria automaticamente as seguintes abas na planilha:

| Aba | Colunas |
|---|---|
| **Rotas** | ID, Nome_Motorista, Placa, Modelo, Obs_Inicial, Hora_Inicio, Lat_Inicio, Long_Inicio, Status, Obs_Viagem, Hora_Fim, Lat_Fim, Long_Fim |
| **GPS_Log** | ID_Rota, Timestamp, Latitude, Longitude, Velocidade |
| **Veiculos** | Placa, Modelo, Status |
| **Config** | Chave, Valor |

### Status dos Veículos
- `Disponivel` — veículo livre
- `Em_Transito` — veículo em rota ativa
- `Finalizado` — rota concluída

---

## 📋 Fluxo do Sistema

```
Supervisor libera veículo
         ↓
Motorista preenche formulário (Nome, Placa, Modelo, Obs. Avaria)
         ↓
Motorista clica "Iniciar Rota"
  → Registra: Hora_Inicio, Lat/Long_Inicio, Status = Em_Transito
  → GPS começa a registrar pontos periodicamente
         ↓
Durante a rota: motorista pode adicionar observações
  → Registradas com timestamp no campo Obs_Viagem
  → Supervisor vê em tempo real no dashboard
         ↓
Motorista clica "Finalizar Rota"
  → Registra: Hora_Fim, Lat/Long_Fim, Status = Finalizado
  → Veículo volta ao status Disponivel
         ↓
Supervisor pode baixar relatório CSV com todos os dados
```

---

## 🔒 Considerações de Segurança

- Para uso **corporativo interno**, considere definir o acesso como "Qualquer pessoa dentro da organização" ao publicar o Web App.
- Os dados são armazenados em uma planilha Google Sheets vinculada à conta do proprietário.
- A geolocalização é solicitada pelo navegador e requer autorização explícita do usuário.

---

## 🛠️ Personalização

- **Intervalo GPS:** Configurável pelo supervisor no painel (de 1 a 30 minutos).
- **Fuso horário:** Definido no `appsscript.json` como `America/Sao_Paulo`. Altere conforme necessário.
- **Acesso:** Modifique o campo `"access"` em `appsscript.json` para `"ANYONE_WITH_GOOGLE"` se quiser restringir a usuários com conta Google.
