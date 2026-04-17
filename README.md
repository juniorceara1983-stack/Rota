# Sistema de Monitoramento de Rotas 🚗

Sistema completo para monitoramento de rotas de veículos corporativos, com **Painel do Motorista** (mobile-first) e **Painel do Supervisor** (desktop), usando Google Apps Script como backend e Google Sheets como banco de dados.

---

## 🔗 Links do Sistema (GitHub Pages)

| Recurso | Link |
|---|---|
| **Página Inicial** | [https://juniorceara1983-stack.github.io/Rota/](https://juniorceara1983-stack.github.io/Rota/) |
| **Painel do Motorista** | [https://juniorceara1983-stack.github.io/Rota/driver.html](https://juniorceara1983-stack.github.io/Rota/driver.html) |
| **Painel do Supervisor** | [https://juniorceara1983-stack.github.io/Rota/supervisor.html](https://juniorceara1983-stack.github.io/Rota/supervisor.html) |
| **Apps Script (API)** | [Implantação](https://script.google.com/macros/s/AKfycbzcXrHlGIxwIcUrk6ul39YJkVaVXMjfXxN8TJbTjaE6PoWQtW5VM2k7EPa82qzalBJhmg/exec) |
| **Planilha Google Sheets** | [1qu5ZkDwNnCnsFM_gKjDwbfAOGgdtb0XlOAc-_YqS7dQ](https://docs.google.com/spreadsheets/d/1qu5ZkDwNnCnsFM_gKjDwbfAOGgdtb0XlOAc-_YqS7dQ) |

---

## ✨ Funcionalidades

### Painel do Motorista
- Formulário de pré-viagem: Nome, Placa, Modelo e Observações de Avaria
- Registro obrigatório de quilometragem inicial e final da rota
- Tela de rota ativa com data/hora de início registrada automaticamente
- Rastreamento GPS periódico (intervalo configurável pelo supervisor)
- Registro imediato de posição ao voltar para o app após bloqueio/desbloqueio da tela (via `visibilitychange`, `focus` e `pageshow`)
- Velocidade registrada em **km/h** (convertida do valor m/s retornado pela API de Geolocalização)
- Campo para registrar ocorrências/imprevistos durante a viagem
- Botão de finalização com data/hora registrada automaticamente
- Tela dedicada de **Registro de Abastecimento** (placa, KM, litros, valor e data/hora) acessível a partir do cadastro inicial ou durante uma rota ativa
- Estado persistido em `localStorage` (motorista pode recarregar a página sem perder a rota)

### Painel do Supervisor
- **Dashboard:** lista em tempo real de todos os veículos em rota (auto-refresh a cada 30 s)
  - Nome, placa, modelo, horário de início, observações
  - Expansão de detalhes com log GPS e link para Google Maps
  - Botão de "Forçar Finalização" para encerrar rotas remotamente
- **Veículos:** liberação/cadastro de veículos na frota, lista com status e botão **Remover** (veículos em rota ativa são protegidos contra remoção)
- **Configurações:** define o intervalo de registro GPS (1 – 30 minutos) e a **retenção de dados** (7, 15, 30, 60, 90, 180 ou 365 dias) com botão "Limpar agora" e função `triggerLimpezaDiaria` para agendamento automático
- **Rastrear:** busca por placa com última posição + lista de localizações e links diretos para Google Maps
- **Abastecimento:** aba apenas de **consulta** (lista e filtro por placa) com botão **Baixar PDF**. O registro é feito pelo motorista no painel do motorista
- **Relatórios:** geração e download de CSV com todas as rotas (filtrável por período e placa), incluindo log GPS completo

---

## 🗂️ Estrutura dos Arquivos

| Arquivo | Descrição |
|---|---|
| `index.html` | Página inicial (GitHub Pages) – links para os dois painéis |
| `driver.html` | Painel do Motorista (GitHub Pages) |
| `supervisor.html` | Painel do Supervisor (GitHub Pages) |
| `Code.gs` | Backend Google Apps Script (API + banco de dados) |
| `Driver.html` | Template original para Apps Script (legado) |
| `Supervisor.html` | Template original para Apps Script (legado) |
| `appsscript.json` | Manifest do projeto Apps Script |

---

## 🚀 Instalação e Configuração

### Pré-requisitos
- Conta Google (Google Workspace ou conta pessoal)
- Acesso ao Google Sheets e Google Apps Script
- Repositório no GitHub com GitHub Pages ativado

---

### Passo a Passo

#### 1. Ativar GitHub Pages
1. No repositório, vá em **Settings → Pages**.
2. Em **Source**, selecione o branch `main` e a pasta `/ (root)`.
3. Clique em **Save**. O site ficará disponível em `https://juniorceara1983-stack.github.io/Rota/`.

#### 2. Criar a Planilha
1. Acesse [sheets.google.com](https://sheets.google.com) e crie uma nova planilha em branco.

#### 3. Criar o Apps Script (API)
1. Na planilha criada, vá em **Extensões → Apps Script**.
2. **Apague** o conteúdo padrão do arquivo `Código.gs`.
3. **Copie** o conteúdo de [`Code.gs`](./Code.gs) e cole no editor.
4. Salve o projeto.

#### 4. Publicar o Apps Script como Web App
1. Clique em **Implantar → Nova implantação**.
2. Selecione o tipo **Aplicativo da Web**.
3. Configure:
   - **Executar como:** Eu (conta do proprietário)
   - **Quem tem acesso:** Qualquer pessoa
4. Clique em **Implantar** e copie a **URL do Web App** gerada.

#### 5. Configurar a URL nos arquivos HTML
1. Em `driver.html`, localize a linha:
   ```js
   var SCRIPT_URL = 'COLE_AQUI_A_URL_DO_SEU_APPS_SCRIPT';
   ```
   Substitua pelo URL copiado no passo anterior.
2. Faça o mesmo em `supervisor.html`.
3. Faça commit e push das alterações para o GitHub.

#### 6. Autorizar Permissões
Na primeira execução do Apps Script, o Google solicitará autorização para acessar a planilha. Siga as instruções na tela.

---

## 📱 Como Usar

### Motorista
Acesse:
```
https://juniorceara1983-stack.github.io/Rota/driver.html
```

### Supervisor
Acesse:
```
https://juniorceara1983-stack.github.io/Rota/supervisor.html
```

---

## 🗃️ Estrutura do Banco de Dados (Google Sheets)

O sistema cria automaticamente as seguintes abas na planilha:

| Aba | Colunas |
|---|---|
| **Rotas** | ID, Nome_Motorista, Placa, Modelo, Obs_Inicial, Hora_Inicio, Lat_Inicio, Long_Inicio, Status, Obs_Viagem, Hora_Fim, Lat_Fim, Long_Fim, Km_Inicio, Km_Fim |
| **GPS_Log** | ID_Rota, Timestamp, Latitude, Longitude, Velocidade |
| **Veiculos** | Placa, Modelo, Status |
| **Config** | Chave, Valor |
| **Abastecimentos** | Placa, Quilometragem, Litros, Valor, Data_Hora |

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
