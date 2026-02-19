## Automação de Cartões – Build do Executável (Windows 11)

Este projeto transforma o script original `robo_cartoes_emsys_v3.py` em um aplicativo Windows com interface gráfica (Tkinter + ttk/ttkbootstrap), sem console, pronto para uso por qualquer colaborador.

### 1. Preparar ambiente Python

1. Instale o Python 3.10+ no Windows (de preferência marcando a opção **“Add Python to PATH”** durante a instalação).
2. Abra o **Prompt de Comando** ou **PowerShell** na pasta do projeto (onde está `app.py`).

### 2. Instalar dependências

```bash
pip install -r requirements.txt
python -m playwright install
```

O segundo comando baixa os navegadores necessários para o Playwright (usado na captura do Good Card via CDP).

### 3. Arquivo de ícone (icon.ico / icon.png)

- O ícone do aplicativo deve ficar **na raiz do projeto**, ao lado de `app.py`:
  - `icon.ico` – usado pelo executável do Windows.
  - `icon.png` – usado para o splash screen e como referência visual.
- Se você quiser trocar o ícone no futuro:
  1. Crie/obtenha um novo ícone **512x512** em PNG com estilo moderno (ex.: azul escuro + verde, tema cartões/robô/automação).
  2. Converta esse PNG para `.ico` (por exemplo, usando um site conversor ou o próprio Pillow em um script separado).
  3. Salve esses arquivos substituindo os existentes:
     - `icon.png`
     - `icon.ico`

O código da GUI já tenta usar automaticamente `icon.ico` na janela principal (`root.iconbitmap("icon.ico")`) e `icon.png` no splash screen.

### 4. Gerar o executável com PyInstaller

No mesmo terminal/prompt, ainda dentro da pasta do projeto, execute:

```bash
pyinstaller --noconsole --onefile --icon=icon.ico --name AutomacaoCartoes app.py
```

Explicação dos parâmetros:

- `--noconsole` → o executável não abre console de fundo (app totalmente gráfico).
- `--onefile` → gera um único `.exe` para distribuição.
- `--icon=icon.ico` → usa o ícone personalizado do aplicativo.
- `--name AutomacaoCartoes` → nome do executável gerado (`AutomacaoCartoes.exe`).
- `app.py` → ponto de entrada da aplicação (GUI principal).

Após a execução bem-sucedida, o arquivo `AutomacaoCartoes.exe` ficará dentro da pasta `dist/`.

### 5. Pastas e arquivos gerados em tempo de execução

Ao rodar o executável (ou `app.py` diretamente), o app:

- Garante que o diretório de trabalho aponte para a pasta onde está o `.exe`/`.py`.
- Usa caminhos relativos para criar/ler arquivos **ao lado do executável**, por exemplo:
  - `capturas_portal/` – pasta onde ficam os arquivos de captura (`captura_001.txt`, etc.).
  - `config_emsys_grid.json` – configuração de calibração do grid do EMSYS.
  - `valecard_despesas.json` – resumo de despesas do Vale Card (último PDF processado).
  - `encontrados.txt`, `nao_encontrados.txt`, `resumo.txt` – relatórios gerados após rodar o EMSYS.
  - `capturas_unificadas.csv` – exportação opcional das capturas unificadas.

Não é necessário criar essas pastas/arquivos manualmente; o aplicativo cria tudo automaticamente quando necessário.

### 6. Comando de build a ser usado sempre

Sempre que quiser gerar/regerar o executável com ícone, use:

```bash
pyinstaller --noconsole --onefile --icon=icon.ico --name AutomacaoCartoes app.py
```

Se alterar o ícone, apenas substitua `icon.ico` (e opcionalmente `icon.png`) antes de rodar o comando acima novamente.

