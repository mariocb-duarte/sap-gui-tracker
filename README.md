# SAP GUI Tracker

Ferramenta leve para explorar telas do **SAP GUI for Windows** e gerar código de automação via **SAP GUI Scripting**, inspirada no [Scripting Tracker](https://tracker.stschnell.de/) de Stefan Schnell.

Não precisa instalar nada: usa apenas HTA + VBScript, que já vêm no Windows.

## Arquivos

| Arquivo | O que faz |
|---|---|
| `Tracker.hta` | Analisador: mostra a árvore de elementos da tela SAP, destaca o elemento no SAP e gera código VBScript e Python |
| `Gravador.vbs` | Gravador: registra as ações feitas no SAP e salva scripts `.vbs` e `.py` em `gravacoes/` |

## Requisitos

- Windows com SAP GUI for Windows
- Scripting habilitado **no servidor** (transação RZ11, parâmetro `sapgui/user_scripting = TRUE`, feito pela equipe Basis)
- Scripting habilitado **no cliente**: SAP GUI → Opções → Acessibilidade e Scripting → Scripting → *Ativar scripting*

## Analisador (`Tracker.hta`)

1. Faça login no SAP e abra a transação desejada.
2. Dê dois cliques em `Tracker.hta`.
3. Escolha a sessão e clique em **Carregar árvore**.
4. Clique em um elemento: ele fica destacado em vermelho no SAP, e o painel mostra ID, tipo, propriedades e código pronto para copiar.

Recursos:

- Árvore de elementos com etiquetas coloridas por tipo, recolher/expandir e filtro por ID, nome ou texto
- Carregamento em blocos com barra de progresso e opção de cancelar
- Grids (ALV / `GridView`): número de linhas e nomes das colunas
- Árvores (`GuiTree`): lista os nós com chave e texto, e permite selecionar, expandir, recolher e dar duplo clique
- Código gerado em VBScript e Python (`win32com`)

A janela não tem barra de título. Para movê-la, arraste a faixa azul do topo. Para redimensionar, use as bordas.

## Gravador (`Gravador.vbs`)

1. Deixe o SAP na tela inicial do fluxo.
2. Dê dois cliques em `Gravador.vbs`. Aparece a caixa "Gravando…".
3. Execute as ações no SAP.
4. Clique em **OK** na caixa para parar.

Os scripts ficam em `gravacoes/gravacao_AAAAMMDD_HHMMSS.vbs` e `.py`. A pasta está no `.gitignore` porque as gravações podem conter dados de negócio.

## Usando o código gerado em Python

```python
import win32com.client  # pip install pywin32

sap = win32com.client.GetObject("SAPGUI")
app = sap.GetScriptingEngine
session = app.ActiveSession

session.findById("wnd[0]/tbar[0]/okcd").text = "/nME23N"
session.findById("wnd[0]").sendVKey(0)
```

## Observações

- `Gravador.vbs` está em codificação Windows-1252 (ANSI), exigida pelo Windows Script Host para exibir acentos. `Tracker.hta` está em UTF-8 com BOM.
- Verifique a política de TI da sua empresa antes de usar automação no SAP.
