# RPA Foto Fazenda

Aplicacao desktop para organizar fotos por fazenda/lote/data e registrar brincos em planilhas Excel.

## Fluxo rapido
1. Coloque imagens em `processar/`.
2. Abra o app, selecione a data, a fazenda e o lote.
3. Digite o numero do brinco para cada foto.
4. O resultado vai para `Dados Processados/<Fazenda>/<Lote>/<DD-MM-YYYY>/`:
   - planilha `DD-MM-YYYY.xlsx` com a coluna `BRINCOS`
   - fotos renomeadas com o numero do brinco

## Estrutura do projeto
- `main.py`: ponto de entrada.
- `gui.py`: interface (CustomTkinter).
- `logic.py`: regras de negocio.
- `processar/`: pasta de entrada das fotos (criada automaticamente se nao existir).
- `Dados Processados/`: saida organizada pelo app.

## Requisitos
- Python 3.x
- Dependencias: `customtkinter`, `tkcalendar`, `pillow`, `openpyxl`

## Criar venv e instalar dependencias (Windows / PowerShell)
```ps1
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install customtkinter tkcalendar pillow openpyxl
```

## Executar
```ps1
python main.py
```

## Gerar executavel com PyInstaller
```ps1
pip install pyinstaller
pyinstaller --onefile --windowed --name RPA_Foto_Fazenda main.py
```
O executavel fica em `dist/` (pasta gerada pelo PyInstaller).

## Informacoes cruciais
- As fotos sao movidas (nao copiadas) para a pasta de destino. Se quiser manter o original, faca backup antes.
- O app evita duplicidade: se o brinco ja estiver na planilha do dia, a foto nao e movida.
- "Desfazer" reverte apenas a ultima foto processada na sessao atual.
- A busca percorre todas as pastas em `Dados Processados/` e sugere numeros com 1 digito de diferenca.
- Codigos de fazenda usados na interface:
  - `VAL` -> Valongo
  - `BARRA` -> Barra
  - `TRES` -> Tres Divisas
