# Empacotar o app em `.exe` (Windows)

## Pré-requisitos
- Python 3.10+ instalado no Windows e disponível no PATH.

## Como gerar
1. Abra o `Prompt de Comando` na pasta do projeto.
2. Execute:

```bat
build_windows.bat
```

## Saída esperada
- Executável em:

```text
dist\BuscaNFSe\BuscaNFSe.exe
```

## Arquivos de build adicionados
- `build_windows.bat`: automação completa (venv, dependências, ícone, pyinstaller).
- `app_nfse_lote_excel.spec`: configuração do PyInstaller.
- `scripts/generate_icon.py`: gera ícone `.ico` automaticamente em `assets/busca_nfse.ico` (arquivo gerado em build, não versionado).


## Observações de desempenho (Windows Server)
- O build está com `UPX` desativado no `.spec` para evitar lentidão/travamentos em alguns ambientes Windows Server/antivírus.
- O processamento do lote foi movido para thread de background para manter a interface responsiva durante execução.
