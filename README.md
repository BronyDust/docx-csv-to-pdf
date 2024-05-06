# Скрипт для шаблонной генерации PDF из DOCX на основе CSV для Windows

1. Запустить скрипт, получится папка docx

```
node ./script.js
```

2. Установить scoop

```
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
Invoke-RestMethod -Uri https://get.scoop.sh | Invoke-Expression
```

3. Установить pipx

```
scoop install pipx
```

4. Обновить PATH (После этого перезайти в учетку)

```
pipx ensurepath
```

5. Установить docx2pdf

```
pipx install docx2pdf
```

6. Запустить конвертацию в pdf

```
docx2pdf docx/ result/
```

## ИЛИ

1. Так же
2. Прогнать файлы через https://tools.pdf24.org/en/docx-to-pdf
