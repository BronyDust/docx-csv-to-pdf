# Скрипт для шаблонной генерации PDF из DOCX на основе CSV для Windows

0. Установить зависимости

```
npm i
```

1. Установить scoop

```
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
Invoke-RestMethod -Uri https://get.scoop.sh | Invoke-Expression
```

2. Установить pipx

```
scoop install pipx
```

3. Обновить PATH (После этого перезайти в учетку)

```
pipx ensurepath
```

4. Установить docx2pdf

```
pipx install docx2pdf
```

5. Запустить конвертацию в pdf

```
run.bat
```
