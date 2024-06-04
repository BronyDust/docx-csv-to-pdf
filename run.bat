@echo OFF
chcp 65001
echo Сейчас будем генерировать PDF в две фазы. Для начала укажи файл шаблонов и данных. Перетащи мышкой в это поле и нажми ENTER
set /P templatePath=Файл шаблона (DOCX): %=%
set /P dataPath=Файл данных (CSV): %=%
node ./script.js %dataPath% %templatePath%
echo ЭТАП ПРЕОБРАЗОВАНИЯ
docx2pdf docx/ result/
echo ПРЕОБРАЗОВАНИЕ ЗАВЕРШЕНО
echo Все готово. Забирай файлы PDF из папки ./result, ибо при следующем запуске скрипта эта папка будет дропнута
pause
