Author: Dmitry A. Borisov, ddaabb@mail.ru (CC BY 4.0)

Необходим пакет LibreOffice, версия 7.2.4.1 (х64) или выше

Для использования (экспорта текста статьи из LibreOffice Writer 
в Habr Flavored Markdown - HFM или HTML):

Вариант 1. Достаточно скачать файл "libre_office_export.odt"
(его можно переименовать), заменить в нем текст на свой текст, 
подготовленный в соответствии с описанием в файле, и вызвать макрос. 
В папке со скачанным файлом появится новый файл в выбранном формате

Вариант 2. Можно экспортировать файлы .ODT в пакетном режиме,
запуская макросы из командной строки без GUI.
Для этого необходимо:
1. Импортировать макросы как в Варианте 3 или 4 (см. далее)
2. В LibreOffice Writer создаем и сохраняем файлы в формате .ODT
с текстами, подготовленными в соответствии с описанием в статье 
https://habr.com/ru/articles/769294/
3. Копируем файлы в формате .ODT в папку, например, "D:\odt"
4. Открываем cmd.exe и запускаем команду:
"C:\Program Files\LibreOffice\program\soffice.exe" --invisible --nofirststartwizard --headless --norestore macro:///DocExport.DocModel.ExportDir("D:\odt",0)
Комментарий:
параметр 0
macro:///DocExport.DocModel.ExportDir("D:\odt",0)
означает выбор формата для экпорта HTML
параметр 1
macro:///DocExport.DocModel.ExportDir("D:\odt",1)
или его отсутствие
macro:///DocExport.DocModel.ExportDir("D:\odt")
означает выбор формата для экпорта HFM
5. В папке "D:\odt" появятся новые файлы - результы экспорта

Вариант 3. Можно импортировать макросы из файла DocExport.oxt
Для этого необходимо:
1. Скачать файл DocExport.oxt
2. Открыть LibreOffice Writer
3. Выбираем в главном меню — Сервис — Макросы — Редактировать макросы

В открывшейся среде разработки (IDE):
1. Выбираем в главном меню — Сервис — Управление расширениями...
2. Жмем кнопку "Добавить" (внизу, в середине)
3. В открывшемся диалоге выбираем файл DocExport.oxt
4. Жмем кнопку "Открыть"

Вариант 4. Можно импортировать макросы вручную
Для этого необходимо:
1. Скачать 6 файлов:
DocModel.bas, DocView.bas, mMath.bas, ViewHfm.bas, ViewHtml.bas, vLatex.bas
2. Открыть LibreOffice Writer
3. Выбираем в главном меню — Сервис — Макросы — Редактировать макросы

В открывшейся среде разработки (IDE):
1. Выбираем в главном меню — Сервис —  Управление макросами - Basic
2. Нажимаем на кнопку "Управление..."
3. На вкладке "Библиотеки" в поле "Приложение/документ" должен быть 
выбран пункт "Мои макросы и диалоги" жмем кнопку "Создать..." (справа)
4. Вводим имя библиотеки, например, "DocExport"
5. На вкладке "Модули" в пункте "Мои макросы" раскрываем пункт "DocExport"
6. Там будет модуль "Module1", созданный автоматически, его можно удалить
7. Жмем кнопку "Создать..." (справа) вводим имя "DocModel", жмем "ОК"
8. Повторяем пункт 7 для следующих имен модулей (названия не менять!):
DocView, mMath, ViewHfm, ViewHtml, vLatex, - итого 6 модулей
9. Выходим из диалога
10. В IDE слева в панели "Каталог объектов":
- раскрываем пункт "Мои макросы и диалоги"
- раскрываем пункт "DocExport"
11. Жмем на пункт "DocModel"
12. Справа в панели редактирования кода удаляем все содержимое!
13. На панели инструментов жмем кнопку "Импорт Basic"
(лист текста со стрелкой влево) или в главном меню — Файл - Импорт Basic
14. Выбираем скачанный файл DocModel.bas
15. В панели редактирования кода отображается импортированный код
16. Повторяем пункты 10 - 14 для файлов:
DocView.bas, mMath.bas, ViewHfm.bas, ViewHtml.bas, vLatex.bas, - итого 6 файлов
17. На панели инструментов жмем кнопку "Сохранить"

Переходим в окно LibreOffice Writer:
1. Создаем и сохраняем файл в формате .ODT с текстом статьи, 
подготовленным в соответствии с описанием в статье 
https://habr.com/ru/articles/769294/
2. Для запуска макроса выбираем в главном меню — Сервис — Макросы — Выполнить макрос
3. В поле "Библиотека":
- раскрываем пункт "Мои макросы",
- раскрываем пункт "DocExport",
- выбираем пункт "DocModel",
4. В поле "Имя макроса" выбираем макрос "MakeDocHfmView"
5. Жмем кнопку "Выполнить"
6. В папке с созданным в пункте 1 файлом появится новый файл:
"имя_файла_export_hfm.txt"
7. Для экспорта в HTML в пункте 4 выбираем макрос "MakeDocHtmlView",
в папке с созданным в пункте 1 файлом появится новый файл:
"имя_файла_export.html"
