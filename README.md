# XMLParse

[![Button Icon]][Link] 

## Как пользоваться программой

- Запустите программу. Она создаст 2 папки (source_file и processed_file)
- Поместите файл с расширением '.xml' в папку source_file
- Запустите программу ещё раз
- Дождитесь окончания обработки файла (это может занять несколько минут)
- Откройте папку processed_file, в которой будет находится выгрузка из xml (файл Выгрузка.csv) и обработанный файл (Свод по операторам.xlsx)

## Файлы программ
- XMLParse.py - программа для выгрузки данных из xml в формате csv
- SvodOperators.py - программа для обработки выгруженных данных в формате csv и сохранение в формате xlsx
- AvtoSborka.py - программа для автоматической сборки 
- MergeFiles.py - программа для вставки id из выгрузки из другой системы в выгрузку (не используется)
- FiasToExcel.py - программа для вставки кода ФИАС в файл выгрузки из другой системы (не используется)

## Сборка программы
Для сборки исполняемого файла запустите командную строку и введите следующую команду:

```sh
pyinstaller --onefile XMLParse.py SvodOperators.py
```

Файлы программ указываются в порядке их выполнения
*Вы также можете сделать сборку с помощью программы AvtoSborka.py на вашем локальном компьютере*

[Button Icon]: https://img.shields.io/badge/Installation-EF2D5E?style=for-the-badge&logoColor=white&logo=DocuSign
[Link]: https://disk.yandex.ru/d/y3RRJiChtpWCww