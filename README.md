# XMLParse

## Как пользоваться программой

- Запустите программу. Она создаст 2 папки
- Поместите файл с расширением '.xml' в папку source_file
- Запустите программу, нажав на Xml Parser.exe
- Дождитесь окончания обработки файла (может занять несколько минут)
- Откройте папку processed_file, в которой будет находится выгрузка из xml (файл itog.csv)

## Файлы программ
- XMLParse.py - программа для выгрузки данных из xml в формате csv
- SvodOperators.py - программа для обработки выгруженных данных в формате csv и сохранение в формате xlsx
- MergeFiles.py - программа для вставки id из выгрузки из другой системы в выгрузку (не используется)
- FiasToExcel.py - программа для вставки кода ФИАС в файл выгрузки из другой системы (не используется)

## Сборка программы
Для сборки исполняемого файла .exe запустите командную строку и введите следующую команду:

```sh
pyinstaller --onefile XMLParse.py SvodOperators.py
```

Файлы программ указываются в порядке их выполнения