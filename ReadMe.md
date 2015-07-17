﻿## Известные проблемы при работе с репозиторием

### Проблемы чтения кодировки windows-1251 в репозитории

1. При просмотре в [репозитории] по умолчанию файлы отображаются в кодировке `UTF-8`.
2. При редактировании в [репозитории] файлы пересохраняются в кодировке `UTF-8`.
3. Отредактированные файлы в папке `/src/*.*` после [скачивания в ZIP] необходимо открыть в **Блокноте** и пересохранить как `ANSI`. При импорте модули должны быть в кодировке `windows-1251`.

[репозитории]://github.com/bopoh13/StatsOKM/tree/dev/src
[скачивания в ZIP]://github.com/bopoh13/StatsOKM/archive/dev.zip

### Автоматическая кодировка файлов через фильтр Git

1. В корне (клона) репозитория необходимо создать файл `.gitattributes`.
2. Добавить в файл макросы и комментарий
	```markdown
	# Custom for Visual Basic (CRLF for classes or modules)
	*.bas	filter=win1251  eol=crlf
	*.cls	filter=win1251  eol=crlf
	```

3. Выполнить в Git Bash комманды
	```bash
	$ git config --global filter.win1251.clean "iconv -f windows-1251 -t utf-8"
	$ git config --global filter.win1251.smudge "iconv -f utf-8 -t windows-1251"
	$ git config --global filter.win1251.required true
	```

4. Теперь можно работать с файлами через Git Bash или Git Client не заботясь о кодировке.

### Возможные проблемы при использовании фильтра в Git Client

После обновления Git Client при выполнении любой команды, например `git status`, может возникнуть ошибка **fatal: [*file*] clean filter 'win1251' faile**. Для устранения необходимо удалить фильтры, установить фильтры и произвести повторное клонирование репозитория.
