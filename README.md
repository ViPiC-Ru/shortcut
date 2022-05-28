# Описание
`JScript` для **создания**, **изменения** и **удаления** ярлычка. Основным преимуществом скрипта является то, что ярлычок не пересоздаётся при каждом запуске, а проверяются его отдельные свойства и изменение происходит только при их отличии. Т.е. можно смело добавлять скрипт в многократный запуск и не переживать, что ярлычки на рабочем столе будут всё время пересоздаваться и менять своё местоположение. А также с помощью данного скрипта можно удалять ненужные ярлычки.

# Использование
В командной строке **Windows** введите следующую команду. Если необходимо скрыть отображение окна консоли, то вместо `cscript` можно использовать `wscript`.
```bat
cscript shortcut.min.js <location> <name> [<target> [<argument> [<directory>
                        [<description> [<icon> [<style> [<hotkey>]]]]]]]
```
- `<location>` - Путь к существующей папке, в которой нужно создать ярлычок (доступны `WshSpecialFolder`).
- `<name>` - Имя ярлычка без расширения, при несовпадении регистра ярлычок переименовывается.
- `<target>` - Путь к файлу, папке или url, при отсутствии ярлычок удаляется (доступны `WshSpecialFolder`).
- `<argument>` - Аргументы для целевого объекта, одинарные кавычки заменяются двойными.
- `<directory>` - Путь к существующей рабочей папке (доступны `WshSpecialFolder`).
- `<description>` - Описание ярлычка, которое отображается в подсказке.
- `<icon>` - Путь к файлу и индекс иконки, отделённый запятой.
- `<style>` - Стиль запуска окна приложения (1 - `Normal`, 3 - `Maximized`, 7 - `Minimized`).
- `<hotkey>` - Комбинация горячих клавиш.

# Примеры использования
Создать ярлычок для **Командной строки** на рабочем столе всех пользователей.
```bat
cscript shortcut.min.js AllUsersDesktop "Командная строка" "%windir%\system32\cmd.exe"
```
Создать ярлычок для главной страницы **Яндекса** на рабочем столе текущего пользователя.
```bat
cscript shortcut.min.js Desktop Яндекс https://yandex.ru/
```
Удалить ярлычок **Internet Explorer** с рабочего стола всех пользователей.
```bat
cscript shortcut.min.js AllUsersDesktop "Internet Explorer" ""
```