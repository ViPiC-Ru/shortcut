/* 0.1.1 создание, изменение и удаление ярлычка

cscript shortcut.min.js <location> <name> <target> [<argument> [<directory> [<description> [<icon> [<style> [<hotkey>]]]]]]

<location>      - путь к существующей папке, в которой нужно создать ярлычок (доступны ключи WshSpecialFolder)
<name>          - имя ярлычка без расширения, при несовпадении регистра ярлычок переименовывается
<target>        - путь к папке или файлу, на который создаётся ярлычок, при отсутствие целевого объекта ярлычок удаляется (доступны ключи WshSpecialFolder)
<argument>      - аргументы целевого объекта, одинарные кавычки заменяются двойными
<directory>     - путь к существующей рабочей папке (доступны ключи WshSpecialFolder)
<description>   - описание ярлычка, которое записывается в подсказку
<icon>          - путь к файлу и индекс иконки отделённый запятой
<style>         - целое число, как стиль запуска окна приложения
<hotkey>        - комбинация горячих клавиш

*/

(function (wsh, undefined) {// замыкаем что бы не сорить глобалы
    var value, flag, pattern, shell, fso, shortcut, file, location, name,
        target, argument, directory, description, icon, style, hotkey,
        delim = ".", ext = "lnk", error = 0;

    shell = new ActiveXObject("WScript.Shell");
    fso = new ActiveXObject("Scripting.FileSystemObject");
    // получаем строковые значения параметров
    if (!error) {// если нет ошибок
        for (var i = 0, iLen = wsh.arguments.length; i < iLen; i++) {
            value = wsh.arguments.item(i);// получаем очередное значение
            switch (i) {// поддерживаемые параметры
                case 0: location = value; break;
                case 1: name = value; break;
                case 2: target = value; break;
                case 3: argument = value; break;
                case 4: directory = value; break;
                case 5: description = value; break;
                case 6: icon = value; break;
                case 7: style = value; break;
                case 8: hotkey = value; break;
            };
        };
    };
    // проверяем местоположение
    if (!error) {// если нету ошибок
        // преобразовываем путь
        value = location ? shell.specialFolders(location) : null;
        if (!value) {// если папка задана не ключём WshSpecialFolder
            value = shell.expandEnvironmentStrings(location);
            value = fso.getAbsolutePathName(value);
        } else location = value;// задаём путь
        // проверяем путь
        if (fso.folderExists(value)) {// если папка существует
        } else if (fso.fileExists(value)) {// если файл фуществует
        } else error = 1;
    };
    // проверяем название
    if (!error) {// если нету ошибок
        pattern = /[\\\/:\*\?"<>\|]/;// запрещённые знаки
        if (name && !pattern.test(name)) {// если проверка пройдена
        } else error = 2;
    };
    // получаем ярлычок
    if (!error) {// если нету ошибок
        value = name + delim + ext;// имя ярлычка
        value = fso.buildPath(location, value);
        shortcut = shell.createShortcut(value);
    };
    // проверяем назначение и удаляем ярлычок если нужно
    if (!error) {// если нету ошибок
        // преобразовываем путь
        value = target ? shell.specialFolders(target) : null;
        if (!value) {// если папка задана не ключём WshSpecialFolder
            value = shell.expandEnvironmentStrings(target);
            value = fso.getAbsolutePathName(value);
        } else target = value;// задаём путь
        // проверяем путь
        if (fso.folderExists(value)) {// если папка существует
        } else if (fso.fileExists(value)) {// если файл фуществует
        } else {// если проверка не пройдена
            // пробуем удалить ярлычок
            if (shortcut.targetPath) {// если ярлычок существует
                try {// пробуем выполнить комманду
                    fso.deleteFile(shortcut.fullName);
                } catch (e) { };// игнорируем исключения
            };
            error = 3;
        };
    };
    // корректируем аргументы
    if (!error && argument) {// если нужно выполнить
        argument = argument.split("'").join('"');
    };
    // проверяем рабочую папку
    if (!error && directory) {// если нужно выполнить
        // преобразовываем путь
        value = directory ? shell.specialFolders(directory) : null;
        if (!value) {// если папка задана не ключём WshSpecialFolder
            value = shell.expandEnvironmentStrings(directory);
            value = fso.getAbsolutePathName(value);
        } else directory = value;// задаём путь
        // проверяем путь
        if (fso.folderExists(value)) {// если папка существует
        } else error = 4;
    };
    // проверяем стиль запуска
    if (!error && style) {// если нужно выполнить
        switch (style) {// поддерживаемые стили
            case "1": case "3": case "7": style = 1 * style; break;
            default: error = 5;
        };
    };
    // задаём назначение
    if (!error && "undefined" != typeof target) {// если нужно выполнить
        value = shell.expandEnvironmentStrings(target);
        value = fso.getAbsolutePathName(value).toLowerCase();
        if (shortcut.targetPath.toLowerCase() != value) {
            try {// пробуем выполнить комманду
                shortcut.targetPath = target;
                flag = true;
            } catch (e) {// если ошибка
                error = 6;
            };
        };
    };
    // задаём аргументы
    if (!error && "undefined" != typeof argument) {// если нужно выполнить
        if (shortcut.arguments != argument) {
            try {// пробуем выполнить комманду
                shortcut.arguments = argument;
                flag = true;
            } catch (e) {// если ошибка
                error = 7;
            };
        };
    };
    // задаём рабочую папку
    if (!error && "undefined" != typeof directory) {// если нужно выполнить
        if (shortcut.workingDirectory != directory) {
            try {// пробуем выполнить комманду
                shortcut.workingDirectory = directory;
                flag = true;
            } catch (e) {// если ошибка
                error = 8;
            };
        };
    };
    // задаём иконку
    if (!error && "undefined" != typeof icon) {// если нужно выполнить
        if (!~icon.indexOf(",")) value = icon + ",0";
        else value = icon.split(", ").join(",");
        if (shortcut.iconLocation != value) {
            try {// пробуем выполнить комманду
                shortcut.iconLocation = value;
                flag = true;
            } catch (e) {// если ошибка
                error = 9;
            };
        };
    };
    // задаём стиль запуска
    if (!error && "undefined" != typeof style) {// если нужно выполнить
        value = style ? style : 1;
        if (shortcut.windowStyle != value) {
            try {// пробуем выполнить комманду
                shortcut.windowStyle = value;
                flag = true;
            } catch (e) {// если ошибка
                error = 10;
            };
        };
    };
    // задаём горячие клавиши
    if (!error && "undefined" != typeof hotkey) {// если нужно выполнить
        if (shortcut.hotkey != hotkey) {
            try {// пробуем выполнить комманду
                shortcut.hotkey = hotkey;
                flag = true;
            } catch (e) {// если ошибка
                error = 11;
            };
        };
    };
    // задаём описание
    if (!error && "undefined" != typeof description) {// если нужно выполнить
        if (shortcut.description != description) {
            try {// пробуем выполнить комманду
                shortcut.description = description;
                flag = true;
            } catch (e) {// если ошибка
                error = 12;
            };
        };
    };
    // пробуем сохранить изменения
    if (!error && flag) {// если нужно выполнить
        try {// пробуем выполнить комманду
            shortcut.save();
        } catch (e) {// если ошибка
            error = 13;
        };
    };
    // переименовываем ярлычок
    if (!error) {// если нету ошибок
        value = name + delim + ext;// имя ярлычка
        file = fso.getFile(shortcut.fullName);
        if (file.name != value) {
            value = fso.buildPath(location, value);
            try {// пробуем выполнить комманду
                fso.moveFile(file.path, value);
            } catch (e) {// если ошибка
                error = 14;
            };
        };
    };
    // завершаем сценарий кодом
    wsh.quit(error);
})(WSH);