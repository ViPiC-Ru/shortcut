/* 0.2.0 создание, изменение и удаление ярлычка

cscript shortcut.min.js <location> <name> [<target> [<argument> [<directory> [<description> [<icon> [<style> [<hotkey>]]]]]]]

<location>      - Путь к существующей папке, в которой нужно создать ярлычок (доступны ключи WshSpecialFolder).
<name>          - Имя ярлычка без расширения, при несовпадении регистра ярлычок переименовывается.
<target>        - Путь к целевому файлу, папке или url, при отсутствии ярлычок удаляется (доступны ключи WshSpecialFolder).
<argument>      - Аргументы для целевого объекта, одинарные кавычки заменяются двойными.
<directory>     - Путь к существующей рабочей папке (доступны ключи WshSpecialFolder).
<description>   - Описание ярлычка, которое отображается в подсказке.
<icon>          - Путь к файлу и индекс иконки, отделённый запятой.
<style>         - Стиль запуска окна приложения (1 - Normal, 3 - Maximized, 7 - Minimized).
<hotkey>        - Комбинация горячих клавиш.

*/

var shortcut = new App({
    lnkExt: "lnk",                  // расширение для стандартного ярлычка
    urlExt: "url",                  // расширение для интернет ярлычка
    escСhar: "^",                   // экранирующий символ
    extDelim: "."                   // разделитель для расширеня
});

// подключаем зависимые свойства приложения
(function (wsh, app, undefined) {
    app.lib.extend(app, {
        fun: {// зависимые функции частного назначения
        },
        init: function () {// функция инициализации приложения
            var key, value, name, index, location, target, length, list, ext, fso,
                shell, path, data, isFound, isUrl, config = {}, error = 0;

            shell = new ActiveXObject("WScript.Shell");
            fso = new ActiveXObject("Scripting.FileSystemObject");
            // получаем параметры
            if (!error) {// если нет ошибок
                length = wsh.arguments.length;// получаем длину
                for (index = 0; index < length; index++) {// пробигаемся по параметрам
                    value = wsh.arguments.item(index);// получаем очередное значение
                    // путь к папке расположения ярлычка
                    key = "location";// ключ проверяемого параметра
                    if (0 == index && !(key in config)) {// если нужно выполнить
                        // корректируем значение
                        path = value ? shell.specialFolders(value) : null;
                        if (path) value = path;// если папка задана ключём WshSpecialFolder
                        // присваиваем значение
                        config[key] = value;// задаём значение
                        continue;// переходим к следующему параметру
                    };
                    // имя ярлычка без расширения
                    key = "name";// ключ проверяемого параметра
                    if (1 == index && !(key in config)) {// если нужно выполнить
                        // присваиваем значение
                        config[key] = value;// задаём значение
                        continue;// переходим к следующему параметру
                    };
                    // путь к целевому объекту
                    key = "target";// ключ проверяемого параметра
                    if (2 == index && !(key in config)) {// если нужно выполнить
                        // корректируем значение
                        path = value ? shell.specialFolders(value) : null;
                        if (path) value = path;// если папка задана ключём WshSpecialFolder
                        // присваиваем значение
                        config[key] = value;// задаём значение
                        continue;// переходим к следующему параметру
                    };
                    // аргументы целевого объекта
                    key = "argument";// ключ проверяемого параметра
                    if (3 == index && !(key in config)) {// если нужно выполнить
                        // корректируем значение
                        value = value.split("'").join('"');
                        // присваиваем значение
                        config[key] = value;// задаём значение
                        continue;// переходим к следующему параметру
                    };
                    // путь к рабочей папке
                    key = "directory";// ключ проверяемого параметра
                    if (4 == index && !(key in config)) {// если нужно выполнить
                        // корректируем значение
                        path = value ? shell.specialFolders(value) : null;
                        if (path) value = path;// если папка задана ключём WshSpecialFolder
                        // присваиваем значение
                        config[key] = value;// задаём значение
                        continue;// переходим к следующему параметру
                    };
                    // описание ярлычка
                    key = "description";// ключ проверяемого параметра
                    if (5 == index && !(key in config)) {// если нужно выполнить
                        // присваиваем значение
                        config[key] = value;// задаём значение
                        continue;// переходим к следующему параметру
                    };
                    // путь к файлу и индекс иконки
                    key = "icon";// ключ проверяемого параметра
                    if (6 == index && !(key in config)) {// если нужно выполнить
                        // присваиваем значение
                        config[key] = value;// задаём значение
                        continue;// переходим к следующему параметру
                    };
                    // стиль запуска окна приложения
                    key = "style";// ключ проверяемого параметра
                    if (7 == index && !(key in config)) {// если нужно выполнить
                        // корректируем значение
                        value = Number(value);
                        // присваиваем значение
                        config[key] = value;// задаём значение
                        continue;// переходим к следующему параметру
                    };
                    // комбинация горячих клавиш
                    key = "hotkey";// ключ проверяемого параметра
                    if (8 == index && !(key in config)) {// если нужно выполнить
                        // присваиваем значение
                        config[key] = value;// задаём значение
                        continue;// переходим к следующему параметру
                    };
                    // если закончились параметры конфигурации
                    break;// остававливаем получние параметров
                };
            };
            // проверяем обязательные поля
            if (!error) {// если нет ошибок
                if (("location" in config) && ("name" in config)) {
                } else error = 1;
            };
            // проверяем стиль запуска
            if (!error && ("style" in config)) {// если нужно выполнить
                if (app.lib.hasValue([1, 3, 7], config.style, false)) {
                } else error = 2;
            };
            // работаем с ярлычком
            if (!error) {// если нет ошибок
                data = {};// данные для создания ярлычка
                isUrl = config.target && ~config.target.indexOf("/");
                // формируем путь к расположению ярлычка
                isFound = false;// найден ли существующий ярлык
                if (!config.target) list = [app.val.urlExt, app.val.lnkExt];
                else list = [isUrl ? app.val.urlExt : app.val.lnkExt];
                for (var i = 0, iLen = list.length; i < iLen && !isFound; i++) {
                    ext = list[i];// получаем очередное расширение
                    name = config.name + app.val.extDelim + ext;// имя с расширением
                    path = fso.buildPath(config.location, name);// путь к ярлычку
                    path = path.split(app.val.escСhar).join("");
                    location = shell.expandEnvironmentStrings(path);
                    if (fso.fileExists(location)) isFound = true;
                };
                data.fullName = path;
                // формируем путь к целевому объекту
                path = config.target;
                if (path && !isUrl) {// если есть целевой объект и это не url
                    path = path.split(app.val.escСhar).join("");
                    target = shell.expandEnvironmentStrings(path);
                    if (!fso.fileExists(target)) path = null;
                };
                if ("target" in config) data.targetPath = path;
                // исправляем экранирование в других параметрах
                for (var key in config) {// пробигаемся по параметрам
                    if (app.lib.hasValue(["argument", "directory", "icon"], key, true)) {
                        path = config[key];// получаем очередное значение
                        path = path.split(app.val.escСhar).join("");
                        config[key] = path;
                    };
                };
                // формируем остальные параметры
                if ("argument" in config) data.arguments = config.argument;
                if ("directory" in config) data.workingDirectory = config.directory;
                if ("description" in config) data.description = config.description;
                if ("icon" in config) data.iconLocation = config.icon;
                if ("style" in config) data.windowStyle = config.style;
                if ("hotkey" in config) data.hotkey = config.hotkey;
                // выполняем действие с ярлычком
                if (obj = app.wsh.setShortcut(data)) {// если выполнено успешно
                } else error = 3;
            };
            // завершаем сценарий кодом
            wsh.quit(error);
        }
    });
})(WSH, shortcut);
// запускаем инициализацию
shortcut.init();