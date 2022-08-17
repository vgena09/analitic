unit FunConst;

interface                    
                                            
var
    PATH_PROG      : String;                       // СЕРВЕР: путь к файлу программы
    FPATH_PROG_INI : String;                       // СЕРВЕР: путь к файлу глобальных настроек
    PATH_BD        : String;                       // СЕРВЕР: путь к банку данных
    PATH_BD0       : String;                       // СЕРВЕР: путь к банку данных до префикса
    PATH_BD_DATA   : String;                       // СЕРВЕР: путь к данным
    PATH_BD_SHABLON: String;                       // СЕРВЕР: путь к шаблонам

    PATH_WORK      : String;                       // ПОЛЬЗОВАТЕЛЬ: путь к рабочему каталогу
    PATH_WORK_TEMP : String;                       // ПОЛЬЗОВАТЕЛЬ: путь к каталогу временных файлов
    PATH_WORK_INI  : String;                       // ПОЛЬЗОВАТЕЛЬ: путь к файлу локальных настроек

    IS_ADMIN       : Boolean;                      // Режим администратора
    IS_NET         : Boolean;                      // Работа в сети
    IS_TERMINATE   : Boolean;                      // Режим отключения пользователей

    IMAIN_REGION   : Integer;                      // Индекс главного региона
    SMAIN_REGION   : String;                       // Название главного региона

const
    STR_UNIQ         = '&А^Н$А!Л=И}Т<И@К~А`';    // Уникальная строка - для временной подмены

    SUB_RGN1         = ' [';     // Выделение субрегиона
    SUB_RGN2         = ']';

    CH1              = '{';
    CH2              = '}';

    CH_SPR           = '/';     // Разделитель в кодировке адреса
    CH_KAV           = Chr(39);
    CH_NEW           = Chr(13)+Chr(10);
    CH_VAL           = ' = ';

    IDLOG            = 'OSA:';
    CELL_LOG         = '$A$1';
    CELL_YEAR        = '$B$1';
    CELL_MONTH       = '$C$1';
    CELL_REGION      = '$D$1';
    CELL_FULL        = '$E$1';
    CELL_CAPTION     = '$F$1';    // Для некоторых форм

    TIMER_INFO_1     = 10000;
    TIMER_INFO_2     = 30000;

{******************  ИДЕНТИФИКАТОРЫ  ******************************************}

        TAG_ANALIZ_COMPARE    = 1;
        TAG_ANALIZ_RATING     = 2;
        TAG_ANALIZ_REGION     = 3;
        TAG_ANALIZ_DYNAMIC    = 4;

{******************  ПУТИ  ****************************************************}

        PATH_MYDOC            = 'Статистика рабочий каталог\'; // Папка рабочая в моих документах
        PATH_TEMP             = 'Temp\';                       // Папка для временных файлов
        PATH_BANK0            = 'Данные';                      // Папка БАЗОВОГО банка
        PATH_DATA             = 'Банк данных\';                // Папка для файлов банка данных
        PATH_FORM             = 'Формы\';                      // Папка для файлов форм
        PATH_SHABLON          = 'Шаблоны\';                    // Папка для шаблонов
        PATH_ARJ              = 'Резервные копии\';            // Папка для файлов архивации
        PATH_TEMP_RAR         = 'Rar\';                        // Папка временная для файлов экспорта/импорта


{******************  Ф А Й Л Ы  ***********************************************}

        PROG_INI              = 'Данные\Глобальные настройки.ini';
        WORK_INI              = 'Локальные настройки.ini';

        FILE_LOCAL_INDEX      = 'Индексы.mdb';                      // Файл индексов
        FILE_LOCAL_INFO       = 'Информатор.txt';                   // Файл информатора
        FILE_LOCAL_HELP_ADMIN = 'Инструкция_администратора.chm';    // Файл помощи администратора
        FILE_LOCAL_HELP_USER  = 'Инструкция_пользователя.chm';      // Файл помощи пользователя

        TEMP_XLS              = 'Текущий отчет.xls';
        TEMP_COMMENT          = 'Комментарий.txt';                  // Файл комментария при экспорте


{******************************************************************************}
        F_COUNTER             = 'Счётчик';
        F_CAPTION             = 'Заголовок';
        F_PARENT              = 'Родитель';
        F_FILE                = 'Файл';
        F_ICON                = 'Иконка';
        F_HINT                = 'Пояснение';
        F_NUMERIC             = 'Номер';
        F_BEGIN               = 'Дата1';
        F_END                 = 'Дата2';
        F_PASSWORD            = 'Пароль';
        F_DATE                = 'Дата';

{*******  БАЗА ДАННЫХ ИНДЕКСОВ  ***********************************************}
        TABLE_REGIONS         = 'Регионы';
           NREGIONS_COUNTER   = 0;
           NREGIONS_CAPTION   = 1;
           NREGIONS_PARENT    = 2;
           NREGIONS_GROUP     = 3;
           NREGIONS_BEGIN     = 4;
           NREGIONS_END       = 5;

        TABLE_REGREPL         = 'Регионы подмена';
           REGREPL_R1         = 'Регион1';
           REGREPL_S1         = 'Субрегион1';
           REGREPL_R2         = 'Регион2';
           REGREPL_S2         = 'Субрегион2';

        TABLE_FORMS           = 'Формы';
           FORMS_EXPORT       = 'Экспорт бланка';
           FORMS_OTHER        = 'Иные регионы';
           FORMS_1            = 'Месячный';
           FORMS_3            = 'Квартальный';
           FORMS_REGION       = 'Регион';
           NFORMS_COUNTER     = 0;
           NFORMS_PARENT      = 1;
           NFORMS_ICON        = 2;
           NFORMS_EXPORT      = 3;
           NFORMS_REGION      = 4;
           NFORMS_CAPTION     = 5;
           NFORMS_FILE        = 6;
           NFORMS_MONTH1      = 7;
           NFORMS_MONTH3      = 8;
           NFORMS_BEGIN       = 9;
           NFORMS_END         = 10;
           NFORMS_CELL_ID     = 11;
           NFORMS_STR_ID      = 12;
           NFORMS_CELL_PERIOD = 13;
           NFORMS_CELL_REGION = 14;
           IDFORM_VIEW_COL    = 40;             // ID формы: просмотр столбца
           IDFORM_ACOMPARE    = 41;             // ID формы: анализ сравнение
           IDFORM_ARATING     = 42;             // ID формы: анализ рейтинг
           IDFORM_AREGION     = 43;             // ID формы: анализ регион
           IDFORM_ADYNAMIC    = 44;             // ID формы: анализ динамика

        TABLE_TABLES          = 'Таблицы';
           NTABLES_COUNTER    = 0;
           NTABLES_CAPTION    = 1;
           NTABLES_NOSTAND    = 2;
           NTABLES_CEP        = 3;
           NTABLES_BEGIN      = 4;
           NTABLES_END        = 5;

        TABLE_ROWS            = 'Строки';
           NROWS_NUMERIC      = 0;
           NROWS_CAPTION      = 1;
           NROWS_TABLE        = 2;

        TABLE_COLS            = 'Столбцы';
           NCOLS_NUMERIC      = 0;
           NCOLS_CAPTION      = 1;
           NCOLS_TABLE        = 2;

        TABLE_FORMULS         = 'Формулы';

        TABLE_ORKEY           = 'Альтернативные коды';
           NORKEY_1           = 0;
           NORKEY_2           = 1;

        TABLE_EXCEPT          = 'Исключения из суммы';
           EXCEPT_ADRES       = 'Адрес';

        TABLE_SEP             = 'СЭП';


{*******************    I N I - ФАЙЛ МАТРИЦЫ    *******************************}
     INI_PARAM                    = 'Параметры';
        INI_PARAM_TABLES          = 'Таблицы';
        INI_PARAM_PAGES           = 'Листы';
        INI_PARAM_PAGELINKS       = 'Ярлычки';

     INI_PAGE                     = 'Лист ';    // Лист №
        INI_PAGE_IN               = 'Ввод';
        INI_PAGE_OUT              = 'Вывод';
        INI_PAGE_PROTECT          = 'Защита';                                   // по умолчанию: true
        INI_PAGE_TYPE             = 'Анализ';
           INI_PAGE_TYPE_STANDART = 'Стандартный';
           INI_PAGE_TYPE_COMPARE  = 'Сравнение';
           INI_PAGE_TYPE_RATING   = 'Рейтинг';
           INI_PAGE_TYPE_REGION   = 'Регион';
           INI_PAGE_TYPE_DYNAMIC  = 'Динамика';
        INI_PAGE_FIXSIZETABLE     = 'Анализ - Фиксированный размер таблиц';
        INI_PAGE_STEP             = 'Шаг в месяцах';                            // по умолчанию: 12
        INI_PAGE_LOOP             = 'Количество шагов';                         // по умолчанию: 2
        INI_PAGE_ROW_BEFORE       = 'Строк в шапке';                            // для нестандартного отчета по умолчанию: 0
        INI_PAGE_ROW_AFTER        = 'Строк в подвале';                          // для нестандартного отчета по умолчанию: 0
        INI_PAGE_COL_KOD          = 'Код';                                      // значение ячейки для нестандартного отчета: номер первого столбца области чисел
        INI_PAGE_ROW_COLOR        = 'Строка цвета';                             // значение ячейки для отчета Сравнение: строка кода условного форматирования - по умолчанию: 41
        INI_PAGE_ROW_CAPTION      = 'Строка заголовка';                         // значение ячейки для отчетов Сравнение и Рейтинг: строка названия формулы - по умолчанию: 45
        INI_PAGE_ROW_REGION       = 'Строка региона';                           // значение ячейки для отчета Регион: строка названия региона - по умолчанию: 46

        INI_PAGE_STAT_PARAM       = 'Параметр';                                 // Параметр+N (c 1) - статистические параметры (SCaption, SFormula, SColor), разделенные |   его отсутствие - брать параметры с правой панели

        INI_PAGE_DAT              = 'Область данных';                           // Адрес области чисел (данных)

        INI_FORMAT_INT            = 'Формат целого числа';                      // для нестандартного отчета по умолчанию: # ###;-# ###;;@
        INI_FORMAT_FLOAT          = 'Формат дробного числа';                    // для нестандартного отчета по умолчанию: # ##0,0;-# ##0,0;;@



{*******************   INI - ФАЙЛ ГЛОБАЛЬНЫХ НАСТРОЕК   ***********************}
     INI_SET                      = 'Установки';
        INI_SET_NET               = 'Работа в сети';
        INI_SET_IMAIN_REGION      = 'Индекс главного региона';
        INI_SET_PAS_IMPORT_EXPORT = 'Пароль экспорта/импорта';
        INI_SET_TERMINATE         = 'Отключение пользователей';


{********************  INI - ФАЙЛ ЛОКАЛЬНЫХ НАСТРОЕК   ************************}
     {Файл сохранений}
     INI_SELECT                   = 'Текущий выбор';
        INI_SELECT_PATH           = 'Путь';
        INI_SELECT_FORM           = 'Форма';

     {Позиция курсора}
     INI_ACTIVE_CELL              = 'Активная ячейка';      // Ключ - ID формы

     {Секции-истории}
     INI_FIND                     = 'Поиск (история)';
     INI_KOD_CORRECT              = 'Коррекция кода (история)';
     INI_KOD_FORMULA              = 'Редактор формулы (история)';

     {Секция: УСТАНОВКИ}
     INI_SET_ARJ_PATH             = 'Архивация - путь';
     INI_SET_EXPORT_PATH          = 'Экспорт - путь';
     INI_SET_IMPORT_SUBREG        = 'Импорт - субрегионы';
     INI_SET_IMPORT_VERIFY_RAR    = 'Импорт - проверка RAR';
     INI_SET_IMPORT_VERIFY_XLS    = 'Импорт - проверка XLS';
     INI_SET_KOD_MODIFY           = 'Код - режим редактирования';
     INI_SET_SEP_DETAIL           = 'СЭП - оптимизация';
     INI_SET_SEP_MONTH            = 'СЭП - число месяцев';
     INI_SET_PARAM_SELECT_ITEM    = 'Выделенный параметр аналитики';

     {Формы}
     INI_FORM_PARAM               = 'Параметры формы: ';
        INI_FORM_PARAM_LEFT       = 'Слева';
        INI_FORM_PARAM_TOP        = 'Сверху';
        INI_FORM_PARAM_WIDTH      = 'Ширина';
        INI_FORM_PARAM_HEIGHT     = 'Высота';
        INI_FORM_PARAM_MAXIMIZE   = 'Развернута';
        INI_FORM_PARAM_SEP_LEFT   = 'Сепаратор слева';
        INI_FORM_PARAM_SEP_RIGHT  = 'Сепаратор справа';

     INI_FORM_KOD                 = 'Форма выбора кода';
        FORM_KOD_TAB              = 'Ширина столбцов - таблицы: ';
        FORM_KOD_ROW              = 'Ширина столбцов - строки: ';
        FORM_KOD_COL              = 'Ширина столбцов - столбцы: ';

     {Аналитика}
     INI_ANALITIC_PARAM           = 'Параметры аналитики';

{*******************   INI - ФАЙЛ ГЛОБАЛЬНЫХ НАСТРОЕК   ***********************}
     INI_INFO_RUN                 = 'Сообщение при запуске';
     INI_INFO_TERMINATE           = 'Требование выхода';


{*******************************  ИКОНКИ  *************************************}
     ICO_PATH_MIN                 = 2;
     ICO_PATH_MAX                 = 12;
     ICO_FORM_MIN                 = 14;
     ICO_FORM_MAX                 = 18;
     ICO_FORM_DETAIL_REPORT       = 14;  /// Иконка отчета
     ICO_FORM_DETAIL_TABLE        = 20;  /// Иконка таблицы
     ICO_FORM_DETAIL_ROW          = 22;  /// Иконка строки
     ICO_FORM_DETAIL_COL          = 24;  /// Иконка столбца
     ICO_FORM_DETAIL_ANALITIC1    = 16;  /// ???
     ICO_FORM_DETAIL_ANALITIC2    = 18;  /// ???

     ICO_FORM_FREE_XLS            = 14;  /// Иконка произвольного XLS - файла
     ICO_FORM_FREE_PDF            = 26;  /// Иконка произвольного PDF - файла
     ICO_FORM_FREE_MDI            = 28;  /// Иконка произвольного MDI - файла

     ICO_NET_ON                   = 5;
     ICO_NET_OFF                  = 6;

     ICO_SYS_FORMULA              = 8;

{********************************  ЦВЕТА  *************************************}
     COLOR_ROW                    = $00FF8000;

{******************************  STATUS BAR ***********************************}
     STATUS_MAIN                  = 0;
     STATUS_KOD                   = 1;
     STATUS_NET                   = 2;
     STATUS_REGION                = 3;
     STATUS_TIME                  = 4;


{******************************************************************************}
     ID_VERSION_PROG = 'Версия программы: ';


{******************************************************************************}
     {Авторство}
     MSGQWEST = 'Тел.:   +375-17-2036508 (Генеральная прокуратура)' + CH_NEW +
                'Тел.:   +375-17-2223633 (Следственный комитет)' + CH_NEW +
                'E-mail: gena09@mail.ru (Разработчик)';

implementation


end.
