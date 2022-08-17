unit FunConst;

interface                    
                                            
var
    PATH_PROG      : String;                       // ������: ���� � ����� ���������
    FPATH_PROG_INI : String;                       // ������: ���� � ����� ���������� ��������
    PATH_BD        : String;                       // ������: ���� � ����� ������
    PATH_BD0       : String;                       // ������: ���� � ����� ������ �� ��������
    PATH_BD_DATA   : String;                       // ������: ���� � ������
    PATH_BD_SHABLON: String;                       // ������: ���� � ��������

    PATH_WORK      : String;                       // ������������: ���� � �������� ��������
    PATH_WORK_TEMP : String;                       // ������������: ���� � �������� ��������� ������
    PATH_WORK_INI  : String;                       // ������������: ���� � ����� ��������� ��������

    IS_ADMIN       : Boolean;                      // ����� ��������������
    IS_NET         : Boolean;                      // ������ � ����
    IS_TERMINATE   : Boolean;                      // ����� ���������� �������������

    IMAIN_REGION   : Integer;                      // ������ �������� �������
    SMAIN_REGION   : String;                       // �������� �������� �������

const
    STR_UNIQ         = '&�^�$�!�=�}�<�@�~�`';    // ���������� ������ - ��� ��������� �������

    SUB_RGN1         = ' [';     // ��������� ����������
    SUB_RGN2         = ']';

    CH1              = '{';
    CH2              = '}';

    CH_SPR           = '/';     // ����������� � ��������� ������
    CH_KAV           = Chr(39);
    CH_NEW           = Chr(13)+Chr(10);
    CH_VAL           = ' = ';

    IDLOG            = 'OSA:';
    CELL_LOG         = '$A$1';
    CELL_YEAR        = '$B$1';
    CELL_MONTH       = '$C$1';
    CELL_REGION      = '$D$1';
    CELL_FULL        = '$E$1';
    CELL_CAPTION     = '$F$1';    // ��� ��������� ����

    TIMER_INFO_1     = 10000;
    TIMER_INFO_2     = 30000;

{******************  ��������������  ******************************************}

        TAG_ANALIZ_COMPARE    = 1;
        TAG_ANALIZ_RATING     = 2;
        TAG_ANALIZ_REGION     = 3;
        TAG_ANALIZ_DYNAMIC    = 4;

{******************  ����  ****************************************************}

        PATH_MYDOC            = '���������� ������� �������\'; // ����� ������� � ���� ����������
        PATH_TEMP             = 'Temp\';                       // ����� ��� ��������� ������
        PATH_BANK0            = '������';                      // ����� �������� �����
        PATH_DATA             = '���� ������\';                // ����� ��� ������ ����� ������
        PATH_FORM             = '�����\';                      // ����� ��� ������ ����
        PATH_SHABLON          = '�������\';                    // ����� ��� ��������
        PATH_ARJ              = '��������� �����\';            // ����� ��� ������ ���������
        PATH_TEMP_RAR         = 'Rar\';                        // ����� ��������� ��� ������ ��������/�������


{******************  � � � � �  ***********************************************}

        PROG_INI              = '������\���������� ���������.ini';
        WORK_INI              = '��������� ���������.ini';

        FILE_LOCAL_INDEX      = '�������.mdb';                      // ���� ��������
        FILE_LOCAL_INFO       = '����������.txt';                   // ���� �����������
        FILE_LOCAL_HELP_ADMIN = '����������_��������������.chm';    // ���� ������ ��������������
        FILE_LOCAL_HELP_USER  = '����������_������������.chm';      // ���� ������ ������������

        TEMP_XLS              = '������� �����.xls';
        TEMP_COMMENT          = '�����������.txt';                  // ���� ����������� ��� ��������


{******************************************************************************}
        F_COUNTER             = '�������';
        F_CAPTION             = '���������';
        F_PARENT              = '��������';
        F_FILE                = '����';
        F_ICON                = '������';
        F_HINT                = '���������';
        F_NUMERIC             = '�����';
        F_BEGIN               = '����1';
        F_END                 = '����2';
        F_PASSWORD            = '������';
        F_DATE                = '����';

{*******  ���� ������ ��������  ***********************************************}
        TABLE_REGIONS         = '�������';
           NREGIONS_COUNTER   = 0;
           NREGIONS_CAPTION   = 1;
           NREGIONS_PARENT    = 2;
           NREGIONS_GROUP     = 3;
           NREGIONS_BEGIN     = 4;
           NREGIONS_END       = 5;

        TABLE_REGREPL         = '������� �������';
           REGREPL_R1         = '������1';
           REGREPL_S1         = '���������1';
           REGREPL_R2         = '������2';
           REGREPL_S2         = '���������2';

        TABLE_FORMS           = '�����';
           FORMS_EXPORT       = '������� ������';
           FORMS_OTHER        = '���� �������';
           FORMS_1            = '��������';
           FORMS_3            = '�����������';
           FORMS_REGION       = '������';
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
           IDFORM_VIEW_COL    = 40;             // ID �����: �������� �������
           IDFORM_ACOMPARE    = 41;             // ID �����: ������ ���������
           IDFORM_ARATING     = 42;             // ID �����: ������ �������
           IDFORM_AREGION     = 43;             // ID �����: ������ ������
           IDFORM_ADYNAMIC    = 44;             // ID �����: ������ ��������

        TABLE_TABLES          = '�������';
           NTABLES_COUNTER    = 0;
           NTABLES_CAPTION    = 1;
           NTABLES_NOSTAND    = 2;
           NTABLES_CEP        = 3;
           NTABLES_BEGIN      = 4;
           NTABLES_END        = 5;

        TABLE_ROWS            = '������';
           NROWS_NUMERIC      = 0;
           NROWS_CAPTION      = 1;
           NROWS_TABLE        = 2;

        TABLE_COLS            = '�������';
           NCOLS_NUMERIC      = 0;
           NCOLS_CAPTION      = 1;
           NCOLS_TABLE        = 2;

        TABLE_FORMULS         = '�������';

        TABLE_ORKEY           = '�������������� ����';
           NORKEY_1           = 0;
           NORKEY_2           = 1;

        TABLE_EXCEPT          = '���������� �� �����';
           EXCEPT_ADRES       = '�����';

        TABLE_SEP             = '���';


{*******************    I N I - ���� �������    *******************************}
     INI_PARAM                    = '���������';
        INI_PARAM_TABLES          = '�������';
        INI_PARAM_PAGES           = '�����';
        INI_PARAM_PAGELINKS       = '�������';

     INI_PAGE                     = '���� ';    // ���� �
        INI_PAGE_IN               = '����';
        INI_PAGE_OUT              = '�����';
        INI_PAGE_PROTECT          = '������';                                   // �� ���������: true
        INI_PAGE_TYPE             = '������';
           INI_PAGE_TYPE_STANDART = '�����������';
           INI_PAGE_TYPE_COMPARE  = '���������';
           INI_PAGE_TYPE_RATING   = '�������';
           INI_PAGE_TYPE_REGION   = '������';
           INI_PAGE_TYPE_DYNAMIC  = '��������';
        INI_PAGE_FIXSIZETABLE     = '������ - ������������� ������ ������';
        INI_PAGE_STEP             = '��� � �������';                            // �� ���������: 12
        INI_PAGE_LOOP             = '���������� �����';                         // �� ���������: 2
        INI_PAGE_ROW_BEFORE       = '����� � �����';                            // ��� �������������� ������ �� ���������: 0
        INI_PAGE_ROW_AFTER        = '����� � �������';                          // ��� �������������� ������ �� ���������: 0
        INI_PAGE_COL_KOD          = '���';                                      // �������� ������ ��� �������������� ������: ����� ������� ������� ������� �����
        INI_PAGE_ROW_COLOR        = '������ �����';                             // �������� ������ ��� ������ ���������: ������ ���� ��������� �������������� - �� ���������: 41
        INI_PAGE_ROW_CAPTION      = '������ ���������';                         // �������� ������ ��� ������� ��������� � �������: ������ �������� ������� - �� ���������: 45
        INI_PAGE_ROW_REGION       = '������ �������';                           // �������� ������ ��� ������ ������: ������ �������� ������� - �� ���������: 46

        INI_PAGE_STAT_PARAM       = '��������';                                 // ��������+N (c 1) - �������������� ��������� (SCaption, SFormula, SColor), ����������� |   ��� ���������� - ����� ��������� � ������ ������

        INI_PAGE_DAT              = '������� ������';                           // ����� ������� ����� (������)

        INI_FORMAT_INT            = '������ ������ �����';                      // ��� �������������� ������ �� ���������: # ###;-# ###;;@
        INI_FORMAT_FLOAT          = '������ �������� �����';                    // ��� �������������� ������ �� ���������: # ##0,0;-# ##0,0;;@



{*******************   INI - ���� ���������� ��������   ***********************}
     INI_SET                      = '���������';
        INI_SET_NET               = '������ � ����';
        INI_SET_IMAIN_REGION      = '������ �������� �������';
        INI_SET_PAS_IMPORT_EXPORT = '������ ��������/�������';
        INI_SET_TERMINATE         = '���������� �������������';


{********************  INI - ���� ��������� ��������   ************************}
     {���� ����������}
     INI_SELECT                   = '������� �����';
        INI_SELECT_PATH           = '����';
        INI_SELECT_FORM           = '�����';

     {������� �������}
     INI_ACTIVE_CELL              = '�������� ������';      // ���� - ID �����

     {������-�������}
     INI_FIND                     = '����� (�������)';
     INI_KOD_CORRECT              = '��������� ���� (�������)';
     INI_KOD_FORMULA              = '�������� ������� (�������)';

     {������: ���������}
     INI_SET_ARJ_PATH             = '��������� - ����';
     INI_SET_EXPORT_PATH          = '������� - ����';
     INI_SET_IMPORT_SUBREG        = '������ - ����������';
     INI_SET_IMPORT_VERIFY_RAR    = '������ - �������� RAR';
     INI_SET_IMPORT_VERIFY_XLS    = '������ - �������� XLS';
     INI_SET_KOD_MODIFY           = '��� - ����� ��������������';
     INI_SET_SEP_DETAIL           = '��� - �����������';
     INI_SET_SEP_MONTH            = '��� - ����� �������';
     INI_SET_PARAM_SELECT_ITEM    = '���������� �������� ���������';

     {�����}
     INI_FORM_PARAM               = '��������� �����: ';
        INI_FORM_PARAM_LEFT       = '�����';
        INI_FORM_PARAM_TOP        = '������';
        INI_FORM_PARAM_WIDTH      = '������';
        INI_FORM_PARAM_HEIGHT     = '������';
        INI_FORM_PARAM_MAXIMIZE   = '����������';
        INI_FORM_PARAM_SEP_LEFT   = '��������� �����';
        INI_FORM_PARAM_SEP_RIGHT  = '��������� ������';

     INI_FORM_KOD                 = '����� ������ ����';
        FORM_KOD_TAB              = '������ �������� - �������: ';
        FORM_KOD_ROW              = '������ �������� - ������: ';
        FORM_KOD_COL              = '������ �������� - �������: ';

     {���������}
     INI_ANALITIC_PARAM           = '��������� ���������';

{*******************   INI - ���� ���������� ��������   ***********************}
     INI_INFO_RUN                 = '��������� ��� �������';
     INI_INFO_TERMINATE           = '���������� ������';


{*******************************  ������  *************************************}
     ICO_PATH_MIN                 = 2;
     ICO_PATH_MAX                 = 12;
     ICO_FORM_MIN                 = 14;
     ICO_FORM_MAX                 = 18;
     ICO_FORM_DETAIL_REPORT       = 14;  /// ������ ������
     ICO_FORM_DETAIL_TABLE        = 20;  /// ������ �������
     ICO_FORM_DETAIL_ROW          = 22;  /// ������ ������
     ICO_FORM_DETAIL_COL          = 24;  /// ������ �������
     ICO_FORM_DETAIL_ANALITIC1    = 16;  /// ???
     ICO_FORM_DETAIL_ANALITIC2    = 18;  /// ???

     ICO_FORM_FREE_XLS            = 14;  /// ������ ������������� XLS - �����
     ICO_FORM_FREE_PDF            = 26;  /// ������ ������������� PDF - �����
     ICO_FORM_FREE_MDI            = 28;  /// ������ ������������� MDI - �����

     ICO_NET_ON                   = 5;
     ICO_NET_OFF                  = 6;

     ICO_SYS_FORMULA              = 8;

{********************************  �����  *************************************}
     COLOR_ROW                    = $00FF8000;

{******************************  STATUS BAR ***********************************}
     STATUS_MAIN                  = 0;
     STATUS_KOD                   = 1;
     STATUS_NET                   = 2;
     STATUS_REGION                = 3;
     STATUS_TIME                  = 4;


{******************************************************************************}
     ID_VERSION_PROG = '������ ���������: ';


{******************************************************************************}
     {���������}
     MSGQWEST = '���.:   +375-17-2036508 (����������� �����������)' + CH_NEW +
                '���.:   +375-17-2223633 (������������ �������)' + CH_NEW +
                'E-mail: gena09@mail.ru (�����������)';

implementation


end.
