import enum

SW_VERSION = 'v0.0.2'

TEMPLATE_FOLDER_NAME = 'template'
ATTENDANCE_FOLDER_NAME = 'attendance'
RESULT_FOLDER_NAME = 'result'
ATTENDANCE_FILE_NAME = '考勤報表.xls'
TEMPLATE_FILE_NAME = '考勤彙總報表.xlsx'
INI_FILENAME = 'settings.ini'

ATTENDANCE_ID_STR = '工號 ：'
ATTENDANCE_ID_KEY = 'ID'
ATTENDANCE_NAME_KEY = 'NAME'
ATTENDANCE_DEPT_KEY = 'DEPT'

ATTENDANCE_ID_INDEX = 2
ATTENDANCE_NAME_INDEX = 9
ATTENDANCE_DEPT_INDEX = 17

SUMMARY_NAME_STR = '姓名'
SUMMARY_DATE_RANGE_STR = '考勤日期：'
SUMMARY_DATE_STR = '日期'


class IniEnum(enum.StrEnum):
    GENERAL_SECTION = 'General'
    START_YEAR = 'start_year'
    START_MON = 'start_mon'
    START_DAY = 'start_day'
    START_FILENAME = 'start_attendance_filename'
    END_YEAR = 'end_year'
    END_MON = 'end_mon'
    END_DAY = 'end_day'
    END_FILENAME = 'end_attendance_filename'


class AttendanceIndexEnum(enum.IntEnum):
    START = 0
    END = 1
