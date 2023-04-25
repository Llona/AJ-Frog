from datetime import datetime
from datetime import timedelta
# from time import sleep
# import random
# import settings
# import numpy as np
import os
import json
import configparser
from collections import OrderedDict


class JsonControl(object):
    def __init__(self, json_full_path):
        self.json_full_path = json_full_path
        self.json_format = 'utf8'
        self.format_list = ['utf8', 'utf-8-sig', 'utf16', None, 'big5', 'gbk', 'gb2312']
        if os.path.exists(self.json_full_path):
            self.try_file_format()

    def try_file_format(self):
        for file_format in self.format_list:
            try:
                with open(self.json_full_path, 'r', encoding=file_format) as file:
                    json.load(file)
                self.json_format = file_format
                # print('find correct format {} in json file: {}'.format(file_format, self.json_full_path))
                return
            except Exception as e:
                # print('checking {} format: {}'.format(self.json_full_path, file_format))
                str(e)

    def read_json(self):
        try:
            with open(self.json_full_path, 'r', encoding=self.json_format) as file:
                return json.load(file)
        except Exception as e:
            print("Error! 讀取cfg設定檔發生錯誤!: {} {}".format(self.json_full_path, e))
            raise

    def write_json(self, json_content):
        try:
            with open(self.json_full_path, 'w', encoding=self.json_format) as file:
                json.dump(json_content, file, ensure_ascii=False, indent=4, separators=(',', ':'))
        except Exception as e:
            print("Error! 寫入cfg設定檔發生錯誤! {} {}".format(self.json_full_path, e))
            # str(e)
            raise


class IniControl(object):
    def __init__(self, ini_full_path):
        self.ini_full_path = ini_full_path
        self.ini_format = 'utf8'
        self.format_list = ['utf8', 'utf-8-sig', 'utf16', None, 'big5', 'gbk', 'gb2312']
        self.try_ini_format()

    def try_ini_format(self):
        for file_format in self.format_list:
            try:
                config_lh = configparser.ConfigParser()
                with open(self.ini_full_path, 'r', encoding=file_format) as file:
                    config_lh.read_file(file)
                self.ini_format = file_format
                print('find correct format {} in ini file: {}'.format(file_format, self.ini_full_path))
                return
            except Exception as e:
                print('checking {} format: {}'.format(self.ini_full_path, file_format))
                str(e)

    def read_config(self, section, key):
        try:
            config_lh = configparser.ConfigParser()
            file_ini_lh = open(self.ini_full_path, 'r', encoding=self.ini_format)
            config_lh.read_file(file_ini_lh)
            file_ini_lh.close()
            return config_lh.get(section, key)
        except Exception as e:
            print("Error! 讀取ini設定檔發生錯誤! " + self.ini_full_path)
            str(e)
            raise

    def read_section_config(self, section):
        try:
            config_lh = configparser.ConfigParser()
            file_ini_lh = open(self.ini_full_path, 'r', encoding=self.ini_format)
            config_lh.read_file(file_ini_lh)
            file_ini_lh.close()
            single_section = config_lh.items(section)

            section_dict = OrderedDict()
            for item in single_section:
                section_dict[item[0]] = item[1]
            return section_dict
        except Exception as e:
            print("Error! 讀取ini設定檔發生錯誤! " + self.ini_full_path)
            str(e)
            raise

    def write_config(self, sections, ini_dict):
        try:
            config_lh = configparser.ConfigParser()
            config_lh.optionxform = str
            file_ini_lh = open(self.ini_full_path, 'r', encoding=self.ini_format)
            config_lh.read_file(file_ini_lh)
            file_ini_lh.close()

            for key, value in ini_dict.items():
                config_lh.set(sections, key, value)
            file_ini_lh = open(self.ini_full_path, 'w', encoding=self.ini_format)
            config_lh.write(file_ini_lh)
            file_ini_lh.close()
        except Exception as e:
            print("Error! 寫入ini設定檔發生錯誤! " + self.ini_full_path)
            str(e)
            raise


def get_interval_mon(start_datetime, end_datetime):
    return (end_datetime.year - start_datetime.year) * 12 + (end_datetime.month - start_datetime.month)


def encode_date_to_str(start_datetime, end_datetime):
    interval = get_interval_mon(start_datetime, end_datetime)
    encode_str = '{}_{}_{}'.format(start_datetime.month, end_datetime.month, interval)
    return encode_str


def decode_str_to_datetime(encode_str):
    decode_str = encode_str.split('_')
    return decode_str


# def random_sleep():
#     sleep(settings.SLEEP_SEC_BASE + random.uniform(settings.SLEEP_RANDOM_DELAY_SEC_MIN,
#                                                    settings.SLEEP_RANDOM_DELAY_SEC_MAX))
#

def get_this_year_start_date():
    return str_to_datetime(str(datetime.today().year)+'/1/1')


def daytime_to_sec(start, end):
    date_start = cover_str_to_datetime(start)
    date_end = cover_str_to_datetime(end)

    date = date_end - date_start

    # 8600 = 1day sec = 24*60*60
    return date.days * 86400


def cover_str_to_datetime(dateime_input):
    if type(dateime_input) == datetime:
        datetime_out = dateime_input
    else:
        datetime_out = str_to_datetime(dateime_input)

    return datetime_out


def roc_to_ad_str(s):
    return str(int(s) + 1911)


def str_to_datetime(s):
    datetime_str = s.replace('-', '/')
    date_time_h = datetime.strptime(datetime_str, "%Y/%m/%d")
    return date_time_h


def get_all_day(begin_date, end_date):
    datetime_list = []
    begin_datetime = cover_str_to_datetime(begin_date)
    end_datetime = cover_str_to_datetime(end_date)

    while begin_datetime <= end_datetime:
        datetime_list.append(begin_datetime)
        begin_datetime += timedelta(days=1)
    return datetime_list

# def round_v2(num, decimal):
#     num = np.round(num, decimal)
#     num = float(num)
#     return num
