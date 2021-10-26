from InputData.wa_excel import GetDataTable
from InputData.robot_dict import region_dict
from BPM.wa_bpm import WaBPM
from Mail.send_mail import SendMail
from Report.get_report import GetReport
import win32com.client as win32
import os
from datetime import datetime
from win_err_bpm import Ui_MainWindow
from win_successfully_bpm import Ui_Main_Successfully


def log(log_str):
    """ Перехват всех исключений и запись в файл """

    path_file = os.getcwd()
    with open(path_file + "\\error_log.txt", "a", ) as file_log:
        print(log_str)
        file_log.writelines([datetime.now().strftime("%d-%m-%Y %H.%M.%S# "), '  ', log_str, '\n'])


def convert_time(start_date, end_date, start_reg_time):
    """ Преобразование времени в соответствии каждому региону для BPM """

    if start_reg_time in ("0:00", "1:00", "2:00", "3:00", "4:00", "5:00", "6:00"):
        start_date_for_bpm = end_date
        end_date_for_bpm = end_date
        start_time_for_bpm = start_reg_time
        end_time_for_bpm = "{}:00".format(int(start_reg_time.split(":")[0]) + 5)
    else:
        start_date_for_bpm = start_date
        start_time_for_bpm = start_reg_time
        if int(start_reg_time.split(":")[0]) + 5 >= 24:
            end_date_for_bpm = end_date
            end_time_for_bpm = "{}:00".format(int(start_reg_time.split(":")[0]) + 5 - 24)
        else:
            end_date_for_bpm = start_date
            end_time_for_bpm = "{}:00".format(int(start_reg_time.split(":")[0]) + 5)

    return [start_date_for_bpm, end_date_for_bpm, start_time_for_bpm, end_time_for_bpm]


def check_folders():
    """ Проверка существования файла с пролетами в папке с программой """

    try:
        path_in_folders = os.getcwd() + "\\" + \
                          [path for path in os.listdir(os.getcwd()) if "WFL_задания на SW расширения РРЛ" in path][0]
    except IndexError:
        path_in_folders = False
    return path_in_folders


def qet_user_path():
    """ Получение последнего файла с пролетами из папки загрузок """

    path_user = os.path.expanduser("~") + "\Downloads"

    list_files_sw = [s for s in os.listdir(path_user)
                     if os.path.isfile(os.path.join(path_user, s))]
    list_files_sw.sort(key=lambda s: os.path.getmtime(os.path.join(path_user, s)))
    path_file_sw = path_user + "\\" + list_files_sw[-1]

    return path_file_sw


def main():
    try:
        path_to_driver = os.getcwd() + "\chromedriver.exe"
        input_data_table = GetDataTable()

        '''Получение файла с работами и данных из шаблона'''
        responsible = input_data_table.get_responsible()
        executor = input_data_table.get_executor()
        start_date = input_data_table.get_start_date()
        end_date = input_data_table.get_end_date()
        get_reg_list = input_data_table.get_reg_list()
        print(responsible)
        print(executor)
        print(get_reg_list)

        if not check_folders():
            GetReport().get_report_run()
            path_file_sw = qet_user_path()
        else:
            path_file_sw = check_folders()

        print(path_file_sw)

        '''Начала цикла по обходу регионов и заведения работ'''
        for mr_name in get_reg_list:
            input_data_table_with_mr_name = GetDataTable(mr_name=mr_name, path_file_sw=path_file_sw)
            rrl_list_sw_file = "\n".join(input_data_table_with_mr_name.get_column_sw_file()[0])
            influence_list_sw_file = "\n".join(input_data_table_with_mr_name.get_column_sw_file()[1])
            print(rrl_list_sw_file)
            print(influence_list_sw_file)

            reg_name = region_dict[mr_name]['region_for_bpm']
            start_reg_time = region_dict[mr_name]['time'].split(":")[0]
            if start_reg_time in "00":
                start_reg_time = "0:00"
            elif start_reg_time in "01":
                start_reg_time = "1:00"
            else:
                start_reg_time = "{}:00".format(start_reg_time)
            print(reg_name)

            time_data_list = convert_time(
                start_date=start_date,
                end_date=end_date,
                start_reg_time=start_reg_time
            )
            start_date_for_bpm = time_data_list[0]
            end_date_for_bpm = time_data_list[1]
            start_time_for_bpm = time_data_list[2]
            end_time_for_bpm = time_data_list[3]

            print(start_date_for_bpm)
            print(end_date_for_bpm)
            print(start_time_for_bpm)
            print(end_time_for_bpm)

            '''Заведение работ'''

            text_wa = WaBPM(
                responsible=responsible,
                executor=executor,
                mr_name=mr_name,
                region_name=reg_name,
                start_date_for_bpm=start_date_for_bpm,
                end_date_for_bpm=end_date_for_bpm,
                start_time_for_bpm=start_time_for_bpm,
                end_time_for_bpm=end_time_for_bpm,
                path_to_driver=path_to_driver,
                rrl_list_sw_file=rrl_list_sw_file,
                influence_list_sw_file=influence_list_sw_file
            ).run_wa()

            '''Отправка письма'''

            SendMail(
                win32=win32,
                text_wa=text_wa,
                region=mr_name,
                start_work=start_date_for_bpm,
                start_time=start_time_for_bpm,
                rrl_list_sw_file=rrl_list_sw_file
            ).send_mail()

        '''Окно успешности'''

        Ui_Main_Successfully().run_win()

    except Exception as err_str:
        err = 'Ошибка: ' + str(err_str)
        log(err)
        Ui_MainWindow().run_win()


if __name__ == "__main__":
    main()
