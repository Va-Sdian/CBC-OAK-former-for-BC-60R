import customtkinter
import tkinter as tk
import csv
from docxtpl import DocxTemplate, RichText
from datetime import datetime
from tkinterdnd2 import DND_FILES, TkinterDnD
from tkinter import messagebox
from tkinter.filedialog import askdirectory
from tkinter import Menu
import os
import sys
import subprocess

# Данные для формирования имени файла и путь его сохранения
oak_former_name: str
oak_former_species_prefix: str
oak_former_client: str
oak_former_patient: str
chosen_directory = None

# Переменные для хранения пути загружаемого файла excel и отслеживание, обновлялся ли он при перезаполнении шаблона
excel_file = None
previous_loaded_file = None
currently_loaded_file_is_refreshed: bool
# Путь к шаблону в формате docx
doc = DocxTemplate("_internal/empty_oak_template.docx")
# Стандартный цвет заполнения переменных в файле docx (чёрный)
# text_color = '#000000'

mchc_high_value: int


def button_callback():
    # Берём значение ввода и заполняем контекст
    context['doctor'] = doctor_entry.get()
    doc.render(context)

    global chosen_directory, excel_file

    # Если файл Excel выбран, сохраняем файл, иначе — в текущую директорию
    directory = chosen_directory if chosen_directory and excel_file else os.getcwd()
    save_file(directory, check_open_folder_var.get())


def save_file(directory, check_open_folder):
    if not directory:
        print("Директория не выбрана.")
        return

    # Создаём путь к файлу и сохраняем
    file_path = os.path.join(directory, oak_former_name)
    doc.save(file_path)
    print(f"Файл сохранён: {file_path}")

    # Открываем директорию, если требуется
    if check_open_folder:
        open_folder(directory)


def open_folder(path):
    """
    Открытие директории в системном файловом менеджере.
    """
    if sys.platform == 'win32':
        os.startfile(path)
    elif sys.platform == 'darwin':  # MacOS
        subprocess.Popen(['open', path])
    else:  # Предполагаем использование Linux
        subprocess.Popen(['xdg-open', path])


def on_drop(event):  # Функция регистрации пути файла, который был перетянут и отпущен на label в окне программы
    try:
        # Преобразование строк события в список файлов с очисткой от лишних символов
        files = [file_path.replace('{', '').replace('}', '') for file_path in app.tk.splitlist(event.data) if file_path]
        if files:
            global excel_file, previous_loaded_file, currently_loaded_file_is_refreshed
            first_file = files[0]  # Работаем только с первым файлом из перетаскиваемых

            currently_loaded_file_is_refreshed = excel_file != first_file
            previous_loaded_file = None if excel_file is None else excel_file
            excel_file = first_file

            open_excel_and_load_data(excel_file)

            file_name = os.path.basename(first_file)
            # Обновляем текст label, используя имя файла
            label.configure(
                text=f"Перетащите файл сюда. \nТекущий файл: \n{file_name}\n{oak_former_client} {oak_former_patient}",
                font=('Helvetica', 12)
            )

    except Exception as e:
        print(f"Ошибка: {e}")


def mchc_error_message_box(_mchc, _mchc_high_value, _hct, _hgb):  # Всплывающее предупреждение, если MCHC выше нормы
    if currently_loaded_file_is_refreshed:
        if int(_mchc) > _mchc_high_value and _hct * 3 < _hgb:
            messagebox.showinfo('MCHC за референсами!',
                                f'MCHC:{_mchc} — за пределами референсного значения: {_mchc_high_value}!')
        elif int(_mchc) > _mchc_high_value and _hct * 3 > _hgb:
            messagebox.showinfo('MCHC за референсами!',
                                f'MCHC:{_mchc} — за пределами референсного значения: {_mchc_high_value}! '
                                f'\n HCT {_hct} * 3 = {float(_hct) * 3} < HGB {_hgb}')


def choose_save_directory():  # Спрашиваем, куда сохранять результаты
    global chosen_directory
    chosen_directory = askdirectory()  # Открытие диалога выбора директории


def checkbox_event_arrows():  # Функция перезаполнения бланка, если был нажат чекбокс про выставление стрелочек
    # print("checkbox toggled, current value:", check_arrows_var.get())
    # При простановке галочки перезаписываются все значения, и проставляются галочки
    global excel_file, currently_loaded_file_is_refreshed
    if excel_file is not None:
        currently_loaded_file_is_refreshed = False  # Нужно, чтобы не всплывало повторно окно предупреждения MCHC
        open_excel_and_load_data(excel_file)


# Инициация окна программы на TkinterDnd
app = TkinterDnD.Tk()
app.configure(bg='gray')  # Установка цвета фона главного окна
app.title('Автозаполнение шаблона ОКА')
app.geometry("400x380")
app.resizable(False, False)

# Создаём меню
check_open_folder_var = tk.BooleanVar(value=False)
menu = Menu(app)
new_item = Menu(menu, tearoff=False)
app.config(menu=menu)
new_item.add_command(label='Сохранить куда...', command=choose_save_directory)
new_item.add_checkbutton(label='Открыть директорию после сохранения', onvalue=True, offvalue=False,
                         variable=check_open_folder_var)
menu.add_cascade(label='Настройки', menu=new_item)

# Создаём виджет Label из библиотеки customtkinter
label = customtkinter.CTkLabel(app, text="Перетащите файл сюда", fg_color="white", corner_radius=20,
                               font=('Helvetica', 18), width=400, height=100)
label.pack(padx=10, pady=10)

# Регистрируем виджет label как целевой объект для операций перетаскивания файлов
label.drop_target_register(DND_FILES)
label.dnd_bind('<<Drop>>', on_drop)

# Создаём чекбокс для выбора функции проставления стрелочек
check_arrows_var = customtkinter.BooleanVar(value=False)
checkbox = customtkinter.CTkCheckBox(app, text="Проставлять ↑ ↓ при выходе за референсы", command=checkbox_event_arrows,
                                     variable=check_arrows_var, onvalue=True, offvalue=False)
checkbox.pack(pady=15)

# Создаём чекбокс для выбора функции окрашивания в красный цвет значений, выходящих за референсы
check_colored_var = customtkinter.BooleanVar(value=False)
checkbox_colored = customtkinter.CTkCheckBox(app, text=f"Менять цвет при выходе за референсы",
                                             command=checkbox_event_arrows, variable=check_colored_var,
                                             onvalue=True, offvalue=False)
checkbox_colored.pack(pady=0)

doctor_label = customtkinter.CTkLabel(app, text='', font=('Helvetica', 24))
doctor_label.pack(padx=20, pady=10)

doctor_entry = customtkinter.CTkEntry(app, placeholder_text='Введите имя врача', font=('Helvetica', 14))
doctor_entry.pack(padx=20, pady=10)

button_fill_template = customtkinter.CTkButton(app, text="Заполнить бланк", command=button_callback)
button_fill_template.pack(padx=20, pady=20)

# Делаем словарь для хранения округлённых процентов лейкоцитарной формулы, и задаём функцию по его округлению
# global met_perc, bond_perc, seg_perc, lym_perc, mon_perc, eos_perc, bas_perc
leuko_percent_values = {
    "met_perc": 0,
    "bond_perc": 0,
    "seg_perc": 0,
    "lym_perc": 0,
    "mon_perc": 0,
    "eos_perc": 0,
    "bas_perc": 0
}


def adjust_percentages(values):
    # Инициализация словарей для хранения целых значений и потерь от округления
    global leuko_percent_values
    int_values = {}
    losses = {}

    # Заполнение словаря целыми значениями и расчёт потерь от округления
    for k, v in values.items():
        int_value = int(v)  # Получаем целую часть числа
        int_values[k] = int_value  # Заносим целое значение в словарь
        losses[k] = v - int_value  # Рассчитываем потерю из-за округления и заносим в словарь

    # Вычисление общей суммы целых частей
    total = sum(int_values.values())

    # Вычисление "недостающей" суммы до 100
    difference = 100 - total

    # Сортировка элементов по значению потерь от округления в убывающем порядке
    sorted_losses = sorted(losses.items(), key=lambda x: x[1], reverse=True)

    # Распределение недостающих процентов по элементам с наибольшими потерями
    for i in range(difference):
        int_values[sorted_losses[i][0]] += 1  # Добавляем 1 к проценту для элемента

    return int_values


# Подготавливаем словарь для хранения значений, которые будут вставлены в шаблон
context = {
    'owner': '',
    'patient_id': '',
    'species': '',
    'date': '',
    'patient': '',

    'rbc': '',
    'hct': '',
    'hgb': '',
    'mcv': '',
    'mch': '',
    'mchc': '',
    'rdw_cv': '',
    'nrbc': '',
    'ret': '',
    'rhe': '',
    'plt': '',
    'mpv': '',
    'pct': '',
    'wbc': '',

    'rbc_ref': '',
    'hct_ref': '',
    'hgb_ref': '',
    'mcv_ref': '',
    'mch_ref': '',
    'mchc_ref': '',
    'rdw_cv_ref': '',
    'nrbc_ref': '',
    'ret_ref': '',
    'rhe_ref': '',
    'plt_ref': '',
    'mpv_ref': '',
    'pct_ref': '',
    'wbc_ref': '',

    'met_abs': '',
    'bond_abs': '',
    'seg_abs': '',
    'lym_abs': '',
    'mon_abs': '',
    'eos_abs': '',
    'bas_abs': '',

    'met_a_ref': '',
    'bond_a_ref ': '',
    'seg_a_ref ': '',
    'lym_a_ref ': '',
    'mon_a_ref ': '',
    'eos_a_ref ': '',
    'bas_a_ref ': '',

    'met_perc': '',
    'bond_perc': '',
    'seg_perc': '',
    'lym_perc': '',
    'mon_perc': '',
    'eos_perc': '',
    'bas_perc': '',

    'met_p_ref': '',
    'bond_p_ref': '',
    'seg_p_ref': '',
    'lym_p_ref': '',
    'mon_p_ref': '',
    'eos_p_ref': '',
    'bas_p_ref': '',

    'comment': '',
    'doctor': ''
}

# Зададим пустые переменные для референсных значений

rbc_ref: str = '-'
hct_ref: str = '-'
hgb_ref: str = '-'
mcv_ref: str = '-'
mch_ref: str = '-'
mchc_ref: str = '-'
rdw_cv_ref: str = '-'
nrbc_ref: str = '-'
ret_ref: str = '-'
rhe_ref: str = '-'
plt_ref: str = '-'
mpv_ref: str = '-'
pct_ref: str = '-'
wbc_ref: str = '-'

met_abs: str = '-'
bond_abs: str = '-'
seg_abs: str = '-'
lym_abs: str = '-'
mon_abs: str = '-'
eos_abs: str = '-'
bas_abs: str = '-'

met_a_ref: str = '-'
bond_a_ref: str = '-'
seg_a_ref: str = '-'
lym_a_ref: str = '-'
mon_a_ref: str = '-'
eos_a_ref: str = '-'
bas_a_ref: str = '-'

met_p_ref: str = '-'
bond_p_ref: str = '-'
seg_p_ref: str = '-'
lym_p_ref: str = '-'
mon_p_ref: str = '-'
eos_p_ref: str = '-'
bas_p_ref: str = '-'

# Для пересчёта абсолютных значений
wbc: float
met_perc: float
bond_perc: float
seg_perc: float
lym_perc: float
mon_perc: float
eos_perc: float
bas_perc: float


def species_references(_species):
    # Объявляем, что работаем с глобальными переменными
    global rbc_ref
    global hct_ref
    global hgb_ref
    global mcv_ref
    global mch_ref
    global mchc_ref
    global rdw_cv_ref
    global nrbc_ref
    global ret_ref
    global rhe_ref
    global plt_ref
    global mpv_ref
    global pct_ref
    global wbc_ref

    global met_a_ref
    global bond_a_ref
    global seg_a_ref
    global lym_a_ref
    global mon_a_ref
    global eos_a_ref
    global bas_a_ref

    global met_p_ref
    global bond_p_ref
    global seg_p_ref
    global lym_p_ref
    global mon_p_ref
    global eos_p_ref
    global bas_p_ref

    global oak_former_species_prefix
    global mchc_high_value

    if _species == 'Пес':
        oak_former_species_prefix = 'с.'
        mchc_high_value = 379
        rbc_ref = '5.4-8.9'
        hct_ref = '37-61'
        hgb_ref = '120-200'
        mcv_ref = '63-74'
        mch_ref = '21-27'
        mchc_ref = '320-379'
        rdw_cv_ref = '11-18'
        nrbc_ref = ''
        ret_ref = '10-110'
        rhe_ref = '22.3-29.6'
        plt_ref = '140-480'
        mpv_ref = '8-14'
        pct_ref = '0.14-0.46'
        wbc_ref = '5.05-16.8'

        met_a_ref = '0'
        bond_a_ref = '0-0.3'
        seg_a_ref = '3.0-11.5'
        lym_a_ref = '1.0-5.0'
        mon_a_ref = '0-1.2'
        eos_a_ref = '0.1-1.2'
        bas_a_ref = '0-0.1'

        met_p_ref = '0'
        bond_p_ref = '0-3'
        seg_p_ref = '45-75'
        lym_p_ref = '21-40'
        mon_p_ref = '3-9'
        eos_p_ref = '0-6'
        bas_p_ref = '0'

    elif _species == 'Кот':
        oak_former_species_prefix = 'к.'
        mchc_high_value = 358

        rbc_ref = '5.0-12.0'
        hct_ref = '28-48'
        hgb_ref = '80-150'
        mcv_ref = '39-53'
        mch_ref = '13-20'
        mchc_ref = '281-358'
        rdw_cv_ref = '15-25'
        nrbc_ref = ''
        ret_ref = '3-50'
        rhe_ref = '13.2-20.8'
        plt_ref = '151-600'
        mpv_ref = '11-21'
        pct_ref = '0.17-0.86'
        wbc_ref = '2.87-17.0'

        met_a_ref = '0'
        bond_a_ref = '0-0.3'
        seg_a_ref = '2.3-12.5'
        lym_a_ref = '0.92-7.5'
        mon_a_ref = '0.05-0.8'
        eos_a_ref = '0.1-1.5'
        bas_a_ref = '0-0.26'

        met_p_ref = '0'
        bond_p_ref = '0-3'
        seg_p_ref = '35-70'
        lym_p_ref = '25-55'
        mon_p_ref = '3-9'
        eos_p_ref = '0-5'
        bas_p_ref = '0-2'

    elif _species == 'Кролик':
        oak_former_species_prefix = 'кр.'
        mchc_high_value = 340

        rbc_ref = '4.15-6.50'
        hct_ref = '25-42'
        hgb_ref = '85-165'
        mcv_ref = '60-80'
        mch_ref = '17-24'
        mchc_ref = '270-340'
        rdw_cv_ref = '12.5-26.0'
        nrbc_ref = ''
        ret_ref = '40-550'
        rhe_ref = '15-24'
        plt_ref = '200-800'
        mpv_ref = '4.3-7.8'
        pct_ref = '0.1-0.56'
        wbc_ref = '3.0-12.5'

        met_a_ref = '0'
        bond_a_ref = '0'
        seg_a_ref = '1.02-5.95'
        lym_a_ref = '1.25-6.35'
        mon_a_ref = '0-1.85'
        eos_a_ref = '0-0.18'
        bas_a_ref = '0-0.65'

        met_p_ref = '0'
        bond_p_ref = '0'
        seg_p_ref = '24-62'
        lym_p_ref = '15-64'
        mon_p_ref = '0-20'
        eos_p_ref = '0-4'
        bas_p_ref = '0-8'

    elif currently_loaded_file_is_refreshed:
        oak_former_species_prefix = str(_species)
        mchc_high_value = 379  # Случайное значение
        messagebox.showinfo('Незапрограммированный вид!',
                            'Референсные значения для данного вида животного не установлены!')


def to_fixed(num_obj, digits=0):
    return f"{num_obj:.{digits}f}"


# Вынужден был добавить функцию переработки значений, из-за неправильного заполнения при 0 и 0.0 и None
def to_str_converting_float_to_int_if_possible(value):
    if isinstance(value, int):
        # Для int просто преобразуем в строку
        return str(value)
    elif isinstance(value, float):
        # Для float проверяем, равна ли дробная часть нулю
        if value.is_integer():
            return str(int(value))
        else:
            return str(value)
    elif isinstance(value, str):
        try:
            float_value = float(value)
            # Дополнительная проверка на случай, если строка представляет float
            if float_value.is_integer():
                # Если это целое число, возвращаем как int
                return str(int(float_value))
            else:
                # В противном случае возвращаем как float
                return str(float_value)
        except ValueError:
            # В случае неудачи с преобразованием строки в число
            return None
    else:
        # Если value не int, не float и не str, возвращаем None
        return None


def check_value_and_get_rich_text(value, ref_range, check_colored=None,
                                  check_arrows=None):
    """
    Определяет положение значения относительно заданного числового диапазона, и возвращает
    объект RichText с учитыванием параметров форматирования.

    value: числовое значение для проверки.
    ref_range: строка в формате 'min-max'.
    check_colored: булева переменная для изменения цвета текста при выходе за границы.
    check_arrows: состояние отображения стрелок, 'on' для включения их в текст.
    """
    # Проверка и получение актуальных значений внутри функции
    if check_colored is None:
        check_colored = check_colored_var.get()
    if check_arrows is None:
        check_arrows = check_arrows_var.get()

    # Преобразование значения value в строку с возможностью конвертации в int
    value_str = to_str_converting_float_to_int_if_possible(value)

    # Проверка на наличие референсного диапазона
    if not ref_range or ref_range == "0" or ref_range == "-":
        return RichText(value_str, color='#000000', bold=True)  # Черный текст, если диапазон не задан

    # Разделение строки на минимальное и максимальное значения и их преобразование
    min_value, max_value = ref_range.split('-')
    min_value, max_value = (float(x) if '.' in x else int(x) for x in (min_value, max_value))
    value_num = float(value_str)

    # Инициализация цвета текста
    text_color = '#000000'

    # Сравнение значения с диапазоном и определение цвета текста
    if value_num < min_value:
        text_color = '#FF0000' if check_colored else '#000000'
        value_str += ' ↓' if check_arrows else ''
    elif value_num > max_value:
        text_color = '#FF0000' if check_colored else '#000000'
        value_str += ' ↑' if check_arrows else ''
        # print(f'value {value_num} > max_value {max_value}, text_color = {text_color}, value_str = {value_str}')

    return RichText(value_str, color=text_color, bold=True)


# Функция абсолютных чисел должна запускаться только после округления процентов
def absolute_numbers(_met, _bond, _seg, _lym, _mon, _eos, _bas, _wbc):
    global met_perc, bond_perc, seg_perc, lym_perc, mon_perc, eos_perc, bas_perc, wbc, met_abs, bond_abs, seg_abs, \
        lym_abs, mon_abs, eos_abs, bas_abs
    met_abs = (wbc * _met) / 100
    bond_abs = (wbc * _bond) / 100

    seg_abs = to_fixed(((wbc * _seg) / 100), 2)
    lym_abs = to_fixed(((wbc * _lym) / 100), 2)
    mon_abs = to_fixed(((wbc * _mon) / 100), 2)
    eos_abs = to_fixed(((wbc * _eos) / 100), 2)
    bas_abs = to_fixed(((wbc * _bas) / 100), 2)

# По неизвестным причинам геманализатор стал выводить строки в другом формате.
# Поэтому добавил функцию определения формата времени.
def parse_datetime(datetime_str):
    formats = [
        '%d.%m.%Y %H:%M:%S',
        '%Y/%m/%d %H:%M:%S'
    ]

    for fmt in formats:
        try:
            datetime_obj = datetime.strptime(datetime_str, fmt)
            return datetime_obj
        except ValueError:
            continue

    raise ValueError("Неизвестный формат даты и времени: {}".format(datetime_str))


# Открытие файла ".csv" таблицы с гемоанализатора BC-60R с результатами как excel файла
def open_excel_and_load_data(_excel_file):
    if ".csv" in _excel_file:
        with open(_excel_file, newline='', encoding='utf-16') as csvfile:
            csv_reader = csv.DictReader(csvfile, delimiter='\t')  # Указываем разделитель

            for row in csv_reader:
                # Для каждой строки 'row' является словарем, где ключи - это названия колонок,
                # А значения - это соответствующие значения в этой строке.

                # Создадим "очищенный" словарь для текущей строки, убрав лишние табуляции
                # Это необязательно, если вам нормально работать с первоначальными ключами
                clean_row = {key.strip('\t'): value for key, value in row.items()}

                datetime_str = clean_row['Время анализа'].lstrip()
                datetime_obj = parse_datetime(datetime_str)
                date_str = datetime_obj.strftime('%d.%m.%Y')

                species = clean_row['Вид'].lstrip()
                species_references(species)

                global wbc, met_perc, bond_perc, seg_perc, lym_perc, mon_perc, eos_perc, bas_perc, leuko_percent_values

                # Сначала извлекаем значения из таблицы
                wbc = float(clean_row['WBC'].lstrip())
                met_perc = 0
                bond_perc = 0
                seg_perc = float(clean_row['Neu%'].lstrip())
                lym_perc = float(clean_row['Lym%'].lstrip())
                mon_perc = float(clean_row['Mon%'].lstrip())
                eos_perc = float(clean_row['Eos%'].lstrip())
                bas_perc = float(clean_row['Bas%'].lstrip())

                # Присваиваем неокруглённые значения словарю
                leuko_percent_values['met_perc'] = met_perc
                leuko_percent_values['bond_perc'] = bond_perc
                leuko_percent_values['seg_perc'] = seg_perc
                leuko_percent_values['lym_perc'] = lym_perc
                leuko_percent_values['mon_perc'] = mon_perc
                leuko_percent_values['eos_perc'] = eos_perc
                leuko_percent_values['bas_perc'] = bas_perc

                # Округляем значения
                adjusted_values = adjust_percentages(leuko_percent_values)

                # Переназначаем старые значения, чтобы мне не переписывать готовый код, оперирующий ими
                met_perc = adjusted_values['met_perc']
                bond_perc = adjusted_values['bond_perc']
                seg_perc = adjusted_values['seg_perc']
                lym_perc = adjusted_values['lym_perc']
                mon_perc = adjusted_values['mon_perc']
                eos_perc = adjusted_values['eos_perc']
                bas_perc = adjusted_values['bas_perc']

                # Пересчитываем абсолютные значения
                absolute_numbers(met_perc, bond_perc, seg_perc, lym_perc, mon_perc, eos_perc, bas_perc, wbc)

                # Даём название файлу и даём фамилию для обновления label в программе
                global oak_former_name, oak_former_client, oak_former_patient
                oak_former_client = clean_row['Клиент'].lstrip()
                oak_former_patient = clean_row['Пациент'].lstrip()
                oak_former_name = (oak_former_species_prefix + ' ' + oak_former_patient + ' ' + clean_row[
                    'ID пациента'].lstrip() + " ОКА" + '.docx')

                # Проверяем MCHC
                global mchc_high_value, currently_loaded_file_is_refreshed
                mchc = clean_row['MCHC'].lstrip()
                hct = clean_row['HCT'].lstrip()
                hgb = clean_row['HGB'].lstrip()
                # Делаем проверку, что был загружен новый файл, а не снимали и ставили галочку на настройках галочек
                if currently_loaded_file_is_refreshed:
                    mchc_error_message_box(mchc, mchc_high_value, hct, hgb)

                # Дальше можно обращаться к данным в 'clean_row' используя названия колонок
                # Например, допустим вам нужен ID пробы и Конт.
                # patient_id = row['ID пациента'].strip()

                if 'Пациент' in clean_row and 'Lym%' in clean_row:
                    context['owner'] = clean_row['Клиент'].lstrip()
                    context['patient_id'] = clean_row['ID пациента'].lstrip()

                    # Названия видов животных в анализаторе установлены некорректно, исправляем это вручную
                    if clean_row['Вид'].lstrip() == 'Пес':
                        context['species'] = 'Собака'
                    elif clean_row['Вид'].lstrip() == 'Кот':
                        context['species'] = 'Кошка'
                    else:
                        context['species'] = clean_row['Вид'].lstrip()
                    context['date'] = date_str
                    context['patient'] = clean_row['Пациент'].lstrip()

                    context['rbc_ref'] = rbc_ref
                    context['hct_ref'] = hct_ref
                    context['hgb_ref'] = hgb_ref
                    context['mcv_ref'] = mcv_ref
                    context['mch_ref'] = mch_ref
                    context['mchc_ref'] = mchc_ref
                    context['rdw_cv_ref'] = rdw_cv_ref
                    context['nrbc_ref'] = nrbc_ref
                    context['ret_ref'] = ret_ref
                    context['rhe_ref'] = rhe_ref
                    context['plt_ref'] = plt_ref
                    context['mpv_ref'] = mpv_ref
                    context['pct_ref'] = pct_ref
                    context['wbc_ref'] = wbc_ref

                    context['met_a_ref'] = met_a_ref
                    context['bond_a_ref'] = bond_a_ref
                    context['seg_a_ref'] = seg_a_ref
                    context['lym_a_ref'] = lym_a_ref
                    context['mon_a_ref'] = mon_a_ref
                    context['eos_a_ref'] = eos_a_ref
                    context['bas_a_ref'] = bas_a_ref

                    context['met_p_ref'] = met_p_ref
                    context['bond_p_ref'] = bond_p_ref
                    context['seg_p_ref'] = seg_p_ref
                    context['lym_p_ref'] = lym_p_ref
                    context['mon_p_ref'] = mon_p_ref
                    context['eos_p_ref'] = eos_p_ref
                    context['bas_p_ref'] = bas_p_ref

                    context['rbc'] = check_value_and_get_rich_text(clean_row['RBC'].lstrip(), rbc_ref)
                    context['hct'] = check_value_and_get_rich_text(clean_row['HCT'].lstrip(), hct_ref)
                    context['hgb'] = check_value_and_get_rich_text(clean_row['HGB'].lstrip(), hgb_ref)
                    context['mcv'] = check_value_and_get_rich_text(clean_row['MCV'].lstrip(), mcv_ref)
                    context['mch'] = check_value_and_get_rich_text(clean_row['MCH'].lstrip(), mch_ref)
                    context['mchc'] = check_value_and_get_rich_text(clean_row['MCHC'].lstrip(), mchc_ref)
                    context['rdw_cv'] = check_value_and_get_rich_text(clean_row['RDW-CV'].lstrip(), rdw_cv_ref)
                    context['nrbc'] = check_value_and_get_rich_text('0', nrbc_ref)
                    context['ret'] = check_value_and_get_rich_text(clean_row['RET#'].lstrip(), ret_ref)
                    context['rhe'] = check_value_and_get_rich_text(clean_row['RHE'].lstrip(), rhe_ref)
                    context['plt'] = check_value_and_get_rich_text(clean_row['PLT'].lstrip(), plt_ref)
                    context['mpv'] = check_value_and_get_rich_text(clean_row['MPV'].lstrip(), mpv_ref)
                    context['pct'] = check_value_and_get_rich_text(clean_row['PCT'].lstrip(), pct_ref)
                    context['wbc'] = check_value_and_get_rich_text(clean_row['WBC'].lstrip(), wbc_ref)

                    context['met_abs'] = check_value_and_get_rich_text(met_abs, met_a_ref)
                    context['bond_abs'] = check_value_and_get_rich_text(bond_abs, bond_a_ref)
                    context['seg_abs'] = check_value_and_get_rich_text(seg_abs, seg_a_ref)
                    context['lym_abs'] = check_value_and_get_rich_text(lym_abs, lym_a_ref)
                    context['mon_abs'] = check_value_and_get_rich_text(mon_abs, mon_a_ref)
                    context['eos_abs'] = check_value_and_get_rich_text(eos_abs, eos_a_ref)
                    context['bas_abs'] = check_value_and_get_rich_text(bas_abs, bas_a_ref)

                    context['met_perc'] = check_value_and_get_rich_text(met_perc, met_p_ref)
                    context['bond_perc'] = check_value_and_get_rich_text(bond_perc, bond_p_ref)
                    context['seg_perc'] = check_value_and_get_rich_text(seg_perc, seg_p_ref)
                    context['lym_perc'] = check_value_and_get_rich_text(lym_perc, lym_p_ref)
                    context['mon_perc'] = check_value_and_get_rich_text(mon_perc, mon_p_ref)
                    context['eos_perc'] = check_value_and_get_rich_text(eos_perc, eos_p_ref)
                    context['bas_perc'] = check_value_and_get_rich_text(bas_perc, bas_p_ref)

                    # context['comment'] = clean_row[].lstrip() Можно добавить комментарий при необходимости,
                    # но легче индивидуально заполнять в готовом файле

                    # Мы обрабатываем только первую строку, подходящую под условия
                    break


app.mainloop()
