import tkinter as tk
from tkinter import *
from tkinter import ttk, messagebox
import qrcode
from PIL import Image, ImageTk, ImageWin, ImageDraw, ImageFont
import win32print
import win32ui
import re
import os
from dotenv import load_dotenv
import requests
import win32gui
import win32api

# Размеры наклеек
templates = [
    {'name': 'Белая 3х3', 'type': 'imei', 'params': (180, 10, 30)},
    {'name': 'Arnavi A4', 'type': 'imei', 'params': (240, 0, 23)},
    {'name': 'АО24 Integral', 'type': 'imei', 'params': (128, 200, 23)},
    {'name': 'АО24 L-Series', 'type': 'imei', 'params': (65, 32, 19)},
    {'name': 'АО24 Beacon', 'type': 'imei', 'params': (45, 170, 28)},
    {'name': 'BTS, BAS', 'type': 'mac', 'params': (180, 0, 30)},
    {'name': 'ESM-P', 'type': 'sn', 'params': (370, 230, 28)},
    {'name': 'QR Copy', 'type': 'copy', 'params': (180, 10, 26)},
    {'name': 'Тех. наклейка', 'type': 'teh', 'params': (20, 0, 55)},
    {'name': 'ДУТ BLS', 'type': 'sn', 'params': (180, 120, 30)},
]

last_qr_code_path = 'qr_code.png'


def setenglang():
    window_handle = win32gui.GetForegroundWindow()
    result = win32api.SendMessage(window_handle, 0x0050, 0, 0x04090409)
    return result


def get_token():
    headers = {
        'accept': 'application/json',
        'Content-Type': 'application/x-www-form-urlencoded',
    }

    data = {
        'username': wbcfg_name,
        'password': wbcfg_pass,
    }

    print('Запрос Токена...')
    response = requests.post('http://ws.arusnavi.ru:8089/external/token', headers=headers, data=data)
    response_json = response.json()
    print('...Токен получен')
    return response_json['access_token']


def extract_imei(input_string):
    imei_pattern = r'\b\d{15}\b'
    imei_match = re.search(imei_pattern, input_string)
    if imei_match:
        return imei_match.group()
    else:
        return None


def extract_mac(raw_string):
    pattern = r'MAC:([0-9A-Fa-f]{12})'
    match = re.search(pattern, raw_string)
    splited_string = raw_string.split(';')
    if match:
        device_info = {'MAC': match.group(1), 'type': splited_string[1], 'ver': splited_string[3]}
        return device_info
    else:
        return None


def get_device_info(token, extracted_imei):
    headers = {
        'accept': 'application/json',
        'Authorization': f'Bearer {token}',
    }

    if extracted_imei:
        response = requests.get(f'http://ws.arusnavi.ru:8089/external-api/v1/devices/by-imei/{extracted_imei}',
                                headers=headers)
        print(response.json())
        device_info = {'id': str(response.json()['identity']),
                       'imei': str(response.json()['imei']),
                       'typeId': str(response.json()['deviceTypeId']),
                       # 'name': str(response.json()['deviceTypeId']),
                       }
        return device_info
    else:
        print('IMEI не найден')
        return None


def generate_qr(device_info):
    qr_data_dict = device_info
    qr_data = f"{qr_data_dict['id']};{qr_data_dict['imei']};{qr_data_dict['typeId']}"
    qr = qrcode.QRCode(version=1, box_size=10, border=2)
    qr.add_data(qr_data)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    img.save('qr_code_raw.png')

    template = Image.new('RGB', (360, 360), 'white')
    qr_code = Image.open('qr_code_raw.png')
    qr_code.convert('RGBA')
    template.paste(qr_code, (40, 47))
    draw = ImageDraw.Draw(template)
    font = ImageFont.truetype('Bebas_Neue.ttf', 46)
    draw.text((131, 20), f'{qr_data_dict["id"]}', font=font, fill='black')
    draw.text((49, 325), f'{qr_data_dict["imei"]}', font=font, fill='black')
    template.save(last_qr_code_path)
    os.remove('qr_code_raw.png')


def generate_teh_qr(device_info, offset):
    qr_data_dict = device_info
    qr_data = f"{qr_data_dict['imei']};{qr_data_dict['id']};{qr_data_dict['typeId']}"
    qr = qrcode.QRCode(version=1, box_size=4, border=2)
    qr.add_data(qr_data)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    img.save('qr_code_raw.png')

    template = Image.new('RGB', (360, 360), 'white')
    qr_code = Image.open('qr_code_raw.png')
    qr_code.convert('RGBA')
    template.paste(qr_code, (15, 120))
    draw = ImageDraw.Draw(template)
    font = ImageFont.truetype('Bebas_Neue.ttf', 30)
    if offset == '':
        offset = 220
    name_coords = (int(offset), 130)
    draw.text(name_coords, name_entry.get(), font=font, fill='black')
    draw.text((175, 165), f'ID {qr_data_dict["id"]}', font=font, fill='black')
    draw.text((135, 200), f'{qr_data_dict["imei"]}', font=font, fill='black')
    template.save(last_qr_code_path)
    os.remove('qr_code_raw.png')


def generate_qr_mac(device_mac):
    qr = qrcode.QRCode(version=1, box_size=11, border=2)
    qr.add_data(device_mac['MAC'])
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    img.save('qr_code_raw.png')

    template = Image.new('RGB', (360, 360), 'white')
    qr_code = Image.open('qr_code_raw.png')
    qr_code.convert('RGBA')
    template.paste(qr_code, (40, 70))
    draw = ImageDraw.Draw(template)
    font = ImageFont.truetype('Bebas_Neue.ttf', 46)
    font_bottom = ImageFont.truetype('Bebas_Neue.ttf', 30)
    draw.text((70, 40), f'{device_mac["MAC"]}', font=font, fill='black')
    draw.text((145, 330), f'{device_mac["type"]} {device_mac["ver"]}', font=font_bottom, fill='black')
    template.save(last_qr_code_path)
    os.remove('qr_code_raw.png')


def generate_qr_id(device_id):
    qr = qrcode.QRCode(version=1, box_size=11, border=2)
    qr.add_data(device_id)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    img.save('qr_code_raw.png')

    template = Image.new('RGB', (360, 360), 'white')
    qr_code = Image.open('qr_code_raw.png')
    qr_code.convert('RGBA')
    template.paste(qr_code, (40, 70))
    draw = ImageDraw.Draw(template)
    font = ImageFont.truetype('Bebas_Neue.ttf', 46)
    draw.text((70, 40), f'{device_id}', font=font, fill='black')
    template.save(last_qr_code_path)
    os.remove('qr_code_raw.png')


def generate_qr_copy(input_string):
    qr = qrcode.QRCode(version=1, box_size=11, border=2)
    qr.add_data(input_string)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    img.save('qr_code_raw.png')
    imei = extract_imei(input_string)
    if imei:
        device_id = imei
    else:
        device_id = input_string
    template = Image.new('RGB', (360, 360), 'white')
    qr_code = Image.open('qr_code_raw.png')
    qr_code.convert('RGBA')
    template.paste(qr_code, (40, 70))
    draw = ImageDraw.Draw(template)
    font = ImageFont.truetype('Bebas_Neue.ttf', 46)
    draw.text((80, 40), f'{device_id}', font=font, fill='black')
    template.save(last_qr_code_path)
    os.remove('qr_code_raw.png')


def print_qr_code(pos_x=0, pos_y=0, size_mm=30):
    printer_name = 'ZDesigner ZD421-300dpi ZPL'
    hprinter = win32print.OpenPrinter(printer_name)

    try:
        # Создание устройства для печати
        hdc = win32ui.CreateDC()
        hdc.CreatePrinterDC(printer_name)
        hdc.StartDoc("QR Code Print")
        hdc.StartPage()

        # Загружаем изображение для печати
        bmp = Image.open(last_qr_code_path)
        if horizontal_var.get():
            bmp = bmp.rotate(270)
        dpi = 300
        size_px = int(size_mm / 25.4 * dpi)
        bmp = bmp.resize((size_px, size_px))
        dib = ImageWin.Dib(bmp)

        # Получаем размеры изображения
        width, height = bmp.size

        # Рисуем изображение на контексте устройства
        dib.draw(hdc.GetHandleOutput(), (pos_x, pos_y, width + pos_x, height + pos_y))
        hdc.EndPage()
        hdc.EndDoc()
    finally:
        win32print.ClosePrinter(hprinter)


def create_gui(token):
    root = tk.Tk()
    root.title("QR Code Generator")
    root.resizable(width=False, height=False)

    for c in range(2):
        root.columnconfigure(index=c, weight=5)
    for r in range(2):
        root.rowconfigure(index=r, weight=9)

    # Строка ввода
    tk.Label(root, text="Сканируйте QR или введите IMEI:").grid(column=0, row=0, columnspan=3)
    raw_entry = tk.Entry(root, width=35)
    raw_entry.grid(column=0, row=1, columnspan=3)

    def coord_mod():
        current_var_params = eval(form_var.get())['params']
        spinbox_var1.set(value=current_var_params[0])
        spinbox_var2.set(value=current_var_params[1])
        spinbox_var3.set(value=current_var_params[2])
        return current_var_params

    def coord_read():
        return int(spinbox_var1.get()), int(spinbox_var2.get()), int(spinbox_var3.get())

    default_template = str(templates[0])
    form_var = StringVar(value=default_template)
    row = 2
    for point in templates:
        btn = ttk.Radiobutton(root, text=point['name'], variable=form_var, value=point, command=coord_mod)
        btn.config(width=15)
        btn.grid(column=0, row=row, sticky='w')
        row += 1

    params_dict = templates[0]
    params_tuple = params_dict['params']
    spinbox_var1, spinbox_var2, spinbox_var3 = (
        StringVar(value=params_tuple[0]), StringVar(value=params_tuple[1]), StringVar(value=params_tuple[2])
    )
    label1 = ttk.Label(text='X координата')
    label1.grid(column=1, row=2)
    label2 = ttk.Label(text='Y координата')
    label2.grid(column=1, row=4)
    label3 = ttk.Label(text='Масштаб')
    label3.grid(column=1, row=6)

    spinbox1 = ttk.Spinbox(from_=1, to=600.0, textvariable=spinbox_var1, width=15)
    spinbox1.grid(column=1, row=3)
    spinbox2 = ttk.Spinbox(from_=1, to=600.0, textvariable=spinbox_var2, width=15)
    spinbox2.grid(column=1, row=5)
    spinbox3 = ttk.Spinbox(from_=10, to=60.0, textvariable=spinbox_var3, width=15)
    spinbox3.grid(column=1, row=7)

    global horizontal_var, name_entry
    horizontal_var = BooleanVar(value=False)
    horizontal_btn = ttk.Radiobutton(root, text='Повернуть на 90°', variable=horizontal_var, value=True)
    horizontal_btn.grid(column=1, row=8)

    name_entry = tk.Entry(root)
    name_entry.grid(column=1, row=9)
    name_offset = tk.Entry(root)
    name_offset.grid(column=1, row=10)

    # Область для отображения QR-кода
    qr_display = tk.Label(root)
    qr_display.grid(column=0, row=row+1, columnspan=3)

    def display_qr_code():
        img = Image.open(last_qr_code_path)
        img.thumbnail((200, 200))
        img_tk = ImageTk.PhotoImage(img)
        qr_display.config(image=img_tk)
        qr_display.image = img_tk

    display_qr_code()

    def entry_change(*args):
        input_string = raw_entry.get()
        data_dict = eval(form_var.get())
        if input_string:
            if data_dict['type'] == 'imei':
                extracted_imei = extract_imei(input_string)
                if extracted_imei:
                    device_info = get_device_info(token, extracted_imei)
                    generate_qr(device_info)
                else:
                    messagebox.showerror("Ошибка", "Не удалось получить информацию об устройстве.")
                    return
            if data_dict['type'] == 'mac':
                extracted_mac = extract_mac(input_string)
                if extracted_mac:
                    generate_qr_mac(extracted_mac)
                else:
                    messagebox.showerror("Ошибка", "Не удалось получить информацию MAC-адрес устройства")
                    return
            if data_dict['type'] == 'sn':
                if input_string:
                    generate_qr_id(input_string)
                else:
                    messagebox.showerror("Ошибка", "Строка ввода пуста!")
                    return
            if data_dict['type'] == 'copy':
                if input_string:
                    generate_qr_copy(input_string)
                else:
                    messagebox.showerror("Ошибка", "Строка ввода пуста!")
                    return
            if data_dict['type'] == 'teh':
                extracted_imei = extract_imei(input_string)
                if extracted_imei:
                    device_info = get_device_info(token, extracted_imei)
                    generate_teh_qr(device_info, name_offset.get())
                else:
                    messagebox.showerror("Ошибка", "Не удалось получить информацию об устройстве.")
                    return
            sticker_template = coord_read()
            display_qr_code()
            print_qr_code(*sticker_template)
            raw_entry.delete(0, tk.END)
        else:
            messagebox.showerror("Ошибка", "Строка ввода пуста!")
            return

    raw_entry.bind("<Return>", entry_change)
    root.mainloop()


if __name__ == '__main__':
    setenglang()
    load_dotenv()
    wbcfg_name = os.getenv('WEBCONF_USER')
    wbcfg_pass = os.getenv('WEBCONF_PASS')
    token = get_token()
    create_gui(token)
