import os
import sys
import subprocess
import configparser
import tkinter as tk
from tkinter import filedialog
from PIL import Image
import ffmpeg
import locale
import json
import shutil
import ctypes
from ctypes import *
import win32clipboard
import win32con
from cursesmenu import CursesMenu
from cursesmenu.items import FunctionItem, SubmenuItem, CommandItem

    
def get_bin_path():
    if getattr(sys, 'frozen', False):
        return os.path.join(sys._MEIPASS, 'bin')
    else:
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), 'bin')

script_dir = os.path.dirname(os.path.abspath(__file__))
bin_path = get_bin_path()
os.environ["PATH"] = bin_path + os.pathsep + os.environ["PATH"]

class DROPFILES(Structure):
    _fields_ = [
        ("pFiles", c_uint32),
        ("pt", c_long * 2),
        ("fNC", c_int),
        ("fWide", c_bool),
    ]

system_language = locale.getlocale()[0]
translations = {
    'ru_RU': {
        'm_welcome': "Добро пожаловать в скрипт для уменьшения размера видео/изображений/аудио для Discord",
        'm_start': "Начать",
        'm_program_func': "Функциональность",
        'm_info': "Инфо",
        'm_chopt': "Выберите пункт меню: ",
        'input_prompt': "Перетащите файл в консоль или введите путь к файлу: ",
        'file_not_found': "Файл не найден. Пожалуйста, проверьте путь и попробуйте снова.",
        'ffmpeg_not_found': "FFmpeg не найден в папке bin. Пожалуйста, убедитесь, что FFmpeg установлен в папке bin.",
        'compressing': "Сжатие файла...",
        'compressed': "Файл сжат и сохранен как: {}",
        'archive_warning': "Ты тупой?",
        'unsupported_format': "Неподдерживаемый формат файла: {}",
        'error_compressing': "Ошибка при сжатии файла: {}",
        'error_getting_duration': "Ошибка при получении длительности файла. Проверьте, что файл не поврежден и содержит видео или аудио поток.",
        'mkv_renamed': "Файл .mkv переименован в .mp4: {}",
        'select_output_folder': "Выберите папку для сохранения сжатых файлов",
        'copied_to_clipboard': "Файл скопирован в буфер обмена",
        'clipboard_error': "Ошибка при копировании в буфер обмена",
        'info_abt_program': "Функциональность программы заключается в том, что она сжимает ваши медиафайлы до размера меньше 10 МБ.\nПоддерживаемые форматы: mp4, mkv, jpg, jpeg, png, mp3.\nЯ не гарантирую 100% результата, но если у вас появились какие-то проблемы, пишите их в \033]8;;https://github.com/MAX1MDEV/messagecord/issues\033\\Issues\033]8;;\033\\.",
        'my_soc': "Код написан: MaximDev\nGithub: https://github.com/MAX1MDEV\nDiscord: https://discord.gg/TpApPfXrXG\nВеб-сайт: https://maximdev.ru\nРандомные steam и промо-ключи: https://maximdev.ru/RSPK",
        'go_to_menu': "\nНажмите Enter для возврата в меню..."
    },
    'en_US': {
        'm_welcome': 'Welcome to the video/image/audio downsizing script for Discord',
        'm_start': "Start",
        'm_program_func': "Functionality",
        'm_info': "Info",
        'm_chopt': "Select the menu item: ",
        'input_prompt': "Drag the file into the console or enter the file path: ",
        'file_not_found': "File not found. Please check the path and try again.",
        'ffmpeg_not_found': "FFmpeg not found in the bin folder. Please make sure FFmpeg is installed in the bin folder.",
        'compressing': "Compressing file...",
        'compressed': "File compressed and saved as: {}",
        'archive_warning': "Are u stupid?",
        'unsupported_format': "Unsupported file format: {}",
        'error_compressing': "Error compressing file: {}",
        'error_getting_duration': "Error getting file duration. Check if the file is not corrupted and contains a video or audio stream.",
        'mkv_renamed': "MKV file renamed to MP4: {}",
        'select_output_folder': "Select folder for saving compressed files",
        'copied_to_clipboard': "File copied to clipboard",
        'clipboard_error': "Error copying to clipboard",
        'info_abt_program': "The functionality of the program is to compress your media files to less than 10 MB.\nSupported formats: mp4, mkv, jpg, jpeg, png, mp3.\nI do not guarantee 100% results, but if you encounter any issues, please report them in \033]8;;https://github.com/MAX1MDEV/messagecord/issues\033\\Issues\033]8;;\033\\.",
        'my_soc': "Code written by: MaximDev\nGithub: https://github.com/MAX1MDEV\nDiscord: https://discord.gg/TpApPfXrXG\nWebSite: https://maximdev.ru\nRandom steam and promo keys: https://maximdev.ru/RSPK",
        'go_to_menu': "\nPress Enter to return to the menu..."
        
    }
}

lang = 'ru_RU' if system_language and (system_language.startswith('ru') or system_language.startswith('Ru')) else 'en_US'
strings = translations[lang]
    
class messagecord:
    
    @staticmethod
    def get_output_folder():
        config = configparser.ConfigParser()
        config_path = os.path.join(script_dir, 'config.ini')
        
        if os.path.exists(config_path):
            config.read(config_path)
            if 'Settings' in config and 'OutputFolder' in config['Settings']:
                output_folder = config['Settings']['OutputFolder']
                if os.path.exists(output_folder):
                    return output_folder
        
        root = tk.Tk()
        root.withdraw()
        output_folder = filedialog.askdirectory(title=strings['select_output_folder'])
        root.destroy()
        
        if output_folder:
            config['Settings'] = {'OutputFolder': output_folder}
            with open(config_path, 'w') as configfile:
                config.write(configfile)
            return output_folder
        return None
    
    @staticmethod
    def copy_file_to_clipboard(file_path):
        try:
            file_path = os.path.abspath(file_path)
            df = DROPFILES()
            df.pFiles = sizeof(DROPFILES)
            df.pt[0] = 0
            df.pt[1] = 0
            df.fNC = 0
            df.fWide = 1
            file_path_buffer = create_string_buffer((file_path + '\0').encode('utf-16le') + b'\0\0')
            total_size = sizeof(DROPFILES) + len(file_path_buffer)
            buffer = create_string_buffer(total_size)
            ctypes.memmove(buffer, bytes(df), sizeof(DROPFILES))
            ctypes.memmove(cast(buffer, c_void_p).value + sizeof(DROPFILES), file_path_buffer, len(file_path_buffer))
            win32clipboard.OpenClipboard()
            try:
                win32clipboard.EmptyClipboard()
                win32clipboard.SetClipboardData(win32con.CF_HDROP, buffer.raw)
            finally:
                win32clipboard.CloseClipboard()
            print(strings['copied_to_clipboard'])
        except Exception as e:
            print(f"{strings['clipboard_error']}: {str(e)}")
    
    @staticmethod
    def check_ffmpeg():
        try:
            ffmpeg_path = os.path.join(bin_path, 'ffmpeg')
            subprocess.run([ffmpeg_path, "-version"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            return True
        except FileNotFoundError:
            print(strings['ffmpeg_not_found'])
            return False
    
    @staticmethod
    def compress_image(input_path, output_path, max_size_mb=9):
        with Image.open(input_path) as img:
            quality = 95
            while True:
                img.save(output_path, optimize=True, quality=quality)
                if os.path.getsize(output_path) <= max_size_mb * 1024 * 1024 or quality <= 5:
                    break
                quality -= 5
    
    @staticmethod
    def get_duration(input_path):
        try:
            probe = ffmpeg.probe(input_path)
            video_stream = next((stream for stream in probe['streams'] if stream['codec_type'] == 'video'), None)
            audio_stream = next((stream for stream in probe['streams'] if stream['codec_type'] == 'audio'), None)
            
            if video_stream and 'duration' in video_stream:
                return float(video_stream['duration'])
            elif audio_stream and 'duration' in audio_stream:
                return float(audio_stream['duration'])
            elif 'format' in probe and 'duration' in probe['format']:
                return float(probe['format']['duration'])
            else:
                raise ValueError(strings['error_getting_duration'])
        except Exception as e:
            raise ValueError(f"{strings['error_getting_duration']} Error: {str(e)}")
    
    @staticmethod
    def compress_video(input_path, output_path, max_size_mb=9):
        try:
            duration = messagecord.get_duration(input_path)
            target_size = max_size_mb * 1024 * 1024 * 8
            bitrate = int(target_size / duration)
    
            (
                ffmpeg
                .input(input_path)
                .output(output_path, video_bitrate=bitrate, audio_bitrate='128k')
                .overwrite_output()
                .run(capture_stdout=True, capture_stderr=True)
            )
        except ffmpeg.Error as e:
            print(strings['error_compressing'].format(str(e)))
        except ValueError as e:
            print(str(e))
    
    @staticmethod
    def compress_audio(input_path, output_path, max_size_mb=9):
        try:
            duration = messagecord.get_duration(input_path)
            target_size = max_size_mb * 1024 * 1024 * 8
            bitrate = int(target_size / duration)
    
            (
                ffmpeg
                .input(input_path)
                .output(output_path, audio_bitrate=f'{bitrate}')
                .overwrite_output()
                .run(capture_stdout=True, capture_stderr=True)
            )
        except ffmpeg.Error as e:
            print(strings['error_compressing'].format(str(e)))
        except ValueError as e:
            print(str(e))
    
    @staticmethod
    def compress_file(input_path, max_size_mb=9):
        file_name, file_extension = os.path.splitext(input_path)
        output_folder = messagecord.get_output_folder()
        
        if not output_folder:
            return
            
        output_file_name = f"{os.path.basename(file_name)}_compressed{file_extension}"
        output_path = os.path.join(output_folder, output_file_name)
        
        print(strings['compressing'])
        
        try:
            if file_extension.lower() in ['.png', '.jpg', '.jpeg']:
                messagecord.compress_image(input_path, output_path, max_size_mb)
            elif file_extension.lower() == '.mkv':
                mp4_output_path = os.path.join(output_folder, f"{os.path.basename(file_name)}_compressed.mp4")
                messagecord.compress_video(input_path, mp4_output_path, max_size_mb)
                output_path = mp4_output_path
                print(strings['mkv_renamed'].format(output_path))
            elif file_extension.lower() in ['.mp4']:
                messagecord.compress_video(input_path, output_path, max_size_mb)
            elif file_extension.lower() == '.mp3':
                messagecord.compress_audio(input_path, output_path, max_size_mb)
            elif file_extension.lower() in ['.zip', '.rar']:
                print(strings['archive_warning'])
                return
            else:
                print(strings['unsupported_format'].format(file_extension))
                return
            
            print(strings['compressed'].format(output_path))
            messagecord.copy_file_to_clipboard(output_path)
        except Exception as e:
            print(strings['error_compressing'].format(str(e)))
        finally:
            input(strings['go_to_menu'])

    
    @staticmethod
    def start():
        input_path = input(strings['input_prompt']).strip()
        input_path = input_path.strip("'\"")
        if os.path.exists(input_path):
            messagecord.compress_file(input_path
        else:
            print(strings['file_not_found'])
    
    @staticmethod
    def get_info():
        print(strings['info_abt_program'])
        input(strings['go_to_menu'])
    
    @staticmethod
    def info():
        print(strings['my_soc'])
        input(strings['go_to_menu'])

if __name__ == '__main__':
    menu = CursesMenu(strings['m_welcome'], strings['m_chopt'])
    function1_item = FunctionItem(strings['m_start'], messagecord.start)
    function2_item = FunctionItem(strings['m_program_func'], messagecord.get_info)
    function3_item = FunctionItem(strings['m_info'], messagecord.info)
    menu.items.append(function1_item)
    menu.items.append(function2_item)
    menu.items.append(function3_item)
    menu.show()