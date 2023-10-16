from datetime import datetime
import os


class Color:
    GREEN = '\033[32m'
    RED = '\033[31m'
    YELLOW = '\033[33m'
    WHITE = '\033[0m'
    CYAN = '\033[96m'


def king_logging(text: str, type_: str = 'succes'):
    date_time = datetime.now().strftime('%Y.%m.%d %H:%M.%S')

    if type_ == 'succes':
        print(f'[{Color.CYAN}{date_time}{Color.WHITE}] - {text}')
    else:
        print(f'[{Color.CYAN}{date_time}{Color.WHITE}] - {Color.RED}{text}{Color.WHITE}')


def check_exists_files(files_paths: list) -> bool:
    result = True
    max_len_file = 0

    for file_path in files_paths:
        if max_len_file < len(file_path):
            max_len_file = len(file_path)

    for file_path in files_paths:
        if not os.path.exists(file_path):
            print(f'{Color.RED}{file_path: <{max_len_file}} - не найден :({Color.WHITE}')
            result = False

    return result
