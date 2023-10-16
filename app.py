from src.services.utils import check_exists_files, Color
from src.config import required_files
from src.services.render_docx import RatstishkaAPI


def main():
    print(
        fr'''{Color.CYAN}
            ╔═════════════════════╗
            ║  {Color.WHITE}Auto-fill-docx{Color.CYAN}     ║
            ║  {Color.WHITE}Version 1.2{Color.CYAN}              ║
            ╚═════════════════════╝
            {Color.WHITE}'''
    )
    if not check_exists_files(required_files):
        return input('\nДля завершения работы нажми Enter...')

    with open('session', 'r', encoding='utf-8') as session:
        rat = RatstishkaAPI(session.read().replace('\n', ''))
        rat.start()


if __name__ == '__main__':
    main()
