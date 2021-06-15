import openpyxl as xl
import unicodedata
import optparse
from tqdm import tqdm

class Normalizer:
    def __init__(self, file, selected_sheets):
        print('[*] Loading file...')
        self.workbook = xl.load_workbook(file)
        self.selected_sheets = [self.workbook.worksheets[i] for i in selected_sheets] if selected_sheets else self.workbook.worksheets

    def normalize(self):
        for sheet in self.selected_sheets:
            print(f'[*] Normalizing sheet {sheet.title}')
            with tqdm(unit=' cells') as bar:
                for row in sheet:
                    for cell in row:
                        if type(cell.value) is str:
                            cell.value = unicodedata.normalize('NFKD', cell.value).encode('ASCII', 'ignore').decode('ASCII')
                            bar.update()

if __name__ == '__main__':
    parser = optparse.OptionParser(usage='%prog file [OPTIONS]')
    parser.add_option('-s', '--sheets', dest='selected_sheets', help='Select specific sheets for normalization')

    options, args = parser.parse_args()
    if len(args) != 1:
        parser.error('[!] Invalid format.')

    file = args[0]
    selected_sheets = options.selected_sheets.split(',') if options.selected_sheets else None

    normalizer = Normalizer(file, selected_sheets)
    normalizer.normalize()

    print('[*] Saving file...')
    normalizer.workbook.save(file)
