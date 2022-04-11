import xlrd, xlsxwriter
import os
import time
import gspread
from oauth2client.service_account import ServiceAccountCredentials

from sys import platform

if platform == 'win32':
    import ctypes
    kernel32 = ctypes.windll.kernel32
    kernel32.SetConsoleMode(kernel32.GetStdHandle(-11), 7)



SUCCESS_MESSAGE = '\033[2;30;42m [SUCCESS] \033[0;0m' 
warning_message = '\033[2;30;43m [WARNING] \033[0;0m'
error_message = '\033[2;30;41m [ ERROR ] \033[0;0m'

non_existent_niks = []
non_existent_articles = []

class OrderNick:
    def __init__(self, google_sheet_id, excelfile):
        self.google_sheet_id = google_sheet_id
        self.excelfile = excelfile

        self.book_reader = xlrd.open_workbook(self.excelfile)
        self.sheet_reader = self.book_reader.sheet_by_index(0)
        self.table_titles = self.sheet_reader.row_values(0)


    def auth_spread(self, table_id):
        scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
        credentials = ServiceAccountCredentials.from_json_keyfile_name('Service Accounts/morbot-338716-b219142d9c70.json')

        gc = gspread.authorize(credentials)
        spread = gc.open_by_key(table_id)

        return spread

    
    def collect_all_nicks(self, excelfile):
        for title_num in range(len(self.table_titles)):
            if self.table_titles[title_num] == 'nik товара':
                col_num = title_num

        all_niks = self.sheet_reader.col_values(col_num)[1:]

        return all_niks


    def collect_status_orders(self, item):
        for title_num in range(len(self.table_titles)):
            if self.table_titles[title_num] == 'Группа статуса':
                group_status_col = title_num

        spam_group = 0
        cancel_group = 0
        sent_group = 0
        processing_group = 0

        group_in_way = 0
        group_bought_out = 0
        group_refund = 0

        for row_num in range(self.sheet_reader.nrows):
            if item in self.sheet_reader.row_values(row_num):
                if self.sheet_reader.cell(row_num, group_status_col).value == 'Ошибка/Спам/Дубль':
                    spam_group += 1
                elif self.sheet_reader.cell(row_num, group_status_col).value in ('Оплачен', 'Отправлен', 'Принят', 'Возврат'):
                    sent_group += 1
                    if self.sheet_reader.cell(row_num, group_status_col).value in ('Отправлен', 'Принят'):
                        group_in_way += 1
                    elif self.sheet_reader.cell(row_num, group_status_col).value == 'Оплачен':
                        group_bought_out += 1

                if self.sheet_reader.cell(row_num, group_status_col).value == 'Возврат':
                    group_refund += 1

                elif self.sheet_reader.cell(row_num, group_status_col).value == 'Отменен':
                    cancel_group += 1
                elif self.sheet_reader.cell(row_num, group_status_col).value == 'Обработка':
                    processing_group += 1                
        return (spam_group, cancel_group, sent_group, processing_group, group_in_way, group_bought_out, group_refund)


    def get_frequrent_dict(self):
        data = self.collect_all_nicks(self.excelfile)

        freq_dict = {}
        for item in data:
            if item == '':
                pass
            else:
                count = 0
                for item2 in data:
                    if item == item2:
                        count += 1
                freq_dict[item] = count
        return freq_dict



    def find_nicks_in_sheet(self, data):
        spread = self.auth_spread(self.google_sheet_id)
        worksheet = spread.get_worksheet(6) # 6
        names_col = worksheet.find('наименование товара').col
        order_names = worksheet.col_values(names_col)
        
        def try_write(item, value):
            try:
                count = len(order_names)
                for name_row in range(len(order_names)):
                    if item in [i.strip() for i in order_names[name_row].split('\n')]:
                        
                        total_number_of_orders_value = worksheet.acell(f'E{name_row+1}').value if worksheet.acell(f'E{name_row+1}').value != None else 0
                        total_number_of_orders_IN_PROCCESING_value = worksheet.acell(f'F{name_row+1}').value if worksheet.acell(f'F{name_row+1}').value != None else 0
                        total_number_of_orders_SPAM_value = worksheet.acell(f'G{name_row+1}').value if worksheet.acell(f'G{name_row+1}').value != None else 0
                        total_number_of_orders_CANCEL_value = worksheet.acell(f'H{name_row+1}').value if worksheet.acell(f'H{name_row+1}').value != None else 0
                        total_number_of_orders_SENT_value = worksheet.acell(f'I{name_row+1}').value if worksheet.acell(f'I{name_row+1}').value != None else 0
                        total_number_of_orders_IN_WAY_value = worksheet.acell(f'J{name_row+1}').value if worksheet.acell(f'J{name_row+1}').value != None else 0
                        total_number_of_orders_BOUGHT_OUT_value = worksheet.acell(f'K{name_row+1}').value if worksheet.acell(f'K{name_row+1}').value != None else 0
                        total_number_of_orders_REFUND_value = worksheet.acell(f'L{name_row+1}').value if worksheet.acell(f'L{name_row+1}').value != None else 0

                        spam_group, cancel_group, sent_group, processing_group, group_in_way, group_bought_out, group_refund = self.collect_status_orders(item)
 
                        worksheet.update(f'E{name_row+1}', int(total_number_of_orders_value) + int(value))
                        worksheet.update(f'F{name_row+1}', int(total_number_of_orders_IN_PROCCESING_value) + int(processing_group))
                        worksheet.update(f'G{name_row+1}', int(total_number_of_orders_SPAM_value) + int(spam_group))
                        worksheet.update(f'H{name_row+1}', int(total_number_of_orders_CANCEL_value) + int(cancel_group))
                        worksheet.update(f'I{name_row+1}', int(total_number_of_orders_SENT_value) + int(sent_group))
                        worksheet.update(f'J{name_row+1}', int(total_number_of_orders_IN_WAY_value) + int(group_in_way))
                        worksheet.update(f'K{name_row+1}', int(total_number_of_orders_BOUGHT_OUT_value) + int(group_bought_out))
                        worksheet.update(f'L{name_row+1}', int(total_number_of_orders_REFUND_value) + int(group_refund))
                        print(SUCCESS_MESSAGE + '\tНайден товар ', item.lower())
                        count -= 1
                        return
                global non_existent_niks
                non_existent_niks.append(item)
            # except BaseException:
            #     print(item)

            except gspread.exceptions.APIError:
                time.sleep(31)
                print(error_message + '\tБоту нужно отдохнуть 30 секунд')
                try_write(item, value)
        for item in data:  
            # print(item)      
            try_write(item, data[item]) 
            # try_write('Cledbel 24K Gold', 2) 
            # break


    def collect_all_articles(self):
        for title_num in range(len(self.table_titles)):
            if self.table_titles[title_num] == 'Товары':
                col_num = title_num
                break
        all_products = list(set(self.sheet_reader.col_values(col_num)[1:]))
        
        return all_products


    def collect_articles_status(self, item):
        for title_num in range(len(self.table_titles)):
            if self.table_titles[title_num] == 'Группа статуса':
                articles_status_col = title_num

        group_refund = 0
        group_bought_out = 0
        for row_num in range(self.sheet_reader.nrows):
            if item in self.sheet_reader.row_values(row_num):
                if self.sheet_reader.cell(row_num, articles_status_col).value == 'Оплачен':
                    group_bought_out += 1
                elif self.sheet_reader.cell(row_num, articles_status_col).value == 'Возврат':
                    group_refund += 1
        return group_refund, group_bought_out


    def find_all_articles_in_sheet(self, data):
        spread = self.auth_spread(self.google_sheet_id)
        worksheet = spread.get_worksheet(6)
        arc_col = worksheet.find('Артикул').col
        all_products_arcs = worksheet.col_values(arc_col)
        
        def try_write(item):
            if item == '':
                return
            try:
                for name_row in range(len(all_products_arcs)):
                    if all_products_arcs[name_row] == item:
                        total_num_of_product_bought_out = worksheet.acell(f'Q{name_row+1}').value if worksheet.acell(f'Q{name_row+1}').value != None else 0
                        total_num_of_product_refund = worksheet.acell(f'R{name_row+1}').value if worksheet.acell(f'R{name_row+1}').value != None else 0

                        group_refund, group_bought_out = self.collect_articles_status(item)
                        worksheet.update(f'Q{name_row+1}', int(total_num_of_product_bought_out) + int(group_bought_out))
                        worksheet.update(f'R{name_row+1}', int(total_num_of_product_refund) + int(group_refund))

                        print(SUCCESS_MESSAGE + '\tНашли артикул ', item)
                        return
                        
                global non_existent_articles
                non_existent_articles.append(item)

            except gspread.exceptions.APIError:
                print('Бот умер')
                time.sleep(31)
                try_write(item)
        for item in data:
            try_write(item)




def save_table(data, filename):
    workbook_writer = xlsxwriter.Workbook(filename+'.xls')
    worksheet_writer = workbook_writer.add_worksheet()

    for line in range(len(data)):
        for col in range(len(data[line])):
            worksheet_writer.write(line, col, data[line][col])

    workbook_writer.close() 



def create_non_existent_TABLE(non_existent_file, excelfile, colname):
    file = open(non_existent_file+'.txt')

    wb_reader = xlrd.open_workbook(excelfile)
    sheet_reader = wb_reader.sheet_by_index(0)
    
    table_titles = sheet_reader.row_values(0)
    for title_col in range(len(table_titles)):
        if table_titles[title_col] == 'Группа статуса':
            status_col = title_col
        if table_titles[title_col] == colname:
            products_col = title_col 
            break
    
    res_data = [table_titles, ]
    all_products_in_table = sheet_reader.col_values(products_col)
    all_status_in_table = sheet_reader.col_values(status_col)
    for product in file:
        for item in range(len(all_products_in_table)):
            if all_products_in_table[item].strip() == product.strip():
                if all_status_in_table[item].strip() in ('Возврат', 'Оплачен'):
                    res_data.append(sheet_reader.row_values(item))
    
    save_table(res_data, non_existent_file)



def create_non_existent_FILE(data, filename):
    if len(data) > 0:
        file = open(f'{filename}.txt', 'a')
        for item in data:
            file.write(item + '\n')
        file.close()


def main():
    upload_folder = 'Upload Excel'
    os.makedirs(upload_folder, exist_ok=True)
    for file in os.listdir(upload_folder):
        if file.endswith('.xls'):
            excel_file = os.path.join(upload_folder, file)
            break
    try:
        nik_order = OrderNick('13N8mWPuPGym1WR0jCdXU_GL4AdMZWYdJQZ22xqW9IDw', excel_file)
        freq_dict = nik_order.get_frequrent_dict()
        nik_order.find_nicks_in_sheet(freq_dict)

        all_products = nik_order.collect_all_articles()
        nik_order.find_all_articles_in_sheet(all_products)

        os.makedirs('NON-EXISTENT', exist_ok=True)
        niks_filename = 'NON-EXISTENT/None NIKS'
        articles_filename = 'NON-EXISTENT/None ARTICLES'

        create_non_existent_FILE(non_existent_niks, niks_filename)
        create_non_existent_FILE(non_existent_articles, articles_filename)

        print(warning_message + '\tСоздаем таблицу с исключениями')
        create_non_existent_TABLE(niks_filename, excel_file, colname='nik товара')
        create_non_existent_TABLE(articles_filename, excel_file, colname='Товары')

    except UnboundLocalError:
        print('В папке нет ни одного excel-файла.')    


# main()

if __name__ == '__main__':
    main()
    print(SUCCESS_MESSAGE + '\tВсе готово!')
    # 11.04.2022
