from dataclasses import dataclass, asdict
from typing import Dict, List
import requests
from bs4 import BeautifulSoup
import sys

from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side

keys = ['Reference', 'Package', 'Flash', 'RAM']
base_url = r'https://www.terraelectronica.ru'


@dataclass
class MicroController:
    partnumber: str
    is_available: bool
    sales_data: list
    package: str
    ram: int
    flash: int

    def __str__(self):
        return str(asdict(self))


def get_start_index(sheet)-> int:
    """
    gets index for title row
    :param sheet:
    :return:
    """
    for i in range(1, sheet.max_row):
        value = sheet['A%i' % i].value
        if value and 'Part No' in value:
            return i
    return 0


def get_column_indexes(sheet, i: int) -> Dict[str, str]:
    """
    gets indexes of service columns
    :param sheet: sheet of xlxs
    :param i: number of row with titles
    :return: dict with indices
    """
    indices = dict()
    for letter in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
        value = sheet['%s%i' % (letter, i)].value
        if sheet['%s%i' % (letter, i)].value:
            for key in keys:
                if key in value:
                    indices[key] = letter
    return indices


def create_mc_list(sheet, indices: Dict[str, str], start_index: int) -> List[MicroController]:
    """

    :param sheet:
    :param indices:
    :param start_index:
    :return:
    """
    microcontrollers: List[MicroController] = list()
    for i in range(start_index+1, sheet.max_row):
        partnumber: str = sheet["%s%i" % (indices['Reference'], i)].value
        package: str = sheet["%s%i" % (indices['Package'], i)].value if 'Package' in indices.keys() else "Unknown"
        flash: str = sheet["%s%i" % (indices['Flash'], i)].value if 'Flash' in indices.keys() else "Unknown"
        ram: str = sheet["%s%i" % (indices['RAM'], i)].value if 'RAM' in indices.keys() else "Unknown"
        microcontroller = MicroController(partnumber=partnumber, is_available=False, sales_data=list(), package=package,
                                          flash=int(flash.split()[0]), ram=int(ram.split()[0]))
        microcontrollers.append(microcontroller)
    return microcontrollers


def update_data_for_catalog(microcontroller: MicroController, request):
    """
    gets data for microcontroller for search redirected to catalog
    :param microcontroller: data for microcontroller to find
    :param request: request
    :return:
    """
    url_present: str = request.url + '&f%5Bpresent%5D=1'
    # print(url_present)
    request = requests.get(url_present)
    soup = BeautifulSoup(request.text)
    items = soup.find_all('tr')
    results = list()
    for item in items[1:]:
        pn, price, count, product_id, instock = "", 0, 0, "", 0
        try:
            content_url = item.contents[3]
            if content_url.attrs['class'][0] == 'table-item-name':
                content_url = content_url.contents[3]
                product_id = content_url.contents[0].attrs['href']
                pn = content_url.contents[0].contents[0]
            content_price = item.contents[11].contents[1].contents[1]
            if content_price.attrs['class'] == ['price-single', 'price-active']:
                price = float(content_price.attrs['data-price'])
                count = int(content_price.attrs['data-count'])
            content_count = item.contents[13].contents[1]
            if content_count.attrs['class'] == ['item-qnt']:
                instock = content_count.contents[0].split()[0]
            product_data = dict(PN=pn, Price=price, Count=count, Url=base_url + product_id, Instock=instock)
            results.append(product_data)
            print(product_id, pn, price, count, instock)
        except (TypeError, IndexError, KeyError):
            print("error!")
            continue
    if len(results) > 0:
        microcontroller.sales_data = results.copy()
        microcontroller.is_available = True


def update_from_common_catalog(microcontroller: MicroController, request):
    """

    :param microcontroller:
    :param request:
    :return:
    """
    soup = BeautifulSoup(request.text)
    items = soup.find_all('tr')
    results = list()
    for item in items[1:]:
        pn, price, count, product_id, instock = "", 0, 0, "", 0
        try:
            content_url = item.contents[3]
            if content_url.attrs['class'][0] == 'table-item-name':
                content_url = content_url.contents[3]
                product_id = content_url.contents[0].attrs['href']
                pn = content_url.contents[0].contents[0]
            content_price = item.contents[11].contents
            if len(content_price) > 1:
                if content_price[1].contents[1].attrs['class'] == ['price-single', 'price-active']:
                    price = float(content_price[1].contents[1].attrs['data-price'])
                    count = int(content_price[1].contents[1].attrs['data-count'])
                content_count = item.contents[13].contents[1]
                if content_count.attrs['class'] == ['item-qnt']:
                    instock = content_count.contents[0].split()[0]
            if int(instock) > 0:
                product_data = dict(PN=pn, Price=price, Count=count, Url=base_url + product_id,
                                    Instock=instock)
                results.append(product_data)
                print(product_id, pn, price, count, instock)
        except (TypeError, IndexError, KeyError):
            print("error!")
            continue
    if len(results) > 0:
        microcontroller.sales_data = results.copy()
        microcontroller.is_available = True


def write_to_file(microcontrollers: List[MicroController]):
    """
    writes list of availale microcontrrollers to xlsx
    :param microcontrollers: list of microcontrollers
    :return:
    """
    wb = Workbook()
    wb.guess_types = True
    ws1 = wb.active
    ft = Font(bold=True)
    alignment = Alignment(horizontal='center', indent=0.2)
    border = Border(bottom=Side(border_style="double", color='00000000'))
    ws1.column_dimensions['A'].width = 20
    ws1.column_dimensions['B'].width = 20
    ws1.column_dimensions['F'].width = 50

    ws1['A1'] = "Model"
    ws1['B1'] = 'PN'
    ws1["C1"] = "Price"
    ws1['D1'] = 'Min count'
    ws1['E1'] = 'InStock'
    ws1['F1'] = 'Url'
    ws1['G1'] = 'Package'
    ws1['H1'] = 'Flash'
    ws1['I1'] = 'RAM'
    cells = [ws1['A1'], ws1['B1'], ws1['C1'], ws1['D1'], ws1['E1'], ws1['F1'], ws1['G1'], ws1['H1'], ws1['I1']]
    for cell in cells:
        cell.font = ft
        cell.border = border
        cell.alignment = alignment
    i = 2
    for microcontroller in microcontrollers:
        if microcontroller.is_available:
            for product in microcontroller.sales_data:
                ws1['A%i' % i] = microcontroller.partnumber
                ws1['B%i' % i] = product['PN']
                ws1['C%i' % i] = product['Price']
                ws1['D%i' % i] = product['Count']
                ws1['E%i' % i] = product['Instock']
                ws1['F%i' % i] = product['Url']
                ws1['G%i' % i] = microcontroller.package
                ws1['H%i' % i] = microcontroller.flash
                ws1['I%i' % i] = microcontroller.ram
                i += 1
    wb.save(filename='Results.xlsx')



def main(filename: str):

    try:
        wb = load_workbook(filename=filename)
        sheet = wb.active
    except FileNotFoundError:
        print("File %a not found" % filename)
        return

    start_index = get_start_index(sheet)
    if start_index == 0:
        print("No column titles")
        return

    indices = get_column_indexes(sheet, start_index)
    if 'Reference' not in indices.keys():
        print("No partnumber column, unable to proceed")
        return

    microcontrollers = create_mc_list(sheet, indices, start_index)

    for microcontroller in microcontrollers:
        search_text: str = microcontroller.partnumber
        search_text = search_text.replace('x', '')
        url: str = base_url + r"/search?text=" + search_text
        r = requests.get(url)
        if 'mikrokontrollery' in r.url:
            update_data_for_catalog(microcontroller, r)
        elif "search?" in r.url:
            soup = BeautifulSoup(r.text)
            links = soup.find('ul', {'class': "search-list"})
            try:
                link = base_url + '/' + links.contents[1].contents[1].attrs['href']
            except (AttributeError, TypeError, IndexError):
                print("Data error %s" % url)
            r = requests.get(link)
            update_data_for_catalog(microcontroller, r)
        elif r'catalog/products/' in r.url:
            update_from_common_catalog(microcontroller, r)
        elif r'/product/' in r.url:

             pass
        else:
            print(r.url+' is not parced')

    write_to_file(microcontrollers)


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Please specify filename")
    else:
        main(sys.argv[1])

