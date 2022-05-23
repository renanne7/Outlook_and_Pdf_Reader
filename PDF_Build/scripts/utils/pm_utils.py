import csv
import json
import PyPDF2
import tabula
import dateparser
from dateutil import parser
from datetime import datetime
from pyjarowinkler import distance

from scripts.constants.pm_constants import CLIENT_LIST_JSON
from scripts.constants.pm_constants import LAST_UPDATED_JSON, LAST_UPDATED_JSON_SAMPLE
from scripts.constants.pm_constants import OUTPUT_CSV_FILE
from scripts.constants.pm_constants import SOLUTION_LIST_JSON
from scripts.constants.pm_constants import TEXT_IDENTIFIER, DATE_IDENTIFIER, SHARE_POINT_SITE_URL
from scripts.utils.pm_logging import get_logger

logger = get_logger()


def time_sorter(data_time, mail_time):
    # logger.info("Starting to sort time")
    # logger.debug("Data Time: " + str(data_time))
    # logger.debug("Mail Time: " + str(mail_time))
    data_time = str(data_time)
    mail_time = str(mail_time)
    data_time = dateparser.parse(data_time)
    mail_time = dateparser.parse(mail_time)
    if data_time < mail_time:
        # logger.info("Time sorter returned True")
        return True
    else:
        # logger.info("Time sorter returned False")
        return False


def read_last_updated_json():
    try:
        with open(LAST_UPDATED_JSON, "r") as read_json:
            json_data = read_json.read()
    except FileNotFoundError:
        json_data = LAST_UPDATED_JSON_SAMPLE
    return json_data


def pdf_extractor(filename, pdf_data):
    # Open allows you to read the file
    pdf_file_obj = open(filename, 'rb')

    # The pdfReader variable is a readable object that will be parsed
    pdf_reader = PyPDF2.PdfFileReader(pdf_file_obj)
    # Extract only tables from pdf using tabula
    tabula.convert_into(filename, OUTPUT_CSV_FILE, output_format='csv', pages="all")
    # Discerning the number of pages will allow us to parse through all the pages
    num_pages = pdf_reader.numPages

    count = 0
    text = ""

    # The while loop will read each page
    while count < num_pages:
        page_obj = pdf_reader.getPage(count)
        count = count + 1
        text = text + page_obj.extractText()
        # print(text)

    text = text.replace('\n', ' ')
    text = text.replace('\t', ' ')
    text = text.replace('   ', ' ')
    text = text.replace('  ', ' ')

    required_date = None
    client_name = None
    sol = None
    qty_sum = None
    required_customer_num = None
    required_client_num = None
    final_solution = None

    if text.strip():
        # Extract Signed on date
        while True:
            try:
                sample_date_text = text.split(TEXT_IDENTIFIER)
                sample_date_text = sample_date_text[1]
                sample_date_text = sample_date_text.split(DATE_IDENTIFIER)
                sample_date_text = sample_date_text[0]
                sample_date_text = sample_date_text.lstrip(' ')
                sample_date_text = sample_date_text.rstrip(' ("')
                date_of_order = parser.parse(sample_date_text)
                required_date = datetime.strftime(date_of_order, "%Y-%m-")
                logger.debug("Pdf Date: " + str(required_date))
                break
            except Exception as e:
                logger.debug("Pdf Date: " + str(required_date))
                required_date = None
            break

        # Extract Client Name
        while True:
            try:
                sample_client_name = text.split('between ')
                sample_client_name = sample_client_name[1]
                sample_client_name = sample_client_name.split(' ("')
                sample_client_name = sample_client_name[0]
                sample_client_name = " ".join(sample_client_name.split(" ", 2)[:2])
                sample_client_name = sample_client_name.lstrip(' ')
                client_name = sample_client_name.rstrip(' ("')
                logger.debug("Pdf Client Name: " + str(client_name))
                break
            except Exception as e:
                logger.debug("Pdf Client Name: " + str(client_name))
                client_name = None
            break

        # Read the output csv file which contains all tables
        file = open(OUTPUT_CSV_FILE, 'rt')
        reader = csv.reader(file)
        table_list = list(reader)

        # Iterate csv to find Solution
        columnIndex = 0
        rowIndex = 0
        for each_line in range(0, len(table_list)):
            sol_list = table_list[each_line]
            if "Solution" in sol_list:
                solution_column = sol_list.index("Solution")
                columnIndex = solution_column
                rowIndex = each_line
                rowIndex = rowIndex + 1

                while len(table_list[rowIndex][columnIndex]) == 0:
                    rowIndex = rowIndex + 1
                else:
                    sol = table_list[rowIndex][columnIndex]
                if sol == "":
                    sol = None

        if sol is not None:
            # Check Solution json and find short solution
            with open(SOLUTION_LIST_JSON, "r") as read_json:
                sol_data = read_json.read()

            SOL_FINDER = json.loads(sol_data)

            dlist = []
            for each_row in range(0, len(SOL_FINDER)):
                sol_list = SOL_FINDER[each_row]
                d = distance.get_jaro_distance(sol, sol_list['solution'], winkler=True, scaling=0.1)
                dlist.append(d)

            match = max(dlist)
            last_index = len(SOL_FINDER)

            if match > 0.9:
                if match in dlist:
                    row = dlist.index(match)
                    logger.debug("Pdf Solution: " + str(SOL_FINDER[row]['shortsolution']))
                    final_solution = str(SOL_FINDER[row]['shortsolution'])

                    if match < 1.0:
                        with open(SOLUTION_LIST_JSON, "w") as f:
                            json.dump([], f)

                        with open(SOLUTION_LIST_JSON, "w") as write_json:
                            entry = {"solution": sol, "shortsolution": SOL_FINDER[row]['shortsolution']}
                            SOL_FINDER.append(entry)
                            json.dump(SOL_FINDER, write_json)

            else:
                final_solution = sol
                logger.debug("Pdf Solution: " + str(sol))
                with open(SOLUTION_LIST_JSON, "w") as f:
                    json.dump([], f)

                with open(SOLUTION_LIST_JSON, "w") as write_json:
                    entry = {"solution": sol, "shortsolution": sol}
                    SOL_FINDER.append(entry)
                    json.dump(SOL_FINDER, write_json)

        else:
            logger.debug("Pdf Solution: " + str(sol))

        # Iterate csv to find Qty
        for each_line in range(0, len(table_list)):
            qty_list = table_list[each_line]
            if "Qty" in qty_list:
                quantity_column = qty_list.index("Qty")
                columnIndex = quantity_column
                rowIndex = each_line

                while True:
                    try:
                        rowIndex = rowIndex + 1
                        if len(table_list[rowIndex][columnIndex]) != 0:
                            data = int(table_list[rowIndex][columnIndex])
                            # print(data)
                            qty_sum = qty_sum + data
                            found = True
                    except Exception as e:
                        qty_sum = None
                        break
        logger.debug("Pdf Quantity: " + str(qty_sum))

        # Iterate csv to find Customer Number
        while True:
            try:
                sample_client_num = text.split("Customer Number: ")
                sample_client_num = sample_client_num[1]
                sample_client_num = sample_client_num.split(" Facility")
                required_customer_num = sample_client_num[0]
                logger.debug("Pdf Customer Number: " + str(required_customer_num))
                break
            except Exception as e:
                logger.debug("Pdf Customer Number: " + str(required_customer_num))
                required_customer_num = None
                break

        # Iterate csv to find Client Number
        while True:
            try:
                sample_client_num = text.split("Client Number: * ")
                sample_client_num = sample_client_num[1]
                sample_client_num = sample_client_num.split(" Facility")
                required_client_num = sample_client_num[0]
                required_client_num = required_client_num.split()
                required_client_num = required_client_num[0]
                logger.debug("Pdf Client Number: " + str(required_client_num))
                break
            except Exception as e:
                logger.debug("Pdf Client Number: " + str(required_client_num))
                required_client_num = None
                break

        pdf_data = {
            "date": str(required_date),
            "clientName": str(client_name),
            "solution": str(final_solution),
            "quantity": str(qty_sum),
            "customerNumber": str(required_customer_num),
            "clientNumber": str(required_client_num)}

    else:
        logger.error("PDF was not read - text was empty")
        pdf_data = {
            "date": str(required_date),
            "clientName": str(client_name),
            "solution": str(final_solution),
            "quantity": str(qty_sum),
            "customerNumber": str(required_customer_num),
            "clientNumber": str(required_client_num)}

    pdf_file_obj.close()
    return pdf_data


def update_last_updated_json(updated_json):
    with open(LAST_UPDATED_JSON, "w") as write_json:
        json.dump(updated_json, write_json)


def update_new_latest_time(uploaded_data):
    try:
        uploaded_data.sort(key=lambda x: dateparser.parse(x["datetime"]))
        update_last_updated_json(uploaded_data[-1])
        return True
    except IndexError:
        return True


def create_path_for_pdf(pdf_data, mail_data):

    if (pdf_data["customerNumber"]) != 'None':
        customer_number = str(pdf_data["customerNumber"])
    elif (pdf_data["clientNumber"]) != 'None':
        customer_number = str(pdf_data["clientNumber"])
    elif (pdf_data["clientName"]) != 'None':
        customer_number = str(pdf_data["clientName"])
    elif (mail_data["mnemonic"]) is not None:
        customer_number = str(mail_data["mnemonic"])
    else:
        customer_number = None

    if (pdf_data["date"]) != 'None':
        required_date = str(pdf_data["date"])
    elif (mail_data["date"]) is not None:
        required_date = str(mail_data["date"])
    else:
        required_date = None

    if (mail_data["pn"]) != '':
        pn = str(mail_data["pn"])
    else:
        pn = None

    if (mail_data["solution"]) != '':
        solution = str(mail_data["solution"])
    elif (pdf_data["solution"]) != 'None':
        solution = str(pdf_data["solution"])
    else:
        solution = None

    site_url = None
    with open(CLIENT_LIST_JSON, "r") as read_path:
        path_data = read_path.read()

    path_finder = json.loads(path_data)

    for each_row in range(0, len(path_finder)):
        site_list = path_finder[each_row]
        if customer_number in site_list['siteName']:
            site_url = path_finder[each_row]['newUrl']
    if site_url is not None and required_date is not None and pn is not None and solution is not None:

        pdf_data["quantity"] = int()
        if pdf_data["quantity"] < 1000:

            sample_url = site_url.split(SHARE_POINT_SITE_URL)
            url_part_one = '/sites/'
            url_part_two = sample_url[1]
            url_part_three = '/Quality Documents/05 - Consulting'
            folder_name = required_date + pn + solution

            storage_path = url_part_one + url_part_two + url_part_three

            logger.debug("Upload site: " + str(site_url))
            logger.debug("Folder name: " + str(folder_name))
            return site_url, storage_path, folder_name
        else:

            sample_url = site_url.split(SHARE_POINT_SITE_URL)
            url_part_one = '/sites/'
            url_part_two = sample_url[1]
            url_part_three = '/Quality Documents/05 - Consulting'
            folder_name = None

            storage_path = url_part_one + url_part_two + url_part_three

            logger.debug("Upload site: " + str(site_url))
            logger.debug("Folder name: " + str(folder_name))
            return site_url, storage_path, folder_name
    else:
        site_url = None
        storage_path = None
        folder_name = None
        return site_url, storage_path, folder_name
