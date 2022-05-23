import csv
import os
import re
import time
import json
from pandas import ExcelWriter
from datetime import datetime
from pyjarowinkler import distance

import win32com.client
from dateutil import parser

from scripts.utils.pm_utils import time_sorter, read_last_updated_json, pdf_extractor, create_path_for_pdf
from scripts.constants.pm_constants import MAIL_CSV_FILE, PDF_UPLOADS_FOLDER
from scripts.constants.pm_constants import PDF_TEMP_DOWNLOAD_PATH, SOLUTION_LIST_JSON
from scripts.utils.pm_logging import get_logger

logger = get_logger()

VERIFICATION_JSON = json.loads(read_last_updated_json())


class OutlookMailReader(object):

    def __init__(self, recipient_name):
        self.outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        self.recipient = self.outlook.CreateRecipient(recipient_name)
        self.inbox = self.outlook.GetSharedDefaultFolder(self.recipient, 6)
        self.messages = self.inbox.Items
        self.mail = self.messages.GetFirst()

    def mail_reader(self):
        # Get all email details
        mail_sender = self.mail.SenderName
        logger.debug("Mail Sender: " + str(mail_sender))
        mail_subject = self.mail.Subject
        logger.debug("Mail Sub: " + str(mail_subject))
        mail_body = self.mail.Body
        mail_receive_time = self.mail.ReceivedTime

        mail_attachment = None

        try:
            mail_attachments = self.mail.Attachments
            mail_attachment = mail_attachments.Item(1)
        except Exception as e:
            logger.info("No attachment")

        # Clean up email body
        mail_body = mail_body.replace('\n', ' ')
        mail_body = mail_body.replace('\t', ' ')
        mail_body = mail_body.replace('   ', ' ')
        mail_body = mail_body.replace('  ', ' ')
        mail_body = mail_body.replace(',', ' ')

        # writing mail body to csv
        f = open(MAIL_CSV_FILE, 'w')
        f.write(mail_body)
        f.close()

        file = open(MAIL_CSV_FILE, 'rt')
        reader = csv.reader(file)
        mail_list = list(reader)

        # Iterate csv to find mnemonic
        mnemonic = None
        for each_line in range(0, len(mail_list)):
            mne_list = mail_list[each_line]
            if " Mnemonic:" in mne_list:
                columnIndex = mne_list.index(" Mnemonic:")
                rowIndex = each_line
                rowIndex = rowIndex + 1
                while mail_list[rowIndex][columnIndex] == " ":
                    rowIndex = rowIndex + 1
                else:
                    mnemonic = mail_list[rowIndex][columnIndex]
                    mnemonic = mnemonic.replace(' ', '')
                if mnemonic == "":
                    mnemonic = None
        logger.debug("Mail Mnemonic: " + str(mnemonic))

        # Iterate csv to find signed on
        required_date = None
        try:
            for each_line in range(0, len(mail_list)):
                date_list = mail_list[each_line]
                if " Signed On:" in date_list:
                    columnIndex = date_list.index(" Signed On:")
                    rowIndex = each_line
                    rowIndex = rowIndex + 1
                    while mail_list[rowIndex][columnIndex] == " ":
                        rowIndex = rowIndex + 1
                    else:
                        signed_on = mail_list[rowIndex][columnIndex]
                        signed_on = signed_on.replace(' ', '')
                        if signed_on != '':
                            date_of_order = parser.parse(signed_on)
                            required_date = datetime.strftime(date_of_order, "%Y-%m-")
                        else:
                            required_date = None
                    if required_date == "":
                        required_date = None
            logger.debug("Mail date: " + str(required_date))
        except Exception as e:
            logger.debug("Mail date: " + str(required_date))

        # Iterate csv to find solution
        sol = None
        sol_list = list()
        solution = ''
        final_solution = ''
        for each_line in range(0, len(mail_list)):
            mail_rows = mail_list[each_line]
            if " Delivery Method:" in mail_rows:
                columnIndex = mail_rows.index(" Delivery Method:")
                rowIndex = each_line
                last_index = len(mail_list) - 40

                while rowIndex < last_index:
                    try:
                        rowIndex = rowIndex + 18
                        if len(mail_list[rowIndex]) != 0:
                            sol = str(mail_list[rowIndex])
                            sol = sol.replace(' ', '')
                            sol = re.findall(r'\'([^]]*)\'', sol)
                            sol_list.append(sol)
                            found = True
                    except Exception as e:
                        break
                if sol == "":
                    solution = None

        if len(sol_list) != 0:
            for item in sol_list:
                for each in item:
                    if solution != '':
                        solution = solution + ","
                    solution = solution + str(each)
        else:
            solution = None

        if solution is not None and solution != 'None' and solution != '':
            solution = solution[:-1]
            # Check Solution json and find short solution
            with open(SOLUTION_LIST_JSON, "r") as read_json:
                sol_data = read_json.read()

            SOL_FINDER = json.loads(sol_data)

            dlist = []
            for each_row in range(0, len(SOL_FINDER)):
                sol_list = SOL_FINDER[each_row]
                d = distance.get_jaro_distance(solution, sol_list['solution'], winkler=True, scaling=0.1)
                dlist.append(d)

            match = max(dlist)
            last_index = len(SOL_FINDER)

            if match > 0.9:
                if match in dlist:
                    row = dlist.index(match)
                    logger.debug("Mail Solution: " + str(SOL_FINDER[row]['shortsolution']))
                    final_solution = str(SOL_FINDER[row]['shortsolution'])

                    if match < 1.0:
                        with open(SOLUTION_LIST_JSON, "w") as f:
                            json.dump([], f)

                        with open(SOLUTION_LIST_JSON, "w") as write_json:
                            entry = {"solution": solution, "shortsolution": SOL_FINDER[row]['shortsolution']}
                            SOL_FINDER.append(entry)
                            json.dump(SOL_FINDER, write_json)

            else:
                final_solution = solution
                logger.debug("Mail Solution: " + str(solution))
                with open(SOLUTION_LIST_JSON, "w") as f:
                    json.dump([], f)

                with open(SOLUTION_LIST_JSON, "w") as write_json:
                    entry = {"solution": solution, "shortsolution": solution}
                    SOL_FINDER.append(entry)
                    json.dump(SOL_FINDER, write_json)

        else:
            logger.debug("Mail Solution: " + str(solution))

        # Extracting PN from subject
        if 'PN Extension' in mail_subject:
            folder_exists = 'Yes'
        else:
            folder_exists = 'No'

        sample_project = mail_subject.split("/")
        pn_list = list()
        required_pn_list = list()
        num_dec_list = list()
        pns = ''

        for each in range(0, len(sample_project)):
            project = sample_project[each]
            data = project.split()
            last_index = len(data)
            prj_num = 0
            while True:
                try:
                    last_index = last_index - 1
                    if len(data[last_index]) != 0:
                        pn = int(data[last_index])
                        pn_list.append(str(pn))
                        found = True
                except Exception as e:
                    break

        if len(pn_list) != 0:
            # Removing duplicates from pn_list
            required_pn = set(pn_list)
            required_pn = list(required_pn)
            required_pn.sort(key=int)

            for item in required_pn:
                item = float(item)
                required_pn = item / 100
                required_pn_list.append(required_pn)

        for item in required_pn_list:
            number_dec = str(item).split('.')
            num_dec_list.append(number_dec)

        pn_dict = dict()
        for pn, pn_part in num_dec_list:
            pn_dict.setdefault(pn, []).append(pn_part)

        for x in pn_dict:
            i = 0
            while i < len(pn_dict[x]):
                if pn_dict[x][i] == '0' or len(pn_dict[x][i]) == 1:
                    pn_dict[x][i] = pn_dict[x][i] + '0'
                if i == 0:
                    pns = pns + x + pn_dict[x][i]
                    i = i + 1
                else:
                    pns = pns + '-' + pn_dict[x][i]
                    i = i + 1
            if pns != '':
                pns = pns + ","

        # Iterate csv to find pns
        try:
            if pns == '':
                pn = None
                for each_line in range(0, len(mail_list)):
                    mail_rows = mail_list[each_line]
                    if " Project #:" in mail_rows:
                        rowIndex = each_line
                        try:
                            rowIndex = rowIndex + 18
                            if len(mail_list[rowIndex]) != 0:
                                pn = str(mail_list[rowIndex])
                                pn = pn.replace(' ', '')
                                pn = re.findall(r'\'([^]]*)\'', pn)
                        except Exception as e:
                            break

                for each in pn:
                    pns = each
                    pns = pns + ","

            if len(pns) != 0:
                logger.debug("Mail PNs: " + str(pns))
            else:
                logger.error("Mail PNs" + str(pns))
        except Exception as e:
            pns = None

        mail_data = {
            "datetime": str(mail_receive_time),
            "sender": str(mail_sender),
            "subject": str(mail_subject),
            "mnemonic": str(mnemonic),
            "date": str(required_date),
            "pn": str(pns),
            "solution": str(final_solution)}

        return mail_data, mail_attachment, folder_exists

    def get_next_mail(self):
        # Gets next email
        self.mail = self.messages.GetNext()

    def get_previous_mail(self):
        self.mail = self.messages.GetPrevious()

    def execute_component(self, sender_name, hs_build_df):
        check_mail = True
        output_data = list()

        # Checking Sender name and reading each email
        while check_mail:
            mail_sender = self.mail.SenderName
            if mail_sender == sender_name:
                logger.info("READING NEW EMAIL")
                mail_data, mail_attachment, folder_exists = OutlookMailReader.mail_reader(self)
                logger.debug("Time Sort: " + str(VERIFICATION_JSON["datetime"]) + " " + str(mail_data["datetime"]))

                # Save pdf to pdfs folder - temp location to read pdfs

                temp_path = None
                if mail_attachment is not None:
                    project_path = os.path.abspath(PDF_TEMP_DOWNLOAD_PATH)
                    temp_path = os.path.join(project_path, mail_attachment.FileName)
                    mail_attachment.SaveAsFile(temp_path)

                # Exception when pdf is not in a readable format/ scanned pdf
                try:
                    pdf_data = {
                        "date": None,
                        "clientName": None,
                        "solution": None,
                        "quantity": None,
                        "customerNumber": None,
                        "clientNumber": None}
                    pdf_data = pdf_extractor(temp_path, pdf_data)
                except Exception as e:
                    # print(str(e))
                    logger.error("PDF was not read - PDF should not be uploaded")
                    pdf_data = {
                        "date": 'None',
                        "clientName": 'None',
                        "solution": 'None',
                        "quantity": 'None',
                        "customerNumber": 'None',
                        "clientNumber": 'None'}

                # Call create path for pdf function to find the url needed and to create path
                site_url = None
                folder_name = None
                pdf_upload_path = None
                if temp_path is not None:
                    site_url, pdf_upload_path, folder_name = create_path_for_pdf(pdf_data=pdf_data, mail_data=mail_data)

                # Upload the pdf to upload_pdfs folder - Folder used by C# code to upload pdfs
                if time_sorter(VERIFICATION_JSON["datetime"], mail_data["datetime"]):

                    # Update excel log with all pdf and mail data
                    hs_build_df = hs_build_df.append(
                        {'Pdf date': pdf_data["date"], 'Pdf solution': pdf_data["solution"],
                         'Pdf client name ': pdf_data["clientName"],
                         'Pdf quantity': pdf_data["quantity"],
                         'Pdf customer num': pdf_data["customerNumber"],
                         'Pdf client num': pdf_data["clientNumber"],
                         'Mail datetime': mail_data["datetime"],
                         'Mail subject': mail_data["subject"],
                         'Mail mnemonic': mail_data["mnemonic"],
                         'Mail date': mail_data["date"],
                         'Mail pn': mail_data["pn"],
                         'Mail solution': mail_data["solution"],
                         'Site url': str(site_url),
                         'Folder': str(folder_name),
                         'PN Extension': str(folder_exists)}, ignore_index=True)
                    writer = ExcelWriter("HS_Build_Log" + "." + time.strftime("%Y-%m-%d") + ".xlsx")
                    hs_build_df.to_excel(writer, index=False)
                    writer.save()

                    if pdf_upload_path is not None:
                        project_path = os.path.abspath(PDF_UPLOADS_FOLDER)
                        temp_each_pdf_folder_path = os.path.join(project_path, mail_attachment.FileName)
                        if not os.path.exists(temp_each_pdf_folder_path):
                            os.mkdir(temp_each_pdf_folder_path)

                        # Save pdf to upload_pdfs folder
                        temp_pdf_storage_path = os.path.join(temp_each_pdf_folder_path, mail_attachment.FileName)
                        mail_attachment.SaveAsFile(temp_pdf_storage_path)

                        # Create a text file in upload_pdfs folder with folder url and folder name
                        file = open(temp_each_pdf_folder_path + '/' + "url.txt", "w")
                        file.write("URL:" + ' ' + site_url)
                        file.write('\n')
                        file.write("FOLDERURL:" + ' ' + pdf_upload_path)
                        file.write('\n')
                        file.write("FOLDERNAME:" + ' ' + folder_name)
                        file.write('\n')
                        file.write("FOLDEREXIST:" + ' ' + folder_exists)
                        file.close()

                        logger.debug("FILE UPLOADED TO UPLOAD PDF FOLDER")

                    else:
                        logger.error("Site URL is none - PDF was not uploaded")
                    output_data.append(mail_data)
                else:
                    check_mail = False

            # Gets next email
            OutlookMailReader.get_next_mail(self)
        return output_data
