#!/usr/bin/env python
# -*- coding: utf-8 -*-
__version__ = '1.0'
__date__ = 'Date: 2018-12-28'
__author__ = 'Rene Salih'

import os
import pandas as pd
from scripts.utils.pm_logging import get_logger
from scripts.utils.pm_utils import update_new_latest_time
from scripts.utils.outlook_mail_reader import OutlookMailReader
from scripts.constants.pm_constants import PDF_TEMP_DOWNLOAD_PATH

logger = get_logger()


def main():
    logger.info("STARTING THE HS BUILD APPLICATION")
    obj = OutlookMailReader(recipient_name="My Lead XYZ")
    hs_build_df = pd.DataFrame()
    # data = obj.execute_component('Salih,Rene', hs_build_df)
    data = obj.execute_component('Project Number Notification', hs_build_df)
    update_new_latest_time(uploaded_data=data)
    for each_file in (os.listdir(PDF_TEMP_DOWNLOAD_PATH)):
        os.remove(os.path.join(PDF_TEMP_DOWNLOAD_PATH, each_file))


if __name__ == '__main__':
    main()
