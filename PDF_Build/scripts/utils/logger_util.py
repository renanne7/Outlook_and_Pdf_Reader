from collections import OrderedDict
from json import loads, dumps
from logging import Formatter, FileHandler, StreamHandler, getLogger, INFO


def logger(name, handler, recordfields=[], level=INFO):
    """
    A function to create logs
    :param name:
    :param handler:
    :param recordfields:
    :param level:
    :return: log
    """
    log = getLogger(name)
    textformatter = JSONFormatter(recordfields)
    handler.setFormatter(textformatter)
    log.addHandler(handler)
    log.setLevel(level)
    return log


def filelogger(logname, recordfields=[], filename='json.log', level=INFO):
    """A convenience function to return a JSON file logger for simple situations.

    Args:
         logname      :   The name of the logger - to allow for multiple logs, and levels of logs in an application
         recordfields :   The metadata fields to add to the JSON record created by the logger
         filename     :   The name of the file to be used in the logger
    Returns:
        A JSON file logger.
    """
    handler = FileHandler(filename, 'w')
    return logger(logname, handler, recordfields, level)


def streamlogger(logname, recordfields=[], outputstream=None, level=INFO):
    """A convenience function to return a JSON stream logger for simple situations.

        Args:
         logname      :   The name of the logger - to allow for multiple logs, and levels of logs in an application
         recordfields :   The metadata fields to add to the JSON record created by the logger
         outputstream :   The outputstream to be used by the logger. sys.stderr is used when outputstream is None.
    Returns:
        A JSON stream logger.
    """
    handler = StreamHandler(outputstream)
    return logger(logname, handler, recordfields, level)


def readJSONlog(logfile, filterfunction=(lambda x: True), customjson=None):
    """Iterate through a log file of JSON records and return a list of JSON records that meet the filterfunction.

    Args:
        logfile          : A file like object consisting of JSON records.
        filterfunction   : A function that returns True if the JSON record should be included in the output and False otherwise.
        customjson       : A decoder function to enable the loading of custom json objects
    Returns:
        A list of Python objects built from JSON records that passed the filterfunction.
    """
    json_records = []
    for x in logfile:
        # if the record in the logfile returns true from the filter function convert it to JSON and add it the records
        # to return
        rec = loads(x[:-1], object_hook=customjson)
        if filterfunction(rec):
            json_records.append(rec)
    return json_records


class JSONFormatter(Formatter):
    """The JSONFormatter class outputs Python log records in JSON format.

       JSONFormatter assumes that log record metadata fields are specified at the fomatter level as opposed to the
       record level. The specification of matadata fields at the formatter level allows for multiple handles to display
       differing levels of detail. For example, console log output might specify less detail to allow for quick problem
       triage while file log output generated from the same data may contain more detail for in-depth investigations.

       Attributes:
           recordfields  : A list of strings containing the names of metadata fields (see Python log record documentation
                           for details) to add to the JSON output. Metadata fields will be added to the JSON record in
                           the order specified in the recordfields list.
           customjson    : A JSONEncoder subclass to enable writing of custom JSON objects.
    """

    def __init__(self, recordfields=[], datefmt=None, customjson=None):
        """__init__ overrides the default constructor to accept a formatter specific list of metadata fields

        Args:
            recordfields : A list of strings referring to metadata fields on the record object. It can be empty.
                           The list of fields will be added to the JSON record created by the formatter.
        """
        Formatter.__init__(self, None, datefmt)
        self.recordfields = recordfields
        self.customjson = customjson

    def usesTime(self):
        """ Overridden from the ancestor to look for the asctime attribute in the recordfields attribute.

        The override is needed because of the change in design assumptions from the documentation for the logging
        module. The implementation in this object could be brittle if a new release changes the name or adds another
        time attribute.

        Returns:
            boolean : True if asctime is in self.recordfields, False otherwise.
        """
        return 'asctime' in self.recordfields

    def _formattime(self, record):
        if self.usesTime():
            record.asctime = self.formatTime(record, self.datefmt)

    def _getjsondata(self, record):
        """ combines any supplied recordfields with the log record msg field into an object to convert to JSON

            Args:
                record   : log record to output to JSON log
            Returns:
                An object to convert to JSON - either an ordered dict if recordfields are supplied or the record.msg attribute
        """
        if len(self.recordfields) > 0:
            fields = []
            for x in self.recordfields:
                fields.append((x, getattr(record, x)))
            fields.append(('msg', record.msg))
            # An OrderedDict is used to ensure that the converted data appears in the same order for every record
            return OrderedDict(fields)
        else:
            return record.msg

    def format(self, record):
        """overridden from the ancestor class to take a log record and output a JSON formatted string.

           Args:
               record    : log record to output to JSON log
           Returns:
               A JSON formatted string
        """
        self._formattime(record)
        json_data = self._getjsondata(record)
        formatted_json = dumps(json_data, cls=self.customjson)
        return formatted_json
