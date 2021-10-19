""" Logger that logs to multiple places at once.

Writes to the console, a regular .log file, and a database events table (for
errors only).

This class implements the python logging module's main methods, at least the
ones used so far.

TODO: Add more methods as needed. See https://docs.python.org/3/library/logging.html
"""

import logging
from logging.handlers import RotatingFileHandler


class AppLogger:
    def __init__(self, proc_name):
        """ Constructor (creates log file and connects to the database).
        """

        self.proc_name = proc_name.upper()

        self.logger = logging.getLogger(self.proc_name)
        handler = RotatingFileHandler(self.proc_name + '.log',
                                      maxBytes=1000000,
                                      backupCount=1)
        formatter = logging.Formatter('%(asctime)s [%(levelname)s] %(message)s',
                                      "%Y-%m-%d %H:%M:%S")
        handler.setFormatter(formatter)

        # Set the logging level
        self.logger.setLevel(logging.INFO)

        # Add the logger
        self.logger.addHandler(handler)

        # Connect to the database in which we log events
        #self.db = StatusDB()

    def info(self, text, *args):
        """ Log informational messages (console and log file only).
        """

        if args:
            # Handle logger.info('%s %s' % (a, b)) format
            if type(args[0]) is dict:
                args = (str(args[0]),)
            elif type(args[0]) is tuple:
                args = args[0]

            try:
                msg = text % args
            except TypeError:
                # Someone passed in as comma delimited not %s
                msg = text + ' ' + ' '.join(args)
        else:
            msg = text

        print(msg)
        self.logger.info(msg)

    def warning(self, text, *args):
        """ Log warning messages (console and log file only).
        """

        if args:
            # Handle logger.warning('%s %s' % (a, b)) format
            if type(args[0]) is dict:
                args = (str(args[0]),)
            elif type(args[0]) is tuple:
                args = args[0]

            try:
                msg = text % args
            except TypeError:
                print('error')
                # Someone passed in as comma delimited not %s
                msg = text + ' ' + ' '.join(args)
        else:
            msg = text

        print(msg)
        self.logger.warning(msg)

    def warn(self, text, *args):
        """ Support deprecated method (should use warning() instead).
        """

        return self.warning(text, args)

    def error(self, text, *args, details=None, doc_id=None, doc_type=None):
        """ Log error messages (to all destinations).
        """

        if args:
            # Handle logger.error('%s %s' % (a, b)) format
            if type(args[0]) is dict:
                args = (str(args[0]),)
            elif type(args[0]) is tuple:
                args = args[0]

            try:
                msg = text % args
            except TypeError:
                # Someone passed in as comma delimited not %s
                msg = text + ' ' + ' '.join(args)
        else:
            msg = text

        print('ERROR:', msg)
        self.logger.error(msg)

        # Also log to the events table
        #try:
        #    self.db.log_event(self.proc_name, msg, details=details,
        #                      doc_id=doc_id, doc_type=doc_type)
        #except Exception as e:
        #    print('WARNING: failed log event:', str(e))
