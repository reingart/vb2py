# Created by Leo from: C:\Development\Python22\Lib\site-packages\vb2py\vb2py.leo

"""Logging infrastructure"""

import logging
import fnmatch
from config import VB2PYConfig
Config = VB2PYConfig()

class VB2PYLogger(logging.StreamHandler):
	"""Logger which can do some interesting filtering"""

	allowed = [] # Loggers which can report
	blocked = [] # Loggers which can't report

	def filter(self, record):
		"""Filter logging events"""
		for allow in self.allowed:
			if fnmatch.fnmatch(record.name, allow) and not record.name in self.blocked:
				return 1

	def initConfiguration(self, conf):
		"""Initialize the configuration"""
		self.allowed = self._makeList(conf["Logging", "Allowed"])
		self.blocked = self._makeList(conf["Logging", "NotAllowed"])

	def _makeList(self, text):
		"""Make a list from a comma separted list of names"""
		names = text.split(",")
		return [name.strip() for name in names]

main_handler = VB2PYLogger()
main_handler.setFormatter(logging.Formatter(logging.BASIC_FORMAT))
main_handler.initConfiguration(Config)

def getLogger(name, level=None):
	"""Create a logger with the usual settings"""
	if level is None:
		level = int(Config["General", "LoggingLevel"])
	log = logging.getLogger(name)
	log.addHandler(main_handler)
	log.setLevel(level)
	return log
