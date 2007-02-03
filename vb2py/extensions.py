# Created by Leo from: C:\Development\Python22\Lib\site-packages\vb2py\vb2py.leo

"""Base classes for all plug-ins"""

import glob
import os
import imp

from utils import rootPath
from config import VB2PYConfig

Config = VB2PYConfig()

import logger
log = logger.getLogger("PlugInLoader")

# << Plug-in functions >> (1 of 2)
def loadAllPlugins():
	"""Load all plug-ins from the plug-in directory and return a list of all the classes"""
	import plugins
	mods = []
	for mod in plugins.mods:
		log.info("Checking '%s' for plugins" % mod)
		#
		f = open(os.path.join(rootPath(), "plugins", "%s.py" % mod), "r")
		try:
			try:
				m = imp.load_module(mod, f, "Plugin-%s" % mod, ('*.py', 'r', 1))
			finally:
				f.close()
		except Exception, err:
			log.warn("Error importing '%s' (%s). Module skipped" % (mod, err))
			continue
		#
		for name in dir(m):
			cls = getattr(m, name)
			#import pdb; pdb.set_trace()
			try:
				is_plugin = cls.__is_plugin__
			except AttributeError:
				is_plugin = 0
			if is_plugin:
				try:
					p = cls()               
					log.info("Added new plug-in: '%s" % p.name)
					mods.append(p)
				except Exception, err:
					log.warn("Error creating plugin '%s' (%s). Class skipped" % (cls, err))
	#
	# Now sort
	mods.sort()                
	return mods
# << Plug-in functions >> (2 of 2)
def disableLogging():
	"""Disable logging in all plugins"""
	#
	# Disable the main logger
	log.setLevel(0)
	#
	# Now do so for pluging
	BasePlugin.logging_level = 0
# -- end -- << Plug-in functions >>
# << Plug-in classes >> (1 of 4)
class BasePlugin(object):
	"""All plug-ins should inherit from this base class or define __is_plugin__"""

	__is_plugin__ = 1 # Set to true if you want to be loaded plug-in

	system_plugin = 0 # True if you are a system plugin
	__enabled = 1   # If false the plugin will not be called
	order = 1000  # Determines order of execution. lower = earlier

	logging_level = int(Config["General", "LoggingLevel"])

	# << BasePlugin methods >> (1 of 6)
	def __init__(self):
		"""Initialize the plugin

		This method should always be called by subclasses as it is required to set up logging etc

		"""
		if not hasattr(self, "name"):
			self.name = self.__class__.__name__

		self.log = logger.getLogger(self.name)
		self.log.setLevel(self.logging_level)
	# << BasePlugin methods >> (2 of 6)
	def preProcessVBText(self, text):
		"""Process raw VB text prior to any conversion

		This method should return a new version of the text with any changes made
		to it. If there is no preprocessing required then do not define this method.

		"""
		return text
	# << BasePlugin methods >> (3 of 6)
	def postProcessPythonText(self, text):
		"""Process Python text following the conversion

		This method should return a new version of the text with any changes made
		to it. If there is no postprocessing required then do not define this method.

		"""
		return text
	# << BasePlugin methods >> (4 of 6)
	def disable(self):
		"""Disable the plugin"""
		self.__enabled = 0
	# << BasePlugin methods >> (5 of 6)
	def isEnabled(self):
		"""Return 1 if plugin is enabled"""
		return self.__enabled
	# << BasePlugin methods >> (6 of 6)
	def __cmp__(self, other):
		"""Used to allow plugins to be sorted to run in a certain order"""
		return cmp(self.order, other.order)
	# -- end -- << BasePlugin methods >>
# << Plug-in classes >> (2 of 4)
import re

class RETextMarkup(BasePlugin):
	"""A utility class to apply regular expression based text markup

	The plug-in allows simple re text replacements as a pre and post conversion
	passes simple by reading from lists of replacements defined as class methods.

	Users can simply create instances of their own classes to handle whatever markup
	they desire.

	"""

	name = "RETextMarkup"
	re_flags = 0 # Put re flags here if you need them

	# Define your patterns by assigning to these properties in the sub-class
	pre_process_patterns = ()
	post_process_patterns = ()

	# << RETextMarkup methods >> (1 of 3)
	def preProcessVBText(self, text):
		"""Process raw VB text prior to any conversion"""
		if self.pre_process_patterns:
			self.log.info("Processing pre patterns")
		return self.processText(text, self.pre_process_patterns)
	# << RETextMarkup methods >> (2 of 3)
	def postProcessPythonText(self, text):
		"""Process Python text following the conversion"""
		if self.post_process_patterns:
			self.log.info("Processing post patterns")
		return self.processText(text, self.post_process_patterns)
	# << RETextMarkup methods >> (3 of 3)
	def processText(self, text, patterns):
		"""Process the text and mark it up"""
		for re_pattern, replace in patterns:
			def doSub(match):
				self.log.info("Replacing '%s' with %s, %s" % (re_pattern, replace, match.groupdict()))
				return replace % match.groupdict()
			r = re.compile(re_pattern, self.re_flags)
			text = r.sub(doSub, text)
		return text
	# -- end -- << RETextMarkup methods >>
# << Plug-in classes >> (3 of 4)
import re

class RenderHookPlugin(BasePlugin):
	"""A utility plugin to hook a render method and apply markup after the render

	The plugin replaces the specified objects normal renderCode method with one which
	calls the plugins addMarkup method when it is complete.

	"""

	name = "RenderHookPlugin"
	hooked_class_name = None # Name of class should go here

	# << RenderHookPlugin methods >> (1 of 2)
	def __init__(self):
		"""Initialize the plugin

		This method should always be called by subclasses as it is required to set up logging etc

		"""
		super(RenderHookPlugin, self).__init__()
		#
		# Look for class and replace its renderAsCode method
		import parserclasses
		self.hooked_class = getattr(parserclasses, self.hooked_class_name)
		old_render_method = self.hooked_class.renderAsCode
		#
		def newRender(obj, indent=0):
			ret = old_render_method(obj, indent)
			return self.addMarkup(indent, ret)
		#    
		self.hooked_class.renderAsCode = newRender
	# << RenderHookPlugin methods >> (2 of 2)
	def addMarkup(self, indent, text):
		"""Add markup to the rendered text"""
		return text
	# -- end -- << RenderHookPlugin methods >>
# << Plug-in classes >> (4 of 4)
class SystemPlugin(BasePlugin):
	"""Special kind of plug-in which is used by the system and cannot be disabled"""

	system_plugin = 1

class SystemPluginREPlugin(RETextMarkup):
	"""Special kind of plug-in which is used by the system and cannot be disabled"""

	system_plugin = 1
# -- end -- << Plug-in classes >>

if __name__ == "__main__":
	loadAllPlugins()
