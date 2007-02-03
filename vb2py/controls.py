# Created by Leo from: C:\Development\Python22\Lib\site-packages\vb2py\vb2py.leo

twips_per_pixel = 15

# << SupportedControls >> (1 of 2)
class VBControl:
	"""Base class for all VB controls"""

	pycard_name = "VBControl"
	is_container = 0 # 1 for container classes like frames

	_attribute_translations = { 
				"Visible" : "visible",
				"BackColor" : "backgroundColor",
				"ForeColor" : "foregroundColor",
				"ToolTipText" : "toolTip",
				}

	# << class VBControl methods >> (1 of 11)
	def _getPropertyList(cls):
		"""Return a list of the properties of this control"""
		items = []
		for item in dir(cls):
			if not (item.startswith("_") or item.startswith("vbobj_")):
				items.append(item)
		return items

	_getPropertyList = classmethod(_getPropertyList)
	# << class VBControl methods >> (2 of 11)
	def _getControlList(cls):
		"""Return a list of the controls contained in this control"""
		items = []
		for item in dir(cls):
			if not item.startswith("_") and item.startswith("vbobj_"):
				items.append(item)
		return items

	_getControlList = classmethod(_getControlList)
	# << class VBControl methods >> (3 of 11)
	def _getControlsOfType(cls, type_name):
		"""Return a control with a given type"""
		lst = []
		for item in dir(cls):
			if not item.startswith("_") and item.startswith("vbobj_"):
				obj = cls._get(item)
				if obj.pycard_name == type_name:
					lst.append(obj)
		return lst

	_getControlsOfType = classmethod(_getControlsOfType)
	# << class VBControl methods >> (4 of 11)
	def _getContainerControls(cls):
		"""Return all container controls"""
		lst = []
		for item in dir(cls):
			if not item.startswith("_"):
				obj = cls._get(item)
				try:
					is_container = obj.is_container
				except AttributeError:
					pass
				else:
					if is_container:
						lst.append(obj)
		return lst

	_getContainerControls = classmethod(_getContainerControls)
	# << class VBControl methods >> (5 of 11)
	def _get(cls, name, default=None):
		"""Get one of our items"""
		try:
			return getattr(cls, name)
		except AttributeError:
			if default is not None:
				return default
			else:
				raise

	_get = classmethod(_get)
	# << class VBControl methods >> (6 of 11)
	def _realName(cls):
		"""Return our real name"""
		return cls.__name__[6:]

	_realName = classmethod(_realName)
	# << class VBControl methods >> (7 of 11)
	def _getPyCardEntry(cls):
		"""Return the pycard representation of this object"""
		#
		# Get dictionary entries for this object
		d = {}
		ret = [d]
		d['name'] = cls._realName()
		d['position'] = (cls.Left/twips_per_pixel, cls.Top/twips_per_pixel)
		d['type'] = cls.pycard_name
		#
		for attr, pycard_attr in cls._attribute_translations.iteritems():
			if hasattr(cls, attr):
				d[pycard_attr] = getattr(cls, attr)

		d.update(cls._getClassSpecificPyCardEntries())
		#
		# Watch out for container objects - we have to recur down them
		if cls.is_container:
			cls._processChildObjects()
			for cmp in cls._getControlList():
				obj = cls._get(cmp)
				entry = obj._getPyCardEntry()
				if entry:
					ret += entry
		return ret

	_getPyCardEntry = classmethod(_getPyCardEntry)
	# << class VBControl methods >> (8 of 11)
	def _getClassSpecificPyCardEntries(cls):
		"""Return additional items for this entry

		This method is normally overriden in the subclass

		"""
		return {}

	_getClassSpecificPyCardEntries = classmethod(_getClassSpecificPyCardEntries)
	# << class VBControl methods >> (9 of 11)
	def _mapNameReference(cls, match):
		"""Map a reference to this object in code to something meaningful

		We have two issues, scope and attributes. The scope we need to map to self.components
		That was easy. But we also need to map attributes. We can't do that in the base class
		but a subclass can hopefully help us out via the _attributeTranslationClass

		"""
		if match.groups()[0] is not None:
			return "self.components.%s.%s" % (cls._realName(), 
											  cls._attributeTranslation(match.groups()[0]))
		else:
			return "self.components.%s" % (cls._realName(),)


	_mapNameReference = classmethod(_mapNameReference)
	# << class VBControl methods >> (10 of 11)
	def _attributeTranslation(cls, name):
		"""Convert a VB attribute to a Python one"""
		try:
			return cls._attribute_translations[name]
		except KeyError:
			return VBControl._attribute_translations.get(name, name)

	_attributeTranslation = classmethod(_attributeTranslation)
	# << class VBControl methods >> (11 of 11)
	def _processChildObjects(cls):
		"""Before we deal with our child object we get a chance to do some processing

		Sub-classed can use this to do special things, like frames adjusting the
		properties of our children

		"""
		for container in cls._getContainerControls():
			container._processChildObject()

	_processChildObjects = classmethod(_processChildObjects)
	# -- end -- << class VBControl methods >>
# << SupportedControls >> (2 of 2)
# << Controls >> (1 of 11)
class Menu(VBControl):
	"""Menu"""
	pycard_name = "Menu"

	def _getPyCardEntry(cls):
		"""Return the pycard representation of this object"""
		return {}

	_getPyCardEntry = classmethod(_getPyCardEntry)

	def _pyCardMenuEntry(cls):		
		"""Return the entry for this menu"""
		return {
					"type" : "Menu",
					"name" : cls._realName(),
					"label" : cls.Caption,
			 }

	_pyCardMenuEntry = classmethod(_pyCardMenuEntry)
# << Controls >> (2 of 11)
class CommandButton(VBControl):
	pycard_name = "Button"

	def _getClassSpecificPyCardEntries(cls):
		"""Return additional items for this entry"""
		return {  "label" : cls.Caption,
				}

	_getClassSpecificPyCardEntries = classmethod(_getClassSpecificPyCardEntries)
# << Controls >> (3 of 11)
class ComboBox(VBControl):
	pycard_name = "ComboBox"

	def _getEntriesFromFRX(self, data):
		"""Get list entries from FRX file"""
		file, offset = data.split("@")
		raw = open(file, "r").read()
		ptr = int(offset, 16)
		num = ord(raw[ptr])
		ptr = ptr + 4
		lst = []
		for i in range(num):
			length = ord(raw[ptr])
			lst.append(raw[ptr+2:ptr+2+length])
			ptr += 2+length
		return lst

	_getEntriesFromFRX = classmethod(_getEntriesFromFRX)			


	def _getClassSpecificPyCardEntries(cls):
		"""Return additional items for this entry"""
		d = {"size" : (cls.Width/twips_per_pixel, cls.Height/twips_per_pixel)}
		if hasattr(cls, "List"):
			d["items"] =  cls._getEntriesFromFRX(cls.List)
		else:
			d["items"] = []
		return d

	_getClassSpecificPyCardEntries = classmethod(_getClassSpecificPyCardEntries)
# << Controls >> (4 of 11)
class ListBox(ComboBox):
	pycard_name = "List"
# << Controls >> (5 of 11)
class Label(VBControl):
	pycard_name = "StaticText"

	def _getClassSpecificPyCardEntries(cls):
		"""Return additional items for this entry"""
		return {  "text" : cls.Caption,
				}

	_getClassSpecificPyCardEntries = classmethod(_getClassSpecificPyCardEntries)
# << Controls >> (6 of 11)
class CheckBox(VBControl):
	pycard_name = "CheckBox"

	_attribute_translations = { 
					"Value" : "checked",
					}
	_attribute_translations.update(VBControl._attribute_translations)

	def _getClassSpecificPyCardEntries(cls):
		"""Return additional items for this entry"""
		return {  "label" : cls.Caption,
				  "checked" : cls._get("Value", 0),
				}

	_getClassSpecificPyCardEntries = classmethod(_getClassSpecificPyCardEntries)
# << Controls >> (7 of 11)
class TextBox(VBControl):
	pycard_name = "TextField"

	_attribute_translations = { 
					"Text" : "text",
					}
	_attribute_translations.update(VBControl._attribute_translations)

	def _getClassSpecificPyCardEntries(cls):
		"""Return additional items for this entry"""
		return {  "text" : cls.Text,
				}

	_getClassSpecificPyCardEntries = classmethod(_getClassSpecificPyCardEntries)
# << Controls >> (8 of 11)
class Form(VBControl):
	"""Form"""
# << Controls >> (9 of 11)
class Frame(VBControl):
	"""Frame"""
	pycard_name = "StaticBox"
	is_container = 1

	def _getClassSpecificPyCardEntries(cls):
		"""Return additional items for this entry"""
		return {  "size" : (cls.Width/twips_per_pixel, cls.Height/twips_per_pixel),
				}

	_getClassSpecificPyCardEntries = classmethod(_getClassSpecificPyCardEntries)


	def _processChildObjects(cls):
		"""Adjust the left and top properties of our children"""
		for item in cls._getControlList():
			obj = cls._get(item)
			print "Offsetting ", obj, cls
			if hasattr(obj, "Left"):
				obj.Left += cls.Left
			if hasattr(obj, "Top"):
				obj.Top += cls.Top
		#for container in cls._getContainerControls():
		#	container._processChildObjects()

	_processChildObjects = classmethod(_processChildObjects)
# << Controls >> (10 of 11)
class OptionButton(VBControl):
	pycard_name = "RadioGroup"

	def _getClassSpecificPyCardEntries(cls):
		"""Return additional items for this entry"""
		return {  "items" : cls.items,
				  "selected" : cls.selected,
				}

	_getClassSpecificPyCardEntries = classmethod(_getClassSpecificPyCardEntries)
# << Controls >> (11 of 11)
class VBUnknownControl(Label):
	"""A control representing an unknown control

	We fake it out as a label

	"""

	Caption = "Unknown control"
# -- end -- << Controls >>

possible_controls = { 
	"VBControl" : VBControl,
	"Form" : Form,
	"CommandButton" : CommandButton,
	"OptionButton" : OptionButton,
	"TextBox" : TextBox,
	"Label" : Label,
	"Menu" : Menu,
	"ComboBox" : ComboBox,
	"ListBox" : ListBox,
	"CheckBox" : CheckBox,
	"Frame" : Frame,

	"VBUnknownControl" : VBUnknownControl,
	"FileListBox" : VBUnknownControl,
	"Timer" : VBUnknownControl,
	"OLE" : VBUnknownControl,
	"Shape" : VBUnknownControl,
}
# -- end -- << SupportedControls >>
