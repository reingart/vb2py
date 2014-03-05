"""Plug-in to convert recordset notation to standard notation

eg, 

RecordSet!FieldName 

Translates to,

RecordSet.Fields("FieldName").Value

This plugin scans code before parsing and replaces the short-form notation
with the full notation, which can be converted automatically.

Contributed by Alexandr Zamaraev
24 Aug 2004

"""


try:
    import vb2py.extensions as extensions
except ImportError:
    import extensions

class RecordSetShort(extensions.SystemPluginREPlugin):
  """Convert RecordSet!FieldName to RecordSet.Fields("FieldName").Value"""

  #name = 'RecordSetShort'
  __enabled = 1

  pre_process_patterns = (
    (r'(?P<RS>[\w\d_]+)!(?P<FN>[\w\d_]+)', r'%(RS)s.Fields("%(FN)s").Value'),
  )
