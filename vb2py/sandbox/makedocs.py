# Created by Leo from: C:\Development\Python22\Lib\site-packages\vb2py\vb2py.leo

import vb2py.vbparser
import re
import glob
import vb2py.utils
import os
import sys

#
# Turn off logging 
import vb2py.extensions
vb2py.extensions.disableLogging()
vb2py.vbparser.log.setLevel(0) # Don't print all logging stuff

# << fn doAutomaticVBConversion >>
def doAutomaticVBConversion(txt):
	"""Convert VB code in the text to Python"""
	vb = re.compile(r'(.*?)<p>VB(\(.*?\))?:</p>.*?<pre class="literal-block">(.*?)</pre>(.*?)', 
					re.DOTALL+re.MULTILINE)
	def convertVB(match):
		"""Convert the match"""
		if match.groups()[1]:
			mod = getattr(vb2py.vbparser, match.groups()[1][1:-1])()
		else:
			mod = vb2py.vbparser.VBCodeModule()
		mod.importStatements = lambda x : ""
		m = vb2py.vbparser.parseVB(match.groups()[2], container=mod)
		return '%s<table style="code-table">' \
			   '<tr><th class="code-header">VB</th><th class="code-header">Python</th></tr>' \
			   '<tr><td class="vb-code-cell"><pre>%s</pre></td>' \
			   '<td class="python-code-cell"><pre>%s</pre></td></tr></table>%s' % (
										   match.groups()[0],
										   match.groups()[2].replace("\n", "<br>"),
										   m.renderAsCode(1).replace("\n", "<br>"),
										   match.groups()[3])
	return vb.sub(convertVB, txt)
# -- end -- << fn doAutomaticVBConversion >>
# << fn addToTemplate >>
def addToTemplate(text, template, token):
	"""Add the text to the template"""
	template_text = open(template, "r").read()
	template_text = template_text.replace(token, text)
	return template_text
# -- end -- << fn addToTemplate >>

if __name__ == "__main__":
	#
	add_to_template = 0
	if len(sys.argv) == 1:
		pattern = "*.htm"
	else:
		pattern = sys.argv[1]
	#
	print "\nvb2Py documentation generator\n"
	for fn in glob.glob(os.path.join(vb2py.utils.rootPath(), "doc", pattern)):
		print "Processing '%s' ... " % fn,
		txt = open(fn, "r").read()
		marked_up_text = doAutomaticVBConversion(txt)
		#
		if add_to_template:
			marked_up_text = addToTemplate(marked_up_text, sys.argv[2], "DOCSGOHERE")
		#
		open(fn+"l", "w").write(marked_up_text)
		print "Done!"
