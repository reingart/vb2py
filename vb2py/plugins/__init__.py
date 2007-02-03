# Created by Leo from: C:\Development\Python22\Lib\site-packages\vb2py\vb2py.leo

import glob
import os

# Depends who imports us ... don't understand this
try:
	from vb2py.utils import rootPath
except ImportError:
	from utils import rootPath	

mods = []
for fn in glob.glob(os.path.join(rootPath(), "plugins", "*.py")):
	name = os.path.splitext(os.path.basename(fn))[0]
	if not name.startswith("_"):
		mods.append(name)
