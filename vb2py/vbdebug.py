# Created by Leo from: C:\Development\Python22\Lib\site-packages\vb2py\vb2py.leo

"""Set up logging"""

import vbclasses

import logger   # For logging output and debugging 
_vb_debug_log = logger.getLogger("vb2PyApp")

vbclasses.Debug._logger = _vb_debug_log
