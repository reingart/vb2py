#@+leo-ver=4-thin
#@+node:pap.20070203162417.1:@thin vb2pyMain.py
#@<< vb2pyUI declarations >>
#@+node:pap.20070203162417.2:<< vb2pyUI declarations >>
"""
__version__ = "$Revision: 1.3 $"
__date__ = "$Date: 2004/08/12 19:14:23 $"
"""

from PythonCard import model

import interactive
#@-node:pap.20070203162417.2:<< vb2pyUI declarations >>
#@nl
#@+others
#@+node:pap.20070203162417.3:class UI
class UI(model.Background):
    """vb2Py UI for doing conversion"""
    
    #@	@+others
    #@+node:pap.20070203172243:on_GoInteractive_mouseClick
    def on_GoInteractive_mouseClick(self, event):
        """Hit the interactive button"""
        form = model.childWindow(self, interactive.Interactive)
        form.visible = True
    #@-node:pap.20070203172243:on_GoInteractive_mouseClick
    #@-others
#@-node:pap.20070203162417.3:class UI
#@-others

if __name__ == '__main__':
    app = model.Application(UI)
    app.MainLoop()
#@-node:pap.20070203162417.1:@thin vb2pyMain.py
#@-leo
