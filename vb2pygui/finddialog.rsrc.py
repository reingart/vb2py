#@+leo-ver=4
#@+node:@file finddialog.rsrc.py
{'type':'CustomDialog',
    'name':'Find',
    'title':'Find',
    'position':(467, 157),
    'size':(437, 164),
    'components': [

{'type':'Button', 
    'name':'btnFindNext', 
    'position':(347, 30), 
    'label':'Find Next', 
    },

{'type':'RadioGroup', 
    'name':'optSearchIn', 
    'position':(10, 40), 
    'size':(151, 47), 
    'items':['Python', 'VB'], 
    'label':'Search in', 
    'layout':'horizontal', 
    'max':1, 
    'selected':'Python', 
    },

{'type':'TextField', 
    'name':'txtFind', 
    'position':(67, 7), 
    'size':(272, -1), 
    },

{'type':'StaticText', 
    'name':'StaticText1', 
    'position':(11, 10), 
    'text':'Find what', 
    },

{'type':'Button', 
    'id':5100, 
    'name':'btnFind', 
    'position':(347, 6), 
    'label':'Find', 
    },

{'type':'Button', 
    'id':5101, 
    'name':'btnCancel', 
    'position':(347, 55), 
    'label':'Close', 
    },

] # end components
} # end CustomDialog
#@-node:@file finddialog.rsrc.py
#@-leo
