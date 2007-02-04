{'application':{'type':'Application',
          'name':'Template',
    'backgrounds': [
    {'type':'Background',
          'name':'bgTemplate',
          'title':'vb2Py Code Conversion',
          'size':(973, 836),
          'backgroundColor':(243, 243, 243),
          'icon':'vb2Py.ico',

         'components': [

{'type':'ImageButton', 
    'name':'HelpConversionStyle', 
    'position':(917, 17), 
    'size':(29, 29), 
    'border':'transparent', 
    'file':'images/help.jpg', 
    },

{'type':'ImageButton', 
    'name':'HelpConvertAs', 
    'position':(289, 15), 
    'size':(30, 29), 
    'border':'transparent', 
    'file':'images/help.jpg', 
    },

{'type':'Image', 
    'name':'PythonPane', 
    'position':(607, 679), 
    'file':'images/pythonpane.jpg', 
    },

{'type':'Image', 
    'name':'VBPane', 
    'position':(134, 679), 
    'file':'images/vbpane.jpg', 
    },

{'type':'RadioGroup', 
    'name':'Pythonicity', 
    'position':(585, 6), 
    'size':(327, -1), 
    'items':['Make sure it works', 'Make it look like Python'], 
    'label':'Conversion style', 
    'layout':'horizontal', 
    'max':1, 
    'stringSelection':'Make sure it works', 
    },

{'type':'RadioGroup', 
    'name':'CodeContext', 
    'position':(19, 5), 
    'items':['Code module', 'Class module', 'Form'], 
    'label':'Convert as', 
    'layout':'horizontal', 
    'max':1, 
    'stringSelection':'Code module', 
    },

{'type':'TextArea', 
    'name':'LogWindow', 
    'position':(15, 678), 
    'size':(914, 91), 
    'backgroundColor':(236, 236, 255, 255), 
    'visible':False, 
    },

{'type':'Button', 
    'name':'Convert', 
    'position':(438, 11), 
    'size':(-1, 35), 
    'label':'Convert -->', 
    },

{'type':'CodeEditor', 
    'name':'Python', 
    'position':(473, 60), 
    'size':(473, 615), 
    'backgroundColor':(255, 255, 255, 255), 
    },

{'type':'CodeEditor', 
    'name':'VB', 
    'position':(16, 60), 
    'size':(448, 615), 
    'backgroundColor':(255, 255, 255, 255), 
    },

] # end components
} # end background
] # end backgrounds
} }
