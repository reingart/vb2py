{'stack':{'type':'Stack',
          'name':'Template',
    'backgrounds': [
    {'type':'Background',
          'name':'TestingView',
          'title':'Testing View',
          'position':(134, 220),
          'size':(1013, 613),
          'style':['resizeable'],

        'menubar': {'type':'MenuBar',
         'menus': [
             {'type':'Menu',
             'name':'menuFile',
             'label':'&File',
             'items': [
                  {'type':'MenuItem',
                   'name':'menuFileExit',
                   'label':'E&xit',
                  },
              ]
             },
             {'type':'Menu',
             'name':'menuTest',
             'label':'Test',
             'items': [
                  {'type':'MenuItem',
                   'name':'menuTestCode',
                   'label':'Test Code\tCtrl+T',
                  },
                  {'type':'MenuItem',
                   'name':'menuGeneratePythonScript',
                   'label':'Generate Python Script\tCtrl+P',
                  },
                  {'type':'MenuItem',
                   'name':'menuGenerateVBScript',
                   'label':'Generate VB Script\tCtrl+B',
                  },
              ]
             },
         ]
     },
         'components': [

{'type':'StaticText', 
    'name':'ResultLabel', 
    'position':(605, 3), 
    'size':(109, 17), 
    'text':'Results', 
    },

{'type':'StaticText', 
    'name':'PythonLabel', 
    'position':(322, 4), 
    'size':(155, 18), 
    'text':'Python Code', 
    },

{'type':'StaticText', 
    'name':'VBLabel', 
    'position':(11, 6), 
    'size':(113, 17), 
    'text':'VB Code', 
    },

{'type':'TextArea', 
    'name':'ResultsView', 
    'position':(599, 21), 
    'size':(328, 449), 
    'alignment':'left', 
    },

{'type':'CodeEditor', 
    'name':'PythonCodeEditor', 
    'position':(311, 23), 
    'size':(279, 452), 
    'backgroundColor':(255, 255, 255), 
    },

{'type':'CodeEditor', 
    'name':'VBCodeEditor', 
    'position':(13, 28), 
    'size':(296, 452), 
    'backgroundColor':(255, 255, 255), 
    },

] # end components
} # end background
] # end backgrounds
} }
