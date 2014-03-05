"""Read a type library and generate a list of objects, methods and properties

This is based on code in win32com.client.tlbbrowse

The output is restructured text which is then converted directly to HTML using
docutils.

"""

# << tlbrowse declarations >>
import win32ui
import win32con
import win32api
import string
import commctrl
import pythoncom
from pywin.mfc import dialog

error = "TypeLib browser internal error"

FRAMEDLG_STD = win32con.WS_CAPTION | win32con.WS_SYSMENU
SS_STD = win32con.WS_CHILD | win32con.WS_VISIBLE
BS_STD = SS_STD  | win32con.WS_TABSTOP
ES_STD = BS_STD | win32con.WS_BORDER
LBS_STD = ES_STD | win32con.LBS_NOTIFY | win32con.LBS_NOINTEGRALHEIGHT | win32con.WS_VSCROLL
CBS_STD = ES_STD | win32con.CBS_NOINTEGRALHEIGHT | win32con.WS_VSCROLL

typekindmap = {
    pythoncom.TKIND_ENUM : 'Enumeration',
    pythoncom.TKIND_RECORD : 'Record',
    pythoncom.TKIND_MODULE : 'Module',
    pythoncom.TKIND_INTERFACE : 'Interface',
    pythoncom.TKIND_DISPATCH : 'Dispatch',
    pythoncom.TKIND_COCLASS : 'CoClass',
    pythoncom.TKIND_ALIAS : 'Alias',
    pythoncom.TKIND_UNION : 'Union'
}

TypeBrowseDialog_Parent=dialog.Dialog
# -- end -- << tlbrowse declarations >>
# << tlbrowse methods >> (1 of 4)
class TypeBrowseDialog(TypeBrowseDialog_Parent):
    # << class TypeBrowseDialog declarations >>
    "Browse a type library"

    IDC_TYPELIST = 1000
    IDC_MEMBERLIST = 1001
    IDC_PARAMLIST = 1002
    IDC_LISTVIEW = 1003
    # -- end -- << class TypeBrowseDialog declarations >>
    # << class TypeBrowseDialog methods >> (1 of 18)
    def __init__(self, typefile = None):
        TypeBrowseDialog_Parent.__init__(self, self.GetTemplate())
        try:
            if typefile:
                self.tlb = pythoncom.LoadTypeLib(typefile)
            else:
                self.tlb = None
        except pythoncom.ole_error:
            self.MessageBox("The file does not contain type information")
            self.tlb = None
        self.HookCommand(self.CmdTypeListbox, self.IDC_TYPELIST)
        self.HookCommand(self.CmdMemberListbox, self.IDC_MEMBERLIST)

        self.members = []
    # << class TypeBrowseDialog methods >> (2 of 18)
    def OnAttachedObjectDeath(self):
        self.tlb = None
        self.typeinfo = None
        self.attr = None
        return TypeBrowseDialog_Parent.OnAttachedObjectDeath(self)
    # << class TypeBrowseDialog methods >> (3 of 18)
    def _SetupMenu(self):
        menu = win32ui.CreateMenu()
        flags=win32con.MF_STRING|win32con.MF_ENABLED
        menu.AppendMenu(flags, win32ui.ID_FILE_OPEN, "&Open...")
        menu.AppendMenu(flags, win32con.IDCANCEL, "&Close")
        mainMenu = win32ui.CreateMenu()
        mainMenu.AppendMenu(flags|win32con.MF_POPUP, menu.GetHandle(), "&File")
        self.SetMenu(mainMenu)
        self.HookCommand(self.OnFileOpen,win32ui.ID_FILE_OPEN)
    # << class TypeBrowseDialog methods >> (4 of 18)
    def OnFileOpen(self, id, code):
        openFlags = win32con.OFN_OVERWRITEPROMPT | win32con.OFN_FILEMUSTEXIST
        fspec = "Type Libraries (*.tlb, *.olb)|*.tlb;*.olb|OCX Files (*.ocx)|*.ocx|DLL's (*.dll)|*.dll|All Files (*.*)|*.*||"
        dlg = win32ui.CreateFileDialog(1, None, None, openFlags, fspec)
        if dlg.DoModal() == win32con.IDOK:
            try:
                self.tlb = pythoncom.LoadTypeLib(dlg.GetPathName())
            except pythoncom.ole_error:
                self.MessageBox("The file does not contain type information")
                self.tlb = None
            self._SetupTLB()
            self.dumpProperties()
    # << class TypeBrowseDialog methods >> (5 of 18)
    def OnInitDialog(self):
        self._SetupMenu()
        self.typelb = self.GetDlgItem(self.IDC_TYPELIST)
        self.memberlb = self.GetDlgItem(self.IDC_MEMBERLIST)
        self.paramlb = self.GetDlgItem(self.IDC_PARAMLIST)
        self.listview = self.GetDlgItem(self.IDC_LISTVIEW)

        # Setup the listview columns
        itemDetails = (commctrl.LVCFMT_LEFT, 100, "Item", 0)
        self.listview.InsertColumn(0, itemDetails)
        itemDetails = (commctrl.LVCFMT_LEFT, 1024, "Details", 0)
        self.listview.InsertColumn(1, itemDetails)

        if self.tlb is None:
            self.OnFileOpen(None,None)
        else:
            self._SetupTLB()
        return TypeBrowseDialog_Parent.OnInitDialog(self)
    # << class TypeBrowseDialog methods >> (6 of 18)
    def _SetupTLB(self):
        self.typelb.ResetContent()
        self.memberlb.ResetContent()
        self.paramlb.ResetContent()
        self.typeinfo = None
        self.attr = None
        if self.tlb is None: return
        items = self.getAll()
        for item in items:
            self.typelb.AddString(item)
    # << class TypeBrowseDialog methods >> (7 of 18)
    def _SetListviewTextItems(self, items):
        self.listview.DeleteAllItems()
        index = -1
        for item in items:
            index = self.listview.InsertItem(index+1,item[0])
            data = item[1]
            if data is None: data = ""
            self.listview.SetItemText(index, 1, data)
    # << class TypeBrowseDialog methods >> (8 of 18)
    def SetupAllInfoTypes(self):
        infos = self._GetMainInfoTypes() + self._GetMethodInfoTypes()
        self._SetListviewTextItems(infos)
    # << class TypeBrowseDialog methods >> (9 of 18)
    def _GetMainInfoTypes(self):
        pos = self.typelb.GetCurSel()
        if pos<0: return []
        docinfo = self.tlb.GetDocumentation(pos)
        infos = [('GUID', str(self.attr[0]))]
        infos.append(('Help File', docinfo[3]))
        infos.append(('Help Context', str(docinfo[2])))
        try:
            infos.append(('Type Kind', typekindmap[self.tlb.GetTypeInfoType(pos)]))
        except:
            pass

        info = self.tlb.GetTypeInfo(pos)
        attr = info.GetTypeAttr()
        infos.append(('Attributes', str(attr)))

        for j in range(attr[8]):
            flags = info.GetImplTypeFlags(j)
            refInfo = info.GetRefTypeInfo(info.GetRefTypeOfImplType(j))
            doc = refInfo.GetDocumentation(-1)
            attr = refInfo.GetTypeAttr()
            typeKind = attr[5]
            typeFlags = attr[11]

            desc = doc[0]
            desc = desc + ", Flags=0x%x, typeKind=0x%x, typeFlags=0x%x" % (flags, typeKind, typeFlags)
            if flags & pythoncom.IMPLTYPEFLAG_FSOURCE:
                desc = desc + "(Source)"
            infos.append( 'Implements' + desc)

        return infos
    # << class TypeBrowseDialog methods >> (10 of 18)
    def _GetMethodInfoTypes(self):
        pos = self.memberlb.GetCurSel()
        if pos<0: return []

        realPos, isMethod = self._GetRealMemberPos(pos)
        ret = []
        if isMethod:
            funcDesc = self.typeinfo.GetFuncDesc(realPos)
            id = funcDesc[0]
            ret.append(("Func Desc", str(funcDesc)))
        else:
            id = self.typeinfo.GetVarDesc(realPos)[0]

        docinfo = self.typeinfo.GetDocumentation(id)
        ret.append(('Help String', docinfo[1]))
        ret.append(('Help Context', str(docinfo[2])))
        return ret
    # << class TypeBrowseDialog methods >> (11 of 18)
    def CmdTypeListbox(self, id, code):
        if code == win32con.LBN_SELCHANGE:
            pos = self.typelb.GetCurSel()
            attributes = self.getTypes(pos)
            self.memberlb.ResetContent()
            for attribute in attributes:
                self.memberlb.AddString(attribute)
    # << class TypeBrowseDialog methods >> (12 of 18)
    def _GetRealMemberPos(self, pos):
        if pos >= self.attr[7]:
            return pos - self.attr[7], 1
        elif pos >= 0:
            return pos, 0
        else:
            raise error, "The position is not valid"
    # << class TypeBrowseDialog methods >> (13 of 18)
    def CmdMemberListbox(self, id, code):
        if code == win32con.LBN_SELCHANGE:
            self.paramlb.ResetContent()
            pos = self.memberlb.GetCurSel()
            names = self.getMembers(pos)
            for i in range(len(names)):
                self.paramlb.AddString(names[i])
        self.SetupAllInfoTypes()
        return 1
    # << class TypeBrowseDialog methods >> (14 of 18)
    def GetTemplate(self):
        "Return the template used to create this dialog"

        w = 272  # Dialog width
        h = 192  # Dialog height
        style = FRAMEDLG_STD | win32con.WS_VISIBLE | win32con.DS_SETFONT | win32con.WS_MINIMIZEBOX
        template = [['Type Library Browser', (0, 0, w, h), style, None, (8, 'Helv')], ]
        template.append([130, "&Type", -1, (10, 10, 62, 9), SS_STD | win32con.SS_LEFT])
        template.append([131, None, self.IDC_TYPELIST, (10, 20, 80, 80), LBS_STD])
        template.append([130, "&Members", -1, (100, 10, 62, 9), SS_STD | win32con.SS_LEFT])
        template.append([131, None, self.IDC_MEMBERLIST, (100, 20, 80, 80), LBS_STD])
        template.append([130, "&Parameters", -1, (190, 10, 62, 9), SS_STD | win32con.SS_LEFT])
        template.append([131, None, self.IDC_PARAMLIST, (190, 20, 75, 80), LBS_STD])

        lvStyle = SS_STD | commctrl.LVS_REPORT | commctrl.LVS_AUTOARRANGE | commctrl.LVS_ALIGNLEFT | win32con.WS_BORDER | win32con.WS_TABSTOP
        template.append(["SysListView32", "", self.IDC_LISTVIEW, (10, 110, 255, 65), lvStyle])

        return template
    # << class TypeBrowseDialog methods >> (15 of 18)
    def dumpProperties(self):
        """Dump all objects and properties

            Objects seem to appear as,

                MyControl - an empty reference
                _MyControl - a reference with all the attributes and sometimes some methods
                _MyControlEvents - a reference holding extra methods

            """
        controls = {}
        for i, item in enumerate(self.getAll()):
            #
            # Create a new control object or get the old one if this is a _Control
            if item.startswith("_"):
                item = item[1:]
            if item.endswith("Events"):
                item = item[:-6]
            try:
                this = controls[item]
            except KeyError:
                this = VBControlObject(item)
                controls[item] = this
            #
            # Find all the types for this object
            types = self.getTypes(i)
            for j, typ in enumerate(types):
                members = self.getMembers(j)
                if members == (None,):
                    # This is an attribute
                    this.addProperty(typ)
                else:
                    # This is a method
                    thismethod = VBMethod(typ)
                    this.addMethod(thismethod)
                    for member in members:
                        thismethod.parameters.append(member)
        return controls
    # << class TypeBrowseDialog methods >> (16 of 18)
    def getTypes(self, pos):
        "Process a single type from the library"
        attributes = []
        if pos >= 0:
            self.typeinfo = self.tlb.GetTypeInfo(pos)
            self.attr = self.typeinfo.GetTypeAttr()
            for i in range(self.attr[7]):
                id = self.typeinfo.GetVarDesc(i)[0]
                attributes.append(self.typeinfo.GetNames(id)[0])
            for i in range(self.attr[6]):
                id = self.typeinfo.GetFuncDesc(i)[0]
                attributes.append(self.typeinfo.GetNames(id)[0])
            #self.SetupAllInfoTypes()
        return attributes
    # << class TypeBrowseDialog methods >> (17 of 18)
    def getMembers(self, pos):
        "Get all the members of this object"	
        try:	
            realPos, isMethod = self._GetRealMemberPos(pos)
        except:
            print "oops!"
            return [None]
        if isMethod:
            id = self.typeinfo.GetFuncDesc(realPos)[0]
            names = self.typeinfo.GetNames(id)
            return names[1:]
        else:
            return [None]
    # << class TypeBrowseDialog methods >> (18 of 18)
    def getAll(self):
        "Get all the types"
        all = []
        n = self.tlb.GetTypeInfoCount()
        for i in range(n):
            all.append(self.tlb.GetDocumentation(i)[0])
        return all
    # -- end -- << class TypeBrowseDialog methods >>
# << tlbrowse methods >> (2 of 4)
class VBControlObject(object):
    """Represents a VB control object"""

    # << VBControlObject methods >> (1 of 4)
    def __init__(self, name, properties=None, methods=None):
        """Initialize the control"""
        self.name = name
        self.properties = properties or {}
        self.methods = methods or {}
    # << VBControlObject methods >> (2 of 4)
    def addProperty(self, name):
        """Add a property"""
        self.properties[name] = name
    # << VBControlObject methods >> (3 of 4)
    def addMethod(self, method):
        """Add a method"""
        self.methods[method.name] = method
    # << VBControlObject methods >> (4 of 4)
    def __repr__(self):
        """Representation of this control"""
        return "VBControlObject('%s', %s, %s)" % (self.name, self.properties, self.methods)
    # -- end -- << VBControlObject methods >>
# << tlbrowse methods >> (3 of 4)
class VBMethod(object):
    """Represents a method of a control object"""

    # << VBMethod methods >> (1 of 3)
    def __init__(self, name, parameters=None):
        """Initialize the method"""
        self.name = name
        self.parameters = parameters or []
    # << VBMethod methods >> (2 of 3)
    def __repr__(self):
        """Representation of this method"""
        return "VBMethod('%s', %s)" % (self.name, self.parameters)
    # << VBMethod methods >> (3 of 3)
    def __str__(self):
    """String representation of this method"""
    try:
        if self.parameters:
            return "%s(*%s*)" % (self.name, ", ".join([str(item) for item in self.parameters]))
        else:
            return "%s()" % (self.name,)
    except:
        print "Failed on ", self.name
    # -- end -- << VBMethod methods >>
# << tlbrowse methods >> (4 of 4)
def enumerate(lst):
    "Mimic 2.3's enumerate function"
    return zip(xrange(sys.maxint), lst)
# -- end -- << tlbrowse methods >>


if __name__=='__main__':
    import sys
    fname = None
    try:
        fname = sys.argv[1]
    except:
        pass

    dlg = TypeBrowseDialog(fname)	
    if fname:
        controls = dlg.dumpProperties()
    else:
        dlg.DoModal()
        sys.exit()

    results = []
    l = results.append

    l("All VB Controls in %s\n\n" % fname)
    l("\n".join(["* %s_" % control for control in controls]))
    l("\n\n")

    for control in controls:
        l("\n%s\n%s\n\n" % (control, "="*len(control)))

        properties = controls[control].properties.keys()
        properties.sort()
        l("    *Properties*\n%s" % ("\n\n".join(["        %s" % 
                    prop for prop in properties]),)	)						

        methods = controls[control].methods.keys()
        methods.sort()					
        l("\n    *Methods*\n%s" % ("\n\n".join(["        %s" % 
                    str(controls[control].methods[method]) for method in methods]),))

    text = "\n".join(results)

    # << Convert to HTML >>
    import StringIO
    rstFile = StringIO.StringIO(text)

    from docutils.core import Publisher
    from docutils.io import StringOutput, StringInput

    pub = Publisher()
    # Initialize the publisher
    pub.source = StringInput(source=text)
    pub.destination = StringOutput(encoding="utf-8")
    pub.set_reader('standalone', None, 'restructuredtext')
    pub.set_writer('html')
    output = pub.publish()

    print output
    # -- end -- << Convert to HTML >>
