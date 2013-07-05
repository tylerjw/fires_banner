'''
Fires Banner Calculator

Dependencies:
Python 2.6.5
xlrd-0.9.2
CoordConverter 0.2 (CGRS support)
xls_parser

Author: Tyler Weaver

Versions:
0.1(06/07/13): initial build, gui demo
0.2(06/08/13):
    pri/tic radio buttons
    parse xls file
    description of contact
    Select first jtac when tab out of mgrs
    output from data
    capitolize items
    key bindings
    invalid team error
    grid outside ao error
    jtac selection errors
    focus in selection on mgrs to end of string
    update TADM file
    pickle TADM file
    Update TADM file error messages
0.3(06/09/13):
    menu bar
    about dialog
    configure dialog
    help file
    pickle parse options

TODO: label row number windage
'''

version = "0.3"
date_updated = "8 June 13"
pickle_filename = "data.pkl"

from CoordConverter import CoordTranslator
from Tkinter import Frame, Button, Entry, Label, Radiobutton, Listbox
from Tkinter import IntVar, StringVar, Menu, Menubutton, SOLID, Toplevel
from Tkinter import TOP, LEFT, BOTTOM, GROOVE, RIDGE, END, YES, BOTH, W, NW, X
import tkFileDialog
from ScrolledText import ScrolledText
from xls_parser import build_fires_dicts, ExcelFileOptions
from cPickle import dump, load
import os
from xlrd import XLRDError
from about_dialog import AboutDialog
from scrolledtext_dialog import ScrolledTextDialog

class FiresBannerBuilder(Frame):
    def __init__(self):
        Frame.__init__(self)
        
        self.pri_tic = IntVar()
        self.team = StringVar()
        self.mgrs = StringVar()
        self.description = StringVar()
        self.filename = ''
        
        self.mgrs.set("42S")
        
        self.pack(expand=YES, fill=BOTH)
        self.master.title("Fires Banner Builder v" + version)

        self.about_dialog = AboutDialog(self,title='Fires Banner', author= 'Tyler Weaver',
                               email='tyler.weaver@cjsotf-a.soccent.centcom.smil.mil\nemail: tylerjw@gmail.com',
                               text='\nVersion: ' + version + '\nDate Updated: ' + date_updated +
                                        '\n\nDependencies:\nPython 2.6.5\nTkinter: 8.5.3\nxlrd 0.9.2',relx=1.5,rely=0.4)

        ## menu
        menu = Frame(self, borderwidth=2, relief=GROOVE)
        menu.pack(side=TOP, fill=X)

        file_m = Menubutton(menu, text="File", underline=0)
        file_m.pack(side=LEFT)
        file_m.menu = Menu(file_m)
        file_m.menu.add_command(label="Update TADM File", command=self.update_dialog)
        file_m.menu.add_command(label="Quit", command=self.quit)
        file_m['menu'] = file_m.menu
        
        edit_m = Menubutton(menu, text="Edit", padx=15, underline=0)
        edit_m.pack(side=LEFT)
        edit_m.menu = Menu(edit_m)
        edit_m.menu.add_command(label='Options', command=self.option_dialog)
        edit_m['menu'] = edit_m.menu

        help_m = Menubutton(menu, text="Help", underline=0)
        help_m.pack(side=LEFT)
        help_m.menu = Menu(help_m)
        help_m.menu.add_command(label='Help', command=self.open_help)
        help_m.menu.add_command(label='About', command=self.about_dialog.show)
        help_m['menu'] = help_m.menu

        menu.tk_menuBar(file_m,edit_m,help_m)

        ## spacer frame
        space = Frame(self, height=20).pack(side=TOP)
        
        ## input frame
        inf = Frame(self, relief=SOLID, borderwidth=2)
        ##  radio button frame
        rbf = Frame(inf)
        Radiobutton(rbf, text="Priority", value=0, variable=self.pri_tic).pack(side=LEFT)
        Radiobutton(rbf, text="TIC", value=1, variable=self.pri_tic).pack(side=LEFT)
        self.pri_tic.set(0)
        rbf.pack(side=TOP)
        ##  team and mgrs frame
        tmf = Frame(inf)
        Label(tmf, text="Team:").pack(side=LEFT, pady=10, padx=5)
        self.team_ent = Entry(tmf, textvariable=self.team, width=7)
        self.team_ent.pack(side=LEFT, padx=5)
        ## mgrs
        Label(tmf, text="  MGRS:").pack(side=LEFT, pady=10, padx=5)
        self.mgrs_ent = Entry(tmf, textvariable=self.mgrs, width=18)
        self.mgrs_ent.pack(side=LEFT, padx=5)
        tmf.pack(side=TOP)
        ## description of contact frame
        dcf = Frame(inf)
        Label(dcf, text="Description of Contact:").pack(side=TOP, anchor=W)
        Entry(dcf, textvariable=self.description, width=40).pack(side=TOP)
        dcf.pack(side=TOP)
        
        inf.pack(side=TOP, ipadx=10, ipady=5, pady=10)

        Label(self, text="JTAC C/S").pack(side=TOP)
        self.jtac_list = Listbox(self, width=20, height=3)
        self.jtac_list.insert(END,"Input Team")
        self.jtac_list.pack(side=TOP)
        
        ## buttons
        btnf = Frame(self)

        self.calc_btn = Button(btnf, text="Calculate CAS Banner", command=self.calculate)
        self.calc_btn.pack(side=LEFT, pady=10, padx=5)

        btnf.pack(side=TOP)

        self.scrolled_text = ScrolledText(self, height=6, width=40,
                                          padx=10, pady=10,
                                          wrap="char")
        self.scrolled_text.pack(side=TOP, padx=10, pady=10)

        #bindings
        self.team_ent.bind('<FocusOut>', self.update_jtac_list)
        self.team_ent.bind('<Key-Return>', self.update_jtac_list)
        self.mgrs_ent.bind('<FocusOut>', lambda e: self.jtac_list.selection_set(0))
        self.mgrs_ent.bind('<FocusIn>', self.mgrs_focus_in)

        try:
            fd = open(pickle_filename, "rb")
            (self.jtacs, self.teams, self.filename, self.file_options) = load(fd)
            fd.close()
        except IOError:
            self.update_stext("Error: loading Pickle file, update TDAM file")
            self.initialize_options()
            self.update_dialog()

    def open_help(self):
        os.system('fires_banner_help.pdf')
        
    def option_dialog(self):
        od = self.OptionsDialog(self)
        od.show()
    
    def update_dialog(self):
        self.filename = tkFileDialog.askopenfilename(parent=self,
                                                  filetypes=[('Excel', '*.xls *.xlsx')],
                                                title="Select TDAM Excel Spreadsheet",
                                                multiple=False)
        try:
            (self.jtacs, self.teams) = build_fires_dicts(self.filename, self.file_options)
            fd = open(pickle_filename, "wb")
            dump((self.jtacs, self.teams, self.filename, self.file_options), fd, -1)
            fd.close()
            self.update_stext("TADM File Updated")
        except IOError:
            self.update_stext("Error: Update TADM File Error")
        except XLRDError:
            self.update_stext("Error: XLRD Error, Confirm selected file")

    def update_stext(self, text):
        self.scrolled_text.delete(1.0, END)
        self.scrolled_text.insert(END, text)

    def initialize_options(self):
        self.file_options = ExcelFileOptions()
    
    def mgrs_focus_in(self, event):
        print 'mgrs focus in'
        # frist test length ( nothing entered yet )
        if len(self.mgrs_ent.get()) < 4:
            self.mgrs_ent.selection_clear()
            
    def update_jtac_list(self, event):
        team = self.team.get().strip()

        #validate team
        if team in self.teams.keys():

            #empty list
            item = self.jtac_list.get(0)
            while item:
                self.jtac_list.delete(0)
                item = self.jtac_list.get(0)

            #add items to list
            for jtac in self.teams[team]:
                self.jtac_list.insert(END, jtac)

            self.jtac_list.selection_set(0)

        else:
            # invalid team
            self.scrolled_text.delete(1.0, END)
            self.scrolled_text.insert(END, "Error: Invalid Team")
        
    def calculate(self):
        ct = CoordTranslator()
        output = ""
        self.mgrs.set(self.mgrs.get().upper())
        self.description.set(self.description.get().upper())
        try:
            cgrs = ct.AsCGRS(self.mgrs.get())
            if cgrs == "": #outside AO
                output = "Error: MGRS outside OEF AO"
        except:
            output = "Error: Invalid MGRS"

        #get jtac and validate
        jtac_idx = self.jtac_list.curselection()
        jtac = ''
        if not jtac_idx:
            output = "Error: Select JTAC"
        else:
            jtac = self.jtac_list.get(jtac_idx[0])
        
        if jtac == "Input Team":
            output = "Error: Set team then select JTAC"

        #validate description
        if self.description.get() == '':
            output = "Error: Set Description of Contact"
        
        if output == "":
            #TODO: calculate values based off team

            if(self.pri_tic.get() == 0):
                output = "PRI/"
            else:
                output = "TIC/"

            #JTAC C/S (column G)            
            output += self.jtacs[jtac]['JTAC CS'] + "/"

            #P FREQ NAME (column J)
            output += self.jtacs[jtac]['P FREQ'] + "/"

            #A FREQ(column K) + NAME(column L)
            output += self.jtacs[jtac]['A FREQ'] + "/"

            #Collection Point MGRS
            output += "CP " + self.mgrs.get() + "/"

            #Collection Point CGRS
            output += "CP " + cgrs + "/"

            #MGRS
            output += self.mgrs.get() + "/"

            #CGRS
            output += cgrs + "/"

            #Description of Contact
            output += self.description.get() + "/"

            #Type of CAS
            output += "REQ ROVER CAPABLE CAS W/ OEF ROE ATT/"

            # KEG + JTAR# (column O)
            output += self.jtacs[jtac]['JTAR']

        self.scrolled_text.delete(1.0, END)
        self.scrolled_text.insert(END, output)

    class OptionsDialog:
        def __init__(self, master, class_=None, relx=0.5, rely=0.3):
            self.master = master
            self.title = 'Fires Banner Options'
            self.class_= class_
            self.relx=relx
            self.rely=rely

            ## defaults
            self.defaults = ExcelFileOptions()

            ## variables
            self.filename=StringVar(value=master.filename)
            self.sh_name=StringVar(value=master.file_options.sh_name)
            self.label_row=StringVar(value=str(master.file_options.label_row + 1))
            self.team_label=StringVar(value=master.file_options.team_label)
            self.pfreq_label=StringVar(value=master.file_options.pfreq_label)
            self.afreq_label=StringVar(value=master.file_options.afreq_label)
            self.jtar_label=StringVar(value=master.file_options.jtar_label)
            self.ignore_label=StringVar(value=', '.join(master.file_options.ignore_label))
            self.jtac_label=StringVar(value=master.file_options.jtac_label) 
            self.jtar_prefix=StringVar(value=master.file_options.jtar_prefix)

        def show(self):
            if self.class_:
                self.root = Toplevel(slef.master, class_=self.class_)
            else:
                self.root = Toplevel(self.master)
            
            self.root.title(self.title)
            self.root.iconname(self.title)
            
            self.frame = Frame(self.root, padx=5, pady=5)

            #### dialog content
            frame1 = Frame(self.frame)
            Label(frame1, text="Sheet Name: ").grid(row=0, sticky=W, ipady=2)
            Label(frame1, text="Label row #: ").grid(row=1, sticky=W, ipady=2)
            Label(frame1, text="Team Label: ").grid(row=2, sticky=W, ipady=2)
            Label(frame1, text="P Freqency Label: ").grid(row=3, sticky=W, ipady=2)
            Label(frame1, text="A Freqency Label: ").grid(row=4, sticky=W, ipady=2)
            Label(frame1, text="JTAC C/S Label: ").grid(row=5, sticky=W, ipady=2)
            Label(frame1, text="Ignore Labels: ").grid(row=6, sticky=W, ipady=2)
            Label(frame1, text="JTAR Label: ").grid(row=7, sticky=W, ipady=2)
            Label(frame1, text="JTAR Prefix: ").grid(row=8, sticky=W, ipady=2)

            Label(frame1, text="", width=20).grid(row=0, column=1)

            self.sheet_en = Entry(frame1, width=40, textvariable=self.sh_name)
            self.label_en = Entry(frame1, width=40, textvariable=self.label_row)
            self.team_en = Entry(frame1, width=40, textvariable=self.team_label)
            self.pfreq_en = Entry(frame1, width=40, textvariable=self.pfreq_label)
            self.afreq_en = Entry(frame1, width=40, textvariable=self.afreq_label)
            self.jtac_en = Entry(frame1, width=40, textvariable=self.jtac_label)
            self.ignore_en = Entry(frame1, width=40, textvariable=self.ignore_label)
            self.jtarl_en = Entry(frame1, width=40, textvariable=self.jtar_label)
            self.jtarp_en = Entry(frame1, width=40, textvariable=self.jtar_prefix)

            self.sheet_en.grid(row=0, column=2, sticky=W)
            self.label_en.grid(row=1, column=2, sticky=W)
            self.team_en.grid(row=2, column=2, sticky=W)
            self.pfreq_en.grid(row=3, column=2, sticky=W)
            self.afreq_en.grid(row=4, column=2, sticky=W)
            self.jtac_en.grid(row=5, column=2, sticky=W)
            self.ignore_en.grid(row=6, column=2, sticky=W)
            self.jtarl_en.grid(row=7, column=2, sticky=W)
            self.jtarp_en.grid(row=8, column=2, sticky=W)

            frame1.pack(side=TOP)

            frame2 = Frame(self.frame)
            Label(frame2, text="Current File:").pack(side=LEFT)
            self.filename_en = Entry(frame2, textvariable=self.filename)
            self.filename_en.pack(side=LEFT, fill=X, expand=YES)
            Button(frame2, text="Browse", command=self.get_filename).pack(side=LEFT)
            frame2.pack(side=TOP, pady=5, fill=X)
            ####

            ### buttons
            frame3 = Frame(self.frame)
            Button(self.frame, text='Test', width=10,
                        command=self.test).pack(side=LEFT, expand=YES)
            Button(self.frame, text='Restore Defaults',
                        command=self.restore_defaults).pack(side=LEFT, expand=YES)
            Button(self.frame, text='Cancel', width=10,
                       command=self.wm_delete_window).pack(side=LEFT, expand=YES)
            Button(self.frame, text='Apply', width=10,
                        command=self.apply_changes).pack(side=LEFT, expand=YES)

            frame3.pack(side=TOP, pady=5)
            self.frame.pack()

            ### enable
            self.root.protocol('WM_DELETE_WINDOW', self.wm_delete_window)
            self._set_transient(self.relx, self.rely)
            self.root.wait_visibility()
            self.root.grab_set()
            self.root.mainloop()
            self.root.destroy()

        def test(self):
            op = ExcelFileOptions()
            op.sh_name = self.sh_name.get()
            op.label_row = int(self.label_row.get()) - 1
            op.team_label = self.team_label.get()
            op.pfreq_label = self.pfreq_label.get()
            op.afreq_label = self.afreq_label.get()
            op.jtar_label = self.jtar_label.get()
            op.ignore_label = self.ignore_label.get()
            op.jtac_label = self.jtac_label.get()
            op.jtar_prefix = self.jtar_prefix.get()

            jtacs = {}
            teams = {}
            output = ''
            
            try:
                (jtacs, teams) = build_fires_dicts(self.filename.get(), op)
                for key in teams:
                    output += key + str(teams[key]) + '\n'
                output += '\n'
                
                for key in jtacs:
                    output += key + str(jtacs[key]) + '\n'
            except IOError:
                output = 'File Parse Error!'
            except ValueError:
                output = 'Check your labels and file, file parse failed.'

            self.std = ScrolledTextDialog(self.root,text=output,title='test')
            self.std.show()

        def restore_defaults(self):
            self.sh_name.set('SOTF-EAST PACE')
            self.label_row.set('2')
            self.team_label.set('3/3 ODA')
            self.pfreq_label.set('P FREQ')
            self.afreq_label.set('A FREQ')
            self.jtar_label.set('JTAR #')
            self.ignore_label.set('LATE DEPLOY')
            self.jtac_label.set('JTAC C/S' )
            self.jtar_prefix.set('KEG')
            

        def apply_changes(self):
            print "apply"
            self.master.file_options.sh_name = self.sh_name.get()
            self.master.file_options.label_row = int(self.label_row.get()) - 1
            self.master.file_options.team_label = self.team_label.get()
            self.master.file_options.pfreq_label = self.pfreq_label.get()
            self.master.file_options.afreq_label = self.afreq_label.get()
            self.master.file_options.jtar_label = self.jtar_label.get()
            self.master.file_options.ignore_label = self.ignore_label.get().split(',')
            self.master.file_options.jtac_label = self.jtac_label.get()
            self.master.file_options.jtar_prefix = self.jtar_prefix.get()
            self.master.filename = self.filename.get()

            try:
                (self.master.jtacs, self.master.teams) = build_fires_dicts(self.master.filename, self.master.file_options)
                fd = open(pickle_filename, "wb")
                dump((self.master.jtacs, self.master.teams, self.master.filename, self.master.file_options), fd, -1)
                fd.close()
                self.master.update_stext("TADM File Updated")
            except IOError:
                self.master.update_stext("Error: Update TADM File Error")
            except XLRDError:
                self.master.update_stext("Error: XLRD Error, Confirm selected file")

            self.wm_delete_window()

        def get_filename(self):
            self.filename.set(tkFileDialog.askopenfilename(parent=self.master,
                                                  filetypes=[('Excel', '*.xls *.xlsx')],
                                                title="Select TDAM Excel Spreadsheet",
                                                multiple=False))

        def _set_transient(self, relx=0.5, rely=0.3):
            widget = self.root
            widget.withdraw() # Remain invisible while we figure out the geometry
            widget.transient(self.master)
            widget.update_idletasks() # Actualize geometry information
            if self.master.winfo_ismapped():
                m_width = self.master.winfo_width()
                m_height = self.master.winfo_height()
                m_x = self.master.winfo_rootx()
                m_y = self.master.winfo_rooty()
            else:
                m_width = self.master.winfo_screenwidth()
                m_height = self.master.winfo_screenheight()
                m_x = m_y = 0
            w_width = widget.winfo_reqwidth()
            w_height = widget.winfo_reqheight()
            x = m_x + (m_width - w_width) * relx
            y = m_y + (m_height - w_height) * rely
            if x+w_width > self.master.winfo_screenwidth():
                x = self.master.winfo_screenwidth() - w_width
            elif x < 0:
                x = 0
            if y+w_height > self.master.winfo_screenheight():
                y = self.master.winfo_screenheight() - w_height
            elif y < 0:
                y = 0
            widget.geometry("+%d+%d" % (x, y))
            widget.deiconify() # Become visible at the desired location

        def wm_delete_window(self):
            self.root.quit()

if __name__ == "__main__":
    FiresBannerBuilder().mainloop()
