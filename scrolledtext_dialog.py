from ScrolledText import ScrolledText
from Tkinter import *

class ScrolledTextDialog:
    def __init__(self, master, text='', title='',
                 class_=None, relx=0.5, rely=0.3):
        self.master = master
        self.text = text
        self.title = title
        self.class_= class_
        self.relx=relx
        self.rely=rely
        

    def show(self):
        if self.class_:
            self.root = Toplevel(slef.master, class_=self.class_)
        else:
            self.root = Toplevel(self.master)
        
        self.root.title(self.title)
        self.root.iconname(self.title)
        
        self.frame = Frame(self.root, padx=5, pady=5)

        self.stf = Frame(self.frame)
        
        stext = ScrolledText(self.stf, bg='white', height=20, width=120)
        stext.insert(END, self.text)
        stext.pack(fill=BOTH, side=LEFT, expand=True)
        stext.focus_set()

        self.stf.pack(expand=True, fill=BOTH)
        self.frame.pack(expand=True, fill=BOTH)
        self.root.bind('<Return>', self.return_evt)
        
        b = Button(self.frame, text='Close', width=10,
                   command=self.wm_delete_window)
        b.pack(side=TOP, pady=5)
        
        self.root.protocol('WM_DELETE_WINDOW', self.wm_delete_window)
        self._set_transient(self.relx, self.rely)
        self.root.wait_visibility()
        self.root.grab_set()
        self.root.mainloop()
        self.root.destroy()

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

    def return_evt(self, event):
        self.root.quit()

    
if __name__=='__main__':
    root = Tk()
    text = 'This is a few lines\n of random text\n that i put in here to test this scrolled text dialog. thank you all.  Tyler'
    about_dialog = ScrolledTextDialog(root,text=text,
                                      title='test')
    Button(root, text='About', command=about_dialog.show).pack()
    Button(root, text='Quit', command=root.quit).pack()
    root.mainloop()
    
