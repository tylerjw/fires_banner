'''
About Dialog loosy based off Message Dialog and IDLE About window.

Author: Tyler Weaver
Date: 06/08/13
'''

from Tkinter import *

class AboutDialog:

    def __init__(self, master, text='', title='', author='', email='',
                 class_=None, relx=0.5, rely=0.3):
        self.master = master
        self.text = 'Author: ' + author + '\n\nSend all comments to:\nemail: ' + email + '\n' + text
        self.title = title
        self.class_= class_
        self.relx=relx
        self.rely=rely
        

    def show(self):
        if self.class_:
            self.root = Toplevel(slef.master, class_=self.class_)
        else:
            self.root = Toplevel(self.master)
        
        self.root.title('About ' + self.title)
        self.root.iconname('About ' + self.title)
        
        self.frame = Frame(self.root, padx=5, pady=5)

        self.mf = Frame(self.frame, borderwidth=2, relief=SUNKEN)
        self.title_message = Message(self.mf, text=self.title + '\n\n',
                                     aspect=400, fg='white', bg='grey50',
                                     font=('Cursor', 24, 'bold'),
                                     pady=5, padx=5, width=300)
        self.message = Message(self.mf, text=self.text + '\n\n\n\n',
                               aspect=400, fg='white', bg='grey50',
                               pady=5, padx=20, width=300)
        self.title_message.pack(expand=1, fill=BOTH)
        self.message.pack(expand=1, fill=BOTH)
        self.mf.pack()
        
        self.frame.pack()
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
    about_dialog = AboutDialog(root,title='Fires Banner', author= 'Tyler Weaver',
                               email='tyler.weaver@cjsotf-a.soccent.centcom.smil.mil\nemail: tylerjw@gmail.com',
                               text='\nPython 2.6.5\nTkinter: 8.5.3\nxlrd 0.9.2')
    Button(root, text='About', command=about_dialog.show).pack()
    Button(root, text='Quit', command=root.quit).pack()
    root.mainloop()
    
