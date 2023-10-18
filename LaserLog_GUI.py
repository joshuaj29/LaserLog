'''
Created on Tue Aug  8 14:26:18 2022

@author: alexs, joshuaj

Copy paste to install exe:
    1. !cd C:/Users/Alexs/.spyder-py3/Production Log Test/laser log exe
    2. !pyinstaller -D --hidden-import "babel.numbers" laserLogGUI.py
or use:
    auto-py-to-exe run from (with no spaces) c: \ users \ joshuaj \ appdata \ local
        \ programs \ python \ python310 \ scripts \ auto-py-to-exe
'''

import tkinter as tk
from tkinter import messagebox as mb
from tkinter import filedialog
from tkcalendar import DateEntry
import os
import pandas as pd
import re
#needed this import for exe to run
import babel.numbers





class LaserLog(tk.Frame):
    def __init__(self, parent, ops):
        tk.Frame.__init__(self, parent)

        self.df = pd.DataFrame()
        self.colNames = ['Work-Order','Tool-Number','Drill-Layers','Panel-Count','Hole-Count','Drill-Time']
        
        self.bggray = self.rgbtohex(41, 39, 39)
        self.fgwhite = self.rgbtohex(201, 201, 201)
        self.pressgray = self.rgbtohex(125, 120, 120)
        self.green = self.rgbtohex(99, 232, 74)
        self.red = self.rgbtohex(232, 74, 74)
        self.blue = self.rgbtohex(52, 125, 235)
        
        self.parent = parent
        self.parent.title('Laser Log')
        self.parent.geometry('%dx%d' % (500, 600))
        self.parent.minsize(width=400, height=500)
        self.parent.maxsize(width=800, height=900)
        self.parent.config(bg=self.bggray)
        self.parent.grid_columnconfigure(0, weight=1)
        self.parent.grid_rowconfigure(0, weight=1)
        
        self.grid(padx=10, pady=10, sticky='news')
        
        # Title frame (grid row 0) with title label within
        self.titleframe = tk.Frame(self, borderwidth=5, relief='groove', bg=self.bggray)
        self.titleframe.grid(row=0, column=0, sticky='news')
        self.grid_columnconfigure(0, weight=1)

        for i in range(4): 
            self.grid_rowconfigure(i, weight=1)
        
        self.title_label = tk.Label(self.titleframe, text='LASER DRILL\nProduction Log Entry Form', font='Calibri 18 bold',
                                    bg=self.bggray, fg='white')
        self.title_label.grid(row=0, column=0, sticky='news')
        self.titleframe.grid_columnconfigure(0, weight=1)
        self.titleframe.grid_rowconfigure(0, weight=1)
        
        self.dates_frame = tk.Frame(self, borderwidth=5, bg=self.bggray)
        self.dates_frame.grid(row=1, column=0, sticky='news')

        for i in range(1): 
            self.dates_frame.grid_rowconfigure(i, weight=1)
        for i in range(2): 
            self.dates_frame.grid_columnconfigure(i, weight=1)
        
        self.before_complete = tk.StringVar()
        
        self.before_complete.set('No Date Selected')
        
        self.completedate_btn = self.new_button(self.dates_frame, 'Complete-Date', self.bggray, 'white', self.pressgray,
                                             self.set_complete_date, fnt='Calibri 16')
        

        
        self.complete_lbl = tk.Label(self.dates_frame, textvariable=self.before_complete, bg=self.bggray, fg=self.red, font='Calibri 16')
        
        self.completedate_btn.grid(row=0, column=0, sticky='news')
        self.complete_lbl.grid(row=0, column=1, sticky='news')
        

        
        # Entry frame (grid row 2) with entry form within it
        self.entryframe = tk.LabelFrame(self, text='', bg=self.bggray, fg=self.fgwhite,
                                        borderwidth=3, relief='groove')
        self.entryframe.grid(row=2, column=0, sticky='news')

        for i in range(7): 
            self.entryframe.grid_rowconfigure(i, weight=1)
        for i in range(2): 
            self.entryframe.grid_columnconfigure(i, weight=1)
        
        self.entry_prompts = ['Work-Order',
                              'Tool-Number',
                              'Drill-Layers',
                              'Panel-Count',
                              'Hole-Count',
                              'Drill-Time (1 Panel)']
        self.operators = ops
        self.filepath = None
        
        self.dd2text = tk.StringVar()
        
        self.dd2text.set('')
        
        self.full_entries = self.get_date_entry() + self.get_entry_form()
        
        self.response = []
        
        self.button_frame = tk.Frame(self, borderwidth=5, bg=self.bggray)
        self.button_frame.grid(row=3, column=0, sticky='news')
        self.button_frame.grid_rowconfigure(0, weight=1)

        for i in range(3): 
            self.button_frame.grid_columnconfigure(i, weight=1)
        
        self.input_button = self.new_button(self.button_frame, 'Input Values', self.bggray, self.green, self.pressgray,
                                            self.confirm_enter_values, fnt='Calibri 16')
        self.load_button = self.new_button(self.button_frame, 'Load Previous', self.bggray, self.blue, self.pressgray,
                                            self.load_previous, fnt='Calibri 16')
        self.quit_button = self.new_button(self.button_frame, 'Exit to Main Menu', self.bggray, self.red, self.pressgray,
                                           self.quit, fnt='Calibri 16')
        
        
        self.input_button.grid(row=0, column=0, sticky='news', padx=5)
        self.load_button.grid(row=0, column=1, sticky='news', padx=5)
        self.quit_button.grid(row=0, column=2, sticky='news', padx=5)
        
        self.center(parent)
    




    def new_button(self, mstr, txt, bgc, fgc, pressc, cmd, fnt='Calibri 13'):
        btn = tk.Button(master=mstr, text=txt, bg=bgc, fg=fgc, activebackground=pressc, command=cmd, font=fnt)
        return btn
    




    def rgbtohex(self, r, g, b):
        return f'#{r:02x}{g:02x}{b:02x}'
    




    def set_complete_date(self):
        
        def fetch(master):
            self.before_complete.set(calen.get_date())
            self.complete_lbl.config(fg=self.green)
            master.destroy()
        
        top = tk.Toplevel(self.parent)
        
        top.geometry('375x200+%d+%d' % (self.parent.winfo_x() + (self.parent.winfo_width() / 4),
                                        self.parent.winfo_y() + (self.parent.winfo_height() / 2)))
        top.config(bg=self.bggray)
        
        f1 = tk.Frame(top, relief='groove', borderwidth=2, bg=self.bggray)
        f1.pack(fill='both', expand=False, padx=10, pady=10)
        
        f2 = tk.Frame(top, bg=self.bggray)
        f2.pack(padx=10, pady=10)
        
        instructions = tk.Label(f1, text='Select when the job was completed'
                                         + '\nYou may change your selection only before submitting the log'
                                         + '\nClick OK when you have made the correct selection',
                                         bg=self.bggray, fg=self.fgwhite)
        instructions.pack(fill='both')
        
        calen = DateEntry(f1, selectmode='day', bg=self.bggray)
        calen.pack(pady=10)
        
        submit_btn = self.new_button(f2, 'OK', self.bggray, self.green, self.pressgray, lambda: fetch(top))
        submit_btn.pack(padx=10, pady=10)
        

    



    def get_date_entry(self):
        return [('Start Date', self.before_complete)]
    



    def center(self, parent):
        parent.update_idletasks()
        
        width = parent.winfo_width()
        frm_width = parent.winfo_rootx() - parent.winfo_x()
        win_width = width + 2 * frm_width
        
        height = parent.winfo_height()
        titlebar_height = parent.winfo_rooty() - parent.winfo_y()
        win_height = height + titlebar_height + frm_width
        
        x = parent.winfo_screenwidth() // 2 - win_width // 2
        y = parent.winfo_screenheight() // 2 - win_height // 2
        
        parent.geometry('{}x{}+{}+{}'.format(width, height, x, y))
    



    def get_entry_form(self):
        rows = []
        ind = list(enumerate(self.entry_prompts))
        
        for i, prompt in ind:
            lbl = tk.Label(self.entryframe, text=prompt, bg=self.bggray, fg='white', font='Calibri 13')
            entry = tk.Entry(self.entryframe, bg=self.bggray, fg='white', justify='center', font='Calibri 13')

            lbl.grid(row=i, column=0, sticky='e', pady=1)
            entry.grid(row=i, column=1, sticky='news', pady=1)
            
            rows.append((prompt, entry))


        
        dd2 = tk.OptionMenu(self.entryframe, self.dd2text, *self.operators)
        dd2.config(bg=self.bggray,
                   fg='white',
                   activebackground=self.pressgray,
                   highlightthickness=0,
                   indicatoron=0,
                   font='Calibri 13')
        dd2lbl = tk.Label(self.entryframe, text='Operator', bg=self.bggray, fg='white', font='Calibri 13')
        

        
        dd2lbl.grid(row=ind[-1][0]+1, column=0, sticky='e', pady=1)
        dd2.grid(row=ind[-1][0]+1, column=1, sticky='news')
        
        rows.append((dd2lbl.cget('text'), self.dd2text))
        
        return rows
    



    def fetch_response(self):
        response_submission = [None] * len(self.full_entries)
        
        for row in enumerate(self.full_entries):
            i = row[0]
            prompt = row[1][0]
            entry = row[1][1].get()

            response_submission[i] = ((prompt, entry))
        
        return response_submission
    



    def regexCheck(self,entry,pattern):
        regex = re.compile(pattern)
        mo = regex.search(entry)
        if mo:
            return True
        else:
            return False



    def num_numCheck(self, entry):
        return self.regexCheck(entry, r'^\d+-\d+$')
    
    def workOCheck(self,entry):
        return self.regexCheck(entry, r'^(\w+\d+-)?\d+-\d+(-\d+)?$')

    def numCheck(self, entry):
        return self.regexCheck(entry, r'^\d+$')

    def timeCheck(self, entry):
        return self.regexCheck(entry, r'^\d+:\d+$')

    def noneSelected(self, entry):
        invalid_entries = ['', 'No Date Selected']
        if entry in invalid_entries:
            return False
        else:
            return True
        
    def errorText(self, check, erText):
        if not check:
            return erText
        

    def errorList(self, entry, errorList, funcMeth, field):
        check = self.noneSelected(entry)
        if check:
            check = funcMeth(entry)
            errorList.append(self.errorText(check, f'Incorrect {field} Format'))
        else:
            errorList.append(self.errorText(check, f'Enter {field}'))





    def check_entries(self):
        entryError = []
        
        for row in self.full_entries:
            tab = row[0]
            entry = row[1].get()

            if tab == 'Start Date':
                stCheck = self.noneSelected(entry)
                entryError.append(self.errorText(stCheck, 'No Date Selected'))

            elif tab == 'Work-Order':
                self.errorList(entry, entryError, self.workOCheck, 'Work Order')
                

            elif tab == 'Tool-Number':
                self.errorList(entry, entryError, self.numCheck, 'Tool-Number')


            elif tab == 'Drill-Layers':
                self.errorList(entry, entryError, self.num_numCheck, 'Drill-Layers')


            elif tab == 'Panel-Count':
                self.errorList(entry, entryError, self.numCheck, 'Panel-Count',)


            elif tab == 'Hole-Count':
                self.errorList(entry, entryError, self.numCheck, 'Hole-Count')


            elif tab == 'Drill-Time (1 Panel)':
                self.errorList(entry, entryError, self.timeCheck, 'Drill-Time')


            elif tab == 'Operator':
                OCheck = self.noneSelected(entry)
                entryError.append(self.errorText(OCheck, 'No Operator Selected'))


        if not all(entry is None for entry in entryError):
            return False, entryError
        else:
            return True, entryError
    
    



    def clear_entries(self):
        
        self.before_complete.set('No Date Selected')
        
        self.complete_lbl.config(fg='red')
        
        self.dd2text.set('')
        
        for row in self.full_entries:
            if isinstance(row[1], tk.Entry):
                row[1].delete(0, tk.END)
                row[1].insert(0, '')
            



    def append_df_to_excel(self, df, excel_path):
        df_excel = pd.read_excel(excel_path)
        result = pd.concat([df_excel, df], ignore_index=True)
        result.to_excel(excel_path, index=False)
    



    def confirm_enter_values(self):
        entry_check = self.check_entries()
        
        if(entry_check[0]):
            idir = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
            self.response = self.fetch_response()
            self.filepath = filedialog.askopenfilename(parent=self, initialdir=idir)
            
            if(self.filepath != ''):
                d = [None] * len(self.response)
                
                for i in range(0, len(self.response)):
                    d[i] = self.response[i][1]
                
                self.df = pd.DataFrame(data={'Start-Date':[d[0]], 'Work-Order':[d[1]],
                                        'Tool-Number':[d[2]], 'Drill-Layers':[d[3]], 'Panel-Count':[d[4]],
                                        'Hole-Count':[d[5]], 'Drill-Time':[d[6]], 'Operator':[d[7]]})

                #print(self.df)
                self.append_df_to_excel(self.df, self.filepath)
                mb.showinfo(title='Success', message='Values stored in target file', parent=self)
                self.clear_entries()

        else:
            errorMessage = ''
            for err in entry_check[1]:
                if err == None:
                    continue
                else:
                    errorMessage += err + '\n'

            mb.showerror(title='Entry error', message=errorMessage, parent=self)




    def load_previous(self):
        if (self.response != []) & (not(self.df.empty)):
            lastDate = self.df.at[0,'Start-Date']
            lastOps = self.df.at[0,'Operator']

            self.clear_entries()
            self.before_complete.set(lastDate)
            
            self.complete_lbl.config(fg=self.green)
            
            self.dd2text.set(lastOps)
            
            i=0
            for row in self.full_entries:
                if isinstance(row[1], tk.Entry):
                    currentColumn = self.df.at[0,self.colNames[i]]
                    row[1].insert(0, f'{currentColumn}')
                    i+=1
        
        else:
            mb.showinfo(title='Load-Previous', message='No Previous Values to Load', parent=self)
                    



    
    def quit(self):
        self.parent.destroy()






if __name__ == '__main__':
    root = tk.Tk()
    LaserLog(root, ['CN', 'JJ', 'MS', 'TS', 'VS'])
    root.mainloop()
