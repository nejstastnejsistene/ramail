#!/usr/bin/env python

import ConfigParser
import collections
import datetime
import re
import urllib
from email.mime.text import MIMEText

import Tkinter as tk

DEBUG = True

GMAIL_HOST = 'smtp.gmail.com:587'
DATE_FMT = '%m/%d/%y'

SUBJECT = 'Reminder: Outstanding {}'
if DEBUG:
    CC = 'mmpozulp@email.wm.edu ncschaaf@email.wm.edu'.split()
else:
    CC = 'jmgarc@wm.edu tnfeenstra@email.wm.edu'.split()

MESG = '''Hello {name},

This is a reminder from the RA on duty that the {item} you checked out on {start_date:{date_fmt}} {sub_mesg}. Regular duty office hours are from 7pm-11pm Sunday-Thursday and 7pm-12:30am on Friday and Saturday.

Thanks and have a great evening.'''

FUTURE = 'must be returned to the duty office by {due_date:{date_fmt}}'
PRESENT = 'must be returned to the duty office tonight'
PAST = 'was due on {due_date:{date_fmt}}. Please return it to the duty office ASAP'

def compose_message(name, item, start_date, due_date, date_fmt=DATE_FMT):
    now = datetime.datetime.now()
    if now.day < due_date.day:
        sub_mesg = FUTURE
    elif now.day > due_date.day:
        sub_mesg = PAST
    else:
        sub_mesg = PRESENT
    sub_mesg = sub_mesg.format(**locals())
    return MESG.format(**locals())

PATTERN = r'<td><a href="details[^>]*?>([^,]+?), ([^\s]+?) [^<]*?</a></td>\s+<td><a[^>]*?>([^<]+?)</a></td>'

Record = collections.namedtuple('Record', 'first last email'.split())

def directory_lookup(phrase, searchtype):
    url = 'http://directory.wm.edu/people/namelisting.cfm'
    if searchtype in 'first last'.split():
        query = {
            'searchtype': searchtype,
            'criteria': 'starts',
            'phrase': phrase,
            }
        data = urllib.urlencode(query)
        text = urllib.urlopen(url, data).read()
        pattern = re.compile(PATTERN)
        for lname, fname, email in pattern.findall(text):
            yield Record(fname, lname.title(), email)


def event_wrapper(func, *args, **kwargs):
    def newfunc(event):
        func(event, *args, **kwargs)
    return newfunc


class MainWindow(tk.Frame):

    def __init__(self, master=None):
        tk.Frame.__init__(self, master)

        self.error = tk.Label(self)
        self.error.grid(row=0, column=0, columnspan=2, sticky=tk.W+tk.E)

        label = tk.Label(self, text='First Name*')
        label.grid(row=1, column=0, sticky=tk.W)

        self.fnamevar = tk.StringVar()
        self.fname = tk.Entry(self, textvariable=self.fnamevar)
        handler = event_wrapper(self.update_names, 'first')
        self.fname.bind('<FocusOut>', handler)
        self.fname.bind('<Return>', handler)
        self.fname.grid(row=1, column=1, sticky=tk.W+tk.E)

        label = tk.Label(self, text='Last Name')
        label.grid(row=2, column=0, sticky=tk.W)

        self.lnamevar = tk.StringVar()
        lname = tk.Entry(self, textvariable=self.lnamevar)
        handler = event_wrapper(self.update_names, 'last')
        lname.bind('<FocusOut>', handler)
        lname.bind('<Return>', handler)
        lname.grid(row=2, column=1, sticky=tk.W+tk.E)

        label = tk.Label(self, text='Email Address*')
        label.grid(row=3, column=0, sticky=tk.W)

        self.emailvar = tk.StringVar()
        self.email = tk.Entry(self, textvariable=self.emailvar)
        self.email.bind('<FocusOut>', self.test_values)
        self.email.bind('<Return>', self.test_values)
        self.email.grid(row=3, column=1, sticky=tk.W+tk.E)

        label = tk.Label(self, text='Item*')
        label.grid(row=4, column=0, sticky=tk.W)

        self.itemvar = tk.StringVar()
        self.item = tk.Entry(self, textvariable=self.itemvar)
        self.item.bind('<FocusOut>', self.test_values)
        self.item.bind('<Return>', self.test_values)
        self.item.grid(row=4, column=1, sticky=tk.W+tk.E)

        label = tk.Label(self, text='Start Date*')
        label.grid(row=5, column=0, sticky=tk.W)

        values = []
        for i in range(-50, 50):
            values.append((datetime.datetime.now() + \
                    datetime.timedelta(days=i)).strftime(DATE_FMT))

        self.startvar = tk.StringVar()
        self.start = tk.Spinbox(self, values=values, textvariable=self.startvar)
        self.start.bind('<FocusOut>', self.test_values)
        self.start.bind('<Return>', self.test_values)
        self.start.grid(row=5, column=1, sticky=tk.E)

        label = tk.Label(self, text='Due Date*')
        label.grid(row=6, column=0, sticky=tk.W)

        self.duevar = tk.StringVar()
        self.due = tk.Spinbox(self, values=values, textvariable=self.duevar)
        self.due.bind('<FocusOut>', self.test_values)
        self.due.bind('<Return>', self.test_values)
        self.due.grid(row=6, column=1, sticky=tk.E)

        for i in range(50):
            self.start.invoke('buttonup')
            self.due.invoke('buttonup')

        button = tk.Button(self, text='Draft Email')
        button.bind('<Button-1>', self.compose)
        button.bind('<Return>', self.compose)
        button.grid(row=7, column=0, columnspan=2, sticky=tk.W+tk.E)

        label = tk.Label(self, text='Select a name:')
        label.grid(row=0, column=2)

        self.names = {}
        self.lbox = tk.Listbox(self)
        self.lbox.bind('<<ListboxSelect>>', self.select_name)
        self.lbox.grid(row=1, rowspan=6, column=2)

        self.grid()

    def update_names(self, event, searchtype):
        if event.widget is self.fname:
            self.test_values(event)
        text = event.widget.get().lower()
        if self.lnamevar.get() or self.fnamevar.get():
            if self.lbox.size():
                self.lbox.delete(0, tk.END)
                for name in sorted(k for (k, v) in self.names.items() \
                        if getattr(v, searchtype).lower().startswith(text)):
                    self.lbox.insert(tk.END, name)
            if not self.lbox.size():
                self.names = {}
                try:
                    for record in directory_lookup(text, searchtype):
                        key = record.first + ' ' + record.last
                        self.names[key] = record
                        if not record.first.lower().startswith(self.fnamevar.get().lower()):
                            continue
                        if not record.last.lower().startswith(self.lnamevar.get().lower()):
                            continue
                        self.lbox.insert(0, key)
                except IOError:
                    self.lbox.config(bg='#ff0000')


    def select_name(self, event):
        record = self.names[self.lbox.get(int(event.widget.curselection()[0]))]
        self.fnamevar.set(record.first)
        self.lnamevar.set(record.last)
        self.emailvar.set(record.email)

    def test_values(self, event=None):
        def should_test(*widgets):
            return event is None or event.widget in widgets
        error = None

        name = self.fnamevar.get()
        if should_test(self.fname) and not name:
            self.fname.config(bg='#ff0000')
            if not error:
                error = 'First name is required.'
        else:
            self.fname.config(bg='#ffffff')

        email = self.emailvar.get()
        if should_test(self.email) and not email:
            self.email.config(bg='#ff0000')
            if not error:
                error = 'Email is required.'
        else:
            self.email.config(bg='#ffffff')

        item = self.itemvar.get()
        if should_test(self.item) and not item:
            self.item.config(bg='#ff0000')
            if not error:
                error = 'Item is required.'
        else:
            self.item.config(bg='#ffffff')

        try:
            start_date = datetime.datetime.strptime(self.startvar.get(), DATE_FMT)
            self.start.config(bg='#ffffff')
        except ValueError:
            if should_test(self.start, self.due):
                self.start.config(bg='#ff0000')
                if not error:
                    error = 'Date must be in format MM/DD/YY'

        try:
            due_date = datetime.datetime.strptime(self.duevar.get(), DATE_FMT)
            self.due.config(bg='#ffffff')
        except ValueError:
            if should_test(self.start, self.due):
                self.due.config(bg='#ff0000')
                if not error:
                    error = 'Date must be in format MM/DD/YY'

        try:
            if should_test(self.start, self.due) and due_date < start_date:
                self.start.config(bg='#ff0000')
                self.due.config(bg='#ff0000')
                if not error:
                    error = 'Start date cannot be after the due date.'
            else:
                self.start.config(bg='#ffffff')
                self.due.config(bg='#ffffff')
        except:
            if not error:
                error = ''

        if error is not None:
            self.error.config(fg='#ff0000', text=error)
            return False
        else:
            self.error.config(fg='#ffffff', text='')
            return name, email, item, start_date, due_date 

    def compose(self, event):
        values = self.test_values()
        if values:
            name, email, item, start_date, due_date = values
            subject = SUBJECT.format('Key' if item.lower() == 'key' else 'Equipment')
            mesg = compose_message(name, item, start_date, due_date)
            CompositionWindow(email, 'asdf', subject, mesg)


class CompositionWindow(tk.Toplevel):

    def __init__(self, toaddr, fromaddr, subject, mesg):
        tk.Toplevel.__init__(self)

        label = tk.Label(self, text='To:')
        label.grid(row=0, sticky=tk.W)

        self.tovar = tk.StringVar(self, value=toaddr)
        toaddr = tk.Entry(self, textvariable=self.tovar)
        toaddr.grid(row=0, column=1, columnspan=11, sticky=tk.W+tk.E)

        label = tk.Label(self, text='From:')
        label.grid(row=1, sticky=tk.W)

        self.fromvar = tk.StringVar(self, value=fromaddr)
        fromaddr = tk.Entry(self, textvariable=self.fromvar)
        fromaddr.grid(row=1, column=1, columnspan=11, sticky=tk.W+tk.E)

        label = tk.Label(self, text='CC:')
        label.grid(row=2, sticky=tk.W)

        self.ccvar = tk.StringVar(self, value=', '.join(CC))
        cc = tk.Entry(self, textvariable=self.ccvar)
        cc.grid(row=2, column=1, columnspan=11, sticky=tk.W+tk.E)

        label = tk.Label(self, text='Subject:')
        label.grid(row=3, sticky=tk.W)

        self.subjectvar = tk.StringVar(self, value=subject)
        subject = tk.Entry(self, textvariable=self.subjectvar)
        subject.grid(row=3, column=1, columnspan=11, sticky=tk.W+tk.E)

        self.text = tk.Text(self, wrap=tk.WORD)
        self.text.insert(tk.INSERT, mesg)
        self.text.grid(row=4, column=0, columnspan=12)

        button = tk.Button(self, text='Send')
        button.bind('<Button-1>', self.send_email)
        button.bind('<Return>', self.send_email)
        button.grid(row=5, column=0, columnspan=12, sticky=tk.W+tk.E)
        
    def send_email(self, event):
        global username, password
        mesg = MIMEText(self.text.get(1.0, tk.END))
        mesg['Subject'] = self.subjectvar.get()
        mesg['From'] = self.fromvar.get()
        mesg['To'] = self.tovar.get()
        s = smtplib.SMTP(HOST)
        s.starttls()
        s.login(username, password)
        s.sendmail(username, [to] + CC, mesg.as_string())
        s.quit()



if __name__ == '__main__':
    config = ConfigParser.RawConfigParser()
    with open('ramail.conf') as f:
        config.readfp(f)
    username = config.get('Authentication', 'email')
    password = config.get('Authentication', 'password')

    app = tk.Tk()
    app.title('give me money to buy a new bike!')
    MainWindow(app)
    app.mainloop()

