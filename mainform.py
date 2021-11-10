from os import close
import tkinter as tk
from tkinter import Entry, Frame, font
from tkinter import ttk
from tkinter import scrolledtext
from tkinter import filedialog
from tkinter.constants import COMMAND, DISABLED, NORMAL
from send_otp import start_session, close_session, send_to_single_mail, write_otp_to_csv
import send_otp as so

#------------------------------------------------------------
def btnRecipientOpenFile_click():
    filepath = filedialog.askopenfilename(filetypes=[('CSV file', '*.csv'), ('All files', '*.*')])
    
    # the user closes the dialog box or clicks Cancel
    if not filepath:
        return
    # clear old path
    entRecipient.delete(0, tk.END)    
    entRecipient.insert(0, filepath)

#------------------------------------------------------------
def btnDestinationSaveFile_click():
    filepath = filedialog.asksaveasfilename(defaultextension='csv',
                                            filetypes=[('CSV file', '*.csv'), ('All files', '*.*')])

    # the user closes the dialog box or clicks Cancel
    if not filepath:
        return

    # clear old path
    entOtpDestination.delete(0, tk.END)    
    entOtpDestination.insert(0, filepath)

#------------------------------------------------------------
# Event handler for button
def btnSend_click():   
    sender = entSender.get()
    password = entPassword.get()
    recipient_csv_file = entRecipient.get()
    otp_length = int(entOtpLength.get())

    mail_subject = entSubject.get()
    opening = scrolltxtOpening.get("1.0", tk.END)
    closing = scrolltxtClosing.get("1.0", tk.END)
    signature = scrolltxtSignature.get("1.0", tk.END)
    csv_destination = entOtpDestination.get()

    mail_otp_dict = so.create_mail_otp_dict(receive_mail_csv_file=recipient_csv_file, otp_len=otp_length)
    
    current_session = so.start_session(sender=sender, password=password)
    
    # enable txtMessage
    scrolltxtMessage.configure(state='normal')

    for r in mail_otp_dict:
        try:
            send_to_single_mail(session=current_session, sender=sender, mail_subject=mail_subject,
                                recipient=r, opening=opening, otp_code=mail_otp_dict[r],
                                closing=closing, signature=signature)
            message = 'OTP is succesfully sent to ' + r + '\n'
            scrolltxtMessage.insert(tk.END, message)
        except Exception as error:
            error_string = 'Mail: ' + r + ' with OTP ' + mail_otp_dict[r] + ' is failed to send: ' + str(error) + '\n'
            scrolltxtMessage.insert(tk.END, error_string)
    
    close_session(current_session)    
    
    write_otp_to_csv(otp_dict=mail_otp_dict, csv_dest_file=csv_destination)
    message = 'mail_otp_dict is written to csv file prosperously.\n' 
    scrolltxtMessage.insert(tk.END, message)
    scrolltxtMessage.configure(state=DISABLED)

#------------------------------------------------------------
# Event handler for Close button
def btnClose_click():
    window.destroy()

# Create a new window
window = tk.Tk()
window.title('Send OTP via email')

highlightFont = font.Font(family='Roboto', name = 'appHighlightFont', size=11)
consolas = font.Font(family='Consolas', name='widgetFont', size=11)

# Create a new frame
frmForm = tk.Frame(relief=tk.FLAT, borderwidth=3)
frmForm.pack(ipadx=5, ipady=5) # pack the frame into the window

# List of labels
labels = ['Sender:', 'Password:', 'Recipient list file:', 
            'Subject:', 'OTP length:',
            'Opening lines (in HTML):', 'Closing lines (in HTML):', 'Signature (in HTML):',
            'Destination csv file:']

# Enumerate the list
for i, text in enumerate(labels):
    # Create label widgets
    label = tk.Label(master=frmForm, text=text, font=highlightFont)

    # Use the grid geometry manager to place the labels, entries and texts
    label.grid(row=i, column=0, sticky='ne', pady=5, ipady=3)

# Sender email
entSender = tk.Entry(master=frmForm, width=50, font=consolas)
entSender.grid(row=0, column=1, sticky='w', pady=5, ipady=3)
entSender.insert(0, so.SENDER)

# Password of sender email
entPassword = tk.Entry(master=frmForm, textvariable=so.PWD, show='*', width=50, font=consolas)
entPassword.grid(row=1, column=1, sticky='w', pady=5, ipady=3)
entPassword.insert(0, so.PWD)

# The recipient list file path
# Create a frame for text and button
frmRecipient = tk.Frame(master=frmForm)
frmRecipient.grid(row=2, column=1, sticky='w')

entRecipient = tk.Entry(master=frmRecipient, width=50, font=consolas)
entRecipient.pack(side=tk.LEFT, pady=5, ipady=3)

btnRecipientOpenFile = tk.Button(master=frmRecipient, text='...', command=btnRecipientOpenFile_click)
btnRecipientOpenFile.pack(padx=5, pady=5)

# The subject of email
entSubject = tk.Entry(master=frmForm, width=50, font=consolas)
entSubject.grid(row=3, column=1, sticky='w', pady=5, ipady=3)
entSubject.insert(0, so.MAIL_SUBJECT)

# OTP length
entOtpLength = tk.Entry(master=frmForm, width=50, font=consolas)
entOtpLength.grid(row=4, column=1, sticky='w', pady=5, ipady=3)
entOtpLength.insert(0, '6')

# The opening paragraph of email
scrolltxtOpening = scrolledtext.ScrolledText(master=frmForm, wrap=tk.WORD, height=5, width=80, font=consolas) # size in characters
scrolltxtOpening.grid(row=5, column=1, sticky='w', pady=5, ipady=3)
scrolltxtOpening.insert('1.0', so.opening_html)

# The closing paragraph of email
scrolltxtClosing = scrolledtext.ScrolledText(master=frmForm, height=5, width=80, font=consolas)
scrolltxtClosing.grid(row=6, column=1, sticky='w', pady=5, ipady=3)
scrolltxtClosing.insert('1.0', so.closing_html)

# The signature of email
scrolltxtSignature = scrolledtext.ScrolledText(master=frmForm, wrap=tk.WORD, height=5, width=80, font=consolas)
scrolltxtSignature.grid(row=7, column=1, sticky='w', pady=5, ipady=3)
scrolltxtSignature.insert('1.0', so.signature_html)

# The destination mail-otp csv file path
# Frame for text and button
frmOtp = tk.Frame(master=frmForm)
frmOtp.grid(row=8, column=1, sticky='w')

entOtpDestination = tk.Entry(master=frmOtp, width=50, font=consolas)
entOtpDestination.pack(side=tk.LEFT, pady=5, ipady=3)

btnDestinationSaveFile = tk.Button(master=frmOtp, text='...', command=btnDestinationSaveFile_click)
btnDestinationSaveFile.pack(padx=5, pady=5)

# Create a new frame to contain the buttons
# The frame fills the whole window in the horizontal direction
frmButtons = tk.Frame()
frmButtons.pack(fill=tk.X, ipadx=5, ipady=5)

# Create Send button
btnSend = tk.Button(master=frmButtons, text='Send', font=highlightFont, command=btnSend_click)
btnSend.pack(side=tk.RIGHT, padx=10, ipadx=20) #padx: padding around the button, ipadx: padding around the text

# Create a frame of output message
frmMessage = tk.Frame()
frmMessage.pack(fill=tk.X, ipadx=5, ipady=5)

scrolltxtMessage = scrolledtext.ScrolledText(master=frmMessage, height=10)
scrolltxtMessage.pack(fill=tk.X, padx=5, pady=5)

# Create a frame to contain Close button
frmClose = tk.Frame()
frmClose.pack(fill=tk.X, ipadx=5, ipady=5)

# button Close
btnClose = tk.Button(master=frmClose, text='Close', font=highlightFont, command=btnClose_click)
btnClose.pack(side=tk.RIGHT, padx=10, ipadx=20)

window.mainloop()