import tkinter as tk
from tkinter import ttk
from tkinter import Text, filedialog
from tkcalendar import DateEntry
import win32com.client as win32
import babel
from datetime import datetime


def schedule_meeting():
    ticket_type = ticket_type_var.get()
    order_number = order_number_entry.get()
    customer_name = customer_name_entry.get()
    start_datetime_str = f"{start_date_entry.get()} {start_hour_spin.get()}:{start_minute_spin.get()}"
    end_datetime_str = f"{end_date_entry.get()} {end_hour_spin.get()}:{end_minute_spin.get()}"
    meeting_content = content_text.get("1.0", tk.END).strip()

    selected_recipients = [recipient_listbox.get(i) for i in recipient_listbox.curselection()]

    try:
        start_time_dt = datetime.strptime(start_datetime_str, '%d-%m-%Y %H:%M')
        end_time_dt = datetime.strptime(end_datetime_str, '%d-%m-%Y %H:%M')

        if end_time_dt <= start_time_dt:
            result_label.config(text="End time must be after start time.")
            return

        start_time_outlook = start_time_dt.strftime('%Y-%m-%d %H:%M:%S')
        end_time_outlook = end_time_dt.strftime('%Y-%m-%d %H:%M:%S')
    except ValueError:
        result_label.config(text="Invalid date or time format. Ensure correct date and time selection.")
        return

    if not selected_recipients:
        result_label.config(text="Please select at least one recipient.")
        return

    if not meeting_content:
        result_label.config(text="Please provide content for the meeting.")
        return

    outlook = win32.Dispatch('outlook.application')
    meeting = outlook.CreateItem(1)  # 1 represents appointment item

    meeting.Subject = f"{ticket_type} - Order Number: {order_number} - Customer Name: {customer_name}"
    meeting.Start = start_time_outlook
    meeting.End = end_time_outlook
    meeting.Location = 'Conference Room'
    meeting.Body = f'Ticket Type: {ticket_type}\nOrder Number: {order_number}\nCustomer Name: {customer_name}\n\nContent:\n{meeting_content}'

    for recipient in selected_recipients:
        meeting.Recipients.Add(recipient)

    # Attach files
    for file_path in file_paths:
        meeting.Attachments.Add(file_path)

    meeting.Save()
    result_label.config(
        text=f"Meeting scheduled with {', '.join(selected_recipients)} from {start_datetime_str} to {end_datetime_str}")


def browse_files():
    global file_paths
    file_paths = filedialog.askopenfilenames(title="Select Files to Attach", filetypes=[("All Files", "*.*")])
    if file_paths:
        file_list_label.config(text="Files selected: " + ", ".join(file_paths))
    else:
        file_list_label.config(text="No files selected")


def create_label_entry(root, row, label_text, entry_var=None):
    ttk.Label(root, text=label_text).grid(row=row, column=0, padx=10, pady=10)
    entry = ttk.Entry(root, textvariable=entry_var) if entry_var else ttk.Entry(root)
    entry.grid(row=row, column=1, padx=10, pady=10)
    return entry


def create_label_dateentry(root, row, label_text):
    ttk.Label(root, text=label_text).grid(row=row, column=0, padx=10, pady=10)
    date_entry = DateEntry(root, date_pattern='dd-mm-yyyy')
    date_entry.grid(row=row, column=1, padx=10, pady=10)
    return date_entry


def create_label_spinbox(root, row, label_text, column=1):
    ttk.Label(root, text=label_text).grid(row=row, column=0, padx=10, pady=10)
    hour_spin = ttk.Spinbox(root, from_=0, to=23, width=5, format='%02.0f')
    minute_spin = ttk.Spinbox(root, from_=0, to=59, width=5, format='%02.0f')
    hour_spin.grid(row=row, column=column, sticky='w')
    minute_spin.grid(row=row, column=column, padx=(0, 0))
    return hour_spin, minute_spin


# Create the main window
root = tk.Tk()
root.title("Meeting Scheduler")

# Initialize global variable for file paths
file_paths = []

# Ticket type
ticket_type_var = tk.StringVar()
ticket_type_combo = ttk.Combobox(root, textvariable=ticket_type_var,
                                 values=["Migration", "Smart Hand", "OAR", "Trouble Ticket", "Others"])
ticket_type_label = ttk.Label(root, text="Select Ticket Type:")
ticket_type_label.grid(row=0, column=0, padx=10, pady=10)
ticket_type_combo.grid(row=0, column=1, padx=10, pady=10)

# Order number and customer name
order_number_entry = create_label_entry(root, 1, "Order Number:")
customer_name_entry = create_label_entry(root, 2, "Customer Name:")

# Start date and time
start_date_entry = create_label_dateentry(root, 3, "Start Date:")
start_hour_spin, start_minute_spin = create_label_spinbox(root, 4, "Start Time:")

# End date and time
end_date_entry = create_label_dateentry(root, 5, "End Date:")
end_hour_spin, end_minute_spin = create_label_spinbox(root, 6, "End Time:")

# Recipients
recipient_label = ttk.Label(root, text="Select Recipients:")
recipient_label.grid(row=7, column=0, padx=10, pady=10)
recipients = [

    "bsouth@ap.equinix.com", "bvu@ap.equinix.com", "snguyen@ap.equinix.com", "cjheng@ap.equinix.com",
    "tmihalatos@ap.equinix.com", "ttamnna@ap.equinix.com", "hsetiawan@ap.equinix.com",
    "alekkat@ap.equinix.com", "ddang@ap.equinix.com", "dpriyanath@ap.equinix.com",
    "fhuynh@ap.equinix.com", "atamang@ap.equinix.com", "ctruong@ap.equinix.com", "svy@ap.equinix.com",
    "ibx-syd-loc@ap.equinix.com", "ibx-sy3-loc@ap.equinix.com", "ibx-sy4-loc@ap.equinix.com",
    "ibx-sy5-loc@ap.equinix.com", "sy1-2securityteam@ap.equinix.com", "ibx-sy1-2-fac@ap.equinix.com"
]
recipient_listbox = tk.Listbox(root, selectmode=tk.MULTIPLE, height=6)
for recipient in recipients:
    recipient_listbox.insert(tk.END, recipient)
recipient_listbox.grid(row=7, column=1, padx=10, pady=10)

# Meeting content
content_label = ttk.Label(root, text="Meeting Content:")
content_label.grid(row=8, column=0, padx=10, pady=10)
content_text = Text(root, height=6, width=40)
content_text.grid(row=8, column=1, padx=10, pady=10)

# File attachment
file_attach_button = ttk.Button(root, text="Browse Files", command=browse_files)
file_attach_button.grid(row=9, column=0, padx=10, pady=10)
file_list_label = ttk.Label(root, text="No files selected")
file_list_label.grid(row=9, column=1, padx=10, pady=10)

# Submit button and result label
submit_button = ttk.Button(root, text="Schedule Meeting", command=schedule_meeting)
submit_button.grid(row=10, column=1, padx=10, pady=10)

result_label = ttk.Label(root, text="")
result_label.grid(row=11, column=0, columnspan=2, padx=10, pady=10)

root.mainloop()
