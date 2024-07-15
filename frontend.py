import customtkinter as tk
from main import getMail, generate_reply, mail_reply
import threading
import pythoncom

def confirm_button_clicked():
    reply = reply_textbox.get("1.0", tk.END)
    mail_reply(reply)
    # Add functionality to handle the reply here

def generate_button_clicked():
    pythoncom.CoInitialize()
    try:
        reply_textbox.delete("1.0", tk.END)
        reply_textbox.insert("1.0", generate_reply(getMail()))
    finally:
        pythoncom.CoUninitialize()

def start_generate_button_clicked():
    threading.Thread(target=generate_button_clicked).start()

def fetch_most_recent_email():
    try:
        email = getMail()
        subject = email.Subject
        body = email.Body
        return f"Subject: {subject}\n\nBody: {body}"
    except Exception as e:
        return f"Error fetching email: {e}"

# Create the main window
window = tk.CTk()
window.title("Email Reply")

# Create a frame for better layout control
frame = tk.CTkFrame(window)
frame.pack(pady=20, padx=20, fill="both", expand=True)

# Fetch the most recent email
email_content = fetch_most_recent_email()

# Create the email label and textbox
email_label = tk.CTkLabel(frame, text="Email:")
email_label.grid(row=0, column=0, padx=10, pady=10, sticky="e")
email_textbox = tk.CTkTextbox(frame, width=250, height=250)
email_textbox.grid(row=0, column=1, padx=10, pady=10, columnspan=2, sticky="ew")
email_textbox.insert("1.0", email_content)  # Insert the fetched email content
email_textbox.configure(state="disabled")

# Create the reply label and textbox
reply_label = tk.CTkLabel(frame, text="Reply:")
reply_label.grid(row=0, column=3, padx=10, pady=10, sticky="e")
reply_textbox = tk.CTkTextbox(frame, width=250, height=250)
reply_textbox.grid(row=0, column=4, padx=10, pady=10, columnspan=2, sticky="ew")

# Create a frame for the buttons
button_frame = tk.CTkFrame(frame)
button_frame.grid(row=1, column=0, columnspan=6, pady=10, padx=10)

# Create the confirm button
confirm_button = tk.CTkButton(button_frame, text="Confirm", command=confirm_button_clicked, hover_color="green")
confirm_button.grid(row=0, column=0, padx=20, pady=10)

# Create the generate button
generate_button = tk.CTkButton(button_frame, text="Generate", command=start_generate_button_clicked, hover_color="green")
generate_button.grid(row=0, column=1, padx=20, pady=10)

# Center the buttons
button_frame.grid_columnconfigure(0, weight=1)
button_frame.grid_columnconfigure(1, weight=1)

# Start the main event loop
window.mainloop()
