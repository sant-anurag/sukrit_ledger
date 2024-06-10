import tkinter as tk
from tkinter import messagebox
from django.contrib.auth import authenticate

# Function to handle login button click
def login():
    username = username_entry.get()
    password = password_entry.get()

    # Authenticate user with Django
    user = authenticate(username=username, password=password)
    if user is not None:
        messagebox.showinfo("Login Successful", "Welcome, " + username)
        # Add code to open main application window or perform other actions
    else:
        messagebox.showerror("Login Failed", "Invalid username or password")

# Create main application window
root = tk.Tk()
root.title("Desktop Application")

# Create and pack username label and entry
username_label = tk.Label(root, text="Username:")
username_label.pack()
username_entry = tk.Entry(root)
username_entry.pack()

# Create and pack password label and entry
password_label = tk.Label(root, text="Password:")
password_label.pack()
password_entry = tk.Entry(root, show="*")
password_entry.pack()

# Create and pack login button
login_button = tk.Button(root, text="Login", command=login)
login_button.pack()

# Run the Tkinter event loop
root.mainloop()
