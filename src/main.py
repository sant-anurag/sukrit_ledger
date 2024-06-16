import ctypes

import pyautogui

from app_defines import *
from ledger_form import *


class ledger_main:
    def __init__(self, master):
        self.master = master

        self.currentUser = ""
        self.main_menu()

    def createledgerentry(self, master):
        LedgerForm(self.master)

    def designmainscreen(self, master, user_category):
        labelFrame = Label(master, text="Inventory & Sales Management", justify=CENTER,
                           font=XXL_FONT,
                           fg='black')
        # labelFrame.place(x=200, y=10)
        result_btnInv = partial(self.createledgerentry, master)
        btn_inventory = Button(master, text="Inventory", fg="Black", command=result_btnInv,
                               font=XL_FONT, width=20, state=NORMAL, bg='RosyBrown1')
        # labelFrame.place(x=200, y=10)
        # result_btnShopper = partial(self.new_shopper_view)
        btn_shopper = Button(master, text="Customer", fg="Black", command=NONE,
                             font=XL_FONT, width=20, state=NORMAL, bg='RosyBrown1')

        # result_sales = partial(self.sales_operations)
        btn_sales = Button(master, text="Sales", fg="Black", command=NONE,
                           font=XL_FONT, width=20, state=NORMAL, bg='RosyBrown1')

        if user_category == 'Admin':
            # result_btnMchd = partial(self.new_center_registration)
            btn_merchandise = Button(master, text="Merchandise", fg="Black", command=NONE,
                                     font=XL_FONT, width=20, state=NORMAL, bg='RosyBrown1')
            # result_btnReport = partial(self.inventory_report)
            btn_reports = Button(master, text="Inv. Reports", fg="Black", command=NONE,
                                 font=XL_FONT, width=20, state=NORMAL, bg='RosyBrown1')

            # result_btnSalesReport = partial(self.sales_report)
            btn_salesreports = Button(master, text="Sales Reports", fg="Black", command=NONE,
                                      font=XL_FONT, width=20, state=NORMAL, bg='RosyBrown1')

            # result_usrLogin = partial(self.user_login_screen)
            btn_usrCtrl = Button(master, text="User Control", fg="Black", command=NONE,
                                 font=XL_FONT, width=20, state=NORMAL, bg='RosyBrown1')
        btn_exit = Button(master, text="Exit", fg="Black", command=master.destroy,
                          font=XL_FONT, width=20, state=NORMAL, bg='RosyBrown1')

        btn_inventory.place(x=65, y=220)
        btn_shopper.place(x=65, y=275)
        btn_sales.place(x=65, y=330)
        if user_category == 'Admin':
            btn_merchandise.place(x=65, y=385)
            btn_reports.place(x=65, y=440)
            btn_salesreports.place(x=65, y=495)
            btn_usrCtrl.place(x=65, y=550)
            btn_exit.place(x=65, y=605)
            master.bind('<M>', lambda event=None: btn_merchandise.invoke())
            master.bind('<m>', lambda event=None: btn_merchandise.invoke())
            master.bind('<R>', lambda event=None: btn_reports.invoke())
            master.bind('<r>', lambda event=None: btn_reports.invoke())
            master.bind('<u>', lambda event=None: btn_usrCtrl.invoke())
            master.bind('<U>', lambda event=None: btn_usrCtrl.invoke())
        elif user_category == 'User':
            btn_exit.place(x=65, y=385)
        else:
            print("Un-reachable code")

        master.bind('<Escape>', lambda event=None: btn_exit.invoke())

        master.bind('<I>', lambda event=None: btn_inventory.invoke())
        master.bind('<i>', lambda event=None: btn_inventory.invoke())
        master.bind('<S>', lambda event=None: btn_sales.invoke())
        master.bind('<s>', lambda event=None: btn_sales.invoke())
        master.bind('<c>', lambda event=None: btn_shopper.invoke())
        master.bind('<C>', lambda event=None: btn_shopper.invoke())

        mainloop()

    def main_menu(self):
        width, height = pyautogui.size()
        self.master.geometry(
            '{}x{}+{}+{}'.format(int(width / 1.35), int(height / 1.25), int(width / 9), int(height / 12)))
        self.master.configure(bg='AntiqueWhite1')
        # canvas designed to display the library image on main screen

        canvas_width, canvas_height = width, height
        canvas = Canvas(self.master, width=canvas_width, height=canvas_height)
        myimage = ImageTk.PhotoImage(
            PIL.Image.open("..\\Images\\Logos\\Geometry-Header-1920x1080.jpg").resize((width * 2, height * 2)))
        canvas.create_image(0, 0, anchor=NW, image=myimage)
        canvas.pack()

        self.master.lift()
        # prevents the application been closed by alt + F4 etc.
        self.login_window()
        # self.designMainScreen(self.master, canvas)
        self.master.mainloop()

    def login_window(self):
        login_window = Toplevel(self.master, takefocus=True)  # create a GUI window
        # login_window.tk.call('tk', 'scaling', 2.0)
        # Get the master screen width and height , and place the child screen accordingly
        xSize = self.master.winfo_screenwidth()
        ySize = self.master.winfo_screenheight()

        # set the configuration of GUI window
        login_window.geometry(
            '{}x{}+{}+{}'.format(410, 200, (int(xSize / 2.7)), (int(ySize / 3.8) + 50)))
        login_window.title("Account Login")  # set the title of GUI window
        login_window.configure(bg="white")
        login_window.protocol('WM_DELETE_WINDOW')
        login_window.configure(background='wheat')
        upperFrame = Frame(login_window, width=300, height=200, bd=8, relief='ridge', bg="white")
        upperFrame.grid(row=1, column=0, padx=20, pady=5, columnspan=2)

        labelLogin = Label(upperFrame, text="System Authentication", width=30, anchor=CENTER, justify=CENTER,
                           font=('arial narrow', 18, 'normal'), fg='blue', bg='light cyan')
        labelLogin.grid(row=0, column=0, padx=1, pady=1)

        lowerFrame = Frame(login_window, width=300, height=110, bd=8, relief='ridge', bg="white")
        lowerFrame.grid(row=2, column=0, padx=20, pady=5)
        userNameLabel = Label(lowerFrame, text="User Name", width=12, anchor=W, justify=LEFT,
                              font=('arial narrow', 15, 'normal'), bg="white", bd=2, relief='ridge')
        userNameLabel.grid(row=2, column=0)
        userNameText = Entry(lowerFrame, width=22, font=('Yu Gothic', 12, 'normal'), bd=2, relief='ridge',
                             bg='light yellow')
        userNameText.grid(row=2, column=1, padx=5)
        userNameText.focus_set()

        passwordLabel = Label(lowerFrame, text="Password", width=12, anchor=W, justify=LEFT,
                              font=('arial narrow', 15, 'normal'), bg="white", bd=2, relief='ridge')
        passwordLabel.grid(row=3, column=0, pady=2)
        passwordText = Entry(lowerFrame, width=22, show='*', font=('Yu Gothic', 12, 'normal'), bd=2, relief='ridge',
                             bg='light yellow')
        passwordText.grid(row=3, column=1, padx=5, pady=2)

        # ---------------------------------Button Frame Start----------------------------------------
        buttonFrame = Frame(login_window, width=200, height=100, bd=4, relief='ridge')
        buttonFrame.grid(row=4, column=0)

        # create a Login Button and place into the button frame window
        login_result = partial(self.validateStaffLogin, userNameText, passwordText, labelLogin, login_window)
        submit = Button(buttonFrame, text="Login", fg="Black", command=login_result,
                        font=NORM_FONT, width=8, bg='light cyan', highlightcolor="snow")
        submit.grid(row=0, column=0)

        # create a Clear Button and place into the self.newItem_window window
        clear_result = partial(self.clear_loginForm, userNameText, passwordText)
        clear = Button(buttonFrame, text="Reset", fg="Black", command=clear_result,
                       font=NORM_FONT, width=8, bg='light cyan', underline=0, highlightcolor="black")
        clear.grid(row=0, column=1)

        # create a Cancel Button and place into the self.newItem_window window
        # cancel_Result = partial(self.destroyWindow, self.newItem_window)
        # close_result = partial(self.closeFromLogin, self.master)
        cancel = Button(buttonFrame, text="Close", fg="Black", command=self.master.destroy,
                        font=NORM_FONT, width=8, bg='light cyan', underline=0, highlightcolor="black")
        cancel.grid(row=0, column=2)
        # ---------------------------------Button Frame End----------------------------------------

        login_window.bind('<Return>', lambda event=None: submit.invoke())
        login_window.bind('<Alt-c>', lambda event=None: cancel.invoke())
        login_window.bind('<Alt-r>', lambda event=None: clear.invoke())
        login_window.focus_set()
        login_window.grab_set()

        login_window.mainloop()  # start the GUI

    def validateStaffLogin(self, userNameText, passwordText, labelLogin, login_window):
        bLoginValid = False

        curUser_category = "Admin"

        if userNameText.get() == curUser_category and passwordText.get() == "password@123":
            bLoginValid = True

        if bLoginValid:
            # self.obj_curUser.setCurrentUser(curUser_category)
            # self.obj_commonUtil.logActivity("Login Success")
            print("Authentication Success for user :", userNameText.get())
            login_window.destroy()
            print("Username = ", self.currentUser, "Category = ", curUser_category)
            self.designmainscreen(self.master, curUser_category)
        else:
            # log the user action
            # self.obj_commonUtil.logActivity("Login failure")
            labelLogin.configure(fg='red')
            labelLogin['text'] = "Login Failed !! Try Again"
            self.clear_loginForm(userNameText, passwordText)
            passwordText.focus()

    def clear_loginForm(self, userNameText, passwordText):
        userNameText.delete(0, END)
        userNameText.configure(fg='black')
        passwordText.delete(0, END)
        passwordText.configure(fg='black')
        userNameText.focus_set()


# obj_animation = LoadingAnimation()
root = Tk()

# Query DPI Awareness (Windows 10 and 8)
awareness = ctypes.c_int()
errorCode = ctypes.windll.shcore.GetProcessDpiAwareness(0, ctypes.byref(awareness))
print(awareness.value)

# Set DPI Awareness  (Windows 10 and 8)
# errorCode = ctypes.windll.shcore.SetProcessDpiAwareness(2)
# the argument is the awareness level, which can be 0, 1 or 2:
# for 1-to-1 pixel control I seem to need it to be non-zero (I'm using level 2)
dpi = root.winfo_fpixels('1i')
factor = dpi / 72
root.call('tk', 'scaling', factor)
libraryObj = ledger_main(root)
