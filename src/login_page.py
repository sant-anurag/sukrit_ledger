import tkinter as tk
from tkinter import Menu, messagebox
from ledger_form import LedgerForm
from app_defines import *
from import_database import * 
from init_database import * 
from account_statement import *

class LoginPage(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack()
        self.create_widgets()
        self.obj_commonUtil = None  # Initialize as needed
        self.obj_initDatabase = InitDatabase()  # Initialize as needed

    def create_widgets(self):
        # Define your login page widgets (labels, entry fields, buttons, etc.)
        self.username_label = tk.Label(self, text="Username:")
        self.username_entry = tk.Entry(self)
        self.password_label = tk.Label(self, text="Password:")
        self.password_entry = tk.Entry(self, show="*")
        self.login_button = tk.Button(self, text="Login", command=self.login)

        # Arrange the widgets using grid, pack, or place geometry managers
        self.username_label.grid(row=0, column=0)
        self.username_entry.grid(row=0, column=1)
        self.password_label.grid(row=1, column=0)
        self.password_entry.grid(row=1, column=1)
        self.login_button.grid(row=2, column=0, columnspan=2)

    def login(self):
        # Implement the login functionality here
        username = self.username_entry.get()
        password = self.password_entry.get()

        # Validate the username and password, and perform login logic
        # Example:
        if username == "admin" and password == "admin":
            print("Login successful!")
            self.main_screen_design()  # Open the main screen design
        else:
            messagebox.showerror("Login Failed", "Invalid username or password")

    def main_screen_design(self):
        # prepares the menu bar
        self.main_bar = tk.Menu(self.master)

        self.file_menu = tk.Menu(self.main_bar, tearoff=0, font=('Helvetica', 10, 'normal'), fg="dark blue",
                                 activeforeground='light yellow')
        self.stock_transaction_menu = Menu(self.main_bar, tearoff=0, font=('Helvetica', 10, 'normal'), fg="dark blue",
                                           activeforeground='light yellow')
        self.edit_menu = Menu(self.main_bar, tearoff=0, font=('Helvetica', 10, 'normal'), fg="dark blue",
                              activeforeground='light yellow')
        self.option_menu = Menu(self.main_bar, tearoff=0, font=('Helvetica', 10, 'normal'), fg="dark blue",
                                activeforeground='light yellow')
        self.view_menu = Menu(self.main_bar, tearoff=0, font=('Helvetica', 10, 'normal'), fg="dark blue",
                              activeforeground='light yellow')
        self.account_menu = Menu(self.main_bar, tearoff=0, font=('Helvetica', 10, 'normal'), fg="dark blue",
                                 activeforeground='light yellow')

        self.main_bar.add_cascade(label="New", menu=self.file_menu, underline=0)
        self.main_bar.add_cascade(label="Stock Transaction", menu=self.stock_transaction_menu, underline=0)
        self.main_bar.add_cascade(label='Edit', menu=self.edit_menu, underline=0)
        self.main_bar.add_cascade(label='View', menu=self.view_menu, underline=0)
        self.main_bar.add_cascade(label='User', menu=self.account_menu, underline=0)

        # Add an item to the "New" menu to open the ledger form
        self.file_menu.add_command(label="Open Ledger Form", command=self.open_ledger_form, underline=0)

        self.account_menu.add_command(label='Import Database', command=self.import_database, state=NORMAL, underline=0)
        self.account_menu.add_command(label='Reset Database', command=self.resetDatabase, state=NORMAL, underline=0)
        self.account_menu.add_command(label='Re-initialize Database', state=NORMAL, underline=0)
        self.account_menu.add_command(label='Create Backup & Exit', command=self.exit_application, state=NORMAL, underline=0)
        self.account_menu.add_command(label='Exit', command=self.simpleExit, state=NORMAL, underline=0)
        
        self.master.config(menu=self.main_bar)

    

    def import_database(self):
        obj_importDatabase = ImportDatabase(self.master)
    
    def resetDatabase(self):
        self.obj_initDatabase.resetallExpanseData()

    def updateMonetaryDonationReceiptBooklet(self, invoice_id, trustType):
        print("updateInvoiceTable -->start")
        today = datetime.now()
        year = today.strftime("%Y")
        dirname = "..\\Expanse_Data\\" + year + "\\Seva_Rashi\\Receipts\\Template"
        if not os.path.exists(dirname):
            print("Current year directory is not available , hence building one")
            os.makedirs(dirname)

        if trustType == VIHANGAM_YOGA_KARNATAKA_TRUST:
            path = dirname + "\\Donation_Receipt_Booklet.xlsx"
        if trustType == SADGURU_SADAFAL_AADARSH_GAUSHALA_TRUST:
            path = dirname + "\\Gaushala_Donation_Receipt_Booklet.xlsx"

        wb_obj = openpyxl.load_workbook(path)
        wb_sheet = wb_obj.active
        total_records = self.totalrecords_excelDataBase(path)
        wb_sheet.cell(row=total_records + 2, column=1).value = str(total_records + 1)
        wb_sheet.cell(row=total_records + 2, column=2).value = str(invoice_id)
        wb_obj.save(path)
        print("Receipt booklet updated")

    def updateExpanseVoucherReceiptBooklet(self, invoice_id, trustType):
        print("updateExpanseVoucherReceiptBooklet -->start")
        today = datetime.now()
        year = today.strftime("%Y")
        dirname = "Expanse_Data\\" + year + "\\Expanse\\Receipts\\Template"
        if not os.path.exists(dirname):
            print("Current year directory is not available , hence building one")
            os.makedirs(dirname)

        if trustType == VIHANGAM_YOGA_KARNATAKA_TRUST:
            path = dirname + "\\Ashram_Expanse_Voucher_Receipt_Booklet.xlsx"
        if trustType == SADGURU_SADAFAL_AADARSH_GAUSHALA_TRUST:
            path = dirname + "\\Gaushala_Expanse_Voucher_Receipt_Booklet.xlsx"

        wb_obj = openpyxl.load_workbook(path)
        wb_sheet = wb_obj.active
        total_records = self.totalrecords_excelDataBase(path)
        wb_sheet.cell(row=total_records + 2, column=1).value = str(total_records + 1)
        wb_sheet.cell(row=total_records + 2, column=2).value = str(invoice_id)
        wb_obj.save(path)
        print("Expanse Voucher booklet updated")

    def generateDonationReceipt(self, donator_idText,
                                seva_amountText,
                                categoryText,
                                collector_nameText,
                                dateOfCollection_calc,
                                paymentMode_text,
                                invoice_id, member_data, print_invoice, submit_deposit):

        currentYearDirectory = self.obj_commonUtil.getCurrentYearFolderName()
        file_name = "Expanse_Data\\" + currentYearDirectory + "\\Seva_Rashi\\Receipts\\Template\\Donation_Receipt_Template.xlsx"
        wb_obj = openpyxl.load_workbook(file_name)
        sheet_obj = wb_obj.active

        sheet_obj.cell(row=6, column=4).value = str(invoice_id)
        sheet_obj.cell(row=7, column=4).value = str(dateOfCollection_calc)
        sheet_obj.cell(row=8, column=4).value = str(member_data[2])
        sheet_obj.cell(row=9, column=4).value = str(member_data[7])  # Address
        sheet_obj.cell(row=10, column=4).value = str(member_data[8])  # city
        sheet_obj.cell(row=11, column=4).value = str(member_data[9])  # state
        sheet_obj.cell(row=12, column=4).value = str(member_data[10])
        sheet_obj.cell(row=13, column=4).value = str(member_data[11])
        sheet_obj.cell(row=15, column=4).value = str(categoryText.get())

        sheet_obj.cell(row=16, column=4).value = str(seva_amountText.get())
        sheet_obj.cell(row=17, column=4).value = str(num2words.num2words(int(seva_amountText.get()))) + " Rs. only"
        sheet_obj.cell(row=18, column=4).value = str(paymentMode_text.get())
        sheet_obj.cell(row=19, column=4).value = str(collector_nameText)

        wb_obj.save(file_name)
        pdf_file = "..\\Expanse_Data\\" + currentYearDirectory + "\\Seva_Rashi\\Receipts\\Receipts\\" + invoice_id + ".pdf"

        self.obj_commonUtil.convertExcelToPdf(file_name, pdf_file)
        print_result = partial(self.printInvoice, pdf_file)
        print_invoice.configure(state=NORMAL, bg='light cyan', command=print_result)
        submit_deposit.configure(state=DISABLED, bg='light grey')

        destdir_repo = self.obj_initDatabase.get_invoice_directory_name() + "\\" + invoice_id + ".pdf"
        self.obj_commonUtil.updateInvoiceTable(invoice_id, destdir_repo)
        desktop_repo = self.obj_initDatabase.get_desktop_invoices_directory_path() + "\\" + invoice_id + ".pdf"
        if categoryText.get() == "Gaushala Seva":
            trustType = SADGURU_SADAFAL_AADARSH_GAUSHALA_TRUST
        else:
            trustType = VIHANGAM_YOGA_KARNATAKA_TRUST

        self.obj_commonUtil.updateMonetaryDonationReceiptBooklet(invoice_id, trustType)
        copyfile(pdf_file, destdir_repo)
        copyfile(pdf_file, desktop_repo)


    def deposit_seva_rashi_Excel(self, new_noncommercial_Item_window,
                                 donator_idText,
                                 seva_amountText,
                                 categoryText,
                                 collector_nameText,
                                 cal,
                                 paymentMode_menu,
                                 paymentMode_text,
                                 authorizedby_Text,
                                 akshayPatra_Text,
                                 invoice_idText,
                                 infolabel, print_invoice, transId_Text):
        dateTimeObj = cal.get_date()
        dateOfCollection_calc = dateTimeObj.strftime("%Y-%m-%d ")
        if donator_idText.get() == "" or \
                seva_amountText.get() == "" or \
                collector_nameText.get() == "":
            infolabel.configure(text="All fields are mandatory", fg='red')

        else:
            today = date.today()
            if dateTimeObj <= today:
                bDonatorIdValid = self.validate_memberlibraryID_Excel(donator_idText.get(), 1)
                bReceiverIdValid = self.validate_memberlibraryID_Excel(collector_nameText.get(), 1)
                bAuthorizorIdValid = self.validate_memberlibraryID_Excel(authorizedby_Text.get(), 1)
                if bDonatorIdValid and bReceiverIdValid and bAuthorizorIdValid and (seva_amountText.get()).isnumeric():
                    if categoryText.get() == "Akshay-Patra Seva" and akshayPatra_Text.get() == "Not Available":
                        infolabel.configure(text="No Akshay patra Assigned to this member !!!", fg='red')
                    else:
                        member_data = self.retrieve_MemberRecords_Excel(donator_idText.get(), 1, SEARCH_BY_MEMBERID)
                        revceiver_data = self.retrieve_MemberRecords_Excel(collector_nameText.get(), 1,
                                                                           SEARCH_BY_MEMBERID)
                        authorizor_data = self.retrieve_MemberRecords_Excel(authorizedby_Text.get(), 1,
                                                                            SEARCH_BY_MEMBERID)
                        print("For debugging Seva Amt :", seva_amountText.get(), "Max donation allowed :",
                              MAX_ALLOWED_DONATION)
                        bGnericSevaCase = True

                        # If the amount > 10000 and by Cash medium,same is considered to be a split candidate
                        if (paymentMode_text.get() == "Cash") and (int(seva_amountText.get()) > MAX_ALLOWED_DONATION):
                            bGnericSevaCase = False

                        # if not a split candidate, proceed with normal deposit to accounts
                        if bGnericSevaCase:

                            filename_MonetarySheet = self.obj_initDatabase.get_seva_deposit_database_name()  # PATH_SEVA_SHEET

                            # writting the credit in Master Seva Sheet - starts
                            if categoryText.get() == "Gaushala Seva":
                                invoice_id = self.obj_commonUtil.generateMonetaryDonationReceiptId(
                                    SADGURU_SADAFAL_AADARSH_GAUSHALA_TRUST)
                            else:
                                invoice_id = self.obj_commonUtil.generateMonetaryDonationReceiptId(
                                    VIHANGAM_YOGA_KARNATAKA_TRUST)
                            # open seva rashi sheet and enter the data --start
                            wb_obj = openpyxl.load_workbook(filename_MonetarySheet)
                            sheet_obj = wb_obj.active
                            total_records = self.totalrecords_excelDataBase(filename_MonetarySheet)

                            if total_records is 0:
                                serial_no = 1
                                row_no = 2
                                balance_amount = 0
                            else:
                                serial_no = total_records + 1
                                row_no = total_records + 2
                                balance_amount = int(str(sheet_obj.cell(row=row_no - 1, column=3).value))

                            # sets the general formatting for the new entry in new row
                            for index in range(1, 14):
                                sheet_obj.cell(row=row_no, column=index).font = Font(size=12, name='Times New Roman',
                                                                                     bold=False)
                                sheet_obj.cell(row=row_no, column=index).alignment = Alignment(horizontal='center',
                                                                                               vertical='center')

                            # new book data is assigned to respective cells in row
                            sheet_obj.cell(row=row_no, column=1).value = str(serial_no)
                            sheet_obj.cell(row=row_no, column=2).value = str(seva_amountText.get())
                            sheet_obj.cell(row=row_no, column=3).value = str(
                                balance_amount + int(seva_amountText.get()))
                            sheet_obj.cell(row=row_no, column=4).value = str(donator_idText.get())
                            sheet_obj.cell(row=row_no, column=5).value = str(member_data[2])
                            sheet_obj.cell(row=row_no, column=6).value = str(dateOfCollection_calc)
                            sheet_obj.cell(row=row_no, column=7).value = str(categoryText.get())
                            sheet_obj.cell(row=row_no, column=8).value = str(collector_nameText.get())
                            sheet_obj.cell(row=row_no, column=9).value = str(revceiver_data[2])
                            sheet_obj.cell(row=row_no, column=10).value = str(member_data[7])  # address
                            if paymentMode_text.get() == "Bank Transfer":
                                paymenttext = str(paymentMode_text.get()) + "Ref. " + str(transId_Text.get())
                            else:
                                paymenttext = str(paymentMode_text.get())
                            sheet_obj.cell(row=row_no, column=11).value = paymenttext
                            sheet_obj.cell(row=row_no, column=12).value = str(authorizedby_Text.get())
                            sheet_obj.cell(row=row_no, column=13).value = str(authorizor_data[2])
                            sheet_obj.cell(row=row_no, column=14).value = str(invoice_id)

                            wb_obj.save(filename_MonetarySheet)

                            # writting the credit in Master Seva Sheet - starts
                            bAlike_Seva = True
                            trust_type = VIHANGAM_YOGA_KARNATAKA_TRUST
                            if categoryText.get() == "Monthly Seva":
                                path_seva_sheet = self.obj_initDatabase.get_monthly_seva_database_name()  # PATH_MONTHLY_SEVA_SHEET
                            elif categoryText.get() == "Gaushala Seva":
                                trust_type = SADGURU_SADAFAL_AADARSH_GAUSHALA_TRUST
                                path_seva_sheet = self.obj_initDatabase.get_gaushala_seva_database_name()  # PATH_GAUSHALA_SEVA_SHEET
                            elif categoryText.get() == "Hawan Seva":
                                path_seva_sheet = self.obj_initDatabase.get_hawan_seva_database_name()  # PATH_HAWAN_SEVA_SHEET
                            elif categoryText.get() == "Event/Prachar Seva":
                                path_seva_sheet = self.obj_initDatabase.get_prachar_event_seva_database_name()  # PATH_EVENT_SEVA_SHEET
                            elif categoryText.get() == "Aarti Seva":
                                path_seva_sheet = self.obj_initDatabase.get_aarti_seva_database_name()  # PATH_AARTI_SEVA_SHEET
                            elif categoryText.get() == "Ashram Seva(Generic)":
                                path_seva_sheet = self.obj_initDatabase.get_ashram_seva_database_name()  # PATH_ASHRAM_GENERIC_SEVA_SHEET
                            elif categoryText.get() == "Ashram Nirmaan Seva":
                                path_seva_sheet = self.obj_initDatabase.get_ashram_nirmaan_seva_database_name()  # PATH_ASHRAM_NIRMAAN_SHEET
                            elif categoryText.get() == "Yoga Fees":
                                path_seva_sheet = self.obj_initDatabase.get_yoga_seva_database_name()  # PATH_YOGA_FEES_SHEET
                            elif categoryText.get() == "Akshay-Patra Seva":
                                path_seva_sheet = self.obj_initDatabase.get_akshay_patra_database_name()  # PATH_AKSHAY_PATRA_DATABASE
                            else:
                                bAlike_Seva = False
                                pass

                            # writting into respective seva sheet - start
                            if bAlike_Seva:
                                wb_sevasheetobj = openpyxl.load_workbook(path_seva_sheet)
                                sevasheet_obj = wb_sevasheetobj.active
                                total_records_seva_sheet = self.totalrecords_excelDataBase(path_seva_sheet)

                                if total_records_seva_sheet is 0:
                                    serial_no = 1
                                    row_no = 2
                                    balance_amount = 0
                                else:
                                    serial_no = total_records_seva_sheet + 1
                                    row_no = total_records_seva_sheet + 2
                                    balance_amount = int(str(sevasheet_obj.cell(row=row_no - 1, column=3).value))

                                # sets the general formatting for the new entry in new row
                                for index in range(1, 14):
                                    sevasheet_obj.cell(row=row_no, column=index).font = Font(size=12,
                                                                                             name='Times New Roman',
                                                                                             bold=False)
                                    sevasheet_obj.cell(row=row_no, column=index).alignment = Alignment(
                                        horizontal='center',
                                        vertical='center')

                                # new book data is assigned to respective cells in row
                                sevasheet_obj.cell(row=row_no, column=1).value = str(serial_no)
                                sevasheet_obj.cell(row=row_no, column=2).value = str(seva_amountText.get())
                                sevasheet_obj.cell(row=row_no, column=3).value = str(
                                    balance_amount + int(seva_amountText.get()))
                                sevasheet_obj.cell(row=row_no, column=4).value = str(donator_idText.get())
                                sevasheet_obj.cell(row=row_no, column=5).value = str(member_data[2])
                                sevasheet_obj.cell(row=row_no, column=6).value = str(dateOfCollection_calc)
                                sevasheet_obj.cell(row=row_no, column=7).value = str(categoryText.get())
                                sevasheet_obj.cell(row=row_no, column=8).value = str(collector_nameText.get())
                                sevasheet_obj.cell(row=row_no, column=9).value = str(revceiver_data[2])
                                if categoryText.get() == "Akshay-Patra Seva":
                                    sevasheet_obj.cell(row=row_no, column=10).value = str(
                                        akshayPatra_Text.get())  # address
                                else:
                                    sevasheet_obj.cell(row=row_no, column=10).value = str(member_data[7])  # address
                                if paymentMode_text.get() == "Bank Transfer":
                                    paymenttext = str(paymentMode_text.get()) + "Ref. " + str(transId_Text.get())
                                else:
                                    paymenttext = str(paymentMode_text.get())
                                sevasheet_obj.cell(row=row_no, column=11).value = paymenttext
                                sevasheet_obj.cell(row=row_no, column=12).value = str(authorizedby_Text.get())
                                sevasheet_obj.cell(row=row_no, column=13).value = str(authorizor_data[2])
                                sevasheet_obj.cell(row=row_no, column=14).value = str(invoice_id)

                                wb_sevasheetobj.save(path_seva_sheet)

                                # writting into respective seva sheet - end

                            # open transaction sheet and enter the data
                            # receiving donation is a credit transaction for the organization
                            if categoryText.get() == "Gaushala Seva":
                                file_name_transaction = self.obj_initDatabase.get_gaushala_transaction_database_name()  # PATH_TRANSACTION_SHEET
                            else:
                                file_name_transaction = self.obj_initDatabase.get_transaction_database_name()  # PATH_TRANSACTION_SHEET

                            transaction_wb_obj = openpyxl.load_workbook(file_name_transaction)
                            transaction_sheet_obj = transaction_wb_obj.active
                            total_records_transaction = self.totalrecords_excelDataBase(file_name_transaction)

                            if total_records_transaction is 0:
                                serial_no = 1
                                row_no = 2
                                balance_amount = 0
                            else:
                                serial_no = total_records_transaction + 1
                                row_no = total_records_transaction + 2
                                balance_amount = int(str(transaction_sheet_obj.cell(row=row_no - 1, column=9).value))

                            # sets the general formatting for the new entry in new row
                            for index in range(1, 10):
                                transaction_sheet_obj.cell(row=row_no, column=index).font = Font(size=12,
                                                                                                 name='Times New Roman',
                                                                                                 bold=False)
                                transaction_sheet_obj.cell(row=row_no, column=index).alignment = Alignment(
                                    horizontal='center',
                                    vertical='center')

                            # new book data is assigned to respective cells in row
                            transaction_sheet_obj.cell(row=row_no, column=1).value = str(serial_no)
                            transaction_sheet_obj.cell(row=row_no, column=2).value = str(dateOfCollection_calc)
                            transaction_sheet_obj.cell(row=row_no, column=3).value = str(seva_amountText.get())
                            transaction_sheet_obj.cell(row=row_no, column=4).value = "---"
                            transaction_sheet_obj.cell(row=row_no, column=5).value = str(categoryText.get())
                            if paymentMode_text.get() == "Bank Transfer":
                                paymenttext = str(paymentMode_text.get()) + "Ref. " + str(transId_Text.get())
                            else:
                                paymenttext = str(paymentMode_text.get())
                            transaction_sheet_obj.cell(row=row_no, column=6).value = paymenttext
                            transaction_sheet_obj.cell(row=row_no, column=7).value = str(
                                authorizedby_Text.get())  # authorizor id
                            transaction_sheet_obj.cell(row=row_no, column=8).value = str(authorizor_data[2])
                            transaction_sheet_obj.cell(row=row_no, column=9).value = str(
                                balance_amount + int(seva_amountText.get()))
                            transaction_sheet_obj.cell(row=row_no, column=10).value = str(invoice_id)
                            transaction_wb_obj.save(file_name_transaction)

                            invoice_idText['text'] = invoice_id

                            # open transaction sheet and enter the data --end

                            self.generateDonationReceipt(donator_idText,
                                                         seva_amountText,
                                                         categoryText,
                                                         revceiver_data[2],
                                                         dateOfCollection_calc,
                                                         paymentMode_text,
                                                         invoice_id, member_data, print_invoice, self.submit_deposit)

                            text_withID = "Seva deposited successfully. Invoice  id :" + invoice_id
                            infolabel.configure(text=text_withID, fg='green')
                            # update the total balance
                            self.obj_commonUtil.calculateTotalAvailableBalance(trust_type)
                        else:
                            print("Donation amount is by CASH and greater than ", MAX_ALLOWED_DONATION)
                            filename_splitCandidate = self.obj_initDatabase.get_splittransaction_database_name()

                            # writting the credit in Master Seva Sheet - starts
                            invoice_id = self.generate_SplitCandidate_ReceiptNo()
                            # open seva rashi sheet and enter the data --start
                            wb_obj = openpyxl.load_workbook(filename_splitCandidate)
                            sheet_obj = wb_obj.active
                            total_records = self.totalrecords_excelDataBase(filename_splitCandidate)

                            if total_records is 0:
                                serial_no = 1
                                row_no = 2
                                balance_amount = 0
                            else:
                                serial_no = total_records + 1
                                row_no = total_records + 2
                                balance_amount = int(str(sheet_obj.cell(row=row_no - 1, column=3).value))

                            # sets the general formatting for the new entry in new row
                            for index in range(1, 11):
                                sheet_obj.cell(row=row_no, column=index).font = Font(size=12, name='Times New Roman',
                                                                                     bold=False)
                                sheet_obj.cell(row=row_no, column=index).alignment = Alignment(horizontal='center',
                                                                                               vertical='center')

                            # new book data is assigned to respective cells in row
                            sheet_obj.cell(row=row_no, column=1).value = str(serial_no)
                            sheet_obj.cell(row=row_no, column=2).value = str(dateOfCollection_calc)
                            sheet_obj.cell(row=row_no, column=3).value = str(seva_amountText.get())
                            sheet_obj.cell(row=row_no, column=4).value = "Open"
                            sheet_obj.cell(row=row_no, column=5).value = str(categoryText.get())
                            sheet_obj.cell(row=row_no, column=6).value = str(paymentMode_text.get())
                            sheet_obj.cell(row=row_no, column=7).value = str(donator_idText.get())
                            sheet_obj.cell(row=row_no, column=8).value = str(member_data[2])
                            sheet_obj.cell(row=row_no, column=9).value = str(seva_amountText.get())
                            sheet_obj.cell(row=row_no, column=10).value = str(invoice_id)
                            wb_obj.save(filename_splitCandidate)

                            self.generateDonationReceipt(donator_idText,
                                                         seva_amountText,
                                                         categoryText,
                                                         revceiver_data[2],
                                                         dateOfCollection_calc,
                                                         paymentMode_text,
                                                         str(invoice_id), member_data, print_invoice,
                                                         self.submit_deposit)

                            text_withID = "Special Donation saved successfully. Receipt  id :" + str(invoice_id)
                            infolabel.configure(text=text_withID, fg='green')
                else:
                    infolabel.configure(text="Invalid Data for ID/IDs/Amount, please correct ...", fg='red')
            else:
                infolabel.configure(text="Transaction Date cannot be future!! please correct ...", fg='red')


    def exit_application(self):
        print("Creating Backup before exit ")
        today = datetime.now()
        backup_folder = today.strftime("%d_%b_%Y_%H%M%S")

        src_folder = "../Expanse_Data\\"
        dest_folder = "C:\\VYOAM\\VYOAM_Backup\\" + backup_folder
        self.obj_commonUtil.create_backup(src_folder, dest_folder)
        print("Creating Backup Completed !!! ")
        self.obj_commonUtil.encryptDatabase()
        self.master.destroy()

    def simpleExit(self):
        self.obj_commonUtil.encryptDatabase()
        self.master.destroy()

    def open_ledger_form(self):
    # Create an instance of LedgerForm
        ledger_form = LedgerForm(self.master)

    def view_main_account_statement(self):
        obj_accountstatement = AccountStatement(root)

    def donothing(self, event=None):
        print("Button is disabled")
        pass

    def clear_loginForm(self, userNameText, passwordText):
        userNameText.delete(0, END)
        userNameText.configure(fg='black')
        userNameText.focus_set()
        passwordText.delete(0, END)
        passwordText.configure(fg='black')

    def resetallExpanseData(self):
        print("VYOAM Deleting Expanse Database")
        today = datetime.now()
        year = today.strftime("%Y")
        dirname1 = "Expanse_Data\\" + year + "\\Expanse"
        dirname2 = "Expanse_Data\\" + year + "\\Expanse\\Receipts"
        dirname3 = "Expanse_Data\\" + year + "\\Expanse\\Receipts\\Receipts"
        dirname4 = "Expanse_Data\\" + year + "\\Expanse\\Receipts\\Template"

        dirname5 = "Expanse_Data\\" + year + "\\Invoices"

        

        dirname10 = "..\\Expanse_Data\\" + year + "\\Seva_Rashi"
        dirname11 = "..\\Expanse_Data\\" + year + "\\Seva_Rashi\\Donation"
        dirname12 = "..\\Expanse_Data\\" + year + "\\Seva_Rashi\\Donation\\Sorted_List"
        dirname13 = "..\\Expanse_Data\\" + year + "\\Seva_Rashi\\Receipts"
        dirname14 = "..\\Expanse_Data\\" + year + "\\Seva_Rashi\\Receipts\\Receipts"
        dirname15 = "..\\Expanse_Data\\" + year + "\\Seva_Rashi\\Receipts\\Template"
        dirname16 = "..\\Expanse_Data\\" + year + "\\Seva_Rashi\\Statements"
        dirname17 = "..\\Expanse_Data\\" + year + "\\Seva_Rashi\\Template"

        dirname18 = "Expanse_Data\\" + year + "\\Statements"

        dirname19 = "Expanse_Data\\" + year + "\\Transaction"
        dirname20 = "Expanse_Data\\" + year + "\\Transaction\\Account_Statement"
        dirname21 = "Expanse_Data\\" + year + "\\Transaction\\Account_Statement\\Statements"
        dirname22 = "Expanse_Data\\" + year + "\\Transaction\\Account_Statement\\Template"
        dirname23 = "Expanse_Data\\" + year + "\\Transaction\\Sorted_List"
        dirname24 = "Expanse_Data\\" + year + "\\Transaction\\StockSell"
        dirname25 = "Expanse_Data\\" + year + "\\Transaction\\StockSell\\Sorted"
        dirname26 = "Expanse_Data\\" + year + "\\Transaction\\StockSell\\Statements"
        dirname27 = "Expanse_Data\\" + year + "\\Transaction\\StockSell\\Template"

        directory_list = [dirname1, dirname2, dirname3, dirname4, dirname5, dirname10, dirname11, dirname12, dirname13, dirname14,
                          dirname15, dirname16, dirname17, dirname18, dirname19, dirname20, dirname21,
                          dirname22, dirname23, dirname24, dirname25, dirname26, dirname27]

        for index in range(0, len(directory_list)):
            if os.path.exists(directory_list[index]):
                print(directory_list[index], " ...deleted")
                shutil.rmtree(directory_list[index])

    def loading_animation_end(self):
        pass

    def main_menu(self):
        width, height = pyautogui.size()
        self.master.geometry('{}x{}+{}+{}'.format(width, height, 0, 0))
        # canvas designed to display the library image on main screen
        canvas_width, canvas_height = width, height
        canvas = Canvas(self.master, width=canvas_width, height=canvas_height)
        myimage = ImageTk.PhotoImage(PIL.Image.open("../Images/Logos/loading2.JPG").resize((width, height)))
        canvas.create_image(0, 0, anchor=NW, image=myimage)
        canvas.pack()

        self.master.lift()
        # prevents the application been closed by alt + F4 etc.
        # self.master.overrideredirect(True)
        self.create_widgets()
        self.master.mainloop()
# Usage example:
if __name__ == "__main__":
    root = tk.Tk()
    app = LoginPage(root)
    root.mainloop()
