import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkinter import PhotoImage

# Database Initialization
def init_db():
    conn = sqlite3.connect("C:/Users/rajpu/OneDrive - EcoSoul Home/Central Repository/Application Programs/Expense Management/Mydatabase/expenseDB.db")
    c = conn.cursor()

    # Create tables
    c.execute('''CREATE TABLE IF NOT EXISTS Users (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    username TEXT, password TEXT,
                    department TEXT, role TEXT)''')

    c.execute('''CREATE TABLE IF NOT EXISTS Expenses (id INTEGER PRIMARY KEY AUTOINCREMENT, email TEXT,
              name TEXT, created_date TEXT, department TEXT, nature TEXT, type TEXT, country TEXT, currency TEXT,
              invoice_date TEXT, invoice_number TEXT, party_name TEXT, amount REAL, due_date TEXT, 
              payment_settlement TEXT, description TEXT, po_owner TEXT, po_date TEXT, po_amount REAL, po_details TEXT,
              balance_due REAL, attached_files TEXT, status TEXT)''')
    conn.commit()
    conn.close()

init_db()

class ExpenseTrackerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Ecosoul EMS💸")
        self.root.geometry("1000x800")
        self.conn = sqlite3.connect('C:/Users/rajpu/OneDrive - EcoSoul Home/Central Repository/Application Programs/Expense Management/Mydatabase/expenseDB.db')
        self.c = self.conn.cursor()
        self.full_access = False  # Initialize full_access flag

        # Style Configuration
        self.style = ttk.Style()
        self.style.configure('TButton', font=('Arial', 12), padding=5)
        self.style.configure('TLabel', font=('Arial', 14, 'bold'))
        self.style.configure('TFrame', background="#E8E8E8")
        self.style.configure('TEntry', padding=5)
        self.style.configure('TCombobox', padding=5)
        self.root.configure(bg="#dee2e6")
        
        self.create_login_page()

    def create_login_page(self):
        self.clear_frame()
 
        # Create a container frame to hold all widgets and center them
        container_frame = tk.Frame(self.root, bg="#dee2e6")
        container_frame.pack(expand=False)

        # Add top margin by placing an empty frame
        tk.Frame(container_frame, height=50, bg="#dee2e6").pack()    

        tk.Label(container_frame, text="Ecosoul Expense Management Solution 💱", font=('Arial', 18, 'bold'), bg="#dee2e6", fg="#05644d").pack(fill=tk.X, pady=20)

        # Labels and Entries
        tk.Label(self.root, text="Username", font=('Arial', 16), bg="#dee2e6",fg="black").pack(pady=5)
        self.username_entry = tk.Entry(self.root, font=('Arial', 14))
        self.username_entry.pack(pady=5)

        tk.Label(self.root, text="Password", font=('Arial', 16), bg="#dee2e6",fg="black").pack(pady=5)
        self.password_entry = tk.Entry(self.root, show="*", font=('Arial', 14))
        self.password_entry.pack(pady=5)

        tk.Label(self.root, text="Department", font=('Arial', 16), bg="#dee2e6",fg="black").pack(pady=5)
        self.department_combobox = ttk.Combobox(self.root, values=[
            "Business | Data Analyst",
            "Digital Marketing",
            "USA - Retail | Ecomm",
            "India - Retail | Ecomm",
            "Logistics | Supply Chain",
            "Administration | HR | PR",
            "Other"], font=('Arial', 16))
        self.department_combobox.pack(pady=5)

        tk.Button(self.root, text="Login", font=('Arial', 14, 'bold'), bg="#fca311", fg="#ffffff", command=self.login).pack(pady=20)

    def login(self):
        username = self.username_entry.get()
        password = self.password_entry.get()
        department = self.department_combobox.get()
        
        self.c.execute("SELECT * FROM Users WHERE username=? AND password=? AND department=?", 
                        (username, password, department))
        
        result = self.c.fetchone()
        if result:
            role = result[4]  # Assuming role is in the 5th column
            self.full_access = (role == "admin")  # Store full_access based on role
            self.department = department 
            self.open_dashboard()
        else:
            messagebox.showerror("Login Failed", "Invalid credentials!")

    def clear_frame(self):
        for widget in self.root.winfo_children():
            widget.destroy()

    def open_dashboard(self):
        self.clear_frame()

        # Sidebar Frame
        sidebar_frame = tk.Frame(self.root)
        sidebar_frame.pack(side=tk.LEFT, fill=tk.Y)

        # Sidebar buttons
        sidebar_frame = tk.Frame(self.root, width=200, bg="#00624E")
        sidebar_frame.pack(side=tk.LEFT, fill=tk.Y)

        # Add "Menu" label at the top of the sidebar, centered
        tk.Label(sidebar_frame, text="Menu", font=('Arial', 18, 'bold'), bg="#f0f0f0", fg="#05644d").pack(fill=tk.X, pady=20)

        tk.Button(sidebar_frame, text="📝 Expense Form", font=('Arial', 14), bg="#edf6f9",compound=tk.LEFT, anchor='w', padx=10, pady=8, command=self.create_expense_form).pack(fill=tk.X, padx=8, pady=8)
        tk.Button(sidebar_frame, text="📚 Track Expenses", font=('Arial', 14), bg="#edf6f9",compound=tk.LEFT, anchor='w', padx=10, pady=8, command=self.track_expenses).pack(fill=tk.X, padx=8, pady=8)
        
        if self.full_access:
            tk.Button(sidebar_frame, text="📒 Approve Expenses", font=('Arial', 14), bg="#edf6f9",compound=tk.LEFT, anchor='w', padx=10, pady=8, command=self.approve_expenses).pack(fill=tk.X, padx=8, pady=8)
        #tk.Button(sidebar_frame, text="Approve Expenses", command=self.track_expenses).pack(fill=tk.X, padx=5, pady=5)

        # Add Logout button to the sidebar
        tk.Button(sidebar_frame, text="Logout", font=('Arial', 14), bg="#fca311",fg="#ffffff", command=self.logout).pack(fill=tk.X, padx=50, pady=50)
        
    def create_expense_form(self):
        self.clear_frame()
        self.open_dashboard()

        # Content Frame
        self.content_frame = tk.Frame(self.root)
        self.content_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        self.content_frame = self.content_frame

        # Centering form entries
        form_frame = tk.Frame(self.content_frame)
        form_frame.place(relx=0.5, rely=0.1, anchor=tk.N)

        # Page Title
        tk.Label(form_frame, text="⟪ New Expense Entry Form ⟫", font=('circular', 20, 'bold'), bg="#f0f0f0", fg="#000000").grid(row=0, column=0, columnspan=3, pady=20)

        # Define font style
        font_style = ('Arial', 16, 'bold')  # Change size and style as needed

        # Labels and Entries
        tk.Label(form_frame, text="Email", font=font_style).grid(row=1, column=0, sticky=tk.W, pady=2)
        self.email_entry = tk.Entry(form_frame,font=('Arial', 12, 'normal'))
        self.email_entry.grid(row=1, column=1, pady=2)

        tk.Label(form_frame, text="Name", font=font_style).grid(row=2, column=0, sticky=tk.W, pady=2)
        self.name_entry = tk.Entry(form_frame,font=('Arial', 12, 'normal'))
        self.name_entry.grid(row=2, column=1, pady=2)

        tk.Label(form_frame, text="Expense Created Date", font=font_style).grid(row=3, column=0, sticky=tk.W, pady=2)
        self.created_date_entry = tk.Entry(form_frame,font=('Arial', 12, 'normal'))
        self.created_date_entry.grid(row=3, column=1, pady=2)

        # Department Dropdown
        tk.Label(form_frame, text="Select your Department", font=font_style).grid(row=4, column=0, sticky=tk.W, pady=2)
        self.department_entry = ttk.Combobox(form_frame,font=('Arial', 12, 'normal'), values=[
            "Business | Data Analyst",
            "Digital Marketing",
            "USA - Retail | Ecomm",
            "India - Retail | Ecomm",
            "Logistics | Supply Chain",
            "Administration | HR | PR",
            "Other"])
        self.department_entry.grid(row=4, column=1, pady=2)

        # Expense Nature Dropdown
        tk.Label(form_frame, text="Expense Nature", font=font_style).grid(row=5, column=0, sticky=tk.W, pady=2)
        self.nature_entry = ttk.Combobox(form_frame,font=('Arial', 12, 'normal'), values=[
            "One-Time",
            "Recurring",
            "Other"])
        self.nature_entry.grid(row=5, column=1, pady=2)

        # Expense Type Dropdown
        tk.Label(form_frame, text="Expense Type", font=font_style).grid(row=6, column=0, sticky=tk.W, pady=2)
        self.type_entry = ttk.Combobox(form_frame,font=('Arial', 12, 'normal'), values=[
            'Admin',
            'Cost',
            'Logistics',
            'Marketing',
            'Procurement',
            'Professional',
            'Reimbursement',
            'Other'])
        self.type_entry.grid(row=6, column=1, pady=2)

        # Country Dropdown
        tk.Label(form_frame, text="Country", font=font_style).grid(row=7, column=0, sticky=tk.W, pady=2)
        self.country_entry = ttk.Combobox(form_frame,font=('Arial', 12, 'normal'), values=[
            'India',
            'USA',
            'Germany',
            'United Kingdom',
            'United Arab Emirates',
            'Canada'])
        self.country_entry.grid(row=7, column=1, pady=2)

        # Currency Dropdown
        tk.Label(form_frame, text="Currency", font=font_style).grid(row=8, column=0, sticky=tk.W, pady=2)
        self.currency_entry = ttk.Combobox(form_frame,font=('Arial', 12, 'normal'), values=[
            'US Dollar: $',
            'INR:  ₹',
            'Euro: €',
            'British Pound: £',
            'Dirham: د.إ'])
        self.currency_entry.grid(row=8, column=1, pady=2)

        # Invoice Date
        tk.Label(form_frame, text="Invoice Date", font=font_style).grid(row=9, column=0, sticky=tk.W, pady=2)
        self.invoice_date_entry = tk.Entry(form_frame,font=('Arial', 12, 'normal'))
        self.invoice_date_entry.grid(row=9, column=1, pady=2)

        # Invoice Number
        tk.Label(form_frame, text="Invoice Number", font=font_style).grid(row=10, column=0, sticky=tk.W, pady=2)
        self.invoice_number_entry = tk.Entry(form_frame,font=('Arial', 12, 'normal'))
        self.invoice_number_entry.grid(row=10, column=1, pady=2)

        # Party Name
        tk.Label(form_frame, text="Party Name", font=font_style).grid(row=11, column=0, sticky=tk.W, pady=2)
        self.party_name_entry = tk.Entry(form_frame,font=('Arial', 12, 'normal'))
        self.party_name_entry.grid(row=11, column=1, pady=2)

        # Amount
        tk.Label(form_frame, text="Amount", font=font_style).grid(row=12, column=0, sticky=tk.W, pady=2)
        self.amount_entry = tk.Entry(form_frame,font=('Arial', 12, 'normal'))
        self.amount_entry.grid(row=12, column=1, pady=2)

        # Due Date
        tk.Label(form_frame, text="Due Date", font=font_style).grid(row=13, column=0, sticky=tk.W, pady=2)
        self.due_date_entry = tk.Entry(form_frame,font=('Arial', 12, 'normal'))
        self.due_date_entry.grid(row=13, column=1, pady=2)

        # Payment Settlement Dropdown
        tk.Label(form_frame, text="Payment Settlement", font=font_style).grid(row=14, column=0, sticky=tk.W, pady=2)
        self.payment_settlement_entry = ttk.Combobox(form_frame,font=('Arial', 12, 'normal'), values=[
            "25%", "50%", "75%", "100%", "Others"])
        self.payment_settlement_entry.grid(row=14, column=1, pady=2)
        self.payment_settlement_entry.bind("<<ComboboxSelected>>", self.handle_payment_settlement)

        # Description
        tk.Label(form_frame, text="Description", font=font_style).grid(row=15, column=0, sticky=tk.W, pady=2)
        self.description_entry = tk.Entry(form_frame,font=('Arial', 12, 'normal'))
        self.description_entry.grid(row=15, column=1, pady=2)

        # PO Owner
        tk.Label(form_frame, text="PO Owner", font=font_style).grid(row=16, column=0, sticky=tk.W, pady=2)
        self.po_owner_entry = tk.Entry(form_frame,font=('Arial', 12, 'normal'))
        self.po_owner_entry.grid(row=16, column=1, pady=2)

        # PO Date
        tk.Label(form_frame, text="PO Date", font=font_style).grid(row=17, column=0, sticky=tk.W, pady=2)
        self.po_date_entry = tk.Entry(form_frame,font=('Arial', 12, 'normal'))
        self.po_date_entry.grid(row=17, column=1, pady=2)

        # PO Amount
        tk.Label(form_frame, text="PO Amount", font=font_style).grid(row=18, column=0, sticky=tk.W, pady=2)
        self.po_amount_entry = tk.Entry(form_frame,font=('Arial', 12, 'normal'))
        self.po_amount_entry.grid(row=18, column=1, pady=2)

        # PO Details
        tk.Label(form_frame, text="PO Details", font=font_style).grid(row=19, column=0, sticky=tk.W, pady=2)
        self.po_details_entry = tk.Entry(form_frame,font=('Arial', 12, 'normal'))
        self.po_details_entry.grid(row=19, column=1, pady=2)

        # Balance Due
        tk.Label(form_frame, text="Balance Due", font=font_style).grid(row=20, column=0, sticky=tk.W, pady=2)
        self.balance_due_entry = tk.Entry(form_frame,font=('Arial', 12, 'normal'), state="readonly")
        self.balance_due_entry.grid(row=20, column=1, pady=2)
   
        # Attached Files
        tk.Label(form_frame, text="Attached Files", font=font_style).grid(row=21, column=0, sticky=tk.W, pady=2)
        self.attached_files_entry = tk.Entry(form_frame,font=('Arial', 12, 'normal'))
        self.attached_files_entry.grid(row=21, column=1, pady=2)
        
        tk.Button(form_frame, text="Browse", font=('Arial', 12), bg="#f0f0f0", command=self.browse_files).grid(row=21, column=2, padx=2)

        # Buttons
        tk.Button(form_frame, text="Submit", font=('Arial', 14, 'bold'), bg="#7be0ad", command=self.submit_expense).grid(row=22, column=1, pady=10)
        tk.Button(form_frame, text="Clear", font=('Arial', 14), bg="#a39594", command=self.clear_expense_form).grid(row=22, column=2, padx=5)

    def browse_files(self):
        file_path = filedialog.askopenfilename()
        self.attached_files_entry.insert(0, file_path)

    def handle_payment_settlement(self, event):
        selected_value = self.payment_settlement_entry.get()
        if selected_value == "Others":
            custom_value = tk.simpledialog.askfloat("Custom Payment", "Enter custom payment amount:")
            if custom_value is not None:
                self.payment_settlement_entry.set(f"{custom_value}%")
            else:
                self.payment_settlement_entry.set("")

        self.calculate_balance_due()

    def calculate_balance_due(self):
        try:
            amount = float(self.amount_entry.get())
            payment_percentage = self.payment_settlement_entry.get().replace('%', '')
            payment_amount = (float(payment_percentage) / 100) * amount
            balance_due = amount - payment_amount
            self.balance_due_entry.config(state=tk.NORMAL)
            self.balance_due_entry.delete(0, tk.END)
            self.balance_due_entry.insert(0, str(balance_due))
            self.balance_due_entry.config(state=tk.DISABLED)
        except ValueError:
            self.balance_due_entry.config(state=tk.NORMAL)
            self.balance_due_entry.delete(0, tk.END)
            self.balance_due_entry.insert(0, "Invalid")
            self.balance_due_entry.config(state=tk.DISABLED)

    def clear_expense_form(self):
        # Clear all entry fields and comboboxes
        for widget in self.content_frame.winfo_children():
            if isinstance(widget, tk.Entry):
                widget.config(state=tk.NORMAL)  # Make sure entry is writable
                widget.delete(0, tk.END)  # Clear entry content
            elif isinstance(widget, ttk.Combobox):
                widget.set('')  # Clear combobox selection
        
        # Reset the Payment Settlement field in case it was set to "Others"
        self.payment_settlement_entry.set('')

        # Clear balance due field
        self.balance_due_entry.config(state=tk.NORMAL)  # Make it writable
        self.balance_due_entry.delete(0, tk.END)  # Clear content
        self.balance_due_entry.config(state=tk.DISABLED)  # Disable editing again

        # Clear the attached files entry
        self.attached_files_entry.delete(0, tk.END)
        self.email_entry.delete(0, tk.END)
        self.name_entry.delete(0, tk.END)
        self.created_date_entry.delete(0, tk.END)
        self.department_entry.delete(0, tk.END)
        self.nature_entry.delete(0, tk.END)
        self.type_entry.delete(0, tk.END)
        self.country_entry.delete(0, tk.END)
        self.currency_entry.delete(0, tk.END)
        self.invoice_date_entry.delete(0, tk.END)
        self.invoice_number_entry.delete(0, tk.END)
        self.party_name_entry.delete(0, tk.END)
        self.amount_entry.delete(0, tk.END)
        self.due_date_entry.delete(0, tk.END)
        self.payment_settlement_entry.delete(0, tk.END)
        self.description_entry.delete(0, tk.END)
        self.po_owner_entry.delete(0, tk.END)
        self.po_date_entry.delete(0, tk.END)
        self.po_amount_entry.delete(0, tk.END)
        self.po_details_entry.delete(0, tk.END)


    def submit_expense(self):
        # Collect data from all entries
        email = self.email_entry.get()
        name = self.name_entry.get()
        created_date = self.created_date_entry.get()
        department = self.department_entry.get()
        nature = self.nature_entry.get()
        type_expense = self.type_entry.get()
        country = self.country_entry.get()
        currency = self.currency_entry.get()
        invoice_date = self.invoice_date_entry.get()
        invoice_number = self.invoice_number_entry.get()
        party_name = self.party_name_entry.get()
        amount = self.amount_entry.get()
        due_date = self.due_date_entry.get()
        payment_settlement = self.payment_settlement_entry.get()
        description = self.description_entry.get()
        po_owner = self.po_owner_entry.get()
        po_date = self.po_date_entry.get()
        po_amount = self.po_amount_entry.get()
        po_details = self.po_details_entry.get()
        balance_due = self.balance_due_entry.get()
        attached_files = self.attached_files_entry.get()

        if not all([email, name, created_date, department, nature, type_expense, country, currency,
                    invoice_date, invoice_number, party_name, amount, due_date, payment_settlement,
                    description, po_owner, po_date, po_amount, po_details, balance_due,attached_files]):
            messagebox.showerror("Error", "All fields are required!")
            return

        # Insert into database
        self.c.execute('''INSERT INTO Expenses 
                        (email, name, created_date, department, nature, type, country, currency,
                        invoice_date, invoice_number, party_name, amount, due_date, payment_settlement, 
                        description, po_owner, po_date, po_amount, po_details, balance_due, attached_files,status) 
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''', 
                        (email, name, created_date, department, nature, type_expense, country, currency,
                        invoice_date, invoice_number, party_name, amount, due_date, payment_settlement, 
                        description, po_owner, po_date, po_amount, po_details, balance_due, attached_files,'Pending'))

        self.conn.commit()

        messagebox.showinfo("Success", "Expense submitted successfully!")
        self.clear_expense_form()

    def logout(self):
        self.root.destroy()  # Close the current window
        root = tk.Tk()  # Create a new instance of Tk
        app = ExpenseTrackerApp(root)  # Initialize the app again (login page)
        root.mainloop()  # Start the main loop
    
    # tracking Expenses form 
    def track_expenses(self,filter_status="All"):
        self.clear_frame()
        self.open_dashboard()
    
        # Content Frame for Treeview
        filter_frame = tk.Frame(self.root)
        filter_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=5)

        tk.Label(filter_frame, text="Filter by Status:").pack(side=tk.LEFT, padx=5)

        status_options = ["All", "Pending", "Approved", "Rejected"]
        self.status_filter_combobox = ttk.Combobox(filter_frame, values=status_options)
        self.status_filter_combobox.set(filter_status)  # Set current selection
        self.status_filter_combobox.pack(side=tk.LEFT, padx=5)
        self.status_filter_combobox.bind("<<ComboboxSelected>>", lambda event: self.track_expenses(self.status_filter_combobox.get()))

        # Export Button
        tk.Button(filter_frame, text="Export Data", command=self.export_data,
                  bg="#FBB216", fg="#071D3B", activebackground="darkgreen", activeforeground="white").pack(side=tk.RIGHT, padx=5)

        # Content Frame for Treeview
        table_frame = tk.Frame(self.root)
        table_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=False)

        # Create a Treeview widget to display expenses
        columns = ("ID", "Email", "Name", "Created Date", "Department", "Expense Nature", "Expense Type", 
               "Country", "Currency", "Invoice Date", "Invoice Number", "Party Name", 
               "Amount", "Due Date", "Payment Settlement", "Description", "PO Owner", 
               "PO Date", "PO Amount", "PO Details","Attached Files", "Balance Due",  "Aprroval Status")

        tree = ttk.Treeview(table_frame, columns=columns, show='headings')
        tree.grid(row=0, column=0, sticky="nsew")

        # Set column headings
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, anchor=tk.CENTER, width=100)

        # Configure grid for resizing
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        # Create a vertical scrollbar linked to the treeview
        scrollbar_vertical = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
        scrollbar_vertical.grid(row=0, column=1, sticky="ns")
        tree.configure(yscroll=scrollbar_vertical.set)

        # Create a horizontal scrollbar linked to the treeview
        scrollbar_horizontal = ttk.Scrollbar(table_frame, orient="horizontal", command=tree.xview)
        scrollbar_horizontal.grid(row=1, column=0, sticky="ew")
        tree.configure(xscroll=scrollbar_horizontal.set)

        # Fetch updated data from the database
        query ="SELECT * FROM Expenses"
        parameters = []
        
        if not self.full_access:  # Check if the user is not an admin
            query += " WHERE Department = ?"
            parameters.append(self.department)  # Replace with actual department value
            
        if filter_status != "All":
            if "WHERE" in query:
                query += " AND status=?"
            else:
                query += " WHERE status=?"
            parameters.append(filter_status)
            
        self.c.execute(query, parameters)
        rows = self.c.fetchall()

        # Insert the updated rows into the Treeview
        for row in rows:
            tree.insert("", tk.END, values=row)

    def approve_expenses(self):
        self.clear_frame()
        self.open_dashboard()

        # Content Frame for approving expenses
        self.content_frame = tk.Frame(self.root)
        self.content_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # Configuring the grid layout for the content frame
        self.content_frame.grid_rowconfigure(0, weight=1) # Treeview gets most of the space
        self.content_frame.grid_rowconfigure(1, weight=0)  # Buttons don't expand vertically
        self.content_frame.grid_columnconfigure(0, weight=1)

        # Fetching all pending expenses
        self.c.execute("SELECT * FROM Expenses WHERE status='Pending'")
        expenses = self.c.fetchall()

        # Displaying expenses in a treeview with approval options
        columns = ('ID', 'Email', 'Name', 'Created Date', 'Department', 'Nature', 'Type', 'Country', 
                   'Currency', 'Invoice Date', 'Invoice Number', 'Party Name', 'Amount', 'Due Date', 
                   'Payment Settlement', 'Description', 'PO Owner', 'PO Date', 'PO Amount', 
                   'PO Details', 'Attached Files','Balance Due', 'Status')
        
        self.expense_tree = ttk.Treeview(self.content_frame, columns=columns, show='headings')
        self.expense_tree.grid(row=0, column=0, sticky="nsew")  # Make the treeview expand in all directions

        for col in columns:
            self.expense_tree.heading(col, text=col)
            self.expense_tree.column(col, anchor=tk.CENTER, width=100,stretch=True)

        for expense in expenses:
            self.expense_tree.insert('', tk.END, values=expense)

        # Add vertical scrollbar
        scrollbar_vertical = ttk.Scrollbar(self.content_frame, orient="vertical", command=self.expense_tree.yview)
        self.expense_tree.configure(yscroll=scrollbar_vertical.set)
        scrollbar_vertical.grid(row=0, column=1, sticky="ns")

        # Add horizontal scrollbar
        scrollbar_horizontal = ttk.Scrollbar(self.content_frame, orient="horizontal", command=self.expense_tree.xview)
        self.expense_tree.configure(xscroll=scrollbar_horizontal.set)
        scrollbar_horizontal.grid(row=1, column=0, sticky="ew")    

        # Add approve and reject buttons in a frame at the bottom
        button_frame = tk.Frame(self.content_frame)
        button_frame.grid(row=2, column=0, columnspan=2, pady=10)

        # Add approve and reject buttons
        tk.Button(button_frame, text="Approve Selected", command=self.approve_selected,
                  bg="green", fg="white", activebackground="darkgreen", activeforeground="white").pack(side=tk.LEFT, padx=10, pady=10)
        tk.Button(button_frame, text="Reject Selected", command=self.reject_selected,
                  bg="red", fg="white", activebackground="darkred", activeforeground="white").pack(side=tk.LEFT, padx=10, pady=10)

    def approve_selected(self):
        selected_items = self.expense_tree.selection()
        for item in selected_items:
            expense_id = self.expense_tree.item(item, 'values')[0]
            self.c.execute("UPDATE Expenses SET status='Approved' WHERE id=?", (expense_id,))
        self.conn.commit()
        self.refresh_expense_tracking()

    def reject_selected(self):
        selected_items = self.expense_tree.selection()
        for item in selected_items:
            expense_id = self.expense_tree.item(item, 'values')[0]
            self.c.execute("UPDATE Expenses SET status='Rejected' WHERE id=?", (expense_id,))
        self.conn.commit()
        self.refresh_expense_tracking()

    def refresh_expense_tracking(self):
        # Refresh the tracking expenses view
        self.track_expenses()

    def export_data(self):
        # Fetching all data from the treeview
        data = [self.tree.item(item)["values"] for item in self.tree.get_children()]
        
        # Get columns from the Treeview
        columns = [self.tree.heading(col)["text"] for col in self.tree["columns"]]
        
        # Create a pandas DataFrame
        df = pd.DataFrame(data, columns=columns) # type: ignore

        # Ask user for file format and path
        file_path = filedialog.asksaveasfilename(defaultextension=".csv",
                                                filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx")],
                                                title="Save as")
        if file_path:
            if file_path.endswith(".csv"):
                df.to_csv(file_path, index=False)
            elif file_path.endswith(".xlsx"):
                df.to_excel(file_path, index=False)

            messagebox.showinfo("Success", f"Data successfully exported to {file_path}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExpenseTrackerApp(root)
    root.mainloop()