        ctk.CTkLabel(tab, text="Customer Code:").grid(row=0, column=0, padx=10, pady=(20, 10), sticky="w")
        self.farmer_customer_var = ctk.StringVar(value="30401154 - BI")
        self.farmer_customer_dropdown = ctk.CTkOptionMenu(tab, variable=self.farmer_customer_var, values=[
            # "30103270 - PBM1 (FARMERS/FM SCP)",
            # "30105314 - PBM1 (FARMERS/MESRA INDAI)",
            # "30202092 - PBM2 (FARMERS/FM SCP)",
            # "30203407 - PBM2 (FARMERS/MESRA INDAI)",
            "30401154 - BI"
        ])
        self.farmer_customer_dropdown.grid(row=0, column=1, padx=10, pady=(20, 10), sticky="ew")

        ctk.CTkLabel(tab, text="File CSV:").grid(row=1, column=0, padx=10, pady=10, sticky="w")
        self.farmer_csv_entry = ctk.CTkEntry(tab)
        self.farmer_csv_entry.grid(row=1, column=1, padx=10, pady=10, sticky="ew")
        self.farmer_csv_entry.insert(0, "C:/Users/TOSHIBA PORTEGE Z30C/Desktop/program python/farmer/PurchaseOrder_3011601648 farmer.csv")
        ctk.CTkButton(tab, text="Browse", command=lambda: browse_files(self.farmer_csv_entry, "csv")).grid(row=1, column=2, padx=(0, 20), pady=10, sticky="e")

        ctk.CTkLabel(tab, text="File Excel Master Data:").grid(row=2, column=0, padx=10, pady=10, sticky="w")
        self.farmer_excel_entry = ctk.CTkEntry(tab)
        self.farmer_excel_entry.grid(row=2, column=1, padx=10, pady=10, sticky="ew")
        self.farmer_excel_entry.insert(0, "C:/Users/TOSHIBA PORTEGE Z30C/Desktop/program python/farmer/NKA.xls")
        ctk.CTkButton(tab, text="Browse", command=lambda: browse_files(self.farmer_excel_entry, "excel")).grid(row=2, column=2, padx=(0, 20), pady=10, sticky="e")

        ctk.CTkLabel(tab, text="Direktori Output:").grid(row=3, column=0, padx=10, pady=10, sticky="w")
        self.farmer_output_entry = ctk.CTkEntry(tab)
        self.farmer_output_entry.grid(row=3, column=1, padx=10, pady=10, sticky="ew")
        self.farmer_output_entry.insert(0, "C:/Users/TOSHIBA PORTEGE Z30C/Desktop/program python/farmer")
        ctk.CTkButton(tab, text="Browse", command=lambda: browse_directory(self.farmer_output_entry)).grid(row=3, column=2, padx=(0, 20), pady=10, sticky="e")

        # Process Button
        ctk.CTkButton(tab, text="Proses", command=process_farmer_files).grid(row=4, column=0, columnspan=3, padx=10, pady=(20, 10), sticky="ew")
