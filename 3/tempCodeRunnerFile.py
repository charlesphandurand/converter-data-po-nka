def browse_files(entry):
    filenames = filedialog.askopenfilenames()
    entry.delete(0, tk.END)
    entry.insert(0, ';'.join(filenames))