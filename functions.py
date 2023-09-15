import main
# GETTING THE VALUES FROM THE INPUTS
def add_item():
    qty = int(main.qt_spinbox.get())
    desc = main.derscription_entry
    price = float(main.up_spinbox.get())
    line_total = qty * price
    invoice_item = [qty, desc, price, line_total]
    main.treeview.insert('', 0, values = invoice_item)