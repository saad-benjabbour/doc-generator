from docxtpl import DocxTemplate

doc = DocxTemplate("invoice_template.docx")

# THIS REPRESENT EACH ROW
invoice_list = [[2, "pen", 0.5, 1],
                [3, "paper", 0.6, 21],
                [5, "paper pack", 0.7, 1],
                [8, "notebook", 0.9, 1]]

doc.render({"name": "saad", 
            "phone":"555-555555",
            "invoice_list":invoice_list,
            "subtotal":10,
            "salestax":"10%",
            "total":9
            })
doc.save("new_invoice.docx")