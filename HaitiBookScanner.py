# isbntools documentation found at https://isbntools.readthedocs.io/en/latest/info.html
# Using this too: https://stackoverflow.com/questions/26360699/how-can-i-get-the-author-and-title-by-knowing-the-isbn-using-google-book-api
import sys
import openpyxl
import time
from isbntools.app import *

# Service set for Google Books
service = "goob"

# Main menu
def main():
    print("Main Menu:")
    print("To register new books, type 'register'")
    print("To check in books, type 'check in'")
    print("To check out books, type 'check out'")
    choice = input("")
    if choice in ["register", "Register", "REGISTER", "'register'"]:
        register_book()
    elif choice in ["check in", "checkin", "Check In", "Check in", "CHECKIN", "CHECK IN", "CheckIn"]:
        check_in()
    elif choice in ["check out", "checkout", "Check Out", "Check out", "CHECKOUT", "CHECK OUT", "CheckOut"]:
        check_out()
# Allows for registration of books into database.
def register_book():
    inventory_workbook = openpyxl.load_workbook("BookDatabase.xlsx")
    book_inventory_sheet = inventory_workbook["Book Inventory"]
    book = input("Scan barcode or type menu: ")
    if book == "menu":
        main()
    else:
        isbn_list = []
        for col in book_inventory_sheet["C"]:
            isbn_list.append(col.value)
        if book in isbn_list:
            #TODO add code to increase "Total Quantity" and "In Stock" by 1
            print("Exists")
        else:
            meta_dict = meta(book, service)
            authors_list = meta_dict["Authors"]
            authors = ",".join(authors_list)
            title = meta_dict["Title"]
            # Appends the info to the last column, and sets "Total Quantity" and "In Stock" to 1
            book_inventory_sheet.append([title, authors, book, "1", "1"])
            inventory_workbook.save("BookDatabase.xlsx")
            main()

        
# Allows the library to scan books in once returned.
def check_in():
    print ("Check in!")
# Allows the library to scan books when checked out.
def check_out():
    print ("Check out!")

main()
