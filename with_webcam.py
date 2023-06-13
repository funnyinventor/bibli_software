import cv2
from pyzbar.pyzbar import decode
import openpyxl
from datetime import datetime
import requests

def scan_barcode(frame):
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    barcode_data = decode(gray)
    for data in barcode_data:
        if data.type == 'EAN13':
            return data.data.decode('utf-8')

def get_book_details(isbn):
    url = f"https://www.googleapis.com/books/v1/volumes?q=isbn:{isbn}"
    response = requests.get(url)
    data = response.json()
    if 'items' in data:
        book = data['items'][0]['volumeInfo']
        title = book.get('title', 'N/A')
        author = ', '.join(book.get('authors', ['N/A']))
        return title, author
    else:
        return None, None

def save_to_excel(isbn, title, author, borrower):
    workbook = openpyxl.load_workbook('library_records.xlsx')
    sheet = workbook.active

    row = (isbn, title, author, borrower, datetime.now())

    sheet.append(row)
    workbook.save('library_records.xlsx')

cap = cv2.VideoCapture(0)

while True:
    ret, frame = cap.read()
    if not ret:
        print("Erreur lors de la capture de l'image de la webcam.")
        break

    cv2.imshow('Barcode Scanner', frame)

    if cv2.waitKey(1) == ord('q'):
        break

    barcode_data = scan_barcode(frame)
    if barcode_data:
        cap.release()
        cv2.destroyAllWindows()
        title, author = get_book_details(barcode_data)
        if title:
            borrower = input("Entrez le nom de l'emprunteur : ")
            save_to_excel(barcode_data, title, author, borrower)
            print("Livre enregistré avec succès.")
        else:
            print("Livre introuvable pour l'ISBN spécifié.")
        break

cap.release()
cv2.destroyAllWindows()
