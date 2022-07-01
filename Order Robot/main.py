# Import Essential Modules

import sys
import subprocess
import os

# Install Required Modules

subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'selenium-driver-updater'])
subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pandas'])
subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'selenium'])
subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'fpdf'])
subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'tk'])

# Import Modules

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select

from tkinter import *
from tkinter.ttk import Combobox
from tkinter import messagebox

import pandas as pd

from fpdf import FPDF


# Create GUI

browserTypes = ["Google Chrome", "Mozilla Firefox"]

root = Tk()

# Order List 

orderListLabel = Label(root, text = "Order List URL: ")
orderListLabel.place(x = 50, y = 50)

orderList = Entry(root, bd = 5)
orderList.place(x = 150, y = 50)

pasteButtonOL = Button(root, text = "Paste URL", command = lambda: orderList.event_generate("<<Paste>>"))
pasteButtonOL.place(x = 300, y = 50)

# Order Form

orderFormURLLabel = Label(root, text = "Order Form URL: ")
orderFormURLLabel.place(x = 50, y = 100)

orderFormURL = Entry(root, bd = 5)
orderFormURL.place(x = 150, y = 100)

pasteButtonOF = Button(root, text = "Paste URL", command = lambda: orderFormURL.event_generate("<<Paste>>"))
pasteButtonOF.place(x = 300, y = 100)

# Browser Selection

browsersLabel = Label(root, text = "Preferred Browser: ")
browsersLabel.place(x = 50, y = 150)

browsers = Combobox(root, values = browserTypes)
browsers.place(x = 200, y = 150)

# Bot Progress

progress = Label(root, text = "The Bot is Ready!")
progress.place(x = 145, y = 250)

# Create Order Bot Function for Button

def orderBot():

    # Checking Data

    if orderList.get() == "":
        messagebox.showerror("Missing Information", "Please insert the link to the CSV file.")

    if orderFormURL.get() == "":
        messagebox.showerror("Missing Information", "Please insert the link to the order form.")

    if browsers.get() == "":
        messagebox.showerror("Missing Information", "Please select your preferred browser type.")
    
    # Fetch Data

    orderLink = orderList.get()
    data = pd.read_csv(orderLink, usecols = range(1, 5)).values
    numOrders = len(data)

    orderURL = orderFormURL.get()
    browserSelection = browsers.get()

    # Update Drivers

    if browserSelection == "Google Chrome":
        webDriver = webdriver.Chrome()
    elif browserSelection == "Mozilla Firefox":
        webDriver = webdriver.Firefox()

    # Open Driver

    webDriver.get(orderURL)

    # Loop Through Orders

    for i in range(numOrders):
        
        # Update Progress Bar

        progress.configure(text = "{} out of {} orders complete.".format(i + 1, numOrders) )

        # Click the OK Button

        webDriver.find_element(By.CLASS_NAME, "btn-dark").click()  # Click Accept Button

        # Body Selection
        # Head
        try:
            head = Select(webDriver.find_element(By.ID, "head"))
            head.select_by_value(str(data[i][0]))

            # Body

            webDriver.find_element(By.XPATH, "//*[@id='id-body-" + str(data[i][1]) +"']").click()

            # Leg

            webDriver.find_element(By.CLASS_NAME, "form-control").send_keys(str(data[i][2]))

            # Shipping Address

            webDriver.find_element(By.ID, "address").send_keys(str(data[i][3]))

            # Preview Order

            webDriver.find_element(By.ID, "preview").click()

            # Ordering Robot

            webDriver.find_element(By.ID, "order").click()

            # Fetching Receipt Data
                                                            
            timeDate    = webDriver.find_element(By.XPATH, "/html/body/div/div/div[1]/div/div[1]/div/div/div[1]").text
            orderNum    = webDriver.find_element(By.CLASS_NAME, "badge").text
            address     = webDriver.find_element(By.XPATH, "/html/body/div/div/div[1]/div/div[1]/div/div/p[2]").text
            headPart    = webDriver.find_element(By.XPATH, "/html/body/div/div/div[1]/div/div[1]/div/div/div[2]/div[1]").text
            bodyPart    = webDriver.find_element(By.XPATH, "/html/body/div/div/div[1]/div/div[1]/div/div/div[2]/div[2]").text
            legPart     = webDriver.find_element(By.XPATH, "/html/body/div/div/div[1]/div/div[1]/div/div/div[2]/div[3]").text
            thanksMSG   = webDriver.find_element(By.XPATH, "/html/body/div/div/div[1]/div/div[1]/div/div/p[3]").text

            # Fetching Image

            image = webDriver.find_element(By.ID, "robot-preview-image").screenshot_as_png

            with open('Image.png', 'wb') as file:
                file.write(image)

            # Create PDF

            pdf = FPDF()
            pdf.add_page()

            pdf.set_font("Arial", 'B', size = 15)
            pdf.set_text_color(r = 255, g = 0, b = 0)
            pdf.cell(200, 5, txt = "Receipt", ln = 1)

            pdf.set_font("Arial", size = 12)
            pdf.set_text_color(r = 1, g = -1, b = -1)
            pdf.cell(200, 5, ln = 2)

            pdf.cell(200, 5, txt = timeDate, ln = 3)
            pdf.cell(200, 5, ln = 4)

            pdf.cell(200, 5, txt = "Order Number: " + orderNum, ln = 5)
            pdf.cell(200, 5, txt = "Delivery Address: " + address, ln = 6)
            pdf.cell(200, 5, ln = 7)

            pdf.cell(200, 5, txt = "Robot Parts:", ln = 8)
            pdf.cell(200, 5, txt = "\t\t\t\t" + headPart, ln = 9)
            pdf.cell(200, 5, txt = "\t\t\t\t" + bodyPart, ln = 10)
            pdf.cell(200, 5, txt = "\t\t\t\t" + legPart, ln = 11)
            pdf.cell(200, 5, ln = 12)

            pdf.multi_cell(200, 5, thanksMSG)

            pdf.image("Image.png", 50, w = 100, h = 150)
            pdf.output("Output/Receipt {}.pdf".format(i + 1))

        except:
            while True:
                if len(webDriver.find_elements(By.ID, "order")) > 0:
                    webDriver.find_element(By.ID, "order").click()
                else:
                    break
                
            # Fetching Receipt Data

            timeDate    = webDriver.find_element(By.XPATH, "/html/body/div/div/div[1]/div/div[1]/div/div/div[1]").text
            orderNum    = webDriver.find_element(By.CLASS_NAME, "badge").text
            address     = webDriver.find_element(By.XPATH, "/html/body/div/div/div[1]/div/div[1]/div/div/p[2]").text
            headPart    = webDriver.find_element(By.XPATH, "/html/body/div/div/div[1]/div/div[1]/div/div/div[2]/div[1]").text
            bodyPart    = webDriver.find_element(By.XPATH, "/html/body/div/div/div[1]/div/div[1]/div/div/div[2]/div[2]").text
            legPart     = webDriver.find_element(By.XPATH, "/html/body/div/div/div[1]/div/div[1]/div/div/div[2]/div[3]").text
            thanksMSG   = webDriver.find_element(By.XPATH, "/html/body/div/div/div[1]/div/div[1]/div/div/p[3]").text

            # Fetching Image

            image = webDriver.find_element(By.ID, "robot-preview-image").screenshot_as_png

            with open('Image.png', 'wb') as file:
                file.write(image)

            # Create PDF

            pdf = FPDF()
            pdf.add_page()

            pdf.set_font("Arial", 'B', size = 15)
            pdf.set_text_color(r = 255, g = 0, b = 0)
            pdf.cell(200, 5, txt = "Receipt", ln = 1)

            pdf.set_font("Arial", size = 12)
            pdf.set_text_color(r = 1, g = -1, b = -1)
            pdf.cell(200, 5, ln = 2)

            pdf.cell(200, 5, txt = timeDate, ln = 3)
            pdf.cell(200, 5, ln = 4)

            pdf.cell(200, 5, txt = "Order Number: " + orderNum, ln = 5)
            pdf.cell(200, 5, txt = "Delivery Address: " + address, ln = 6)
            pdf.cell(200, 5, ln = 7)

            pdf.cell(200, 5, txt = "Robot Parts:", ln = 8)
            pdf.cell(200, 5, txt = "\t\t\t\t" + headPart, ln = 9)
            pdf.cell(200, 5, txt = "\t\t\t\t" + bodyPart, ln = 10)
            pdf.cell(200, 5, txt = "\t\t\t\t" + legPart, ln = 11)
            pdf.cell(200, 5, ln = 12)

            pdf.multi_cell(200, 5, thanksMSG)

            pdf.image("Image.png", 50, w = 100, h = 150)
            pdf.output("Output/Receipt {}.pdf".format(i + 1))

        # Order New Robot

        webDriver.find_element(By.ID, "order-another").click()

    os.remove("Image.png")

    from zipfile import ZipFile

    currentDir = os.getcwd()
    os.chdir("Output")

    Receipts = os.listdir()

    with ZipFile("Receipts.zip", 'w') as archive:
        for i in Receipts:
            if ".pdf" in i:
                archive.write(i)
            else:
                continue

    for i in Receipts:
        if ".pdf" in i:
            os.remove(i)

    os.chdir(currentDir)

startButton = Button(root, text = "Start Bot", bg = 'lime', command = orderBot)
startButton.place(x = 170, y = 200)

root.title("Orderbot")
root.geometry("400x300+10+20")

root.mainloop()
