from PIL import Image
import pytesseract
import sys
from pdf2image import convert_from_path
import os
import sys
import xlwt
from xlwt import Workbook

# Path of the pdf
from InvoiceSummary import InvoiceSummary

PDF_file = "/Users/geetparekh/Documents/Suchi/XYZInc.pdf"

''' 
Part #1 : Converting PDF to images 
'''

# Store all the pages of the PDF in a variable
pages = convert_from_path(PDF_file, 500)

# Counter to store images of each page of PDF to image
image_counter = 1

# Iterate through all the pages stored above
for page in pages:
    # Declaring filename for each page of PDF as JPG
    # For each page, filename will be:
    # PDF page 1 -> page_1.jpg
    # PDF page 2 -> page_2.jpg
    # PDF page 3 -> page_3.jpg
    # ....
    # PDF page n -> page_n.jpg
    filename = "page_" + str(image_counter) + ".jpg"

    # Save the image of the page in system
    page.save(filename, 'JPEG')

    # Increment the counter to update filename
    image_counter = image_counter + 1

''' 
Part #2 - Recognizing text from the images using OCR 
'''
3
# Variable to get count of total number of pages
filelimit = image_counter - 1

# Creating a text file to write the output
outfile = "out_text.txt"

# Open the file in append mode so that
# All contents of all images are added to the same file
f = open(outfile, "w")

# Iterate from 1 to total number of pages
for i in range(1, filelimit + 1):
    # Set filename to recognize text from
    # Again, these files will be:
    # page_1.jpg
    # page_2.jpg
    # ....
    # page_n.jpg
    filename = "page_" + str(i) + ".jpg"

    # Recognize the text as string in image using pytesserct
    text = str(((pytesseract.image_to_string(Image.open(filename)))))

    # The recognized text is stored in variable text
    # Any string processing may be applied on text
    # Here, basic formatting has been done:
    # In many PDFs, at line ending, if a word can't
    # be written fully, a 'hyphen' is added.
    # The rest of the word is written in the next line
    # Eg: This is a sample text this word here GeeksF-
    # orGeeks is half on first line, remaining on next.
    # To remove this, we replace every '-\n' to ''.
    text = text.replace('-\n', '')

    # Finally, write the processed text to the file.
    f.write(text)

# Close the file after writing all the text.
f.close()


def identify_template():
    with open("out_text.txt") as fp:
        cnt = 0
        check_template1 = False
        for line in fp:
            line = line.strip()
            if not line:
                continue
            if cnt > 50:
                return 0
            print("line {} contents {}".format(cnt, line))

            if check_template1:
                match_str = "Code Description Use Provider Loc ID Period Quantity Unit Price Amount"
                if match_str == line:
                    return 1
                else:
                    check_template1 = False
            if "License" in line:
                check_template1 = True
            cnt += 1
    fp.close()
    return 0


template = identify_template()
print(template)


def parse_based_on_template1():
    summary_list = []
    invoice_summary = None
    match_str = "Code Description Use Provider Loc ID Period Quantity Unit Price Amount"
    match_str_count = 0
    sub_total_without_tax = 0
    fetch_company_summary = False
    fetch_sub_total = False

    # Workbook is created
    wb = Workbook()

    # add_sheet is used to create sheet.
    sheet1 = wb.add_sheet('Sheet 1')

    # sheet1.write(1, 0, 'ISBT DEHRADUN')
    # sheet1.write(2, 0, 'SHASTRADHARA')
    # sheet1.write(3, 0, 'CLEMEN TOWN')
    # sheet1.write(4, 0, 'RAJPUR ROAD')
    # sheet1.write(5, 0, 'CLOCK TOWER')
    sheet1.write(0, 0, 'Company Name')
    sheet1.write(0, 1, 'City')
    sheet1.write(0, 2, 'State')
    sheet1.write(0, 3, 'Zip Code')
    sheet1.write(0, 4, 'Country')
    sheet1.write(0, 5, 'Subtotal')

    # wb.save('xlwt example.xls')

    row_counter = 1
    company_name = None
    city = None
    state = None
    zip_code = None
    country = None

    with open("out_text.txt") as fp:
        for line in fp:
            line = line.strip()
            if not line:
                continue
            if fetch_company_summary:
                company_summary = line.split(',')
                if len(company_summary) < 5:
                    continue

                company_name = company_summary[0].strip()
                city = company_summary[-4].strip()
                state = company_summary[-3].strip()
                zip_code = company_summary[-2].strip()
                country = company_summary[-1].strip()

                print("Company Name: ", company_name)
                print("City: ", city)
                print("State: ", state)
                print("Zip Code: ", zip_code)
                print("Country: ", country)

                # block summary
                invoice_summary = InvoiceSummary(company_name,
                                                 city,
                                                 state,
                                                 zip_code,
                                                 country,
                                                 )
                # TODO: get tax based on the zip code and set it in the invoice summary
                fetch_company_summary = False
                fetch_sub_total = True
                continue

            if fetch_sub_total:
                if "Subtotal" in line:
                    sub_total_line = line.split()
                    print("Subtotal: ", sub_total_line[-1])

                    sheet1.write(row_counter, 0, company_name)
                    sheet1.write(row_counter, 1, city)
                    sheet1.write(row_counter, 2, state)
                    sheet1.write(row_counter, 3, zip_code)
                    sheet1.write(row_counter, 4, country)
                    sheet1.write(row_counter, 5, sub_total_line[-1])
                    row_counter = row_counter + 1

                    sub_total_without_tax = sub_total_without_tax + float(sub_total_line[-1].replace(',', ''))
                    print("Subtotal without tax: ", sub_total_without_tax)

                    print("---------------------------------------------------------------")
                    invoice_summary.set_sub_total(sub_total_line[-1])
                    # TODO: compute subtotal with tax and set it to the invoice summary
                    summary_list.append(invoice_summary)
                    fetch_sub_total = False
                    fetch_company_summary = True
                    continue

            if match_str == line and match_str_count == 0:
                fetch_company_summary = True
                match_str_count = 1
                continue

            # if "Please remit to:" in line:
            #    break
    fp.close()
    print("Subtotal without tax: ", sub_total_without_tax)
    sheet1.write(row_counter, 5, sub_total_without_tax)
    wb.save('xlwt example.xls')


if template == 1:
    parse_based_on_template1()
