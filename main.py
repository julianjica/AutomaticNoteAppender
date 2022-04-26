from PyPDF2 import PdfFileWriter, PdfFileReader, PdfFileMerger
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import pandas as pd
import os

def split_pdf(pdf_path):
    # Main reference: 
    #https://stackoverflow.com/questions/490195/split-a-multi-page-pdf-file-into-multiple-pdf-files-with-python
    inputpdf = PdfFileReader(open(pdf_path, "rb"))
    for i in range(inputpdf.numPages):
        output = PdfFileWriter()
        output.addPage(inputpdf.getPage(i))
        with open("Data/ops/%s.pdf" % i, "wb") as outputStream:
            output.write(outputStream) 

# Checking the files inside Data
excel = []
pdf = []
for file in os.listdir("Data/"):
    if file.endswith(".xlsx"):
        excel.append(file)
    elif file.endswith(".pdf"):
        pdf.append(file)
# Doing the operations for every file
for file in excel:
    print(file)
    data = pd.read_excel("Data/"+file)
    data = data.iloc[:, :6]
    # Splitting PDF
    split_pdf("Data/"+file[:-5]+".pdf")
    shape = data.shape[0] - 60
    shape = [30, 30, shape]
    
    # Writting in each splitting segment
    for ii in range(0,3):
        packet = io.BytesIO()
        # Create a new PDF with Reportlab
        can = canvas.Canvas(packet, pagesize=letter)
        can.setFont('Helvetica', 12.5)
        # Measurements are made in points (GIMP)
        # First Column: 330, Second Column: 370, Third Column: 410,
        # Fourth Column: 450
        # First level: 638, Distance between levels: 20
        count = 2
        for jj in range(330, 451, 40):
            for kk in range(0, shape[ii]):
                try:
                    can.drawString(jj, 628 - 18 * kk, "%1.2f"%data.iloc[30*ii+kk,
                        count])
                except:
                    print(data.iloc[30*ii+kk,-3 + count], 30*ii+kk,-3 + count)
            count += 1
        can.showPage()
        can.save()

        # Move to the beginning of the StringIO buffer
        packet.seek(0)
        new_pdf = PdfFileReader(packet)
        # Read your existing PDF
        existing_pdf = PdfFileReader(open("Data/ops/%s.pdf"%ii, "rb"))
        output = PdfFileWriter()
        # Add the "watermark" (which is the new pdf) on the existing page
        page = existing_pdf.getPage(0)
        page.mergePage(new_pdf.getPage(0))
        output.addPage(page)
        # Finally, write "output" to a real file
        outputStream = open("Data/ops/out_%s.pdf"%ii, "wb")
        output.write(outputStream)
        outputStream.close()

    # Merge results and save
    
    pdfs = ["Data/ops/out_%s.pdf"%ii for ii in range(3)]
    merger = PdfFileMerger()
    for pdf in pdfs:
        merger.append(pdf)
    merger.write("Data/Result/%s.pdf"%file[:-5])
    merger.close()

    # Remove unparent files
    for ii in range(0,3):
        os.remove("Data/ops/%s.pdf"%ii)
        os.remove("Data/ops/out_%s.pdf"%ii)
    
