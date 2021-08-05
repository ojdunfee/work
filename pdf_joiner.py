from PyPDF2 import PdfFileMerger, PdfFileReader
import PyPDF2
import os

def merge_td(dir):
    """
    scan a folder of pdf files combining the cover page of bank statements with pdf documents provided by coworker.
    """
    merger = PdfFileMerger()
    if not len(os.listdir(dir)) % 2 == 0:
        print('unequal lengths')
        return None

    beg = sorted([file for file in os.listdir(dir) if file.endswith('XFER.pdf')])
    end = sorted([file for file in os.listdir(dir) if file.endswith('FEES.pdf')])

    for i, file in enumerate(beg):
        pdf1 = open(dir + '/' + file, 'rb')
        pdf2 = open(dir + '/' + end[i], 'rb')

        pdf1Reader = PdfFileReader(pdf1)
        pdf2Reader = PdfFileReader(pdf2)

        pdfWriter = PyPDF2.PdfFileWriter()

        for pageNum in range(pdf1Reader.numPages):
            pageObj = pdf1Reader.getPage(pageNum)
            pdfWriter.addPage(pageObj)

        for pageNum in range(pdf2Reader.numPages):
            pageObj = pdf2Reader.getPage(pageNum)
            pdfWriter.addPage(pageObj)

        pdfOutputFile = open('fees/{}'.format(end[i]), 'wb')
        pdfWriter.write(pdfOutputFile)

        pdfOutputFile.close()
        pdf1.close()
        pdf2.close()

def merge_crystal(dir):
    """
    Merge pdf documents for accounts receivable posting backup.
    """
    merger = PdfFileMerger()
    beg = sorted([file for file in os.listdir(dir) if not file.endswith('FEES.pdf')])
    end = sorted([file for file in os.listdir(dir) if file.endswith('FEES.pdf')])

    for i, file in enumerate(beg):
        pdf1 = open(dir + '/' + file, 'rb')
        pdf2 = open(dir + '/' + end[i], 'rb')

        pdf1Reader = PdfFileReader(pdf1)
        pdf2Reader = PdfFileReader(pdf2)

        pdfWriter = PyPDF2.PdfFileWriter()

        for pageNum in range(pdf1Reader.numPages):
            pageObj = pdf1Reader.getPage(pageNum)
            pdfWriter.addPage(pageObj)

        for pageNum in range(pdf2Reader.numPages):
            pageObj = pdf2Reader.getPage(pageNum)
            pdfWriter.addPage(pageObj)

        pdfOutputFile = open('final/{}'.format(beg[i]), 'wb')
        pdfWriter.write(pdfOutputFile)

        pdfOutputFile.close()
        pdf1.close()
        pdf2.close()
