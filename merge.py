import os
import glob
import comtypes.client
from PyPDF2 import PdfFileMerger
import time
import shutil
from docx import Document
from docx.shared import Inches

if os.path.exists('files'):
	shutil.rmtree('files')

if not os.path.exists('files'):
  os.mkdir('files')

if not os.path.exists('resultPDF'):
  os.mkdir('resultPDF')

def docs_to_pdf():
	word = comtypes.client.CreateObject('Word.Application')
	pdfslist = PdfFileMerger()
	x = 0
	for f in glob.glob("*"):
		split_tup = os.path.splitext(f)[1]
		if split_tup == ".docx":
			input_file = os.path.abspath(f)
			output_file = os.path.abspath("files/demo" + str(x) + ".pdf")
			doc = word.Documents.Open(input_file)
			doc.SaveAs(output_file, FileFormat=16+1)
			doc.Close()
			pdfslist.append(open(output_file, 'rb'))
			x += 1

		elif split_tup == ".jpg" or split_tup == ".jpeg" or split_tup == ".png":
			input_file = os.path.abspath(f)
			document = Document()
			document.add_picture(f)
			document.save("files/images.docx")
			input_file = os.path.abspath("files/images.docx")
			output_file = os.path.abspath("files/demo" + str(x) + ".pdf")
			doc = word.Documents.Open(input_file)
			doc.SaveAs(output_file, FileFormat=16+1)
			doc.Close()
			pdfslist.append(open(output_file, 'rb'))
			x += 1

		elif split_tup == ".pdf":
			src_dir= os.curdir
			dst_dir= os.path.join(os.curdir , "files")
			src_file = os.path.join(src_dir, f)
			shutil.copy(src_file,dst_dir)
			dst_file = os.path.join(dst_dir, f)
			new_dst_file_name = os.path.join(dst_dir, "demo" + str(x) + ".pdf")
			os.rename(dst_file, new_dst_file_name)
			output_file = os.path.abspath("files/demo" + str(x) + ".pdf")
			pdfslist.append(open(output_file, 'rb'))
			x += 1

	word.Quit()
	return pdfslist

def joinpdf(pdfs):
	with open("resultPDF/result.pdf", "wb") as result_pdf:
		pdfs.write(result_pdf)
	pdfs.close()

def main():
	pdfs = docs_to_pdf()
	joinpdf(pdfs)

main()