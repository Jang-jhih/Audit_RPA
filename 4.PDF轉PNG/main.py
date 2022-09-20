import os
from pdf2image import convert_from_path

Poppler = os.path.join('poppler-0.68.0','bin')
PDFlist = os.listdir(os.path.join('PDF'))
PDFlist = [os.path.join('PDF',_) for _ in PDFlist]


PNGlist = [_.replace('PDF','Export').split('.')[0] for _ in PDFlist]



for pdf,png in zip(PDFlist,PNGlist):
    images = convert_from_path(pdf, 500,poppler_path=Poppler)
    for i, image in enumerate(images):
        fname = f'{png}_page{i}.png'
        image.save(fname, "PNG")
