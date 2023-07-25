# сгребаем из текущей папки все документы .docx , конвертируем их в pdf
# и склеиваем

import os
import sys
from pathlib import Path

import docx2pdf
import pypdf

def main():
    foldermask = 'temp'
    outputPath = Path('./buffer/out.pdf')

    hasFiles=False
    for i in os.listdir(foldermask):
        if i.endswith('.docx'):
            hasFiles=True
    if not hasFiles:
        print('No files to convert. Shutting down.')
        return

    print('converting...')
    for i in os.listdir(foldermask):
        if i.endswith('.docx'):
            tmp=docx2pdf.convert(Path(foldermask)/i,Path(foldermask)/f"{i[:-5]}.pdf")
    
    print ('merging...')
    pile = pypdf.PdfMerger()
    for i in os.listdir(foldermask):
        if i.endswith('.pdf'):
            pile.append(Path(foldermask)/i)
    
    print('cleaning up...')
    try:
        os.unlink(outputPath)
    except FileNotFoundError:
        pass # ну и ладно
    pile.write(outputPath.as_posix())
    pile.close()
    for i in os.listdir(foldermask):
        os.unlink(Path(foldermask)/i)
    print('complete!')
if __name__ == '__main__':
    main()