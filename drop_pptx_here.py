import zipfile
import os, sys
import re, tempfile

def rm_txt(str):
    return re.sub(r'<p:txBody>.*?<\/p:txBody>', '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody>', str)

def main(fn):
    print("Processing %s -- please wait." % fn)
    if fn[-5:] != '.pptx':
        raise RuntimeError("Files need to be .pptx files.")
    fn2 = fn.replace(".pptx","-clean.pptx")
    with zipfile.ZipFile(fn, "r") as inzip, zipfile.ZipFile(fn2, "w") as outzip:
        for inzipinfo in inzip.infolist():
            with inzip.open(inzipinfo) as infile:
                if inzipinfo.filename.startswith("ppt/notesSlides/notesSlide") and inzipinfo.filename.endswith(".xml"):
                    data = infile.read().decode('utf-8','ignore')
                    print(". . .", "cleaning", inzipinfo.filename)
                    data = rm_txt(data)
                    outzip.writestr(inzipinfo.filename, data)
                else: 
                    outzip.writestr(inzipinfo.filename, infile.read())

if __name__ == '__main__':
    try:
        if len(sys.argv) > 1:
            for arg in sys.argv[1:]:
                fn = os.path.abspath(arg)
                main(fn)
                print("---")
        else:
            print("You need to drag your .pptx file(s) onto this one.")
    finally:
        input("Press any key to quit.")
