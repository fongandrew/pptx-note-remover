import zipfile
import os, sys
import re, tempfile

def rm_txt(str):
    return re.sub(r'<p:txBody>.*</p:txBody>', '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody>', str)

def main(fn):
    print "Processing %s -- please wait." % fn
    if fn[-5:] != '.pptx':
        raise RuntimeError("Files need to be .pptx files.")
    old = zipfile.ZipFile(fn, "r")
    fn2 = fn.replace(".pptx","-clean.pptx")
    new = zipfile.ZipFile(fn2, "w")
    for item in old.infolist():
        data = old.read(item.filename)
        if item.filename.startswith("ppt/notesSlides/notesSlide") \
                and item.filename.endswith(".xml"):
            print ". . .", "cleaning", item.filename
            data = rm_txt(data)
        new.writestr(item, data)
    new.close()
    old.close()
    print "Complete. Saved as", fn2

if __name__ == '__main__':
    try:
        if len(sys.argv) > 1:
            for arg in sys.argv[1:]:
                fn = os.path.abspath(arg)
                main(fn)
                print "---"
        else:
            print "You need to drag your .pptx file(s) onto this one."
    finally:
        raw_input("Press any key to quit.")
        
        