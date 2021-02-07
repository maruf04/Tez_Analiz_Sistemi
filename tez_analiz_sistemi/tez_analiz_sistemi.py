from tabula import read_pdf
from os import listdir, path, mkdir, system, remove
from docx2pdf import convert
from time import sleep
from PyPDF2 import PdfFileReader, generic
from re import findall


#üstteki kütüphaneler:
#tabula---->tablo çekmek için 
#os terminal renklendirme ve temizleme için
#docx2pdf docx dosyasını pdf e dönüştürmek için
# time sistemi bekletmek için
#PyPDF2 pdf dokumanı okumak ve içeriklerini analiz edebilmek için

class main:
    #constructor fonksiyonu
    def __init__(self):
        super().__init__()
        system("cls")
        system("color d")
        self.docx = False
        self.dic = {"BEYAN": 0, "ÖNSÖZ": 0, "İÇİNDEKİLER": 0, "ÖZET": 0, "ABSTRACT": 0, "ŞEKİLLER LİSTESİ": 0, "TABLOLAR LİSTESİ": 0,
                    "EKLER LİSTESİ": 0, "SİMGELER VE KISALTMALAR": 0, "GİRİŞ": 0, "SONUÇLAR": 0, "ÖNERİLER": 0, "KAYNAKLAR": 0, "EKLER": 0, "ÖZGEÇMİŞ": 0}
        self.first()
    #terminalin sayfa arayüzünü gösterir
    def first(self):
        self.page_start = """
                          WELCOME
                    TEZ ANALYSİS SYSTEM
                SOFTWARE ENGINEER MARUF AKAN
                        yes--->[E/e]
                        no---->[H/h]
        """
        self.page_end = """
                          GOOD BYE
                    TEZ ANALYSİS SYSTEM
                SOFTWARE ENGINEER MARUF AKAN
        """

        print(self.page_start)
        print("START---->KEY PRESS...")
        input()
        print("please waiting...")
        sleep(1)
        print("starting...")
        sleep(1)

        self.start()
        
    #sistemin çalışmasına devam etmek için evet veya hayır ile cevap vermemizi bekler
    def start(self):
        system("cls")
        self.answer = input(
            "Do you want your files to be scanned ?\n   [E/e]-[H/h]")
        if self.answer.lower() == "e":
            system("color a")
            print("control...")
            sleep(2)
            self.control()
        elif self.answer.lower() == "h":
            print("good bye ☻")
            sleep(2)
            exit()
        else:
            print(self.answer)
            system("color c")
            print("please try again...")
            sleep(2)
            self.start()
    def control(self):
        system("cls")
        self.files = listdir(".")
        
        self.file_ = ""
        for file in self.files:
            if file.endswith(".pdf") or file.endswith(".doc"):
                self.file_ = file
                break
        if self.file_ == "":
            system("color c")
            print("You do not have files in the system")
            sleep(1)
            print("system exit\n good bye")
            sleep(2)
            exit()
        if self.file_.endswith(".doc"):
            system("color a")
            print(self.file_+" convert to pdf ...")
            sleep(2)
            convert(self.file_)
            self.docx = True
            self.control()
        else:
            print(self.file_+" analysis starting...")
            sleep(2)
            self.distribute()

    def distribute(self):
        self.analysis = open("analysis.txt","w", encoding="utf-8")
        system("cls")
        self.about()
        sleep(3)
        system("cls")

    def about(self):
        print("about write...")
        with open(self.file_, "rb") as reader_:
            reader = PdfFileReader(reader_)
            info = reader.getDocumentInfo()
            self.analysis.write("--------------about----------------\n")
            self.analysis.write(f"author: "+info.author+"\n")
            self.analysis.write(f"program: "+info.creator+"\n")
            self.analysis.write(f"producer: "+info.producer+"\n")
            self.analysis.write(f"pages: "+str(reader.getNumPages())+"\n")
        reader_.close()
        self.table()

    def table(self):
        print("tables extract....")
        sleep(3)
        tables = read_pdf(self.file_, pages="all")
        file_name = "tables"
        if not path.isdir(file_name):
            mkdir(file_name)
        for i, table in enumerate(tables, start=1):
            system("cls")
            table.to_excel(path.join(file_name, "table_" +
                                     str(i) + ".xlsx"), index=False)
        self.font()

    def font(self):
        system("cls")
        print("font analysis...")
        sleep(1)
        pdf = PdfFileReader(self.file_)
        fonts = set()
        embedded = set()
        for page in pdf.pages:
            obj = page.getObject()
            if type(obj) == generic.ArrayObject:
                for i in obj:
                    if hasattr(i, 'keys'):
                        f, e = self.font_write(i, fonts, embedded)
                        fonts = fonts.union(f)
                        embedded = embedded.union(e)
            else:
                f, e = self.font_write(obj['/Resources'], fonts, embedded)
                fonts = fonts.union(f)
                embedded = embedded.union(e)
        unembedded = fonts - embedded
        self.analysis.write('--------------Font List-------------\n')
        for font in fonts:
            font = font.replace("/ABCDEE+", "")
            font = font.replace("/", "")
            self.analysis.write(font + "\n")

        self.header()

    def font_write(self, obj, fnt, emb):
        if not hasattr(obj, 'keys'):
            return None, None
        fontkeys = {'/FontFile', '/FontFile2', '/FontFile3'}
        if '/BaseFont' in obj:
            fnt.add(obj['/BaseFont'])
        if '/FontName' in obj:
            if [x for x in fontkeys if x in obj]:
                emb.add(obj['/FontName'])
        for k in obj.keys():
            self.font_write(obj[k], fnt, emb)

        return fnt, emb

    def header(self):
        system("cls")
        print("header analysis")
        with open(self.file_, "rb") as data:
            data_ = PdfFileReader(data)
            pages = data_.getNumPages()
            i = 0
            while i < pages:
                num = data_.getPage(i)
                num = num.extractText()
                num = num.replace("\n", "").replace(" ", "")
                self.headercontrol(num)
                i += 1
        data.close()
        self.analysis.write(
            "-----------------header analysis----------------------\n")
        for i in self.dic:
            if self.dic[i] == 0:
                self.analysis.write(str(i)+"---->tez dosyasında bulunamadı! (Hata sebebi punto farklılığı veya dosyada bulundurmamasıdır)\n")
        self.extractUrl()

    def headercontrol(self, file):
        for i in self.dic:
            if file.find(i) >= 0:
                self.dic[i] = 1
                print(i)

    def extractUrl(self):
        system("cls")
        system("color e")
        print("Url Analysis")
        self.analysis.write("--------------URL--------------\n")
        with open(self.file_, 'rb') as objects:
            read = PdfFileReader(objects)
            for page in range(read.getNumPages()):
                text = read.getPage(page).extractText()
                regex = r"(https?://\S+)"
                url =findall(regex, text)
                for i in url:
                    print(i)
                    self.analysis.write(i + "\n")
        objects.close()
        sleep(3)
        self.finish()

    def finish(self):
        system("cls")
        system("color d")
        if self.docx:
            remove(self.file_)

        print(self.page_end)
        sleep(4)
        print("good bye")

        self.analysis.close()
        system("cls")
        system("start analysis.txt")
        exit()


if __name__ == '__main__':
    main()
























