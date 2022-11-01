from docxtpl import DocxTemplate
import win32com.client as win32



class Word_Helper():
    def __init__(self):
        pass
        
    def write_to_word_file(self, context: dict, path: str, name:str="Laborauswertung.docx"):
        self.doc = DocxTemplate(path)
        self.doc.render(context)
        self.doc.save(name)
        return True

    def open_word(self, path: str):
        print(path)
        word = win32.gencache.EnsureDispatch('Word.Application')
        word.Visible = True
        doc = word.Documents.Open(path)

if __name__ == "__main__":
    w = Word_Helper()
    w.write_to_word_file({"projekt_nr": "Florian Test"}, r"items\vorlagen\Bericht Vorlage.docx", name="FLOIAN_TEST")