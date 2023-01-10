from docxtpl import DocxTemplate
import win32com.client as win32



class Word_Helper():
    def __init__(self):
        pass
        
    def write_to_word_file(self, context: dict, vorlage_path: str, name: str="Laborauswertung.docx"):
        try:
            self.doc = DocxTemplate(vorlage_path)
            self.doc.render(context)
            self.doc.save(name)
        except Exception as ex:
            print("FEHLER BEI WORD: ", str(ex))
            return False
        return True

    def open_word(self, vorlage_path: str):
        word = win32.gencache.EnsureDispatch('Word.Application')
        word.Visible = True
        word.Documents.Open(vorlage_path)

if __name__ == "__main__":
    w = Word_Helper()
    w.write_to_word_file({"projekt_nr": "Florian Test"}, r"vorlagen\Bericht Vorlage.docx", name="FLOIAN_TEST.docx")