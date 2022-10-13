from docxtpl import DocxTemplate


class Word_Helper():
    def __init__(self):
        pass
        
    def write_to_worfd_file(self, context: dict, path: str, name:str="Laborauswertung.docx"):
        if ".doc" or ".docx" not in name:
            name = name + ".docx"
        self.doc = DocxTemplate(path)
        self.doc.render(context)
        self.doc.save(name)
        return True

if __name__ == "__main__":
    w = Word_Helper()
    w.write_to_worfd_file({"projekt_nr": "Florian Test"}, r"items\vorlagen\Bericht Vorlage.docx", name="FLOIAN_TEST")