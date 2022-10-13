from docxtpl import DocxTemplate


class Word_Helper():
    def __init__(self, path):
        self.path = path
        self.doc = DocxTemplate(path)




    def write_to_worfd_file(self, context, name="Laborauswertung.docx"):
        self.doc.render(context)
        self.doc.save(name)