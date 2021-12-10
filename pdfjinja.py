
from jinja2 import Environment
from PyPDF2.generic import NameObject, TextStringObject, createStringObject
from PyPDF2.pdf import ContentStream, PdfFileReader, PdfFileWriter
from PyPDF2.utils import b_


def map_text_objects(content_stream, mapper_func):
    """
    Iterates over each TextStringObject in the given ContentStream, and applies the
    given function to it.
    """
    # Note: we check all strings are TextStringObjects. ByteStringObjects
    # are strings where the byte->cstring encoding was unknown,
    # so their text here would be gibberish.
    def map_text_object(parent, index):
        obj = parent[index]
        if isinstance(obj, TextStringObject):
            replacement = mapper_func(obj)
            if isinstance(replacement, TextStringObject):
                parent[index] = replacement
            else:
                parent[index] = createStringObject(replacement)

    for operands, operator in content_stream.operations:
        if operator in (b_("Tj"), b_("'")):
            map_text_object(operands, 0)
        elif operator == b_('"'):
            map_text_object(operands, 2)
        elif operator == b_("TJ"):
            objs = operands[0]
            for i in range(len(objs)):
                map_text_object(objs, i)


class PDFTemplate(object):
    def __init__(self, template_path, jinja_env=None):
        self.reader = PdfFileReader(template_path)
        self.writer = PdfFileWriter()
        self.jinja_env = jinja_env or Environment()
        self.params = None

    def map_text(self, text):
        """
        Translates the given text using the template's parameters.
        """
        print(text)
        return self.jinja_env.from_string(text).render(self.params)

    def save(self, output_path):
        """
        Saves the rendered PDF to the given path.
        """
        with open(output_path, 'wb') as f:
            self.writer.write(f)

    def render(self, params):
        """
        Renders the PDF template with the given parameters.
        """
        self.params = params
        for i in range(self.reader.getNumPages()):
            page = self.reader.getPage(i)
            content = page.getContents().getObject()
            if not isinstance(content, ContentStream):
                content = ContentStream(content, self.reader)
            # self._process_stream_object(content, params)
            map_text_objects(content, self.map_text)
            page[NameObject("/Contents")] = content
            self.writer.addPage(page)

