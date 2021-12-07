
from PyPDF2.pdf import ContentStream, PdfFileReader, PdfFileWriter
from PyPDF2.generic import DecodedStreamObject, EncodedStreamObject, TextStringObject, NameObject
from PyPDF2.utils import b_
from jinja2 import Environment


class PDFTemplate(object):
    def __init__(self, template_path, jinja_env=None):
        self.reader = PdfFileReader(template_path)
        self.writer = PdfFileWriter()
        self.jinja_env = jinja_env or Environment()

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
        for i in range(self.reader.getNumPages()):
            page = self.reader.getPage(i)
            content = page.getContents()
            if isinstance(content, DecodedStreamObject) or isinstance(content, EncodedStreamObject):
                self._process_stream_object(content, params)
            else:
                for obj in content:
                    self._process_stream_object(obj, params)
            page[NameObject("/Contents")] = content.decodedSelf
            self.writer.addPage(page)

    def get_text_objects(self, page):
        """
        Returns a generator that iterates over each TextStringObject in the
        given page.
        """
        content = page["/Contents"].getObject()
        if not isinstance(content, ContentStream):
            content = ContentStream(content, self.pdf)
        # Note: we check all strings are TextStringObjects. ByteStringObjects
        # are strings where the byte->cstring encoding was unknown, so adding
        # them to the text here would be gibberish.
        for operands, operator in content.operations:
            if operator in (b_("Tj"), b_("'")):
                obj = operands[0]
                if isinstance(obj, TextStringObject):
                    yield obj
            elif operator == b_('"'):
                obj = operands[2]
                if isinstance(obj, TextStringObject):
                    yield obj
            elif operator == b_("TJ"):
                for obj in operands[0]:
                    if isinstance(obj, TextStringObject):
                        yield obj

    def _process_stream_object(self, object, params):
        data = object.getData()
        try:
            encoding = 'utf-8'
            decoded_data = data.decode(encoding)
        except UnicodeDecodeError:
            encoding = 'unicode_escape'
            decoded_data = data.decode(encoding)

        new_data = self._translate_content(decoded_data, params)
        encoded_data = new_data.encode(encoding)

        if object.decodedSelf is not None:
            object.decodedSelf.setData(encoded_data)
            return object.decodedSelf
        else:
            object.setData(encoded_data)
            return object

    def _translate_content(self, content, params):
        return self.jinja_env.from_string(content).render(params)
