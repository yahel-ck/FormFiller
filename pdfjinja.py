from jinja2 import Environment
from PyPDF2.generic import ByteStringObject, NameObject, TextStringObject, createStringObject
from PyPDF2.pdf import ContentStream, PdfFileReader, PdfFileWriter
from PyPDF2.utils import b_, str_, ord_
import re


CMAP_ENTRY_RE = re.compile("\\s*<([0-9a-fA-F]+)>\\s+<([0-9a-fA-F]+)>\\s*")
CMAP_RE = re.compile(
    "\nbegincmap\n(?:.*?\n)?[0-9]* beginbfchar\n(.*?)\nendbfchar\n(?:.*?\n)?"
    "endcmap\n", 
    re.DOTALL)


def map_text_objects(page, content_stream, mapper_func):
    """
    Iterates over each string object in the given ContentStream, and applies the
    given function to it.

    Note: Works only for TextStringObjects or ByteStringObjects with cmap
    defined. ByteStringObjects are strings where the byte->cstring encoding was 
    unknown, but if they have a cmap field we can translate them to unicode,
    otherwise their text here would be gibberish.
    """
    def map_text_object(parent, index):
        obj = translate(parent[index])
        if obj is None:
            return
        replacement = mapper_func(obj)
        if isinstance(replacement, TextStringObject):
            parent[index] = replacement
        else:
            parent[index] = createStringObject(replacement)

    c_map = None
    c_maps = {}
    # Concatenate TextStringObjects and try to translate ByteStringObjects
    # when we have a CMap, when we don't, then byte->string encoding is unknown,
    # so adding them to the text here would be gibberish.

    def translate(text):
        if isinstance(text, TextStringObject):
            return text
        if isinstance(text, ByteStringObject):
            newText = "".join(c_map.get(ord_(c), '?') for c in text)
            return newText
        return None

    for operands, operator in content_stream.operations:
        if operator == b_("Tf"):
            try:
                font = operands[0]
                c_map = c_maps.get(font)
                if (c_map == None):
                    c_map = parse_cmap(get_cmap_data(page, font))
                    c_maps[font] = c_map
            except KeyError:
                c_map = None
        elif operator in (b_("Tj"), b_("'")):
            map_text_object(operands, 0)
        elif operator == b_('"'):
            map_text_object(operands, 2)
        elif operator == b_("TJ"):
            objs = operands[0]
            for i in range(len(objs)):
                map_text_object(objs, i)


def get_cmap_data(page, font):
    return str_(page["/Resources"]["/Font"][font]["/ToUnicode"].getData())


def parse_cmap(cstr):
    cmap_def = CMAP_RE.search(cstr)
    if cmap_def == None:
        return None
    cmap = {}
    for entry in cmap_def.group(1).split("\n"):
        entry_def = CMAP_ENTRY_RE.match(entry)
        if entry_def is not None:
            cmap[int(entry_def.group(1), base=16)] = \
                chr(int(entry_def.group(2), base=16))
    return cmap


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
        new_text = self.jinja_env.from_string(text).render(self.params)
        print(' '.join(str(ord(c)) for c in text))
        print(' '.join(str(ord(c)) for c in new_text))
        print()
        return new_text

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
            map_text_objects(page, content, self.map_text)
            page[NameObject("/Contents")] = content
            self.writer.addPage(page)
