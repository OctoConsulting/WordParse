import spacy

from document.sentence import Sentence


class Paragraph:
    def __init__(self, text, bold=False, colored=False,
                 font_names=None, font_sizes=None, format_string=None,
                 hyperlink=None, indent=None, level_number=None, ilvl=None,
                 num_id=None, style=None, section_num=None,
                 paragraph_id=None, line_spacing=None, left_margin=None,
                 right_margin=None, top_margin=None, bottom_margin=None,
                 page_height=None, page_width=None, html_text=None,
                 is_table=None, font_name_char_counts=None, font_size_char_counts=None):
        self.text = text
        self.bold = bold
        self.colored = colored
        self.font_names = font_names
        self.font_sizes = font_sizes
        self.format_string = format_string
        self.indent = indent
        self.level_number = level_number
        self.hyperlink = hyperlink
        self.ilvl = ilvl
        self.num_id = num_id
        self.style = style
        self.section_num = section_num
        self.paragraph_id = paragraph_id
        self.line_spacing = line_spacing
        self.left_margin = left_margin
        self.right_margin = right_margin
        self.top_margin = top_margin
        self.bottom_margin = bottom_margin
        self.page_height = page_height
        self.page_width = page_width
        self.html_text = html_text
        self.is_table = is_table
        self.font_name_char_counts = font_name_char_counts
        self.font_size_char_counts = font_size_char_counts

        self._is_chapter = False  # Set by Document.chapterize()

        self._sentences = list()

    def set_sentences(self, sentences):
        for span in sentences:
            assert isinstance(span, spacy.tokens.span.Span)
            self._sentences.append(Sentence(span))

    @property
    def is_chapter(self):
        return self._is_chapter

    def set_is_chapter(self, flag):
        self._is_chapter = flag

    def sentences(self):
        return self._sentences
