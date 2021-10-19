from collections import defaultdict

from spacy.lang.en import English

from document.paragraph import Paragraph


class Document:
    def __init__(self):
        self.paragraphs = list()

        self.nlp = English()
        sentencizer = self.nlp.create_pipe("sentencizer")
        self.nlp.add_pipe(sentencizer)

    def add_paragraph(self, p):
        doc = self.nlp(p.text)
        p.set_sentences(doc.sents)
        self.paragraphs.append(p)

        return p

    def chapterize(self):
        use_headings = False
        headings = defaultdict(int)
        for p in self.paragraphs:
            if p.style is not None and 'Heading' in p.style:
                headings[p.style] += 1

        if 'Heading1' in headings and headings['Heading1'] >= 3:
            for p in self.paragraphs:
                if p.style is not None and 'Heading' in p.style:
                    p.set_is_chapter(True)
            use_headings = True

        if not use_headings:
            for p in self.paragraphs:
                if p.format_string is not None and p.format_string == '%1.0':
                    p.set_is_chapter(True)

        # Risky
        for p in self.paragraphs:
            if p.style is not None and 'TOC' in p.style:
                continue
            if p.text.upper().startswith('APPENDIX'):
                p.set_is_chapter(True)

    def to_csv(self, fn):
        with open(fn, 'w') as f:
            for paragraph in self.paragraphs:
                for i, sentence in enumerate(paragraph.sentences()):
                    if i == 0 and paragraph.level_name is not None:
                        prefix = paragraph.level_name + ' '
                        # if paragraph.num_levels is not None:
                        #    prefix = '[%d] ' % paragraph.num_levels
                        # else:
                        #    prefix = ''
                    else:
                        prefix = ''
                    f.write(prefix + str(sentence.span) + '\n')
