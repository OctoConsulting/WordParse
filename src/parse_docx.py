from document.paragraph import Paragraph
from document.document import Document
import json
import os
import sys
import copy
import shutil

from collections import defaultdict
from lxml import etree as ET
import pandas as pd
import zipfile
import re

sys.path.append(os.path.join(os.path.dirname(__file__), '..'))


ns = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
ns_draw = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
ns_exp = "{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}"

ROMAN_LOWER = ['i', 'ii', 'iii', 'iv', 'v', 'vi', 'vii', 'viii', 'ix', 'x']
ROMAN_UPPER = ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X']

unicode_dict = {
    b'\xe2\x80\xa2':    b'-',   # black dot bullet
    b'\xef\x82\xa7':    b'-',   # big black dot bullet
    b'\xef\x82\xb7':    b'-',   # square bullet
    b'\xef\x80\xad':    b'-',   # single dash bullet
    b'\xc2\xad':        b'-',   # double dash bullet
    b'\xe2\x97\x8f':    b'-',   # text black dot bullet
    b'\xe2\x96\xaa':    b'-',   # text black square bullet
    b'\xe2\x97\x8b':    b'-',   # text black circle bullet
    b'\xe2\x80\x94':    b'-',   # em dash
    b'\xe2\x80\x93':    b'-',   # en dash
    b'\xe2\x80\x83':    b' ',   # em space
    b'\xe2\x80\x82':    b' ',   # en space
    b'\xe2\x80\x85':    b' ',   # 1/4 em space
    b'\xc2\xa0':        b' ',   # nonbreaking space
    b'\xc2\xa9':        b'COPYRIGHT',  # Copyright
    b'\xc2\xae':        b'REGISTERED',  # Registered
    b'\xe2\x84\xa2':    b'TRADEMARK',  # Trademark
    b'\xc2\xa7':        b'SECTION',    # Section
    b'\xc2\xb6':        b'PARAGRAPH',  # Paragraph
    b'\xe2\x80\xa6':    b'...',  # ellipsis
    b'\xe2\x80\x98':    b"'",   # single opening quote
    b'\xe2\x80\x99':    b"'",   # single closing quote
    b'\xe2\x80\x9c':    b'"',   # double opening quote
    b'\xe2\x80\x9d':    b'"',   # double closing quote
    b'\xe2\x96\xa1':    b'[]',   # check box
    b'\xe2\x81\x84':    b'/'    # forward slash
}


def parse_unicode(text_str):
    text_bytes = text_str.encode('utf-8')
    for key, value in unicode_dict.items():
        if key in text_bytes:
            text_bytes = text_bytes.replace(key, value)

    final_text_str = text_bytes.decode('utf-8')
    return final_text_str


def load_numbering(fn):
    tree = ET.parse(fn)

    # TODO: handle lvlRestart

    abstract = dict()
    num = dict()
    for abstract_node in tree.getiterator(ns + 'abstractNum'):
        abstract_num_id = abstract_node.attrib[ns + 'abstractNumId']
        abstract[abstract_num_id] = dict()
        found_level_node = False
        for level_node in abstract_node.getiterator(ns + 'lvl'):
            found_level_node = True
            d = dict()
            ilvl = level_node.attrib[ns + 'ilvl']
            if level_node.find(ns + 'start') is not None:
                d['start'] = int(level_node.find(
                    ns + 'start').attrib[ns + 'val']) - 1
            d['num_fmt'] = level_node.find(ns + 'numFmt').attrib[ns + 'val']
            d['level_text'] = parse_unicode(
                level_node.find(ns + 'lvlText').attrib[ns + 'val'])
            for indent_node in level_node.getiterator(ns + 'ind'):
                d['ind'] = indent_node.attrib.get(ns + 'left', None)
            abstract[abstract_num_id][ilvl] = d
        # If we don't find any numbering levels, look for a numStyleLink,
        #   which links the abstract to a style in styles.xml
        #   This abstract will use the numbering definition in that style
        if found_level_node is False:
            d = dict()
            num_style_link_node = abstract_node.find(ns + 'numStyleLink')
            if num_style_link_node is not None:
                d['num_style'] = num_style_link_node.attrib.get(
                    ns + 'val', None)
                abstract[abstract_num_id]['0'] = d

    for num_node in tree.getiterator(ns + 'num'):
        num_id = num_node.attrib[ns + 'numId']
        parent = num_node.find(ns + 'abstractNumId')
        ref_id = parent.attrib[ns + 'val']
        num[num_id] = copy.deepcopy(abstract[ref_id])
        for lvl_override_node in num_node.getiterator(ns + 'lvlOverride'):
            override_ilvl = lvl_override_node.attrib[ns + 'ilvl']
            if lvl_override_node.find(ns + 'startOverride') is not None:
                num[num_id][override_ilvl]['override_start'] = int(
                    lvl_override_node.find(ns + 'startOverride').attrib[ns + 'val'])

        for ilvl, lvl_dict in num[num_id].items():
            lvl_dict['abstract_ref_id'] = ref_id
        # TODO - add code to process <w:lvlOverride w:ilvl="X">

    return abstract, num


def load_theme(fn):
    tree = ET.parse(fn)

    theme_major_latin = None
    theme_minor_latin = None

    for font_scheme_node in tree.getiterator(ns_draw + 'fontScheme'):
        major_font_node = font_scheme_node.find(ns_draw + 'majorFont')
        if major_font_node is not None:
            major_font_latin = major_font_node.find(ns_draw + 'latin')
            if major_font_latin is not None:
                theme_major_latin = major_font_latin.attrib.get(
                    'typeface', None)

        minor_font_node = font_scheme_node.find(ns_draw + 'minorFont')
        if minor_font_node is not None:
            minor_font_latin = minor_font_node.find(ns_draw + 'latin')
            if minor_font_latin is not None:
                theme_minor_latin = minor_font_latin.attrib.get(
                    'typeface', None)

    return theme_major_latin, theme_minor_latin


def load_styles(fn, theme_major_latin, theme_minor_latin):
    tree = ET.parse(fn)

    default_font_name = None
    default_font_size = None
    for default_node in tree.getiterator(ns + 'docDefaults'):
        for rfonts_node in default_node.getiterator(ns + 'rFonts'):
            default_font_name = rfonts_node.attrib.get(ns + 'ascii', None)
            default_font_theme = rfonts_node.attrib.get(
                ns + 'asciiTheme', None)
            if default_font_name is None and default_font_theme is not None:
                if 'major' in default_font_theme:
                    default_font_name = theme_major_latin
                elif 'minor' in default_font_theme:
                    default_font_name = theme_minor_latin
        for sz_node in default_node.getiterator(ns + 'sz'):
            default_font_size = str(float(int(sz_node.attrib[ns + 'val']) / 2))
    if default_font_size is None:
        # if docDefaults does not specify a default font size, Word defaults to size 10 font
        default_font_size = '10'

    # So far I've only seen this with the SSA docx and only ilvl is there
    styles = dict()
    for style_node in tree.getiterator(ns + 'style'):
        style_name = style_node.attrib[ns + 'styleId']
        styles[style_name] = dict()
        for level_node in style_node.getiterator(ns + 'outlineLvl'):
            styles[style_name]['outline_lvl'] = level_node.attrib[ns + 'val']
        for level_node in style_node.getiterator(ns + 'ilvl'):
            styles[style_name]['ilvl'] = level_node.attrib[ns + 'val']
        for level_node in style_node.getiterator(ns + 'numId'):
            styles[style_name]['num_id'] = level_node.attrib[ns + 'val']
        for level_node in style_node.getiterator(ns + 'basedOn'):
            styles[style_name]['based_on'] = level_node.attrib[ns + 'val']
        rpr_node = style_node.find(ns + 'rPr')
        if rpr_node is not None:
            for level_node in rpr_node.getiterator(ns + 'b'):
                styles[style_name]['bold'] = True
            for level_node in rpr_node.getiterator(ns + 'color'):
                styles[style_name]['colored'] = True
            for level_node in rpr_node.getiterator(ns + 'rFonts'):
                styles[style_name]['ascii_font'] = level_node.attrib.get(
                    ns + 'ascii', None)
                styles[style_name]['ascii_font_theme'] = level_node.attrib.get(
                    ns + 'asciiTheme', None)
                if styles[style_name]['ascii_font'] is None and styles[style_name]['ascii_font_theme'] is not None:
                    if 'major' in styles[style_name]['ascii_font_theme']:
                        styles[style_name]['ascii_font'] = theme_major_latin
                    elif 'minor' in styles[style_name]['ascii_font_theme']:
                        styles[style_name]['ascii_font'] = theme_minor_latin

            for level_node in rpr_node.getiterator(ns + 'sz'):
                styles[style_name]['font_size'] = str(
                    float(int(level_node.attrib[ns + 'val']) / 2))
        for spacing_node in style_node.getiterator(ns + 'spacing'):
            styles[style_name]['spacing_before'] = spacing_node.attrib.get(
                ns + 'before', None)
            styles[style_name]['spacing_after'] = spacing_node.attrib.get(
                ns + 'after', None)
            styles[style_name]['line_rule'] = spacing_node.attrib.get(
                ns + 'lineRule', None)
            styles[style_name]['line_spacing'] = spacing_node.attrib.get(
                ns + 'line', None)

    # Once XML is loaded can follow basedOn links
    for style_name in styles:
        current_style = style_name
        while 'based_on' in styles[current_style]:
            based_on = styles[current_style]['based_on']
            # Inherit settings not overridden by this style
            for k in styles[based_on]:
                if k not in styles[style_name]:
                    styles[style_name][k] = styles[based_on][k]
            current_style = based_on

    return styles, default_font_name, default_font_size


def save_as_html(document, fn_out):
    html = list()
    html.append('<html>')
    html.append('<head><meta charset="UTF-8"></head>')
    html.append('<table border=1>')
    html.append('<tr><th>ParagraphID</th><th>SectionNumber</th><th>Style</th><th>Fonts</th><th>FontSizes</th><th>Color</th><th>Bold</th><th>HL</th><th>ID</th><th>Indent</th><th>Level</th><th>Format</th><th>Value</th><th>LineSpacing</th><th>Text</th><th>HTML_Text</th></tr>')

    for p in document.paragraphs:
        if p.bold:
            bold = 'Bold'
        else:
            bold = ''
        if p.colored:
            colored = 'Color'
        else:
            colored = ''
        if p.hyperlink:
            hyperlink = 'Link'
        else:
            hyperlink = ''

        if p.is_chapter:
            text = '<font color="red">' + p.text + '</font>'
        else:
            text = p.text

        font_names = ','.join(p.font_names)
        font_sizes = ','.join(p.font_sizes)

        html.append('<tr><td>' + z(p.paragraph_id) + '</td><td>' + z(p.section_num) + '</td><td>' + z(p.style) + '</td><td>' + font_names + '</td><td>' + font_sizes + '</td><td>' + colored + '</td><td>' + bold + '</td><td>' + hyperlink + '</td><td>' +
                    z(p.num_id) + '</td><td>' + z(p.indent) + '</td><td>' + z(p.ilvl) + '</td><td>' + z(p.format_string) + '</td><td>' + z(p.level_number) + '</td><td>' + p.line_spacing + '</td><td>' + text + '</td><td>' + p.html_text + '</td></tr>')

    html.append('</table>')
    html.append('</html>')

    with open(fn_out, 'w', encoding='utf8') as f:
        f.write('\n'.join(html))


def z(x):
    if x is None:
        return ''
    else:
        return str(x)


def get_creator_app(fn):
    app_tree = ET.parse(fn)
    creator_app = app_tree.findtext(ns_exp + "Application")
    return creator_app


def parse_docx(extract_dir, logger, debug=False):
    def parse(docx_file):
        extract_full_dir = '%s/%s' % (extract_dir,
                                      os.path.basename(docx_file).replace('.docx', '').strip())

        with zipfile.ZipFile(docx_file, 'r') as zip_ref:
            zip_ref.extractall(extract_full_dir)
            print(docx_file, " was successfully extracted to: ")
            print(extract_full_dir)

        creator_app = get_creator_app(extract_full_dir + '/docProps/app.xml')
        print("This docx was created by:", creator_app)

        # Some docx files don't have numbering.xml
        if os.path.exists(extract_full_dir + '/word/numbering.xml'):
            abstract, numbering = load_numbering(
                extract_full_dir + '/word/numbering.xml')
        else:
            abstract = dict()
            numbering = dict()

        numbering_initialized = dict()
        for num_id, ilvls in numbering.items():
            numbering_initialized[num_id] = dict()
            for ilvl, values in ilvls.items():
                numbering_initialized[num_id][ilvl] = False

        levels = dict()
        for abstract_num_id, ilvls in abstract.items():
            levels[abstract_num_id] = dict()
            for ilvl, values in ilvls.items():
                if abstract[abstract_num_id][ilvl].get('num_fmt', None) is not None:
                    if abstract[abstract_num_id][ilvl]['num_fmt'] != 'bullet':
                        levels[abstract_num_id][ilvl] = abstract[abstract_num_id][ilvl]['start']

        # print('abstract')
        #print(json.dumps(abstract, indent=2))
        # print('numbering')
        #print(json.dumps(numbering, indent=2))

        # Some docx files don't have theme1.xml
        if os.path.exists(extract_full_dir + '/word/theme/theme1.xml'):
            theme_font_major, theme_font_minor = load_theme(
                extract_full_dir + '/word/theme/theme1.xml')
        else:
            theme_font_major = None
            theme_font_minor = None

        styles, default_font_name, default_font_size = load_styles(
            extract_full_dir + '/word/styles.xml', theme_font_major, theme_font_minor)
        #print(json.dumps(styles, indent=2))

        normal_font_name = styles['Normal'].get('ascii_font', None)
        normal_font_size = styles['Normal'].get('font_size', None)

        tree = ET.parse(extract_full_dir + '/word/document.xml')

        # NOTE: may want to undo some of these removals later
        # Removing tables here
        #ET.strip_elements(tree, ns + 'tbl')
        # Removing drawings here
        ET.strip_elements(tree, ns + 'drawing')
        # Removing pict's here
        ET.strip_elements(tree, ns + 'pict')

        #root = tree.getroot()

        if debug == False:
            try:
                if extract_dir.startswith("extract"):
                    shutil.rmtree(extract_dir)
                    print("Removing extract directory:")
                    print(extract_dir)
                    print("Extract directory successfully deleted")
            except:
                print("WARNING: could not delete extract directory")

        BULLET = '-'

        section_info = {}
        new_section_num = 1

        paragraph_id = 0

        paragraphs = list()

        for paragraph in tree.getiterator(ns + 'p'):
            level_name = None  # 2.0, 2.1.1, a), etc.
            paragraph_style = None
            outline_lvl = None  # outlineLvl value
            ilvl = None  # 0, 1, 2, etc.
            num_id = None  # index into numbering.xml (2, 23, 25, etc)
            num_fmt = None  # how to format the number
            format_string = None  # %1.%2 and similar
            indent = None
            paragraph_bold = False
            hyperlink = False
            colored = False
            text = ''  # paragraph text
            section_num = new_section_num
            paragraph_id += 1
            first_page_num = None
            page_num_format = None
            pstyle_font_name = None
            pstyle_font_size = None
            pstyle_line_spacing = None
            pstyle_ilvl = None
            paragraph_line_spacing = None
            numbering_lvl = None

            if paragraph.getparent().tag == ns + 'tc':
                is_table = True
            else:
                is_table = False

            for text_node in paragraph.getiterator('*'):
                if text_node.tag == ns + 't':
                    text += text_node.text
                # replace tabs with space
                elif text_node.tag == ns + 'tab':
                    text += ' '
                # replace line breaks with newline character
                elif text_node.tag == ns + 'br':
                    text += '\n'

            text = parse_unicode(text)

            # Regular expressions to find numbering

            dash_re = re.compile('^ *-')
            dash_re_search = dash_re.match(text)
            if dash_re_search is not None:
                numbering_lvl = dash_re_search.group()
                text = re.sub(dash_re, '', text).lstrip()

            numbering_lvl_re = re.compile('^ *[A-Za-z]?[\-\.]?(\d+[.]?)+\s+')
            numbering_lvl_search = numbering_lvl_re.match(text)
            if numbering_lvl_search is not None:
                try:
                    numbering_lvl_float = float(numbering_lvl_search.group())
                    if numbering_lvl_float >= 100:
                        numbering_lvl_final = None
                    else:
                        numbering_lvl_final = numbering_lvl_search.group()
                except:
                    numbering_lvl_final = numbering_lvl_search.group()
                finally:
                    if numbering_lvl is None:
                        numbering_lvl = numbering_lvl_final

            section_header_re = re.compile('^ *SECTION \w+', re.IGNORECASE)
            section_header_search = section_header_re.match(text)
            section_re = re.compile('section', re.IGNORECASE)
            if section_header_search is not None:
                section_header = section_header_search.group()
                if numbering_lvl is None:
                    numbering_lvl = re.sub(section_re, '', section_header)

            paren_numbering_re = re.compile('^ *[(]?\w+[)]')
            paren_numbering_search = paren_numbering_re.match(text)
            if paren_numbering_search is not None:
                if numbering_lvl is None:
                    numbering_lvl = paren_numbering_search.group()

            roman_numeral_re = re.compile('^ *([IVXivx]+[.]?)+[.]')
            roman_numeral_search = roman_numeral_re.match(text)
            if roman_numeral_search is not None:
                if numbering_lvl is None:
                    numbering_lvl = roman_numeral_search.group()

            if numbering_lvl is not None:
                numbering_lvl = numbering_lvl.strip()

            for hyperlink_node in paragraph.getiterator(ns + 'hyperlink'):
                hyperlink = True

            ppr_node = paragraph.find(ns + 'pPr')
            if ppr_node is not None:

                framepr_node = ppr_node.find(ns + 'framePr')
                if framepr_node is not None:
                    # Possibly undo this later
                    # Here we will skip any paragraph that is part of a text frame
                    continue

                pstyle_node = ppr_node.find(ns + 'pStyle')
                if pstyle_node is not None:
                    paragraph_style = pstyle_node.attrib[ns + 'val']

                outline_node = ppr_node.find(ns + 'outlineLvl')
                if outline_node is not None:
                    outline_lvl = outline_node.attrib[ns + 'val']

                numpr_node = ppr_node.find(ns + 'numPr')
                if numpr_node is not None:

                    ilvl_node = numpr_node.find(ns + 'ilvl')
                    if ilvl_node is not None:
                        ilvl = ilvl_node.attrib[ns + 'val']

                    #lvl_format = None
                    numid_node = numpr_node.find(ns + 'numId')
                    if numid_node is not None:
                        num_id = numid_node.attrib[ns + 'val']

            for sectpr_node in paragraph.getiterator(ns + 'sectPr'):
                new_section_num = section_num + 1
                for pgmar_node in sectpr_node.getiterator(ns + 'pgMar'):
                    left_margin = float(
                        pgmar_node.attrib.get(ns + 'left', -1)) / 1440
                    right_margin = float(
                        pgmar_node.attrib.get(ns + 'right', -1)) / 1440
                    top_margin = float(
                        pgmar_node.attrib.get(ns + 'top', -1)) / 1440
                    bottom_margin = float(
                        pgmar_node.attrib.get(ns + 'bottom', -1)) / 1440
                for pgnumtype_node in sectpr_node.getiterator(ns + 'pgNumType'):
                    first_page_num = pgnumtype_node.attrib.get(
                        ns + 'start', -1)
                    page_num_format = pgnumtype_node.attrib.get(
                        ns + 'fmt', None)
                for pgsz_node in sectpr_node.getiterator(ns + 'pgSz'):
                    page_height = float(
                        pgsz_node.attrib.get(ns + 'h', -1)) / 1440
                    page_width = float(
                        pgsz_node.attrib.get(ns + 'w', -1)) / 1440
                section_info[section_num] = {'section_num': section_num,
                                             'start_page': first_page_num,
                                             'page_num_format': page_num_format,
                                             'left_margin': left_margin,
                                             'right_margin': right_margin,
                                             'top_margin': top_margin,
                                             'bottom_margin': bottom_margin,
                                             'page_height': page_height,
                                             'page_width': page_width}

            # for outline_node in paragraph.getiterator(ns + 'pStyle'):
                #paragraph_style = outline_node.attrib[ns + 'val']

            if paragraph_style is not None:

                pstyle_font_name = styles[paragraph_style].get(
                    'ascii_font', None)
                pstyle_font_size = styles[paragraph_style].get(
                    'font_size', None)
                pstyle_line_spacing = styles[paragraph_style].get(
                    'line_spacing', None)

                # Don't assign bold here, it would only be for paragraph mark
                if paragraph_bold is False:
                    if 'bold' in styles[paragraph_style]:
                        paragraph_bold = styles[paragraph_style]['bold']

                if colored is False:
                    if 'colored' in styles[paragraph_style]:
                        colored = styles[paragraph_style]['colored']

                # Use outline levels first (Heading1, etc.)

                pstyle_num_id = styles[paragraph_style].get('num_id', None)
                if num_id is None:
                    num_id = pstyle_num_id

                pstyle_ilvl = styles[paragraph_style].get('ilvl', None)

            if ilvl is None:
                if pstyle_ilvl is not None:
                    ilvl = pstyle_ilvl
                else:
                    ilvl = '0'  # default?

            # TODO - understand this better
            # A value of 0 for the @val attribute shall never be used to point to
            #  a numbering definition instance, and shall instead only be used
            #  to designate the removal of numbering properties at a particular level
            #  in the style hierarchy (typically via direct formatting).
            if num_id == "0":
                num_id = None

            for spacing_node in paragraph.getiterator(ns + 'spacing'):
                paragraph_line_spacing = spacing_node.attrib.get(
                    ns + 'line', None)

            if paragraph_line_spacing is not None:
                line_spacing_num = paragraph_line_spacing
            elif pstyle_line_spacing is not None:
                line_spacing_num = pstyle_line_spacing
            else:
                line_spacing_num = "240"

            line_spacing = str(float(line_spacing_num) / 240)

            """
            if line_spacing_num == "240":
                line_spacing = 'S'
            elif line_spacing_num == "360":
                line_spacing = 'H'
            elif line_spacing_num == "480":
                line_spacing = 'D'
            else:
                line_spacing = 'O'
            """

            # TODO: Colin to add font name(s) and font size(s)
            # Rollup from the runs within this paragraph
            font_names = set()
            font_sizes = set()
            html_text = ''
            font_name_char_counts = defaultdict(int)
            font_size_char_counts = defaultdict(int)

            for run_node in paragraph.getiterator(ns + 'r'):
                font_name = None
                font_size = None
                run_text = ''
                bold = paragraph_bold

                for text_node in run_node.getiterator('*'):
                    if text_node.tag == ns + 't':
                        run_text += text_node.text
                    elif text_node.tag == ns + 'tab':
                        run_text += ' '
                    elif text_node.tag == ns + 'br':
                        run_text += '<br/>'

                # Get bold from run if available
                for bold_node in run_node.getiterator(ns + 'b'):
                    bold_val = bold_node.attrib.get(ns + 'val', None)
                    if bold_val is None or bold_val == "1":
                        bold = True

                # Get ascii font from run if available
                for rfonts_node in run_node.getiterator(ns + 'rFonts'):
                    font_name = rfonts_node.attrib.get(ns + 'ascii', None)
                    font_theme = rfonts_node.attrib.get(
                        ns + 'asciiTheme', None)
                    if font_name is None and font_theme is not None:
                        if 'major' in font_theme:
                            font_name = theme_font_major
                        elif 'minor' in font_theme:
                            font_name = theme_font_minor

                # Else get from run style if available
                if font_name is None:
                    for style_node in run_node.getiterator(ns + 'rStyle'):
                        run_style = style_node.attrib[ns + 'val']
                        if styles[run_style].get('ascii_font', None) is not None:
                            font_name = styles[run_style]['ascii_font']

                if font_name is None:
                    if pstyle_font_name is not None:
                        font_name = pstyle_font_name
                    elif normal_font_name is not None:
                        font_name = normal_font_name
                    elif default_font_name is not None:
                        font_name = default_font_name

                if font_name is not None:
                    if run_text.isspace() is False and len(run_text) > 0:
                        font_names.add(font_name)
                        font_name_char_counts[font_name] += len(run_text)
                else:
                    print("Found a run with none font in paragraph:", paragraph_id)
                    print(run_text)

                # Get font size from run if available
                for sz_node in run_node.getiterator(ns + 'sz'):
                    font_size = str(float(int(sz_node.attrib[ns + 'val']) / 2))

                # Else get from run style if available
                if font_size is None:
                    for style_node in run_node.getiterator(ns + 'rStyle'):
                        run_style = style_node.attrib[ns + 'val']
                        if styles[run_style].get('font_size', None) is not None:
                            font_size = styles[run_style]['font_size']

                if font_size is None:
                    if pstyle_font_size is not None:
                        font_size = pstyle_font_size
                    elif normal_font_size is not None:
                        font_size = normal_font_size
                    elif default_font_size is not None:
                        font_size = default_font_size

                if font_size is not None:
                    if run_text.isspace() is False and len(run_text) > 0:
                        font_sizes.add(font_size)
                        font_size_char_counts[font_size] += len(run_text)
                else:
                    print("Found a run with no font size in paragraph:", paragraph_id)
                    print(run_text)

                if bold:
                    run_text = '<b>%s</b>' % run_text

                html_text += '<span style="font-size:%spt;font-family:%s">%s</span>' % (
                    font_size, "'" + font_name + "'", run_text)

            # TODO: Calculate these within this paragraph

            lvl_format = None
            if ilvl is not None and num_id is not None:

                #lvl_format = None
                if num_id in numbering:
                    # If there is a numStyleLink reference,
                    #  use the num_id found within that style
                    if numbering[num_id][ilvl].get('num_style', None) is not None:
                        num_style = numbering[num_id][ilvl]['num_style']
                        num_id = styles[num_style]['num_id']
                    lvl_format = numbering[num_id][ilvl]
                else:
                    print("num_id of", num_id, " not found in numbering!")

                format_string = lvl_format['level_text']
                num_fmt = lvl_format['num_fmt']

                if 'ind' in lvl_format:
                    indent = lvl_format['ind']

                #print('  lvl_format:', lvl_format)

                abstract_referenced = lvl_format['abstract_ref_id']

                if num_fmt == 'bullet':
                    level_name = BULLET
                else:
                    # Handle Level override here
                    if lvl_format.get('override_start', None) is not None and numbering_initialized[num_id][ilvl] == False:
                        numbering_initialized[num_id][ilvl] = True
                        levels[abstract_referenced][ilvl] = lvl_format['override_start']
                    else:
                        # Increment level number
                        for i in range(int(ilvl)):
                            if abstract[abstract_referenced][str(i)]['num_fmt'] != 'bullet':
                                if levels[abstract_referenced][str(i)] == abstract[abstract_referenced][str(i)]['start']:
                                    levels[abstract_referenced][str(i)] += 1

                        levels[abstract_referenced][ilvl] += 1

                    # Restart numbering below
                    for k in levels[abstract_referenced]:
                        if int(k) > int(ilvl):
                            levels[abstract_referenced][k] = abstract[abstract_referenced][k]['start']

                    # Get formatting pattern
                    level_name = lvl_format['level_text']

                    # Format number
                    for i in range(int(ilvl) + 1):
                        s = '%s%d' % ('%', i + 1)
                        if s in level_name:
                            if num_fmt == 'decimal':
                                new_value = levels[abstract_referenced][str(i)]
                            elif num_fmt == 'lowerLetter':
                                new_value = chr(
                                    ord('a') + (levels[abstract_referenced][str(i)] - 1))
                            elif num_fmt == 'upperLetter':
                                new_value = chr(
                                    ord('A') + (levels[abstract_referenced][str(i)] - 1))
                            elif num_fmt == 'none':
                                new_value = ''
                            elif num_fmt == 'lowerRoman':
                                val = levels[abstract_referenced][str(i)] - 1
                                if val < len(ROMAN_LOWER):
                                    new_value = ROMAN_LOWER[val]
                                else:
                                    new_value = 'xxx'  # TODO: Handle higher values
                            elif num_fmt == 'upperRoman':
                                val = levels[abstract_referenced][str(i)] - 1
                                if val < len(ROMAN_UPPER):
                                    new_value = ROMAN_UPPER[val]
                                else:
                                    new_value = 'xxx'  # TODO: Handle higher values
                            else:
                                raise RuntimeError(
                                    'Unsupported level format %s' % num_fmt)
                            level_name = level_name.replace(s, str(new_value))

                    # Check if we could not format the level
                    if '%' in level_name:
                        print("Could not format the level!!")
                        print(ilvl, num_id, lvl_format)

            if level_name is None and numbering_lvl is not None:
                level_name = numbering_lvl

            # if len(text.strip()) > 0 and level_name not in [None, BULLET]:
            if len(text.strip()) == 0:
                continue

            # Prepare font char counts for output to csv as dictionary
            font_name_char_counts = json.dumps(font_name_char_counts)
            font_size_char_counts = json.dumps(font_size_char_counts)

            # Convert to dict for loading into DataFrame
            p = dict()
            p['bold'] = bold
            p['colored'] = colored
            if format_string is None:
                format_string = ''
            p['format_string'] = format_string
            p['level_number'] = level_name
            p['hyperlink'] = hyperlink
            p['ilvl'] = ilvl
            p['indent'] = indent
            p['num_id'] = num_id
            p['style'] = paragraph_style
            p['text'] = text
            p['font_names'] = font_names
            p['font_sizes'] = font_sizes
            p['section_num'] = section_num
            p['paragraph_id'] = paragraph_id
            p['line_spacing'] = line_spacing
            p['html_text'] = html_text
            p['is_table'] = is_table
            p['font_name_char_counts'] = font_name_char_counts
            p['font_size_char_counts'] = font_size_char_counts

            if len(p['font_names']) == 0:
                print("Found None Font!!!")
                print(p['font_names'])
                print(p)
                p['font_names'].add('None')

            if len(p['font_sizes']) == 0:
                print("Found None Font Size!!!")
                print(p['font_sizes'])
                print(p)
                p['font_sizes'].add('None')

            paragraphs.append(p)

        # Get information from final section, sectPr child node of body
        for body_node in tree.getiterator(ns + 'body'):
            sectpr_node = body_node.find(ns + 'sectPr')
            if sectpr_node is not None:
                if section_info:
                    section_num = max(section_info.keys()) + 1
                else:
                    section_num = 1
                for pgmar_node in sectpr_node.getiterator(ns + 'pgMar'):
                    left_margin = float(
                        pgmar_node.attrib.get(ns + 'left', -1)) / 1440
                    right_margin = float(
                        pgmar_node.attrib.get(ns + 'right', -1)) / 1440
                    top_margin = float(
                        pgmar_node.attrib.get(ns + 'top', -1)) / 1440
                    bottom_margin = float(
                        pgmar_node.attrib.get(ns + 'bottom', -1)) / 1440
                for pgnumtype_node in sectpr_node.getiterator(ns + 'pgNumType'):
                    first_page_num = pgnumtype_node.attrib.get(
                        ns + 'start', -1)
                    page_num_format = pgnumtype_node.attrib.get(
                        ns + 'fmt', None)
                for pgsz_node in sectpr_node.getiterator(ns + 'pgSz'):
                    page_height = float(
                        pgsz_node.attrib.get(ns + 'h', -1)) / 1440
                    page_width = float(
                        pgsz_node.attrib.get(ns + 'w', -1)) / 1440
                section_info[section_num] = {'section_num': section_num,
                                             'start_page': first_page_num,
                                             'page_num_format': page_num_format,
                                             'left_margin': left_margin,
                                             'right_margin': right_margin,
                                             'top_margin': top_margin,
                                             'bottom_margin': bottom_margin,
                                             'page_height': page_height,
                                             'page_width': page_width}

        # If we can't find any paragraphs in the document, break out of parse_docx
        if len(paragraphs) == 0:
            logger.warn('Unsupported docx with filename: %s' % docx_file)
            return

        # Now that we have the entire document we can break it up into chapters
        # based on various heuristics/rules
        document = Document()
        for p in paragraphs:
            p['left_margin'] = section_info[p['section_num']]['left_margin']
            p['right_margin'] = section_info[p['section_num']]['right_margin']
            p['top_margin'] = section_info[p['section_num']]['top_margin']
            p['bottom_margin'] = section_info[p['section_num']]['bottom_margin']
            p['page_height'] = section_info[p['section_num']]['page_height']
            p['page_width'] = section_info[p['section_num']]['page_width']
            paragraph = Paragraph(p['text'], bold=p['bold'], colored=p['colored'],
                                  font_names=p['font_names'],
                                  font_sizes=p['font_sizes'],
                                  format_string=p['format_string'],
                                  level_number=p['level_number'],
                                  hyperlink=p['hyperlink'],
                                  ilvl=p['ilvl'], indent=p['indent'],
                                  num_id=p['num_id'], style=p['style'],
                                  section_num=p['section_num'],
                                  paragraph_id=p['paragraph_id'],
                                  line_spacing=p['line_spacing'],
                                  left_margin=p['left_margin'],
                                  right_margin=p['right_margin'],
                                  top_margin=p['top_margin'],
                                  bottom_margin=['bottom_margin'],
                                  page_height=['page_height'],
                                  page_width=['page_width'],
                                  html_text=p['html_text'],
                                  is_table=p['is_table'])
            document.add_paragraph(paragraph)

        document.chapterize()

        # TODO: Better way of doing this (don't rely on index)
        for i, p in enumerate(document.paragraphs):
            paragraphs[i]['is_chapter'] = p.is_chapter

        # Save as DataFrame
        df = pd.DataFrame(paragraphs)
        df['font_names'] = df['font_names'].apply(lambda x: ','.join(x))
        df['font_sizes'] = df['font_sizes'].apply(lambda x: ','.join(x))
        # df.to_csv(fn_out, index=False)
        # print('Wrote', fn_out)

        return df
    return parse
