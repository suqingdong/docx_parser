from pathlib import Path

import docx
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.oxml.text.run import CT_R

from docx.table import Table
from docx.text.paragraph import Paragraph

from docx_parser import util


class DocumentParser(object):
    def __init__(self, filename, image_as='base64', image_type='png', media_dir='media', **kwargs):
        self.image_as = image_as
        self.media_dir = Path(media_dir)
        self.image_type = image_type
        self.document = docx.Document(filename)

    def parse(self):
        for n, element in enumerate(self.document.element.body.iterchildren()):
            # print(n, element)
            if isinstance(element, CT_P):
                yield from self.parse_paragraph(Paragraph(element, self.document))
            elif isinstance(element, CT_Tbl):
                yield self.parse_table(Table(element, self.document))

    def parse_paragraph(self, paragraph):
        """parse paragraph
        """
        if paragraph._element.xpath('.//a:graphic|.//w:hyperlink'):
            yield 'multipart', self._parse_child_paragraph(paragraph._element)
        else:
            text = util.get_element_text(paragraph._element)
            if text:
                yield 'paragraph', {'text': text, 'style_id': paragraph.style.style_id}

    def _parse_child_paragraph(self, element):
        """parse paragraph with graphic or hyperlink
        """
        data = []
        for child in element.iterchildren():

            if isinstance(child, CT_R) and child.xpath('.//a:graphic'):
                part = self._parse_graphic(child)
            elif isinstance(child, docx.oxml.etree._Element):
                part= self._parse_hyperlink(child)
            else:
                part = util.get_element_text(child)
                if not part:
                    continue
            if part is None:
                continue
            data.append(part)
        return data

    def _parse_graphic(self, element):
        """parse graphic, return image data
        """
        rid = None
        try:
            rid = element.xpath('.//a:blip/@*')[0]
            im = self.document.part.rels[rid]._target
            image, filename = util.blob_to_image(
                im.blob,
                image_as=self.image_as,
                image_type=self.image_type,
                filename=im.sha1,
                media_dir=self.media_dir)
        except:
            return None
        return {
            'type': self.image_as,
            'filename': filename,
            'image': image,
        }

    def _parse_hyperlink(self, element):
        """parse hyperlink, return text and href
        """
        for value in element.values():
            if value.startswith('rId'):
                href = self.document.part.rels[value]._target
                text = util.get_element_text(element)
                return {'text': text, 'href': href}

    def parse_table(self, table, strip=True):
        """parse table, return table data and merged_cells
        """
        data = [
            [cell.text.strip() if strip else cell.text for cell in row.cells]
            for row in table.rows
        ]

        merged_cells = {}
        for x, row in enumerate(table.rows):
            for y, cell in enumerate(row.cells):
                if cell._tc.vMerge or cell._tc.grid_span != 1:
                    tc = (cell._tc.top, cell._tc.bottom,
                          cell._tc.left, cell._tc.right)
                    merged_cells['_'.join(map(str, tc))] = cell.text

        return 'table', {'data': data, 'merged_cells': merged_cells}
