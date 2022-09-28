import sys
import json

import click

from docx_parser.core.parser import DocumentParser


@click.command('docx-parser', no_args_is_help=True)
@click.argument('infile')
@click.option('-o', '--outfile', help='the output filename [stdout]')
@click.option('-A', '--image-as', help='extract image as file, base64 or blob',
              type=click.Choice(['file', 'base64', 'blob']), show_choices=True,
              default='file', show_default=True)
@click.option('-T', '--image-type', help='extract image as file, base64 or blob',
              type=click.Choice(['jpeg', 'png']), show_choices=True,
              default='jpeg', show_default=True)
@click.option('-D', '--media-dir', help='the media directory to save files', default='media', show_default=True)
def main(**kwargs):
    out = open(kwargs['outfile'], 'w') if kwargs['outfile'] else sys.stdout
    with out:
        doc = DocumentParser(kwargs['infile'], **kwargs)
        for each in doc.parse():
            line = json.dumps(each, ensure_ascii=False)
            out.write(line + '\n')


if __name__ == "__main__":
    main()
