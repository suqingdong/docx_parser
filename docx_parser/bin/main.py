import sys
import json

import click

from docx_parser import version_info
from docx_parser.core.parser import DocumentParser


epilog = click.style('''\n
\b
examples:
    docx_parser --help
    docx_parser tests/demo.docx
    docx_parser tests/demo.docx -A base64 -o out.jl

contact: {author} <{author_email}>
'''.format(**version_info), fg='green')


@click.command('docx-parser',
               help=click.style(version_info['desc'], bold=True, fg='cyan'),
               no_args_is_help=True,
               epilog=epilog)
@click.argument('infile')
@click.option('-o', '--outfile', help='the output filename [stdout]')
@click.option('-A', '--image-as', help='extract image as file, base64 or blob',
              type=click.Choice(['file', 'base64']), show_choices=True,
              default='file', show_default=True)
@click.option('-T', '--image-type', help='extract image as file, base64 or blob',
              type=click.Choice(['jpeg', 'png']), show_choices=True,
              default='jpeg', show_default=True)
@click.option('-D', '--media-dir', help='the media directory to save files', default='media', show_default=True)
@click.version_option(version=version_info['version'], prog_name=version_info['prog'])
def main(**kwargs):
    out = open(kwargs['outfile'], 'w') if kwargs['outfile'] else sys.stdout
    with out:
        doc = DocumentParser(kwargs['infile'], **kwargs)
        for each in doc.parse():
            line = json.dumps(each, ensure_ascii=False)
            out.write(line + '\n')


if __name__ == "__main__":
    main()
