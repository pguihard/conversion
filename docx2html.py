"""
convert a docx file to an html file using the mammoth package
py -m pylint docx2html.py
-------------------------------------------------------------------
Your code has been rated at 10.00/10
"""
import sys
import os
from datetime import datetime
import mammoth
from git import Repo

from constants import TAGS
from local_settings import PATH

def upload2github(filename):
    """
    upload the resulting html file to the dropbox cloud
    """
    now = datetime.now()
    dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
    #
    # To initiate new Git repo in the mentioned directory
    repo = Repo.init(f'{PATH}')
    assert isinstance(repo, Repo)

    print(f'INFO : Git repo is initiated.')

    # To check configuration values, use `config_reader()`
    # using ~/.ssh/id_ed25519
    url = 'git@github.com:pguihard/doc.git'
    repo.remotes.origin.set_url(url)
    #
    # Stage the change
    index = repo.index
    index.add([filename])

    print(f'INFO : Git changes to html are staged.')

    # Commit the change
    index.commit(dt_string)

    print(f'INFO : Git stages are committed.')
    # push the commit
    origin = repo.remote(name='origin')
    origin.push()

    print(f'INFO : Git commits are pushed to {url}.')  

def process(fname):
    """ can be processed for one specific year
    """

    input_filename = PATH + fname + ".docx"
    with open(input_filename, "rb") as docx_file:
        result = mammoth.convert_to_html(docx_file)
        html = result.value

    print(f'INFO : docx is converted to html.')
    
    html = '<!doctype html><html lang="fr"><head><meta charset="utf-8">' \
            '<title>Mutation managériale</title>' \
            '<link rel="stylesheet" href="css/style.css" />' \
            '</head><body>' + html + '</body></html>'

    for tag in TAGS:
        html = html.replace(f'<{tag}', f'\n<{tag}')

    print(f'INFO : html is reformatted.')

    output_filename = PATH + "index.html"
    with open(output_filename, "w", encoding="utf-8") as fout:
        fout.writelines(html)

    print(f'INFO : html file is saved.')

    upload2github(output_filename)

#You can pass in argument the file name of the docx file
if __name__ == "__main__":
    if len(sys.argv) > 1:
        process(sys.argv[1])
    else:
        process("CA")
