#!/usr/bin/env python3
# Author: Adam Hlavacek, 2018, https://adamhlavacek.com/
# LICENSE: MIT

import re
import time
import tempfile

from sys import argv
from subprocess import call
from os.path import isfile, join, basename, abspath
from os import listdir
from shutil import rmtree

SOFFICE_PATH = "soffice"


try:
    assert isfile(argv[1])
except IndexError or AssertionError:
    print('usage: %s pathToTheFileToConvert [outputPath]' % argv[0])
    exit(1)

try:
    call([SOFFICE_PATH, "--version"])
except FileNotFoundError:
    print('this program requires libre office\'s core to be installed')
    print('it is not installed, or the executable path ("%s") is invalid' % SOFFICE_PATH)
    exit(2)

temp_dir = tempfile.mkdtemp()

call([SOFFICE_PATH, "--headless", "--convert-to", "html", "--outdir", temp_dir, argv[1]])
time.sleep(2)

new_temp_file_path = join(temp_dir, listdir(temp_dir)[0])

try:
    output_file_path = argv[2]
except IndexError:
    output_file_path = basename(new_temp_file_path)
    
output_file_path = abspath(output_file_path)

with open(new_temp_file_path, 'r') as f:
    content = f.read()

# remove the temporary directory - we have the file in memory, the file on dict is not needed anymore
rmtree(temp_dir)

head_html = """
<meta name="viewport" content="width=device-width, initial-scale=1.0">
"""

upper_html = """
<style>
#document{
    overflow: hidden;
    height: 100%;
    width: calc(100% - .5em);
    max-width: 35cm;
}

html, body {
    width: 100%;
    height: 100%;
    padding: 0;
    margin: 0;
    display: flex;
    flex-direction: column;
    align-items: center;
}

#menuBtn {
    position: fixed;
    opacity: 0.5;
    right: 0;
    font-size: 20px;
}

#percentage {
    background: #07bfab;;
    height: 5px;
    width: 0;
}

#info, #pageNum {
    width: 100%;
    z-index: 5000;
}

#pageNum {
    text-align: center;
}
    </style>
    <style id="colorPalette">
    </style>
<div id="info">
<div id="percentage"></div>
<div id="pageNum"></div>
</div>
<span id="menuBtn">&#9776;</span>
<div id="document">
"""

lower_html = """
</div><script>
const $ = document.querySelector.bind(document);
const body = $('body');
const doc = $('#document');
const colorPalete = $('#colorPalette');
const percentage = $('#percentage');
const pageNum = $('#pageNum');


const currentColorPalette = {bck: null, fg: null, fnt: null};

function movePages(pageDelta){
    doc.scrollTop += pageDelta * body.getBoundingClientRect().height * 0.9;
    showStats();
}

document.addEventListener('click', (e) => {
    pageDelta = e.pageX > (body.getBoundingClientRect().width / 2) ? 1 : -1;
    movePages(pageDelta);
});

document.addEventListener('keydown', e => {
    e = e || window.event;
    pageNext = [37, 38, 33];
    pagePrev = [39, 40, 34];
    if (pageNext.indexOf(e.keyCode) !== -1)
        movePages(-1);
    else if (pagePrev.indexOf(e.keyCode) !== -1)
       movePages(1);
});

function showStats(){
    let percentageNum = (doc.scrollTop + doc.getBoundingClientRect().height) * 100 / doc.scrollHeight;
    percentage.style.width = `${percentageNum}%`;
    let pages = Math.round(doc.scrollHeight / (body.getBoundingClientRect().height * 0.9));
    let currentPage = Math.round(
        (doc.scrollTop + doc.getBoundingClientRect().height) / (body.getBoundingClientRect().height * 0.9)
    );
    pageNum.innerText = `${currentPage}/${pages}`;
}

function setColorPalette(backgroundColor=null, foregroundColor=null, fontSize=null){
    if (backgroundColor !== null)
        currentColorPalette.bck = backgroundColor;
    if (foregroundColor !== null)
        currentColorPalette.fg = foregroundColor;
    if (fontSize !== null)
        currentColorPalette.fnt = fontSize;
    colorPalette.innerHTML = `
body, #info {
    background: ${currentColorPalette.bck} !important;
}

p,a,span,#pageNum, li, table, td, tr, #document {
    color: ${currentColorPalette.fg} !important;
    font-size: ${currentColorPalette.fnt} !important;
    border-color:  ${currentColorPalette.fg} !important;
}
`;}

window.onload = _ => {
    setColorPalette("#1e1e1e", "#effdff", "20px");
    showStats();
}

</script>
"""

# disable all attributes changing font size or color
content = content.replace("font-size", "nope").replace("size=\"", "sizeIs=\"").replace("color=\"", "coolor=\"")

# put head code right after the head tag
content = re.sub("<head(.*>)", "<head \\1" + head_html, content, re.I | re.M)

# put upper body code right after the body tag
content = re.sub("<body(.*>)", "<body \\1" + upper_html, content, re.I | re.M)

# put lower body code right before the closing of body tag
content = re.sub("</body(.*>)", lower_html + "</body \\1", content, re.I | re.M)

with open(output_file_path, 'w') as f:
    f.write(content)

print('Your new html doc was saved as %s' % output_file_path)
