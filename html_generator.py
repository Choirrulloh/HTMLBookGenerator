#!/usr/bin/env python3
# Author: Adam Hlavacek, 2018, https://adamhlavacek.com/
# LICENSE: MIT

import re
import time
import tempfile
import base64

from sys import argv
from subprocess import call
from os.path import isfile, join, basename, abspath
from os import listdir
from shutil import rmtree
from codecs import open
from mimetypes import MimeTypes
from urllib.request import pathname2url
from urllib.parse import unquote


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

new_temp_file_path = join(temp_dir, [file for file in listdir(temp_dir) if file.endswith('.html')][0])

try:
    output_file_path = argv[2]
except IndexError:
    output_file_path = basename(new_temp_file_path)
    
output_file_path = abspath(output_file_path)

with open(new_temp_file_path, 'r', 'utf-8', errors='replace') as f:
    content = f.read()

head_html = """
<meta name="viewport" content="width=device-width, initial-scale=1.0">
"""

upper_html = """
<style>
#document{
    overflow: hidden;
    height: 100%;
    width: calc(100% - .5em);
    max-width: 30cm;
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

pre {
    white-space: pre-wrap;
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
let disableClicks = false;  // can be set to true from developer tools in browser for easier debugging

const $ = document.querySelector.bind(document);
const body = $('body');
const doc = $('#document');
const colorPalete = $('#colorPalette');
const percentage = $('#percentage');
const pageNum = $('#pageNum');


const currentColorPalette = {bck: null, fg: null, fnt: null, lh: null};

let scrollHeight = null;

function countScrollHeight(){
    let div = document.createElement('div');
    div.style.position = "fixed";
    div.style.visibility = "hidden";
    div.style.width = `calc(${currentColorPalette.fnt} * ${currentColorPalette.lh})`;
    document.body.appendChild(div);
    scrollHeight = doc.getBoundingClientRect().height - div.getBoundingClientRect().width;
    div.parentNode.removeChild(div);
}

function movePages(pageDelta){
    countScrollHeight();
    doc.scrollTop += pageDelta * scrollHeight;
    showStats();
}

document.addEventListener('click', (e) => {
    if (disableClicks)
        return;
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
    else if (e.key.toLowerCase() === "f")
       toggleFullscreen();
});

// Switch fullscreen when stats are pressed
$('#info').onclick = (e) => {
    if (disableClicks)
        return;
    toggleFullscreen();
    e.stopPropagation();
}

function toggleFullscreen(){
    if(document.fullscreenElement)
        document.exitFullscreen();
    else
        body.requestFullscreen();
    for (let time of [1, 100, 250, 500, 1000, 1500, 2000])
        setTimeout(() => {
            showStats();
            countScrollHeight();
            showStats();
        }, time);
}

function showStats(){
    let percentageNum = (doc.scrollTop + doc.getBoundingClientRect().height) * 100 / doc.scrollHeight;
    percentage.style.width = `${percentageNum}%`;
    
    let pages = Math.round(doc.scrollHeight / scrollHeight);
    let currentPage = Math.round(
        (doc.scrollTop + doc.getBoundingClientRect().height) / scrollHeight
    );
    
    pageNum.innerText = `${currentPage}/${pages}`;
}

function setColorPalette(backgroundColor=null, foregroundColor=null, fontSize=null, lineHeight=null){
    if (backgroundColor !== null)
        currentColorPalette.bck = backgroundColor;
    if (foregroundColor !== null)
        currentColorPalette.fg = foregroundColor;
    if (fontSize !== null)
        currentColorPalette.fnt = fontSize;
    if (lineHeight !== null)
        currentColorPalette.lh = lineHeight;
    colorPalette.innerHTML = `
html, body, #info {
    background: ${currentColorPalette.bck} !important;
}

p,a,span,#pageNum, li, table, td, tr, #document, h1, h2, h3, h4, h5, h6, h7, h8, h9 {
    color: ${currentColorPalette.fg} !important;
    font-size: ${currentColorPalette.fnt} !important;
    border-color:  ${currentColorPalette.fg} !important;
    line-height: ${currentColorPalette.lh}  !important;
}
`;}

window.onload = _ => {
    setColorPalette("#1e1e1e", "#effdff", "20px", "1.5"); // dark pallete
    // setColorPalette("#ffffff", "#000000", "20px", "1.5"); // light pallete
    // setColorPalette("#fee5b3", "#695445", "20px", "1.5"); // warm pallete
    showStats();  // this call is required to initialize GUI
    countScrollHeight();
    showStats();  // this call is required because first call used scrollHeight which was null at the time
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

# base64 include all external files
mime = MimeTypes()
while True:
    result = re.search('(<img)(.*src=")(.*?)(".*?>)', content, re.I | re.M)
    if result is None or result.groups()[2].endswith('base64-ed '):
        break

    external_file = join(temp_dir, basename(unquote(result.groups()[2])))
    if not isfile(external_file):
        pass

    file_url = pathname2url(external_file)
    mime_type = mime.guess_type(file_url)[0]
    if mime_type is None:
        mime_type = 'unknown/unknown'
        print('WRONG MIME', mime.guess_type(file_url), file_url, external_file)
    b64_data = "data:" + mime_type + ";base64,"
    with open(external_file, 'rb') as f:
        b64_data += base64.b64encode(f.read()).decode('ascii')
    content = re.sub('(<img)(.*src=")(.*?>)', '<b64-img' + result.groups()[1] + b64_data + result.groups()[3],
                     content, flags=re.I | re.M, count=1)
    del b64_data

content = content.replace('<b64-img', '<img')

# remove the temporary directory - we have the file in memory, the file on dict is not needed anymore
rmtree(temp_dir)

with open(output_file_path, 'w', 'utf-8') as f:
    f.write(content)

print('Your new html doc was saved as %s' % output_file_path)
