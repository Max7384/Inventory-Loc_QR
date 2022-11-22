#!pip install dominate
#!pip install opencv-python
#!pip install pandas
#!pip install qrcode[pil]
#!pip install -e git+http://github.com/Max7384/pymaging.git#egg=pymaging
#!pip install -e git+http://github.com/Max7384/pymaging-png.git#egg=pymaging-png
#%pip install xlrd #xls handling

import dominate
from dominate.tags import *
import cv2
import os
import datetime

import pandas as pd
import zipfile as zp
import qrcode
import qrcode.image.svg
from qrcode.image.pure import PymagingImage
import pandas as pd
from pathlib import Path

from configparser import ConfigParser

# set up everything

cfg = ConfigParser()
cfg.read('config.ini')

viertel = False
total = True

sizeQR = cfg.get('sizeQR','sizeQR')

listInvStartsAndEnds = []


# helper functions
def createListstotal(filterIdCol, sizeQR):
    for e in filterIdCol:
        listInvstartEndtotal = createQRCodeStartAndEnd(e, sizeQR)
        #print(e, listInvstartEndtotal)
    #print(listInvstartEndtotal)
    return listInvstartEndtotal

def dirExists(dirName):
    Path(os.path.join(dirName)).mkdir(parents=True, exist_ok=True)

def qr_object(box_size):
    qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_M,
            box_size=box_size,
            border=4
        )
    return qr

def createQRCodeStartAndEnd(code,box_size=5):
    method='fragment'
    if method == 'basic':
        factory = qrcode.image.svg.SvgImage
    elif method == 'fragment':
        factory = qrcode.image.svg.SvgFragmentImage
    else:
        factory = qrcode.image.svg.SvgPathImage


    folder_string = str(("QR_Gr{}/{}").format(str(box_size), str(code[1][0])))
    dirExists(folder_string)

    loc_name = str(code[1])
    bereich = str(code[3])

    InvStart = cfg.get('commands','InvStart')
    data_start = ("{}{}").format(str(InvStart),str(code[0]))
    qr_start = qr_object(box_size)
    qr_start.add_data(data_start)
    qr_start.make(fit=True)
    img_start = qr_start.make_image(fill_color="black", back_color="white")
    filename_start = str(("{}/{}_26_Gr{}_{}.png").format(str(folder_string),loc_name,str(box_size),str(code[0])))
    img_start.save(filename_start)

    InvEnd = cfg.get('commands','InvEnd')
    data_end = ("{}{}").format(str(InvEnd),str(code[0]))
    qr_end = qr_object(box_size)
    qr_end.add_data(data_end)
    qr_end.make(fit=True)
    img_end = qr_end.make_image(fill_color="black", back_color="white")
    filename_end = str(("{}/{}_27_Gr{}_{}.png").format(str(folder_string),loc_name,str(box_size),str(code[0])))
    img_end.save(filename_end)

    #onlyloc
    data_loc = f'loc{chr(35)}0{str(code[4])}{chr(35)}{code[0]}'
    qr_loc = qr_object(box_size)
    qr_loc.add_data(data_loc)
    qr_loc.make(fit=True)
    img_loc = qr_loc.make_image(fill_color="black", back_color="white")
    filename_loc = str(("{}/{}_loconly_Gr{}_{}.png").format(str(folder_string),loc_name,str(box_size),str(code[0])))
    img_loc.save(filename_loc)

    listInvStartsAndEnds.append([data_start, filename_start, data_end,filename_end, filename_loc, code[2], bereich, data_loc])
    return listInvStartsAndEnds

def severalonPage(items,htmlname,sizeMarker,decr,neuerbereich):
    h = html()
    with h.add(body()).add(div(id='content')):
        link(rel='stylesheet', href='style.css')
        script(type='text/javascript', src='script.js')
        style(
           """empty {
                 text-align: center;
                 vertical-align: center;
                 color: #2C232A;
                 font-family: sans-serif;
                 font-size: 1em;
                 }
              .location {
                 text-align: center;
                 vertical-align: bottom;
                 #color: #2C232A;
                 font-family: sans-serif;
                 font-size: 1em;
                 }
              .command {
                 text-align: center;
                 vertical-align: center;
                 color: #2C232A;
                 font-family: sans-serif;
                 font-size: 0.6em;
                 }
              .photo {
                 text-align: center;
                 vertical-align: center;
                 }
              .ID {
                 text-align: center;
                 vertical-align: top;
                 color: #2C232A;
                 font-family: sans-serif;
                 font-size: 0.5em;
                 }
               .pbreak {
                 page-break-after: always;
                 }
                  """)


        for i in items:


            commanddescription_start = i[0]
            path_start_img = i[1]
            commanddescription_end = i[2]
            path_end_img = i[3]
            path_loc_img = i[4]
            location_descr = i[5]
            bereich = i[6]

            with table().add(tbody()):

                l = tr()
                l += td(div(img(src=path_start_img), _class='photo'), div((commanddescription_start), _class='ID'),div((''),_class='empty'))
                l.add(td(div(p((str(location_descr),str(decr))), _class='location')))
                with l:
                    td(div(img(src=path_end_img), _class='photo'), div((commanddescription_end), _class='ID'))
                if neuerbereich != bereich:
                    neuerbereich = bereich
                    #print("neuerbereich: ",bereich)
                    div( _class='pbreak')
                    h1('Bereich: ',bereich)

    name = "{}_{}_win_{}.html".format(htmlname,sizeMarker,str(datetime.datetime.now())[20:]) # Windows

    with open(name, 'w') as f:
        f.write(h.render(pretty=True))
    print("{} was created!".format(name))

def oneperpage_startinv(items,htmlname,sizeMarker):
    onepp = html()
    with onepp.add(body()).add(div(id='content')):
        link(rel='stylesheet', href='style.css')
        script(type='text/javascript', src='script.js')
        style(
           """ID {
                 text-align: center;
                 vertical-align: center;
                 color: #2C232A;
                 font-family: sans-serif;
                 font-size: 0.6em;
                 }
               .location {
                 text-align: left;
                 vertical-align: center;
                 color: #2C232A;
                 font-family: sans-serif;
                 font-size: 2.5em;
                 }
              .command {
                 text-align: center;
                 vertical-align: top;
                 color: #2C232A;
                 font-family: sans-serif;
                 font-size: 0.6em;
                 }
              .pbreak {
                 page-break-before: always;
                 }
                  """)
        for i in items:
            location_descr = i[5]
            commanddescription_start = i[0]
            path_start_img = i[1]
            commanddescription_end = i[2]
            path_end_img = i[3]
            path_loc_img = i[4]
            location_descr = i[5]
            bereich = i[6]

            #print(i[1])
            div((location_descr), _class='location')
            div(img(src=path_start_img), _class='photo')
            div((commanddescription_start), _class='ID')
            div( _class='pbreak')
    #if windows == 1:
    name = "{}_{}_win_{}.html".format(sizeMarker,htmlname,str(datetime.datetime.now())[20:])

    with open(name, 'w') as f:
        f.write(onepp.render(pretty=True))
    print("{} was created!".format(name))

def oneperpage_endinv(items,htmlname,sizeMarker):
    onepp = html()
    with onepp.add(body()).add(div(id='content')):
        link(rel='stylesheet', href='style.css')
        script(type='text/javascript', src='script.js')
        style(
           """ID {
                 text-align: center;
                 vertical-align: center;
                 color: #2C232A;
                 font-family: sans-serif;
                 font-size: 0.6em;
                 }
               .location {
                 text-align: left;
                 vertical-align: center;
                 color: #2C232A;
                 font-family: sans-serif;
                 font-size: 2.5em;
                 }
              .command {
                 text-align: center;
                 vertical-align: top;
                 color: #2C232A;
                 font-family: sans-serif;
                 font-size: 0.6em;
                 }
              .pbreak {
                 page-break-before: always;
                 }
                  """)
        for i in items:
            location_descr = i[5]
            commanddescription_start = i[0]
            path_start_img = i[1]
            commanddescription_end = i[2]
            path_end_img = i[3]
            path_loc_img = i[4]
            location_descr = i[5]
            bereich = i[6]

            div((location_descr), _class='location')
            div(img(src=path_end_img), _class='photo')
            div((commanddescription_end), _class='ID')
            div( _class='pbreak')
    #if windows == 1:
    name = "{}_{}_win_{}.html".format(sizeMarker,htmlname,str(datetime.datetime.now())[20:])

    with open(name, 'w') as f:
        f.write(onepp.render(pretty=True))
    print("{} was created!".format(name))

def oneperpage_loc(items,htmlname,sizeMarker):
    onepp = html()
    with onepp.add(body()).add(div(id='content')):
        link(rel='stylesheet', href='style.css')
        script(type='text/javascript', src='script.js')
        style(
           """ID {
                 text-align: center;
                 vertical-align: center;
                 color: #2C232A;
                 font-family: sans-serif;
                 font-size: 0.6em;
                 }
               .location {
                 text-align: left;
                 vertical-align: center;
                 color: #2C232A;
                 font-family: sans-serif;
                 font-size: 2.5em;
                 }
              .command {
                 text-align: center;
                 vertical-align: top;
                 color: #2C232A;
                 font-family: sans-serif;
                 font-size: 0.6em;
                 }
              .pbreak {
                 page-break-before: always;
                 }
                  """)
        for i in items:
            location_descr = i[5]
            commanddescription_start = i[0]
            path_start_img = i[1]
            commanddescription_end = i[2]
            path_end_img = i[3]
            path_loc_img = i[4]
            location_descr = i[5]
            bereich = i[6]
            commanddescription_loc = i[7]

            #print(i[1])
            div((location_descr), _class='location')
            div(img(src=path_loc_img), _class='photo')
            div((commanddescription_loc), _class='ID')
            div( _class='pbreak')
    #if windows == 1:
    name = "{}_{}_win_{}.html".format(sizeMarker,htmlname,str(datetime.datetime.now())[20:])

    with open(name, 'w') as f:
        f.write(onepp.render(pretty=True))
    print("{} was created!".format(name))





def main():
    neuerbereich = None
    xlsname = cfg.get('sourcexls','filenamexls')
    sheetname = cfg.get('sourcexls','sheetname')
    df = pd.read_excel(open(xlsname,'rb'),sheet_name=sheetname)
    filterIdCol = list(zip(df.ID,df.filename,df.name,df.bereich,df.idwl))
    listInvstartEndtotal = createListstotal(filterIdCol, sizeQR)

    severalonPage(listInvstartEndtotal,'SeveralOnPage_StartEndInv',sizeQR,' ',neuerbereich)
    oneperpage_startinv(listInvstartEndtotal,'OnePerPage_StartInv',sizeQR)
    oneperpage_endinv(listInvstartEndtotal,'OnePerPage_EndInv',sizeQR)
    oneperpage_loc(listInvstartEndtotal,'OnePerPage_Loc',sizeQR)


main()
