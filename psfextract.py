#!/usr/bin/env python3
#
# Microsoft PSTREAM File Extractor
# Copyright (C) 2023 asdcorp
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.

import struct
import ctypes
import os
import sys
import logging
import xml.etree.ElementTree as ET

def unpack(file):
    logging.debug(f'Unpacking {file}')

    result = ctypes.windll.msdelta.ApplyDeltaW(
        ctypes.c_longlong(0),
        None,
        ctypes.c_wchar_p(file),
        ctypes.c_wchar_p(file)
    )

    if result:
        return True

    logging.error(f'Failed to unpack {file}')
    return False

def extract(f, off, leng, dest):
    logging.debug(f'Extracting data from PSTREAM to {dest}')
    os.makedirs(os.path.dirname(dest), exist_ok=True)

    buff = 1024*1024
    f.seek(off)
    read = 0
    with open(dest, 'wb') as d:
        while(read < leng):
            if (read+buff) >= leng:
                buff = leng-read

            d.write(f.read(buff))
            read += buff

if __name__ == '__main__':
    if sys.platform != 'win32':
        print('PSTREAM files can be only extracted on Windows')
        exit(1)

    if len(sys.argv) != 3:
        print('Usage:')
        print('psfextract.py <psf_file> <destination>')
        exit(1)

    logging.basicConfig(level=logging.INFO)

    psf_file = sys.argv[1]
    dest_dir = sys.argv[2]
    dir = os.path.abspath(dest_dir).replace('\\', '/') + '/'

    if os.path.exists(dir) and len(os.listdir(dir)) != 0:
        logging.critical(f'{dest_dir} is not empty!')
        exit(1)
    elif not os.path.exists(dir):
        os.mkdir(dir)

    with open(psf_file, 'rb') as f:
        if f.read(7) != b'PSTREAM':
            logging.critical('Specified file is not PSTREAM!')
            exit(1)

        f.seek(40)
        offset, length = struct.unpack('ll', f.read(8))

        logging.debug(f'Manifest offset: {offset}')
        logging.debug(f'Manifest packed length: {length}')

        extract(f, offset, length, dir + 'manifest.cix.xml')
        if not unpack(dir + 'manifest.cix.xml'):
            logging.critical('Failed to unpack the manifest!')
            exit(1)

        logging.debug('Parsing manifest XML')
        tree = ET.parse(dir + 'manifest.cix.xml')
        root = tree.getroot()

        logging.debug('Getting files')
        files = root.findall('./{urn:ContainerIndex}Files/{urn:ContainerIndex}File')
        filecount = len(files)
        extracted = 0

        for file in files:
            fil_name = file.get('name')
            name = dir + fil_name.replace('\\', '/')
            logging.debug(fil_name)

            logging.debug(f'Getting delta for {name}')
            delta = file.find('./{urn:ContainerIndex}Delta/{urn:ContainerIndex}Source')

            type = delta.get('type')
            offset = int(delta.get('offset'))
            length = int(delta.get('length'))

            logging.debug(f'Delta type: {type}')
            logging.debug(f'Delta offset: {offset}')
            logging.debug(f'Delta length: {length}')

            extract(f, offset, length, name)

            if type == 'PA30' and not unpack(name):
                logging.critical('Failed to unpack {name}!')
                exit(1)

            extracted += 1
            if extracted % 100 == 0 or extracted == filecount:
                percent = extracted / filecount
                logging.info(f'{extracted}/{filecount} files extracted ({percent:.2%})')
