Microsoft PSTREAM File Extractor
================================
Utility for extracting PSTREAM files (.psf) written in Python.
Created to replace other laughably bad approaches to this.

Usage
-----
```
psfextract.py <psf_file> <destination>
```

To extract a PSTREAM file you need to specify the file you want to extract and
the destination directory.

The destination directory should not exist or be empty.

Requirements
------------
This script requires at least Python 3 and works only on Windows because of
the *msdelta.dll* library dependency.

License
-------
This script is licensed under the terms of the GNU General Public License v3.0.
