# coding: utf-8

from distutils.core import setup
import py2exe

setup( console=['MainControl.py'],
       zipfile = None,
       options = {'py2exe' : {
                     "bundle_files": 1,
                     "dll_excludes": ["MSVCP90.dll","w9xpopen.exe"]
                     }
                  }
 )

