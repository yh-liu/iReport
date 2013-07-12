from distutils.core import setup
import py2exe

setup(name = "export_from_CIM",
      console=["export_from_CIM.py"], 
      options = {'py2exe': {
          "bundle_files" : 2,
          "optimize": 1, }
                 }
      )
