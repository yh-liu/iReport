
# Execute this as 'setup.py py2exe' to create a .exe from docmaker.py
from distutils.core import setup
import py2exe

inc_module = ["win32com.client"]  

py2exe_options = dict(
##    typelibs = [
##        # typelib for 'Word.Application.8' - execute 
##        # 'win32com/client/makepy.py -i' to find a typelib.
##        ('{00020905-0000-0000-C000-000000000046}', 0, 8, 6),
##        #('{00020813-0000-0000-C000-000000000046}', 0, 1, 8),
##        #('{00020905-0000-0000-C000-000000000046}', lcid=0, major=8, minor=6)
##    ],
    includes = inc_module,
    packages = ['win32com.client'], 
    bundle_files = 2,
    optimize = 1
)

setup(name="iReport32_2010",
      console=["IReport32_2010.py"],
      options = {"py2exe" : py2exe_options},
      )

setup(name="iReport32_2013",
      console=["IReport32_2013.py"],
      options = {"py2exe" : py2exe_options},
      )
