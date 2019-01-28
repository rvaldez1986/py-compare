import sys, os
from cx_Freeze import setup, Executable

project_name = "My Tkinter App"
exe_name = "MyTkinterApp"
version = "1"
main_script = "main.py"
optimization = 0
packages = ["tkinter","numpy","pandas","openpyxl"]
resources = ['Existing_File.xlsx','logo.gif']

ex_packages = ["scipy", "PyQt4.QtSql", "sqlite3", 
                       "scipy.lib.lapack.flapack", "PyQt4.QtNetwork", "PyQt4.QtScript", "PyQt5", "matplotlib", 
                           "Cython", "PyQt4", "tornado", "zmq",  
                           "jinja2","cffi",
                           "numexpr","babel","notebook","lxml", 
                           "cryptography","PIL","bottleneck","boto","IPython","docutils", #new from this line
                            "statsmodels","tables","sqlalchemy","test","sphinx"                           
                           ]




base = None
if sys.platform == "win32":
    base = "Win32GUI"  #if you want back console, comment line
    executables = [Executable("main.py", base=base, targetName=exe_name+".exe")]
    
    
    os.environ['TCL_LIBRARY'] = sys.exec_prefix + "\\tcl\\tcl8.6"
    os.environ['TK_LIBRARY'] = sys.exec_prefix + "\\tcl\\tk8.6"
    resources.extend([sys.exec_prefix + "\\DLLs\\tcl86t.dll", sys.exec_prefix + "\\DLLs\\tk86t.dll"])
    
elif sys.platform == "darwin":
    executables = [Executable("main.py", base=base)]

build_exe_options = {"optimize": optimization, "include_files": resources, "packages": packages, "excludes": ex_packages}
mac_options = {"bundle_name": exe_name}


setup(  name = project_name,
        version = version,
        options = {"build_exe": build_exe_options},
        bdist_mac = mac_options,
        executables = [Executable(main_script, base=base, targetName=exe_name+".exe")]
)