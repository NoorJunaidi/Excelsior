#!"c:\program files\python27\python.exe"
# EASY-INSTALL-ENTRY-SCRIPT: 'oletools==0.55.1','console_scripts','olevba'
__requires__ = 'oletools==0.55.1'
import re
import sys
from pkg_resources import load_entry_point

if __name__ == '__main__':
    sys.argv[0] = re.sub(r'(-script\.pyw?|\.exe)?$', '', sys.argv[0])
    sys.exit(
        load_entry_point('oletools==0.55.1', 'console_scripts', 'olevba')()
    )