import os
from inspect import getsourcefile
from version import __version__

# get the path of the current python script
PATH_TB = file_path = os.path.dirname(getsourcefile(lambda: 0))

PATH_TB_ICONS = os.path.join(PATH_TB, "Resources", "Icons")
