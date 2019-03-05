#-----------------------------------------------------------------------------
# Copyright (c) 2013-2018, PyInstaller Development Team.
#
# Distributed under the terms of the GNU General Public License with exception
# for distributing bootloader.
#
# The full license is in the file COPYING.txt, distributed with this software.
#-----------------------------------------------------------------------------


import sys
import os
import spacy
from pathlib import Path

#add the path to spacy.data
spacy_data_path = os.path.join('spacy', 'data')
dp = os.path.join(sys._MEIPASS, spacy_data_path)
spacy.util._data_path = Path(dp)
