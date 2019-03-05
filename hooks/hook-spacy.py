#-----------------------------------------------------------------------------
# Copyright (c) 2013-2018, PyInstaller Development Team.
#
# Distributed under the terms of the GNU General Public License with exception
# for distributing bootloader.
#
# The full license is in the file COPYING.txt, distributed with this software.
#-----------------------------------------------------------------------------
from PyInstaller.utils.hooks import collect_data_files, collect_submodules, get_package_paths
import os

hiddenimports = (
collect_submodules('spacy') +
collect_submodules('spacy.lang.en') +
collect_submodules('thinc')
)
hiddenimports.append('spacy.strings')

pkg_base, pkg_dir = get_package_paths('spacy')
root_dir = pkg_dir

def recursive_walk(base_folder, parent_folder, base_prefix):
    result = []
    folder = os.path.join(base_folder, parent_folder)
    for folderName, subfolders, filenames in os.walk(folder):
        bn = os.path.basename(folderName)
        if subfolders:
            for subfolder in subfolders:
                pf = os.path.join(parent_folder, subfolder)
                r = recursive_walk(base_folder, pf, base_prefix)
                if r:
                    result = result + r
        for filename in filenames:
            i = len(base_folder)
            package_name = base_prefix + folderName[i:]
            full_path = os.path.join(folderName, filename)
            entry = (full_path, package_name)
            result.append(entry)
    return result
datas = recursive_walk(root_dir, 'data', 'spacy')