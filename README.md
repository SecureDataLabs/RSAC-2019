# RSA Conference 2019 - Rise of the Machines

## Presented by: Etienne Greeff and Wicus Ross

This repository contains the topic modelling Python code as well as the Cobalt Strike Aggressor script. The two Python scripts are independent. Both generate topic modelling information. The _topicmodelscan.py_ script must be converted to an Windows executable and is intended to be used with the Cobalt Strike Aggressor script.

## Requirements

You will need Python 3.7 to run these examples. You will also need to be able to install several modules.

PyInstaller version 3.4 is required to convert _topicmodelscan.py_ to a Windows executable. Unfortunately there are several issues with PyInstaller, including a bug in the NTLD loader modules of PyInstaller and this will require manual fixing. More on that later.

Cobalt Strike 3.13 or newer is required to run the Aggressor script. Cobal Strike is a commercial product. It is possible to get a trial license for it if you wish to play with the Aggressor script.

## Setup

Ensure PIP is up to date before starting. To update it run the following:

    python -m pip install --upgrade pip

In the directory with the requirements.txt file run:

    pip install -r requirements.txt

This will install all the Python modules necessary to run _topicmodel.py_.

PyInstaller is not included in the requirements.txt file as its is not a dependency of either Python applications. PyInstaller is used to package the respective files in a self extracting and self contained executable.

To download and install PyInstaller run the following from the command-line:

    pip install pyinstaller

The _spaCy_ module requires us to download files for a specif language group. If you run the following command on Windows then run it from a command prompt with elevated privileges that allows file system hard links. Running it as Administrator is sufficient. To initialise _spaCy_ run:

    python -m spacy download en

See <https://spacy.io/usage/> for more information on this.

## Topic Model - GUI

The graphical user interface for the topic model application is a Python application. It is recommended to run it with _--help_ argument. For example:

    python topicmodel.py --help

To access the GUI run this Python application on the command-line as:

    python topicmodel.py gui

It can also be used to run in command-line only mode. This can be achieved by substituting _gui_ with either _match_ or _local_:

    python topicmodel.py _match_ --help

or

    python topicmodel.py _local_ --help

The following is an optional step and is not be required. To regenerate the related Qt UI Python class data run:

    python -m PyQt5.uic.pyuic -x topicmodelling.ui -o topicui.py

## Preparing PyInstaller

The scan implementation was created to run with Cobalt Strike without a GUI. To do so we need to convert it to a Windows executable. This is achieved with the help of PyInstaller. There is a problem with PyInstaller version 3.4 as the NLTK and spaCy modules are incorrectly defined or not supported by the current version of PyInstaller. Manual editing is required to fix this.

Locate the Python environment (virtualenv or conda environment) that houses the PyInstaller. On Windows 10 it is typically located at:

    C:\Users\username\AppData\Local\Continuum\anaconda3\envs\

For this example lets assume the environment is _rsaenv37_. When PyInstaller is present (installed with pip) it should be located here:

    C:\Users\username\AppData\Local\Continuum\anaconda3\envs\rsaenv37\Lib\site-packages\PyInstaller

Two fixes are required. First copy the file named *pyi_rth_spacy.py* from this repository to:

    C:\Users\username\AppData\Local\Continuum\anaconda3\envs\rsaenv37\Lib\site-packages\PyInstaller\loader\rthooks

Next, in the same destination location, rename **pyi_rth__nltk.py** to **pyi_rth_nltk.py**. Notice there are two underscores between *rth* and *nltk*.

Now open *C:\Users\username\AppData\Local\Continuum\anaconda3\envs\rsaenv37\Lib\site-packages\PyInstaller\loader\rthooks.dat* with a text editor.

Locate **pyi_rth_nltk.py** in the rthooks.data file. PyInstaller will now match the renamed file to this one.

Finally we need to add a mapping for spaCy. This is the file we copied across earlier. In the **rthooks.dat** file add the following to the list. Please pay careful attention to the file structure. The file ends with a curly brace \}. The line we introduce must precede the curly brace. Our new line follows an existing **rthooks.dat** entry. This preceding line must end with a comma.

When we add the spaCy mapping our file should end with something like this:

        'nltk': ['pyi_rth_nltk.py'],
        'spacy': ['pyi_rth_spacy.py'],
    }

## Topic Model Scan

Before using PyInstaller ensure you followed the instructions in the Preparing PyInstaller section. PyInstaller will fail if you did not. After amending PyInstaller we can proceed to create an executable.

In the directory that contains our Python and spec files run:

    pyinstaller topicmodelscan.spec

*NOTE:* We ran this using Continuum's Anaconda Python 3.7 runtime environment on Windows 10. It was not tested on Linux or MacOS.

PyInstaller will create two directories namely; build and dist. Our executable will be deployed to dist.

## Files or Directories of Interest in this repository

The following is list of files or directories and a brief explanation of purpose.

* hooks
  * This is a custom directory that we specify in the spec file for PyInstaller. These hook files are used by PyInstaller to include modules or resources that it unable to determine on its own.
* DroidSansMono.ttf and stopwords
  * These are required by the wordcloud module when used in the GUI mode.
* pyi_rth_spacy.py
  * This is the PyInstaller loader file that we need to copy to the PyInstaller's rthooks folder.
* requirements.txt
  * The required modules. Feed this file to pip to pull all the dependent modules.
* rsac2019.cna
  * The Aggressor script for Cobalt Strike.
* topicmodel.py, topicmodelling.ui, topicui.py
  * The Graphical User Interface for our Topic Modelling software.
* topicmodelscan.py
  * This standalone command-line version of our Topic Modelling software.
* topicmodel.spec and topicmodelscan.spec
  * The PyInstaller build specification files