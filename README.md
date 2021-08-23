Boilerplate code for renaming files, using an Excel spreadsheet as input.
Please be sure to test this program thoroughly and backup files before 
using it.

# Setup

This program was written on a Mac. If you're using Windows some of the steps
below might differ. I use [homebrew](https://brew.sh/) to install things like
Python 3 and git on my system- other options might make more sense for you 
depending on how you have everything set up.

Before starting, confirm that Python 3 is installed on your machine. To do
that, open a terminal prompt and type the following:

```console
$ python3 --version
```

If you see an error message, please install Python 3 on your system before
continuing.

Next, navigate to a place where you'd like to clone this repo. Type the
following commands to clone the repo and cd into it:

```console
$ git clone https://github.com/johnjung/rename_boilerplate
$ cd rename_boilerplate
```

I use [Python virtual
environments](https://docs.python.org/3/tutorial/venv.html) to keep packages
separate from each other. To set up and activate a virtual environment for
this project, type the following commands:

```console
$ python3 -m venv venv
$ source venv/bin/activate
```

This program uses the [OpenPyXL](https://openpyxl.readthedocs.io/en/stable/)
library to read Excel files. Use pip to install that module:

```console
$ pip install -r requirements.txt
```

Copy your spreadsheet into the directory containing this script. The name
of the spreadsheet is hardcoded into the script- you can change the name
of the spreadsheet at the top of the code if you need to.

To run the script, do:

```console
$ python renamer.py --dry-run
```

The program was written in as defensive of a way as possible, but because it 
hasn't been tested, and especially because it hasn't been tested with your
files, you should review it carefully. Use it in --dry-run mode and review
the code until you're confident enough to use it to modify files, and be 
sure to create a backup before renaming files.

