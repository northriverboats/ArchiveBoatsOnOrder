#!/usr/bin/env python3
""" /samba/shares/production/Boats on Order - Smartsheet edition/

copy  the current sheets in the Boats on Order - Smartsheet edition/
  to  Boats on Order - Smartsheet edition/Archived/Y%_%m_%d oldname
"""

import click
import datetime
from dotenv import load_dotenv
import glob
import os
import shutil
import sys
import time


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def remove(path):
    """
    Remove the file or directory
    """
    if os.path.isdir(path):
        try:
            os.rmdir(path)
        except OSError:
            print("Unable to remove folder: %s" % path)
    else:
        try:
            if os.path.exists(path):
                os.remove(path)
        except OSError:
            print("Unable to remove file: %s" % path)

def cleanup(number_of_days, path):
    """
    Removes files from the passed in path that are older than or equal
    to the number_of_days
    """
    time_in_secs = time.time() - (number_of_days * 24 * 60 * 60)
    for root, dirs, files in os.walk(path, topdown=False):
        for file_ in files:
            full_path = os.path.join(root, file_)
            stat = os.stat(full_path)

            if stat.st_mtime <= time_in_secs:
                remove(full_path)

        if not os.listdir(root):
            remove(root)


def copy_to_archive(source, destination):
    """
    Copy spreadsheets from the current folder to the Archvie folder
    """
    os.chdir(source)
    for file in glob.glob("*.xlsx"):
        filename = destination+datetime.datetime.today().strftime('%Y_%m_%d')+"  "+file
        print(file, filename)
        shutil.copyfile(file, filename)

def delete_from_archive(destination):
    """
    delete files with a name pattern that matches exactly 6 months ago today
    """
    os.chdir(destination)
    for file in glob.glob((datetime.datetime.today()-datetime.timedelta(days=27*7)).strftime('%Y_%m_%d')+"  "+"*.xlsx"):
        if os.path.exists(file):
            os.remove(file)

@click.command()
def main():
    """ /samba/shares/production/Boats on Order - Smartsheet edition/

    copy  the current sheets in the Boats on Order - Smartsheet edition/
      to  Boats on Order - Smartsheet edition/Archived/Y%_%m_%d oldname
    """

    # set python environment
    if getattr(sys, 'frozen', False):
        bundle_dir = sys._MEIPASS
    else:
        # we are running in a normal Python environment
        bundle_dir = os.path.dirname(os.path.abspath(__file__))

    # load environmental variables
    load_dotenv(dotenv_path=resource_path(".env"))

    if os.getenv('HELP'):
      with click.get_current_context() as ctx:
        click.echo(ctx.get_help())
        ctx.exit()

    source=os.getenv('SOURCE')
    destination=os.getenv('DESTINATION')
    copy_to_archive(source, destination)
    # cleanup(26*7+1, destination) # anything over 6 months old


if __name__ == "__main__":
    main()
