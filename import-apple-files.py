#!/usr/bin/env python2.7
# coding=utf-8
"""
A command-line script for Windows to download all files from connected iphone/ipad to specified directory

Based on https://github.com/dblume/list-photos-on-phone

Project home: https://github.com/abcdenis/import-apple-files
"""

import codecs
import datetime
import logging
import os
import re
import sys
import time
from argparse import ArgumentParser

import pythoncom
import pywintypes
import win32con
import win32file
from win32com.shell import shell, shellcon
from pywintypes import IID

log = logging.getLogger("main")

UNIX_START = datetime.datetime(1970, 1, 1)
BUFFER_SIZE = 16 << 20  # 16 MiB
REFRESH_ENABLED = False  # does not work now

# TODO: find a way to refer to constants from
#   pywin32-219\com\win32comext\propsys\pscon.py
PKEY_Size = (IID('{B725F130-47EF-101A-A5F1-02608C9EEBAC}'), 12)
PKEY_DateCreated = (IID('{B725F130-47EF-101A-A5F1-02608C9EEBAC}'), 15)


def console(msg):
    log.info(u"[console] " + msg)
    sys.stdout.write(msg + "\n")
    sys.stdout.flush()


def fix_timezone(utc):
    dt = datetime.datetime(
        year=utc.year,
        month=utc.month,
        day=utc.day,
        hour=utc.hour,
        minute=utc.minute,
        second=utc.second,
        microsecond=utc.msec * 1000
    )
    result = (dt - UNIX_START).total_seconds()
    return result


def change_file_creation_time(fname, newtime):
    # https://stackoverflow.com/questions/4996405/how-do-i-change-the-file-creation-date-of-a-windows-file
    wintime = pywintypes.Time(newtime)
    winfile = win32file.CreateFile(
        fname,
        win32con.GENERIC_WRITE,
        win32con.FILE_SHARE_READ | win32con.FILE_SHARE_WRITE | win32con.FILE_SHARE_DELETE,
        None,
        win32con.OPEN_EXISTING,
        win32con.FILE_ATTRIBUTE_NORMAL,
        None)

    win32file.SetFileTime(winfile, wintime, None, None)
    winfile.close()


def looks_like_disk_root(pidl):
    if not isinstance(pidl, list):
        return False
    if not pidl:
        return False

    match = re.match(r"^/([A-Z]:)\\", pidl[0])
    if not match:
        return False

    log.debug("looks like root of disk %s", match.group(1))
    return True


def is_my_computer_path_obj(path_obj, max_num=100):
    disk_root_count = 0
    for pidl, idx in zip(path_obj, xrange(max_num)):
        if looks_like_disk_root(pidl):
            disk_root_count += 1
    return disk_root_count > 0


def get_computer_shell_folder():
    """
    Return the local computer's shell folder.
    """
    desktop = shell.SHGetDesktopFolder()
    log.debug(u"desktop: %s", desktop)
    candidates = list()

    for pidl in desktop.EnumObjects(0, shellcon.SHCONTF_FOLDERS):
        display_name = desktop.GetDisplayNameOf(pidl, shellcon.SHGDN_NORMAL)
        log.debug(u"desktop folder: %s", display_name)

        some_obj = desktop.BindToObject(pidl, None, shell.IID_IShellFolder2)
        if is_my_computer_path_obj(some_obj):
            log.info(u"LOOKS LIKE MyComputer: %s", display_name)
            candidates.append(some_obj)

    if not candidates:
        console("ERROR: cant find My Computer")
        sys.exit(1)

    if len(candidates) > 1:
        console(u"ERROR: too many candidates to My Computer (%d): %s" % (len(candidates), candidates))
        sys.exit(1)

    return candidates[0]


def get_dcim_folder(device_pidl, parent):
    """
    Tries to find an iPhone by searching the pidl for the path
    "Internal Storage\\DCIM".
    :param device_pidl: A candidate Windows PIDL for the iPhone
    :param parent: The parent folder of the PIDL
    """
    device_name = parent.GetDisplayNameOf(device_pidl, shellcon.SHGDN_NORMAL)
    name = None
    pidl = None

    folder = parent.BindToObject(device_pidl, None, shell.IID_IShellFolder2)
    try:
        top_dir_name = ""
        for pidl in folder.EnumObjects(0, shellcon.SHCONTF_FOLDERS):
            top_dir_name = folder.GetDisplayNameOf(pidl, shellcon.SHGDN_NORMAL)
            break  # Only want to see the first folder.
        if top_dir_name != "Internal Storage":
            return None, None, device_name
    except pywintypes.com_error:
        return None, None, device_name  # No problem, must not be an iPhone

    folder = folder.BindToObject(pidl, None, shell.IID_IShellFolder2)
    for pidl in folder.EnumObjects(0, shellcon.SHCONTF_FOLDERS):
        name = folder.GetDisplayNameOf(pidl, shellcon.SHGDN_NORMAL)
        break  # Only want to see the first folder.
    if name != "DCIM":
        if name is None:
            console(u"found empty folder '%s\\%s'. Make sure you unlocked iPhone." % (device_name, top_dir_name))
        log.debug(u"%s's '%s' has '%s', not a 'DCIM' dir." %
                  (device_name, top_dir_name, name))
        return None, None, device_name

    return pidl, folder, device_name


def do_refresh_on_shell_view(pidl):
    """
    Does not work now.
    TODO: fix to avoid manual refreshing folders in Windows Explorer

    Got from:
    https://stackoverflow.com/questions/29004302/refreshing-a-folder-that-doesnt-exist-in-the-file-system
    """
    if REFRESH_ENABLED:
        console("refresh...")
        shell.SHChangeNotify(shellcon.SHCNE_MKDIR, shellcon.SHCNF_IDLIST | shellcon.SHCNF_FLUSH, pidl, None)
        shell.SHChangeNotify(shellcon.SHCNE_CREATE, shellcon.SHCNF_IDLIST | shellcon.SHCNF_FLUSH, pidl, None)
        shell.SHChangeNotify(shellcon.SHCNE_ASSOCCHANGED, shellcon.SHCNF_IDLIST | shellcon.SHCNF_FLUSH, pidl, None)
        console("done.")


class AppleFilesImporter(object):

    def __init__(self, target_dir):
        assert os.path.isdir(target_dir), "directory not found: " + target_dir
        self.target_dir = target_dir
        self.new_files = 0
        self.total_files = 0
        self.dir_stats = dict()

    def run(self):
        # Find the iPhone in the virtual folder for the local computer.
        computer_folder = get_computer_shell_folder()
        dcim_processed = False
        for pidl in computer_folder:
            # If this is the iPhone, get the PIDL of its DCIM folder.
            dcim_pidl, parent, iphone_name = get_dcim_folder(pidl, computer_folder)
            if dcim_pidl is None:
                continue

            do_refresh_on_shell_view(dcim_pidl)

            self.walk_dcim_folder(dcim_pidl, parent)
            dcim_processed = True
            break

        if not dcim_processed:
            console("Unable to find 'My Computer/<iphone name>/Internal Storage/DCIM' folder. "
                    "Make sure you plugged in and unlock iPhone.")
            sys.exit(1)

        console("---")
        for dir_name, count in sorted(self.dir_stats.items(), key=lambda x: x[0]):
            console("%s: %d files" % (dir_name, count))
        console("---")
        console("new files: %d" % self.new_files)
        console("total files: %d" % self.total_files)

        return self.new_files

    def save_stream_to_file(self, stream, basename, expected_size, creation_time):
        self.total_files += 1
        full_path = os.path.join(self.target_dir, basename)

        if os.path.isfile(full_path):
            if os.path.getsize(full_path) == expected_size:
                change_file_creation_time(full_path, creation_time)
                return False
            else:
                os.unlink(full_path)
                # and rewrite

        bytes_read = 0
        with open(full_path, "wb") as fh:
            while True:
                buf = stream.read(BUFFER_SIZE)
                if not buf:
                    break
                fh.write(buf)
                bytes_read += len(buf)
        log.debug("Saved %d bytes to %s", bytes_read, full_path)
        if bytes_read != expected_size:
            console(u"ERROR: saved %d of %d expected bytes - file %s" % (bytes_read, expected_size, full_path))
            # TODO: delete file?
        else:
            sys.stdout.write(".")
            sys.stdout.flush()

        change_file_creation_time(full_path, creation_time)
        return True

    def walk_dcim_folder(self, dcim_pidl, parent):
        """
        Iterates all the subfolders of the iPhone's DCIM directory, gathering
        photos that need to be processed in photo_dict.

        :param dcim_pidl: A PIDL for the iPhone's DCIM folder
        :param parent: The parent folder of the PIDL
        """
        do_refresh_on_shell_view(dcim_pidl)

        dcim_folder = parent.BindToObject(dcim_pidl, None, shell.IID_IShellFolder2)
        for pidl in dcim_folder.EnumObjects(0, shellcon.SHCONTF_FOLDERS):
            folder_name = dcim_folder.GetDisplayNameOf(pidl, shellcon.SHGDN_NORMAL)
            console(u"processing folder: DCIM\\%s" % folder_name)

            do_refresh_on_shell_view(pidl)

            folder = dcim_folder.BindToObject(pidl, None, shell.IID_IShellFolder2)

            self.dir_stats[folder_name] = self.process_photos(folder)
            console("")

    def process_photos(self, folder):
        """
        Adds photos to photo_dict if they are newer than prev_index.
        :param folder: The PIDL of the folder to walk.
        """
        processed_count = 0
        new_count = 0
        for pidl in folder.EnumObjects(0, shellcon.SHCONTF_NONFOLDERS):
            name = folder.GetDisplayNameOf(pidl, shellcon.SHGDN_FORADDRESSBAR)
            size = folder.GetDetailsEx(pidl, PKEY_Size)

            created_utc = folder.GetDetailsEx(pidl, PKEY_DateCreated)

            stream = folder.BindToStorage(pidl, None, pythoncom.IID_IStream)
            is_new = self.save_stream_to_file(stream, os.path.basename(name), size, fix_timezone(created_utc))
            processed_count += 1

            if is_new:
                self.new_files += 1
                new_count += 1
                if new_count % 25 == 0:
                    sys.stdout.write(str(new_count))
                    sys.stdout.flush()

        return processed_count


def init_logging():
    log_file = sys.argv[0] + ".log"
    fh = codecs.open(log_file, "a+", "utf8")
    logging.basicConfig(stream=fh,
                        level=logging.DEBUG,
                        format='%(asctime)s %(levelname)s - %(message)s',
                        datefmt='%Y-%m-%d %H:%M:%S')


def main():
    init_logging()

    parser = ArgumentParser()
    parser.add_argument("target_dir", help="existing dir to save files")
    args = parser.parse_args()

    start_time = time.time()

    downloader = AppleFilesImporter(args.target_dir)
    downloader.run()

    # iterative brute-force scanning
    # prev_new_count = -1
    # while True:
    #     new_count = downloader.run()
    #     if new_count == prev_new_count:
    #         break
    #     prev_new_count = new_count

    console("Done. Elapsed %1.2fs." % (time.time() - start_time,))


if __name__ == '__main__':
    main()
