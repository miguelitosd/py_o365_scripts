py_o365_scripts
===============

Small set of scripts I cobbled together after doing a bunch of google searches for how to delete/move emails via python in o365

## Requirements:
* A windows box with python and the win32com.client module installed
* outlook configured and running on the box against the exchange/o365 environment you want the scripts to work against.

## Files

* README.md
* daily.sh
    * Small shell script that calls the 2 python scripts and uses tee to redirect their output to files that are then parsed to give a summary, sample in script itself.
* email_archival.py
    * Main script I wrote to walk through emails with some logic:
    * Only messages at least 28 days old
    * Delete some mails based on a couple subject regexes
    * Delete some mails based on sender(s)
    * Emails with no sender are skipped
    * Delete based on a to regex
    * All other mails will be moved into a folder based on naming archives/$year/$month e.g. archives/2021/07
* purge.py
    * Script I wrote to go through emails between 24 and 96 hours old and purge/delete everything that matches based on some subject, to or from address, as long as it has been marked as read (downloaded)which is simply to purge the large amount of automated emails I get and don't need to keep copies around.
    * This one helps to trim down the usage while allowing my process to have pulled a copy via fetchmail already.
