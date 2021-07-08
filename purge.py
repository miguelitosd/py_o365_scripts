import datetime
import pprint
import re
import sys
import time
import win32com.client

def uprint(*objects, sep=' ', end='\n', file=sys.stdout):
    enc = file.encoding
    if enc == 'UTF-8':
        print(*objects, sep=sep, end=end, file=file)
    else:
        f = lambda obj: str(obj).encode(enc, errors='backslashreplace').decode(enc)
        print(*map(f, objects), sep=sep, end=end, file=file)

start_time = time.time()
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

#inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case,
#                                    # the inbox. You can change that number to reference
#                                    # any other folder

inbox = outlook.Folders.Item("__YOUR_EMAIL__").Folders.Item("Inbox")

hours_ago_new    = 24
hours_ago_old    = 96
date_newest  = format((datetime.datetime.today()-datetime.timedelta(hours=hours_ago_new)).strftime("%Y-%m-%d %H:%M:%S"))
date_oldest  = format((datetime.datetime.today()-datetime.timedelta(hours=hours_ago_old)).strftime("%Y-%m-%d %H:%M:%S"))

counter     = 0
uprint("Calling inboxItems")
Filter = ("@SQL=(\"urn:schemas:httpmail:datereceived\" >= '" + date_oldest + "' AND \"urn:schemas:httpmail:datereceived\" <= '" + date_newest + "')")
allMessages = inbox.Items
uprint("Applying filter: %s" % Filter)
messages = allMessages.Restrict(Filter)
totMsg = messages.Count
countDel = 0
countSkipped = 0
uprint("Count: %d, enter loop" % totMsg)
didone=0
for message in list(messages):
    if (counter==0 or didone==0):
        message = messages.GetLast()
    else:
        message = messages.GetPrevious()

    counter+=1
    uprint("==========================")

    #uprint("Sender = %s" % message.Sender.Address)
    #continue

    try:
        if (message.UnRead == True):
            uprint("Skipping unread (%d/%d)" % (counter, totMsg))
            didone=1
            continue
    except:
        uprint("Message: %s" % message)
        break


    if ( re.match("__SUB_DEL_REGEX__",message.Subject) ):
        countDel+=1
        uprint("Deleting (%d/%d) on Subject match... %s" % (countDel, totMsg, message.Subject))
        message.Delete()
        continue

    didone=1
    try:
        uprint("%s = %s" % (message.ReceivedTime,message.Subject))
    except:
        uprint("NO TIME = %s" % (message.Subject))
        continue

    try:
        if (re.match("^__SENDER__",message.Sender.Address)):
            uprint("Deleting on Sender match... %s = %s" % (message.Sender.Address,message.Subject))
            countDel+=1
            message.Delete()
            continue
    except:
        uprint("No Sender.Address = %s" % message.Subject)

    try:
        emailTo = message.To
    except:
        emailTo = "UNKNOWN"

    
    if (re.match("__TO_DEL_REGEX__",emailTo)):
        countDel+=1
        uprint("Deleting (%d/%d) on To match... %s = %s" % (countDel, totMsg, emailTo,message.Subject))
        message.Delete()
        continue
    else:
        uprint("No To Address, Sub = %s" % message.Subject)

    countSkipped+=1
    uprint("Skip (%d/%d)" % (countSkipped, totMsg))

endTime = time.time()
spent = endTime - start_time
perTime = counter / spent
uprint("DONE")
uprint("Deleted: %d Skipped: %s Total: %d" % (countDel, countSkipped, counter))
uprint("Ran for %d seconds for ~ %.2f msgs/s" % (spent, perTime))
