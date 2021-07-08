from datetime import date,timedelta
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

def moveFolder(f1, f2, msg):
    year = str(f1)
    month = str(f2)

    root = outlook.Folders.Item("__MY_EMAIL__").Folders.Item("archives")
    try:
        yearFolder = outlook.Folders.Item("__MY_EMAIL__").Folders.Item("archives").Folders.Item(year)
        uprint("folder archives/%s exists" % year)
    except:
        uprint("Making folder archives/%s" % year)
        root.Folders.Add(year)

    try:
        newFolder = outlook.Folders.Item("__MY_EMAIL__").Folders.Item("archives").Folders.Item(year).Folders.Item(month)
        uprint("Folder archives/%s/%s exists" % (year,month))
    except:
        yearFolder = outlook.Folders.Item("__MY_EMAIL__").Folders.Item("archives").Folders.Item(year)
        uprint("Making Folder archives/%s/%s" % (year,month))
        yearFolder.Folders.Add(month)

    archiveFolder = outlook.Folders.Item("__MY_EMAIL__").Folders.Item("archives").Folders.Item(year).Folders.Item(month)

    uprint("archiveFolder name: %s" % archiveFolder.Name)
    if message.UnRead == True:
        message.UnRead = False

    uprint("Doing message.Move call")
    message.Move(archiveFolder)
    

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

#inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case,
#                                    # the inbox. You can change that number to reference
#                                    # any other folder

inbox = outlook.Folders.Item("__MY_EMAIL__").Folders.Item("Inbox")

days_ago    = 28
date_limit  = (date.today()-timedelta(days=days_ago)).isoformat()
counter     = 0
uprint("Calling inboxItems")
Filter      = "[ReceivedTime]<'" + date_limit + "'"
allMessages = inbox.Items
uprint("Applying filter: %s" % Filter)
messages = allMessages.Restrict(Filter)
totMsg = messages.Count
countDel = 0
countMoved = 0
uprint("Count: %d, enter loop" % totMsg)
start_time = time.time()
for message in list(messages):
    message = messages.GetLast()

    counter+=1
    uprint("==========================")

    if ( re.match("__SUB_REGEX__",message.Subject) ):
        countDel+=1
        uprint("Deleting (%d/%d) on Subject match... %s" % (countDel, totMsg, message.Subject))
        message.Delete()
        continue

    try:
        uprint("%s = %s" % (message.ReceivedTime,message.Subject))
    except:
        if (re.match("^(Recall:|\[Cloud Audit AWS Alert\] )", message.Subject)):
            uprint("Deleting Recall/AWS Alert message: %s" % message.Subject)
            countDel+=1
            message.Delete()
            continue
        else:
            uprint("No Time and not Recall: %s" % message.Subject)
            continue

    try:
        if (re.match("__SENDER_DEL_REGEX__",message.Sender.Address)):
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

    date = str(message.ReceivedTime)
    d = date.split("-")
    folder = "archives/%s/%s" % (d[0],d[1])
    year = int(d[0])
    month = int(d[1])

    countMoved+=1
    uprint("Moving email (%d/%d) to box: %s" % (countMoved, totMsg, folder))
    moveFolder(d[0],d[1],message)

end_time = time.time()
took = end_time - start_time
per = counter / took
uprint("DONE")
uprint("Archived: %d Deleted: %d Total: %d" % (countMoved, countDel, counter))
uprint("Took %d seconds for %.2f msgs/s" % (took, per))
