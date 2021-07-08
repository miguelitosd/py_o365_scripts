cd ~/Downloads/
d=$(date +%m%d)
py ./purge.py 2>&1 | tee out.purge.$d
py ./email_archival.py 2>&1 | tee out.archive.$d
tail -n 2 out.purge.$d out.archive.$d | perl -e 'my $tot = 0; while (<>) { print $_; if (/Archived:\s+(\d+)\s+/) { $tot+=$1; } if (/Deleted:\s+(\d+)\s+/) { $tot+=$1};} print "\nTotal Deleted/Archived: $tot\n";'

#==> out.purge.0420 <==
#No To Address, Sub = RE: Forno testing (Bakeoff 1 AMD Rome/Skylake Refresh (CL) 
#Skip (180/453)
#DONE
#Deleted: 273 Skipped: 180 Total: 453
#Ran for 18 seconds for ~ 24.55 msgs/s
#
#==> out.archive.0420 <==
#archiveFolder name: 03
#Doing message.Move call
#DONE
#Archived: 135 Deleted: 4 Total: 139
#Took 23 seconds for 5.96 msgs/s
