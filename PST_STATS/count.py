import sys, stat, os
import win32com.client

def RecurseFolderCount(folder, pst_path, TotalCount, rdoSession, msgCount):
  TotalCount = TotalCount + folder.Items.Count
	#for item in folder.Items:	   
	for subFolder in folder.Folders:
		TotalCount = RecurseFolderCount(subFolder, pst_path, TotalCount, rdoSession, msgCount)
	return TotalCount

#dirs=os.listdir(".")
msgCount = 0
directory = "."
extension = ".pst"
pstpath = ''
TTemp = 0
OverAllSize=0
TotalCount = 0
OverAllCount = 0
mbSize = 0
gbSize=0

dirs=[file for file in os.listdir(directory) if file.endswith(extension)]
out = open('counts.txt', 'w') 
#dirs=[filename for filename in dirs if filename[0] != '.' or '*.py' or '*.txt']
print dirs

for d in dirs:
	try:
		size = os.path.getsize(d)
		#size /= 1024*1024.0
		#gsize /= 1024*1024*1024.0
		OverAllSize = size + OverAllSize
		rdoSession = win32com.client.Dispatch("Redemption.RDOSession")
		rdoSession.LogonPstStore(d)
		pstStore = rdoSession.Stores.DefaultStore
	except Exception as e:
		print e
	rootFolder = pstStore.IPMRootFolder
	print "working on " + d
	print "RootCount: " + str(rootFolder.Items.Count)
	TTemp = TotalCount
	TotalCount = 0
	TotalCount = RecurseFolderCount(rootFolder, pstStore.PstPath, TotalCount, rdoSession, msgCount)
	
	OverAllCount = TotalCount + TTemp
	gbSize = OverAllSize / (1024*1024*1024.0)
	mbSize = OverAllSize / (1024*1024.0)

	print "End Count: " + str(TotalCount)
	out.write(d +"\t"+str(TotalCount)+"\t"+str(size)+"\n")
	
out.write("\nTotal msg count: "+str(OverAllCount)+"\tTotal Data Size: "+str(OverAllSize)+"bytes\t"+str(mbSize)+"mb\t"+str(gbSize)+"gb")
	
	#print "We are here "+ d
	

#raw_input()
