__author__ = 'leosamu@upv.es'
import os, tarfile, shutil




# from xlrd import open_workbook,cellname, cellnameabs, colname
# wb = open_workbook('C:\Users\leosamu\Documents\XLS\MOOCExcelExperimental.xlsm')
# print cellname(0,0),cellname(0,10),cellname(100,100)
# print cellnameabs(3,1),cellnameabs(41,59),cellnameabs(265,358)
# print colname(0),colname(10),colname(100)


# for s in wb.sheets():
#     print 'Sheet:',s.name
#     for row in range(s.nrows):
#         values = []
#         for col in range(s.ncols):
#             values.append(s.cell(row,col).value)
#         try:
#             print ','.join(values)
#         except:
#             print
#     print



path = "C:\Users\leosamu\Documents\edx\tools\carpetagenerada"
path = "carpetagenerada"

def generate_Edx():
        # CreateDirectoryTree ok
        # GenerateCourseIdFile
        # GenerateRoots
        # GenerateCourseMainFile
        # GenerateChapterFiles
        # GenerateChapterSequentialFiles
        # GenerateExamSequentialFiles
        # GenerateExamVerticalFiles
        # GenerateAbout
        # GenerateInfo
        # GeneratePolicies
        # GenerateTARGZ ok
    clean()
    create_directory_tree()
    make_tarfile()

def clean():
    """
    Deletes the directory from previous generations
    """
    try:
        shutil.rmtree(path)
    except:
        print "cannot remove " + path


def create_directory_tree():
    """
    Generates the directory structure needed for our xml project
    """
    print os.getcwd()
    if not os.path.exists(path):
        os.makedirs(path)
        if not os.path.exists(path+"/course"):
            os.makedirs(path+"/course")
        if not os.path.exists(path+"/problem"):
            os.makedirs(path+"/problem")
        if not os.path.exists(path+"/sequential"):
            os.makedirs(path+"/sequential")
        if not os.path.exists(path+"/vertical"):
            os.makedirs(path+"/vertical")
        if not os.path.exists(path+"/video"):
            os.makedirs(path+"/video")
        if not os.path.exists(path+"/policies"):
            os.makedirs(path+"/policies")
        if not os.path.exists(path+"/chapter"):
            os.makedirs(path+"/chapter")
        if not os.path.exists(path+"/roots"):
            os.makedirs(path+"/roots")
        if not os.path.exists(path+"/html"):
            os.makedirs(path+"/html")
        if not os.path.exists(path+"/about"):
            os.makedirs(path+"/about")
        if not os.path.exists(path+"/info"):
            os.makedirs(path+"/info")
        if not os.path.exists(path+"/discussion"):
            os.makedirs(path+"/discussion")

def make_tarfile():
    """
    Packs all in a targz file ready to import.
    """
    with tarfile.open(path + '/' + path + '.tar.gz', 'w:gz') as tar:
        for f in os.listdir(path):
            tar.add(path + "/" + f, arcname=os.path.basename(f))
        tar.close()




generate_Edx()