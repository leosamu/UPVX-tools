__author__ = 'leosamu@upv.es'
import os, tarfile, shutil , xlrd
from lxml import etree


"""
    Const declaration
    this consts will help us in case of modifications in the excel sheet
    TO-DO: readapt the consts to the new xls format.
"""

"""
    sheet->datosgenerales
"""
CDATOSGENERALESSHEET = "DatosGenerales2"
CDATOSGENERALESNOMBREPOS = [1, 0]
CDATOSGENERALESCATEGORIAPOS = [2, 8]
CDATOSGENERALESEDICIONPOS = [2, 9]
CDATOSGENERALESDISPLAYNAMEPOS = [2, 2]
CDATOSGENERALESSTARTDATEPOS = [2, 5]
CDATOSGENERALESABOUTPOS = [6, 6]
CDATOSGENERALESPREREQUISITESPOS = [6, 5]
CDATOSGENERALESTEACHERSPOS = [6, 1]
CDATOSGENERALESEFFORTPOS = [2, 7]

"""
    sheet->curso
"""
CCURSOSHEET = "curso"
CCURSOCHAPTERIDCOL = 1
CCURSOCHAPTERDISPLAYNAMECOL = 2
CCURSOSUBSECTIONIDCOL = 3
CCURSOSUBSECTIONDISPLAYNAMECOL = 4
CCURSOSUBSECTIONFORMATCOL = 5
CCURSOUNITIDCOL = 7
CCURSOUNITDISPLAYNAMECOL = 8
CCURSOVIDEOURLCOL = 11
CCURSOUNITTEXTCOL = 12
CCURSOUNITRESUMECOL = 13
CCURSOUNITFORUMCOL = 14
CCURSOINFOUPDATESPOS = [2, 10]
CCURSOABOUTVIDEOPOS = [2, 11]
#CCURSOPROBLEMINITIALCOL = 17
CCURSOCHAPTERINITIALROW = 3
#CCURSOMAXATTEMPTS AS INTEGER = 2 <.- THIS WILL GO TO DATOS GENERALES SHEET
#CCURSOSHOANSWER AS STRING = "FINISHED" <.- THIS WILL GO TO DATOS GENERALES SHEET
#CCURSOPROBLEMWEIGHT = 1 <.- THIS WILL GO TO PROBLEM SHEET
"""
    sheet->Examenes
"""
CEXAMENESSHEET = "Examenes"
CEXAMENESQUESTIONINITIALROW = 2
CEXAMENESCHAPTERIDCOL = 1
CEXAMENESQUESTIONCOL = 5
CEXAMENESANSWERCOL = 6
CEXAMENESRIGHTANSWERCOL = 10
CEXAMENESUNITIDCOL = 2
CEXAMENESUNITTITLECOL = 3
CEXAMENESPREVIOUSTEXTCOL = 4
#CEXAMENESMAXATEMPTS = 2
#CEXAMENESSHOWANSWER = "never"
#CEXAMENESPROBLEMWEIGHT = 1

# from xlrd import open_workbook,cellname, cellnameabs, colname
# wb = open_workbook('C:\Users\leosamu\Documents\XLS\MOOCExcelExperimental.xlsm')
# print cellname(0,0),cellname(0,10),cellname(100,100)
# print cellnameabs(3,1),cellnameabs(41,59),cellnameabs(265,358)
# print colname(0),colname(10),colname(100)

path = ""
xlsmPath = "C:\Users\leosamu\Documents\XLS\MOOCExcelExperimental.xlsm"

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
    select_base_path()
    clean()
    create_directory_tree()
    create_course_id_file()
    make_tarfile()

def select_base_path():
    wb = xlrd.open_workbook(xlsmPath)
    sheet = wb.sheet_by_name(CDATOSGENERALESSHEET)
    global path
    path = sheet.cell_value(CDATOSGENERALESNOMBREPOS[0],CDATOSGENERALESNOMBREPOS[1])


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

def create_course_id_file():
    """
    generates course.xml in the main dir

    Private Sub GenerateCourseIdFile()

        'courseCat and courseID may change the cell numbers due to the current
        'inexistence of that information in the xls
        courseCat = Sheets(cDatosGeneralesSheet).Cells(cDatosGeneralesCategoriaRow, cDatosGeneralesCategoriaCol)
        courseID = Sheets(cDatosGeneralesSheet).Cells(cDatosGeneralesNombreRow, cDatosGeneralesNombreCol) + Sheets(cDatosGeneralesSheet).Cells(cDatosGeneralesEdicionRow, cDatosGeneralesEdicionCol)
        courseName = Sheets(cDatosGeneralesSheet).Cells(cDatosGeneralesNombreRow, cDatosGeneralesNombreCol)

    """
    xmlfile = path + "/course.xml"
    wb = xlrd.open_workbook(xlsmPath)

    sheet = wb.sheet_by_name(CDATOSGENERALESSHEET)
    print (sheet.cell(3,4))
    # Create the root element

    courseCat = sheet.cell(CDATOSGENERALESCATEGORIAPOS[0],CDATOSGENERALESCATEGORIAPOS[1])
    courseID = sheet.cell(CDATOSGENERALESNOMBREPOS[0],CDATOSGENERALESNOMBREPOS[1])
    courseName = "asdfads"
    # Make a new document tree
    page = etree.Element('course', org=courseCat,course=courseName,url_name=courseID)
    doc = etree.ElementTree(page)

    # Add the subelements
  #  pageElement = etree.SubElement(page, 'Country',
   #                                   name='Germany',
    #                                  Code='DE',
     #                                 Storage='Basic')
    # For multiple multiple attributes, use as shown above

    # Save to XML file
    doc.write(xmlfile, pretty_print=True, xml_declaration=False, encoding='utf-8')

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


def make_tarfile():
    """
    Packs all in a targz file ready to import.
    """
    with tarfile.open(path + '/' + path + '.tar.gz', 'w:gz') as tar:
        for f in os.listdir(path):
            tar.add(path + "/" + f, arcname=os.path.basename(f))
        tar.close()




generate_Edx()