__author__ = 'leosamu@upv.es'
import os, tarfile, shutil, xlrd, datetime
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
CDATOSGENERALESCATEGORIAPOS = [5, 7]
CDATOSGENERALESEDICIONPOS = [1, 8]
CDATOSGENERALESDISPLAYNAMEPOS = [1, 1]
CDATOSGENERALESSTARTDATEPOS = [1, 4]
CDATOSGENERALESABOUTPOS = [5, 5]
CDATOSGENERALESPREREQUISITESPOS = [5, 4]
CDATOSGENERALESTEACHERSPOS = [5, 0]
CDATOSGENERALESEFFORTPOS = [5, 3]

"""
    sheet->unidad
"""
CUNIDADSHEET = "Unidades2"
CUNIDADCHAPTERIDCOL = 1
CUNIDADCHAPTERNAMECOL = 2
CUNIDADSUBSECTIONIDCOL = 3
CUNIDADSUBSECTIONNAMECOL = 4
CUNIDADFORMATCOL = 5
"""
    sheet->problemas
"""
CPROBLEMASSHEET = "Problemas2"
CPROBLEMASIDUNIDADCOL = 0
CPROBLEMASIDSUBSECCIONCOL = 1
CPROBLEMASIDLECCIONCOL = 2

"""
    sheet->curso
"""
CCURSOSHEET = "Curso2"
CCURSOCHAPTERIDCOL = 0
CCURSOSUBSECTIONIDCOL = 1
CCURSOLESSONIDCOL = 2
CCURSOLESSONDISPLAYNAMECOL = 3
CCURSOVIDEOCOL = 6
CCURSOOBJETIVOSCOL = 7
CCURSORESUMECOL = 8
CCURSOFORUMCOL = 9

#CCURSOPROBLEMINITIALCOL = 17
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
wb = xlrd.open_workbook(xlsmPath)


def generate_Edx():
    # CreateDirectoryTree ok
    # GenerateCourseIdFile ok
    # GenerateRoots ok
    # GenerateCourseMainFile ok
    # GenerateChapterFiles ok
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
    create_roots()
    create_course()
    make_tarfile()


def select_base_path():
    sheet = wb.sheet_by_name(CDATOSGENERALESSHEET)
    global path
    path = sheet.cell_value(CDATOSGENERALESNOMBREPOS[0], CDATOSGENERALESNOMBREPOS[1])


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
        if not os.path.exists(path + "/course"):
            os.makedirs(path + "/course")
        if not os.path.exists(path + "/problem"):
            os.makedirs(path + "/problem")
        if not os.path.exists(path + "/sequential"):
            os.makedirs(path + "/sequential")
        if not os.path.exists(path + "/vertical"):
            os.makedirs(path + "/vertical")
        if not os.path.exists(path + "/video"):
            os.makedirs(path + "/video")
        if not os.path.exists(path + "/policies"):
            os.makedirs(path + "/policies")
        if not os.path.exists(path + "/chapter"):
            os.makedirs(path + "/chapter")
        if not os.path.exists(path + "/roots"):
            os.makedirs(path + "/roots")
        if not os.path.exists(path + "/html"):
            os.makedirs(path + "/html")
        if not os.path.exists(path + "/about"):
            os.makedirs(path + "/about")
        if not os.path.exists(path + "/info"):
            os.makedirs(path + "/info")
        if not os.path.exists(path + "/discussion"):
            os.makedirs(path + "/discussion")


def create_course_id_file():
    """
    generates course.xml in the main dir
    """
    xmlfile = path + "/course.xml"

    sheet = wb.sheet_by_name(CDATOSGENERALESSHEET)

    courseCat = sheet.cell_value(CDATOSGENERALESCATEGORIAPOS[0], CDATOSGENERALESCATEGORIAPOS[1])
    courseID = sheet.cell_value(CDATOSGENERALESNOMBREPOS[0], CDATOSGENERALESNOMBREPOS[1]) + sheet.cell_value(
        CDATOSGENERALESEDICIONPOS[0], CDATOSGENERALESEDICIONPOS[1])
    courseName = sheet.cell_value(CDATOSGENERALESNOMBREPOS[0], CDATOSGENERALESNOMBREPOS[1])

    # Create the root element
    page = etree.Element('course', org=courseCat, course=courseName, url_name=courseID)
    # Make a new document tree
    doc = etree.ElementTree(page)

    # Save to XML file
    doc.write(xmlfile, pretty_print=True, xml_declaration=False, encoding='utf-8')


def create_roots():
    """
    generates the xml in roots for current course
    """
    sheet = wb.sheet_by_name(CDATOSGENERALESSHEET)

    courseCat = sheet.cell_value(CDATOSGENERALESCATEGORIAPOS[0], CDATOSGENERALESCATEGORIAPOS[1])
    courseID = sheet.cell_value(CDATOSGENERALESNOMBREPOS[0], CDATOSGENERALESNOMBREPOS[1]) + sheet.cell_value(
        CDATOSGENERALESEDICIONPOS[0], CDATOSGENERALESEDICIONPOS[1])
    courseName = sheet.cell_value(CDATOSGENERALESNOMBREPOS[0], CDATOSGENERALESNOMBREPOS[1])

    xmlfile = path + "/roots/" + courseID + ".xml"

    # Create the root element
    page = etree.Element('course', org=courseCat, course=courseName, url_name=courseID)
    # Make a new document tree
    doc = etree.ElementTree(page)

    # Save to XML file
    doc.write(xmlfile, pretty_print=True, xml_declaration=False, encoding='utf-8')


def create_course():
    """
    this generates the course main file wich
    contains a list of chapters for the course
    """
    sheet = wb.sheet_by_name(CDATOSGENERALESSHEET)
    sheetUnidad = wb.sheet_by_name(CUNIDADSHEET)

    courseID = sheet.cell_value(CDATOSGENERALESNOMBREPOS[0], CDATOSGENERALESNOMBREPOS[1]) + sheet.cell_value(
        CDATOSGENERALESEDICIONPOS[0], CDATOSGENERALESEDICIONPOS[1])
    courseDisplayName = sheet.cell_value(CDATOSGENERALESDISPLAYNAMEPOS[0], CDATOSGENERALESDISPLAYNAMEPOS[1])
    courseStartDate = sheet.cell_value(CDATOSGENERALESSTARTDATEPOS[0], CDATOSGENERALESSTARTDATEPOS[1])
    courseStartDate = datetime.datetime(*xlrd.xldate_as_tuple(courseStartDate, xlrd.Book.datemode))
    xmlfile = path + "/course/" + courseID + ".xml"

    # Create the root element
    page = etree.Element('course', display_name=courseDisplayName, start=str(courseStartDate))
    # Make a new document tree
    doc = etree.ElementTree(page)
    currentChapter = ""
    urlName = ""
    for row in range(1, sheetUnidad.nrows):
        if currentChapter != sheetUnidad.cell_value(row, CUNIDADCHAPTERIDCOL):
            currentChapter = sheetUnidad.cell_value(row, CUNIDADCHAPTERIDCOL)
            if (sheetUnidad.cell_value(row, CUNIDADCHAPTERIDCOL) != "final"):
                urlName = 'Unidad' + str(int(sheetUnidad.cell_value(row, CUNIDADCHAPTERIDCOL)))
                etree.SubElement(page, 'chapter',
                                 url_name=urlName)
            else:
                urlName = 'Final'
                etree.SubElement(page, 'chapter',
                                 url_name=urlName)
            create_chapter(row, urlName)
    # Save to XML file
    doc.write(xmlfile, pretty_print=True, xml_declaration=False, encoding='utf-8')


def create_chapter(_startRow, _urlName):
    """
    creates a chapter.xml file
    witch contains a list of sections
    """
    sheetUnidad = wb.sheet_by_name(CUNIDADSHEET)
    sheetProblems = wb.sheet_by_name(CPROBLEMASSHEET)
    currentChapter = sheetUnidad.cell_value(_startRow, CUNIDADCHAPTERIDCOL)
    chapterDisplayName = sheetUnidad.cell_value(_startRow, CUNIDADCHAPTERNAMECOL)
    strChapterID = _urlName
    if (currentChapter != "final"):
        strExamenID = "e" + str(int(currentChapter))
        xmlfile = path + "/chapter/" + strChapterID + '.xml'
    else:
        strExamenID = "final"
        xmlfile = path + "/chapter/Final.xml"

    # Create the root element
    page = etree.Element('chapter', display_name=chapterDisplayName)
    # Make a new document tree
    doc = etree.ElementTree(page)
    urlName = ""
    if (currentChapter != "final"):
        # Add normal childrens
        for row in range(_startRow, sheetUnidad.nrows):
            if currentChapter == sheetUnidad.cell_value(row, CUNIDADCHAPTERIDCOL):
                urlName = strChapterID + 'Subsection' + str(
                    int(sheetUnidad.cell_value(row, CUNIDADSUBSECTIONIDCOL))) + "sequential"
                etree.SubElement(page, 'sequential',
                                 url_name=urlName)
                createSequential(sheetUnidad.cell_value(row, CUNIDADCHAPTERIDCOL),
                                 sheetUnidad.cell_value(row, CUNIDADSUBSECTIONIDCOL),
                                 sheetUnidad.cell_value(row, CUNIDADCHAPTERNAMECOL),
                                 sheetUnidad.cell_value(row, CUNIDADSUBSECTIONNAMECOL), urlName,
                                 str(sheetUnidad.cell_value(row, CUNIDADFORMATCOL)), "", "")
            else:
                break

        # Add exam if exists
        for rowExam in range(1, sheetProblems.nrows):
            if strExamenID == sheetProblems.cell_value(rowExam, CPROBLEMASIDUNIDADCOL):
                urlName = strChapterID + 'Examensequential'
                etree.SubElement(page, 'sequential',
                                 url_name=urlName)
                #TO-DO call to createExam
                break
    else:
        urlName = 'FinalExamensequential'
        etree.SubElement(page, 'sequential',
                         url_name=urlName)
        #TO-DO call to createExam
    # Save to XML file
    doc.write(xmlfile, pretty_print=True, xml_declaration=False, encoding='utf-8')


def createSequential(_section, _subsection, _sectionDisplayName, _subsectionDisplayName, _urlName, _format, _startDate,
                     _endDate):
    """
    creates sequential files
    wich contains a list of vertical files for each lesson and the problems of that lesson
    """
    sheetCurso = wb.sheet_by_name(CCURSOSHEET)
    #sheetProblem = wb.sheet_by_name(CPROBLEMASSHEET)

    currentChapter = _section
    currentSubsection = _subsection
    strSubsectionID = _urlName
    xmlfile = path + "/sequential/" + strSubsectionID + ".xml"

    #TO-DO check if endDate goes to the end field
    page = etree.Element('sequential', display_name=_subsectionDisplayName, format=_format, start=_startDate,
                         end=_endDate)
    # Make a new document tree
    doc = etree.ElementTree(page)
    #Unidad1Subsection1Vertical1
    #TO-DO generate the range with a binary search
    for row in range(1, sheetCurso.nrows):
        if currentChapter == sheetCurso.cell_value(row,
                                                   CCURSOCHAPTERIDCOL) and currentSubsection == sheetCurso.cell_value(
                row, CCURSOSUBSECTIONIDCOL):
            urlName = "Unidad" + str(int(currentChapter)) + "Subsection" + str(
                int(currentSubsection)) + "Vertical" + str(int(sheetCurso.cell_value(row, CCURSOLESSONIDCOL)))
            etree.SubElement(page, 'vertical', url_name=urlName)
            createVertical(currentChapter, currentSubsection, sheetCurso.cell_value(row, CCURSOLESSONIDCOL), row,
                           urlName, _sectionDisplayName, _subsectionDisplayName)
            problemRow = findProblems(currentChapter, currentSubsection, sheetCurso.cell_value(row, CCURSOLESSONIDCOL))
            if problemRow > 0:
                urlName = urlName + "Problems"
                etree.SubElement(page, 'vertical', url_name=urlName)
                #TO-DO call to createProblemSet
                #createProblemSet(currentChapter,currentSubsection,sheetCurso.cell_value(row, CCURSOLESSONIDCOL,problemRow,urlName)
        if currentChapter < sheetCurso.cell_value(row, CCURSOCHAPTERIDCOL):
            break

    doc.write(xmlfile, pretty_print=True, xml_declaration=False, encoding='utf-8')

def createProblemSet(_Chapter, _Subsection, _Lesson, _row, _urlName):
    """
    creates the sets of problems related with a lesson
    <vertical display_name='Actividad  1'>
        <problem url_name='problem_1' />
    </vertical>
    """
    xmlfile = path + "/vertical/" + _urlName + ".xml"

def createVertical(_Chapter, _Subsection, _Lesson, _row, _urlName, _ChapterDisplayName, _SubsectionDisplayName):
    """
    creates the vertical files wich has links to every element in the vertical
    html Objetivos
    video Video
    html Resumen
    forumlink Foro
    """
    baseName = "Unidad" + str(int(_Chapter)) + "Subsection" + str(int(_Subsection)) + "Vertical" + str(int(_Lesson))
    sheetCurso = wb.sheet_by_name(CCURSOSHEET)
    sheetDatosGenerales = wb.sheet_by_name(CDATOSGENERALESSHEET)
    courseName = sheetDatosGenerales.cell_value(CDATOSGENERALESNOMBREPOS[0], CDATOSGENERALESNOMBREPOS[1])
    displayName = sheetCurso.cell_value(_row, CCURSOLESSONDISPLAYNAMECOL)
    xmlfile = path + "/vertical/" + _urlName + ".xml"
    page = etree.Element('vertical', display_name=displayName)
    # Make a new document tree
    doc = etree.ElementTree(page)

    if sheetCurso.cell_value(_row, CCURSOOBJETIVOSCOL) != "":
        #Unidad1Subsection1Vertical1Objetivos
        urlName = baseName + "Objetivos"
        etree.SubElement(page, 'html', url_name=urlName)
        generateHtml(urlName, sheetCurso.cell_value(_row, CCURSOOBJETIVOSCOL),displayName)
        displayName=""

    if sheetCurso.cell_value(_row, CCURSOVIDEOCOL) != "":
        #Unidad1Subsection1Vertical1Video
        urlName = baseName + "Video"
        etree.SubElement(page, 'video', url_name=urlName)
        generateVideo(urlName, sheetCurso.cell_value(_row, CCURSOVIDEOCOL),displayName)
        displayName=""

    if sheetCurso.cell_value(_row, CCURSORESUMECOL) != "":
        #Unidad1Subsection1Vertical1Resumen
        urlName = baseName + "Resumen"
        etree.SubElement(page, 'html', url_name=urlName)
        generateHtml(urlName, sheetCurso.cell_value(_row, CCURSORESUMECOL),displayName)
        displayName=""

    if sheetCurso.cell_value(_row, CCURSOFORUMCOL) != "":
        #Unidad1Subsection1Vertical1Discussion
        urlName = baseName + "Discussion"
        etree.SubElement(page, 'discussion', url_name=urlName)
        discussionCategory = "Tema " + str(int(_Chapter)) + ": " + _ChapterDisplayName
        discussionID = courseName + str(int(_Chapter)) + "_" + str(int(_Subsection))
        generateDiscussion(urlName, discussionCategory, _SubsectionDisplayName, discussionID,displayName)
        displayName=""

    doc.write(xmlfile, pretty_print=True, xml_declaration=False, encoding='utf-8')

def generateVideo(_urlName,_videoURL,_displayName):
    """
    generates the video xml file
    """
    xmlfile = path + "/video/" + _urlName + ".xml"
    page = etree.Element('video', youtube="1.00" + _videoURL, display_name=_displayName, youtube_id_1_0=_videoURL)
    # Make a new document tree
    doc = etree.ElementTree(page)

    doc.write(xmlfile, pretty_print=True, xml_declaration=False, encoding='utf-8')

def generateHtml(_urlName, _htmlText,_displayName):
    """
    generates the xml and html file wich will link the html
    into the course
    """
    xmlfile = path + "/html/" + _urlName + ".xml"
    htmlfile = path + "/html/" + _urlName + ".html"
    page = etree.Element('html', filename=_urlName, display_name=_displayName)
    # Make a new document tree
    doc = etree.ElementTree(page)

    doc.write(xmlfile, pretty_print=True, xml_declaration=False, encoding='utf-8')
    #TO-DO VALIDATE HTMLS
    html = open(htmlfile, "w")
    html.write(_htmlText.encode('utf8'))
    html.close()


def generateDiscussion(_urlName, _discussionCategory, _SubsectionDisplayName, _discussionID,_displayName):
    """
    generates the discussion file wich will link the discussion
    into the course
    """
    xmlfile = path + "/discussion/" + _urlName + ".xml"
    page = etree.Element('discussion', discussion_category=_discussionCategory,
                         discussion_target=_SubsectionDisplayName, discussion_id=_discussionID, display_name=_displayName)
    # Make a new document tree
    doc = etree.ElementTree(page)

    doc.write(xmlfile, pretty_print=True, xml_declaration=False, encoding='utf-8')


#needs to change the way we do the search
def findProblems(_chapter, _subSection, _lesson):
    sheetProblems = wb.sheet_by_name(CPROBLEMASSHEET)
    sheetProblems.nrows
    for row in range(1, sheetProblems.nrows):
        if _chapter == sheetProblems.cell_value(row, CPROBLEMASIDUNIDADCOL) and _subSection == sheetProblems.cell_value(
                row, CPROBLEMASIDSUBSECCIONCOL) and _lesson == sheetProblems.cell_value(row, CPROBLEMASIDLECCIONCOL):
            return row
        else:
            if _chapter < sheetProblems.cell_value(row, CPROBLEMASIDUNIDADCOL):
                return -1
    return -1


# def bisectSheetDuplicates(sheet,targetvalue,firstindex=0,lastindex=0):
#
#     size = firstindex-lastindex
#     if size ==0:
#         return -1
#     if size==1:
#         if sheet.cell_value(firstindex,0) == targetvalue:
#             return firstindex
#         else:
#             return -1
#     center = int(round(size/2))
#
#     if sheet.cell_value(center,0) == targetvalue:
#         #found it now moveback to the first appearance
#         while sheet.cell_value(center,0) == targetvalue:
#             center = center -1
#         return firstindex + center+1
#     if targetvalue > sheet.cell_value(center,0):
#         return bisectSheetDuplicates(sheet,targetvalue,center+1,lastindex)
#     else:
#         return bisectSheetDuplicates(sheet,targetvalue,firstindex,center-1)


def make_tarfile():
    """
    Packs all in a targz file ready to import.
    """
    with tarfile.open(path + '/' + path + '.tar.gz', 'w:gz') as tar:
        for f in os.listdir(path):
            tar.add(path + "/" + f, arcname=os.path.basename(f))
        tar.close()


generate_Edx()