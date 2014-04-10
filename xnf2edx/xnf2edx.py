
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
CDATOSGENERALESSHEET = "DatosGenerales"
CDATOSGENERALESNOMBREPOS = [1, 0]
CDATOSGENERALESCATEGORIAPOS = [7, 1]
CDATOSGENERALESEDICIONPOS = [7, 2]
CDATOSGENERALESDISPLAYNAMEPOS = [1, 1]
CDATOSGENERALESSTARTDATEPOS = [1, 4]
CDATOSGENERALESENDDATEPOS = [1, 5]
CDATOSGENERALESINFOPOS=[5,5]
CDATOSGENERALESABOUTPOS = [5, 4]
CDATOSGENERALESPREREQUISITESPOS = [5, 3]
CDATOSGENERALESTEACHERSPOS = [5, 0]
CDATOSGENERALESEFFORTPOS = [5, 2]
CDATOSGENERALESABOUTVIDEOPOS=[5, 6]


"""
    sheet->unidad
"""
CUNIDADSHEET = "Unidades"
CUNIDADTYPECOL = 0
CUNIDADCHAPTERIDCOL = 1
CUNIDADCHAPTERNAMECOL = 2
CUNIDADSUBSECTIONIDCOL = 3
CUNIDADSUBSECTIONNAMECOL = 4
CUNIDADFORMATCOL = 5
CUNIDADSTARTDATECOL = 6
CUNIDADENDDATECOL = 7

"""
    sheet->problemas
"""
CPROBLEMASSHEET = "Problemas"
CPROBLEMASIDUNIDADCOL = 0
CPROBLEMASIDSUBSECCIONCOL = 1
CPROBLEMASIDLECCIONCOL = 2
CPROBLEMASPREVIACOL = 4
CPROBLEMASTIPOCOL = 5
CPROBLEMASENUNCIADOCOL = 6
CPROBLEMASCOMENTARIOCOL = 7
CPROBLEMASCORRECTACOL = 8
CPROBLEMASRESPUESTACOL = 9

"""
    sheet->leccion
"""
CCURSOSHEET = "Leccion"
CCURSOCHAPTERIDCOL = 0
CCURSOSUBSECTIONIDCOL = 1
CCURSOLESSONIDCOL = 2
CCURSOLESSONDISPLAYNAMECOL = 4
CCURSOVIDEOCOL = 7
CCURSOOBJETIVOSCOL = 8
CCURSORESUMECOL = 9
CCURSOFORUMCOL = 10

#CCURSOPROBLEMINITIALCOL = 17
#CCURSOMAXATTEMPTS AS INTEGER = 2 <.- THIS WILL GO TO DATOS GENERALES SHEET
#CCURSOSHOANSWER AS STRING = "FINISHED" <.- THIS WILL GO TO DATOS GENERALES SHEET
#CCURSOPROBLEMWEIGHT = 1 <.- THIS WILL GO TO PROBLEM SHEET
"""
    sheet->Format
"""
CFORMATSHEET = "TipodeTarea"
CFORMATNAMECOL = 0
CFORMATABBREVIATIONCOL = 1
CFORMATWEIGHTCOL = 2
CFORMATDROPABLECOL = 3
CFORMATMAXATEMPTSCOL = 4
CFORMATSHOWANSWERCOL = 5
CFORMATPROBLEMWEIGHTCOL = 6
CFORMATRANDOMIZECOL = 7



problemSetID = 1
path = ""

"""
hardcoded xlsmpath must change to a parameter
"""
xlsmPath = "XNF.xlsm"
wb = xlrd.open_workbook(xlsmPath)


def generate_Edx():
    '''
    Main script makes the calls in order to clean the resulting thir and after that generate that dir and the targz
    that we will use to import the course
    '''
    select_base_path()
    clean()
    create_directory_tree()
    create_course_id_file()
    create_roots()
    create_course()
    create_about()
    create_info()
    create_policies()
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


def create_about():
    """
    creates the about files
    """
    create_about_video()
    create_about_overview()


def create_about_video():
    """
    creates the about video file
    """
    sheet = wb.sheet_by_name(CDATOSGENERALESSHEET)
    htmlfile = path + "/about/video.html"
    page = etree.Element('iframe', width='560',
                         height='315', src='//www.youtube.com/embed/' + sheet.cell_value(CDATOSGENERALESABOUTVIDEOPOS[0],CDATOSGENERALESABOUTVIDEOPOS[1]) +'?autoplay=0&rel=0',
                         frameborder='0', allowfullscreen='')
    # Make a new document tree
    doc = etree.ElementTree(page)

    doc.write(htmlfile, pretty_print=True, xml_declaration=False, encoding='utf-8')



def create_about_overview():
    """
    generates the overview file
    """
    sheet = wb.sheet_by_name(CDATOSGENERALESSHEET)
    htmlpath = path + "/about/overview.html"
    about = sheet.cell_value(CDATOSGENERALESABOUTPOS[0],CDATOSGENERALESABOUTPOS[1])
    if about[:1] != "<":
            about = "<p>" + about + "</p>"
    prerequisites = sheet.cell_value(CDATOSGENERALESPREREQUISITESPOS[0],CDATOSGENERALESPREREQUISITESPOS[1])
    if prerequisites[:1] != "<":
            prerequisites = "<p>" + prerequisites + "</p>"

    aboutroot = etree.Element('section', Class='about')
    about_writer = etree.ElementTree(aboutroot)
    aboutroot.append(etree.fromstring('<h2>Acerca de este curso</h2>\n'))
    aboutroot.append(etree.fromstring(about))

    prerequisiteroot = etree.Element('section', Class='prerequisites')
    prerequisitewriter = etree.ElementTree(prerequisiteroot)
    prerequisiteroot.append(etree.fromstring('<h2>Prerrequisitos</h2>\n'))
    prerequisiteroot.append(etree.fromstring(prerequisites))

    coursestaffroot = etree.Element('section', Class='course-staff')
    coursestaffwriter = etree.ElementTree(coursestaffroot)
    coursestaffroot.append(etree.fromstring('<h2>Profesores del curso</h2>\n'))
    teacherRow = CDATOSGENERALESTEACHERSPOS[0]
    while sheet.nrows > teacherRow and sheet.cell_value(teacherRow, CDATOSGENERALESTEACHERSPOS[1]) != "":
        article = etree.SubElement(coursestaffroot, 'article', Class='teacher')
        div = etree.SubElement(article,'div', Class='teacher-image')
        etree.SubElement(div, 'img', src='', align='left', style='margin:0 20 px 0 ')
        teacherDescription = sheet.cell_value(teacherRow, CDATOSGENERALESTEACHERSPOS[1])
        if teacherDescription[:1] != "<":
            teacherDescription = "<p>" + teacherDescription + "</p>"
        article.append(etree.fromstring(teacherDescription))
        teacherRow += 2

    about_writer.write(htmlpath, pretty_print=True, xml_declaration=False, encoding='utf-8')
    htmlfile = open(htmlpath, 'a')
    prerequisitewriter.write(htmlfile, pretty_print=True, xml_declaration=False, encoding='utf-8')
    coursestaffwriter.write(htmlfile, pretty_print=True, xml_declaration=False, encoding='utf-8')
    htmlfile.close()


def create_policies():
    """
    this creates the policies file
    """


def create_info():
    """
    this creates the info file
    """
    create_info_updates()
    create_info_handouts()

def create_info_updates():
    """
    creates info updates
    """
    sheet = wb.sheet_by_name(CDATOSGENERALESSHEET)
    htmlfile = path + "/info/updates.html"
    info = sheet.cell_value(CDATOSGENERALESINFOPOS[0],CDATOSGENERALESINFOPOS[1])
    if info != "":
        if info[:1] != "<":
            info = "<p>" + info + "</p>"
        page = etree.fromstring(info)
        doc = etree.ElementTree(page)

        doc.write(htmlfile, pretty_print=True, xml_declaration=False, encoding='utf-8')

def create_info_handouts():
    """
    creates info handouts
    """
    print "blas"


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
            urlName = 'Unidad' + str(int(sheetUnidad.cell_value(row, CUNIDADCHAPTERIDCOL)))
            etree.SubElement(page, 'chapter',url_name=urlName)
            create_chapter(row, urlName)
    # Save to XML file
    doc.write(xmlfile, pretty_print=True, xml_declaration=False, encoding='utf-8')


def create_chapter(_startRow, _urlName):
    """
    creates a chapter.xml file
    witch contains a list of sections
    :param _startRow:
    :param _urlName:
    """
    sheetUnidad = wb.sheet_by_name(CUNIDADSHEET)
    sheetProblems = wb.sheet_by_name(CPROBLEMASSHEET)
    currentChapter = sheetUnidad.cell_value(_startRow, CUNIDADCHAPTERIDCOL)
    chapterDisplayName = sheetUnidad.cell_value(_startRow, CUNIDADCHAPTERNAMECOL)
    strChapterID = _urlName
    #this will serve as counter for the problems in the chapter
    global problemSetID
    problemSetID = 1
    xmlfile = path + "/chapter/" + strChapterID + '.xml'

    # Create the root element
    page = etree.Element('chapter', display_name=chapterDisplayName)
    # Make a new document tree
    doc = etree.ElementTree(page)
    urlName = ""
    # Add normal childrens
    for row in range(_startRow, sheetUnidad.nrows):
        if currentChapter == sheetUnidad.cell_value(row, CUNIDADCHAPTERIDCOL):
            urlName = strChapterID + 'Subsection' + str(int(sheetUnidad.cell_value(row, CUNIDADSUBSECTIONIDCOL))) + "Sequential"
            etree.SubElement(page, 'sequential',url_name=urlName)
            createSequential(sheetUnidad.cell_value(row, CUNIDADTYPECOL),
                                 sheetUnidad.cell_value(row, CUNIDADCHAPTERIDCOL),
                                 sheetUnidad.cell_value(row, CUNIDADSUBSECTIONIDCOL),
                                 sheetUnidad.cell_value(row, CUNIDADCHAPTERNAMECOL),
                                 sheetUnidad.cell_value(row, CUNIDADSUBSECTIONNAMECOL), urlName,
                                 str(sheetUnidad.cell_value(row, CUNIDADFORMATCOL)), "", "")
        else:
            break


    # Save to XML file
    doc.write(xmlfile, pretty_print=True, xml_declaration=False, encoding='utf-8')


def createSequential(_type, _section, _subsection, _sectionDisplayName, _subsectionDisplayName, _urlName, _format, _startDate,
                     _endDate):
    """
    creates sequential files
    wich contains a list of vertical files for each lesson and the problems of that lesson
    :param _type:
    :param _section:
    :param _subsection:
    :param _sectionDisplayName:
    :param _subsectionDisplayName:
    :param _urlName:
    :param _format:
    :param _startDate:
    :param _endDate:
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

            urlName = "Unidad" + str(int(currentChapter)) + "Subsection" + str(int(currentSubsection)) + "Vertical" + str(int(sheetCurso.cell_value(row, CCURSOLESSONIDCOL)))
            if sheetCurso.cell_value(row, CCURSOOBJETIVOSCOL) != "" or sheetCurso.cell_value(row,CCURSOVIDEOCOL) != "" or sheetCurso.cell_value(row, CCURSORESUMECOL) != "" or sheetCurso.cell_value(row, CCURSOFORUMCOL) != "":
                etree.SubElement(page, 'vertical', url_name=urlName)
                createVertical(currentChapter, currentSubsection, sheetCurso.cell_value(row, CCURSOLESSONIDCOL), row, urlName, _sectionDisplayName, _subsectionDisplayName)

            problemRow = findProblems(currentChapter, currentSubsection, sheetCurso.cell_value(row, CCURSOLESSONIDCOL))
            if problemRow > 0:
                urlName += "Problems"
                if _type == "A":
                    if sheetCurso.cell_value(row, CCURSOLESSONDISPLAYNAMECOL) != "":
                        displayName = sheetCurso.cell_value(row, CCURSOLESSONDISPLAYNAMECOL)
                    else:
                        displayName = "Examen"
                else:
                    displayName = "Actividad " + str(problemSetID)
                etree.SubElement(page, 'vertical', url_name=urlName)
                createProblemSet(currentChapter, currentSubsection, sheetCurso.cell_value(row, CCURSOLESSONIDCOL),
                                 problemRow, urlName,displayName)

        if currentChapter < sheetCurso.cell_value(row, CCURSOCHAPTERIDCOL):
            break

    doc.write(xmlfile, pretty_print=True, xml_declaration=False, encoding='utf-8')


def createProblemSet(_Chapter, _Subsection, _Lesson, _row, _urlName, _displayName):
    """
    creates the sets of problems related with a lesson
    :param _Chapter:
    :param _Subsection:
    :param _Lesson:
    :param _row:
    :param _urlName:
    """
    global problemSetID
    displayName = _displayName
    xmlfile = path + "/vertical/" + _urlName + ".xml"
    problemID = 1
    page = etree.Element('vertical', display_name=displayName)
    # Make a new document tree
    doc = etree.ElementTree(page)
    sheetProblem = wb.sheet_by_name(CPROBLEMASSHEET)
    for row in range(_row, sheetProblem.nrows):
        if _Chapter == sheetProblem.cell_value(row, CPROBLEMASIDUNIDADCOL) \
                and _Subsection == sheetProblem.cell_value(row, CPROBLEMASIDSUBSECCIONCOL) \
                and _Lesson == sheetProblem.cell_value(row, CPROBLEMASIDLECCIONCOL):
            urlName = _urlName + str(problemID)
            #if the problem has a previa add an html element
            if sheetProblem.cell_value(row, CPROBLEMASPREVIACOL) != "":
                etree.SubElement(page, 'html', url_name=urlName + "Previa")
                createHtml(urlName + "Previa", sheetProblem.cell_value(row, CPROBLEMASPREVIACOL), displayName)
                displayName = ""
            etree.SubElement(page, 'problem', url_name=urlName)
            #call generate problem
            createProblem(displayName, row, urlName)
            displayName = ""  #due to platform issues only the first problem on a problemSet will have displayName
            problemID += 1


        else:
            break

    problemSetID += 1
    doc.write(xmlfile, pretty_print=True, xml_declaration=False, encoding='utf-8')


def createProblem(_displayName, _row, _urlName):
    """
    creates a problem object xml
    :param _displayName:
    :param _row:
    :param _urlName:
    """
    sheetProblem = wb.sheet_by_name(CPROBLEMASSHEET)
    xmlfile = path + "/problem/" + _urlName + ".xml"
    max_attempts = 2
    showanswer = "finished"
    weight = 1
    type = sheetProblem.cell_value(_row, CPROBLEMASTIPOCOL)
    nounce = sheetProblem.cell_value(_row, CPROBLEMASENUNCIADOCOL)

    if type == "Custom":
        page = etree.fromstring(nounce)
        # Make a new document tree
        doc = etree.ElementTree(page)
    else:
        #if nounce[:1] != "<":
        nounce = "<p>" + nounce + "</p>"



        comentary = sheetProblem.cell_value(_row, CPROBLEMASCOMENTARIOCOL)
        if comentary != "":
            comentary = "<div class='detailed-solution'>" + comentary + "</div>"

        #markdown='null' max_attempts='2' showanswer='finished' weight='1' display_name='Actividad  3'
        page = etree.Element('problem', display_name=_displayName, markdown="null", max_attempts=str(max_attempts),
                             showanswer=showanswer, weight=str(weight))
        # Make a new document tree
        doc = etree.ElementTree(page)
        #add the nounce of the problem
        page.append(etree.fromstring(nounce))
        #switch(type):
        #    case "Multichoice":
        #    break
        if type == "MultiChoice":
            problemMultiChoice(page, _row)
        elif type == "CheckBox":
            problemCheckBox(page, _row)
        elif type == "NumericalInput":
            problemNumerical(page, _row)
        elif type == "TextInput":
            problemText(page, _row)
        else:
            print "Error"
            exit()



        #add the solution (unique comentary)
        if comentary != "":
            solution = etree.SubElement(page, 'solution')
            solution.append(etree.fromstring(comentary))

    doc.write(xmlfile, pretty_print=True, xml_declaration=False, encoding='utf-8')


def problemText(_page, _row):
    """
    add the text input response box
    <stringresponse answer="Michigan" type="ci" >
    <textline size="20"/>
    </stringresponse>
    :param _page:
    :param _row:
    """
    sheetProblem = wb.sheet_by_name(CPROBLEMASSHEET)
    tolerance = sheetProblem.cell_value(_row, CPROBLEMASCORRECTACOL)
    answerCol = CPROBLEMASRESPUESTACOL

    root = etree.SubElement(_page, 'stringresponse', answer=unicode(sheetProblem.cell_value(_row, answerCol)))
    etree.SubElement(root, 'textline', size=unicode(unicode(sheetProblem.cell_value(_row, answerCol)).__sizeof__()))


def problemNumerical(_page, _row):
    """
    add the numerical input response box
    <numericalresponse answer="3.14159">
    <responseparam type="tolerance" default=".02" />
    <formulaequationinput />
    </numericalresponse>
    :param _page:
    :param _row:
    """
    sheetProblem = wb.sheet_by_name(CPROBLEMASSHEET)
    tolerance = sheetProblem.cell_value(_row, CPROBLEMASCORRECTACOL)
    answerCol = CPROBLEMASRESPUESTACOL

    root = etree.SubElement(_page, 'numericalresponse', answer=unicode(sheetProblem.cell_value(_row, answerCol)))
    etree.SubElement(root, 'responseparam', type="tolerance", default=unicode(tolerance))
    etree.SubElement(root, 'formulaequationinput')


def problemMultiChoice(_page, _row):
    """
    adds the options in a multichoice problem type
    :param _page:
    :param _row:
    """
    sheetProblem = wb.sheet_by_name(CPROBLEMASSHEET)
    rightAnswer = sheetProblem.cell_value(_row, CPROBLEMASCORRECTACOL)
    answerCol = CPROBLEMASRESPUESTACOL
    currentAnswer = 1
    root = etree.SubElement(_page, 'multiplechoiceresponse')
    choicegroup = etree.SubElement(root, 'choicegroup', type="MultipleChoice")
    while sheetProblem.ncols > answerCol and sheetProblem.cell_value(_row, answerCol) != "":
        try:
            choice = etree.SubElement(choicegroup, 'choice', correct=str((rightAnswer == currentAnswer)))
            choice.text = unicode(sheetProblem.cell_value(_row, answerCol))
            #choice.append(etree.fromstring(sheetProblem.cell_value(_row, answerCol)))
            currentAnswer += 1
            answerCol += 2
        except:
            print "error adding option", sys.exc_info()[0]
            raise


def problemCheckBox(_page, _row):
    """
    adds the options in a checkbox problem type
    :param _page:
    :param _row:
    """

    sheetProblem = wb.sheet_by_name(CPROBLEMASSHEET)
    rightAnswer = sheetProblem.cell_value(_row, CPROBLEMASCORRECTACOL)
    #comma separated
    rightAnswer = rightAnswer.split(";")
    answerCol = CPROBLEMASRESPUESTACOL
    currentAnswer = 1
    root = etree.SubElement(_page, 'choiceresponse')
    choicegroup = etree.SubElement(root, 'checkboxgroup', direction="Vertical")
    while sheetProblem.ncols > answerCol and sheetProblem.cell_value(_row, answerCol) != "":
        try:
            choice = etree.SubElement(choicegroup, 'choice',
                                      correct=str(unicode(currentAnswer) in map(unicode.lower, rightAnswer)))
            choice.text = unicode(sheetProblem.cell_value(_row, answerCol))
            #choice.append(etree.fromstring(sheetProblem.cell_value(_row, answerCol)))
            currentAnswer += 1
            answerCol += 2
        except:
            print "error adding option", sys.exc_info()[0]
            raise


def createVertical(_Chapter, _Subsection, _Lesson, _row, _urlName, _ChapterDisplayName, _SubsectionDisplayName):
    """
    creates the vertical files wich has links to every element in the vertical
    html Objetivos
    video Video
    html Resumen
    forumlink Foro
    :param _Chapter:
    :param _Subsection:
    :param _Lesson:
    :param _row:
    :param _urlName:
    :param _ChapterDisplayName:
    :param _SubsectionDisplayName:
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
        createHtml(urlName, sheetCurso.cell_value(_row, CCURSOOBJETIVOSCOL), displayName)
        displayName = ""

    if sheetCurso.cell_value(_row, CCURSOVIDEOCOL) != "":
        #Unidad1Subsection1Vertical1Video
        urlName = baseName + "Video"
        etree.SubElement(page, 'video', url_name=urlName)
        createVideo(urlName, sheetCurso.cell_value(_row, CCURSOVIDEOCOL), displayName)
        displayName = ""

    if sheetCurso.cell_value(_row, CCURSORESUMECOL) != "":
        #Unidad1Subsection1Vertical1Resumen
        urlName = baseName + "Resumen"
        etree.SubElement(page, 'html', url_name=urlName)
        createHtml(urlName, sheetCurso.cell_value(_row, CCURSORESUMECOL), displayName)
        displayName = ""

    if sheetCurso.cell_value(_row, CCURSOFORUMCOL) != "":
        #Unidad1Subsection1Vertical1Discussion
        urlName = baseName + "Discussion"
        etree.SubElement(page, 'discussion', url_name=urlName)
        discussionCategory = "Tema " + str(int(_Chapter)) + ": " + _ChapterDisplayName
        discussionID = courseName + str(int(_Chapter)) + "_" + str(int(_Subsection))
        createDiscussion(urlName, discussionCategory, _SubsectionDisplayName, discussionID, displayName)
        displayName = ""

    doc.write(xmlfile, pretty_print=True, xml_declaration=False, encoding='utf-8')


def createVideo(_urlName, _videoURL, _displayName):
    """
    generates the video xml file
    :param _urlName:
    :param _videoURL:
    :param _displayName:
    """
    xmlfile = path + "/video/" + _urlName + ".xml"
    page = etree.Element('video', youtube="1.00" + _videoURL, display_name=_displayName, youtube_id_1_0=_videoURL)
    # Make a new document tree
    doc = etree.ElementTree(page)

    doc.write(xmlfile, pretty_print=True, xml_declaration=False, encoding='utf-8')


def createHtml(_urlName, _htmlText, _displayName):
    """
    generates the xml and html file wich will link the html
    into the course
    :param _urlName:
    :param _htmlText:
    :param _displayName:
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


def createDiscussion(_urlName, _discussionCategory, _SubsectionDisplayName, _discussionID, _displayName):
    """
    generates the discussion file wich will link the discussion
    into the course
    :param _urlName:
    :param _discussionCategory:
    :param _SubsectionDisplayName:
    :param _discussionID:
    :param _displayName:
    """
    xmlfile = path + "/discussion/" + _urlName + ".xml"
    page = etree.Element('discussion', discussion_category=_discussionCategory,
                         discussion_target=_SubsectionDisplayName, discussion_id=_discussionID,
                         display_name=_displayName)
    # Make a new document tree
    doc = etree.ElementTree(page)

    doc.write(xmlfile, pretty_print=True, xml_declaration=False, encoding='utf-8')


def findProblems(_chapter, _subSection, _lesson):
    """
    :param _chapter:
    :param _subSection:
    :param _lesson:
    :return:
    """
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