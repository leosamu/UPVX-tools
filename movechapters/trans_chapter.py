# coding=utf-8
__author__ = u"Leonardo Salom Muñoz"
__credits__ = u"Leonardo Salom Muñoz"
__version__ = u"0.0.1-SNAPSHOT"
__maintainer__ = u"Leonardo Salom Muñoz"
__email__ = u"leosamu@upv.es"
__status__ = u"Development"

import pymongo
import bson.son
import sys

srcconnection = pymongo.Connection("mongodb://" + sys.argv[1], safe=True)
dstconnection = pymongo.Connection("mongodb://" + sys.argv[4], safe=True)
srcdb = srcconnection.edxapp
dstdb = dstconnection.edxapp
srcmodulestore = srcdb.modulestore
dstmodulestore = dstdb.modulestore
#modulestore_location = db['modulestore.location_map']

#clones a chapter _chapter from the _src course to the _dest course
#course names comes in the format [org]/[course]/[name]
#for example: Matematicas/cursografos/2014-001
#chapter name is the name of the chapter in the mongodb
def trans_chapter(_src, _chapter, _dest):
    try:
        src = _src.split('/')
        dest = _dest.split('/')

        dbDstCourse = dstmodulestore.find_one(
            {'_id.org': dest[0], '_id.category': 'course', '_id.course': dest[1], '_id.name': dest[2]})
        dbSrcChapter = srcmodulestore.find_one(
            {'_id.org': src[0], '_id.category': 'chapter', '_id.course': src[1], '_id.name': _chapter})
        if dbDstCourse == None:
            print "Destination course does not exist"
            print '{', '_id.org:', dest[0], '_id.category:course', '_id.course:', dest[1], '_id.name:', dest[2], '}'
        if dbSrcChapter == None:
            print "Source chapter does not exist"
            print '{', '_id.org:', src[0], '_id.category:chapter', '_id.course:', src[1], '_id.name:', _chapter, '}'
        if dbSrcChapter <> None and dbDstCourse <> None:
            #choose a new chapter name in case there is another chapter with the same name in the destination course
            chp_name = create_document_name(dbDstCourse, dbSrcChapter)
            #atach the new chapter to the child list of the destination course
            add_chapter(dbDstCourse, chp_name)
            dbSrcChapter['_id']['name'] = chp_name
            clone_document_to(dbDstCourse, dbSrcChapter)
    except:
        print "Error in trans chapter ", sys.exc_info()[0]
        raise


#we will clone a document (_dbSrcDocument) and all its childs recursively to the destiny course (_dbDstCourse)
#known bug if a element already exists it will add it correctly but not the links of the parent to him.
def clone_document_to(_dbDstCourse, _dbSrcDocument):
    try:
        dbSrcDocument = _dbSrcDocument
        if 'definition' in dbSrcDocument.keys() and 'children' in dbSrcDocument['definition'].keys():
            for i in range(len(dbSrcDocument['definition']['children'])):
                str_child = dbSrcDocument['definition']['children'][i]
                child = str_child.split('/')
                dbChild = srcmodulestore.find_one(
                    {'_id.org': child[2], '_id.category': child[4], '_id.course': child[3], '_id.name': child[5]})
                dbSrcDocument['definition']['children'][i] = 'i4x://' + _dbDstCourse['_id']['org'] + '/' + \
                                                             _dbDstCourse['_id']['course'] + '/' + dbChild['_id'][
                                                                 'category'] + '/' + create_document_name(_dbDstCourse,
                                                                                                          dbChild)
                clone_document_to(_dbDstCourse, dbChild)

        #delete xml link to avoid problems with linkage
        if 'metadata' in dbSrcDocument.keys() and 'xml_attributes' in dbSrcDocument['metadata'].keys():
           del dbSrcDocument['metadata']['xml_attributes']
        #if the document has format and that format exist on the course grading formats
        #we will add graded=true to the document
        #otherwise we will remove the document format
        if 'metadata' in dbSrcDocument.keys() and 'format' in dbSrcDocument['metadata'].keys() and dbSrcDocument['metadata']['format']<>"":
            if dstmodulestore.find({'_id.name':_dbDstCourse['_id']['name'],'_id.org':_dbDstCourse['_id']['org'],'_id.category':'course','_id.course':_dbDstCourse['_id']['course'],'definition.data.grading_policy.GRADER.type':dbSrcDocument['metadata']['format']}).count() >0:
                dbSrcDocument['metadata']['graded']='true'
            else:
                dbSrcDocument['metadata']['format']=''

        dbSrcDocument['_id']['org'] = _dbDstCourse['_id']['org']
        dbSrcDocument['_id']['course'] = _dbDstCourse['_id']['course']
        dbSrcDocument['_id']['name'] = create_document_name(_dbDstCourse, dbSrcDocument)


        if sys.argv.__contains__('-v'):
            print "Inserting :", dbSrcDocument

        id = bson.son.SON([('tag',dbSrcDocument['_id']['tag']),('org', dbSrcDocument['_id']['org']), ('course', dbSrcDocument['_id']['course']),('category',dbSrcDocument['_id']['category']),('name', dbSrcDocument['_id']['name']),('revision',dbSrcDocument['_id']['revision'])])
        dstmodulestore.insert( {'_id': id  ,'definition': dbSrcDocument['definition'] ,'metadata' : dbSrcDocument['metadata']})

    except:
        print "Error cloning the document ", _dbSrcDocument, " into the course ", _dbDstCourse, sys.exc_info()[0]
        raise


#we add the new chapter to the destination course children list'

#adds a chapter to the childlist of a course
def add_chapter(_dbDstCourse, _newChapterName):
    try:
        children = 'i4x://' + _dbDstCourse['_id']['org'] + "/" + _dbDstCourse['_id'][
            'course'] + "/chapter/" + _newChapterName
        dstmodulestore.update({'_id.org': _dbDstCourse['_id']['org'], '_id.category': 'course',
                            '_id.course': _dbDstCourse['_id']['course'], '_id.name': _dbDstCourse['_id']['name']},
                           {'$push': {'definition.children': children}})

    except:
        print "Error adding chapter as child of the course ", sys.exc_info()[0]
        raise


#to avoid duplicateIDs and weird stuff in the database we create a new chapter instead of updating
#the existing one in case that a chapter with the same name of the chapter we are importing already
#exists on the database
def create_document_name(_dbDstCourse, _dbSrcDocument):
    try:
        new_name = ''
        ordinal = 1
        exist = dstmodulestore.find(
            {'_id.org': _dbDstCourse['_id']['org'], '_id.category': _dbSrcDocument['_id']['category'],
             '_id.course': _dbDstCourse['_id']['course'], '_id.name': _dbSrcDocument['_id']['name']}).count()
        if (exist == 0):
            new_name = _dbSrcDocument['_id']['name']
        while new_name == '':
            exist = dstmodulestore.find(
                {'_id.org': _dbDstCourse['_id']['org'], '_id.category': _dbSrcDocument['_id']['category'],
                 '_id.course': _dbDstCourse['_id']['course'],
                 '_id.name': _dbSrcDocument['_id']['name'] + str(ordinal)}).count()
            if exist == 0:
                new_name = _dbSrcDocument['_id']['name'] + str(ordinal)
            ordinal = ordinal + 1
        return new_name
    except:
        print "Error creating the new document name", sys.exc_info()[0]
        raise


trans_chapter(sys.argv[2], sys.argv[3], sys.argv[5])
#python trans_chapter.py srvOrigen courseOrigen ChapterName srvDestination courseDestination


