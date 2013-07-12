# wsbase.py
#
# author: Yahui Liu (yh.liu@autosoft.com.cn)

# This script requires suds that provides SOAP bindings for python.
#
# Download suds from https://fedorahosted.org/suds/
#
# unpack it and then run:
# python setup.py install
#
# This may require you to install setuptools (an .exe from python.org)

import sys
#sys.path.append("/home/coverity/python-suds-0.4")

from suds import *
from suds.client import Client
from suds.wsse import *

from optparse import OptionParser
import cgi
import re

#imported by get_file_contents()
import zlib
from base64 import standard_b64decode

# uncomment to show SOAP xml
#import logging
#logging.basicConfig()
#logging.getLogger('suds.client').setLevel(logging.DEBUG)
#logging.getLogger('suds.transport').setLevel(logging.DEBUG)


# -----------------------------------------------------------------------------
# Base class for all the web service clients
class WebServiceClient:
    def __init__(self, webservice_type, host, port, user, password):
        url = 'http://' + host + ':' + port
        if webservice_type == 'configuration':
            self.wsdlFile = url + '/ws/v7/configurationservice?wsdl'
        elif webservice_type == 'defect':
            self.wsdlFile = url + '/ws/v7/defectservice?wsdl'
        else:
            raise "unknown web service type: " + webservice_type

        self.client = Client(self.wsdlFile)
        self.security = Security()
        self.token = UsernameToken(user, password)
        self.security.tokens.append(self.token)
        self.client.set_options(wsse=self.security)

    def getwsdl(self):
        print(self.client)


# -----------------------------------------------------------------------------
# Class that implements webservices methods for Defects
class DefectServiceClient(WebServiceClient):
    def __init__(self, host, port, user, password):
        WebServiceClient.__init__(self, 'defect', host, port, user, password)

    def set_ext_ref(self, cid, externalReference):
        triageStoreIdDO = self.client.factory.create('triageStoreIdDataObj');
        triageStoreIdDO.name = 'Default Triage Store';
        defectStateSpecDO = self.client.factory.create('defectStateSpecDataObj');
        defectStateSpecDO.externalReference = externalReference
        self.client.service.updateTriageForCIDsInTriageStore(triageStoreIdDO, cid, defectStateSpecDO)


    def get_merged_defects_for_project(self, project):
        print "TODO: get merged defects for " + project;
        PAGE_SIZE = 2500
        projectId  = self.client.factory.create('projectIdDataObj')
        projectId.name = project
        filterSpec = self.client.factory.create('mergedDefectFilterSpecDataObj') 
        pageSpec   = self.client.factory.create('pageSpecDataObj') 
        pageSpec.pageSize = PAGE_SIZE
        pageSpec.sortAscending = True
        pageSpec.startIndex = 0

        all_defects = []
        i = 0

        while True:
            pageSpec.startIndex = i
            defectsPage = self.client.service.getMergedDefectsForProject(projectId, filterSpec, pageSpec)
            i += PAGE_SIZE

            if (defectsPage.totalNumberOfRecords > 0):
                all_defects.extend(defectsPage.mergedDefects)

            if (i >= defectsPage.totalNumberOfRecords):
                break;

        return all_defects
    

    def get_stream_defects(self, cid):
        sdfsDO  = self.client.factory.create('streamDefectFilterSpecDataObj')
        sdfsDO.includeDefectInstances = True
        sdfsDO.includeHistory = True        
        return self.client.service.getStreamDefects([cid], sdfsDO)
    
    def get_merged_defects_for_stream(self, stream, componentList, componentExclude):
        #print "TODO: Defect Service---get merged defects for " + stream;
        PAGE_SIZE = 2500
        streamId  = self.client.factory.create('streamIdDataObj')
        streamId.name = stream
        filterSpec = self.client.factory.create('mergedDefectFilterSpecDataObj')
        temp = []
        print componentList
        #print componentList.split(',')
        #if componentList is None in config.xml  
        if (componentList !=None) :
            for x in componentList.split(','):
                componentIdList = self.client.factory.create('componentIdDataObj')
                componentIdList.name = x
                temp.append(componentIdList)

            filterSpec.componentIdList = temp
            if (componentExclude == "Y"):
                filterSpec.componentIdExclude = True
            else:
                filterSpec.componentIdExclude = False
        else :
            pass
        
        #print temp


        pageSpec   = self.client.factory.create('pageSpecDataObj') 
        pageSpec.pageSize = PAGE_SIZE
        pageSpec.sortAscending = True
        pageSpec.startIndex = 0

        all_defects = []
        i = 0

        while True:
            pageSpec.startIndex = i
            defectsPage = self.client.service.getMergedDefectsForStreams(streamId, filterSpec, pageSpec)
            i += PAGE_SIZE

            if (defectsPage.totalNumberOfRecords > 0):
                all_defects.extend(defectsPage.mergedDefects)

            if (i >= defectsPage.totalNumberOfRecords):
                break;

        return all_defects
    class SourceLine(object):  
        def __init__(self, num, text):  
            self.lineNum = num  
            self.text = text  
    def _cache_key(self, stream, file):  
            return '??'.join([file.contentsMD5,file.filePathname])
        
    def get_file_contents(self, stream):
	print "TODO: get file contents "
	#global _cache  
        #cache = {}
	streamId = self.client.factory.create('streamIdDataObj')
	streamId.name = "demo"
	field = self.client.factory.create('fileIdDataObj')
	field.filePathname = "/copy_demo/demo/inc/leak.c"
	field.contentsMD5 = "c30e5254860cec982befff145eb14e5c"
	key = self._cache_key(streamId, field)
	
	print streamId
	print field
	print key

	#if key not in _cache['files']:  
        base64Src = self.client.service.getFileContents(streamId, field)
        text = zlib.decompress(standard_b64decode(base64Src.contents))
        l = text.splitlines()
        print text
        lines = [self.SourceLine(*x) for x in zip(range(1,len(l)+1), l)]
        #print lines
        return text
        #cache['files'][key] = (text,lines)  
        #(self.contents,self._split_contents) = cache['files'][key]  

    def get_CIDs_For_Streams(self, stream):
        print "TODO: GET cids for streams: " + stream
        streamId = self.client.factory.create('streamIdDataObj')
	streamId.name = stream
	filterSpec = self.client.factory.create('mergedDefectFilterSpecDataObj')
	return self.client.service.getCIDsForStreams(streamId)

    def get_MergedDefect_History(self, cids, stream):
        print "TODO: get MergedDefect History:"
        streamId = self.client.factory.create('streamIdDataObj')
        streamId.name = stream
        return self.client.service.getMergedDefectHistory(cids, streamId)

    def get_CheckerSubcategories_ForStreams(self, stream):
        print "TODO: Defect Service---get checker subcategories for streams.\n"
        streamIdData = self.client.factory.create('streamIdDataObj')
        streamIdData.name = stream
        return self.client.service.getCheckerSubcategoriesForStreams(streamIdData)

    def get_ComponentMetrics_ForProject(self, project, component):
        projectId = self.client.factory.create('projectIdDataObj')
        projectId.name = project
        temp = []
        for x in component:
            componentIds = self.client.factory.create('componentIdDataObj')
            componentIds.name = x
            temp.append(componentIds)
##        print "wsbase::get_ComponentMetrics_ForProject, component = "
##        print temp
        return self.client.service.getComponentMetricsForProject(projectId, temp)
    
class SourceFile(object):  
    ''''' 
    Helper class for a source code file 
    '''  
    class SourceLine(object):  
        def __init__(self, num, text):  
            self.lineNum = num  
            self.text = text  
    def _cache_key(self, stream, file):  
            return '??'.join([file.contentsMD5,file.filePathname])  
          
    def __init__(self, stream, file):  
            global _cache  
            key = self._cache_key(stream, file)  
            if key not in _cache['files']:  
                src = client.defect.getFileContents(stream, file)  
                text = zlib.decompress(standard_b64decode(src.contents))  
                l = text.splitlines()  
                lines = [self.SourceLine(*x) for x in zip(range(1,len(l)+1), l)]  
                _cache['files'][key] = (text,lines)  
                (self.contents,self._split_contents) = _cache['files'][key]  
          
    def snippet(self, line, caption=None, context=7):  
            start = line - (context/2) - 1  
            if start < 0:  
                context += start  
                start = 0  
            lines = self._split_contents[start:start+context]  
            if caption:  
                lines.insert(context/2, self.SourceLine('',caption))  
            return lines  
# -----------------------------------------------------------------------------
# Class that implements webservices methods for Configuration
class ConfigurationServiceClient(WebServiceClient):
    def __init__(self, host, port, user, password):
        WebServiceClient.__init__(self, 'configuration', host, port, user, password)

    def getProjects(self):
        return self.client.service.getProjects();

    def addUser(self, username):
        print "TODO: add user " + username
        userSpec = self.client.factory.create('userSpecDataObj')
        userSpec.username = username
        userSpec.password = "coverity"
        userSpec.email = "addApiT@163.com"
        userSpec.familyName = "Liu"
        userSpec.givenName = "DemoAPI"
        userSpec.local = True
        
        return self.client.service.createUser(userSpec)

    def addStreamInProject(self, streamName, projectName):
        print "TODO: add stream "
        streamSpec = self.client.factory.create('streamSpecDataObj')
        streamSpec.name = "demoAddByAPIS"
        streamSpec.language = "CXX"
        streamSpec.description = "This is added by API."
        triageStoreId = self.client.factory.create('triageStoreIdDataObj')
        triageStoreId.name = "Default Triage Store"
        streamSpec.triageStoreId = triageStoreId
        componentMapId = self.client.factory.create('componentMapIdDataObj')
        componentMapId.name = "Default"
        streamSpec.componentMapId = componentMapId

        projectSpec = self.client.factory.create('projectIdDataObj')
        projectSpec.name = "huawei"
        return self.client.service.createStreamInProject(projectSpec, streamSpec)

    def get_commit_gate(self):
        return self.client.service.getCommitState()

    def get_license_state(self):
        return self.client.service.getLicenseState()

    def notify(self, message):
        usernames = ['admin', 'test']
        subject = "This is a test mail from Web Service API"
        text  = message
        return self.client.service.notify(usernames, subject, text)

    def get_Checker_Properties(self, checker, category, project):
        #print "TODO: Configuration Service --- get checker properties.\n"
        #print project
        filterSpec = self.client.factory.create('checkerPropertyFilterSpecDataObj')
        filterSpec.checkerNameList = checker
        filterSpec.subcategoryList = category
        filterSpec.projectId.name = project
        filterSpec.domainList = "STATIC_C"
        #print filterSpec.checkerNameList
        #print filterSpec.subcategoryList
        #print filterSpec.projectId.name
        return self.client.service.getCheckerProperties(filterSpec)

    def get_Component(self, component):
        filterSpec = self.client.factory.create('componentMapFilterSpecDataObj')
        filterSpec.namePattern = component
        return self.client.service.getComponentMaps(filterSpec)
# end-----------------------------------------------------------------------------
