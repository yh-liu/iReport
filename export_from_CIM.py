# export_from_CIM.py
#
# author: Yahui Liu (yh.liu@autosoft.com.cn)

# This script requires wsbase.py.

# This may require you to install setuptools (an .exe from python.org)

from wsbase import *
import xml.etree.ElementTree
import xml.etree.cElementTree as et

def get_component_list(service, map):
##    componentList = []
##    with open ('defectList.csv') as fn:
##        for each_line in fn:
##            data = each_line.split(',')
##            #The eighth field of defectlist.csv line is component name, so the array index should be 7
##            componentList.append(data[7].strip())
##    #print componentList
##    return list(set(componentList))
      readComList = tree.find('.*/componentList').text
            
      all = service.get_Component(map)
      
      #print all[0].components
      
      temp = []
      for item in all[0].components:
        temp.append(item[0].name)
    # this means all the component  
      if (readComList == None):
         return temp

      else :
        if((tree.find('.*/componentExclude').text == 'N')) :
            readList = []
            for x in readComList.split(','):
                  readList.append(x)
            return readList
        else :
            for i in readList :
                temp.remove(i)
            return temp
            

#def write_component_metric_to_file():
    

if __name__ == '__main__':
    
    tree = et.ElementTree(file="config.xml")
    parser = OptionParser()

    parser.set_defaults(host=tree.find('*/host').text)
    parser.set_defaults(port=tree.find('*/port').text)
    parser.set_defaults(user=tree.find('*/user').text)
    parser.set_defaults(password=tree.find('*/password').text)
    parser.set_defaults(project=tree.find('*/project').text)
    parser.set_defaults(stream=tree.find('*/stream').text)
    parser.set_defaults(component=tree.find('*/componentList').text)
    parser.set_defaults(componentExclude=tree.find('*/componentExclude').text)

    parser.add_option("--host",      dest="host",      help="host of Coverity Connect");
    parser.add_option("--port",      dest="port",      help="port of Coverity Connect");
    parser.add_option("--user",      dest="user",      help="Coverity Connect user");
    parser.add_option("--password",  dest="password",  help="Coverity Connect password");
    parser.add_option("--project",   dest="project",   help="Coverity Connect project");
    parser.add_option("--stream",   dest="stream",   help="Coverity Connect stream");
    parser.add_option("--xml",       dest="xml",       help="path to cov export file");
    parser.add_option("--component", dest="component", help="Coverity Connect Component");
    parser.add_option("--componentExclude", dest="componentExclude", help="component included or excluded? ");
    (options, args) = parser.parse_args()
    if (not options.host or
        not options.port or
        not options.user or
        not options.password or not options.project):
        parser.print_help()
        sys.exit(-1)
    #print options.host, options.port, options.user, options.password
    defectService = DefectServiceClient(options.host, options.port, options.user, options.password);
    configurationService = ConfigurationServiceClient(options.host, options.port, options.user, options.password);

    item = []
    checker=[]
    checkerList = []
    checkerPropertieSet = []
    defectCSV = open("defectList.csv","w")
    #defectCSV.write("CID,       subcategoryShortDescription,    checkerName,    checkerSubcategory,     Category,   Description,    Impact,     componentName,   DefectStatus,    Classification,   Severity,     Owner,     Comment,   functionName,       filePathname    \n")

##    0 DefectStatus New
##    1 Classification Unclassified
##    2 Action Undecided
##    3 Severity Unspecified
##    4 Fix Target Untargeted
##    5 Owner Unassigned
##    6 TranslatedOwner Unassigned
##    7 OwnerName None
##    8 Comment None  ###
    
    for bugs in defectService.get_merged_defects_for_stream(options.stream, options.component, options.componentExclude):
        #print bugs
        if [bugs.checkerName, bugs.checkerSubcategory] not in checkerList:
            checkerList.append([bugs.checkerName, bugs.checkerSubcategory])
            checker = configurationService.get_Checker_Properties(bugs.checkerName, bugs.checkerSubcategory, options.project)
            checkerPropertieSet.extend(checker)
        else:
            #checker = checkerPropertieSet[-1]
            #print checker
            for checkerProper in checkerPropertieSet:
##                print checkerPropertieSet
##                print "#########"
##                print checkerProper.checkerSubcategoryId.checkerName
                if (bugs.checkerName == checkerProper.checkerSubcategoryId.checkerName) & (bugs.checkerSubcategory == checkerProper.checkerSubcategoryId.subcategory):
                    checker = [checkerProper]
                    break
                
        #print checkerPropertieSet
        #print checkerList
        #print checker
        #the following comment is for comment export ----bugs.defectStateAttributeValues[8].attributeValueId.name
        #if the comment is needed, it is seperated with defectlist.csv better and another csv file is necessory
##        print bugs
##        print "@@@@@@@@@@@@@@@@@@@@@@@@@@@@"
##        print bugs.defectStateAttributeValues[3].attributeValueId.name
##        print bugs.defectStateAttributeValues[7].attributeValueId.name
##        print bugs.defectStateAttributeValues[8].attributeValueId.name
##        print "####################################"
        item = [str(bugs.cid), checker[0].subcategoryShortDescription, bugs.checkerName, bugs.checkerSubcategory, checker[0].category, checker[0].categoryDescription, checker[0].impact, bugs.componentName, \
                        bugs.defectStateAttributeValues[0].attributeValueId.name, bugs.defectStateAttributeValues[1].attributeValueId.name,  \
                        bugs.defectStateAttributeValues[3].attributeValueId.name, \
                        bugs.defectStateAttributeValues[7].attributeValueId.name] #\
                        #bugs.defectStateAttributeValues[8].attributeValueId.name,  \
                        #bugs.functionName, bugs.filePathname]
        
        
        #print item
        defectCSV.write(str(item).strip("[]"))
        defectCSV.write("\n")
        item = []
        checker=[]
        #for bugAttribute in bugs.defectStateAttributeValues:
        #    print bugAttribute.attributeDefinitionId.name, bugAttribute.attributeValueId.name
        #quit()
    defectCSV.close()
    
    with open('defectList.csv') as fn:
        line = fn.readline()
        if (len(line) == 0) :
            print "No defects found. \n"
            exit(1)
        else:
            line = line.split(',')
        print line
        print line[7]
        map = line[7].strip().split('.')[0]
    #print "map = ", map 
    component_list = get_component_list(configurationService, map)
##    print "component_list = "
##    print component_list
    componentMetricsCSV = open('componentMetrics.csv','w')
    #componentMetricsCSV.write(component name, codeLine, blankLine, commentLine, totalCount, outstandingCount, inspectedCount, fixedCount, dismissedCount, triagedCount, newCount)

    for item in defectService.get_ComponentMetrics_ForProject(options.project, component_list):
        line = [item.componentId.name, item.codeLineCount, item.blankLineCount, item.commentLineCount, item.totalCount, item.outstandingCount, item.inspectedCount, item.fixedCount, item.dismissedCount, item.triagedCount, item.newCount]
        componentMetricsCSV.write(str(line).strip("[]"))
        componentMetricsCSV.write("\n")
        line = []
    componentMetricsCSV.close()

    # New a test file for next step if the script can run here.
    test =  open ('test','w')
    test.close()
