'''
This library contains a listener to log information to the console
as the execution happens and also trigger an email when a test case fails.

Required and optional arguments that can be passed with this listener are

1. Required argument: Recipients who need to get the mail when a test case fails.        
2. Optional Argument: SMTP server which need to be used for sending email.

        Ex: --listener ConsoleLogListener.py:abc@gmail.com:xyz@gmail.com:SMTPHostname
        
Created on 10/5/2012 by Grace
Modified latest on 11/22/2012 by Grace
Modified latest on 16/01/2013 by Grace
Modified latest on 25/07/2013 by Grace

'''

from robot.api import logger
from robot.errors import ExecutionFailed
from selenium.common.exceptions import WebDriverException

import datetime
import sendingmail
import sys
import re
import os.path, time
import shutil
import Selenium2Library
from robot.libraries.OperatingSystem import OperatingSystem
from Selenium2Library import keywords
from robot.libraries.BuiltIn import BuiltIn
from robot.libraries.String import String
import string
import socket


class ConsoleLogListener:

    ROBOT_LISTENER_API_VERSION = 2    
    email = []
    args = []
    SMTPHost = 'inhbcasarray.ftdcorp.net'
    OutPath = ""    
    OutputFolder = ""
    fHtmlSourceFile =""
    SelLibInstance = ""
    KeywordStarttime = ""
    TestCaseOutPutFile =""
    HtmlSourceCode = ""
    Tname = ""
    SourceDir = ""
    Failed = 0
    Passed = 0
    ImagePath=""
    logIndentation = ""
    
    def __init__(self,e,filename='',smtp='inhbcasarray.ftdcorp.net'): # this will be executed when the listener is called
        dt = datetime.datetime.now()

        self.email=e # following steps are to get the details from the aruments passed at commandline
        self.SMTPHost = smtp
        self.args=sys.argv # this will give the entire arguments passed as list
        outputdirectory = self.ReturnOutputFileDirectory()# this is to get the output directory, incase passed        
                
        if (outputdirectory!=""):  #if output directory is passed, then setting the path for saving the HtmlSourceCode and temporary log file                
                dirpath= string.find(self.args[outputdirectory],"C:\Jenkins")
                if (dirpath==-1): # if complete workspace path is not given, then consider taking the working directory and specified output directory
                    self.OutputFolder = os.getcwd()+"/"+self.args[outputdirectory]
                    
                else:  # if workspace path is given, then directly create the files in the given workspace path
                    self.OutputFolder = self.args[outputdirectory]
        else: # if output directory is not passed, then save in the current working directory
                self.OutputFolder = os.getcwd()

        self.OutputFolder = self.OutputFolder+"/"+dt.strftime("%Y/%m/%d")
        
        if not os.path.exists(self.OutputFolder): # create the directory to save source files with datetime as folder name under the given output directory      
                  os.makedirs(self.OutputFolder)
        
        self.TestCaseOutPutFilePath = self.OutputFolder +'/Log_'+dt.strftime("%H-%M-%S.%f")+'.txt'
        
    def start_test(self,name,attrs):       # executed at the start of every test
        
        self.TestCaseOutPutFile = open(self.TestCaseOutPutFilePath,'w') #make sure to close the file in end_test
        self.Tname=name.replace(" ","_")
        self.HtmlSourceCode = ""
        self.fHtmlSourceFile = ""
        self.ImagePath = ""
        self.logIndentation = ""
        
        self.PrintMsg('\n*******************************************'+          
                     '\n%s : START OF THE TEST: %s'%(attrs['starttime'],name))
        self.SelLibInstance = BuiltIn().get_library_instance('Selenium2Library') # this will return the current selenium libary instance
        
        
    def start_keyword(self,name,attrs):        # executed at the start of every keyword within a test
        self.KeywordStarttime = time.asctime( time.localtime(time.time()) )    # to get the time when the test started, later used to caluclate the screenshots generated while test execution
        variablePattern = '.*(\$\{.*?\})'
        
        #Check comments: Don't try to replace in comments, as there might be invalid variables
        if (name.lower().strip() != 'builtin.comment'):
            args = []
            for arg in attrs['args']:
                argActual = arg
                match = re.search(variablePattern,arg)
                while match:
                    try:
                        variableName = match.group(1)
                        variableValue = BuiltIn().get_variables()[variableName]
                        variableValue = str(variableValue)
                    except UnicodeEncodeError: pass  #this exception will occur if there are special chars
                    except:
                        variableValue = ""
                    	
                    arg = arg.replace(variableName, variableValue);
                    match = re.search(variablePattern,arg)
                if argActual == arg:
                    args.append("["+arg+"]")
                else:
                    args.append("["+argActual+"="+arg+"]")
                
            strArguments = ','.join(args)
        else:
            strArguments = ','.join(attrs['args'])
		
        self.PrintMsg('%s : %s KEYWORD : %s (  %s  )'%(attrs['starttime'],self.logIndentation, name,strArguments))
		
		#Add a tab to logIndentation, so that if any other keyword is called before end_keyword, the log will be indented by a tab
		#remove the tab at the end of end_keyword
        self.logIndentation = self.logIndentation + "\t"
		       
    def end_keyword(self,name,attrs): # executed at the end of every keyword.        
       if (attrs['status']=='FAIL'): # if the keyword fails, the following code will retrieve the source code of the opened browser page into a html file.
         FaliureTime = datetime.datetime.now().strftime("%H-%M-%S.%f")

         try: # this will throw a run time error if no browser is open. hence handling the exception
          
         	#Take the screen shot only for the keyword failure (ignore tear down keywords failures)
          
          if attrs['type'] == "Keyword": 
            self.HtmlSourceCode = Selenium2Library.Selenium2Library.get_source(self.SelLibInstance)
            
            #Set the file names
            self.fHtmlSourceFile 	= self.OutputFolder+"\\"+self.Tname+"_"+FaliureTime+".html" # creating source code file with timestamp
            self.ImagePath 			= self.OutputFolder+"\\"+self.Tname+"_"+FaliureTime+".png"
            
            #wrie html source to file
            if self.HtmlSourceCode != "":
            	with open(self.fHtmlSourceFile, 'w') as f: f.write(self.HtmlSourceCode.encode('utf8'))
            
            #Take screen shot
            Selenium2Library.Selenium2Library.capture_page_screenshot(self.SelLibInstance,filename=self.ImagePath) # capture screenshot and save it with test case name and timestamp
          
         except RuntimeError:
            self.PrintMsg('%s'%("No browser open"))                
            
         except WebDriverException as e:
            self.PrintMsg("	Unable to get screenshot. Reason: "+e.__str__())
            self.HtmlSourceCode = ""
            self.ImagePath = ""
         
       #Remove one tab, as the end_keyword is completed
       self.logIndentation = self.logIndentation[:-1]
                
    def end_test(self,name,attrs):         # executed once a test ends.        
        self.logIndentation = ""
        self.PrintMsg('%s : %s %s'%(attrs['endtime'],"END OF THE TEST:",name))                
        self.PrintMsg('\n\nTest Case Status: %s '%(attrs['status']))
        if (attrs['status']=='FAIL'):                
                
                self.PrintMsg('Error Message : %s '%(attrs['message']))
                self.TestCaseOutPutFile.close()
                
                self.Failed = self.Failed+1
                
                browser = self.GetCommandLineArgument("browser")     # to retrive the browser information passed at commandline.
                env = self.GetCommandLineArgument("env")         # to retrive the QA environment information passed at commandline.
                hostname = self.GetCommandLineArgument("remote_url") 
                if (hostname==""):	#treat at localhost, if there is no remote_url                        
                	hostname = str(socket.gethostname())
                        
                outputfile = self.ReturnOutputFileDirectory()# to retrive the output directory information passed at commandline.
                
                if (outputfile!=""):
                    outputfile=self.args[outputfile]                    

                # send a mail with testcase name, receivers of the email, SMTPHost to be used, browser, qa environment, hostname, temp log file path, source code file path,test start time and failure time
                sendingmail.dropmail(name,self.email,self.SMTPHost,browser,env,hostname,self.TestCaseOutPutFilePath, self.fHtmlSourceFile,self.ImagePath)            
        else:
                self.Passed = self.Passed+1
                self.TestCaseOutPutFile.close()
        
        logger.console('\n%s'%("##########################################"))        
        logger.console('%s %s %s %s'%("Status So far :: PASSED: ",str(self.Passed)," FAILED: ",str(self.Failed)))
        logger.console('%s'%("##########################################\n"))
    
    def close(self): # executed once the listener is called off
        dt = datetime.datetime.now()        
        logger.console('%s'%("*******************************************"))       
        logger.console('%s : All tests executed'%(str(dt.time())))        
        logger.console('%s'%("*******************************************"))
        os.remove(self.TestCaseOutPutFilePath)
        
    def GetCommandLineArgument(self, argument):# method to retreive the qa environment information passed at command line. Is will search with "env:" string in the entire arguments list. It will return "" if this information is not found
        GetValueCommandline = [i for i,item in enumerate(self.args) if re.search(argument+":*".lower(), item.lower())]
        if (len(GetValueCommandline)>0):
            argValue = self.args[GetValueCommandline[0]]
            argValue = self.args[GetValueCommandline[0]][len(argument)+1:]
            return argValue
        else:
            return ""
            
    def ReturnOutputFileDirectory(self):# method to retreive the browser information passed at command line. Is will search with "-d" or "--outputdir:" string in the entire arguments list. It will return "" if this information is not found
       str1="-d *"
       str2="--outputdir *"
       GetValueCommandline = [i for i,item in enumerate(self.args)if (re.search(str1.lower(), item.lower())  or re.search(str2.lower(), item.lower()) )]
       
       if (len(GetValueCommandline)>0):               
         item = self.args[GetValueCommandline[0]]
         ind = self.args.index(item)
         ind = ind+1         
         return ind
       if (len(GetValueCommandline)<=0):
         return ""
               
    def PrintMsg(self,msg): # method to print the passed message on to the console and also write the same message to a file
       msg = msg.replace("\n",os.linesep)
       logger.console(msg.encode('latin1'))
       self.TestCaseOutPutFile.write(os.linesep+msg.encode('latin1'))
