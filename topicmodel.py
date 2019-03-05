# -*- coding: utf-8 -*-
"""
Created on Thu Dec 20 09:47:14 2018

@author: etienneg
"""


# Run in terminal or command prompt
# python -m spacy download en
   
# Run in terminal or command prompt
# python3 -m spacy download en


import os,sys,fnmatch,argparse,datetime
import glob
from email import policy
from email.parser import BytesParser
import win32com.client,docx
import PyPDF2

import numpy as np
import pandas as pd
import pickle
import ast
from pprint import PrettyPrinter
import matplotlib.pyplot as plt
from sklearn.cluster import KMeans
from collections import Counter


from wordcloud import WordCloud
import matplotlib.colors as mcolors

import warnings
warnings.filterwarnings("ignore", category=Warning)
warnings.filterwarnings("ignore", category=PendingDeprecationWarning) 
warnings.filterwarnings("ignore", category=DeprecationWarning) 
warnings.filterwarnings("ignore", category=FutureWarning)
warnings.filterwarnings("ignore", category=UserWarning)




import re, spacy, gensim
from nltk.corpus import stopwords
original_stop_words = stopwords.words('english')

# Sklearn
from sklearn.decomposition import LatentDirichletAllocation, TruncatedSVD, NMF
from sklearn.feature_extraction.text import CountVectorizer, TfidfVectorizer
from sklearn.model_selection import GridSearchCV


#from gensim.utils import simple_preprocess
from gensim.parsing.preprocessing import STOPWORDS
from nltk.stem import WordNetLemmatizer, SnowballStemmer
#from nltk.stem.porter import *
#import nltk
import PyQt5
from PyQt5 import QtWidgets,QtCore,QtGui
from PyQt5.QtWidgets import (QMessageBox,QProgressDialog,QFileDialog)
PyQt5.QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling, True)
import pyLDAvis.sklearn

from topicui import Ui_Form



class TopicModeltoPickle:
    def __init__(self,
                 email_vectoriser=None,
                 doc_vectoriser=None,
                 LDAEmailModel=None,
                 LDADocModel=None,
                 NMFEmailModel=None,
                 NMFDocModel=None,
                 removestopwords=True,
                 stopwords="",
                 processtext=True,
                 cleantext=True,
                 lemmapattern="['NOUN','ADJ','VERB','ADV']"):
        #Need everyting here to process future text to see which topic it belongs to so
        
        ##  clean text in same way
        ## vectorise the text using the vectoriser
        ## transform the text using the vectoriser
        self.removestopwords=removestopwords
        self.stopwords=stopwords
        self.email_vectoriser=email_vectoriser
        self.doc_vectoriser=doc_vectoriser
        self.cleantext=cleantext
        self.processtext=processtext
        self.lemmapattern=lemmapattern
        self.NMFEmailModel=NMFEmailModel
        self.NMFDocModel=NMFDocModel
        self.LDAEmailModel=LDAEmailModel
        self.LDADocModel=LDADocModel
        



class toPickle:
    def __init__(self,email_txt,email_txt_emailname,doc_txt,doc_txt_emailname):
        self.email_txt=email_txt
        self.email_txt_emailname=email_txt_emailname
        self.doc_txt=doc_txt
        self.doc_txt_emailname=doc_txt_emailname
    
        
        
        
        
        
class TopicModel:
    stemmer = SnowballStemmer('english')
    word = win32com.client.Dispatch("Word.Application")
    word.visible = False
    
    def __init__(self,emaildir=None,tempdir=None,readdrivedir=None):

        # Initialize spacy 'en' model, keeping only tagger component (for efficiency)
        self.nlp = spacy.load('en', disable=['parser', 'ner'])

        
        self.emaildirectory=emaildir
        self.readdrivedirectory=readdrivedir
        self.tempdir=tempdir
        

       
        
        #now do all the initialisation
        self.emailtext=[]
        self.emailtext_original=[]
        self.emailtext_words=[]
        self.emailtext_forvectorize=[]
        self.emailtext_emailname=[]
        self.emailtext_topic=[]
        
        self.emailtexttopwords=[]
        self.emailtexttopwordweights=[]
        
        self.emaildoctopwords=[]
        self.emaildoctopwordweights=[]
        
        self.doctext=[]
        self.doctext_original=[]
        self.doctext_forvectorize=[]
        self.doctext_words=[]
        self.doctext_emailname=[]
        self.doctext_topic=[]
        self.removewords=False
        self.removelist=[]
        self.cleantext=False
        self.processtext=False
        self.lemma_args=[]
        self.ngram_args="1,1"
        self.bigram_args=[]
        self.regex_pattern=[]
        
        self.doattachments=True
        self.attachmentdirectory=[]
        self.dopdf=True
        
        self.text_nmf=[]
        self.doctext_nmf=[]
        self.tfidfTextFeatureNames=[]
        self.tfidfDocTextFeatureNames=[]
        self.text_lda=[]
        self.text_number_topics=None
        
        self.doctext_lda=[]
        self.doctext_number_topics=None
        
        
        self.countTextFeatureNames=[]
        self.countDocTextFeatureNames=[]
              
        self.currentEmailVectorizer=None
        self.currentDocVectorizer=None   
        self.currentEmailVectorised=None
        self.currentDocVectorised=None
        self.stoplistextended=False
        self.stop_words=original_stop_words
        
        self.Algorithm="LDA"
        self.VectorizerMethod="Count Vectorizer"
        self.LDALearningMethod="batch"
        
        self.df_email_TextTopics=pd.DataFrame()
        self.df_email_DocTopics=pd.DataFrame()
        self.df_email_doc_topic_distribution = pd.DataFrame()
        self.df_email_text_topic_distribution = pd.DataFrame()      
                
        
 
        
    def populateWordWeightMatrix(self,no_top_words=10):
        model=self.text_lda
        if not(model==None):
            weights=model.components_/model.components_.sum(axis=1)[:,None]
            for idx,topic in enumerate(weights):
                self.emailtexttopwords.append([self.currentEmailVectorizer.get_feature_names()[i] for i in topic.argsort()[:-no_top_words -1:-1]])
                self.emailtexttopwordweights.append([topic[i] for i in topic.argsort()[:-no_top_words -1:-1]])
        model=self.doctext_lda
        if not(model==None):
            weights=model.components_/model.components_.sum(axis=1)[:,None]
            for idx,topic in enumerate(weights):
                self.emaildoctopwords.append([self.currentDocVectorizer.get_feature_names()[i] for i in topic.argsort()[:-no_top_words -1:-1]])
                self.emaildoctopwordweights.append([topic[i] for i in topic.argsort()[:-no_top_words -1:-1]])
                
            
            
            
    def displayTopics(self, model, vectorizer, no_top_words=10,printvalues=False):
        for idx,topic in enumerate(model.components_):
            print("Topic %d:" % (idx))
            if printvalues:
                print([(vectorizer.get_feature_names()[i],topic[i]) for i in topic.argsort()[:-no_top_words -1:-1]])
            else:    
                print(" ".join([vectorizer.get_feature_names()[i] for i in topic.argsort()[:-no_top_words -1:-1]]))
    
    def stringTopics(self, model, vectorizer, no_top_words=10):
        toreturn=""
        for idx,topic in enumerate(model.components_):
            toreturn=toreturn+"Topic %d:" % (idx)
            toreturn=toreturn +" ".join([vectorizer.get_feature_names()[i] for i in topic.argsort()[:-no_top_words -1:-1]])
            toreturn=toreturn +"\n"
        return toreturn
    
    def listTopics(self,model,vectorizer,no_top_words=10):
        toreturn=[]
        for idx,topic in enumerate(model.components_):
            string="Topic %d:" % (idx)
            string=string +" ".join([vectorizer.get_feature_names()[i] for i in topic.argsort()[:-no_top_words -1:-1]])
            toreturn.append(string)
        return toreturn
    
    def displayEmailsMatchingTextTopic(self,topicnumber):
        result=[]
        for idx,topic in enumerate(self.emailtext_topic):
            if topic==topicnumber:
                result.append(self.emailtext_emailname[idx])
        return result

    def displayEmailsMatchingDocTopic(self,topicnumber):
        result=[]
        for idx,topic in enumerate(self.doctext_topic):
            if topic==topicnumber:
                result.append(self.doctext_emailname[idx])
        return result

            
    def _save_file(self,fn, cont):
        """Saves cont to a file fn"""
        #print("Trying to create file:",fn)
        filename=os.path.join(self.attachmentdirectory,fn)  
        if(os.path.exists(filename)):
            #if the file exists delete it as we must have run this before
            os.remove(filename)         
        try:
           
            # handle error here
            file = open(filename, "wb")
            file.write(cont)
            file.close()
        except OSError:
            print("Could not create file: ",filename)
            #create filename with current date and time
            
        return (filename)
    
    def _read_text_msword(self,filename):
        if(os.path.exists(filename)):
            try:
                wb = self.word.Documents.Open(filename)
                doc = self.word.ActiveDocument
                Text=doc.Range().Text
                wb = self.word.Documents.Close()
                return Text
            except:
                return ""
        else:
            return ""
        
    def _read_text_pdf(self,filename):
        Text=""
        if(os.path.exists(filename)):
            try:
                pdfFileObj = open(filename, 'rb') 
  
                # creating a pdf reader object 
                pdfReader = PyPDF2.PdfFileReader(pdfFileObj) 
  
                for page in pdfReader.pages:
                    Text=Text+page.extractText() 
                pdfFileObj.close() 
                return Text
            except:
                return ""
        else:
            return ""
        
          
    def _preprocess(self,text,english_regex=False,RemoveList=None):
        resulttext=[]
        if RemoveList==None:
            RemoveList=""
        
        
        for token in gensim.utils.simple_preprocess(text):
            if english_regex:
                if token not in STOPWORDS and token not in RemoveList and len(token) > 3 and re.match('[a-zA-Z\-][a-zA-Z\-]{2,}', token):
                    resulttext.append(self.stemmer.stem(WordNetLemmatizer().lemmatize(token,pos='v')))
            else:
                if token not in STOPWORDS and token not in RemoveList and len(token) > 3:
                    resulttext.append(self.stemmer.stem(WordNetLemmatizer().lemmatize(token,pos='v')))

        return resulttext
    
    def sentencesToWords(self,sentences):
        
        for sentence in sentences:
            yield(gensim.utils.simple_preprocess(str(sentence), deacc=True))  # deacc=True removes punctuations
    
    
    
    def remove_stopwords(self,texts):
        return [[word for word in gensim.utils.simple_preprocess(str(doc)) if word not in self.stop_words] for doc in texts]
    
    def lemmatizeText(self,texts,allowed_postags=['NOUN', 'ADJ', 'VERB', 'ADV']):

        """https://spacy.io/api/annotation"""
        texts_out = []
        self.nlp.max_length=2000000
        for sent in texts:
            doc = self.nlp(" ".join(sent)) 
            texts_out.append(" ".join([token.lemma_ if token.lemma_ not in ['-PRON-'] else '' for token in doc if token.pos_ in allowed_postags]))
        return texts_out

    
    def cleanandProcessText(self):
        
        self.stop_words=original_stop_words
        self.stop_words.extend(re.sub("\'","",self.removelist).split(','))
        
        self.doctext_words[:]=[]
        self.emailtext_words[:]=[]
        
        if self.emailtext_original:
            self.emailtext=self.emailtext_original
            if self.cleantext: # was the clean text box ticked/comannd line arg chosen ?
                self.emailtext = [re.sub('\S*@\S*\s?', '', sent) for sent in self.emailtext] 
                # Remove new line characters
                self.emailtext = [re.sub('\s+', ' ', sent) for sent in self.emailtext]
    
                # Remove distracting single quotes
                self.emailtext = [re.sub("\'", "", sent) for sent in self.emailtext]

            
            self.emailtext_words=list(self.sentencesToWords(self.emailtext))               
            if self.removewords:
                self.emailtext_words=self.remove_stopwords(self.emailtext_words)
                
            #now lemmatize this lot
            if self.processtext:  ## Was lemmatize selelcted or was command line switch invoked ?
                self.emailtext_forvectorize=self.lemmatizeText(self.emailtext_words,allowed_postags=self.lemma_args)
            else:
                self.emailtext_forvectorize=[" ".join(txt) for txt in self.emailtext_words]
                
                
            
        if self.doctext_original:
            self.doctext=self.doctext_original
            if self.cleantext:
                self.doctext = [re.sub('\S*@\S*\s?', '', sent) for sent in self.doctext] 
                # Remove new line characters
                self.doctext = [re.sub('\s+', ' ', sent) for sent in self.doctext]
    
                # Remove distracting single quotes
                self.doctext = [re.sub("\'", "", sent) for sent in self.doctext]

            self.doctext_words=list(self.sentencesToWords(self.doctext))
            if self.removewords:               
                self.doctext_words=self.remove_stopwords(self.doctext_words)
                
            #now lemmatize this lot
            if self.processtext:
                self.doctext_forvectorize=self.lemmatizeText(self.doctext_words,allowed_postags=self.lemma_args)
            else:
                self.doctext_forvectorize=[" ".join(txt) for txt in self.doctext_words]
                
            
    

    def displayErrorMessage(self,message):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Warning)
        msg.setText(message)
        
    def initialiseTextFields(self):    
        self.emailtext_original[:]=[]
        self.emailtext[:]=[]
        self.emailtext_words[:]=[]
        self.emailtext_emailname[:]=[]
        
        self.doctext[:]=[]
        self.doctext_words[:]=[]
        self.doctext_original[:]=[]
        self.doctext_emailname[:]=[]
        
        self.doctext_forvectorize[:]=[]
        self.emailtext_forvectorize[:]=[]        
        
        
    def resetVectorisersLDA(self):

        self.text_nmf=None
        

        self.doctext_nmf=None
        
        self.tfidfTextFeatureNames=[]
        self.tfidfDocTextFeatureNames=[]
        
        self.text_lda=None

        self.doctext_lda=None
       
        
        self.countTextFeatureNames=[]
        self.countDocTextFeatureNames=[]
        
        self.currentEmailVectorizer=None
        

        self.currentDocVectorizer=None  
        
        self.currentEmailVectorised=None
        
        self.currentDocVectorised=None    
        
        self.doctext_forvectorize[:]=[]
        self.emailtext_forvectorize[:]=[]

        self.df_email_TextTopics=pd.DataFrame()
        self.df_email_DocTopics=pd.DataFrame()
        self.df_email_doc_topic_distribution = pd.DataFrame()
        self.df_email_text_topic_distribution = pd.DataFrame() 
        
        
    def readEmailsExchange(self,progressbar,folderstoread,numberemails,readattachments=False,dopdf=False):
        ##use iterator as mailbox can mutate while busy reading
        if self.tempdir==None:
            self.tempdir='c:\\temp\\'
        basepart = "read_from_exchange"
        self.attachmentdirectory=self.tempdir+basepart
        if os.path.exists(self.attachmentdirectory):
            files = glob.glob(self.attachmentdirectory+'\\*')
            for f in files:
                os.remove(f)
        else:
            os.mkdir(self.attachmentdirectory)
            
        if numberemails==0:
            numberemails=0.001
        
        increment=100.0/(numberemails*len(folderstoread))
        completed=0
        loop = QtCore.QEventLoop()   
        ## Need to make sure the email text and doc text variables are empty
        self.initialiseTextFields()

        
        outlook =win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        for currfolder in folderstoread:
            folder = outlook.GetDefaultFolder(currfolder)
            
            messages=folder.Items
            try:
                message=messages.GetLast() #Start from top on inbox
            except:
                return
                
            i=0
            while message:

                if i>= numberemails:
                    break
                completed += increment
                progressbar.setValue(completed)
                QtCore.QTimer.singleShot(0, loop.exit)
                loop.exec_()
                
                #print (message.subject)
                try:
                    messageid=message.CreationTime.strftime("%Y-%m-%d %Hh%M")+" "+message.subject[:10]
                    
                    ## Take the message Body into the array
                    if(len(message.Body.strip())):
                        self.emailtext_original.append(message.Body.strip())
                        ##new need to add the email name from whence this comes so we can see interesting e-mails later
                        self.emailtext_emailname.append(messageid)  
                    
                    
                    if readattachments:
                        for attachment in message.Attachments:
         
                            if attachment.Type == 1:
                                ## Think we have an attachment now...
                                ## Need to check the attachment type
                                if attachment.FileName.endswith('.doc') or attachment.FileName.endswith('.docx') or attachment.FileName.endswith('.rtf'):
                                    filepath=self.attachmentdirectory+'\\'+message.ReceivedTime.strftime("%Y-%m-%d %Hh%M")+" "+attachment.FileName.lower()
                                    attachment.SaveAsFile (filepath)
                                    ReadText=self._read_text_msword(filepath) 
                                    
                                    ##We need to treat eachc attachement seperately
                                    if(len(ReadText.strip())):
                                        self.doctext_original.append(ReadText.strip())
                                        self.doctext_emailname.append(filepath)
                                if attachment.FileName.endswith('.pdf') and dopdf==True:
                                    filepath=self.attachmentdirectory+'\\'+message.ReceivedTime.strftime("%Y-%m-%d %Hh%M")+" "+attachment.FileName.lower()
                                    attachment.SaveAsFile (filepath)
                                    ReadText=self._read_text_pdf(filepath)
                                    if(len(ReadText.strip())):
                                        self.doctext_original.append(ReadText.strip())
                                        self.doctext_emailname.append(filepath)
                                    
                                    
                except:
                    next
                try:        
                    message=messages.GetPrevious()
                    i+=1
                except:
                    next
        return True
            

    def is_document_file(self,filename, extensions=['.doc', '.docx', '.rtf', '.pdf']):
        return any(filename.endswith(e) for e in extensions)   

    def readLocalComputerDocuments(self,doPDF=True,doWord=True):
        #Progressbar is for whenb we run in a go ...
 
        
        if self.readdrivedirectory==None:
            
            return False

        if not os.path.exists(self.readdrivedirectory):
            #can't construct the class this is an error
            print('Computer directory {} does not exist'.format(self.readdrivedirectory))
            
#            raise Exception('.eml directory {} does not exist'.format(self.emaildirectory))
            return False
      
        ## Need to make sure the email text and doc text variables are empty
        self.initialiseTextFields()
        documents=[]
        if doWord==True:
            documents.append('*.docx')
        if doPDF==True:
            documents.append('*.pdf')
        
        for subdir, dirs, files in os.walk(self.readdrivedirectory):
            for extension in documents:             
                for file in fnmatch.filter(files,extension):       
                    filepath = subdir + os.sep + file
                    
                    #print os.path.join(subdir, file)
                        
                    filepath = subdir + os.sep + file
                    #print("current file : ",filepath)
                        
                    ##Now check if 
                    if filepath.endswith(".docx"):
                        #print (filepath)
                       DocText=[] 
                       

                       try:
                           doc = docx.Document(filepath)
                           for para in doc.paragraphs:
                               DocText.append(para.text)
                           ReadText='\n'.join(DocText)
                       except:
                           ReadText=""

                       
                       if(len(ReadText.strip())):
                            self.doctext_original.append(ReadText.strip())
                            self.doctext_emailname.append(filepath)                            
                    if (filepath.endswith(".pdf") and  self.dopdf==True):  
                                        
                        ReadText=self._read_text_pdf(filepath)                    
                        if(len(filepath.strip())):
                            self.doctext_original.append(ReadText.strip())
                            self.doctext_emailname.append(file)
                                        
                           
        return True


    def readComputerDocuments(self,progressbar):
        #Progressbar is for whenb we run in a go ...
 
        
        if self.readdrivedirectory==None:
            #can't construct the class this is an error
            progressbar.close()
            QMessageBox.warning(progressbar.parent(),'Reading documents',str('Computer directory not yet specified...'))
            
            return False

        if not os.path.exists(self.readdrivedirectory):
            #can't construct the class this is an error
            progressbar.close()
            QMessageBox.warning(progressbar.parent(),"Reading documents", 'Computer directory {} does not exist'.format(self.readdrivedirectory))
            
#            raise Exception('.eml directory {} does not exist'.format(self.emaildirectory))
            return False
      
        ## Need to make sure the email text and doc text variables are empty
        self.initialiseTextFields()
        documents=['*.doc', '*.docx', '*.rtf', '*.pdf']    
        numfiles = 0
        for root, dirs, files in os.walk(self.readdrivedirectory):
            for file in files:   
                
                if file.endswith('.doc') or file.endswith('.docx') or file.endswith('.rtf') or file.endswith('.pdf'):
                    numfiles += 1
        if numfiles==0:
            numfiles=0.0001
        increment=100.0/numfiles
        completed=0
        for subdir, dirs, files in os.walk(self.readdrivedirectory):
            for extension in documents:             
                for file in fnmatch.filter(files,extension):       
                    loop = QtCore.QEventLoop()    
                    completed += increment
                    progressbar.setValue(completed)
                    filepath = subdir + os.sep + file
                    
                    #print os.path.join(subdir, file)
                        
                    filepath = subdir + os.sep + file
                    #print("current file : ",filepath)
                        
                    ##Now check if 
                    if filepath.endswith(".doc") or filepath.endswith(".docx") or filepath.endswith("*.rtf"):
                        #print (filepath)
                        progressbar.setLabelText("Reading: "+filepath)
                        QtCore.QTimer.singleShot(0, loop.exit)
                        loop.exec_()
                        ReadText=self._read_text_msword(filepath) 
        
                        if(len(ReadText.strip())):
                            self.doctext_original.append(ReadText.strip())
                            self.doctext_emailname.append(filepath)                            
                    if (filepath.endswith(".pdf") and  self.dopdf==True):  
                        progressbar.setLabelText("Reading: "+filepath)
                        QtCore.QTimer.singleShot(0, loop.exit)
                        loop.exec_()                              
                        ReadText=self._read_text_pdf(filepath)                    
                        if(len(filepath.strip())):
                            self.doctext_original.append(ReadText.strip())
                            self.doctext_emailname.append(file)
                                        
                           
        return True


        
    def readEmails(self,progressbar,readattachments=False):
        #Progressbar is for whenb we run in a go ...
        if self.tempdir==None:
            self.tempdir='c:\\temp\\'
        # take the last bit of directory and create tempdir+lastbit
        
        if self.emaildirectory==None:
            #can't construct the class this is an error
            progressbar.close()
            QMessageBox.warning(progressbar.parent(),'Reading emails',str('Email directory not yet specified...'))
            
            return False

        if not os.path.exists(self.emaildirectory):
            #can't construct the class this is an error
            progressbar.close()
            QMessageBox.warning(progressbar.parent(),"Reading emails", '.eml directory {} does not exist'.format(self.emaildirectory))
            
#            raise Exception('.eml directory {} does not exist'.format(self.emaildirectory))
            return False

        basepart = os.path.basename(os.path.normpath(self.emaildirectory))
        self.attachmentdirectory=self.tempdir+basepart
        #check if the directory exists... if so clear it... if not create it
        if os.path.exists(self.attachmentdirectory):
            files = glob.glob(self.attachmentdirectory+'\\*')
            for f in files:
                os.remove(f)
        else:
            os.mkdir(self.attachmentdirectory)
        
        ## Need to make sure the email text and doc text variables are empty
        self.initialiseTextFields()
        
        mswordtype=["application/msword",
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    "application/rtf"]
        
        
        for subdir, dirs, files in os.walk(self.emaildirectory):
            numfiles=len(files) # for progress bar
            if numfiles==0:
                numfiles=0.0001
            increment=100.0/numfiles
            completed=0
            loop = QtCore.QEventLoop()            
            for file in files:
                completed += increment
                progressbar.setValue(completed)
                QtCore.QTimer.singleShot(0, loop.exit)
                loop.exec_()
                #print os.path.join(subdir, file)
                
                filepath = subdir + os.sep + file
                if filepath.endswith(".eml"):
                    #print (filepath)
                    Text="" # Will concatenate all Text into here
                    
                    #Now need to do error handling here as AV package might stop me getting to some files
                    try:
                        fp = open(filepath,'rb')

                        
                        msg = BytesParser(policy=policy.default).parse(fp)
                        for part in msg.walk():
        
                            ##Now depending on content type extract what we need to..
                            cp = part.get_content_type()
                            if cp=="text/plain": 
                                Text += part.get_content()
                            if (cp in mswordtype and readattachments==True):
                                
                                fn=part.get_filename() 
                                #append the e-mail number to this so we know from which email this attachment is
                                fn=file+'_'+fn
                                ## Now this is important... 
                                ## Remember this is a typosquatting inbox so need to save to disk to make sure your local AV picks up any nasties
                                ## This would not have gone via your normal AV hygiene...
                                filesaved=self._save_file(fn, part.get_payload(decode=True))
                                ReadText=self._read_text_msword(filesaved) 
                                
                                ##We need to treat eachc attachement seperately
                                if(len(ReadText.strip())):
                                    self.doctext_original.append(ReadText.strip())
                                    self.doctext_emailname.append(file)
                            if (cp=="application/pdf" and readattachments==True and self.dopdf==True) :
                                fn=part.get_filename() 
                                #append the e-mail number to this so we know from which email this attachment is
                                fn=file+'_'+fn
                                ## Now this is important... 
                                ## Remember this is a typosquatting inbox so need to save to disk to make sure your local AV picks up any nasties
                                ## This would not have gone via your normal AV hygiene...
                                filesaved=self._save_file(fn, part.get_payload(decode=True))
                                ReadText=self._read_text_pdf(filesaved) 
                                
                                ##We need to treat eachc attachement seperately
                                if(len(ReadText.strip())):
                                    self.doctext_original.append(ReadText.strip())
                                    self.doctext_emailname.append(file)
                                
                                
                                
                                
                                
                        if(len(Text.strip())):
                            self.emailtext_original.append(Text)
                            ##new need to add the email name from whence this comes so we can see interesting e-mails later
                            self.emailtext_emailname.append(file)                           

                        fp.close()
                    except:
                        next
        return True
        
    def tfidfVectorise(self,max_df=0.9,min_df=5.0,no_txt_features=10000,no_doc_features=10000,regex_pattern="",ngramrange="1,1"):
        self.tfidfVectText=None
        self.tfidfVectTextInitialized=False
        self.tfidfVectorizerText=None
        self.tfidfTextFeatureNames=None
        self.currentEmailVectorizer=None
        self.currentEmailVectorised=None
        
        self.tfidfVectDocText=None
        self.tfidfVectDocTextInitialized=False
        self.tfidfVectorizerDocText=None
        self.tfidfVectDocText=None
        self.tfidfDocTextFeatureNames=None
        self.currentEmailVectorizer=None
        self.currentDocVectorizer=None
        
        if(len(self.emailtext_forvectorize)):
            #now we have already stripped out english stop words earlier so here we can strip out optional other stop words if stoplist was passed
            self.tfidfVectorizerText = TfidfVectorizer(analyzer='word',token_pattern=regex_pattern,max_df=max_df, min_df=min_df, max_features=no_txt_features, stop_words=None,ngram_range=tuple(map(int,ngramrange.split(','))))
            
            #might not have enough e-mails to satisfy degrees of freedom so do some error handling here:
            try:
                self.tfidfVectText = self.tfidfVectorizerText.fit_transform(self.emailtext_forvectorize)
                self.tfidfTextFeatureNames = self.tfidfVectorizerText.get_feature_names()
                self.tfidfVectTextInitialized=True
                self.currentEmailVectorizer=self.tfidfVectorizerText
                self.currentEmailVectorised=self.tfidfVectText
   
            except:
                self.tfidfVectText=None
                self.tfidfVectTextInitialized=False
                self.tfidfVectorizerText=None
                self.tfidfVectDocText=None
                self.tfidfTextFeatureNames=None
                self.currentEmailVectorizer=None
        
                
                
            
        if(len(self.doctext_forvectorize)):
            #now we have already stripped out english stop words earlier so here we can strip out optional other stop words if stoplist was passed
            self.tfidfVectorizerDocText = TfidfVectorizer(token_pattern=regex_pattern,max_df=max_df, min_df=min_df, max_features=no_doc_features, stop_words=None,ngram_range=tuple(map(int,ngramrange.split(','))))
            
            #might not have enough attachments to satisfy the min degrees of freedom here so do a try with fit
            try:
                self.tfidfVectDocText = self.tfidfVectorizerDocText.fit_transform(self.doctext_forvectorize)
                self.tfidfDocTextFeatureNames = self.tfidfVectorizerDocText.get_feature_names()   
                self.tfidfVectDocTextInitialized=True
                self.currentDocVectorizer=self.tfidfVectorizerDocText
                self.currentDocVectorised=self.tfidfVectDocText
            except:
                ## reset things to what they were before we tried to vectorize
                self.tfidfVectDocText=None
                self.tfidfVectDocTextInitialized=False
                self.tfidfVectorizerDocText=None
                self.tfidfVectDocText=None
                self.tfidfDocTextFeatureNames=None
                self.currentDocVectorizer=None
                
                
        return self.tfidfVectText,self.tfidfVectDocText
                
                    
    def countVectorise(self,max_df=0.95,min_df=2,RemoveList=None,no_txt_features=1000,no_doc_features=1000,regex_pattern="",ngramrange="1,1"):
        self.countVectText=None
        self.countVectTextInitialized=False
        self.countVectorizerText=None
        self.countVectText=None
        self.countfTextFeatureNames=None
        
        self.countVectDocText=None
        self.countVectDocTextInitialized=False
        self.countVectorizerDocText=None
        self.countVectDocText=None
        self.countDocTextFeatureNames=None

        self.currentDocVectorizer=None 
        self.currentEmailVectorizer=None
        self.currentDocVectorised=None
        self.currentEmailVectorised=None
        
        if(len(self.emailtext_forvectorize)):
            #now we have already stripped out english stop words earlier so here we can strip out optional other stop words if stoplist was passed
            self.countVectorizerText = CountVectorizer(max_df, min_df, max_features=no_txt_features, stop_words=None,token_pattern=regex_pattern,decode_error='ignore',ngram_range=tuple(map(int,ngramrange.split(','))))
            self.countVectText = self.countVectorizerText.fit_transform(self.emailtext_forvectorize)
            self.countTextFeatureNames = self.countVectorizerText.get_feature_names()
            self.countVectTextInitialized=True
            self.currentEmailVectorizer=self.countVectorizerText 
            self.currentEmailVectorised=self.countVectText
            
        if(len(self.doctext_forvectorize)):
            #now we have already stripped out english stop words earlier so here we can strip out optional other stop words if stoplist was passed
            self.countVectorizerDocText = CountVectorizer(max_df, min_df, max_features=no_doc_features, stop_words=None,token_pattern=regex_pattern,decode_error='ignore',ngram_range=tuple(map(int,ngramrange.split(','))))
            self.countVectDocText = self.countVectorizerDocText.fit_transform(self.doctext_forvectorize)
            self.countDocTextFeatureNames = self.countVectorizerDocText.get_feature_names()   
            self.countVectDocTextInitialized=True
            self.currentDocVectorizer=self.countVectorizerDocText 
            self.currentDocVectorised=self.countVectDocText
        return self.countVectText,self.countVectDocText
                    
   
    def populateTopicDataframe(self):
        # Styling
       
        if not(self.currentEmailVectorised==None or self.text_lda==None):
            lda_output_text = self.text_lda.transform(self.currentEmailVectorised)
            topicnames = ["Topic" + str(i) for i in range(len(self.text_lda.components_))]
            self.df_email_TextTopics = pd.DataFrame(np.round(lda_output_text, 2), columns=topicnames, index=self.emailtext_emailname)
            dominant_topic = np.argmax(self.df_email_TextTopics.values, axis=1)
            self.df_email_TextTopics['dominant_topic'] = dominant_topic
            
            self.df_email_text_topic_distribution = self.df_email_TextTopics['dominant_topic'].value_counts().reset_index(name="Num Documents")
            
            self.df_email_text_topic_distribution.columns = ['Topic_Num', 'Num_Documents']
            
        if not(self.currentDocVectorised==None or self.doctext_lda==None):
            lda_output_doc = self.doctext_lda.transform(self.currentDocVectorised)
            topicnames = ["Topic" + str(i) for i in range(len(self.doctext_lda.components_))]
            self.df_email_DocTopics = pd.DataFrame(np.round(lda_output_doc, 2), columns=topicnames, index=self.doctext_emailname)
            dominant_topic = np.argmax(self.df_email_DocTopics.values, axis=1)
            self.df_email_DocTopics['dominant_topic'] = dominant_topic
                        


            self.df_email_doc_topic_distribution = self.df_email_DocTopics['dominant_topic'].value_counts().reset_index(name="Num Documents")
            
            self.df_email_doc_topic_distribution.columns = ['Topic_Num', 'Num_Documents']

# Get dominant topic for each document







        
    def doNMF(self,text_vectorized=None,doc_text_vectorized=None,text_n_components=10, text_alpha=0.1, text_l1_ratio=.5, 
                     doc_n_components=10,doc_alpha=0.1,doc_l1_ratio=0.5,
                     random_state=1,init='nndsvd'):
            

        self.text_nmf=None
        self.doctext_nmf=None
        
        if not(text_vectorized==None):
            self.text_nmf = NMF(n_components=text_n_components, random_state=random_state,alpha=text_alpha, l1_ratio=text_l1_ratio, init=init).fit(text_vectorized)
            transformed=self.text_nmf.transform(text_vectorized)
            self.emailtext_topic=np.argmax(transformed,axis=1)
        if not(doc_text_vectorized==None):
            self.doctext_nmf = NMF(n_components=doc_n_components, random_state=random_state,alpha=doc_alpha, l1_ratio=doc_l1_ratio, init=init).fit(doc_text_vectorized)
            transformed=self.doctext_nmf.transform(doc_text_vectorized)
            self.doctext_topic=np.argmax(transformed,axis=1)
        return self.text_nmf,self.doctext_nmf
        
    def doLDA(self,text_vectorized=None,doc_text_vectorized=None,text_n_components=10, doc_n_components=10, max_iter=10, 
              learning_method='batch',learning_decay=0.7,batch_size=128,perp_tol=0.1,mean_change_tol=0.001,evaluate_every=0):

        self.text_lda=None
        

        self.doctext_lda=None
        
        if not(text_vectorized==None):
            #depending on learning method call with slightly different parameters
            if self.LDALearningMethod=="online":
                self.text_lda = LatentDirichletAllocation(n_components=text_n_components,max_iter=10, 
                                                          learning_method='online',learning_decay=learning_decay,
                                                          evaluate_every=evaluate_every,batch_size=batch_size,perp_tol=perp_tol,
                                                          mean_change_tol=mean_change_tol,random_state=100).fit(text_vectorized)
            elif self.LDALearningMethod=="batch":
                self.text_lda = LatentDirichletAllocation(n_components=text_n_components,max_iter=10, 
                                                          learning_method='batch',
                                                          evaluate_every=evaluate_every,perp_tol=perp_tol,
                                                          mean_change_tol=mean_change_tol,random_state=100).fit(text_vectorized)
                
            #use transform here to get all the topics for each for the emails
            transformed=self.text_lda.transform(text_vectorized)
            self.emailtext_topic=np.argmax(transformed,axis=1)
            
            
        if not(doc_text_vectorized==None):
            if self.LDALearningMethod=="online":
                self.doctext_lda = LatentDirichletAllocation(n_components=doc_n_components,max_iter=10, learning_method='online',
                                                             learning_decay=learning_decay,
                                                             evaluate_every=evaluate_every,batch_size=batch_size,perp_tol=perp_tol,
                                                             mean_change_tol=mean_change_tol,random_state=100).fit(doc_text_vectorized)
            elif self.LDALearningMethod=="batch":
                self.doctext_lda = LatentDirichletAllocation(n_components=text_n_components,max_iter=10, 
                                                          learning_method='batch',
                                                          evaluate_every=evaluate_every,perp_tol=perp_tol,
                                                          mean_change_tol=mean_change_tol,random_state=100).fit(doc_text_vectorized)
            
            transformed=self.doctext_lda.transform(doc_text_vectorized)
            self.doctext_topic=np.argmax(transformed,axis=1)
        
        self.populateTopicDataframe()
        
        return self.text_lda,self.doctext_lda    
    
    def GridSearchLDA(self,gridsearch_args={'n_components': [10, 15, 20, 25, 30], 'learning_decay': [.5, .7, .9]}):
        self.text_lda=None
        self.doctext_lda=None
        model_text=None
        model_doctext=None
        besttextparams=""
        bestdoctextparams=""
        grid_args=ast.literal_eval(gridsearch_args)
        if not(self.currentEmailVectorised==None):
            
            self.text_lda = LatentDirichletAllocation()
            model_text = GridSearchCV(self.text_lda,param_grid=grid_args)
            model_text.fit(self.currentEmailVectorised)
            self.text_lda=model_text.best_estimator_
            besttextparams=PrettyPrinter().pformat(model_text.best_params_)
        if not(self.currentDocVectorised==None):
            self.doctext_lda = LatentDirichletAllocation()
            model_doctext = GridSearchCV(self.doctext_lda,param_grid=grid_args)
            model_doctext.fit(self.currentDocVectorised)
            self.doctext_lda=model_doctext.best_estimator_
            bestdoctextparams=PrettyPrinter().pformat(model_doctext.best_params_)
        return self.text_lda,besttextparams,self.doctext_lda,bestdoctextparams
            
     

class Ui(QtWidgets.QMainWindow):

    def __init__(self):
        
        #All the vaiables that will be used by the buttons
        PyQt5.QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling, True)        
        super(Ui, self).__init__()
        self.ui=Ui_Form()
        self.ui.setupUi(self)   

        #now do the bindings for the buttons
        ## First three is sources, i.e. choose directory with .eml files, connect to exchange, read from drive on computer
        self.ui.readEmailsAttachmentsExchange.clicked.connect(self.readEmailsAttachmentsExchange)
        self.ui.choosedirectory.clicked.connect(self.choosedotemlsourcedir)
        self.ui.readfromdrive.clicked.connect(self.choosecomputersourcedir)
        
        
        self.ui.Doattachments.clicked.connect(self.doattachmentsclicked)
        
        self.ui.dotopicmodel.clicked.connect(self.doTopicModelling)
        self.ui.docsmatchtopicdisplay.clicked.connect(self.showdocmatchtopic)
        self.ui.emailmatchtopicdisplay.clicked.connect(self.showemailmatchtopic)

        self.ui.setemailvisfile.clicked.connect(self.choosetxtvisfile)
        self.ui.setdocvisfile.clicked.connect(self.choosedocvisfile)
        self.ui.saveprocessed.clicked.connect(self.chooseprocessedsavefile)
        self.ui.loadprocessed.clicked.connect(self.chooseprocessedloadfile)
        self.ui.savewords.clicked.connect(self.choosewordssavefile)
        self.ui.loadwords.clicked.connect(self.choosewordsloadfile)      
        self.ui.cleantext.clicked.connect(self.cleantextclicked)
        self.ui.processtext.clicked.connect(self.processclicked)
        
        self.ui.savetopicmodel.clicked.connect(self.saveTopicModel)
        
        self.ui.AlgorithmCombo.currentIndexChanged.connect(self.ChooseAlgorithm)
        self.ui.ldalearningmethod.currentIndexChanged.connect(self.ChooseLDALearningMethod)
        self.setWindowTitle("Typosquatting topic modelling RSAC 2019")
        self.setWindowIcon(QtGui.QIcon('sd.png'))
        self.Topic=TopicModel(tempdir="c:\\temp\\")
        self.dir=None
        self.LDADoc=None
        self.LDA=None
        self.NMF=None
        self.NMFDoc=None


        self.Topic.removelist=self.ui.removelist.toPlainText().strip()
        self.Topic.Algorithm=self.ui.AlgorithmCombo.currentText().strip()
        self.hideshowfields()
        self.show()
     
       
    def showMessage(self,title="",text1="",text2=""):
        box = QtWidgets.QMessageBox(self)
        box.setWindowTitle(title)
        box.setText(text1)
        box.setInformativeText(text2)
        box.setBaseSize(QtCore.QSize(600,600))
        box.setStandardButtons(QtWidgets.QMessageBox.Ok)
        box.setDefaultButton(QtWidgets.QMessageBox.Ok)
        box.setIcon(QtWidgets.QMessageBox.Information)
        return box.exec()
    
    def hideshowfields(self):
        
        if self.Topic.Algorithm=="LDA":
            self.ui.emailalphalabel.hide()
            self.ui.docalphalabel.hide()
            self.ui.emailalpha.hide()
            self.ui.docalpha.hide()
            self.ui.emaill1ratio.hide()
            self.ui.docl1ratio.hide()
            self.ui.emaill1ratiolabel.hide()
            self.ui.docl1ratiolabel.hide()
            
            self.ui.ldalearningmethodlabel.show()
            self.ui.ldalearningmethod.show()
            self.ui.maxiterlabel.show()
            self.ui.maxiter.show()
            self.ui.labelgridsearch.show()
            self.ui.dogridsearch.show()
            self.ui.gridsearchargs.show()
            self.ui.labelgridsearchargs.show()
            self.ui.evaluateeverylabel.show()
            self.ui.evaluateevery.show()
            self.ui.meanchangetollabel.show()
            self.ui.meanchangetol.show()
            #Now depending if online or batch hide or show fields:
            if self.Topic.LDALearningMethod=="batch":
                self.ui.perptollabel.show()
                self.ui.perptol.show()
                


               
                self.ui.learningdecaylabel.hide()
                self.ui.learningdecay.hide()
                self.ui.batchsizelabel.hide()
                self.ui.batchsize.hide()
            else:
                self.ui.learningdecaylabel.show()
                self.ui.learningdecay.show()
                self.ui.batchsizelabel.show()
                self.ui.batchsize.show()
                
                self.ui.perptollabel.hide()
                self.ui.perptol.hide()
        
        elif self.Topic.Algorithm=="NMF":
            self.ui.evaluateevery.hide()
            self.ui.evaluateeverylabel.hide()
            self.ui.meanchangetollabel.hide()
            self.ui.meanchangetol.hide()
            self.ui.perptollabel.hide()
            self.ui.perptol.hide()
            self.ui.batchsizelabel.hide()
            self.ui.batchsize.hide()
            
            self.ui.ldalearningmethodlabel.hide()
            self.ui.ldalearningmethod.hide()
            self.ui.maxiterlabel.hide()
            self.ui.maxiter.hide()
            self.ui.labelgridsearch.hide()
            self.ui.dogridsearch.hide()
            self.ui.gridsearchargs.hide()
            self.ui.labelgridsearchargs.hide()
            self.ui.learningdecaylabel.hide()
            self.ui.learningdecay.hide()
            
            self.ui.emailalphalabel.show()
            self.ui.docalphalabel.show()
            self.ui.emailalpha.show()
            self.ui.docalpha.show()
            self.ui.emaill1ratio.show()
            self.ui.docl1ratio.show()
            self.ui.emaill1ratiolabel.show()
            self.ui.docl1ratiolabel.show()
     
      
    def ChooseAlgorithm(self):
        self.Topic.Algorithm=self.ui.AlgorithmCombo.currentText().strip()
        #now hide certain controls
        self.hideshowfields()
        

    def ChooseLDALearningMethod(self):
        self.Topic.LDALearningMethod=self.ui.ldalearningmethod.currentText().strip()
        #hide or display controls depending on what was selected
        self.hideshowfields()
        
        
    def saveTopicModel(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getSaveFileName(self,"File to save Topic Model info in..","",filter=('PKL file (*.pkl)'), options=options)
        if fileName:
            self.ui.topicmodelsavefilename.setText(fileName)
            ##now fill out the data structure will all the elements of the Topic model
            #def __init__(self,stopwords,vectoriser,cleantext,cleantextpattern,processtext,lemmapattern,NMFModel=None, LDAModel=None):
            #Need everyting here to process future text to see which topic it belongs to so
            
            toSaveTopicModel=TopicModeltoPickle(removestopwords=self.Topic.removewords,
                                                stopwords=self.Topic.removelist,
                                                email_vectoriser=self.Topic.currentEmailVectorizer,
                                                doc_vectoriser=self.Topic.currentDocVectorizer,
                                                cleantext=self.Topic.cleantext,
                                                processtext=self.Topic.processtext,
                                                lemmapattern=self.Topic.lemma_args,
                                                NMFEmailModel=self.Topic.text_nmf,
                                                NMFDocModel=self.Topic.doctext_nmf,
                                                LDAEmailModel=self.Topic.text_lda,
                                                LDADocModel=self.Topic.doctext_lda)
            afile = open(fileName, 'wb')
            pickle.dump(toSaveTopicModel, afile)
            afile.close()            

        

    def chooseprocessedsavefile(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getSaveFileName(self,"File to save eMail info in..","",filter=('PKL file (*.pkl)'), options=options)
        if fileName:
            #print(fileName)
            #get the variables to save like RemoveList,doattach,regex,removewords, emaildir

            #now get the status of the clean_text and do_attachments click boxes
            
            self.ui.saveprocessedfilename.setText(fileName)
            #now we need to save into this file...
            picklethis=toPickle(self.Topic.emailtext_original,self.Topic.emailtext_emailname,
                                self.Topic.doctext_original,self.Topic.doctext_emailname)
                                
            afile = open(fileName, 'wb')
            pickle.dump(picklethis, afile)
            afile.close()
    
    def choosewordssavefile(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getSaveFileName(self,"File to save remove word list to","",filter=('PKL file (*.pkl)'), options=options)
        if fileName:
            self.Topic.removelist=self.ui.removelist.toPlainText().strip()
            print(fileName)
            #get the variables to save like RemoveList,doattach,regex,removewords, emaildir

            #now get the status of the clean_text and do_attachments click boxes
            
            self.ui.savewordsfilename.setText(fileName)
            #now we need to save into this file...
                               
            afile = open(fileName, 'wb')
            pickle.dump(self.Topic.removelist, afile)
            afile.close()
            
            
    def choosewordsloadfile(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getOpenFileName(self,"File to load word list from","",filter=('PKL file (*.pkl)'), options=options)
        if fileName:
            print(fileName)
            self.ui.loadwordsfilename.setText(fileName)
            #now we need to save into this file...
            afile = open(fileName, 'rb')
            self.Topic.removelist=str(pickle.load(afile))
            afile.close()

            ## now back populate all the options as they were when this was saved.. 
            self.ui.removelist.setPlainText(self.Topic.removelist)
            
    def chooseprocessedloadfile(self):

        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getOpenFileName(self,"File to load eMail info in..","",filter=('PKL file (*.pkl)'), options=options)
        unpicklethis=toPickle("","","","")
        if fileName:
            self.Topic.initialiseTextFields()
            #print(fileName)
            self.ui.loadprocessedfilename.setText(fileName)
            #now we need to save into this file...
            afile = open(fileName, 'rb')
            unpicklethis=pickle.load(afile)
            afile.close()

            ## now back populate all the options as they were when this was saved.. 
             
            self.Topic.emailtext_original=unpicklethis.email_txt
            self.Topic.emailtext_emailname=unpicklethis.email_txt_emailname
            self.Topic.doctext_original=unpicklethis.doc_txt
            self.Topic.doctext_emailname=unpicklethis.doc_txt_emailname   

            self.Topic.resetVectorisersLDA()            
            self.ui.emailtopics.setText("")
            self.ui.doctopics.setText("")
            
            self.ui.emailsmatchingdoctopic.setText("")
            self.ui.emailsmatchingtexttopic.setText("")
        
            
    def choosedotemlsourcedir(self):
        file = str(QFileDialog.getExistingDirectory(self, "Select Directory"))
        self.ui.sourcedir.setText(file)
        self.Topic.emaildirectory=file
        self.readEmailsAttachments()

    def choosecomputersourcedir(self):
        file = str(QFileDialog.getExistingDirectory(self, "Select Directory to read documents from"))
        self.ui.readdrivesourcedir.setText(file)
        self.Topic.readdrivedirectory=file
        self.readComputerDocuments()        
        
    def choosetxtvisfile(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getSaveFileName(self,"Choose eMail visualation file","",filter=('HTMLfile (*.htm*)'), options=options)
        if fileName:
            print(fileName)
            self.ui.emailpyldavisfilename.setText(fileName)
            return fileName
        else:
            self.ui.emailpyldavisfilename.setText("")
            return ""

    def choosedocvisfile(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getSaveFileName(self,"Choose attachment visualisation file","",filter=('HTMLfile (*.htm*)'), options=options)
        if fileName:
            print(fileName)
            self.ui.pyldavisdocfilename.setText(fileName)
        else:
            self.ui.pyldavisdocfilename.setText("")

        
    def showdocmatchtopic(self):
        topicnumber=int(self.ui.docsmatchingtopicnum.value())
        docs=self.Topic.displayEmailsMatchingDocTopic(topicnumber)
        self.ui.emailsmatchingdoctopic.setText(",".join(docs))
        
    def showemailmatchtopic(self):
        topicnumber=int(self.ui.emailsmatchingtopicnum.value())
        emails=self.Topic.displayEmailsMatchingTextTopic(topicnumber)
        self.ui.emailsmatchingtexttopic.setText(",".join(emails))
        


    def cleantextclicked(self):
        self.Topic.cleantext=not(self.ui.cleantext.checkState()==0)
        
    def doattachmentsclicked(self):
        self.Topic.doattachments=not(self.ui.Doattachments.checkState()==0)

    def processclicked(self):
        self.Topic.processtext=not(self.ui.processtext.checkState()==0)

    def readEmailsAttachmentsExchange(self):
        #First get the folder(s) we need to read:
        folders=self.ui.exchangefoldernumbers.text().strip()
        numberemails=int(self.ui.numberemailsload.text().strip())
        self.Topic.doattachments=not(self.ui.Doattachments.checkState()==0)
        self.Topic.dopdf=not(self.ui.dopdf.checkState()==0) 
        if len(folders):
            folderstoread=[int(i) for i in folders.split(",")]
        if self.Topic.doattachments:
            msg="Finished reading eMails, Doc attachments from Exchange"
            progress="Reading eMails, Doc attachments from Exchange"

        else:
            msg="Finished reading eMails from Exchange"
            progress="Reading eMails from Exchange"
            
            
        self.progress=QProgressDialog(progress,None,0,100,self)
        self.progress.setWindowTitle("Working on it")
        self.progress.setWindowModality(QtCore.Qt.WindowModal)
        self.progress.show()
        self.progress.adjustSize()
        self.progress.setValue(0)
        
        if (self.Topic.readEmailsExchange(self.progress,folderstoread,numberemails,readattachments=self.Topic.doattachments,dopdf=self.Topic.dopdf)):
            self.progress.close()
            #Need a popup so say this is done...
            QMessageBox.about(self, "Topic modelling", msg)     
        #print("We are done...")
               
                
        
        
        
    def readEmailsAttachments(self):    
        self.Topic.doattachments=not(self.ui.Doattachments.checkState()==0) 
        self.Topic.dopdf=not(self.ui.dopdf.checkState()==0) 
        if self.Topic.doattachments:
            msg="Finished reading eMails, Doc attachments"
            progress="Reading eMails, Doc attachments"

        else:
            msg="Finished reading eMails"
            progress="Reading eMails"
            
        self.progress=QProgressDialog(progress,None,0,100,self)
        self.progress.setWindowTitle("Working on it")
        self.progress.setWindowModality(QtCore.Qt.WindowModal)
        self.progress.show()
        self.progress.adjustSize()
        self.progress.setValue(0)
        if (self.Topic.readEmails(self.progress,readattachments=self.Topic.doattachments)):
            self.progress.close()
            #Need a popup so say this is done...
            QMessageBox.about(self, "Topic modelling", msg)   
        else:
            self.progress.close()              
            
    def readComputerDocuments(self):    
        self.Topic.dopdf=not(self.ui.dopdf.checkState()==0)        
        if self.Topic.dopdf:
            msg="Finished reading Doc(x), RTFs & PDFs"
            progress="Reading Doc & PDFs"

        else:
            msg="Finished reading Docs"
            progress="Reading Doc(x) & RTFs"
            
        self.progress=QProgressDialog(progress,None,0,100,self)
        self.progress.setWindowTitle("Working on it")
        self.progress.setWindowModality(QtCore.Qt.WindowModal)
        self.progress.show()
        self.progress.adjustSize()
        self.progress.setValue(0)
        if (self.Topic.readComputerDocuments(self.progress)):
            self.progress.close()
            #Need a popup so say this is done...
            QMessageBox.about(self, "Topic modelling", msg)   
        else:
            self.progress.close() 
    
    def plotDocumentDistribution(self):
        # Plot two distributions
        
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(10, 4), dpi=120)


        # Topic Distribution by Dominant Topics
        if not(self.Topic.df_email_TextTopics.empty):

            
            email_topic_top3words=[]
            for i in range(0,self.Topic.text_number_topics):
                for j in range(0,4):
                    email_topic_top3words.append((i,self.Topic.emailtexttopwords[i][j]))
                    
#            df_email_top3words_stacked=pd.DataFrame(email_topic_top3words,columns=['topic_id','words'])
#            df_email_top3words=df_email_top3words_stacked.groupby('topic_id').agg(', \n'.join)
#            df_email_top3words.reset_index(level=0,inplace=True)                                               
                    
            

            ax1.bar(x='Topic_Num', height='Num_Documents', data=self.Topic.df_email_text_topic_distribution, width=.5, color='firebrick')
            ax1.set_xticks(range(0,self.Topic.text_lda.n_components))
#            tick_formatter=FuncFormatter(lambda x, pos: 'Topic '+str(x)+'\n'+df_email_top3words.loc[df_email_top3words.topic_id==x, 'words'].values[0])
#            ax1.tick_params(axis='both',which='minor', labelsize=5)
#            ax1.xaxis.set_major_formatter(tick_formatter)
            ax1.set_title('Number of Emails by Dominant Topic', fontdict=dict(size=10))
            ax1.set_ylabel('Number of Emails')
            ax1.set_ylim(0, np.max(self.Topic.df_email_text_topic_distribution.Num_Documents)*1.1)
        else:
            fig.delaxes(ax1)

        # Topic Distribution by Topic Weights
        if not(self.Topic.df_email_DocTopics.empty):

#            doc_topic_top3words=[]
#            for i in range(0,self.Topic.doctext_number_topics):
#                for j in range(0,4):
#                    doc_topic_top3words.append((i,self.Topic.emaildoctopwords[i][j]))
                    
#            df_doc_top3words_stacked=pd.DataFrame(doc_topic_top3words,columns=['topic_id','words'])
#            df_doc_top3words=df_doc_top3words_stacked.groupby('topic_id').agg(', \n'.join)
#            df_doc_top3words.reset_index(level=0,inplace=True)    
            
            ax2.bar(x='Topic_Num', height='Num_Documents', data=self.Topic.df_email_doc_topic_distribution, width=.5, color='steelblue')
#            ax2.tick_params(axis='both',which='minor', labelsize=5)
            ax2.set_xticks(range(0,self.Topic.doctext_lda.n_components))
#            tick_formatter=FuncFormatter(lambda x, pos: 'Topic '+str(x)+'\n'+df_doc_top3words.loc[df_doc_top3words.topic_id==x, 'words'].values[0])
#            ax2.xaxis.set_major_formatter(tick_formatter)

    
            ax2.set_title('Number of Documents by Dominant Topic', fontdict=dict(size=10))
            ax2.set_ylabel('Number of Documents')
            ax2.set_ylim(0, np.max(self.Topic.df_email_doc_topic_distribution.Num_Documents)*1.1)
        else:
            fig.delaxes(ax2)
        fig.canvas.set_window_title('Email/Document distribution by dominant topic')          
        plt.show()
    
    def clusterDocumentsSimilarTopic(self):
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 7), dpi=120)
        #fig(num='Cluster Emails/Documents based on similarity by topic')  
        if self.Topic.text_lda:
        
            txt_lda_output=self.Topic.text_lda.transform(self.Topic.currentEmailVectorised)
            if len(txt_lda_output)>self.Topic.text_number_topics:
            
                clusterstxt = KMeans(n_clusters=self.Topic.text_number_topics, random_state=100).fit_predict(txt_lda_output)
        
                # Build the Singular Value Decomposition(SVD) model
                txt_svd_model = TruncatedSVD(n_components=2)  # 2 components
                txt_lda_output_svd = txt_svd_model.fit_transform(txt_lda_output)
                
                # X and Y axes of the plot using SVD decomposition
                x = txt_lda_output_svd[:, 0]
                y = txt_lda_output_svd[:, 1]
    
                ax1.scatter(x, y, c=clusterstxt)
                ax1.set_ylabel('Component 2')
                ax1.set_xlabel('Component 1')
                ax1.set_title("Segregation of Topic Clusters eMail text", )
            else:
                fig.delaxes(ax1)
        else:
            fig.delaxes(ax1)
        if self.Topic.doctext_lda:
            doctxt_lda_output=self.Topic.doctext_lda.transform(self.Topic.currentDocVectorised)
            
            if len(doctxt_lda_output)>self.Topic.doctext_number_topics:
                clustersdoctxt = KMeans(n_clusters=self.Topic.doctext_number_topics, random_state=100).fit_predict(doctxt_lda_output)
        
                # Build the Singular Value Decomposition(SVD) model
                doctxt_svd_model = TruncatedSVD(n_components=2)  # 2 components
                doctxt_lda_output_svd =doctxt_svd_model.fit_transform(doctxt_lda_output)
                
                # X and Y axes of the plot using SVD decomposition
                x = doctxt_lda_output_svd[:, 0]
                y = doctxt_lda_output_svd[:, 1]
    
                ax2.scatter(x, y, c=clustersdoctxt)
                ax2.set_ylabel('Component 2')
                ax2.set_xlabel('Component 1')
                ax2.set_title("Segregation of Topic Clusters eMail Document text", )
            else:
                fig.delaxes(ax2)
        else:
            fig.delaxes(ax2)
        
        fig.canvas.set_window_title('Segregation  of Email/Document text Topic clusters')                   
        plt.show()

    def showWordCloud(self):
        
        cols = [color for name, color in mcolors.TABLEAU_COLORS.items()]  # more colors: 'mcolors.XKCD_COLORS'
        
        cloud = WordCloud(stopwords=self.Topic.stop_words,
                          background_color='white',
                          width=2500,
                          height=1800,
                          max_words=10,
                          colormap='tab10',
                          color_func=lambda *args, **kwargs: cols[i],
                          prefer_horizontal=1.0)
        
       
        if len(self.Topic.emailtexttopwords):        
            #So we plot with 4 columns so divided number of topics by 4 and round up for the the num cols
            rows=int(np.ceil(self.Topic.text_number_topics/4.0))
            
            fig, axes = plt.subplots(rows, 4, figsize=(10,10), sharex=True, sharey=True)
            fig.suptitle('Word cloud for eMail text')
            for i, ax in enumerate(axes.flatten()):
                if i<self.Topic.text_number_topics:
                    fig.add_subplot(ax)
                    topic_words = dict(zip(self.Topic.emailtexttopwords[i],self.Topic.emailtexttopwordweights[i]))
                    cloud.generate_from_frequencies(topic_words, max_font_size=300)
                    plt.gca().imshow(cloud)
                    plt.gca().set_title('Topic ' + str(i), fontdict=dict(size=16))
                    plt.gca().axis('off')
                else:
                    fig.delaxes(ax)
        
            
            plt.subplots_adjust(wspace=0, hspace=0)
            
            plt.axis('off')
            plt.margins(x=0, y=0)
            plt.tight_layout()
            fig.canvas.set_window_title('Word cloud for eMail text')    

            plt.show()


        if len(self.Topic.emaildoctopwords):
            rows=int(np.ceil(self.Topic.doctext_number_topics/4.0))
            
            fig, axes = plt.subplots(rows, 4, figsize=(10,10), sharex=True, sharey=True)
            fig.suptitle('Word cloud for eMail documents')
            for i, ax in enumerate(axes.flatten()):
                if i<self.Topic.doctext_number_topics:
                    fig.add_subplot(ax)
                    topic_words = dict(zip(self.Topic.emaildoctopwords[i],self.Topic.emaildoctopwordweights[i]))
                    cloud.generate_from_frequencies(topic_words, max_font_size=300)
                    plt.gca().imshow(cloud)
                    plt.gca().set_title('Topic ' + str(i), fontdict=dict(size=16))
                    plt.gca().axis('off')
                else:
                    fig.delaxes(ax)
            
            
            plt.subplots_adjust(wspace=0, hspace=0)
            
            plt.axis('off')
            plt.margins(x=0, y=0)
            plt.tight_layout()
            fig.canvas.set_window_title('Word cloud for eMail Document text')    
            plt.show()



    def showWordCountImportance(self):

        
        if len(self.Topic.emailtext_words):
            
            data_flat = [w for w_list in self.Topic.emailtext_words for w in w_list]
            counter = Counter(data_flat)
            
            out = []
            for i in range(0,self.Topic.text_number_topics):
                for j in range(0,len(self.Topic.emailtexttopwords[i])):
                    word=self.Topic.emailtexttopwords[i][j]
                    weight=self.Topic.emailtexttopwordweights[i][j]
                    out.append([word, i , weight, counter[word]])
            
            df = pd.DataFrame(out, columns=['word', 'topic_id', 'importance', 'word_count'])        
            
            # Plot Word Count and Weights of Topic Keywords
            rows=int(np.ceil(self.Topic.text_number_topics/4.0))
            
            fig, axes = plt.subplots(rows, 4, figsize=(12,6), sharey=True, dpi=160)
            cols = [color for name, color in mcolors.TABLEAU_COLORS.items()]
            for i, ax in enumerate(axes.flatten()):
                if i<self.Topic.text_number_topics:
                    ax.bar(x='word', height="word_count", data=df.loc[df.topic_id==i, :], color=cols[i], width=0.5, alpha=0.3, label='Word Count')
                    ax_twin = ax.twinx()
                    ax_twin.bar(x='word', height="importance", data=df.loc[df.topic_id==i, :], color=cols[i], width=0.2, label='Weights')
                    ax.set_ylabel('Word Count', color=cols[i],fontsize=6)
                    ax_twin.set_ylim(0,np.max(df.importance)*1.1); ax.set_ylim(0,np.max(df.word_count)*1.1)

                    ax.set_title('Topic: ' + str(i), color=cols[i], fontsize=6)
                    ax.tick_params(axis='y', left=False,labelsize=8)
                    ax.set_xticklabels(df.loc[df.topic_id==i, 'word'], rotation=30, horizontalalignment= 'right',fontsize=6)
                    ax.legend(loc='upper left',prop={'size': 6}); ax_twin.legend(loc='upper right',prop={'size': 6})
                else:
                    fig.delaxes(ax)
            
            fig.tight_layout(w_pad=2)    
            fig.canvas.set_window_title('Word Count and Importance of Topic Keywords for eMail text')     
            plt.show()

        if len(self.Topic.doctext_words):
            
            data_flat = [w for w_list in self.Topic.doctext_words for w in w_list]
            counter = Counter(data_flat)
            
            out = []
            for i in range(0,self.Topic.doctext_number_topics):
                for j in range(0,len(self.Topic.emaildoctopwords[i])):
                    word=self.Topic.emaildoctopwords[i][j]
                    weight=self.Topic.emaildoctopwordweights[i][j]
                    out.append([word, i , weight, counter[word]])
            
            dfdoc = pd.DataFrame(out, columns=['word', 'topic_id', 'importance', 'word_count'])        
            
            # Plot Word Count and Weights of Topic Keywords
            rows=int(np.ceil(self.Topic.doctext_number_topics/4.0))
            
            fig, axes = plt.subplots(rows, 4, figsize=(12,6), sharey=True, dpi=160)
            cols = [color for name, color in mcolors.TABLEAU_COLORS.items()]
            for i, ax in enumerate(axes.flatten()):
                if i<self.Topic.doctext_number_topics:
                    ax.bar(x='word', height="word_count", data=dfdoc.loc[dfdoc.topic_id==i, :], color=cols[i], width=0.5, alpha=0.3, label='Word Count')
                    ax_twin = ax.twinx()
                    ax_twin.bar(x='word', height="importance", data=dfdoc.loc[dfdoc.topic_id==i, :], color=cols[i], width=0.2, label='Weights')
                    ax.set_ylabel('Word Count', color=cols[i],fontsize=6)
                    ax_twin.set_ylim(0,np.max(dfdoc.importance)*1.1); ax.set_ylim(0,np.max(dfdoc.word_count)*1.1)
                    ax.set_title('Topic: ' + str(i), color=cols[i], fontsize=6)
                    ax.tick_params(axis='y', left=False,labelsize=8)
                    ax.set_xticklabels(dfdoc.loc[dfdoc.topic_id==i, 'word'], rotation=30, horizontalalignment= 'right',fontsize=6)
                    ax.legend(loc='upper left',prop={'size': 6}); ax_twin.legend(loc='upper right',prop={'size': 6})
                else:
                    fig.delaxes(ax)
            
            fig.tight_layout(w_pad=2)    
            fig.canvas.set_window_title('Word Count and Importance of Topic Keywords for eMail Document text')    
            plt.show()
            
    def doTopicModelling(self):

        if not(len(self.Topic.emailtext_original)) and not(len(self.Topic.doctext_original)):
            ##Create warning box saying there is nothing to topic modelling...
            QMessageBox.warning(self,'Topic modelling',str('No content to perform topic modelling on. Read eMails or load previously loaded emails...'))
            return
        ##First we need to do text processing
        ##before we can do text processing needs to get status of all the dialog box items
        self.Topic.removewords=not(self.ui.removewords.checkState()==0)
        self.Topic.cleantext=not(self.ui.cleantext.checkState()==0)
        self.Topic.processtext=not(self.ui.processtext.checkState()==0)
        self.Topic.removelist=self.ui.removelist.toPlainText().strip()  
        self.Topic.lemma_args=self.ui.lemmaargs.text().strip()
        self.Topic.ngram_args=self.ui.ngram_args.text().strip()
        self.Topic.regex_pattern=self.ui.regexpattern.text().strip()
        self.Topic.text_number_topics=self.ui.TopicsEmail.value()
        self.Topic.doctext_number_topics=self.ui.TopicsDoc.value() 


        ### Do a lot of heavy lifting to clean the text
        self.Topic.cleanandProcessText()
        #print ("Have cleaned text")
        VectorizerMethod=self.ui.EmbeddingMethod.currentText().strip()

        Algorithm=self.ui.AlgorithmCombo.currentText().strip()

        emailnumtopics=self.ui.TopicsEmail.value()

        
        docnumtopics=self.ui.TopicsDoc.value()

        
        wordsemailtopic=self.ui.WordsperEmailTopic.value()
        wordsdoctopic=self.ui.WordsDocTopic.value()
        
        mindf=int(self.ui.mindf.text().strip())   ### This is wrong !!
        maxdf=int(self.ui.maxdf.text().strip())
        
        numemailfeatures=int(self.ui.numemailfeatures.text())
        numdocfeatures=int(self.ui.numdocfeatures.text())
        

        
        #pyldavisemail=not(self.ui.dovisemail.checkState()==0)
        pyldavisemailfile=self.ui.emailpyldavisfilename.text().strip()
        if len(pyldavisemailfile):
            pyldavisemail=True
        else:
            pyldavisemail=False
            
        
        #pyldavisdoc=not(self.ui.dovisdoc.checkState()==0)
        pyldavisdocfile=self.ui.pyldavisdocfilename.text().strip()  
        if len(pyldavisdocfile):
            pyldavisdoc=True
        else:
            pyldavisdoc=False
        
        dogridsearch=not(self.ui.dogridsearch.checkState()==0)
        gridsearchparams=self.ui.gridsearchargs.text().strip()
        
        showmodelperformance=not(self.ui.showmodelperformance.checkState()==0)
        showprettypictures=not(self.ui.showprettypictures.checkState()==0)
        showdistributionbytopic=not(self.ui.showdistributionbytopic.checkState()==0)
        showwordcloud=not(self.ui.showwordcloud.checkState()==0)
        showwordcountimportance=not(self.ui.showwordcountimportance.checkState()==0)
        
        
        
        ## So lets first vectorize depending on the vectorizer chosen
        if VectorizerMethod=="Count Vectorizer" and not(self.Topic==None):
            vectorizedtxt,vectorizeddoctxt=self.Topic.countVectorise(min_df=mindf,max_df=maxdf,no_txt_features=numemailfeatures,no_doc_features=numdocfeatures,regex_pattern=self.Topic.regex_pattern,ngramrange=self.Topic.ngram_args) 
            #print("Have count vectorized...")
        elif VectorizerMethod=="Tfidf Vectorizer" and not(self.Topic==None):
            vectorizedtxt,vectorizeddoctxt=self.Topic.tfidfVectorise(min_df=mindf,max_df=maxdf,no_txt_features=numemailfeatures,no_doc_features=numdocfeatures,regex_pattern=self.Topic.regex_pattern,ngramrange=self.Topic.ngram_args) 
            #print("Have tfidf vectorized...")
            

        todisplayemail=""
        todisplaydoc=""
        if   Algorithm=="NMF" and not(self.Topic==None):
            #first get arguments from the form...
            
             self.NMF, self.NMFDoc=self.Topic.doNMF(text_vectorized=vectorizedtxt,doc_text_vectorized=vectorizeddoctxt,text_n_components=emailnumtopics,doc_n_components=docnumtopics)
             #Now display the topics
             if not(self.NMF==None):
                 todisplayemail=self.Topic.stringTopics(NMF,self.Topic.currentEmailVectorizer,no_top_words=wordsemailtopic)
             if not(self.NMFDoc==None):
                              
                 todisplaydoc=self.Topic.stringTopics(self.NMFDoc,self.Topic.currentDocVectorizer,no_top_words=wordsdoctopic)    
              
             
             
        elif Algorithm=="LDA" and not(self.Topic==None):

            if dogridsearch:
                #Do a grid search and display the optimum parameters from the search grid
                self.LDA,text_params,self.LDADoc,doctext_params=self.Topic.GridSearchLDA(gridsearch_args=gridsearchparams)
                #Now display this in an infnormation box
              
                
#                message='Best eMail parameters from: '+gridsearchparams+' is : '+text_params+'\n'+'Best eMail Doc parameters from: '+gridsearchparams+' is : '+doctext_params'+'\n'
                
                message="Given : "+gridsearchparams+"\n"
                message=message+"Best eMail parameters are: "+ text_params+"\n"
                message=message+"Best eMail Doc parameters are: "+doctext_params
                 
                self.showMessage(text1="Grid Search", text2=message) 
                
            
                
                
            else:

                self.Topic.emailtexttopwords[:] = []  
                self.Topic.emailtexttopwordweights[:]=[]
        
                self.Topic.emaildoctopwords[:]=[]
                self.Topic.emaildoctopwordweights[:]=[]
                
                learning=self.ui.ldalearningmethod.currentText().strip()
                #first parameters work with both online and batch
                iterations=int(self.ui.maxiter.value())
                evaluate_every=int(self.ui.evaluateevery.value())
                mean_change_tol=float(self.ui.meanchangetol.value())
                

                
                
                
                #now we have two different learning methods so remember there will be two different calls 
                ##initialize all values
                learning_decay=0.7 #online used in online
                batch_size=128 #only used in online
                perp_tol=0.1 #only used in batch learning

                
                
                if self.Topic.LDALearningMethod=="batch":
                    perp_tol=self.ui.perptol.value()
                    
                elif self.Topic.LDALearningMethod=="online":
                    learning_decay=float(self.ui.learningdecay.value())
                    batch_size=int(self.ui.batchsize.value())
                
                self.LDA, self.LDADoc=self.Topic.doLDA(text_vectorized=vectorizedtxt,doc_text_vectorized=vectorizeddoctxt,
                        text_n_components=emailnumtopics,doc_n_components=docnumtopics,
                        learning_method=learning,max_iter=iterations,learning_decay=learning_decay,
                        batch_size=batch_size,perp_tol=perp_tol,mean_change_tol=mean_change_tol,evaluate_every=evaluate_every)   
              

                self.Topic.populateWordWeightMatrix(no_top_words=10)
                #print("Finished LDA topic modelling")
                #print("About to create panel for text")
                if pyldavisemail:
                    #get right Vectorizer
                    
                    emailpanel = pyLDAvis.sklearn.prepare(self.LDA, vectorizedtxt, self.Topic.currentEmailVectorizer, mds='tsne')
                    ##create in tempdir with filename 
                    pyLDAvis.save_html(emailpanel,pyldavisemailfile)
                if pyldavisdoc:
                    docpanel=pyLDAvis.sklearn.prepare(self.LDADoc,vectorizeddoctxt,self.Topic.currentDocVectorizer,mds='tsne')
                    pyLDAvis.save_html(docpanel,pyldavisdocfile)
                message=""
                if not(self.LDA==None):
                    todisplayemail=self.Topic.stringTopics(self.LDA,self.Topic.currentEmailVectorizer,no_top_words=wordsemailtopic)
                    if showmodelperformance:
                        message = "Log Likelihood: "+str(self.LDA.score(self.Topic.currentEmailVectorised))+"\n"

                        message = message+ "Perplexity: "+str( self.LDA.perplexity(self.Topic.currentEmailVectorised))+"\n"
                        message=message    + "Model parameters:\n"+PrettyPrinter().pformat(self.LDA.get_params())+"\n"
                    


                if not(self.LDADoc==None):  
                    todisplaydoc=self.Topic.stringTopics(self.LDADoc,self.Topic.currentDocVectorizer,no_top_words=wordsdoctopic)    
                    if showmodelperformance:
                        message = message+"Log Likelihood: "+str(self.LDADoc.score(self.Topic.currentDocVectorised))+"\n"
                        message = message+ "Perplexity: "+str( self.LDADoc.perplexity(self.Topic.currentDocVectorised))+"\n"
                        message=message    + "Model parameters:\n"+PrettyPrinter().pformat(self.LDADoc.get_params())
                    
                if not(self.LDADoc==None and self.LDA==None) and showmodelperformance:
                    self.showMessage(text1="Model Peformance",text2=message)                    
        
                if showwordcloud:
                    self.showWordCloud()
                    
                if showdistributionbytopic:
                    self.plotDocumentDistribution()
                    
                if showprettypictures:
                    self.clusterDocumentsSimilarTopic()
                    
                if showwordcountimportance:
                    self.showWordCountImportance()
                    
        #Now update the text field...
        self.ui.emailtopics.setText(todisplayemail)
        self.ui.doctopics.setText(todisplaydoc)

        
def handler_rungui(args):
    app = QtWidgets.QApplication(sys.argv)
    application=Ui()
    application.show()
    
    sys.exit(app.exec_())       
 
def handler_match_topic(args):
    
    ##Need to initialise the topic modelling class 
    Topic=TopicModel(tempdir="c:\\temp\\")




    #print("we are going to match topics now...")
    whattomatch=args.dest_what
    if whattomatch=='a':
        Topic.doattachments=True
    else:
        Topic.doattachments=False
    ##strip down the topics... First split on , 
    topicstodo=[]
    dest_topics=args.dest_topics
    if "-" in dest_topics:
        listtopics=list(map(int,dest_topics.split("-")))
        first=listtopics[0]
        last=listtopics[1]
        for topic in range(first,last+1):
            topicstodo.append(topic)
            #print (topic)
    elif "," in dest_topics:
        listtopics=list(map(int,dest_topics.split(",")))
        for topic in listtopics:
            topicstodo.append(topic)
            #print (topic)
    else:
        #print("Only have to do : ",int(dest_topics))
        topicstodo.append(int(dest_topics))
    ## now lets see if the topic model file to load exists
    unpickleModel=TopicModeltoPickle()
    FileName=args.model_file
    if(os.path.exists(FileName)):
        try:
           
            # handle error here
            file = open(FileName, "rb")
            unpickleModel=pickle.load(file)
            file.close()
        except OSError:
            print("Could not load file: ",args.model_file)
    else:
        print("Could not locate file: ",args.model_file)


    Topic.lemma_args=unpickleModel.lemmapattern
    Topic.removewords=unpickleModel.removestopwords
    Topic.removelist=unpickleModel.stopwords
    Topic.text_lda=unpickleModel.LDAEmailModel
    Topic.doctext_lda=unpickleModel.LDADocModel
    Topic.text_nmf=unpickleModel.NMFEmailModel
    Topic.doctext_nmf=unpickleModel.NMFDocModel
    Topic.currentEmailVectorizer=unpickleModel.email_vectoriser
    Topic.currentDocVectorizer=unpickleModel.doc_vectoriser
    Topic.cleantext=unpickleModel.cleantext
    Topic.processtext=unpickleModel.processtext

            
    ##Now loop through all the .eml files in the directory specified and transform each of e-mails/docs to see what dominant topic is
    ##if dominant topic in topistodo then we write it out
    Topic.emaildirectory=args.dest_dir
    if not os.path.exists(Topic.emaildirectory):
            #can't construct the class this is an error
            print("Reading emails to see which match topic(s)",args.dest_topics,' .eml directory {} does not exist'.format(Topic.emaildirectory))
            
            raise Exception('.eml directory {} does not exist'.format(Topic.emaildirectory))
    basepart = os.path.basename(os.path.normpath(Topic.emaildirectory))
    
    #check if the directory exists... if so clear it... if not create it
    if Topic.doattachments:
        Topic.attachmentdirectory=Topic.tempdir+basepart
        Topic.dopdf=True
        if os.path.exists(Topic.attachmentdirectory):
            files = glob.glob(Topic.attachmentdirectory+'\\*')
            for f in files:
                os.remove(f)
        else:
            os.mkdir(Topic.attachmentdirectory)

    ## Need to make sure the email text and doc text variables are empty
    Topic.initialiseTextFields()
        
    mswordtype=["application/msword",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "application/rtf"]
        
        
    for subdir, dirs, files in os.walk(Topic.emaildirectory):
           
        for file in files:
               
            filepath = subdir + os.sep + file
            if filepath.endswith(".eml"):
                #print (filepath)
                Text="" # Will concatenate all Text into here
                    

                    #Now need to do error handling here as AV package might stop me getting to some files
                try:
                    fp = open(filepath,'rb')                 
                    msg = BytesParser(policy=policy.default).parse(fp)
                    for part in msg.walk():
        
                        ##Now depending on content type extract what we need to..
                        cp = part.get_content_type()
                        if cp=="text/plain": 
                            Text += part.get_content()
                        elif (cp in mswordtype and Topic.doattachments==True):
                                
                            fn=part.get_filename() 
                            #append the e-mail number to this so we know from which email this attachment is
                            fn=file+'_'+fn
                            ## Now this is important... 
                            ## Remember this is a typosquatting inbox so need to save to disk to make sure your local AV picks up any nasties
                            ## This would not have gone via your normal AV hygiene...
                            filesaved=Topic._save_file(fn, part.get_payload(decode=True))
                            ReadText=Topic._read_text_msword(filesaved) 
                                
                            ##We need to treat eachc attachement seperately
                            if(len(ReadText.strip())):
                                Topic.doctext_original.append(ReadText.strip())
                                Topic.doctext_emailname.append(file)
                        elif (cp=="application/pdf" and Topic.doattachments==True and Topic.dopdf==True) :
                            fn=part.get_filename() 
                            #append the e-mail number to this so we know from which email this attachment is
                            fn=file+'_'+fn
                            ## Now this is important... 
                            ## Remember this is a typosquatting inbox so need to save to disk to make sure your local AV picks up any nasties
                            ## This would not have gone via your normal AV hygiene...
                            filesaved=Topic._save_file(fn, part.get_payload(decode=True))
                            ReadText=Topic._read_text_pdf(filesaved) 
                                
                            ##We need to treat eachc attachement seperately
                            if(len(ReadText.strip())):
                                Topic.doctext_original.append(ReadText.strip())
                                Topic.doctext_emailname.append(file)
                                
                                
                                
                                
                                
                    if(len(Text.strip()) and Topic.doattachments==False):
                        Topic.emailtext_original.append(Text)
                        ##new need to add the email name from whence this comes so we can see interesting e-mails later
                        Topic.emailtext_emailname.append(file)                           

                    fp.close()
                except:
                    next
                    
                    
    ##now clean the text...
    
    Topic.cleanandProcessText()
    
    
    ##With text cleaned and processed we now need to vectorize text using vectorizer and then do a transform
    if(len(Topic.emailtext_forvectorize) and not(Topic.currentEmailVectorizer==None)):
        #now we have already stripped out english stop words earlier so here we can strip out optional other stop words if stoplist was passed
        
        Topic.currentEmailVectorised = Topic.currentEmailVectorizer.transform(Topic.emailtext_forvectorize)
            
    if(len(Topic.doctext_forvectorize) and not(Topic.currentDocVectorizer==None)):
        #now we have already stripped out english stop words earlier so here we can strip out optional other stop words if stoplist was passed
        Topic.currentDocVectorised = Topic.currentDocVectorizer.transform(Topic.doctext_forvectorize)


    ## Now transform each of the bits of currentEmailVectorised & currentDocVectorised
    topic_scores=None

    if Topic.doattachments:
        if not(Topic.doctext_lda==None) and not(Topic.currentDocVectorised==None):
            topic_scores=Topic.doctext_lda.transform(Topic.currentDocVectorised)
            Topic.doctext_topic=np.argmax(topic_scores,axis=1)
            
        elif not(Topic.doctext_nmf==None) and not(Topic.currentDocVectorised==None):
            topic_scores=Topic.doctext_nmf.transform(Topic.currentDocVectorised)
            Topic.doctext_topic=np.argmax(topic_scores,axis=1)
    else:
        if not(Topic.text_lda==None) and not(Topic.currentEmailVectorised==None):
            topic_scores=Topic.text_lda.transform(Topic.currentEmailVectorised)
            Topic.emailtext_topic=np.argmax(topic_scores,axis=1)
        elif not(Topic.text_nmf==None) and not(Topic.currentEmailVectorised==None):
            topic_scores=Topic.text_nmf.tranform(Topic.currentEmailVectorised)
            Topic.emailtext_topic=np.argmax(topic_scores,axis=1)
    
    
    ### now go through each of the items and see which match the topic...
    if Topic.doattachments and len(Topic.doctext_topic):
        if not(Topic.doctext_lda==None):
            model=Topic.doctext_lda
        elif not(Topic.doctext_nmf==None):
            model=Topic.doctext_nmf
        TopicString=Topic.listTopics(model, Topic.currentDocVectorizer)
        for topic in topicstodo:
            
            print("Topic: ",TopicString[topic])
            for i in range(0,len(Topic.doctext_topic)):
                if Topic.doctext_topic[i]==topic:
                    print(" ",Topic.doctext_emailname[i]," ")
            print("\n")    
    elif len(Topic.emailtext_topic):
        if not(Topic.text_lda==None):
            model=Topic.text_lda
        elif not(Topic.text_nmf==None):
            model=Topic.text_nmf
        TopicString=Topic.listTopics(model,Topic.currentEmailVectorizer)
        for topic in topicstodo:
            
            
            print("Topic: ",TopicString[topic])
            for i in range(0,len(Topic.emailtext_topic)):
                if Topic.emailtext_topic[i]==topic:
                    print(" ",Topic.emailtext_emailname[i]," ")
            print("\n")        
                    
                    
            
                    
    return True
    

def handler_scan_local(args):
    Topic=TopicModel(tempdir="c:\\temp\\")
    ##need to decode here either docs or pdfs or both
    if args.dir_to_scan:
        Topic.readdrivedirectory=args.dir_to_scan
    else:
        print("need argument -d <directory to scan>")
        return False
    
    if args.scan_what:
        scanwhat=args.scan_what
        if scanwhat=='W':
            doWord=True
            doPDF=False
            Topic.doattachments=True
        elif scanwhat=='P':
            doPDF=True
            doWord=False
            Topic.doattachments=True
            Topic.dopdf=True
        elif scanwhat=='B':
            doPDF=True
            doWord=True
            Topic.doattachments=True
            Topic.dopdf=True
    else:
        doWord=True
        doPDF=False
        Topic.doattachments=True      
        
    Topic.readLocalComputerDocuments(doPDF,doWord)
    ##see if there is a stopword file
    if args.scan_stop_words:
        ## now load the stopword file
        stopfile=args.scan_stop_words
        f = open(stopfile, "r")
        stopread=f.read()
        Topic.removelist=stopread
        f.close()
        Topic.removewords=True
    else:
        Topic.removewords=False
        Topic.removelist=""
        
    Topic.cleantext=True
    Topic.processtext=True
    
    if args.scan_lemmaargs:
        Topic.lemma_args=args.scan_lemmaargs
        Topic.processtext=True
    else:
        Topic.lemma_args="['NOUN','ADJ','VERB','ADV']"
        Topic.processtext=True
    
    if args.scan_ngram_args:
        Topic.ngram_args=args.scan_ngram_args
    else:
        Topic.ngram_args="1,1"
        
    if args.scan_regex_pattern:
        Topic.regex_pattern=args.scan_regex_pattern
    else: 
        Topic.regex_pattern="[a-zA-Z0-9]{3,}"
        
    
    Topic.text_number_topics=10
    
    if args.scan_number_topics:
        Topic.doctext_number_topics=args.scan_number_topics
    else:
        Topic.doctext_number_topics=10 


    ### Do a lot of heavy lifting to clean the text
    Topic.cleanandProcessText()
    #print ("Have cleaned text")
      
    emailnumtopics=10
    wordsdoctopic=10
        
    mindf=2
    maxdf=20
    numdocfeatures=10000
    numemailfeatures=10000
        

        #pyldavisdoc=not(self.ui.dovisdoc.checkState()==0)
    if args.scan_dovis:
        pyldavisdocfile="c:\\temp\\"+datetime.datetime.now().strftime("%B %d %Y %Hh%M")+" pyldavisdoc.html"  
        pyldavisdoc=True
    else:
        pyldavisdoc=False
        
    vectorizedtxt,vectorizeddoctxt=Topic.countVectorise(min_df=mindf,max_df=maxdf,no_txt_features=numemailfeatures,no_doc_features=numdocfeatures,
                                                        regex_pattern=Topic.regex_pattern,ngramrange=Topic.ngram_args) 
 

    todisplaydoc=""

        
    Topic.emaildoctopwords[:]=[]
    Topic.emaildoctopwordweights[:]=[]


    learning_decay=0.7 #online used in online
    batch_size=128 #only used in online
    perp_tol=0.1 #only used in batch learning
    
    if args.scan_max_iter:
        iterations=args.scan_maxiter
    else:
        iterations=20
        
    evaluate_every=1
    mean_change_tol=0.001
              
    LDA, LDADoc=Topic.doLDA(text_vectorized=vectorizedtxt,doc_text_vectorized=vectorizeddoctxt,
                        text_n_components=emailnumtopics,doc_n_components=Topic.doctext_number_topics,
                        learning_method="batch",max_iter=iterations,learning_decay=learning_decay,
                        batch_size=batch_size,perp_tol=perp_tol,mean_change_tol=mean_change_tol,evaluate_every=evaluate_every)   
              

    if pyldavisdoc:
        try:
            docpanel=pyLDAvis.sklearn.prepare(LDADoc,vectorizeddoctxt,Topic.currentDocVectorizer,mds='tsne')
            pyLDAvis.save_html(docpanel,pyldavisdocfile)
        except:
            print("Could not create visualisation file")

    if not(LDADoc==None):                  
        todisplaydoc=Topic.stringTopics(LDADoc,Topic.currentDocVectorizer,no_top_words=wordsdoctopic)    
        print(todisplaydoc)
        






    
    
    
    
        
    
def main():
    parser=argparse.ArgumentParser()
    subparsers = parser.add_subparsers(help='help for subcommand')

    # create the parser for the "-match" command
    parser_matchtopic = subparsers.add_parser('match', help='display emails that match topic provied in t')
    parser_matchtopic.add_argument('-f', type=str, dest='model_file',help='The model file to load -f <modelfile.pkl>')
    parser_matchtopic.add_argument("-d", type=str, dest='dest_dir',help="Directory to match e-mails from -d directory")
    parser_matchtopic.add_argument("-t", type=str, dest='dest_topics',help='Print e-mail names whose dominant topic match list of topics -t x-y|x,y|x')
    parser_matchtopic.add_argument("-w", type=str, dest='dest_what',help='Match what either topics in e-mail or topics in attachments -w e | a i.e. -w e or -w a')
    parser_matchtopic.set_defaults(func=handler_match_topic)
    
    
    parser_interpret = subparsers.add_parser('gui', help='Run the full topic modelling GUI')
    parser_interpret.set_defaults(func=handler_rungui)
    
    parser_scanlocal = subparsers.add_parser('local',help='Do topic modelling against local files on device')
    parser_scanlocal.add_argument('-d',type=str,dest='dir_to_scan',help='Directory to scan for *.doc(x) & *.pdf files')
    parser_scanlocal.add_argument('-w',type=str,dest='scan_what',help='Scan what -w W|P|B i.e. Word or PDF or Both')
    parser_scanlocal.add_argument('-l',type=str,dest='scan_lemmaargs',help='-l [\'NOUN\',\'ADJ\',\'VERB\',\'ADV\'] Optional Lemma arguments for lemmatisation')
    parser_scanlocal.add_argument('-n',type=str,dest='scan_ngram_args',help='-n x,y optional ngrams from x to y default 1,1')
    parser_scanlocal.add_argument('-i',type=int,dest='scan_max_iter',help='Optional maximum number of iterations for batch mode during LDA eval')
    parser_scanlocal.add_argument('-r',type=str,dest='scan_regex_pattern',help='Optional regular expression pattern for words default is [a-zA-Z0-9]{3,}')
    parser_scanlocal.add_argument('-t',type=int,dest='scan_number_topics',help='Option number of topics default is -t 10')
    parser_scanlocal.add_argument('-s',type=str,dest='scan_stop_words',help='optional -f <stopwordsfile> : text file containing comma delimited list of stopwords')
    parser_scanlocal.add_argument('-v',action='store_true',dest='scan_dovis',help="create visualisation file")
    parser_scanlocal.set_defaults(func=handler_scan_local)
    
    args=parser.parse_args()
    args.func(args)
    
    
    

       
if __name__== "__main__":
    main()