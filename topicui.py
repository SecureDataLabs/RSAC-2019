# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'topicmodelling.ui'
#
# Created by: PyQt5 UI code generator 5.9.2
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets

class Ui_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(1419, 916)
        self.EmbeddingMethod = QtWidgets.QComboBox(Form)
        self.EmbeddingMethod.setGeometry(QtCore.QRect(80, 210, 121, 22))
        self.EmbeddingMethod.setMaxVisibleItems(2)
        self.EmbeddingMethod.setObjectName("EmbeddingMethod")
        self.EmbeddingMethod.addItem("")
        self.EmbeddingMethod.addItem("")
        self.label = QtWidgets.QLabel(Form)
        self.label.setGeometry(QtCore.QRect(20, 210, 53, 16))
        self.label.setObjectName("label")
        self.removewords = QtWidgets.QCheckBox(Form)
        self.removewords.setGeometry(QtCore.QRect(130, 110, 16, 20))
        self.removewords.setText("")
        self.removewords.setChecked(True)
        self.removewords.setObjectName("removewords")
        self.label_2 = QtWidgets.QLabel(Form)
        self.label_2.setGeometry(QtCore.QRect(30, 110, 101, 16))
        self.label_2.setObjectName("label_2")
        self.AlgorithmCombo = QtWidgets.QComboBox(Form)
        self.AlgorithmCombo.setGeometry(QtCore.QRect(80, 240, 51, 22))
        self.AlgorithmCombo.setObjectName("AlgorithmCombo")
        self.AlgorithmCombo.addItem("")
        self.AlgorithmCombo.addItem("")
        self.label_3 = QtWidgets.QLabel(Form)
        self.label_3.setGeometry(QtCore.QRect(20, 240, 61, 16))
        self.label_3.setObjectName("label_3")
        self.emailtopics = QtWidgets.QTextBrowser(Form)
        self.emailtopics.setGeometry(QtCore.QRect(20, 450, 671, 192))
        self.emailtopics.setReadOnly(True)
        self.emailtopics.setObjectName("emailtopics")
        self.doctopics = QtWidgets.QTextBrowser(Form)
        self.doctopics.setGeometry(QtCore.QRect(700, 450, 701, 192))
        self.doctopics.setObjectName("doctopics")
        self.TopicsEmail = QtWidgets.QSpinBox(Form)
        self.TopicsEmail.setGeometry(QtCore.QRect(110, 310, 46, 22))
        self.TopicsEmail.setMinimum(1)
        self.TopicsEmail.setProperty("value", 10)
        self.TopicsEmail.setObjectName("TopicsEmail")
        self.TopicsDoc = QtWidgets.QSpinBox(Form)
        self.TopicsDoc.setGeometry(QtCore.QRect(820, 310, 46, 22))
        self.TopicsDoc.setMinimum(1)
        self.TopicsDoc.setProperty("value", 10)
        self.TopicsDoc.setObjectName("TopicsDoc")
        self.label_4 = QtWidgets.QLabel(Form)
        self.label_4.setGeometry(QtCore.QRect(710, 310, 111, 16))
        self.label_4.setObjectName("label_4")
        self.label_5 = QtWidgets.QLabel(Form)
        self.label_5.setGeometry(QtCore.QRect(20, 310, 91, 16))
        self.label_5.setObjectName("label_5")
        self.label_7 = QtWidgets.QLabel(Form)
        self.label_7.setGeometry(QtCore.QRect(720, 210, 51, 20))
        self.label_7.setObjectName("label_7")
        self.WordsperEmailTopic = QtWidgets.QSpinBox(Form)
        self.WordsperEmailTopic.setGeometry(QtCore.QRect(290, 310, 46, 22))
        self.WordsperEmailTopic.setMinimum(1)
        self.WordsperEmailTopic.setProperty("value", 10)
        self.WordsperEmailTopic.setObjectName("WordsperEmailTopic")
        self.label_9 = QtWidgets.QLabel(Form)
        self.label_9.setGeometry(QtCore.QRect(190, 310, 101, 16))
        self.label_9.setObjectName("label_9")
        self.Doattachments = QtWidgets.QCheckBox(Form)
        self.Doattachments.setGeometry(QtCore.QRect(150, 0, 21, 20))
        self.Doattachments.setText("")
        self.Doattachments.setChecked(True)
        self.Doattachments.setObjectName("Doattachments")
        self.doattachmentlabel = QtWidgets.QLabel(Form)
        self.doattachmentlabel.setGeometry(QtCore.QRect(30, 0, 111, 16))
        self.doattachmentlabel.setObjectName("doattachmentlabel")
        self.dotopicmodel = QtWidgets.QPushButton(Form)
        self.dotopicmodel.setGeometry(QtCore.QRect(20, 390, 171, 28))
        self.dotopicmodel.setObjectName("dotopicmodel")
        self.label_11 = QtWidgets.QLabel(Form)
        self.label_11.setGeometry(QtCore.QRect(20, 430, 151, 16))
        self.label_11.setObjectName("label_11")
        self.label_12 = QtWidgets.QLabel(Form)
        self.label_12.setGeometry(QtCore.QRect(700, 430, 151, 16))
        self.label_12.setObjectName("label_12")
        self.sourcedir = QtWidgets.QLineEdit(Form)
        self.sourcedir.setGeometry(QtCore.QRect(240, 30, 281, 22))
        self.sourcedir.setObjectName("sourcedir")
        self.emailpyldavisfilename = QtWidgets.QLineEdit(Form)
        self.emailpyldavisfilename.setGeometry(QtCore.QRect(200, 350, 491, 22))
        self.emailpyldavisfilename.setObjectName("emailpyldavisfilename")
        self.emailmatchtopicdisplay = QtWidgets.QPushButton(Form)
        self.emailmatchtopicdisplay.setGeometry(QtCore.QRect(20, 650, 191, 28))
        self.emailmatchtopicdisplay.setObjectName("emailmatchtopicdisplay")
        self.emailsmatchingtopicnum = QtWidgets.QSpinBox(Form)
        self.emailsmatchingtopicnum.setGeometry(QtCore.QRect(220, 650, 46, 22))
        self.emailsmatchingtopicnum.setMinimum(0)
        self.emailsmatchingtopicnum.setProperty("value", 0)
        self.emailsmatchingtopicnum.setObjectName("emailsmatchingtopicnum")
        self.docsmatchingtopicnum = QtWidgets.QSpinBox(Form)
        self.docsmatchingtopicnum.setGeometry(QtCore.QRect(990, 650, 46, 22))
        self.docsmatchingtopicnum.setMinimum(0)
        self.docsmatchingtopicnum.setProperty("value", 0)
        self.docsmatchingtopicnum.setObjectName("docsmatchingtopicnum")
        self.docsmatchtopicdisplay = QtWidgets.QPushButton(Form)
        self.docsmatchtopicdisplay.setGeometry(QtCore.QRect(710, 650, 271, 28))
        self.docsmatchtopicdisplay.setObjectName("docsmatchtopicdisplay")
        self.emailsmatchingdoctopic = QtWidgets.QTextBrowser(Form)
        self.emailsmatchingdoctopic.setGeometry(QtCore.QRect(700, 690, 671, 171))
        self.emailsmatchingdoctopic.setObjectName("emailsmatchingdoctopic")
        self.emailsmatchingtexttopic = QtWidgets.QTextBrowser(Form)
        self.emailsmatchingtexttopic.setGeometry(QtCore.QRect(20, 690, 671, 171))
        self.emailsmatchingtexttopic.setObjectName("emailsmatchingtexttopic")
        self.pyldavisdocfilename = QtWidgets.QLineEdit(Form)
        self.pyldavisdocfilename.setGeometry(QtCore.QRect(930, 350, 471, 22))
        self.pyldavisdocfilename.setObjectName("pyldavisdocfilename")
        self.removelist = QtWidgets.QPlainTextEdit(Form)
        self.removelist.setGeometry(QtCore.QRect(150, 110, 781, 61))
        self.removelist.setObjectName("removelist")
        self.processtext = QtWidgets.QCheckBox(Form)
        self.processtext.setGeometry(QtCore.QRect(230, 180, 16, 20))
        self.processtext.setChecked(True)
        self.processtext.setObjectName("processtext")
        self.label_15 = QtWidgets.QLabel(Form)
        self.label_15.setGeometry(QtCore.QRect(150, 180, 71, 16))
        self.label_15.setObjectName("label_15")
        self.label_20 = QtWidgets.QLabel(Form)
        self.label_20.setGeometry(QtCore.QRect(900, 310, 101, 20))
        self.label_20.setObjectName("label_20")
        self.WordsDocTopic = QtWidgets.QSpinBox(Form)
        self.WordsDocTopic.setGeometry(QtCore.QRect(1000, 310, 46, 22))
        self.WordsDocTopic.setMinimum(1)
        self.WordsDocTopic.setProperty("value", 10)
        self.WordsDocTopic.setObjectName("WordsDocTopic")
        self.numemailfeatures = QtWidgets.QLineEdit(Form)
        self.numemailfeatures.setGeometry(QtCore.QRect(1010, 210, 41, 22))
        self.numemailfeatures.setObjectName("numemailfeatures")
        self.label_8 = QtWidgets.QLabel(Form)
        self.label_8.setGeometry(QtCore.QRect(910, 210, 91, 20))
        self.label_8.setObjectName("label_8")
        self.label_21 = QtWidgets.QLabel(Form)
        self.label_21.setGeometry(QtCore.QRect(1060, 210, 131, 20))
        self.label_21.setObjectName("label_21")
        self.numdocfeatures = QtWidgets.QLineEdit(Form)
        self.numdocfeatures.setGeometry(QtCore.QRect(1190, 210, 41, 22))
        self.numdocfeatures.setObjectName("numdocfeatures")
        self.maxdf = QtWidgets.QLineEdit(Form)
        self.maxdf.setGeometry(QtCore.QRect(770, 210, 31, 22))
        self.maxdf.setObjectName("maxdf")
        self.mindf = QtWidgets.QLineEdit(Form)
        self.mindf.setGeometry(QtCore.QRect(860, 210, 31, 22))
        self.mindf.setObjectName("mindf")
        self.label_22 = QtWidgets.QLabel(Form)
        self.label_22.setGeometry(QtCore.QRect(810, 210, 41, 20))
        self.label_22.setObjectName("label_22")
        self.emailalphalabel = QtWidgets.QLabel(Form)
        self.emailalphalabel.setGeometry(QtCore.QRect(20, 270, 81, 16))
        self.emailalphalabel.setObjectName("emailalphalabel")
        self.emaill1ratiolabel = QtWidgets.QLabel(Form)
        self.emaill1ratiolabel.setGeometry(QtCore.QRect(150, 270, 91, 16))
        self.emaill1ratiolabel.setObjectName("emaill1ratiolabel")
        self.emailalpha = QtWidgets.QLineEdit(Form)
        self.emailalpha.setGeometry(QtCore.QRect(80, 270, 31, 22))
        self.emailalpha.setObjectName("emailalpha")
        self.emaill1ratio = QtWidgets.QLineEdit(Form)
        self.emaill1ratio.setGeometry(QtCore.QRect(220, 270, 31, 22))
        self.emaill1ratio.setObjectName("emaill1ratio")
        self.docalphalabel = QtWidgets.QLabel(Form)
        self.docalphalabel.setGeometry(QtCore.QRect(710, 270, 51, 16))
        self.docalphalabel.setObjectName("docalphalabel")
        self.docalpha = QtWidgets.QLineEdit(Form)
        self.docalpha.setGeometry(QtCore.QRect(770, 270, 31, 22))
        self.docalpha.setObjectName("docalpha")
        self.docl1ratio = QtWidgets.QLineEdit(Form)
        self.docl1ratio.setGeometry(QtCore.QRect(890, 270, 31, 22))
        self.docl1ratio.setObjectName("docl1ratio")
        self.docl1ratiolabel = QtWidgets.QLabel(Form)
        self.docl1ratiolabel.setGeometry(QtCore.QRect(820, 270, 61, 16))
        self.docl1ratiolabel.setObjectName("docl1ratiolabel")
        self.maxiterlabel = QtWidgets.QLabel(Form)
        self.maxiterlabel.setGeometry(QtCore.QRect(140, 240, 71, 16))
        self.maxiterlabel.setObjectName("maxiterlabel")
        self.maxiter = QtWidgets.QSpinBox(Form)
        self.maxiter.setGeometry(QtCore.QRect(210, 240, 51, 22))
        self.maxiter.setMinimum(1)
        self.maxiter.setMaximum(200)
        self.maxiter.setProperty("value", 10)
        self.maxiter.setObjectName("maxiter")
        self.ldalearningmethod = QtWidgets.QComboBox(Form)
        self.ldalearningmethod.setGeometry(QtCore.QRect(360, 240, 61, 22))
        self.ldalearningmethod.setObjectName("ldalearningmethod")
        self.ldalearningmethod.addItem("")
        self.ldalearningmethod.addItem("")
        self.ldalearningmethodlabel = QtWidgets.QLabel(Form)
        self.ldalearningmethodlabel.setGeometry(QtCore.QRect(270, 240, 91, 16))
        self.ldalearningmethodlabel.setObjectName("ldalearningmethodlabel")
        self.choosedirectory = QtWidgets.QPushButton(Form)
        self.choosedirectory.setGeometry(QtCore.QRect(30, 30, 201, 28))
        self.choosedirectory.setObjectName("choosedirectory")
        self.setemailvisfile = QtWidgets.QPushButton(Form)
        self.setemailvisfile.setGeometry(QtCore.QRect(20, 350, 171, 28))
        self.setemailvisfile.setObjectName("setemailvisfile")
        self.setdocvisfile = QtWidgets.QPushButton(Form)
        self.setdocvisfile.setGeometry(QtCore.QRect(700, 350, 221, 28))
        self.setdocvisfile.setObjectName("setdocvisfile")
        self.saveprocessed = QtWidgets.QPushButton(Form)
        self.saveprocessed.setGeometry(QtCore.QRect(30, 70, 191, 28))
        self.saveprocessed.setObjectName("saveprocessed")
        self.saveprocessedfilename = QtWidgets.QLineEdit(Form)
        self.saveprocessedfilename.setGeometry(QtCore.QRect(230, 70, 291, 22))
        self.saveprocessedfilename.setObjectName("saveprocessedfilename")
        self.loadprocessedfilename = QtWidgets.QLineEdit(Form)
        self.loadprocessedfilename.setGeometry(QtCore.QRect(760, 70, 321, 22))
        self.loadprocessedfilename.setObjectName("loadprocessedfilename")
        self.loadprocessed = QtWidgets.QPushButton(Form)
        self.loadprocessed.setGeometry(QtCore.QRect(540, 70, 211, 28))
        self.loadprocessed.setObjectName("loadprocessed")
        self.lemmaargs = QtWidgets.QLineEdit(Form)
        self.lemmaargs.setGeometry(QtCore.QRect(250, 180, 991, 22))
        self.lemmaargs.setObjectName("lemmaargs")
        self.label_23 = QtWidgets.QLabel(Form)
        self.label_23.setGeometry(QtCore.QRect(1240, 210, 71, 20))
        self.label_23.setObjectName("label_23")
        self.ngram_args = QtWidgets.QLineEdit(Form)
        self.ngram_args.setGeometry(QtCore.QRect(1310, 210, 61, 22))
        self.ngram_args.setObjectName("ngram_args")
        self.labelgridsearch = QtWidgets.QLabel(Form)
        self.labelgridsearch.setGeometry(QtCore.QRect(930, 240, 61, 16))
        self.labelgridsearch.setObjectName("labelgridsearch")
        self.dogridsearch = QtWidgets.QCheckBox(Form)
        self.dogridsearch.setGeometry(QtCore.QRect(1000, 240, 16, 20))
        self.dogridsearch.setText("")
        self.dogridsearch.setObjectName("dogridsearch")
        self.labelgridsearchargs = QtWidgets.QLabel(Form)
        self.labelgridsearchargs.setGeometry(QtCore.QRect(1020, 240, 81, 20))
        self.labelgridsearchargs.setObjectName("labelgridsearchargs")
        self.gridsearchargs = QtWidgets.QLineEdit(Form)
        self.gridsearchargs.setGeometry(QtCore.QRect(1110, 240, 301, 22))
        self.gridsearchargs.setObjectName("gridsearchargs")
        self.cleantext = QtWidgets.QCheckBox(Form)
        self.cleantext.setGeometry(QtCore.QRect(80, 180, 16, 20))
        self.cleantext.setChecked(True)
        self.cleantext.setObjectName("cleantext")
        self.label_6 = QtWidgets.QLabel(Form)
        self.label_6.setGeometry(QtCore.QRect(20, 180, 61, 16))
        self.label_6.setObjectName("label_6")
        self.regexpatternlabel = QtWidgets.QLabel(Form)
        self.regexpatternlabel.setGeometry(QtCore.QRect(220, 210, 81, 20))
        self.regexpatternlabel.setObjectName("regexpatternlabel")
        self.regexpattern = QtWidgets.QLineEdit(Form)
        self.regexpattern.setGeometry(QtCore.QRect(300, 210, 391, 22))
        self.regexpattern.setObjectName("regexpattern")
        self.savewords = QtWidgets.QPushButton(Form)
        self.savewords.setGeometry(QtCore.QRect(940, 110, 101, 28))
        self.savewords.setObjectName("savewords")
        self.savewordsfilename = QtWidgets.QLineEdit(Form)
        self.savewordsfilename.setGeometry(QtCore.QRect(1050, 110, 321, 22))
        self.savewordsfilename.setObjectName("savewordsfilename")
        self.loadwordsfilename = QtWidgets.QLineEdit(Form)
        self.loadwordsfilename.setGeometry(QtCore.QRect(1050, 140, 321, 22))
        self.loadwordsfilename.setText("")
        self.loadwordsfilename.setObjectName("loadwordsfilename")
        self.loadwords = QtWidgets.QPushButton(Form)
        self.loadwords.setGeometry(QtCore.QRect(940, 140, 101, 28))
        self.loadwords.setObjectName("loadwords")
        self.learningdecaylabel = QtWidgets.QLabel(Form)
        self.learningdecaylabel.setGeometry(QtCore.QRect(430, 240, 81, 16))
        self.learningdecaylabel.setObjectName("learningdecaylabel")
        self.learningdecay = QtWidgets.QDoubleSpinBox(Form)
        self.learningdecay.setGeometry(QtCore.QRect(510, 240, 51, 22))
        self.learningdecay.setDecimals(1)
        self.learningdecay.setMaximum(1.0)
        self.learningdecay.setSingleStep(0.1)
        self.learningdecay.setProperty("value", 0.7)
        self.learningdecay.setObjectName("learningdecay")
        self.showprettypictures = QtWidgets.QCheckBox(Form)
        self.showprettypictures.setGeometry(QtCore.QRect(510, 390, 16, 20))
        self.showprettypictures.setText("")
        self.showprettypictures.setChecked(True)
        self.showprettypictures.setObjectName("showprettypictures")
        self.showmodelperformance = QtWidgets.QCheckBox(Form)
        self.showmodelperformance.setGeometry(QtCore.QRect(350, 390, 16, 20))
        self.showmodelperformance.setText("")
        self.showmodelperformance.setObjectName("showmodelperformance")
        self.label_24 = QtWidgets.QLabel(Form)
        self.label_24.setGeometry(QtCore.QRect(200, 390, 151, 16))
        self.label_24.setObjectName("label_24")
        self.label_25 = QtWidgets.QLabel(Form)
        self.label_25.setGeometry(QtCore.QRect(400, 390, 111, 16))
        self.label_25.setObjectName("label_25")
        self.label_26 = QtWidgets.QLabel(Form)
        self.label_26.setGeometry(QtCore.QRect(560, 390, 151, 16))
        self.label_26.setObjectName("label_26")
        self.showdistributionbytopic = QtWidgets.QCheckBox(Form)
        self.showdistributionbytopic.setGeometry(QtCore.QRect(710, 390, 16, 20))
        self.showdistributionbytopic.setText("")
        self.showdistributionbytopic.setChecked(True)
        self.showdistributionbytopic.setObjectName("showdistributionbytopic")
        self.label_27 = QtWidgets.QLabel(Form)
        self.label_27.setGeometry(QtCore.QRect(790, 390, 71, 16))
        self.label_27.setObjectName("label_27")
        self.showwordcloud = QtWidgets.QCheckBox(Form)
        self.showwordcloud.setGeometry(QtCore.QRect(860, 390, 16, 20))
        self.showwordcloud.setText("")
        self.showwordcloud.setChecked(True)
        self.showwordcloud.setObjectName("showwordcloud")
        self.showwordcountimportance = QtWidgets.QCheckBox(Form)
        self.showwordcountimportance.setGeometry(QtCore.QRect(1150, 390, 16, 20))
        self.showwordcountimportance.setText("")
        self.showwordcountimportance.setChecked(True)
        self.showwordcountimportance.setObjectName("showwordcountimportance")
        self.label_28 = QtWidgets.QLabel(Form)
        self.label_28.setGeometry(QtCore.QRect(890, 390, 251, 16))
        self.label_28.setObjectName("label_28")
        self.readEmailsAttachmentsExchange = QtWidgets.QPushButton(Form)
        self.readEmailsAttachmentsExchange.setGeometry(QtCore.QRect(950, 30, 141, 23))
        self.readEmailsAttachmentsExchange.setObjectName("readEmailsAttachmentsExchange")
        self.label_13 = QtWidgets.QLabel(Form)
        self.label_13.setGeometry(QtCore.QRect(1100, 30, 61, 16))
        self.label_13.setObjectName("label_13")
        self.label_14 = QtWidgets.QLabel(Form)
        self.label_14.setGeometry(QtCore.QRect(1250, 30, 61, 20))
        self.label_14.setObjectName("label_14")
        self.numberemailsload = QtWidgets.QSpinBox(Form)
        self.numberemailsload.setGeometry(QtCore.QRect(1310, 30, 61, 22))
        self.numberemailsload.setMinimum(1)
        self.numberemailsload.setMaximum(999999)
        self.numberemailsload.setProperty("value", 10000)
        self.numberemailsload.setObjectName("numberemailsload")
        self.exchangefoldernumbers = QtWidgets.QLineEdit(Form)
        self.exchangefoldernumbers.setGeometry(QtCore.QRect(1160, 30, 71, 22))
        self.exchangefoldernumbers.setObjectName("exchangefoldernumbers")
        self.readdrivesourcedir = QtWidgets.QLineEdit(Form)
        self.readdrivesourcedir.setGeometry(QtCore.QRect(670, 30, 271, 22))
        self.readdrivesourcedir.setObjectName("readdrivesourcedir")
        self.readfromdrive = QtWidgets.QPushButton(Form)
        self.readfromdrive.setGeometry(QtCore.QRect(530, 30, 131, 28))
        self.readfromdrive.setObjectName("readfromdrive")
        self.dopdflabel = QtWidgets.QLabel(Form)
        self.dopdflabel.setGeometry(QtCore.QRect(190, 0, 71, 16))
        self.dopdflabel.setObjectName("dopdflabel")
        self.dopdf = QtWidgets.QCheckBox(Form)
        self.dopdf.setGeometry(QtCore.QRect(260, 0, 21, 20))
        self.dopdf.setText("")
        self.dopdf.setChecked(True)
        self.dopdf.setObjectName("dopdf")
        self.savetopicmodel = QtWidgets.QPushButton(Form)
        self.savetopicmodel.setGeometry(QtCore.QRect(20, 870, 171, 28))
        self.savetopicmodel.setObjectName("savetopicmodel")
        self.topicmodelsavefilename = QtWidgets.QLineEdit(Form)
        self.topicmodelsavefilename.setGeometry(QtCore.QRect(200, 870, 491, 22))
        self.topicmodelsavefilename.setObjectName("topicmodelsavefilename")
        self.batchsizelabel = QtWidgets.QLabel(Form)
        self.batchsizelabel.setGeometry(QtCore.QRect(570, 240, 51, 16))
        self.batchsizelabel.setObjectName("batchsizelabel")
        self.batchsize = QtWidgets.QLineEdit(Form)
        self.batchsize.setGeometry(QtCore.QRect(620, 240, 31, 22))
        self.batchsize.setObjectName("batchsize")
        self.perptollabel = QtWidgets.QLabel(Form)
        self.perptollabel.setGeometry(QtCore.QRect(430, 240, 81, 16))
        self.perptollabel.setObjectName("perptollabel")
        self.perptol = QtWidgets.QDoubleSpinBox(Form)
        self.perptol.setGeometry(QtCore.QRect(510, 240, 41, 22))
        self.perptol.setDecimals(2)
        self.perptol.setMaximum(1.0)
        self.perptol.setSingleStep(0.05)
        self.perptol.setProperty("value", 0.1)
        self.perptol.setObjectName("perptol")
        self.meanchangetollabel = QtWidgets.QLabel(Form)
        self.meanchangetollabel.setGeometry(QtCore.QRect(660, 240, 91, 16))
        self.meanchangetollabel.setObjectName("meanchangetollabel")
        self.meanchangetol = QtWidgets.QDoubleSpinBox(Form)
        self.meanchangetol.setGeometry(QtCore.QRect(740, 240, 61, 22))
        self.meanchangetol.setDecimals(4)
        self.meanchangetol.setMaximum(1.0)
        self.meanchangetol.setSingleStep(0.0005)
        self.meanchangetol.setProperty("value", 0.001)
        self.meanchangetol.setObjectName("meanchangetol")
        self.evaluateeverylabel = QtWidgets.QLabel(Form)
        self.evaluateeverylabel.setGeometry(QtCore.QRect(810, 240, 61, 20))
        self.evaluateeverylabel.setObjectName("evaluateeverylabel")
        self.evaluateevery = QtWidgets.QSpinBox(Form)
        self.evaluateevery.setGeometry(QtCore.QRect(870, 240, 41, 22))
        self.evaluateevery.setMinimum(0)
        self.evaluateevery.setProperty("value", 0)
        self.evaluateevery.setObjectName("evaluateevery")

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Form"))
        self.EmbeddingMethod.setItemText(0, _translate("Form", "Count Vectorizer"))
        self.EmbeddingMethod.setItemText(1, _translate("Form", "Tfidf Vectorizer"))
        self.label.setText(_translate("Form", "Vectorizer"))
        self.removewords.setToolTip(_translate("Form", "<html><head/><body><p>By ticking this box all the words to the right will be removed from the corpus</p></body></html>"))
        self.label_2.setText(_translate("Form", "Remove words ?"))
        self.AlgorithmCombo.setItemText(0, _translate("Form", "LDA"))
        self.AlgorithmCombo.setItemText(1, _translate("Form", "NMF"))
        self.label_3.setText(_translate("Form", "Algorithm"))
        self.label_4.setText(_translate("Form", "Document # topics"))
        self.label_5.setText(_translate("Form", "E-mail # topics"))
        self.label_7.setText(_translate("Form", "max_df"))
        self.label_9.setText(_translate("Form", "Words  per topic"))
        self.Doattachments.setToolTip(_translate("Form", "<html><head/><body><p>If this is ticked attachments will be processed when reading e-mails from .eml directory or from Exchange. Does not apply when docs are read from local computer</p></body></html>"))
        self.doattachmentlabel.setText(_translate("Form", "Do Attachments ?"))
        self.dotopicmodel.setText(_translate("Form", "Perform Topic Modelling"))
        self.label_11.setText(_translate("Form", "E-mail topics"))
        self.label_12.setText(_translate("Form", "Document topics"))
        self.emailmatchtopicdisplay.setText(_translate("Form", "Show e-mails matching topic"))
        self.docsmatchtopicdisplay.setText(_translate("Form", "Show e-mails/docs with docs matching topic"))
        self.removelist.setPlainText(_translate("Form", "\'http\',\'https\',\'outlook\',\'click\',\'email\',\'iphone\',\'aspx\',\'utm_source\',\'utm_content\',\'utm_medium\',\'utm_campaign\',\'utm_term\',\'unsubscribe\'"))
        self.processtext.setToolTip(_translate("Form", "<html><head/><body><p>By clicking  this box will lemmatize corpus and only return items to the right. See spacy nlp module for documentation on allowed_posttags</p></body></html>"))
        self.processtext.setText(_translate("Form", "regex cleaning"))
        self.label_15.setText(_translate("Form", "Process text "))
        self.label_20.setText(_translate("Form", "Words per Topic"))
        self.numemailfeatures.setText(_translate("Form", "10000"))
        self.label_8.setText(_translate("Form", "No text features"))
        self.label_21.setText(_translate("Form", "No document features"))
        self.numdocfeatures.setText(_translate("Form", "10000"))
        self.maxdf.setText(_translate("Form", "10"))
        self.mindf.setText(_translate("Form", "2"))
        self.label_22.setText(_translate("Form", "min_df"))
        self.emailalphalabel.setText(_translate("Form", "E-mail alpha"))
        self.emaill1ratiolabel.setText(_translate("Form", "E-mail L1 ratio"))
        self.emailalpha.setText(_translate("Form", "0.1"))
        self.emaill1ratio.setText(_translate("Form", "0.5"))
        self.docalphalabel.setText(_translate("Form", "Doc alpha"))
        self.docalpha.setText(_translate("Form", "0.1"))
        self.docl1ratio.setText(_translate("Form", "0.5"))
        self.docl1ratiolabel.setText(_translate("Form", "Doc  L1 ratio"))
        self.maxiterlabel.setText(_translate("Form", "Max iterations"))
        self.ldalearningmethod.setItemText(0, _translate("Form", "batch"))
        self.ldalearningmethod.setItemText(1, _translate("Form", "online"))
        self.ldalearningmethodlabel.setText(_translate("Form", "Learning Method"))
        self.choosedirectory.setText(_translate("Form", "Read .eml eMails/Docs from  Dir"))
        self.setemailvisfile.setText(_translate("Form", "Visualisation  File for eMail"))
        self.setdocvisfile.setText(_translate("Form", "Visualisation File for Attachments"))
        self.saveprocessed.setToolTip(_translate("Form", "<html><head/><body><p>Save emails/attachments that have been read from Directory or from Exchange</p></body></html>"))
        self.saveprocessed.setText(_translate("Form", "Save  read emails/docs"))
        self.loadprocessed.setText(_translate("Form", "Load previously saved emails/docs"))
        self.lemmaargs.setText(_translate("Form", "[\'NOUN\', \'ADJ\', \'VERB\', \'ADV\']"))
        self.label_23.setText(_translate("Form", "ngram_args"))
        self.ngram_args.setText(_translate("Form", "1,1"))
        self.labelgridsearch.setText(_translate("Form", "Gridsearch ?"))
        self.dogridsearch.setToolTip(_translate("Form", "<html><head/><body><p>When clicked will perform gridsearch i..e will see which of the parameters specified in gridsearch args create the \'best\' model according to log likelihood measure. Note gridsearch is only for the LDA algorithm bit. The vectorizer will vectoise according to the parameters on the form. Note  will do search both for email and attachment text so might take a while... Will display optimum values for both.</p></body></html>"))
        self.labelgridsearchargs.setText(_translate("Form", "Gridsearch args"))
        self.gridsearchargs.setToolTip(_translate("Form", "<html><head/><body><p>Gridsearch arguments. See sklearnn <span style=\" font-family:\'monaco,Source Code Pro,monospace,Courier New\'; font-size:13px; color:#660066; background-color:#fafafa;\">LatentDirichletAllocation documentation for parameters</span></p></body></html>"))
        self.gridsearchargs.setText(_translate("Form", "{\'n_components\': [10, 15, 20, 25, 30]} "))
        self.cleantext.setToolTip(_translate("Form", "<html><head/><body><p>by clicking this box will remove single quotes, email addresses and new line characters from corpus</p></body></html>"))
        self.cleantext.setText(_translate("Form", "regex cleaning"))
        self.label_6.setText(_translate("Form", "Clean text"))
        self.regexpatternlabel.setText(_translate("Form", "regex pattern"))
        self.regexpattern.setToolTip(_translate("Form", "<html><head/><body><p>This regular expression pattern will be applied during the vectorizer for parameter token_pattern</p></body></html>"))
        self.regexpattern.setText(_translate("Form", "[a-zA-Z0-9]{3,}"))
        self.savewords.setText(_translate("Form", "save words"))
        self.loadwords.setText(_translate("Form", "load words"))
        self.learningdecaylabel.setText(_translate("Form", "Learning decay"))
        self.showprettypictures.setToolTip(_translate("Form", "<html><head/><body><p><span style=\" font-family:\'geomanistregular,system-ui,Open Sans,Oxygen,sans-serif\'; font-size:10pt; font-weight:296; color:#333333; background-color:#ffffff;\">Use k-means clustering on the document-topic probabilioty matrix. Number of clusters is equal to the number of topics. For the X and Y,  use SVD on the </span><span style=\" font-family:\'monaco,Source Code Pro,monospace,Courier New\'; font-size:10pt; font-weight:296; color:#000000; background-color:#fafafa;\">lda_output</span><span style=\" font-family:\'geomanistregular,system-ui,Open Sans,Oxygen,sans-serif\'; font-size:10pt; font-weight:296; color:#333333; background-color:#ffffff;\"> object with </span><span style=\" font-family:\'monaco,Source Code Pro,monospace,Courier New\'; font-size:10pt; font-weight:296; color:#000000; background-color:#fafafa;\">n_components</span><span style=\" font-family:\'geomanistregular,system-ui,Open Sans,Oxygen,sans-serif\'; font-size:10pt; font-weight:296; color:#333333; background-color:#ffffff;\"> as 2. SVD ensures that these two columns captures the maximum possible amount of information from </span><span style=\" font-family:\'monaco,Source Code Pro,monospace,Courier New\'; font-size:10pt; font-weight:296; color:#000000; background-color:#fafafa;\">lda_output</span><span style=\" font-family:\'geomanistregular,system-ui,Open Sans,Oxygen,sans-serif\'; font-size:10pt; font-weight:296; color:#333333; background-color:#ffffff;\"> in the first 2 components. ( lda_output is the transformation of each document using the model we built )</span></p></body></html>"))
        self.showmodelperformance.setToolTip(_translate("Form", "<html><head/><body><p><span style=\" font-family:\'geomanistregular,system-ui,Open Sans,Oxygen,sans-serif\'; font-size:15px; font-weight:296; color:#333333; background-color:#ffffff;\">A model with higher log-likelihood and lower perplexity (exp(-1. * log-likelihood per word)) is considered to be good.</span></p></body></html>"))
        self.label_24.setText(_translate("Form", "Show model performance"))
        self.label_25.setText(_translate("Form", "Cluster documents"))
        self.label_26.setText(_translate("Form", "Show distribution by topic"))
        self.showdistributionbytopic.setToolTip(_translate("Form", "<html><head/><body><p>Show frequency distribution of emails and documents by topic</p></body></html>"))
        self.label_27.setText(_translate("Form", "Word cloud"))
        self.showwordcloud.setToolTip(_translate("Form", "<html><head/><body><p><span style=\" font-family:\'geomanistregular,system-ui,Open Sans,Oxygen,sans-serif\'; font-size:15px; font-weight:296; color:#333333; background-color:#ffffff;\">We do have topic keywords in each topic the  word cloud with the size of the words proportional to the weight of the words in relation to the topic</span></p></body></html>"))
        self.showwordcountimportance.setToolTip(_translate("Form", "<html><head/><body><p><span style=\" font-family:\'geomanistregular,system-ui,Open Sans,Oxygen,sans-serif\'; font-size:15px; font-weight:296; color:#333333; background-color:#ffffff;\">You want to keep an eye out on the words that occur in multiple topics and the ones whose relative frequency is more than the weight. Often such words turn out to be less important so can be added to stop word list in the beginning and re-running the training process.</span></p></body></html>"))
        self.label_28.setText(_translate("Form", "Word count & importance of Topic keywords"))
        self.readEmailsAttachmentsExchange.setToolTip(_translate("Form", "<html><head/><body><p>Use Outlook on client to connect to exchange using current authenticated user. Read folder(s) in dialog box. Format is 5,6 or just single number. 3: Deleted Items, 5: Sent Items, 6: Inbox</p><p><br/></p><p><br/></p></body></html>"))
        self.readEmailsAttachmentsExchange.setText(_translate("Form", "Read from Exchange"))
        self.label_13.setText(_translate("Form", "Folder(s)"))
        self.label_14.setText(_translate("Form", "# to read"))
        self.exchangefoldernumbers.setToolTip(_translate("Form", "<html><head/><body><p>Format is either x or x,y,z for instance 6 for Inbox or 3,5,6 for Deleted Items, Sent Items and Inbox. See win32.com python documentaiton for complete list of folder numbers</p><p><br/></p></body></html>"))
        self.exchangefoldernumbers.setText(_translate("Form", "6"))
        self.readfromdrive.setText(_translate("Form", "Read Docs from  Dir"))
        self.dopdflabel.setText(_translate("Form", "Do PDFs ?"))
        self.dopdf.setToolTip(_translate("Form", "<html><head/><body><p>Process .pdf files in attachments or when files read from local machine</p></body></html>"))
        self.savetopicmodel.setText(_translate("Form", "Save Topic Model"))
        self.batchsizelabel.setText(_translate("Form", "Batch size"))
        self.batchsize.setText(_translate("Form", "128"))
        self.perptollabel.setText(_translate("Form", "Perp tol"))
        self.meanchangetollabel.setText(_translate("Form", "Mean change tol"))
        self.evaluateeverylabel.setText(_translate("Form", "Eval every"))
        self.evaluateevery.setToolTip(_translate("Form", "<html><head/><body><p>How often to evaluate perplexity. Set to 0 or negative number not to evaluate perplexity at all. Evaluating perplexity can help you  check convergence in  training process but it will also increase total training time. Evaluating perplexity in every iteration might increase raining tim up to two-fold</p></body></html>"))


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Form = QtWidgets.QWidget()
    ui = Ui_Form()
    ui.setupUi(Form)
    Form.show()
    sys.exit(app.exec_())

