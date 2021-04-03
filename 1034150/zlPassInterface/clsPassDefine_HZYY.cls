VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPassDefine_HZYY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--------------------------------------------------------------------------------------------------------------------------------------
' --�ӿڶ��� ������������Ϣ�������޹�˾��
'--------------------------------------------------------------------------------------------------------------------------------------
'<-begin------------------------------------------------------------------------basexml----------------------------------------------------->

Public Type HZYYBASE
    strHospCode         As String
    strPatiID           As String
    strEventNO          As String
    strSource           As String
End Type
'<-begin------------------------------------------------------------------------1.2 optpatient���ﻼ�߾����ǩ��ʽ����:----------------------------------------------------->
Public Type OPTPATIENT
    strSex              As String
    strName             As String
    strIDType           As String
    strIDNO             As String
    strBirthWeight      As String           '��������
    strBirthDay         As String           '��������
    strEthnicGroup      As String           '����
    strNativePlace      As String           '����
    strRace             As String           '����
    strMedCardNO        As String           '���￨��
    strEventTime       As String           '����ʱ��
    strDeptID           As String           '�Һſ���ID
    strDeptName         As String           '�Һſ�������
    strPayType          As String           '��������
    strPregnancy        As String           '�Ƿ���
    strTimeOfPreg       As String           '����
    strBreastFeeding    As String          '�Ƿ���
    strHeight           As String                 '����
    strWeight           As String                 '����
    strAddress          As String                '���˵�ַ
    strPhoneNo          As String                '���˵绰
    strDialysis         As String               '�Ƿ�͸��
    strmarital          As String                 '����״��
    strOccupation       As String             'ְҵ
    strSpecialConstitution As String    '��������
    strVisitType        As String              '�������
    strPatiCondition    As String          '����״̬
End Type

'<-begin------------------------------------------------------------------------1.3 iptpatientסԺ���߾����ǩ:----------------------------------------------------->
Public Type IPTPATIENT
    strSex                      As String
    strName                     As String
    strIDType                   As String
    strIDNO                     As String
    strBirthWeight              As String           '��������
    strBirthDay                 As String           '��������
    strEthnicGroup              As String           '����
    strNativePlace              As String           '����
    strRace                     As String           '����
    strMedCardNO                As String           '���￨��
    strPayType                  As String           '��������
    strPregnancy                As String           '�Ƿ���
    strTimeOfPreg               As String           '����
    strBreastFeeding            As String           '�Ƿ���
    strHeight                   As String           '����
    strWeight                   As String           '����
    strAddress                  As String           '���˵�ַ
    strPhoneNo                  As String           '���˵绰
    strDialysis                 As String           '�Ƿ�͸��
    strmarital                  As String           '����״��
    strOccupation               As String           'ְҵ
    strSpecialConstitution      As String           '��������
    strINDeptId                 As String           '��Ժ����ID
    strINDeptName               As String           '��Ժ��������
    strHospitalTime             As String           '��Ժʱ��
    strInWardID                 As String           '��Ժ����ID
    strInWardName               As String           '��Ժ��������
    strInWardBedNo              As String           '��Ժ������
    strInConditon               As String           '��Ժ����
    strWeightOfBaby             As String           '��������Ժ����
    strPatientConditon          As String           '����״̬
End Type
'<-begin------------------------------------------------------------------------1.4 optprescriptions�����ʹ�����ϸ��Ϣ��ǩ:----------------------------------------------------->

Public Type OPTPRESCRIPTIONSINFO
    strRecipeId             As String                   '����id
    strRecipeNo             As String                   '������
    strRecipeSource         As String                   '������Դ
    strRecipeCategory       As String                   '�������
    strRecipeType           As String                   '��������
    strDeptID               As String                   '��������ID
    strDeptName             As String                   '������������
    strRecipeDocTitle       As String                   '����ҽ��ְ��
    strRecipeDocId          As String                   '����ҽ������
    strRecipeDocName        As String                   '����ҽ������
    strRecipeTime           As String                   '����ʱ��
    strHerbUnitPrice        As String                   '��Ƭ�����۸�
    strHerbPacketCount     As String                    '��Ƭ����
    strIsCream              As String                   '�෽
    strRecipeFeeTotal       As String                   '�������
    strOriginalRecipeId     As String                   'ԭʼ������
    strRecipeStatus         As String                   '����״̬
    strUrgentFlag           As String                   '������־
    strReviewPharmId        As String                   '�˶�ҩʦ����
    strReviewPharmName      As String                   '�˶�ҩʦ����
    strReviewPharmNTitle    As String                   '�˶�ҩʦְ��
    strPrepPharmId          As String                   '����ҩʦ����
    strPrepPharmName        As String                   '����ҩʦ����
    strPrepPharmTitle       As String                   '����ҩʦְ��
    strCheckPharmID         As String                   '���ҩʦ����
    strCheckPharmName       As String                   '���ҩʦ����
    strCheckPharmTitle      As String                   '���ҩʦְ��
    strDespensingPharmId    As String                   '��ҩҩʦ����
    strDespensingPharmName  As String                   '��ҩҩʦ����
    strDespensingPharmTitle As String                   '��ҩҩʦ����
End Type

Public Type OptPRESCRIPTIONSITEM
    strRecipeItemId             As String               '������ϸ���
    strRecipeId                 As String               '����id
    strDrugPurpose              As String               'ҩ��Ŀ��
    strGroupNO                  As String               '���
    strDrugID                   As String               'ҩƷID
    strDrugName                 As String               'ҩƷ����
    strCountUnit                As String               '��װ�������
    strPackUnit                 As String               '��װ���λ
    strManufacturerID           As String               '��������id
    strManufacturerName         As String               '������������
    strDrugdose                 As String               'ÿ�θ�ҩ����
    strDrugadminRouteName       As String               '��ҩ;��
    strDrugUsingFreq            As String               '��ҩƵ��
    strDrugUsingTimePoint       As String               '��ҩʱ��
    strDrugUsingAim             As String               '��ҩĿ��
    strDrugUsingArea            As String               '��ҩ��λ
    strDuration                 As String               '�Ƴ�
    strPreparation              As String               '����
    strSpecification            As String               '���
    strUnitPrice                As String               '����
    strDespensingNum            As String               '��ҩ����
    strFeeTotal                 As String               '�ܼ�
    strStartTime                As String               '��ʼʹ��ʱ��
    strEndTime                  As String               'ֹͣʹ��ʱ��
    strSpecialPrompt            As String               '����Ҫ��
    strSkinTestFlag             As String               'Ƥ�Ա�־
    strSkinTestResult           As String               'Ƥ�Խ��
    strSkinTestTime             As String               'Ƥ��ʱ��
    strDrugSource               As String               'ҩƷ��Դ
    strdrugReturnFlag           As String               '�Ƿ���ҩ��־
    strOuvasFlag                As String               '���ﾲ���־
    strDrippingSpeed            As String               '����
    strLimitTime                As String               '����ʱ��
    strTherapeuticRegimen       As String               '��ҩ����
    strDispensingWindow         As String               '��ҩ���ں�
    strDrugstoreArea            As String               '��Ʒ���ܺ�
End Type

'<-begin------------------------------------------------------------------------1.5 ordersҽ����Ϣ��ǩ:----------------------------------------------------->
Public Type NonMedicalOrderItem          '���һ����ҩƷҽ��
    strOrderId              As String       'ҽ��id
    strOrderTime            As String       'ҽ��ʱ��
    strOrderDeptID          As String       'ҽ������id
    strOrderDeptName        As String       'ҽ����������
    strDocGroup             As String       'ҽ��ҽ��������
    strOrderDocName         As String       'ҽ��ҽ������
    strOrderDocID           As String       'ҽ��ҽ������
    strOrderDocTitle        As String       'ҽ��ҽ��ְ��
    strOrderType            As String       'ҽ������
    strOrderCode            As String       'ҽ������
    strOrderName            As String       'ҽ������
    strOrderCategory        As String       'ҽ�����
    strOrderFreq            As String       'ҽ��Ƶ��
    strOrderValidTime       As String       'ҽ����Чʱ��
    strOrderInvalidTime     As String       'ҽ��ʧЧʱ��
    strDuration             As String       '�Ƴ�
    strCheckTime            As String       '����ʱ��
    strCheckNurseID         As String       '���˻�ʿid
    strCheckNurseName       As String       '���˻�ʿ����
    strStopFlag             As String       'ҽ��ֹͣ��־
End Type

Public Type MedicalOrderItem              '��ǩһ��ҩ��ҽ��������ҩ��ҽ�����
    strOrderId                     As String         'ҽ��id
    strOrderTime                   As String         'ҽ��ʱ��
    strOrderDeptID                As String         'ҽ�����Ҵ���
    strOrderDeptName             As String         'ҽ����������
    strDocGroup                    As String         'ҽ��ҽ��������
    strOrderDocID                 As String         'ҽ��ҽ��id
    strOrderDocName              As String         'ҽ��ҽ������
    strOrderDocTitle              As String         'ҽ��ҽ��ְ��
    strOrderType                   As String         'ҽ������
    strDrugPurpose                 As String         'ҩ��Ŀ��
    strGroupNO                     As String         '���
    strDrugID                      As String         'ҩƷID
    strDrugName                   As String         'ҩƷͨ����
    strCountUnit                   As String         '��װ�������
    strPackUnit                    As String         '��װ���λ
    strManufacturerID              As String         '��������id
    strManufacturerName           As String         '������������
    strDrugdose                    As String         'ÿ�θ�ҩ����
    strDrugadminRouteName       As String         '��ҩ;��
    strDrugUsingFreq              As String         '��ҩƵ��
    strDrugUsingTimePoint        As String         '��ҩʱ��
    strDrugUsingAim               As String         '��ҩĿ��
    strDrugUsingArea              As String         '��ҩ��λ
    strDrugSource                  As String         'ҩƷ��Դ
    strDuration                     As String         '�Ƴ�
    strPreparation                  As String         '����
    strSpecifications               As String         '���
    strUnitPrice                   As String         '����
    strDespensingNum               As String         '��ҩ����
    strFeeTotal                    As String         '�ܼ�
    strCheckTime                   As String         '����ʱ��
    strCheckNurseID               As String         '���˻�ʿid
    strCheckNurseName            As String         '���˻�ʿ����
    strOrderValidTime             As String         'ҽ����Чʱ��
    strOrderInvalidTime           As String         'ҽ��ʧЧʱ��
    strSpecialPrompt               As String         '����Ҫ��
    strSkinTestTime               As String         'Ƥ��ʱ��
    strSkinTestFlag               As String         'Ƥ�Ա�־
    strSkinTestResult             As String         'Ƥ�Խ��
    strdrugReturnFlag             As String         '�Ƿ���ҩ��־
    strStopFlag                    As String         '�Ƿ�ͣҩ��־
    strPivasFlag                   As String         'סԺ�����־
    strUrgentFlag                  As String         '������־
    strDrippingSpeed               As String         '����
    strLimitTime                   As String         '����ʱ��
    strTherapeuticRegimen          As String         '��ҩ����
    strExeDeptID                  As String         'ҽ��ִ�п���id
    strExeDeptName               As String         'ҽ��ִ�п�������
    strDispensingWindow            As String         '��ҩ���ں�
    strDrugstoreArea              As String         '��Ʒ���ܺ�
End Type



Public Type HerbMedicalOrderInfo
    strOrderId                      As String                  'ҽ��id
    strOrderTime                    As String                  'ҽ��ʱ��
    strOrderDeptID                  As String                  'ҽ�����Ҵ���
    strOrderDeptName                As String                  'ҽ����������
    strDocGroup                     As String                  'ҽ��ҽ��������
    strOrderDocID                   As String                  'ҽ��ҽ��id
    strOrderDocName                 As String                  'ҽ��ҽ������
    strOrderDocTitle                As String                  'ҽ��ҽ��ְ��
    strOrderType                    As String                  'ҽ������
    strHerbUnitPrice                As String                  '��Ƭ�����۸�
    strHerbPacketCount              As String                  '��Ƭ����
    strIsCream                      As String                  '�෽
    strCheckTime                    As String                  '����ʱ��
    strCheckNurseID                 As String                  '���˻�ʿid
    strCheckNurseName               As String                  '���˻�ʿ����
    strOrderValidTime               As String                  'ҽ����Чʱ��
    strOrderInvalidTime             As String                  'ҽ��ʧЧʱ��
    strdrugReturnFlag               As String                  '�Ƿ���ҩ��־
    strStopFlag                     As String                  '�Ƿ�ͣҩ��־
    strUrgentFlag                   As String                  '������־
    strExeDeptID                    As String                  'ҽ��ִ�п���id
    strExeDeptName                  As String                  'ҽ��ִ�п�������
End Type

Public Type HerbMedicalOrderItem                     '��ǩһ����Ƭҽ����ϸ�� ������Ƭҽ����ϸ���
    strOrderId                     As String                                                    'ҽ��id
    strOrderitemID                As String                                                    'ҽ����ϸ
    strGroupNO                     As String                                                    '���
    strDrugID                      As String                                                    'ҩƷID
    strDrugName                   As String                                                    'ҩƷͨ����
    strManufacturerID              As String                                                    '��������id
    strManufacturerName           As String                                                    '������������
    strDrugdose                    As String                                                    'ÿ�θ�ҩ����
    strDrugadminRouteName       As String                                                    '��ҩ;��
    strDrugUsingFreq              As String                                                    '��ҩƵ��
    strPreparation                  As String                                                    '����
    strSpecifications               As String                                                    '���
    strUnitPrice                   As String                                                    '����
    strDespensingNum               As String                                                    '��ҩ����
    strFeeTotal                    As String                                                    '�ܼ�
    strSpecialPrompt               As String                                                    '����Ҫ��
End Type
'<-begin------------------------------------------------------------------------1.6 Diagnoses�����Ϣ��ǩ:----------------------------------------------------->

Public Type Diagnosis               'һ�������Ϣ�����������Ϣ���
    strDiagID                      As String                 '�����ˮ��
    strDiagDeptID                 As String                 '��Ͽ���id
    strDiagDeptName              As String                 '��Ͽ�������
    strDiagDocID                  As String                 '���ҽ������
    strDiagDocName               As String                 '���ҽ������
    strDiagDocTitle               As String                 '���ҽ��ְ��
    strDiagDate                    As String                 '���ʱ��
    strDiagCategory                As String                 '������
    strDiagType                    As String                 '�������
    strDiagName                   As String                 '�������
    strDiagCode                    As String                 '��ϱ���
    strDiagCodeType               As String                 '��ϱ�������
    strDiseaseClassification       As String                 '��������
    strDiseaseStaging              As String                 '��������
    strDiseaseScore                As String                 '��������
End Type

'<-begin------------------------------------------------------------------------1.7 allergies������Ϣ��ǩ:----------------------------------------------------->
Public Type Allergy
    strAllergyID           As String          '������Ϣ
    strAllergyDrug         As String          '����ҩ��
    strAnaphylaxis          As String          '������Ӧ
    strRecordTime          As String          '��¼ʱ��
End Type

'<-begin------------------------------------------------------------------------1.8 Operations������Ϣ��ǩ:----------------------------------------------------->

Public Type Operation
    strOperationID                 As String                 '����id
    strOperationCode               As String                 '��������
    strOperationName              As String                 '��������
    strDeptID                      As String                 '����id
    strDeptName                   As String                 '��������
    strOperationStartTime         As String                 '������ʼʱ��
    strOperationEndTime           As String                 '��������ʱ��
    strOperationIncisionType      As String                 '�п�����
    strAnesthesiaCode              As String                 '��������
    strAsa                          As String                 '�����ּ�
    strAnesthesiaEndTime          As String                 '����ֹͣʱ��
    strAnesthesiaStartTime        As String                 '������ʼʱ��
    strIsImplant                   As String                 '�Ƿ���ֲ����
    strImplantNO                   As String                 'ֲ����id
    strImplantName                As String                 'ֲ��������
    strIsReOperation               As String                 '�Ƿ��������
    strOperationDocID             As String                 '����ҽ������
    strOperationDocName          As String                 '����ҽ������
    strOperationlevel              As String                 '�����ȼ�
    strOperationSiteCode          As String                 '������λ
    strhemorrhageVolume            As String                 '��Ѫ��
    strOperationSource             As String                 '������Դ
    strpreOPDiagCode             As String                 '��ǰ��ϱ���
    strpreOPDiagName            As String                 '��ǰ�������
    strpostOPDiagCode            As String                 '������ϱ���
    strpostOPDiagName           As String                 '�����������
    strNnis                         As String                 '�������շּ�
    strisSelectiveOperation       As String                 '��������
End Type

'<-begin------------------------------------------------------------------------1.9  Exams�����ǩ:----------------------------------------------------->
Public Type ExamInfo
    strReportID                    As String                 '���鱨�浥��
    strExamItemCode               As String                 '������Ŀ����
    strExamItemName              As String                 '������Ŀ����
    strSampleCollectTime          As String                 '����ʱ��
    strSampleCode                  As String                 '�걾����
    strSampleName                 As String                 '�걾����
    strSampleCollectOpporunity    As String                 '����ʱ��
    strApplyNO                     As String                 '���뵥���
    strApplicantID                 As String                 '����ҽ������
    strApplicantName              As String                 '����ҽ������
    strApplicantDeptID            As String                 '�������id
    strApplicantDeptName         As String                 '�����������
    strReporterID                  As String                 '����ҽ������
    strReporterName               As String                 '����ҽ������
    strReportTime                  As String                 '����ʱ��
    strExamAim                    As String                 '����Ŀ��
End Type

Public Type ExamItem
    strReportID                    As String                 '���鱨�浥��
    strReportItemID               As String                 '���鱨����ϸ���
    strindicatorCode               As String                 '����ָ�����
    strindicatorName              As String                 '����ָ������
    strindicatorename              As String                 '����ָ��Ӣ�ļ��
    strExamResult                  As String                 '�����
    strExamResultUnit             As String                 '���������λ
    strreferenceResult             As String                 '�ο����
    strupperlimit                  As String                 '�ο�ֵ����
    strlowerlimit                  As String                 '�ο�ֵ����
    strcriticalFlag                As String                 'Σ��ֵ��־
End Type
'<-begin------------------------------------------------------------------------1.10 ImageinfosӰ���ǩ:----------------------------------------------------->
Public Type ImageInfo
    strImageID                 As String                 'Ӱ����
    strImageCode               As String                 '�����Ŀ����
    strImageName              As String                 '�����Ŀ����
    strperformMethod           As String                 '��鷽��
    strperformSite             As String                 '��鲿λ
    strimagingPosition         As String                 '��λ
    strimagingDiagnosis        As String                 'Ӱ�����
    strApplyNO                 As String                 '���뵥���
    strApplicantID             As String                 '����ҽ������
    strApplicantName          As String                 '����ҽ������
    strApplicantDeptID        As String                 '�������id
    strApplicantDeptName     As String                 '�����������
    strPerformTime             As String                 '���ʱ��
    strReporterID              As String                 '����ҽ������
    strReporterName           As String                 '����ҽ������
    strReportTime              As String                 '����ʱ��
    strperformAim             As String                 '���Ŀ��
End Type

'<-begin------------------------------------------------------------------------1.11 SpecialExams��������Ŀ��ǩ:----------------------------------------------------->
Public Type SpecialExam
    strReportID                As String                 '�����鱨����
    strExamItemCode           As String                 '�����Ŀ����
    strExamItemName          As String                 '�����Ŀ����
    strExamConclusion          As String                 '������
    strApplyNO                 As String                 '���뵥���
    strApplicantID             As String                 '����ҽ������
    strApplicantName          As String                 '����ҽ������
    strApplicantDeptID        As String                 '�������id
    strApplicantDeptName     As String                 '�����������
    strPerformTime             As String                 '���ʱ��
    strReporterID              As String                 '����ҽ������
    strReporterName           As String                 '����ҽ������
    strReportTime              As String                 '����ʱ��
    strperformAim             As String                 '���Ŀ��
End Type

'<-begin------------------------------------------------------------------------1.12    ElectronicMedical ������Ӳ�����ǩ:----------------------------------------------------->
Public Type ElectronicMedical
    strElectronicMedicalID        As String                 '���Ӳ������
    strChiefComplaint              As String                 '����
    strMedicalHistory              As String                 '�ֲ�ʷ
    strPastHistory                 As String                 '����ʷ
    strPersonalHistory             As String                 '����ʷ
    strFamilyDiseaseHistory       As String                 '����ʷ
    strMenstrualHistory            As String                 '�¾�ʷ
    strObstericalHistory           As String                 '����ʷ
    strRecordDocID                As String                 '��¼ҽ������
    strRecordDocName             As String                 '��¼ҽ������
    strRecordTime                  As String                 '��¼ʱ��
End Type
'<-begin------------------------------------------------------------------------1.13    AdmissionRecord��Ժ��¼��ǩ:----------------------------------------------------->
Public Type AdmissionRecord
    strAdmissionRecordID              As String                 '��Ժ��¼ID
    strAdmissionRecordType            As String                 '��Ժ��¼����
    strChiefComplaint                  As String                 '����
    strMedicalHistory                  As String                 '�ֲ�ʷ
    strPastHistory                     As String                 '����ʷ
    strPersonalHistory                 As String                 '����ʷ
    strFamilyDiseaseHistory           As String                 '����ʷ
    strMenstrualHistory                As String                 '�¾�ʷ
    strObstericalHistory               As String                 '����ʷ
    strOperationHistory                As String                 '����ʷ
    strTransfusionHistory              As String                 '��Ѫʷ
    strInfectionHistory                As String                 '��Ⱦ��ʷ
    strVaccinationHistory              As String                 'Ԥ������ʷ
    strPhysicalExam                    As String                 '�����
    strSpecialExam                     As String                 'ר�Ƽ��
    strAuxiliaryExam                   As String                 '���������
    strRecordDocID                    As String                 '��¼ҽ������
    strRecordDocName                 As String                 '��¼ҽ������
    strRecordTime                      As String                 '��¼ʱ��
End Type


'<-begin------------------------------------------------------------------------1.14    Progressnotes����¼��ǩ:----------------------------------------------------->
Public Type HZYYProgressNote
    strProgressNoteID         As String                 '����¼ID
    strProgressNoteType       As String                 '����¼����
    strProgressNoteContent    As String                 '����¼����
    strRecordDocID            As String                 '��¼ҽ������
    strRecordDocName         As String                 '��¼ҽ������
    strRecordTime              As String                 '��¼ʱ��
End Type


'<-begin------------------------------------------------------------------------1.15    vitalsigns����������ǩ:----------------------------------------------------->
Public Type VitalSign
    strVitalSignNO        As String                 '������¼���
    strTemperature          As String                 '����
    strSbp                  As String                 '����ѹ
    strDbp                  As String                 '����ѹ
    strBreathingRate       As String                 '����Ƶ��
    strPulseRate           As String                 '����
    strHeartRate           As String                 '����
    strPainScore           As String                 '��ʹ����
    strHour24Amountin     As String                 '24Сʱ����
    strHour24Amountout    As String                 '24Сʱ����
    strTestTime            As String                 '�ⶨʱ��
End Type

'<-begin------------------------------------------------------------------------1.16    pathologicalExams������Ϣ��ǩ:----------------------------------------------------->
Public Type PathologicalExam
    strReportID                As String                 '���鱨�浥��
    strExamItemCode           As String                 '�����Ŀ����
    strExamItemName          As String                 '�����Ŀ����
    strSampleName             As String                 '�걾����
    strPathologicDiagnosis     As String                 '�������
    strApplyNO                 As String                 '���뵥���
    strApplicantID             As String                 '����ҽ������
    strApplicantName          As String                 '����ҽ������
    strApplicantDeptID        As String                 '�������id
    strApplicantDeptName     As String                 '�����������
    strPerformTime             As String                 '�ͼ�ʱ��
    strReporterID              As String                 '����ҽ������
    strReporterName           As String                 '����ҽ������
    strReportTime              As String                 '����ʱ��
    strExamAim                As String                 '���Ŀ��
End Type


'<-begin------------------------------------------------------------------------1.17    bacterialreportsϸ�����������ǩ:----------------------------------------------------->
Public Type BacterialReportInfo
      strReportID                      As String                 '���鱨�浥��
      strExamItemCode                 As String                 '������Ŀ����
      strExamItemName                As String                 '������Ŀ����
      strSampleCollectTime            As String                 '����ʱ��
      strSampleCode                    As String                 '�걾����
      strSampleName                   As String                 '�걾����
      strSampleCollectOpporunity      As String                 '����ʱ��
      strApplyNO                       As String                 '���뵥���
      strApplicantID                   As String                 '����ҽ������
      strApplicantName                As String                 '����ҽ������
      strApplicantDeptID      As String                 '�������id
      strApplicantDeptName   As String                 '�����������
      strReporterID            As String                 '����ҽ������
      strReporterName         As String                 '����ҽ������
      strReportTime            As String                 '����ʱ��
End Type

Public Type BacterialReportItem
    strReportID            As String                 '���鱨�浥��
    strReportItemID       As String                 'ϸ��������ϸ���
    strExamItemResult     As String                 '������
End Type

'<-begin------------------------------------------------------------------------1.18    drugsensitivesҩ������������Ϣ��ǩ:----------------------------------------------------->
Public Type DrugSensitiveinfo
    strReportID                    As String                 '���鱨�浥��
    strExamItemCode               As String                 '������Ŀ����
    strExamItemName              As String                 '������Ŀ����
    strSampleCollectTime          As String                 '����ʱ��
    strSampleCode                  As String                 '�걾����
    strSampleName                 As String                 '�걾����
    strSampleCollectOpporunity    As String                 '����ʱ��
    strgermID                      As String                 'ϸ��id
    strgermName                   As String                 'ϸ������
    strApplyNO                     As String                 '���뵥���
    strApplicantID                 As String                 '����ҽ������
    strApplicantName              As String                 '����ҽ������
    strApplicantDeptID            As String                 '�������id
    strApplicantDeptName         As String                 '�����������
    strReporterID                  As String                 '����ҽ������
    strReporterName               As String                 '����ҽ������
    strReportTime                  As String                 '����ʱ��
End Type

Public Type DrugSensitiveItem
    strReportID                    As String                 '���鱨�浥��
    strReportItemID               As String                 'ҩ����ϸ���
    strantibioticID                As String                 '����ҩ��id
    strantibioticName             As String                 '����ҩ��
    strsensitivity                  As String                 '���г̶�
    strmic                          As String                 '����־�Ũ��
End Type

'<-end------------------------------------------------------------------------xml----------------------------------------------------->

Public Type OptPrescription
    udtOptPresInfo                  As OPTPRESCRIPTIONSINFO        '������Ϣ
    colPresItem                     As Collection                    '������ϸ
End Type

'������Ƭҽ����Ϣ�ṹ
Public Type HerbMedicalOrder
    udtHerbInfo         As HerbMedicalOrderInfo
    colItemHerb  As Collection                       'HerbMedicalOrderItem�ļ���
End Type
'ҽ����ϢXML��ǩ
Public Type Order
    colNonMedical                   As Collection           '������ҩ��ҽ��
    colMedical                      As Collection           '����ҩ��ҽ�����
    colHerbMedical                  As Collection           '������Ƭҽ����ϢHerbMedicalOrder
End Type
'
Public Type HZYYExam
    udtInfo As ExamInfo
    colExamItem As Collection                                      '����������ϸ ExamItem
End Type

'ϸ�����������ǩ
Public Type BacterialReport
    udtInfo As BacterialReportInfo
    colItem As Collection                       'BacterialReportItem�ļ���
End Type


Public Type DrugSensitive
    udtInfo As DrugSensitiveinfo
    colItem As Collection                       'DrugSensitiveItem�ļ���
End Type
