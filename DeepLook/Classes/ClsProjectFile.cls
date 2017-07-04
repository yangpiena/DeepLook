VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsProjectFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*************************************************************************
''               人人为我，我为人人
''      枕善居VB及.NET源码博客汉化收藏整理
''网    站：http://www.Mndsoft.com/
''e-mail  ：mndsoft@126.com
''发布日期：2009-10-08 10:12:30
''QQ      ：88382850
''   如果您有新的、好的代码可以提供给枕善居上发布，让大家学习哦!
''----------------------------------------------------------------------
'**系统名称：VB及.NET工程源代码扫描分析工具 V4.12.0
'**模块描述：
'**模 块 名：ClsProjectFile
'**创 建 人：
'**汉 化 者：枕善居(mndsoft)
'**日    期：2009-10-08 10:12:30
'**修 改 人：
'**日    期：
'**描    述：
'**版    本：V4.12.0
'*************************************************************************

Option Explicit

' -----------------------------------------------------------------------------------------------
Private PrjName As String
Private PrjTitle As String
Private PrjType As String
Private PrjStartUpItem As String
Private PrjPath As String
Private PrjRefComCount As Long

Private PrjClass As Long
Private PrjModule As Long
Private PrjForm As Long
Private PrjUserControl As Long
Private PrjUserDocument As Long
Private PrjPropertyPage As Long
Private PrjDesigner As Long

Private PrjTotalLines As Long
Private PrjTotalLinesNB As Long
Private PrjTotalCommentLines As Long

Private PrjVersion As String

Private PrjVariables As Long

Private PrjConstants As Long
Private PrjTypes As Long
Private PrjEnums As Long

Private PrjRefCom() As String
Private PrjRefCom_Elements As Long
Private PrjDecDll() As String
Private PrjDecDll_Elements As Long
Private PrjCreateObj() As String
Private PrjCreateObj_Elements As Long

Private PrjSubLines As Long
Private PrjFuncLines As Long
Private PrjPropLines As Long

Private PrjEXEPath As String
Private PrjEXEFName As String

Public Enum PFileDataTypes
    SPF_DecDll = 0
    SPF_CreateObj = 1
    SPF_RefCom = 2
End Enum
'------------------------------------------------------------------------------

Private Function FixProjType(PrjType As String)
    FixProjType = UCase$(PrjType)

    Select Case FixProjType
        Case "EXE"
            FixProjType = "Standard EXE"
        Case "OLEEXE"
            FixProjType = "ActiveX EXE"
        Case "OLEDLL"
            FixProjType = "ActiveX DLL"
        Case Else
            FixProjType = "ActiveX Control"
    End Select
End Function

Private Sub Class_Initialize()
    ReDim PrjRefCom(6) As String
    ReDim PrjDecDll(6) As String
    ReDim PrjCreateObj(6) As String
End Sub

'------------------------------------------------------------------------------

Property Get ProjectName() As String
    ProjectName = PrjName
End Property

Property Let ProjectName(NewName As String)
    PrjName = NewName
End Property

Property Get ProjectRefComCount() As Long
    ProjectRefComCount = PrjRefComCount
End Property

Property Let ProjectRefComCount(AddAmount As Long)
    PrjRefComCount = PrjRefComCount + AddAmount
End Property

Property Get ProjectTitle() As String
    ProjectTitle = PrjTitle
End Property

Property Let ProjectTitle(NewTitle As String)
    PrjTitle = NewTitle
End Property

Property Get ProjectProjectType() As String
    ProjectProjectType = FixProjType(PrjType)
End Property

Property Let ProjectProjectType(NewType As String)
    PrjType = NewType
End Property

Property Get ProjectStartupItem() As String
    ProjectStartupItem = PrjStartUpItem
End Property

Property Let ProjectStartupItem(NewStartupItem As String)
    PrjStartUpItem = NewStartupItem
End Property

Property Get ProjectPath() As String
    ProjectPath = PrjPath
End Property

Property Let ProjectPath(NewPath As String)
    PrjPath = NewPath
End Property

Property Get ProjectSubLines() As Long
    ProjectSubLines = PrjSubLines
End Property

Property Let ProjectSubLines(AddAmount As Long)
    PrjSubLines = PrjSubLines + AddAmount
End Property

Property Get ProjectFuncLines() As Long
    ProjectFuncLines = PrjFuncLines
End Property

Property Let ProjectFuncLines(AddAmount As Long)
    PrjFuncLines = PrjFuncLines + AddAmount
End Property

Property Get ProjectPropLines() As Long
    ProjectPropLines = PrjPropLines
End Property

Property Let ProjectPropLines(AddAmount As Long)
    PrjPropLines = PrjPropLines + AddAmount
End Property

Property Get ProjectForms() As Long
    ProjectForms = PrjForm
End Property

Property Let ProjectForms(AddAmount As Long)
    PrjForm = PrjForm + AddAmount
End Property

Property Get ProjectModules() As Long
    ProjectModules = PrjModule
End Property

Property Let ProjectModules(AddAmount As Long)
    PrjModule = PrjModule + AddAmount
End Property

Property Get ProjectClasses() As Long
    ProjectClasses = PrjClass
End Property

Property Let ProjectClasses(AddAmount As Long)
    PrjClass = PrjClass + AddAmount
End Property

Property Get ProjectUserControls() As Long
    ProjectUserControls = PrjUserControl
End Property

Property Let ProjectUserControls(AddAmount As Long)
    PrjUserControl = PrjUserControl + AddAmount
End Property

Property Get ProjectUserDocuments() As Long
    ProjectUserDocuments = PrjUserDocument
End Property

Property Let ProjectUserDocuments(AddAmount As Long)
    PrjUserDocument = PrjUserDocument + AddAmount
End Property

Property Get ProjectPropertyPages() As Long
    ProjectPropertyPages = PrjPropertyPage
End Property

Property Let ProjectPropertyPages(AddAmount As Long)
    PrjPropertyPage = PrjPropertyPage + AddAmount
End Property

Property Get ProjectDesigners() As Long
    ProjectDesigners = PrjDesigner
End Property

Property Let ProjectDesigners(AddAmount As Long)
    PrjDesigner = PrjDesigner + AddAmount
End Property

Property Get ProjectLines() As Long
    ProjectLines = PrjTotalLines
End Property

Property Let ProjectLines(AddAmount As Long)
    PrjTotalLines = PrjTotalLines + AddAmount
End Property

Property Get ProjectCommentLines() As Long
    ProjectCommentLines = PrjTotalCommentLines
End Property

Property Let ProjectCommentLines(AddAmount As Long)
    PrjTotalCommentLines = PrjTotalCommentLines + AddAmount
End Property

Property Get ProjectLinesNB() As Long
    ProjectLinesNB = PrjTotalLinesNB
End Property

Property Let ProjectLinesNB(AddAmount As Long)
    PrjTotalLinesNB = PrjTotalLinesNB + AddAmount
End Property

Property Get ProjectVariables() As Long
    ProjectVariables = PrjVariables
End Property

Property Let ProjectVariables(AddAmount As Long)
    PrjVariables = PrjVariables + AddAmount
End Property

Property Get ProjectConstants() As Long
    ProjectConstants = PrjConstants
End Property

Property Let ProjectConstants(AddAmount As Long)
    PrjConstants = PrjConstants + AddAmount
End Property

Property Get ProjectTypes() As Long
    ProjectTypes = PrjTypes
End Property

Property Let ProjectTypes(AddAmount As Long)
    PrjTypes = PrjTypes + AddAmount
End Property

Property Get ProjectEnums() As Long
    ProjectEnums = PrjEnums
End Property

Property Let ProjectEnums(AddAmount As Long)
    PrjEnums = PrjEnums + AddAmount
End Property

Property Get ProjectVersion() As String
    ProjectVersion = PrjVersion
End Property

Property Let ProjectVersion(AddString As String)
    PrjVersion = PrjVersion & AddString
End Property

Property Get ProjectRefCom(ItemIndex As Long) As String
    ProjectRefCom = PrjRefCom(ItemIndex)
End Property

Property Get ProjectRefComArrayCount() As Long
    ProjectRefComArrayCount = PrjRefCom_Elements
End Property

Property Get ProjectDecDlls(ItemIndex As Long) As String
    ProjectDecDlls = PrjDecDll(ItemIndex)
End Property

Property Get ProjectDecDllsArrayCount() As Long
    ProjectDecDllsArrayCount = PrjDecDll_Elements
End Property

Sub AddData(AddDataType As PFileDataTypes, AddData As String)
    If AddDataType = SPF_CreateObj Then
        PrjCreateObj_Elements = PrjCreateObj_Elements + 1
        If UBound(PrjCreateObj) < PrjCreateObj_Elements Then
            ReDim Preserve PrjCreateObj(PrjCreateObj_Elements + 6) As String
        End If

        PrjCreateObj(PrjCreateObj_Elements) = AddData
    ElseIf AddDataType = SPF_DecDll Then
        PrjDecDll_Elements = PrjDecDll_Elements + 1
        If UBound(PrjDecDll) < PrjDecDll_Elements Then
            ReDim Preserve PrjDecDll(PrjDecDll_Elements + 6) As String
        End If

        PrjDecDll(PrjDecDll_Elements) = AddData
        If AddDataType = SPF_RefCom Then
            PrjRefCom_Elements = PrjRefCom_Elements + 1
            If UBound(PrjRefCom) < PrjRefCom_Elements Then
                ReDim Preserve PrjRefCom(PrjRefCom_Elements + 6) As String
            End If

            PrjRefCom(PrjRefCom_Elements) = AddData
        End If
    End If
End Sub

Property Get ProjectCreateObjects(ItemIndex As Long) As String
    ProjectCreateObjects = PrjCreateObj(ItemIndex)
End Property

Property Get ProjectCreateObjectsArrayCount() As Long
    ProjectCreateObjectsArrayCount = PrjCreateObj_Elements
End Property

Property Get ProjectEXEPath() As String
    ProjectEXEPath = PrjEXEPath
End Property

Property Let ProjectEXEPath(Path As String)
    If Right$(Path, 1) <> "\" Then Path = Path & "\"
    PrjEXEPath = Path
End Property

Property Get ProjectEXEFName() As String
    ProjectEXEFName = PrjEXEFName
End Property

Property Let ProjectEXEFName(FName As String)
    PrjEXEFName = FName
End Property
