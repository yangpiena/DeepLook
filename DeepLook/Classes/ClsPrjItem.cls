VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsProjectItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*************************************************************************
''               人人为我，我为人人
''      枕善居VB及.NET源码博客汉化收藏整理
''网    站：http://www.Mndsoft.com/
''e-mail  ：mndsoft@126.com
''发布日期：2009-10-08 10:12:34
''QQ      ：88382850
''   如果您有新的、好的代码可以提供给枕善居上发布，让大家学习哦!
''----------------------------------------------------------------------
'**系统名称：VB及.NET工程源代码扫描分析工具 V4.12.0
'**模块描述：
'**模 块 名：ClsProjectItem
'**创 建 人：
'**汉 化 者：枕善居(mndsoft)
'**日    期：2009-10-08 10:12:34
'**修 改 人：
'**日    期：
'**描    述：
'**版    本：V4.12.0
'*************************************************************************

Option Explicit

' -----------------------------------------------------------------------------------------------
Private ItemName As String
Private ItemControls As Long
Private ItemControlsNoArrays As Long

Private ItemInControls As Boolean
Private ItemSeenAttributes As Boolean

Private ItemCodeLines As Long
Private ItemCodeLinesNoBlanks As Long
Private ItemCommentLines As Long
Private ItemHybridLines As Long

Private ItemSubs() As String
Private ItemSubs_Elements As Long
Private ItemFunctions() As String
Private ItemFunctions_Elements As Long
Private ItemProperties() As String
Private ItemProperties_Elements As Long
Private ItemEvents() As String
Private ItemEvents_Elements As Long

Private ItemVariables As Long
Private ItemConstants As Long
Private ItemTypes As Long
Private ItemEnums As Long

Private ItemStatementDO As Long
Private ItemStatementWHILE As Long
Private ItemStatementIF As Long
Private ItemStatementFOR As Long
Private ItemStatementSELECT As Long

Public Enum StatementType
    STDO = 1
    STWHILE = 2
    STIF = 3
    STFOR = 4
    STSELECT = 5
End Enum

Public Enum SPFTypes
    SPF_Sub = 0
    SPF_Function = 1
    SPF_Property = 2
    SPF_Event = 3
End Enum
'------------------------------------------------------------------------------

Sub SortArrays()
    BubbleSortArray ItemSubs
    BubbleSortArray ItemFunctions
    BubbleSortArray ItemProperties
    BubbleSortArray ItemEvents
End Sub

Property Get PrjItemName() As String
    PrjItemName = ItemName
End Property

Property Let PrjItemName(NewName As String)
    ItemName = NewName
End Property

Property Get PrjItemControls() As Long
    PrjItemControls = ItemControls
End Property

Property Let PrjItemControls(AddAmount As Long)
    ItemControls = ItemControls + AddAmount
End Property

Property Get PrjItemControlsNoArrays() As Long
    PrjItemControlsNoArrays = ItemControlsNoArrays
End Property

Property Let PrjItemControlsNoArrays(AddAmount As Long)
    ItemControlsNoArrays = ItemControlsNoArrays + AddAmount
End Property

Property Get PrjItemSeenAttributes() As Boolean
    PrjItemSeenAttributes = ItemSeenAttributes
End Property

Property Let PrjItemSeenAttributes(YesNo As Boolean)
    ItemSeenAttributes = YesNo
End Property

Property Get PrjItemInControls() As Boolean
    PrjItemInControls = ItemInControls
End Property

Property Let PrjItemInControls(YesNo As Boolean)
    ItemInControls = YesNo
End Property

Property Get PrjItemCodeLines() As Long
    PrjItemCodeLines = ItemCodeLines
End Property

Property Let PrjItemCodeLines(AddAmount As Long)
    ItemCodeLines = ItemCodeLines + AddAmount
End Property

Property Get PrjItemCodeLinesNoBlanks() As Long
    PrjItemCodeLinesNoBlanks = ItemCodeLinesNoBlanks
End Property

Property Let PrjItemCodeLinesNoBlanks(AddAmount As Long)
    ItemCodeLinesNoBlanks = ItemCodeLinesNoBlanks + AddAmount
End Property

Property Get PrjItemCommentLines() As Long
    PrjItemCommentLines = ItemCommentLines
End Property

Property Let PrjItemCommentLines(AddAmount As Long)
    ItemCommentLines = ItemCommentLines + AddAmount
End Property

Property Get PrjItemHybridLines() As Long
    PrjItemHybridLines = ItemHybridLines
End Property

Property Let PrjItemHybridLines(AddAmount As Long)
    ItemHybridLines = ItemHybridLines + AddAmount
End Property

Property Get PrjItemItemSubs(ItemIndex As Long) As String
    PrjItemItemSubs = ItemSubs(ItemIndex)
End Property

Property Get PrjItemItemSubsCount() As Long
    PrjItemItemSubsCount = ItemSubs_Elements
End Property

Property Get PrjItemItemFunctions(ItemIndex As Long) As String
    PrjItemItemFunctions = ItemFunctions(ItemIndex)
End Property

Property Get PrjItemItemFunctionsCount() As Long
    PrjItemItemFunctionsCount = ItemFunctions_Elements
End Property

Property Get PrjItemItemProperties(ItemIndex As Long) As String
    PrjItemItemProperties = ItemProperties(ItemIndex)
End Property

Property Get PrjItemItemPropertiesCount() As Long
    PrjItemItemPropertiesCount = ItemProperties_Elements
End Property

Property Get PrjItemItemEvents(ItemIndex As Long) As String
    PrjItemItemEvents = ItemEvents(ItemIndex)
End Property

Property Get PrjItemItemEventsCount() As Long
    PrjItemItemEventsCount = ItemEvents_Elements
End Property

Sub AddSPF(AddName As String, SPFType As SPFTypes)
    If SPFType = SPF_Sub Then
        ItemSubs_Elements = ItemSubs_Elements + 1
        If UBound(ItemSubs) < ItemSubs_Elements Then
            ReDim Preserve ItemSubs(ItemSubs_Elements + 10) As String
        End If
    
        ItemSubs(ItemSubs_Elements) = AddName
    ElseIf SPFType = SPF_Function Then
        ItemFunctions_Elements = ItemFunctions_Elements + 1
        If UBound(ItemFunctions) < ItemFunctions_Elements Then
            ReDim Preserve ItemFunctions(ItemFunctions_Elements + 10) As String
        End If
        
        ItemFunctions(ItemFunctions_Elements) = AddName
    ElseIf SPFType = SPF_Property Then
        ItemProperties_Elements = ItemProperties_Elements + 1
        If UBound(ItemProperties) < ItemProperties_Elements Then
            ReDim Preserve ItemProperties(ItemProperties_Elements + 10) As String
        End If
        
        ItemProperties(ItemProperties_Elements) = AddName
    ElseIf SPFType = SPF_Event Then
        ItemEvents_Elements = ItemEvents_Elements + 1
        If UBound(ItemEvents) < ItemEvents_Elements Then
            ReDim Preserve ItemEvents(ItemEvents_Elements + 10) As String
        End If
        
        ItemEvents(ItemEvents_Elements) = AddName
    End If
End Sub

Property Get PrjItemVariables() As Long
    PrjItemVariables = ItemVariables
End Property

Property Let PrjItemVariables(AddAmount As Long)
    ItemVariables = ItemVariables + AddAmount
End Property

Sub AddToStatement(STTYPE As StatementType)
    If STTYPE = STDO Then
        ItemStatementDO = ItemStatementDO + 1
    ElseIf STTYPE = STFOR Then
        ItemStatementFOR = ItemStatementFOR + 1
    ElseIf STTYPE = STIF Then
        ItemStatementIF = ItemStatementIF + 1
    ElseIf STTYPE = STWHILE Then
        ItemStatementWHILE = ItemStatementWHILE + 1
    Else
        ItemStatementSELECT = ItemStatementSELECT + 1
    End If
End Sub

Property Get PrjItemStatements(STTYPE As StatementType)
    If STTYPE = STDO Then
        PrjItemStatements = ItemStatementDO
    ElseIf STTYPE = STFOR Then
        PrjItemStatements = ItemStatementFOR
    ElseIf STTYPE = STIF Then
        PrjItemStatements = ItemStatementIF
    ElseIf STTYPE = STWHILE Then
        PrjItemStatements = ItemStatementWHILE
    Else
        PrjItemStatements = ItemStatementSELECT
    End If
End Property

Property Get PrjItemConstants() As Long
    PrjItemConstants = ItemConstants
End Property

Property Let PrjItemConstants(AddAmount As Long)
    ItemConstants = PrjItemConstants + AddAmount
End Property

Property Get PrjItemTypes() As Long
    PrjItemTypes = ItemTypes
End Property

Property Let PrjItemTypes(AddAmount As Long)
    ItemTypes = PrjItemTypes + AddAmount
End Property

Property Get PrjItemEnums() As Long
    PrjItemEnums = ItemEnums
End Property

Property Let PrjItemEnums(AddAmount As Long)
    ItemEnums = PrjItemEnums + AddAmount
End Property

Private Sub Class_Initialize()
    ReDim ItemSubs(10) As String
    ReDim ItemFunctions(10) As String
    ReDim ItemProperties(10) As String
    ReDim ItemEvents(10) As String
End Sub
