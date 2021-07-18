VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ITA_XMLClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' classe che consente di gestire un file XML di personalizzazioni

' definizione variabili a livello classe
Dim objXML As DOMDocument
Public objDOM As DOMDocument
Public objNodeList As Object
Public objNode As Object
Public objNodeSearch As IXMLDOMNode

Public Function InizializzaXML(ByVal strXMLName As String) As Boolean
    ' consente di caricare un file XML in un oggetto
    Set objXML = New DOMDocument
    
    If Not (objXML Is Nothing) And strXMLName <> "" Then
        'Call objXML.Load(strXMLName)
        Call objXML.Load(strXMLName)
        Set objDOM = objXML
        InizializzaXML = True
    Else
        Set objDOM = Nothing
        InizializzaXML = True
    End If
    
    
        
End Function


Public Function DebugPrintXMLNode(n As IXMLDOMNode)

    Dim i As Integer
    Debug.Print "node value: " & n.nodeValue
    
    Call DebugPrintXMLNodeAttributes(n)
    If n.hasChildNodes = True Then
        For i = 0 To n.childNodes.length - 1
            'Debug.Print "node value: " & n.ChildNodes(i).NodeValue
            Call DebugPrintXMLNode(n.childNodes(i))
        Next
    End If

End Function

Public Function DebugPrintXMLNodeAttributes(n As IXMLDOMNode)

    Dim i As Integer
    Debug.Print "attributes for node value: " & n.nodeName
    
    If Not (n.Attributes Is Nothing) Then
        For i = 0 To n.Attributes.length - 1
            Debug.Print "attr nome: " & n.Attributes(i).baseName & " attr value: " & n.Attributes(i).nodeValue
        Next
    End If

End Function

Public Function getAttrValue(ByVal nl As IXMLDOMNodeList, ByVal attrName As String) As Variant
    Dim n As IXMLDOMElement
    
    If Not (nl Is Nothing) And attrName <> "" Then
        If Not (nl.Item(0) Is Nothing) Then
            
            Set n = nl.Item(0)
            
            getAttrValue = n.getAttributeNode(attrName).nodeValue
                        
            Set n = Nothing
            
        End If
        
        
    End If

End Function


Public Function getNodeAttrValue(ByVal n As IXMLDOMElement, ByVal attrName As String) As Variant
    
    If Not (n Is Nothing) And attrName <> "" Then
        getNodeAttrValue = n.getAttributeNode(attrName).nodeValue
    End If

End Function


Public Function GetNodeList(ByVal strNodeName As String) As Object
    ' funzione che restituisce un nodo del file XML in un oggetto
    
    Set objNodeList = objXML.getElementsByTagName(strNodeName)
    
End Function
Public Function GetNode(ByVal strNodeName As String) As Object
    ' funzione che restituisce un nodo del file XML in un oggetto
    Dim objNode As Object
        
    Set objNode = objXML.nodeFromID(strNodeName)
    
End Function

Public Function GetNodeValue(ByVal strNodeName As String) As String
    ' funzione che restituisce il valore di un nodo del file XML in una string
    Dim objNode As Object
    
    Set objNode = GetNode(strNodeName)
    
    If Not (objNode Is Nothing) Then
        GetNodeValue = objNode.Text
    End If

End Function


'Lettura di un valore dati sezione e chiave.
Public Function GetValue(Section As String, key As String) As String
    Dim objNode As IXMLDOMNode

    GetValue = ""
    If objDOM Is Nothing Then Exit Function
    
    Set objNode = objXML.selectSingleNode("//" & Section & "/" & key)
    If objNode Is Nothing Then Exit Function
    
    GetValue = objNode.Text
End Function


Function GetSearchList() As Variant
    Dim Args() As Variant
    Dim iNode As Integer

    
    If objXML Is Nothing Then
        Call MXNU.MsgBoxEX("Controllare il file XML di configurazione", vbCritical, "Lettura Lista Ricerche")
    Else
        MsgBox "apro il nodo ricerche"
        Set objNodeSearch = objXML.selectSingleNode("//configMxpKnos/Searches")
              
        If objNodeSearch Is Nothing Then
            Args() = Array("")
        Else
            MsgBox "Carico la lista di ricerche"
            For iNode = 1 To objNodeSearch.childNodes.length
                MsgBox "Carico la ricerca: " & objNodeSearch.childNodes.Item(iNode).Text
                Args(iNode) = objNodeSearch.childNodes.Item(iNode).Text
            Next
            Set objNodeSearch = Nothing
        End If
    
    End If
    
    GetSearchList = Args()
    
End Function
