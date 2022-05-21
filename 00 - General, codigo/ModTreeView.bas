Attribute VB_Name = "ModTreeView"
'           .---.        .-----------
'          /     \  __  /    ------
'         / /     \(..)/    -----
'        //////   ' \/ `   ---
'       //// / // :    : ---
'      // /   /  /`    '--
'     // /        //..\\
'   o===|========UU====UU=====-  -==========================o
'                '//||\\`
'                       DEVELOPED BY JGH
'
'   -=====================|===o  o===|======================-
Option Explicit
'----------------------------------------------------------------------------------------
' Module    : ModTreeView
' DateTime  : 06/03/2013
' Author    : José García Herruzo
' Purpose   : This module contents functions and procedures applied to treeview
' References: N/A
' Functions :
'               1-AddParentNode
'               2-AddChildrenNode
' Procedures: N/A
' Updates   :
'       DATE        USER    DESCRIPTION
'       N/A
'----------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
' Function  : AddParentNode
' DateTime  : 06/03/2013
' Author    : José García Herruzo
' Purpose   : Add parent node
' Arguments :
'             tv_myTreeView            --> Treeview object
'             str_myNodeName           --> Node name
'             str_myNodeKey            --> Node Key
'---------------------------------------------------------------------------------------
Public Function AddParentNode(ByVal tv_myTreeView As TreeView, ByVal str_myNodeName As String, ByVal str_myNodeKey As String) As Node

Dim nd As Node

Set nd = tv_myTreeView.Nodes.Add(, , str_myNodeKey, str_myNodeName)

nd.Tag = str_myNodeName
   
Set AddParentNode = nd
   
Set nd = Nothing
   
End Function

'---------------------------------------------------------------------------------------
' Function  : AddChildrenNode
' DateTime  : 06/03/2013
' Author    : José García Herruzo
' Purpose   : Add new children node
' Arguments :
'             tv_myTreeView            --> Treeview object
'             str_myNodeName           --> Parent node name
'             str_myChildrenNodeName   --> Children node name
'             str_myNodeKey            --> Node Key
'---------------------------------------------------------------------------------------
Public Function AddChildrenNode(ByVal tv_myTreeView As TreeView, ByVal str_myNodeName As Node, ByVal str_myChildrenNodeName As String, _
                             ByVal str_myNodeKey As String) As Node

Dim nd As Node

Set nd = tv_myTreeView.Nodes.Add(str_myNodeName, tvwChild, str_myNodeKey, str_myChildrenNodeName)

nd.Tag = str_myChildrenNodeName

Set AddChildrenNode = nd

Set nd = Nothing

End Function
