Attribute VB_Name = "modExplorerTreeView"
Option Explicit
'You can load the Control like this, this happens To be In the Form
'load Event.
'
'LoadExplorerTreeView Me.tvFileTree
'
'
'Handle Click Events like this;
'
'Private Sub tvFileTree_NodeClick(ByVal node As MSComctlLib.node)
'    tvHandleClick node, Me.tvFileTree
'End Sub
'
'Private Sub tvFileTree_Expand(ByVal node As MSComctlLib.node)
'    tvHandleClick node, Me.tvFileTree
'End Sub


Public Function GetDriveIcon(drv As Drive) As String
    '************************************************************
    ' Author: EDDIE MERKEL
    ' Date: Thursday, Jul 5, 2001   Time: 10:47 AM
    ' Purpose: get back the key value in the image list for the icon
    '            for that specific type of drive
    ' Arguments: drv as drive, the drive that we are working with
    ' Return: string value equal to the key of the image in the image list
    '************************************************************
    Select Case drv.DriveType
        Case 0: GetDriveIcon = "drive" 'here "drive" is the key in my imagelist control for the icon i want to use for a drive node
        Case 1: GetDriveIcon = "floppy"
        Case 2: GetDriveIcon = "drive"
        Case 3: GetDriveIcon = "network"
        Case 4: GetDriveIcon = "cd"
        Case 5: GetDriveIcon = "drive"
    End Select
    
End Function

Public Function GetTruePath(oNode As Node) As String
    '************************************************************
    ' Author: EDDIE MERKEL
    ' Date: Thursday, Jul 5, 2001   Time: 10:47 AM
    ' Purpose: get back the fixed path to the directory represented by the
    '       nodes path
    ' Arguments: onode as node, the node that we are working with
    ' Return: string value equal to the path to the file system folder that
    '          relates to this node
    '************************************************************

    Dim sPath As String
    Dim vTmp As Variant
    
    ' gotta get the folder to start with
    vTmp = Split(oNode.FullPath, "\")
    sPath = oNode.FullPath
    sPath = Replace(sPath, vTmp(0) & "\", "")
    sPath = Replace(sPath, "\\", "\")
    
    GetTruePath = sPath
    
End Function

Public Sub LoadExplorerTreeView(ByRef tvFileTree As TreeView)
    '************************************************************
    ' Author: EDDIE MERKEL
    ' Date: Thursday, Jul 5, 2001   Time: 10:47 AM
    ' Purpose: Load the initial nodes into the treeview control
    ' Arguments: tvFileTree the treeview control that we are working with
    ' Return:
    '************************************************************

    Dim oNode As Node
    Dim fso As FileSystemObject
    Dim drv As Drive
    Dim fld As Folder
    Dim sDriveIcon As String
    
    On Error GoTo ErrorHandler
    
    Set fso = New FileSystemObject
    
    tvFileTree.Nodes.Add , , "ROOT", "COMPUTER", "computer"
    
    With fso
        For Each drv In .Drives
            sDriveIcon = GetDriveIcon(drv)
            Set oNode = tvFileTree.Nodes.Add("ROOT", tvwChild, drv.DriveLetter, drv.DriveLetter & ":\", sDriveIcon)
        Next
    End With
    
    
PROC_EXIT:

    If Not fso Is Nothing Then
        Set fso = Nothing
    End If
    
    If Not drv Is Nothing Then
        Set drv = Nothing
    End If
    
    Exit Sub   'Need to change this to Function, Property, etc.

ErrorHandler:

    Screen.MousePointer = vbDefault
    MsgBox Err.Source & vbCrLf & vbCrLf & CStr(Err.Number) & ": " & Err.Description, vbCritical, "LoadExplorerTreeView"
    Resume PROC_EXIT
    
    
End Sub

Private Sub LoadFolders(drv As Drive, oNode As Node, fld As Folder, tvFileTree As TreeView)
    '************************************************************
    ' Author: EDDIE MERKEL
    ' Date: Thursday, Jul 5, 2001   Time: 10:47 AM
    ' Purpose: add the directory structure of the computer to the treeview control
    '           a chunk at a time
    ' Arguments: drv as drive, the drive that we are working with
    '            oNode, the node to put the new nodes under
    '            fld, the folder to search under
    ' Return:
    '************************************************************
    Dim subFld As Folder
    Dim iSubFolderCount As Integer
    Dim sNode As Node
    Dim i As Integer
    Dim bRemoved As Boolean
    
    On Error Resume Next
    
    iSubFolderCount = fld.SubFolders.Count
    
    If iSubFolderCount <> 0 Then
        ' first remove any dummy node
        If oNode.Children > 0 Then
            i = oNode.Child.Index
            Do
               If InStr(1, oNode.Child.Key, "DUMMY", vbTextCompare) <> 0 Then
                    tvFileTree.Nodes.Remove oNode.Child.Key
                    bRemoved = True
                End If
            Loop While Not bRemoved
        End If
        
        For Each subFld In fld.SubFolders
            Set sNode = tvFileTree.Nodes.Add(oNode, tvwChild, oNode.FullPath & "\" & subFld.Name, subFld.Name, "closed", "open")
            If subFld.SubFolders.Count <> 0 Then
                ' add a dummy node
                tvFileTree.Nodes.Add sNode, tvwChild, sNode.FullPath & "\" & "DUMMY", "DUMMY"
            End If
        Next
    End If
    
    ' clean up
    If Not sNode Is Nothing Then
        Set sNode = Nothing
    End If
    
End Sub



Public Sub tvHandleClick(ByVal Node As MSComctlLib.Node, tvFileTree As TreeView)
    '************************************************************
    ' Author: EDDIE MERKEL
    ' Date: Thursday, Jul 5, 2001   Time: 03:52 PM
    ' Purpose: this just handles the users click in the treeview either on the node
    '           or on the + sign to fill the treeview control with drive and folder
    '           information just like windows explorer, works in conjunction with
    '           and dependent onloadTreeView, which loads the drives into the
    '           treeview and loadFolders which this calls to add child nodes
    ' Arguments: treeview node object
    ' Return:
    '************************************************************
    Dim drv As Drive
    Dim fso As FileSystemObject
    Dim fld As Folder
    Dim bFill As Boolean
    Dim bDrive As Boolean
    Dim bLoad As Boolean
    Dim vTmp As Variant
    Dim sPath As String
    Dim lRet As Long

    bDrive = False
    bFill = False
    bLoad = True

    If Node.Children = 0 Then
        bFill = True
    ElseIf InStr(1, Node.Child.Key, "DUMMY") Then
        bFill = True
    End If
    
    If bFill Then
        If Not Node Is Nothing Then
            Set fso = New FileSystemObject
    
            With Node
                With fso
                    If InStr(1, Node.Text, ":") Then
                        Set drv = .GetDrive(Left(Node.Text, 1))
                        bDrive = True
                    Else
                        bDrive = False
                        vTmp = Split(Node.FullPath, "\")
                        If UBound(vTmp) > 0 Then
                            Set drv = .GetDrive(Left(vTmp(1), 1))
                        Else
                            bLoad = False
                        End If
                    End If
                End With
            End With
    
            If bLoad Then
                If bDrive Then
                    ' do this to get past if the a drive or cd rom doesn't have a disk in it
                    If drv.IsReady Then
                        LoadFolders drv, Node, drv.RootFolder, tvFileTree
                    End If
                Else
'                    ' gotta get the folder to start with
'                    vTmp = Split(node.FullPath, "\")
                    ' ok the ubound element is our node so
                    sPath = GetTruePath(Node) 'node.FullPath
'                    sPath = Replace(sPath, vTmp(0) & "\", "")
'                    sPath = Replace(sPath, "\\", "\")
                    With fso
                        Set fld = .GetFolder(sPath)
                    End With
    
                    LoadFolders drv, Node, fld, tvFileTree
                End If
            End If
            
            ' check to see if the node now has real children, if it does then
            ' expand it
            If Node.Children > 0 Then
                If InStr(1, Node.Child.Key, "DUMMY") = 0 Then
                    ' expand the node
                    Node.Expanded = True
                End If
            End If
            
            ' clean up
            Set fso = Nothing
            Set drv = Nothing
            Set fld = Nothing
            
            
        End If
    End If
    
End Sub


