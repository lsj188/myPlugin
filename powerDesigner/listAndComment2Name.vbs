'******************************************************************************
'* File:     List Tables.vbs
'* Purpose:  This VB Script shows how to display properties of the first 5 tables
'*           defined in the current active PDM using message box.
'* Title:    Display tables properties in message box
'* Category: Display tables properties
'* Version:  1.0
'* Company:  Sybase Inc. 
'******************************************************************************

Option Explicit

'-----------------------------------------------------------------------------
' Main function
'-----------------------------------------------------------------------------

' Get the current active model
Dim model
Set model = ActiveModel
If (model Is Nothing) Or (Not model.IsKindOf(PdPDM.cls_Model)) Then
   MsgBox "The current model is not a PDM model."
Else
   ShowProperties model
End If


'-----------------------------------------------------------------------------
' Display tables properties defined in a folder
'-----------------------------------------------------------------------------
Sub ShowProperties(package)
   ' Get the Tables collection
   Dim ModelTables
   Set ModelTables = package.Tables
   MsgBox "The model or package '" + package.Name + "' contains " + CStr(ModelTables.Count) + " tables."

   ' For each table
   Dim noTable
   Dim tbl
   Dim bShortcutClosed
   Dim Desc
   noTable = 1
   For Each tbl In ModelTables
      If IsObject(tbl) Then
         bShortcutClosed = false
         If tbl.IsShortcut Then
            If Not (tbl.TargetObject Is Nothing) Then
               Set tbl = tbl.TargetObject
            Else
               bShortcutClosed = true
            End If
         End If

         Desc = "Table " + CStr(noTable) + ":"
         If Not bShortcutClosed Then
            Desc = Desc + vbCrLf + "ObjectID: "   + tbl.ObjectID
            Desc = Desc + vbCrLf + "Name: "       + tbl.Name
            Desc = Desc + "    " + "Code: "       + tbl.Code
            If IsObject(tbl.Parent) Then
               Desc = Desc + vbCrLf + "Parent: "  + tbl.Parent.Name
            Else
               Desc = Desc + vbCrLf + "Parent: "  + "<None>"
            End If
            Desc = Desc + vbCrLf + "DisplayName: " + tbl.DisplayName
            Desc = Desc + vbCrLf + "ObjectType: " + tbl.ObjectType
            Desc = Desc + vbCrLf + "CreationDate: " + CStr(tbl.CreationDate)
            Desc = Desc + vbCrLf + "Creator: "    + tbl.Creator
            Desc = Desc + vbCrLf + "ModificationDate: " + CStr(tbl.ModificationDate)
            Desc = Desc + vbCrLf + "Modifier: "   + tbl.Modifier
            Desc = Desc + vbCrLf + "Comment: "    + tbl.Comment
            Desc = Desc + vbCrLf + "Description: "  + Rtf2Ascii(tbl.Description)
            
            tbl.Name = tbl.Comment
			
			dim col
			For Each col In tbl.Columns
			    col.Name=col.Comment
				
				output "[ table: " + tbl.Name + " Comment: " + col.Comment + " => " + " Column Name: " + col.Name +  "]"
			Next
			
         Else
            Desc = Desc + vbCrLf + "The target object of the table shortcut "   + tbl.Code + " is not accessible."
         End If
         MsgBox Desc
      Else
         MsgBox "Not an object!"
      End If
      noTable = noTable + 1
      
      If noTable > 5 Then
         Exit For
      End If
   Next
   
   ' Display tables defined in subpackages
   Dim subpackage
   For Each subpackage in package.Packages
      If Not subpackage.IsShortcut Then
         ShowProperties subpackage
      End If
   Next
End Sub