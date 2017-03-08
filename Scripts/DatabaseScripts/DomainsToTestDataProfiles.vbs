'******************************************************************************
'* File:     List Columns.vbs
'* Purpose:  This VB Script shows how to display properties of the first 5 tables
'*           defined in the current active PDM using message box.
'* Title:    Display tables properties in message box
'* Category: Display tables columns
'* Version:  1.0
'* Company:  Randy Buchholz 
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
   Dim ModelColumns
   Set ModelTables = package.Tables
   MsgBox "The model or package '" + package.Name + "' contains " + CStr(ModelTables.Count) + " tables."

   ' For each table
   Dim noTable
   Dim tbl
   Dim bShortcutClosed
   Dim Desc
   Dim col
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
         Set ModelColumns = tbl.Columns
         Dim noCol
         noCol = 1
         Dim Keys
         Dim key
         Set Keys = tbl.Keys
         
         For Each key in Keys
            key.ConstraintName = Key.Name
            
         Next
         
         'output tbl.Keys
         For Each col in ModelColumns
            If IsObject(col) Then
            bShortcutClosed = false
            If col.IsShortcut Then
               If Not (col.TargetObject Is Nothing) Then
                  Set tbl = tbl.TargetObject
               Else
                  bShortcutClosed = true
               End If
            End If
            
            
            'output col.GetExtendedAttribute ("SqlServer.ExtDeftConstName")
            col.SetExtendedAttribute "SqlServer.ExtDeftConstName", "DF_" + Replace(col.Table.Code, "_", "") + "_" + Replace(col.Name, "_", "")
            col.CheckConstraintName = "CKC_" + Replace(col.Table.Code, "_", "") + "_" + Replace(col.Name, "_", "")          
            
         
         End If           
      Next
   End If   
   Next
   
   Dim ModelRefs
   Dim ref
   Set ModelRefs = package.References
   For Each ref in ModelRefs
      ref.ForeignKeyConstraintName = "FK_" + Replace(ref.ChildTable.Code, "_", "") + "_" + Replace(ref.ForeignKeyColumnList, "_", "")
      ref.Name = ref.ForeignKeyConstraintName
      ref.Code = ref.ForeignKeyConstraintName
   Next
    
End Sub
