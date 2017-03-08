'******************************************************************************
'* File:     TableLovToTestDataProfiles.vbs
'* Purpose:  Creates TestDataProfiles from all tables with Stereotype "LOV"
'*             where profile does not exist
'* Title:    Test Data from ValueLists
'* Category: Create Objects
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
   Dim table
   
   Dim ModelDataProfiles
   Dim newProfile
   
   Dim existingProfile
   
   Set ModelTables = model.Tables
   Set ModelDataProfiles = model.TestDataProfiles

   Dim testTypeDictionary
   Set testTypeDictionary = CreateObject("Scripting.Dictionary")
   For Each existingProfile in model.TestDataProfiles
      testTypeDictionary.Add UCase(existingProfile.Name), existingProfile.Name
   next

   For Each table in ModelTables
      If table.Stereotype = "LOV" then
         If testTypeDictionary.Exists("PICK_" + UCase(table.Code)) then
         else
            set newProfile = model.TestDataProfiles.createnew
            newProfile.Name = "pick_" + table.Code
            newProfile.Code = "pick_" + table.Code
            newProfile.ProfileClass = 1
            newProfile.CharacterCase = 2
            newProfile.ValuesSource = 2
            output "Created: " + newProfile.Name
         end if
      end if
   next    
End Sub
