Attribute VB_Name = "User_Macros"
'Option Explicit
Private PatPck As Integer
Private ShrSty As Integer
Private SpkPAL As Integer
Private SurNet As Integer
Private EduMat As Integer

Public Sub TemplateMacro(oRow As IBBQueryRow)
    Const ConsID = 1

    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
    
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init REApplication.SessionContext
        
        Dim oID As Long
        
        oID = oRow.Field(ConsID)
        oCons.Load oID
        
        
        oCons.Closedown
        Set oCons = Nothing
    End If

End Sub
Public Sub GetUserID()

   Dim reS As REServices
   Set reS = New REServices
   Dim oUI As IBBUtilityCode
   Set oUI = reS
   Dim lID As Long
   
   reS.Init REApplication.SessionContext

   'Get the user ID by passing the User Name
'   lID = oUI.GetUserID("19673_relointegrationuser")
    lID = oUI.GetUserID("lcoronado")
  
   MsgBox "UserID: " & lID
  
   reS.Closedown
   Set reS = Nothing
   Set oUI = Nothing

End Sub

Private Sub AddTableEntry(ByVal sClub As String)
'Private Sub AddTableEntry()
    Dim oTables As CCodeTables
    Set oTables = New CCodeTables
    oTables.Init REApplication.SessionContext
    
    Dim REService As REServices
    Set REService = New REServices
    REService.Init REApplication.SessionContext
    
    Dim ocodetableserver As CCodeTablesServer
    Set ocodetableserver = REService.CreateServiceObject(bbsoCodeTablesServer)
    ocodetableserver.Init REApplication.SessionContext
    
    Dim otablelookuphandler As CTableLookupHandler
    Set otablelookuphandler = REService.CreateServiceObject(bbsoTableLookupServer)
    otablelookuphandler.Init REApplication.SessionContext
    
    
    Dim oTable As CCodeTable
    Dim lID As Long
    'Loop through the tables to find the one you need
    For Each oTable In oTables
        If oTable.Fields(ctfNAME) = "TR team name" Then
            'Get the ID
            lID = oTable.Fields(ctfCODETABLEID)
            Exit For
        End If
    oTable.Closedown
    Next oTable
    
    oTables.Closedown
    Set oTables = Nothing
    
    Set oTable = Nothing
    
    'Load the collection of table entries for the requested table
    Dim oTEs As CTableEntries
    Set oTEs = New CTableEntries
    oTEs.Init REApplication.SessionContext, lID
    
    'Get a count of table entries
    Debug.Print oTEs.Count
    
    Dim oTE As CTableEntry
    
'    Dim sClub As String
'    sClub = "Ritz"
    
    If ocodetableserver.GetTableEntryID(sClub, lID) = 0 Then
    
'        omsgbox = MsgBox("Do you want to add '" & sClub & "' to the " & _
'        ocodetableserver.TABLENAME(lID) & " table?", vbYesNo)
'        Select Case omsgbox
'        Case vbYes
        otablelookuphandler.AddEntry True, lID, , sClub
'        Case vbNo
'        Debug.Print "Nothing"
'        End Select
    End If
    
    oTEs.Closedown
    Set oTEs = Nothing
    
    Set oTE = Nothing
    
    ocodetableserver.Closedown
    Set ocodetableserver = Nothing
    
    otablelookuphandler.Closedown
    Set otablelookuphandler = Nothing
    
    REService.Closedown
    Set REService = Nothing

End Sub
Public Sub OpeningPDF()
    Dim pa As String, pat1 As String, pat2 As String

'    pat1 = "C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"
'    pat2 = "C:\my documents\Batch 2320.pdf"
'    pa = pat1 & pat2

    pa = "C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe C:\Users\lcoronado\Desktop\re720exp_processdoc.pdf"
    
    Shell pa, vbNormalFocus

End Sub
'======================================================
'
'  End of the functions
'
'======================================================



Public Sub DeleteInvalidPrefAddressEmail(oRow As IBBQueryRow)
Const ConsID = 1

    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
    
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init REApplication.SessionContext
        
        Dim oID As Long
        
        oID = oRow.Field(ConsID)
        oCons.Load oID
        
        Dim oPhone As CConstitAddressPhone
        
        For Each oPhone In oCons.PreferredAddress.Phones
            If oPhone.Fields(CONSTIT_ADDRESS_PHONES_fld_PHONETYPE) = "E-mail" Then
                Debug.Print oCons.Fields(RECORDS_fld_FIRST_NAME) & " " & oPhone.Fields(CONSTIT_ADDRESS_PHONES_fld_PHONETYPE) & " " & oPhone.Fields(CONSTIT_ADDRESS_PHONES_fld_NUM)
                oCons.PreferredAddress.Phones.Remove oPhone
            End If
        Next oPhone
                
        On Error Resume Next
        oCons.Save
        
        oCons.Closedown
        Set oPhone = Nothing
        Set oCons = Nothing
    End If

End Sub

Public Sub CleanDuplicatePhoneType(oRow As IBBQueryRow)
    Const ConsID = 1

    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
    
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init REApplication.SessionContext
        
        Dim oID As Long
        
        Dim HomePn As Integer
        Dim BussPn As Integer
        Dim HomeFx As Integer
        Dim BussFx As Integer
        Dim email  As Integer
        Dim Web    As Integer
        
        HomePn = 0
        BussPn = 0
        HomeFx = 0
        BussFx = 0
        email = 1
        Web = 0
        
        oID = oRow.Field(ConsID)
        oCons.Load oID
        
        Dim oPhone As IBBPhone
        Dim oAddress As CConstitAddress
        
        For Each oAddress In oCons.Addresses
            For Each oPhone In oAddress.Phones
                If oAddress.Fields(CONSTIT_ADDRESS_fld_PREFERRED) = True Then
'            Debug.Print oPhone.Fields(Phone_fld_PhoneType) & " " & oPhone.Fields(Phone_fld_Num)
                If (Left(oPhone.Fields(Phone_fld_PhoneType), 6) = "E-mail") Then
                    Debug.Print "fixing... " & oCons.Fields(RECORDS_fld_CONSTITUENT_ID)
'                    If (Right(oPhone.Fields(Phone_fld_PhoneType), 1) <> Email) Then
                        Select Case email
                        Case 1
                            oPhone.Fields(Phone_fld_PhoneType) = "E-mail 1"
                        Case 2
                            oPhone.Fields(Phone_fld_PhoneType) = "E-mail 2"
                        Case 3
                            oPhone.Fields(Phone_fld_PhoneType) = "E-mail 3"
                        Case 4
                            oPhone.Fields(Phone_fld_PhoneType) = "E-mail 4"
                        End Select
'                    End If
                    email = email + 1
                End If
'                If (Left(oPhone.Fields(Phone_fld_PhoneType), 4) = "Work") Then
'                    If (Right(oPhone.Fields(Phone_fld_PhoneType), 1) <> BussPn) Then
'                        Select Case BussPn
'                        Case 1
'                            oPhone.Fields(Phone_fld_PhoneType) = "Work 1"
'                        Case 2
'                            oPhone.Fields(Phone_fld_PhoneType) = "Work 2"
'                        End Select
'                    End If
'                    BussPn = BussPn + 1
'                End If
'                If (Left(oPhone.Fields(Phone_fld_PhoneType), 4) = "Home") Then
'                    If (Right(oPhone.Fields(Phone_fld_PhoneType), 1) <> HomePn) Then
'                        Select Case HomePn
'                        Case 1
'                            oPhone.Fields(Phone_fld_PhoneType) = "Home 1"
'                        Case 2
'                            oPhone.Fields(Phone_fld_PhoneType) = "Home 2"
'                        End Select
'                    End If
'                    HomePn = HomePn + 1
'                End If
                End If
            Next oPhone
        Next oAddress
        
        On Error Resume Next
        oCons.Save
        
        oCons.Closedown
        Set oPhone = Nothing
        Set oAddress = Nothing
        Set oCons = Nothing
    End If

End Sub

Public Sub MoveTRTeamNames(oRow As IBBQueryRow)
    Const ConsID = 1

    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
'
'
'   Get the attribute ID
'
        Dim oService As REServices
        Set oService = New REServices
        oService.Init REApplication.SessionContext
           
        Dim oAttributeServer As CAttributeTypeServer
        Set oAttributeServer = oService.CreateServiceObject(bbsoAttributeTypeServer)
        oAttributeServer.Init REApplication.SessionContext
        
        Dim lAttributeID As Long
        Dim lAttributeID2 As Long
            
        lAttributeID = oAttributeServer.GetAttributeTypeID("TR team names", bbAttributeRecordType_PARTICIPANT)
        lAttributeID2 = oAttributeServer.GetAttributeTypeID("TR team name", bbAttributeRecordType_PARTICIPANT)
        
        Dim oAttribute As IBBAttribute
        Dim oPart As CParticipant
        Dim lPartID As Long
        
'
'   Load the constituent record
'
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init REApplication.SessionContext
        
        Dim oID As Long
        
        oID = oRow.Field(ConsID)
        oCons.Load oID
        
        Dim PartType As String
        PartType = ""
        
        For Each oPart In oCons.Participants
            If oPart.EventObject.Fields(SPECIAL_EVENT_fld_NAME) = oRow.Field("Event Name") Then
                
                Debug.Print oRow.Field("Event Name") & " ID: " & lAttributeID
                
                For Each oAttribute In oPart.Attributes
                    Debug.Print oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID)
                    If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttributeID Then
                        Debug.Print "Im here..." & oAttribute.Fields(Attribute_fld_VALUE)
                        PartType = oAttribute.Fields(Attribute_fld_VALUE)
                        oPart.Attributes.Remove oAttribute
                        Exit For
                    End If
                Next oAttribute
                
                Debug.Print "PartType Before: " & PartType
                PartType = Left(PartType, 59)
                Debug.Print "PartType After: " & PartType
                
                AddTableEntry (PartType)

                With oPart
                    Set oAttribute = oPart.Attributes.Add
                    oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttributeID2
                    oAttribute.Fields(Attribute_fld_VALUE) = PartType
                    
                    On Error Resume Next
                    oPart.Save
                End With
                
            End If
            oPart.Closedown
        Next oPart
'
'   Clean up
'
        Set oPart = Nothing
        
        oCons.Closedown
        oService.Closedown
        Set oService = Nothing
        Set oAttributeServer = Nothing
        Set oAttribute = Nothing
        Set oCons = Nothing
    End If

End Sub
Public Sub AddSpeaktoPCAction(oRow As IBBQueryRow)
Const ConsID = 1

    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
'
'
'   Get the attribute ID
'
        Dim oService As REServices
        Set oService = New REServices
        oService.Init REApplication.SessionContext
           
        Dim oAttributeServer As CAttributeTypeServer
        Set oAttributeServer = oService.CreateServiceObject(bbsoAttributeTypeServer)
        oAttributeServer.Init REApplication.SessionContext
        
        Dim lAttributeID As Long
        Dim lAttributeID2 As Long
            
        lAttributeID = oAttributeServer.GetAttributeTypeID("SpeakToPatientCentral", bbAttributeRecordType_CONSTITUENT)
        lAttributeID2 = oAttributeServer.GetAttributeTypeID("", bbAttributeRecordType_PARTICIPANT)
        
        Dim oAttribute As IBBAttribute
        Dim oPart As CParticipant
        Dim lPartID As Long
        
'
'   Load the constituent record
'
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init REApplication.SessionContext

        Dim oID As Long

        oID = oRow.Field(ConsID)
        oCons.Load oID
'
'   find the constituent attribute
'
        Dim NoteType As String
        NoteType = ""
        
        For Each oAttribute In oCons.Attributes
            Debug.Print oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) & " " & lAttributeID
            If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttributeID And _
                oAttribute.Fields(Attribute_fld_VALUE) = True Then
                If Trim(oAttribute.Fields(Attribute_fld_COMMENTS)) = "Patient" Then
                    NoteType = "Patient Packet Form"
                End If
                If Trim(oAttribute.Fields(Attribute_fld_COMMENTS)) = "JOML" Then
                    NoteType = "JOML"
                End If
                Exit For
            End If
        Next oAttribute
        If NoteType <> "" Then
'
'   Create the action and note
'
            Dim oAction As CAction
            Set oAction = New CAction
            oAction.Init REApplication.SessionContext
            
            Dim oAction2 As IBBAction2
            Set oAction2 = oAction
            
            With oAction
                oAction.Fields(ACTION_fld_CATEGORY) = "Email"
                oAction.Fields(ACTION_fld_DATE) = Date
                oAction.Fields(ACTION_fld_TYPE) = "PALS-SpeakToPatientCentral"
                oAction.Fields(ACTION_fld_RECORDS_ID) = oID
                
                oAction.Fields(ACTION_fld_AUTO_REMIND) = True
                oAction.Fields(ACTION_fld_NOTIFY_USING) = "Raiser's Edge reminders"
                oAction.Fields(ACTION_fld_PRIORITY) = "Normal"
                oAction.Fields(ACTION_fld_REMIND_FREQUENCY) = "day(s)"
                oAction.Fields(ACTION_fld_REMIND_VALUE) = 1
'
' Hard coded to add mgarcia - 74 as the user to be notified....
' Hard coded to add jbolduc - 379 as the user to be notified....
'
                .Remindees.Add.Fields(ActionRemindee_fld_USER_ID) = 74
            End With
            With oAction2.Notepads.Add
                .Fields(NOTEPAD_fld_NotepadType) = "Patient Central"
                .Fields(NOTEPAD_fld_Description) = NoteType
                .Fields(NOTEPAD_fld_NotepadDate) = CStr(Date)
                .Fields(NOTEPAD_fld_Author) = "jbolduc"
            End With
            oAction.Save
            
            With oAttribute
                oAttribute.Fields(Attribute_fld_VALUE) = False
            End With
            oCons.Save
    
            oAction.Closedown
            Set oAction = Nothing
        End If
        
        oCons.Closedown
        Set oService = Nothing
        Set oAttributeServer = Nothing
        Set oAttribute = Nothing
        Set oCons = Nothing
    End If

End Sub

Public Sub AddSpeaktoPCActionEvent(oRow As IBBQueryRow)
Const ConsID = 1

    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
'
'
'   Get the attribute ID
'
        Dim oService As REServices
        Set oService = New REServices
        oService.Init REApplication.SessionContext
           
        Dim oAttributeServer As CAttributeTypeServer
        Set oAttributeServer = oService.CreateServiceObject(bbsoAttributeTypeServer)
        oAttributeServer.Init REApplication.SessionContext
        
        Dim lAttributeID As Long
        Dim lAttributeID2 As Long
            
        lAttributeID = oAttributeServer.GetAttributeTypeID("SpeakToPatientCentral", bbAttributeRecordType_CONSTITUENT)
        lAttributeID2 = oAttributeServer.GetAttributeTypeID("Would you like to speak to a Patient Central Assoc", bbAttributeRecordType_PARTICIPANT)
        
        Dim oAttribute As IBBAttribute
        
'
'   Load the constituent record
'
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init REApplication.SessionContext

        Dim oID As Long

        oID = oRow.Field(ConsID)
        oCons.Load oID

        Dim NoteType As String
        NoteType = ""
        
        If oRow.Field("Event ID") <> "" Then
            NoteType = "TeamRisers-PS Registration"
        End If
        
        If NoteType <> "" Then
'
'   Create the action and note
'
            Dim oAction As CAction
            Set oAction = New CAction
            oAction.Init REApplication.SessionContext
            
            Dim oAction2 As IBBAction2
            Set oAction2 = oAction
            
            With oAction
                oAction.Fields(ACTION_fld_CATEGORY) = "Email"
                oAction.Fields(ACTION_fld_DATE) = Date
                oAction.Fields(ACTION_fld_TYPE) = "PALS-SpeakToPatientCentral"
                oAction.Fields(ACTION_fld_RECORDS_ID) = oID
                
                oAction.Fields(ACTION_fld_AUTO_REMIND) = True
                oAction.Fields(ACTION_fld_NOTIFY_USING) = "Raiser's Edge reminders"
                oAction.Fields(ACTION_fld_PRIORITY) = "Normal"
                oAction.Fields(ACTION_fld_REMIND_FREQUENCY) = "day(s)"
                oAction.Fields(ACTION_fld_REMIND_VALUE) = 1
'
' Hard coded to add mgarcia - 74 as the user to be notified....
' Hard coded to add jbolduc - 379 as the user to be notified...
'
                .Remindees.Add.Fields(ActionRemindee_fld_USER_ID) = 74
            End With
            With oAction2.Notepads.Add
                .Fields(NOTEPAD_fld_NotepadType) = "Patient Central"
                .Fields(NOTEPAD_fld_Description) = NoteType
                .Fields(NOTEPAD_fld_NotepadDate) = CStr(Date)
                .Fields(NOTEPAD_fld_Author) = "jbolduc"
            End With
            oAction.Save
            
            oAction.Closedown
            Set oAction = Nothing
'
            Dim oPart As CParticipant
            
            For Each oPart In oCons.Participants
                If oPart.EventObject.Fields(SPECIAL_EVENT_fld_EVENTID) = oRow.Field("Event ID") Then
                    For Each oAttribute In oPart.Attributes
                        If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttributeID2 And _
                            UCase(oAttribute.Fields(Attribute_fld_VALUE)) = "YES" Then
                            With oPart
                                oAttribute.Fields(Attribute_fld_VALUE) = "No"
                                oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE) = CStr(Date)
                            End With
                            oPart.Save
                        End If
                    Next oAttribute
                End If
                oPart.Closedown
            Next oPart
            
            Set oPart = Nothing
                        
        End If
        
        oCons.Closedown
        Set oService = Nothing
        Set oAttributeServer = Nothing
        Set oAttribute = Nothing
        Set oCons = Nothing
    End If

End Sub

Public Sub DeleteSpecificAppeal(oRow As IBBQueryRow)
Const ConsID = 1

    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
    
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init REApplication.SessionContext
        
        Dim oID As Long
        
        oID = oRow.Field(ConsID)
        oCons.Load oID
        
        Dim oAppeal As CConstitAppeal
        
        For Each oAppeal In oCons.Appeals
            Debug.Print oAppeal.Fields(CONSTITUENT_APPEALS_fld_Appeal)
            If oAppeal.Fields(CONSTITUENT_APPEALS_fld_Appeal) = "2016 Patient Survivor Housefile Engagement" And _
                oAppeal.Fields(CONSTITUENT_APPEALS_fld_Package) = "Download-Rollout Balance of Caregivers/Family" Then
                With oAppeal
                    Debug.Print "Deleting... "
                    oCons.Appeals.Remove oAppeal
                End With
                On Error Resume Next
                oCons.Save
            End If
        Next oAppeal
        
        oCons.Closedown
        Set oCons = Nothing
    End If

End Sub
Public Sub AddShareStoryAction(oRow As IBBQueryRow)
Const ConsID = 1

    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
'
'
'   Get the attribute ID
'
        Dim oService As REServices
        Set oService = New REServices
        oService.Init REApplication.SessionContext
           
        Dim oAttributeServer As CAttributeTypeServer
        Set oAttributeServer = oService.CreateServiceObject(bbsoAttributeTypeServer)
        oAttributeServer.Init REApplication.SessionContext
        
        Dim lAttributeID As Long
        Dim lAttributeID2 As Long
            
        lAttributeID = oAttributeServer.GetAttributeTypeID("ShareYourStoryPatientCentral", bbAttributeRecordType_CONSTITUENT)
        lAttributeID2 = oAttributeServer.GetAttributeTypeID("", bbAttributeRecordType_PARTICIPANT)
        
        Dim oAttribute As IBBAttribute
        Dim oPart As CParticipant
        Dim lPartID As Long
        
'
'   Load the constituent record
'
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init REApplication.SessionContext

        Dim oID As Long

        oID = oRow.Field(ConsID)
        oCons.Load oID
'
'   find the constituent attribute
'
        Dim NoteType As String
        NoteType = ""
        
        For Each oAttribute In oCons.Attributes
            Debug.Print oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) & " " & lAttributeID
            If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttributeID And _
                oAttribute.Fields(Attribute_fld_VALUE) = True Then
                If Trim(oAttribute.Fields(Attribute_fld_COMMENTS)) = "Patient" Then
                    NoteType = "Patient Packet Form"
                End If
                If Trim(oAttribute.Fields(Attribute_fld_COMMENTS)) = "JOML" Then
                    NoteType = "JOML"
                End If
                Exit For
            End If
        Next oAttribute
        If NoteType <> "" Then
'
'   Create the action and note
'
            Dim oAction As CAction
            Set oAction = New CAction
            oAction.Init REApplication.SessionContext
            
            Dim oAction2 As IBBAction2
            Set oAction2 = oAction
            
            With oAction
                oAction.Fields(ACTION_fld_CATEGORY) = "Email"
                oAction.Fields(ACTION_fld_DATE) = Date
                oAction.Fields(ACTION_fld_TYPE) = "Share Your Story Interest"
                oAction.Fields(ACTION_fld_RECORDS_ID) = oID
                
                oAction.Fields(ACTION_fld_AUTO_REMIND) = True
                oAction.Fields(ACTION_fld_NOTIFY_USING) = "Raiser's Edge reminders"
                oAction.Fields(ACTION_fld_PRIORITY) = "Normal"
                oAction.Fields(ACTION_fld_REMIND_FREQUENCY) = "day(s)"
                oAction.Fields(ACTION_fld_REMIND_VALUE) = 1
'
' Hard coded to add Anica Lamkin as the user to be notified....
'
                .Remindees.Add.Fields(ActionRemindee_fld_USER_ID) = 211
            End With
'            With oAction2.Notepads.Add
'                .Fields(NOTEPAD_fld_NotepadType) = "Patient Central"
'                .Fields(NOTEPAD_fld_Description) = NoteType
'                .Fields(NOTEPAD_fld_NotepadDate) = CStr(Date)
'                .Fields(NOTEPAD_fld_Author) = "alamkin"
'            End With
            On Error Resume Next
            oAction.Save
            
            With oAttribute
                oAttribute.Fields(Attribute_fld_VALUE) = False
            End With
            On Error Resume Next
            oCons.Save
    
            oAction.Closedown
            Set oAction = Nothing
        End If
        
        oCons.Closedown
        Set oService = Nothing
        Set oAttributeServer = Nothing
        Set oAttribute = Nothing
        Set oCons = Nothing
    End If

End Sub
Public Sub GetTargetAnalyticScores(oRow As IBBExportRow)
    Const ConsID = 1
' Use to Extract Any List
'    Const TAScore1 = 13
'    Const TAScore2 = 14
'    Const TAScore3 = 15
'    Const TAScore4 = 16
'    Const WPScore1 = 17
    
    
    Dim BTAScore1 As String
    Dim BTAScore2 As String
    Dim BTAScore3 As String
    Dim BTAScore4 As String
    Dim BTAScore5 As String
    
    BTAScore1 = ""
    BTAScore2 = ""
    BTAScore3 = ""
    BTAScore4 = ""
    BTAScore5 = ""

'
' Use to Extract Direct Mail
'
    Const TAScore1 = 8
    Const TAScore2 = 9
    Const TAScore3 = 10
    Const TAScore4 = 11
    Const WPScore1 = 12

    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
        
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init REApplication.SessionContext
        
        Dim oID As Long
        
        oID = oRow.Field(ConsID)
        oCons.Load oID
        
        Debug.Print oCons.Fields(RECORDS_fld_FIRST_NAME) & " " & oCons.Fields(RECORDS_fld_LAST_NAME) & " " & oCons.Fields(RECORDS_fld_CONSTITUENT_ID)
        
        Dim oPros As CProspect
        Set oPros = oCons.Prospect
        
        Dim oRating As CRating
        
        For Each oRating In oPros.Ratings
            With oRating
                If oRating.Fields(RATING_fld_SOURCE) = "Blackbaud Analytics' Custom Modeling Service" Then
                    If oRating.Fields(RATING_fld_CATEGORY) = "CMS Target Gift Dollar Range" Then
                        Debug.Print oRating.Fields(RATING_fld_DESCRIPTION)
'                        oRow.Field(TAScore1) = oRating.Fields(RATING_fld_DESCRIPTION)
                        BTAScore1 = oRating.Fields(RATING_fld_DESCRIPTION)
                    End If
                    If oRating.Fields(RATING_fld_CATEGORY) = "CMS Major Gift Likelihood" Then
'                        oRow.Field(TAScore2) = oRating.Fields(RATING_fld_DESCRIPTION)
                        BTAScore2 = oRating.Fields(RATING_fld_DESCRIPTION)
                    End If
                    If oRating.Fields(RATING_fld_CATEGORY) = "CMS Planned Gift Likelihood" Then
'                        oRow.Field(TAScore3) = oRating.Fields(RATING_fld_DESCRIPTION)
                        BTAScore3 = oRating.Fields(RATING_fld_DESCRIPTION)
                    End If
                End If
                If oRating.Fields(RATING_fld_SOURCE) = "Target Analytics Custom Modeling Service" Then
                    If oRating.Fields(RATING_fld_CATEGORY) = "CMS Principal Giving Solution" Then
                        If Trim(oRating.Fields(RATING_fld_DESCRIPTION)) <> "" Then
'                            oRow.Field(TAScore4) = oRating.Fields(RATING_fld_DESCRIPTION)
                            BTAScore4 = oRating.Fields(RATING_fld_DESCRIPTION)
                        End If
                    End If
                    If oRating.Fields(RATING_fld_CATEGORY) = "CMS WealthPoint Rating" Then
'                        oRow.Field(WPScore1) = oRating.Fields(RATING_fld_DESCRIPTION)
                        BTAScore5 = oRating.Fields(RATING_fld_DESCRIPTION)
                    End If
                End If
'                If oRating.Fields(RATING_fld_SOURCE) = "Target Analytics Custom Modeling Service" Then
'                    If oRating.Fields(RATING_fld_CATEGORY) = "CMS WealthPoint Rating" Then
'                        oRow.Field(WPScore1) = oRating.Fields(RATING_fld_DESCRIPTION)
'                    End If
'                End If
                Set oRating = Nothing
            End With
        Next oRating
        
        Set oPros = Nothing
        
        oRow.Field(TAScore1) = BTAScore1
        oRow.Field(TAScore2) = BTAScore2
        oRow.Field(TAScore3) = BTAScore3
        oRow.Field(TAScore4) = BTAScore4
        oRow.Field(WPScore1) = BTAScore5
        
        oCons.Closedown
        Set oCons = Nothing
    End If

End Sub

Public Sub AddTargetAnalyticScores(oRow As IBBQueryRow)
    Const ConsID = 1
    Const TAScore1 = 13
    Const TAScore2 = 14
    Const TAScore3 = 15
    Const TAScore4 = 16
    Const WPScore1 = 17

'    Const TAScore1 = 8
'    Const TAScore2 = 9
'    Const TAScore3 = 10
'    Const TAScore4 = 11

    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
        
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init REApplication.SessionContext
        
        Dim oID As Long
        
        oID = oRow.Field(ConsID)
        oCons.Load oID
        
        Debug.Print oCons.Fields(RECORDS_fld_FIRST_NAME) & " " & oCons.Fields(RECORDS_fld_LAST_NAME)
        
        Dim oPros As CProspect
        Set oPros = oCons.Prospect
        
        Dim oRating As CRating
        
        For Each oRating In oPros.Ratings
            With oRating
                If oRating.Fields(RATING_fld_SOURCE) = "Blackbaud Analytics' Custom Modeling Service" Then
                    If oRating.Fields(RATING_fld_CATEGORY) = "CMS Target Gift Dollar Range" Then
                        Debug.Print oRating.Fields(RATING_fld_DESCRIPTION)
                        oRow.Field("CMS Target Gift Dollar Range") = oRating.Fields(RATING_fld_DESCRIPTION)
                    End If
                    If oRating.Fields(RATING_fld_CATEGORY) = "CMS Major Gift Likelihood" Then
                        oRow.Field("CMS Major Gift Likelihood") = oRating.Fields(RATING_fld_DESCRIPTION)
                    End If
                    If oRating.Fields(RATING_fld_CATEGORY) = "CMS Planned Gift Likelihood" Then
                        oRow.Field("CMS Planned Gift Likelihood") = oRating.Fields(RATING_fld_DESCRIPTION)
                    End If
                End If
                If oRating.Fields(RATING_fld_SOURCE) = "Target Analytics Custom Modeling Service" Then
                    If oRating.Fields(RATING_fld_CATEGORY) = "CMS Principal Giving Solution" Then
                        oRow.Field("CMS Principal Giving Solution") = oRating.Fields(RATING_fld_DESCRIPTION)
                    End If
                End If
                If oRating.Fields(RATING_fld_SOURCE) = "Target Analytics Custom Modeling Service" Then
                    If oRating.Fields(RATING_fld_CATEGORY) = "CMS WealthPoint Rating" Then
                        oRow.Field(WPScore1) = oRating.Fields(RATING_fld_DESCRIPTION)
                    End If
                End If
                Set oRating = Nothing
            End With
        Next oRating
        
        Set oPros = Nothing
        
        oCons.Closedown
        Set oCons = Nothing
    End If

End Sub


Public Sub CountEventRegistrations(oRow As IBBExportRow)
    Const ConsID = 1
    Const AD = 6
    Const PS = 7
    Const OT = 8

    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
    
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init REApplication.SessionContext
        
        Dim oID As Long
        
        oID = oRow.Field(ConsID)
        oCons.Load oID
        
        Dim AdvocacyDay As Integer
        Dim PurpleStride As Integer
        Dim OtherEvent As Integer
        
        AdvocacyDay = 0
        PurpleStride = 0
        OtherEvent = 0
        
        Dim oPart As CParticipant
              
        Debug.Print oCons.Fields(RECORDS_fld_ID)
        For Each oPart In oCons.Participants
            With oPart
                Debug.Print oPart.EventObject.Fields(SPECIAL_EVENT_fld_TYPEID) & " " & _
                    oPart.EventObject.Fields(SPECIAL_EVENT_fld_EVENTID)
                If oPart.EventObject.Fields(SPECIAL_EVENT_fld_TYPEID) = "Advocacy Day" Then
                    AdvocacyDay = AdvocacyDay + 1
                End If
                If oPart.EventObject.Fields(SPECIAL_EVENT_fld_TYPEID) = "PurpleStride" Then
                    PurpleStride = PurpleStride + 1
                End If
                If oPart.EventObject.Fields(SPECIAL_EVENT_fld_TYPEID) <> "Advocacy Day" And _
                    oPart.EventObject.Fields(SPECIAL_EVENT_fld_TYPEID) <> "PurpleStride" Then
                    OtherEvent = OtherEvent + 1
                End If
            End With
        Next oPart
        
        Debug.Print AdvocacyDay
        Debug.Print PurpleStride
        Debug.Print OtherEvent
        
        oRow.Field(AD) = AdvocacyDay
        oRow.Field(PS) = PurpleStride
        oRow.Field(OT) = OtherEvent
        
        oCons.Closedown
        Set oCons = Nothing
        Set oPart = Nothing
    End If

End Sub
Public Sub FindIMOLastGiftDate(oRow As IBBQueryRow)
    Const ConsID = 1

    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
    
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init REApplication.SessionContext
        
        Dim oID As Long
        
        oID = oRow.Field(ConsID)
        oCons.Load oID
        
        Dim oTribute As CTribute
        Dim oGift As CGift
        Dim LastGift As Date
        LastGift = "01/01/1900"
        
        For Each oTribute In oCons.Tributes
'            Debug.Print oTribute.Fields(Tribute_fld_TRIBUTE_TYPE)
            If oTribute.Fields(Tribute_fld_TRIBUTE_TYPE) = "In Memory of" Then
                For Each oGift In oTribute.Gifts
                    If oGift.Fields(GIFT_fld_Date) > LastGift Then
                        LastGift = oGift.Fields(GIFT_fld_Date)
                    End If
                Next oGift
            End If
        Next oTribute
        
        oRow.Field("IMO Last Gift Date") = LastGift
        
        Set oGift = Nothing
        Set oTribute = Nothing
        
        oCons.Closedown
        Set oCons = Nothing
    End If
    
End Sub

Public Sub FindPiPMembership(oRow As IBBQueryRow)
    Const ConsID = 1

    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
    
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init REApplication.SessionContext
'
'
'   Get the attribute ID
'
        Dim oService As REServices
        Set oService = New REServices
        oService.Init REApplication.SessionContext
           
        Dim oAttributeServer As CAttributeTypeServer
        Set oAttributeServer = oService.CreateServiceObject(bbsoAttributeTypeServer)
        oAttributeServer.Init REApplication.SessionContext
        
        Dim lAttID1 As Long
        Dim lAttID2 As Long
        Dim lAttID3 As Long
        Dim lAttID4 As Long
        Dim lAttID5 As Long
        Dim lAttID6 As Long
        Dim lAttID7 As Long
        Dim lAttID8 As Long
        Dim lAttID9 As Long
        Dim lAttID10 As Long
            
        lAttID1 = oAttributeServer.GetAttributeTypeID("CORP Partners in Progress - $1,000-2,499", bbAttributeRecordType_CONSTITUENT)
        lAttID2 = oAttributeServer.GetAttributeTypeID("CORP Partners in Progress - $2,500-4,999", bbAttributeRecordType_CONSTITUENT)
        lAttID3 = oAttributeServer.GetAttributeTypeID("CORP Partners in Progress - $5,000-9,999", bbAttributeRecordType_CONSTITUENT)
        lAttID4 = oAttributeServer.GetAttributeTypeID("CORP Partners in Progress - $10,000-24,999", bbAttributeRecordType_CONSTITUENT)
        lAttID5 = oAttributeServer.GetAttributeTypeID("CORP Partners in Progress - $25,000+", bbAttributeRecordType_CONSTITUENT)
        lAttID6 = oAttributeServer.GetAttributeTypeID("Partners in Progress - $1,000-2,499", bbAttributeRecordType_CONSTITUENT)
        lAttID7 = oAttributeServer.GetAttributeTypeID("Partners in Progress - $2,500-4,999", bbAttributeRecordType_CONSTITUENT)
        lAttID8 = oAttributeServer.GetAttributeTypeID("Partners in Progress - $5,000-9,999", bbAttributeRecordType_CONSTITUENT)
        lAttID9 = oAttributeServer.GetAttributeTypeID("Partners in Progress - $10,000-24,999", bbAttributeRecordType_CONSTITUENT)
        lAttID10 = oAttributeServer.GetAttributeTypeID("Partners in Progress - $25,000+", bbAttributeRecordType_CONSTITUENT)

        Dim oAttribute As IBBAttribute
Debug.Print lAttID10
'
'
'   Load the constituent record
'
        
        Dim oID As Long
        
        oID = oRow.Field(ConsID)
        oCons.Load oID
'
        Dim NoteType As String
        NoteType = ""
        
        For Each oAttribute In oCons.Attributes
        
            Debug.Print oAttribute.Fields(Attribute_fld_ATTRIBUTES_ID)
     
            If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttID1 Then
                If NoteType = "" Then
                    NoteType = "CORP Partners in Progress - $1,000-2,499"
                Else
                    NoteType = NoteType & ", CORP Partners in Progress - $1,000-2,499"
                End If
            End If
            If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttID2 Then
                If NoteType = "" Then
                    NoteType = "CORP Partners in Progress - $2,500-4,999"
                Else
                    NoteType = NoteType & ", CORP Partners in Progress - $2,500-4,999"
                End If
            End If
            If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttID3 Then
                If NoteType = "" Then
                    NoteType = "CORP Partners in Progress - $5,000-9,999"
                Else
                    NoteType = NoteType & ", CORP Partners in Progress - $5,000-9,999"
                End If
            End If
            If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttID4 Then
                If NoteType = "" Then
                    NoteType = "CORP Partners in Progress - $10,000-24,999"
                Else
                    NoteType = NoteType & ", CORP Partners in Progress - $10,000-24,999"
                End If
            End If
            If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttID5 Then
            Debug.Print oAttribute.Fields(Attribute_fld_ATTRIBUTES_ID)
                If NoteType = "" Then
                    NoteType = "CORP Partners in Progress - $25,000+"
                Else
                    NoteType = NoteType & ", CORP Partners in Progress - $25,000+"
                End If
            End If
            If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttID6 Then
                If NoteType = "" Then
                    NoteType = "Partners in Progress - $1,000-2,499"
                Else
                    NoteType = NoteType & ", Partners in Progress - $1,000-2,499"
                End If
            End If
            If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttID7 Then
                If NoteType = "" Then
                    NoteType = "Partners in Progress - $2,500-4,999"
                Else
                    NoteType = NoteType & ", Partners in Progress - $2,500-4,999"
                End If
            End If
            If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttID8 Then
                If NoteType = "" Then
                    NoteType = "Partners in Progress - $5,000-9,999"
                Else
                    NoteType = NoteType & ", Partners in Progress - $5,000-9,999"
                End If
            End If
            If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttID9 Then
                If NoteType = "" Then
                    NoteType = "Partners in Progress - $10,000-24,999"
                Else
                    NoteType = NoteType & ", Partners in Progress - $10,000-24,999"
                End If
            End If
            If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttID10 Then
                If NoteType = "" Then
                    NoteType = "Partners in Progress - $25,000+"
                Else
                    NoteType = NoteType & ", CPartners in Progress - $25,000+"
                End If
            End If
        Next oAttribute
        
        oRow.Field("PiP Membership") = NoteType
'
        oCons.Closedown
        Set oService = Nothing
        Set oAttributeServer = Nothing
        Set oAttribute = Nothing
        Set oCons = Nothing
    
    End If

End Sub
Public Sub MakeGiftAdjustment(oRow As IBBQueryRow)
    Const Gift_ID = 1
'    Const AdjReason = "A-Fund Adj tool (152905 to 222335)- report to finance"
'    Const AdjReason = "A - Fund adjustments (part of Streisand agreement) from 146477 to 217884"
    Const AdjReason = "Global adjustment appeal from NVLVSpring2017 to NVLVSpring2018."
    
    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
        '
        ' load the gift record
        '
        Dim oGift As CGift
        Set oGift = New CGift
        oGift.Init RE7.SessionContext
        
        Dim lID As Long
        lID = oRow.Field(Gift_ID)
        
        oGift.Load lID

        If oGift.Fields(GIFT_fld_Post_Status) = "Posted" Then
            Debug.Print oGift.Fields(GIFT_fld_Amount)
'
' Search for any unposted adjustments and post it
'
            Dim oAdj As IBBAdjustment
            Dim oGLd As CGiftGLDistribution
            Dim Adj_Counter As Integer
            Dim Adj_Date As Date
            Adj_Date = oGift.Fields(GIFT_fld_Date)
            
'            For Each oAdj In oGift.Adjustments
'                If oAdj.Fields(ADJUSTMENT_fld_Post_Status) = "Not Posted" Then
'                    If oAdj.Fields(ADJUSTMENT_fld_Date) > Adj_Date Then
'                        Adj_Date = oAdj.Fields(ADJUSTMENT_fld_Date)
'                    End If
'                End If
'            Next oAdj
'
'    '        Debug.Print Adj_Date & " "; Adj_Counter
'
'            Adj_Counter = 0
'
'            For Each oAdj In oGift.Adjustments
'                Adj_Counter = Adj_Counter + 1
'                If oAdj.Fields(ADJUSTMENT_fld_Date) = Adj_Date And oAdj.Fields(ADJUSTMENT_fld_Post_Status) = "Not Posted" Then
'                    Dim oAdjs As CAdjustmentServer
'                    Set oAdjs = New CAdjustmentServer
'                    oAdjs.Init REApplication.SessionContext, oGift
'    '                Debug.Print oAdj.Fields(ADJUSTMENT_fld_Date) & " " & oAdj.Fields(ADJUSTMENT_fld_Post_Status) & " " & Adj_Date & " " & Adj_Counter
'                    With oAdjs
'                        Set oAdj = .EditAdjustment(oGift.Adjustments(Adj_Counter))
'                        With oAdj
'                            .Fields(ADJUSTMENT_fld_Post_Status) = "Posted"
'                        End With
'                        .Save
'                        .Closedown
'                    End With
'                    Set oAdjs = Nothing
'                End If
'            Next oAdj
'            Set oAdj = Nothing
'
'
            Dim oAdjustment As IBBAdjustment
            Dim oAdjustmentServer As CAdjustmentServer

            Set oAdjustmentServer = New CAdjustmentServer

            With oAdjustmentServer
                .Init REApplication.SessionContext, oGift
'                On Error Resume Next
                Set oAdjustment = .AddAdjustment()
                With oAdjustment
                Debug.Print "creating adjustment on " & oGift.Fields(GIFT_fld_ID) & " " & oGift.Fields(GIFT_fld_Amount) & " " & oGift.Fields(GIFT_fld_Date)
                    .Fields(ADJUSTMENT_fld_Date) = Date
                    .Fields(ADJUSTMENT_fld_Amount) = oGift.Fields(GIFT_fld_Amount)
                    .Fields(ADJUSTMENT_fld_Reason) = AdjReason
                    .Fields(ADJUSTMENT_fld_Fund) = oGift.Fields(GIFT_fld_Fund)
                    .Fields(ADJUSTMENT_fld_Campaign) = oGift.Fields(GIFT_fld_Campaign)
                    .Fields(ADJUSTMENT_fld_Appeal) = "NVLVSpring2018"
                    .Fields(ADJUSTMENT_fld_Package) = oGift.Fields(GIFT_fld_Package)
                End With
                'Validate and save the adjustment
'                On Error Resume Next
                .Validate
                oAdjustmentServer.Save
            End With

            Set oAdjustmentServer = Nothing
'            On Error Resume Next
            oGift.Save

        End If

        If oGift.Fields(GIFT_fld_Post_Status) <> "Posted" Then
            With oGift
'                .Fields(GIFT_fld_Campaign) = "ANN"
                .Fields(GIFT_fld_Appeal) = "NVLVSpring2018"
'                .Fields(GIFT_fld_Package) = ""
'                On Error Resume Next
                oGift.Save
            End With
            With oGift.Notepads.Add
                .Fields(NOTEPAD_fld_Author) = "lcoronado"
                .Fields(NOTEPAD_fld_Description) = "Appeal adjustment"
                .Fields(NOTEPAD_fld_NotepadDate) = Date
                .Fields(NOTEPAD_fld_NotepadType) = "Adjustment History"
                .Fields(NOTEPAD_fld_ActualNotes) = AdjReason
                .Fields(NOTEPAD_fld_Notes) = .Fields(NOTEPAD_fld_ActualNotes)
'                On Error Resume Next
                oGift.Save
            End With
        End If
        
        oGift.Closedown
        Set oGift = Nothing

    End If
End Sub
Public Sub RemoveConstAttribute(oRow As IBBQueryRow)
    Const ConsID = 1

    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
'
'
'   Get the attribute ID
'
        Dim oService As REServices
        Set oService = New REServices
        oService.Init REApplication.SessionContext
           
        Dim oAttributeServer As CAttributeTypeServer
        Set oAttributeServer = oService.CreateServiceObject(bbsoAttributeTypeServer)
        oAttributeServer.Init REApplication.SessionContext
        
        Dim lAttributeID As Long
        Dim lAttributeID2 As Long
            
        lAttributeID = oAttributeServer.GetAttributeTypeID("Source/General", bbAttributeRecordType_CONSTITUENT)
        lAttributeID2 = oAttributeServer.GetAttributeTypeID("Blah", bbAttributeRecordType_CONSTITUENT)
        
        Dim oAttribute As IBBAttribute
        Dim oPart As CParticipant
        Dim lPartID As Long
        
'
'   Load the constituent record
'
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init REApplication.SessionContext
        
        Dim oID As Long
        
        oID = oRow.Field(ConsID)
        oCons.Load oID
        
        For Each oAttribute In oCons.Attributes
            If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttributeID And _
                oAttribute.Fields(Attribute_fld_VALUE) = "Perthera Contact" Then
                Debug.Print oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) & " " & lAttributeID
'                oCons.Attributes.Remove oAttribute
'                oCons.Save
            End If
        Next oAttribute
                
'
'   Clean up
'
        Set oPart = Nothing
        
        oCons.Closedown
        oService.Closedown
        Set oService = Nothing
        Set oAttributeServer = Nothing
        Set oAttribute = Nothing
        Set oCons = Nothing
    End If

End Sub
Public Sub CreateFlagStripAction(oRow As IBBQueryRow)
    Const ConsID = 1

    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
    
        Dim oID As Long
        oID = oRow.Field(ConsID)
        
        Dim ActionDate As Date
        ActionDate = Date
        If oRow.Field("Gift Date") <> "" Then
            ActionDate = oRow.Field("Gift Date")
        End If
        
        
        Dim oAction As CAction

        Set oAction = New CAction
        oAction.Init REApplication.SessionContext
        
        With oAction
            .Fields(ACTION_fld_CATEGORY) = "Advocacy"
            .Fields(ACTION_fld_TYPE) = "Advocacy"
            .Fields(ACTION_fld_DELIVERY_METHOD) = "MAIL"
            .Fields(ACTION_fld_DATE) = ActionDate
'            .Fields(ACTION_fld_DATE) = oRow.Field("Gift Date")
'            .Fields(ACTION_fld_DATE) = "05/27/2016"
            .Fields(ACTION_fld_COMPLETED) = True
            .Fields(ACTION_fld_COMPLETED_DATE) = ActionDate
'            .Fields(ACTION_fld_COMPLETED_DATE) = oRow.Field("Gift Date")
'            .Fields(ACTION_fld_COMPLETED_DATE) = "05/27/2016"
            .Fields(ACTION_fld_ALERT_TITLE) = "AD 2016 Alert Direct Response"
            .Fields(ACTION_fld_PRIORITY) = "Normal"
            'system record ID of the constituent
            .Fields(ACTION_fld_RECORDS_ID) = oID
            .Save
        End With
                
        oAction.Closedown
        Set oAction = Nothing

    End If

End Sub
Public Sub Query_LastNextActions(oRow As IBBQueryRow)
    
    Dim Const_ID As Long
    Dim Solicitor_ID As Long
    
    Const_ID = 1
    Solicitor_ID = 2
    
    If oRow.BOF Then
        Rem MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        Rem MsgBox "End processing"
    Else
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init RE7.SessionContext
        
        Dim LastActionDate As Date
        Dim LastActionType As String
        Dim LastActionNote As String
        Dim LastActionCatg As String
        Dim NextActionDate As Date
        Dim NextActionType As String
        Dim NextActionNote As String
        Dim NextActionCatg As String
        Dim NoNextAction As Boolean
        
        LastActionDate = "1/1/1900"
        LastActionType = ""
        LastActionNote = ""
        LastActionCatg = ""
        NextActionDate = "1/1/9999"
        NextActionType = ""
        NextActionNote = ""
        NextActionCatg = ""
        
        
        Dim lID As Long
        'get the id
        lID = oRow.Field(Const_ID)
        oCons.Load lID
        
        Dim sID As Long
        sID = oRow.Field(Solicitor_ID)
        
        Debug.Print oCons.Fields(RECORDS_fld_PRIMARY_ADDRESSEE)
        
        Dim oActions As CAction
        Dim oSolicitor As CActionSolicitor
        
        NoNextAction = True
        
        Debug.Print oCons.Actions.Count
        
        If oCons.Actions.Count > 0 Then
            For Each oActions In oCons.Actions
                For Each oSolicitor In oActions.Solicitors
                    If oSolicitor.Fields(ActionSolicitor_fld_RECORDS_ID) = sID Then
                        If oActions.Fields(ACTION_fld_COMPLETED) = True And _
                            oActions.Fields(ACTION_fld_DATE) > LastActionDate Then
                            LastActionDate = oActions.Fields(ACTION_fld_DATE)
                            LastActionType = oActions.Fields(ACTION_fld_TYPE)
                            LastActionNote = oActions.Fields(ACTION_fld_NOTES)
                            LastActionCatg = oActions.Fields(ACTION_fld_CATEGORY)
                            LastActionDate = oActions.Fields(ACTION_fld_DATE)
                        End If
                        If oActions.Fields(ACTION_fld_COMPLETED) = False And _
                            oActions.Fields(ACTION_fld_DATE) < NextActionDate Then
                            NextActionDate = oActions.Fields(ACTION_fld_DATE)
                            NextActionType = oActions.Fields(ACTION_fld_TYPE)
                            NextActionNote = oActions.Fields(ACTION_fld_NOTES)
                            NextActionCatg = oActions.Fields(ACTION_fld_CATEGORY)
                            NextActionDate = oActions.Fields(ACTION_fld_DATE)
                        End If
                    End If
                Next
            Next
        End If
'
        Set oActions = Nothing

        If LastActionDate = "1/1/1900" Then oRow.Field("Last Action Date") = ""
        If LastActionDate <> "1/1/1900" Then oRow.Field("Last Action Date") = LastActionDate
        oRow.Field("Last Action Type") = LastActionType
        oRow.Field("Last Action Category") = LastActionCatg
        oRow.Field("Last Action Note") = LastActionNote
        If NextActionDate = "1/1/9999" Then oRow.Field("Next Action Date") = ""
        If NextActionDate <> "1/1/9999" Then oRow.Field("Next Action Date") = NextActionDate
        oRow.Field("Next Action Type") = NextActionType
        oRow.Field("Next Action Category") = NextActionCatg
        oRow.Field("Next Action Note") = NextActionNote
        
        
        oCons.Closedown
        Set oCons = Nothing
    End If
End Sub

Public Sub AddPatientPacket(oRow As IBBQueryRow)
Const ConsID = 1

    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
'
'
'   Get the attribute ID
'
        Dim oService As REServices
        Set oService = New REServices
        oService.Init REApplication.SessionContext
           
        Dim oAttributeServer As CAttributeTypeServer
        Set oAttributeServer = oService.CreateServiceObject(bbsoAttributeTypeServer)
        oAttributeServer.Init REApplication.SessionContext
        
        Dim lAttributeID As Long
        Dim lAttributeID2 As Long
            
        lAttributeID = oAttributeServer.GetAttributeTypeID("PALS Patient Packet-Online", bbAttributeRecordType_CONSTITUENT)
        lAttributeID2 = oAttributeServer.GetAttributeTypeID("PALS-Patient Packet", bbAttributeRecordType_ACTION)
        
        Dim oAttribute  As IBBAttribute
        Dim oAttribute2 As IBBAttribute
        Dim oPart As CParticipant
        Dim lPartID As Long
        
        Dim ActionDate As Date
        ActionDate = Date
'
'   Load the constituent record
'
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init REApplication.SessionContext

        Dim oID As Long

        oID = oRow.Field(ConsID)
        oCons.Load oID
'
'   find the constituent attribute
'
        Dim NoteType As String
        NoteType = ""

        For Each oAttribute In oCons.Attributes
'            Debug.Print oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) & " " & lAttributeID
            If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttributeID And _
                oAttribute.Fields(Attribute_fld_VALUE) = True Then
                If Trim(oAttribute.Fields(Attribute_fld_COMMENTS)) = "Patient" Then
                    NoteType = "Patient Packet Form"
                End If
                If Trim(oAttribute.Fields(Attribute_fld_COMMENTS)) = "JOML" Then
                    NoteType = "JOML"
                End If
                If Trim(oAttribute.Fields(Attribute_fld_COMMENTS)) <> "" Then
                    NoteType = Trim(oAttribute.Fields(Attribute_fld_COMMENTS))
                End If
                ActionDate = oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE)
                Exit For
            End If
        Next oAttribute
        If NoteType <> "" Then
'
'   Create the action and note
'
            Dim oAction As CAction
            Set oAction = New CAction
            oAction.Init REApplication.SessionContext
            
            Dim oAction2 As IBBAction2
            Set oAction2 = oAction
            
            With oAction
                oAction.Fields(ACTION_fld_CATEGORY) = "Mailing"
                oAction.Fields(ACTION_fld_DATE) = ActionDate
                oAction.Fields(ACTION_fld_TYPE) = "PALS Patient Packet-Online"
                oAction.Fields(ACTION_fld_RECORDS_ID) = oID
                
                oAction.Fields(ACTION_fld_STATUS) = "Order Fulfillment"
                oAction.Fields(ACTION_fld_AUTO_REMIND) = True
                oAction.Fields(ACTION_fld_NOTIFY_USING) = "Raiser's Edge reminders"
                oAction.Fields(ACTION_fld_PRIORITY) = "Normal"
                oAction.Fields(ACTION_fld_REMIND_FREQUENCY) = "day(s)"
                oAction.Fields(ACTION_fld_REMIND_VALUE) = 1
'
' Hard coded to add mgarcia - 74 as the user to be notified....
' Hard coded to add jbolduc - 379 as the user to be notified....
' Hard coded to add ahoward - 322 as the user to be notified....
'
                .Remindees.Add.Fields(ActionRemindee_fld_USER_ID) = 74
'                .Remindees.Add.Fields(ActionRemindee_fld_USER_ID) = 379
'                .Remindees.Add.Fields(ActionRemindee_fld_USER_ID) = 322
        
                Set oAttribute2 = oAction.Attributes.Add
                oAttribute2.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttributeID2
                oAttribute2.Fields(Attribute_fld_VALUE) = "1"
                oAttribute2.Fields(Attribute_fld_ATTRIBUTEDATE) = ActionDate
                oAttribute2.Fields(Attribute_fld_COMMENTS) = NoteType
            End With
            oAction.Save
            
            
            With oAttribute
                oAttribute.Fields(Attribute_fld_VALUE) = False
            End With
            oCons.Save
    
            oAction.Closedown
            Set oAction = Nothing
        End If
        
        oCons.Closedown
        Set oService = Nothing
        Set oAttributeServer = Nothing
        Set oAttribute = Nothing
        Set oCons = Nothing
    End If

End Sub

Public Sub QueryCurrentSolicitor(oRow As IBBQueryRow)
    Dim Const_ID As Long
    Const_ID = 1
    
    If oRow.BOF Then
        Rem MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        Rem MsgBox "End processing"
    Else
        Dim DevSolicitor As String
        Dim SecondaryRelManager As String
        Dim PrimaryRelManager As String
        Dim OtherSolicitor As String
        Dim SDevSolicitor As String
        Dim SSecondaryRelManager As String
        Dim SPrimaryRelManager As String
        Dim SOtherSolicitor As String
        
        DevSolicitor = ""
        SecondaryRelManager = ""
        PrimaryRelManager = ""
        OtherSolicitor = ""
        SDevSolicitor = ""
        SSecondaryRelManager = ""
        SPrimaryRelManager = ""
        SOtherSolicitor = ""
        
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init RE7.SessionContext
        
        Dim lID As Long
        lID = oRow.Field(Const_ID)
        oCons.Load lID
        
        Debug.Print oCons.Fields(RECORDS_fld_PRIMARY_ADDRESSEE)
        
        Dim oSolicitor As CSolicitorActions
        Dim oASolicitor As CAssignedSolicitor
        Dim oSolID As CActionSolicitor
        
        'loop through the constituent's assigned solicitors
        For Each oASolicitor In oCons.Relations.AssignedSolicitors
            If oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_DATE_TO) = "" Then
                If (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) = "Development Solicitor") Then
                    DevSolicitor = oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_NAME)
                End If
                If (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) = "Primary Relationship Manager") Then
                    PrimaryRelManager = oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_NAME)
                End If
                If (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) = "Secondary Relationship Manager") Then
                    SecondaryRelManager = oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_NAME)
                End If
                If (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) <> "Development Solicitor") And _
                    (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) <> "Primary Relationship Manager") And _
                    (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) <> "Secondary Relationship Manager") Then
                    OtherSolicitor = oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_NAME)
                End If
            End If
        Next oASolicitor
        
'
        Dim oCons2 As CRecord
        Set oCons2 = New CRecord
        oCons2.Init RE7.SessionContext
        
        Set REService = New REServices
        REService.Init REApplication.SessionContext
        
        If oRow.Field("Spouse System Record ID") <> "" Then
            Dim lid2 As Long
            lid2 = oRow.Field("Spouse System Record ID")
            oCons2.Load lid2
            For Each oASolicitor In oCons2.Relations.AssignedSolicitors
                If oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_DATE_TO) = "" Then
                    If (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) = "Development Solicitor") Then
                        SDevSolicitor = oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_NAME)
                    End If
                    If (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) = "Primary Relationship Manager") Then
                        SPrimaryRelManager = oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_NAME)
                    End If
                    If (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) = "Secondary Relationship Manager") Then
                        SSecondaryRelManager = oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_NAME)
                    End If
                    If (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) <> "Development Solicitor") And _
                        (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) <> "Primary Relationship Manager") And _
                        (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) <> "Secondary Relationship Manager") Then
                        SOtherSolicitor = oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_NAME)
                    End If
                End If
            Next oASolicitor
        End If

'
        oRow.Field("Development Solicitor") = DevSolicitor
        oRow.Field("Primary Relationship Manager") = PrimaryRelManager
        oRow.Field("Secondary Relationship Manager") = SecondaryRelManager
'        oRow.Field("Other Solicitor") = OtherSolicitor
'        oRow.Field("Spouse Development Solicitor") = SDevSolicitor
'        oRow.Field("Spouse Primary Relationship Manager") = SPrimaryRelManager
'        oRow.Field("Spouse Secondary Relationship Manager") = SSecondaryRelManager
'        oRow.Field("Spouse Other Solicitor") = SOtherSolicitor
                
        Set oASolicitor = Nothing
        Set oSolID = Nothing
        
        oCons.Closedown
        oCons2.Closedown
        Set oCons2 = Nothing
        Set oCons = Nothing
    
    End If
End Sub

Public Sub PrefAddressCheckSendMail(oRow As IBBQueryRow)
    Const ConsID = 1

    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
    
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init REApplication.SessionContext
        
        Dim oID As Long
        
        oID = oRow.Field(ConsID)
        oCons.Load oID
'
        Dim oAddress As CConstitAddress
        
        For Each oAddress In oCons.Addresses
            If oAddress.Fields(CONSTIT_ADDRESS_fld_PREFERRED) = True Then
                oAddress.Fields(CONSTIT_ADDRESS_fld_SENDMAIL) = True
                On Error Resume Next
                oCons.Save
            End If
        Next
        
'
        Set oAddress = Nothing
        oCons.Closedown
        Set oCons = Nothing
    End If

End Sub
Public Sub QueryProspectsSolicitorInfo(oRow As IBBQueryRow)
    Dim Const_ID As Long
    Const_ID = 1
    Const TAScore1 = 13
    Const TAScore2 = 14
    Const TAScore3 = 15
    Const TAScore4 = 16
    Const WPScore1 = 17
    
    If oRow.BOF Then
        Rem MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        Rem MsgBox "End processing"
    Else
        Dim DevSolicitor As String
        Dim SecondaryRelManager As String
        Dim PrimaryRelManager As String
        Dim OtherSolicitor As String
        Dim SDevSolicitor As String
        Dim SSecondaryRelManager As String
        Dim SPrimaryRelManager As String
        Dim SOtherSolicitor As String
        
        DevSolicitor = ""
        SecondaryRelManager = ""
        PrimaryRelManager = ""
        OtherSolicitor = ""
        SDevSolicitor = ""
        SSecondaryRelManager = ""
        SPrimaryRelManager = ""
        SOtherSolicitor = ""
        
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init RE7.SessionContext
        
        Dim lID As Long
        lID = oRow.Field(Const_ID)
        oCons.Load lID
        
        Debug.Print oCons.Fields(RECORDS_fld_PRIMARY_ADDRESSEE)
        
        Dim oSolicitor As CSolicitorActions
        Dim oASolicitor As CAssignedSolicitor
        Dim oSolID As CActionSolicitor
        
        'loop through the constituent's assigned solicitors
        For Each oASolicitor In oCons.Relations.AssignedSolicitors
            If oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_DATE_TO) = "" Then
                If (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) = "Development Solicitor") Then
                    DevSolicitor = oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_NAME)
                End If
                If (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) = "Primary Relationship Manager") Then
                    PrimaryRelManager = oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_NAME)
                End If
                If (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) = "Secondary Relationship Manager") Then
                    SecondaryRelManager = oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_NAME)
                End If
                If (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) <> "Development Solicitor") And _
                    (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) <> "Primary Relationship Manager") And _
                    (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) <> "Secondary Relationship Manager") Then
                    OtherSolicitor = oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_NAME)
                End If
            End If
        Next oASolicitor
'
        Dim oPros As CProspect
        Set oPros = oCons.Prospect
        
        Dim oRating As CRating
        
        For Each oRating In oPros.Ratings
            With oRating
                If oRating.Fields(RATING_fld_SOURCE) = "Blackbaud Analytics' Custom Modeling Service" Then
                    If oRating.Fields(RATING_fld_CATEGORY) = "CMS Target Gift Dollar Range" Then
                        Debug.Print oRating.Fields(RATING_fld_DESCRIPTION)
                        oRow.Field("CMS Target Gift Dollar Range") = oRating.Fields(RATING_fld_DESCRIPTION)
                    End If
                    If oRating.Fields(RATING_fld_CATEGORY) = "CMS Major Gift Likelihood" Then
                        oRow.Field("CMS Major Gift Likelihood") = oRating.Fields(RATING_fld_DESCRIPTION)
                    End If
                    If oRating.Fields(RATING_fld_CATEGORY) = "CMS Planned Gift Likelihood" Then
                        oRow.Field("CMS Planned Gift Likelihood") = oRating.Fields(RATING_fld_DESCRIPTION)
                    End If
                End If
                If oRating.Fields(RATING_fld_SOURCE) = "Target Analytics Custom Modeling Service" Then
                    If oRating.Fields(RATING_fld_CATEGORY) = "CMS Principal Giving Solution" Then
                        oRow.Field("CMS Principal Giving Solution") = oRating.Fields(RATING_fld_DESCRIPTION)
                    End If
                End If
                If oRating.Fields(RATING_fld_SOURCE) = "Target Analytics Custom Modeling Service" Then
                    If oRating.Fields(RATING_fld_CATEGORY) = "CMS WealthPoint Rating" Then
                        oRow.Field("CMS WealthPoint Rating") = oRating.Fields(RATING_fld_DESCRIPTION)
                    End If
                End If
                Set oRating = Nothing
            End With
        Next oRating
        
        Set oPros = Nothing

'
'        Dim oCons2 As CRecord
'        Set oCons2 = New CRecord
'        oCons2.Init RE7.SessionContext
'
'        Set REService = New REServices
'        REService.Init REApplication.SessionContext
'
'        If oRow.Field("Spouse System Record ID") <> "" Then
'            Dim lid2 As Long
'            lid2 = oRow.Field("Spouse System Record ID")
'            oCons2.Load lid2
'            For Each oASolicitor In oCons2.Relations.AssignedSolicitors
'                If oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_DATE_TO) = "" Then
'                    If (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) = "Development Solicitor") Then
'                        SDevSolicitor = oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_NAME)
'                    End If
'                    If (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) = "Primary Relationship Manager") Then
'                        SPrimaryRelManager = oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_NAME)
'                    End If
'                    If (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) = "Secondary Relationship Manager") Then
'                        SSecondaryRelManager = oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_NAME)
'                    End If
'                    If (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) <> "Development Solicitor") And _
'                        (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) <> "Primary Relationship Manager") And _
'                        (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) <> "Secondary Relationship Manager") Then
'                        SOtherSolicitor = oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_NAME)
'                    End If
'                End If
'            Next oASolicitor
'        End If

'
        oRow.Field("Development Solicitor") = DevSolicitor
        oRow.Field("Primary Relationship Manager") = PrimaryRelManager
        oRow.Field("Secondary Relationship Manager") = SecondaryRelManager
'        oRow.Field("Other Solicitor") = OtherSolicitor
'        oRow.Field("Spouse Development Solicitor") = SDevSolicitor
'        oRow.Field("Spouse Primary Relationship Manager") = SPrimaryRelManager
'        oRow.Field("Spouse Secondary Relationship Manager") = SSecondaryRelManager
'        oRow.Field("Spouse Other Solicitor") = SOtherSolicitor
                
        Set oASolicitor = Nothing
        Set oSolID = Nothing
        
        oCons.Closedown
'        oCons2.Closedown
'        Set oCons2 = Nothing
        Set oCons = Nothing
    
    End If
End Sub

Public Sub ExportProspectsSolicitorInfo(oRow As IBBExportRow)
    Dim Const_ID As Long
    Const_ID = 1
    Const TAScore1 = 13
    Const TAScore2 = 14
    Const TAScore3 = 15
    Const TAScore4 = 16
    Const WPScore1 = 17
    Const DevSol = 18
    Const PRM = 19
    Const SRM = 20
    Const PAC = 21
    
    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
        Dim DevSolicitor As String
        Dim SecondaryRelManager As String
        Dim PrimaryRelManager As String
        Dim OtherSolicitor As String
        Dim SDevSolicitor As String
        Dim SSecondaryRelManager As String
        Dim SPrimaryRelManager As String
        Dim SOtherSolicitor As String
        Dim SPatientCentralAss As String
        
        DevSolicitor = ""
        SecondaryRelManager = ""
        PrimaryRelManager = ""
        OtherSolicitor = ""
        SDevSolicitor = ""
        SSecondaryRelManager = ""
        SPrimaryRelManager = ""
        SOtherSolicitor = ""
        SPatientCentralAss = ""
        
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init RE7.SessionContext
        
        Dim lID As Long
        lID = oRow.Field(Const_ID)
        oCons.Load lID
        
        Debug.Print oCons.Fields(RECORDS_fld_CONSTITUENT_ID) & ": " & oCons.Fields(RECORDS_fld_PRIMARY_ADDRESSEE)
        
        Dim oSolicitor As CSolicitorActions
        Dim oASolicitor As CAssignedSolicitor
        Dim oSolID As CActionSolicitor
        
        Debug.Print "finding solicitor... "
        'loop through the constituent's assigned solicitors
        For Each oASolicitor In oCons.Relations.AssignedSolicitors
            If oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_DATE_TO) = "" Then
                If (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) = "Development Solicitor") Then
                    DevSolicitor = oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_NAME)
                End If
                If (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) = "Primary Relationship Manager") Then
                    PrimaryRelManager = oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_NAME)
                End If
                If (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) = "Secondary Relationship Manager") Then
                    SecondaryRelManager = oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_NAME)
                End If
                If (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) = "Patient Central Associate") Then
                    SPatientCentralAss = oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_NAME)
                End If
                If (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) <> "Development Solicitor") And _
                    (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) <> "Primary Relationship Manager") And _
                    (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) <> "Secondary Relationship Manager") And _
                    (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) <> "Patient Central Associate") Then
                    OtherSolicitor = oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_NAME)
                End If
            End If
        Next oASolicitor
'
        Dim oPros As CProspect
        Set oPros = oCons.Prospect
        
        Dim oRating As CRating
        
        Debug.Print "finding rating..."
        For Each oRating In oPros.Ratings
            Debug.Print "finding...."
            With oRating
                If oRating.Fields(RATING_fld_SOURCE) = "Blackbaud Analytics' Custom Modeling Service" Then
                    If oRating.Fields(RATING_fld_CATEGORY) = "CMS Target Gift Dollar Range" Then
                        Debug.Print oRating.Fields(RATING_fld_DESCRIPTION)
                        oRow.Field(TAScore1) = oRating.Fields(RATING_fld_DESCRIPTION)
                    End If
                    If oRating.Fields(RATING_fld_CATEGORY) = "CMS Major Gift Likelihood" Then
                        oRow.Field(TAScore2) = oRating.Fields(RATING_fld_DESCRIPTION)
                    End If
                    If oRating.Fields(RATING_fld_CATEGORY) = "CMS Planned Gift Likelihood" Then
                        oRow.Field(TAScore3) = oRating.Fields(RATING_fld_DESCRIPTION)
                    End If
                End If
                If oRating.Fields(RATING_fld_SOURCE) = "Target Analytics Custom Modeling Service" Then
                    If oRating.Fields(RATING_fld_CATEGORY) = "CMS Principal Giving Solution" Then
                        oRow.Field(TAScore4) = oRating.Fields(RATING_fld_DESCRIPTION)
                    End If
                End If
                If oRating.Fields(RATING_fld_SOURCE) = "Target Analytics Custom Modeling Service" Then
                    If oRating.Fields(RATING_fld_CATEGORY) = "CMS WealthPoint Rating" Then
                        oRow.Field(WPScore1) = oRating.Fields(RATING_fld_DESCRIPTION)
                    End If
                End If
                Set oRating = Nothing
            End With
        Next oRating
        


'
'        Dim oCons2 As CRecord
'        Set oCons2 = New CRecord
'        oCons2.Init RE7.SessionContext
'
'        Set REService = New REServices
'        REService.Init REApplication.SessionContext
'
'        If oRow.Field("Spouse System Record ID") <> "" Then
'            Dim lid2 As Long
'            lid2 = oRow.Field("Spouse System Record ID")
'            oCons2.Load lid2
'            For Each oASolicitor In oCons2.Relations.AssignedSolicitors
'                If oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_DATE_TO) = "" Then
'                    If (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) = "Development Solicitor") Then
'                        SDevSolicitor = oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_NAME)
'                    End If
'                    If (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) = "Primary Relationship Manager") Then
'                        SPrimaryRelManager = oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_NAME)
'                    End If
'                    If (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) = "Secondary Relationship Manager") Then
'                        SSecondaryRelManager = oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_NAME)
'                    End If
'                    If (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) <> "Development Solicitor") And _
'                        (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) <> "Primary Relationship Manager") And _
'                        (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) <> "Secondary Relationship Manager") Then
'                        SOtherSolicitor = oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_NAME)
'                    End If
'                End If
'            Next oASolicitor
'        End If

'
        Debug.Print "storing... "
        oRow.Field(DevSol) = DevSolicitor
        oRow.Field(PRM) = PrimaryRelManager
        oRow.Field(SRM) = SecondaryRelManager
        oRow.Field(PAC) = SPatientCentralAss
'        oRow.Field("Other Solicitor") = OtherSolicitor
'        oRow.Field("Spouse Development Solicitor") = SDevSolicitor
'        oRow.Field("Spouse Primary Relationship Manager") = SPrimaryRelManager
'        oRow.Field("Spouse Secondary Relationship Manager") = SSecondaryRelManager
'        oRow.Field("Spouse Other Solicitor") = SOtherSolicitor
                
        Set oPros = Nothing
        Set oASolicitor = Nothing
        Set oSolID = Nothing
        
        oCons.Closedown
'        oCons2.Closedown
'        Set oCons2 = Nothing
        Set oCons = Nothing
    
    End If
End Sub

Public Sub AddDIYpreEventActions(oRow As IBBQueryRow)
Const ConsID = 1

    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
'
'
'   Get the attribute ID
'
        Dim oService As REServices
        Set oService = New REServices
        oService.Init REApplication.SessionContext
           
        Dim oAttributeServer As CAttributeTypeServer
        Set oAttributeServer = oService.CreateServiceObject(bbsoAttributeTypeServer)
        oAttributeServer.Init REApplication.SessionContext
        
        Dim lAttributeID As Long
        Dim lAttributeID2 As Long
        Dim lAttributeID3 As Long
        Dim lAttributeID4 As Long
        Dim lAttributeID5 As Long
        Dim lAttributeID6 As Long
        Dim lAttributeID7 As Long
        Dim lAttributeID8 As Long
        Dim lAttributeID9 As Long
        Dim lAttributeID10 As Long
        Dim lAttributeID11 As Long
            
        lAttributeID = oAttributeServer.GetAttributeTypeID("DIY-What is your event name?", bbAttributeRecordType_ACTION)
        lAttributeID2 = oAttributeServer.GetAttributeTypeID("DIY-What is the location of your event?", bbAttributeRecordType_ACTION)
        lAttributeID3 = oAttributeServer.GetAttributeTypeID("DIY-How many attendess will there be?", bbAttributeRecordType_ACTION)
        lAttributeID4 = oAttributeServer.GetAttributeTypeID("DIY-Are you intrested in mailed materials?", bbAttributeRecordType_ACTION)
        lAttributeID5 = oAttributeServer.GetAttributeTypeID("DIY-Address to receive the materials?", bbAttributeRecordType_ACTION)
        lAttributeID6 = oAttributeServer.GetAttributeTypeID("DIY-Considering listing your event on a calendar?", bbAttributeRecordType_ACTION)
        lAttributeID7 = oAttributeServer.GetAttributeTypeID("DIY-If yes, what details should be included?", bbAttributeRecordType_ACTION)
        lAttributeID8 = oAttributeServer.GetAttributeTypeID("DIY-Interested in having a local representative?", bbAttributeRecordType_ACTION)

        Dim Question1 As String
        Dim Question2 As String
        Dim Question3 As String
        Dim Question4 As Boolean
        Dim Question5 As String
        Dim Question6 As Boolean
        Dim Question7 As String
        Dim Question8 As Boolean
        Dim AddTYNote As Boolean
        
        Question1 = ""
        Question2 = ""
        Question3 = ""
        Question4 = False
        Question5 = ""
        Question6 = False
        Question7 = ""
        Question8 = False
        AddTYNote = False

        Dim oAttribute  As IBBAttribute
        Dim oAttribute2 As IBBAttribute
        Dim oPart As CParticipant
        Dim lPartID As Long
        
        Dim ActionDate As Date
        ActionDate = Date
'
'   Load the constituent record
'
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init REApplication.SessionContext

        Dim oID As Long

        oID = oRow.Field(ConsID)
        oCons.Load oID
'
'   Create the action and note
'
            Dim oAction As CAction
            Set oAction = New CAction
'            oAction.Init REApplication.SessionContext
            
'            Dim oAction2 As IBBAction2
'            Set oAction2 = oAction
            
            
            For Each oAction In oCons.Actions
                If oAction.Fields(ACTION_fld_TYPE) = "DIY-Pre-Event Survey" And _
                    oAction.Fields(ACTION_fld_COMPLETED) = False Then
                    For Each oAttribute In oAction.Attributes
                        If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttributeID Then
                            Question1 = oAttribute.Fields(Attribute_fld_VALUE)
                        End If
                        If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttributeID2 Then
                            Question2 = oAttribute.Fields(Attribute_fld_VALUE)
                        End If
                        If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttributeID3 Then
                            Question3 = oAttribute.Fields(Attribute_fld_VALUE)
                        End If
                        If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttributeID4 Then
                            Question4 = oAttribute.Fields(Attribute_fld_VALUE)
                        End If
                        If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttributeID5 Then
                            Question5 = oAttribute.Fields(Attribute_fld_VALUE)
                        End If
                        If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttributeID6 Then
                            Question6 = oAttribute.Fields(Attribute_fld_VALUE)
                        End If
                        If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttributeID7 Then
                            Question7 = oAttribute.Fields(Attribute_fld_VALUE)
                        End If
                        If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttributeID8 Then
                            Question8 = oAttribute.Fields(Attribute_fld_VALUE)
                        End If
                        ActionDate = oAction.Fields(ACTION_fld_DATE)
                    Next oAttribute
'
'   Mark the main action as complete
'
                    With oAction
                        oAction.Fields(ACTION_fld_COMPLETED) = True
                        oAction.Fields(ACTION_fld_COMPLETED_DATE) = Date
                        oAction.Save
                    End With
                    AddTYNote = True
                End If
            Next oAction
            
            Dim oAction2 As CAction
            Set oAction2 = New CAction
            oAction2.Init REApplication.SessionContext
'
'   Create the action for the mail in materials
'
            If Question4 = True Then
                With oAction2
                    oAction2.Fields(ACTION_fld_CATEGORY) = "Mailing"
                    oAction2.Fields(ACTION_fld_DATE) = ActionDate
                    oAction2.Fields(ACTION_fld_TYPE) = "DIY-Mail in Materials"
                    oAction2.Fields(ACTION_fld_RECORDS_ID) = oID

                    oAction2.Fields(ACTION_fld_AUTO_REMIND) = True
                    oAction2.Fields(ACTION_fld_NOTIFY_USING) = "Raiser's Edge reminders"
                    oAction2.Fields(ACTION_fld_PRIORITY) = "Normal"
                    oAction2.Fields(ACTION_fld_REMIND_FREQUENCY) = "day(s)"
                    oAction2.Fields(ACTION_fld_REMIND_VALUE) = 1
'
'     Hard coded to add kseccombe - 213 as the user to be notified....
'     Hard coded to add ematteucci - 282 as the user to be notified....
'
                    .Remindees.Add.Fields(ActionRemindee_fld_USER_ID) = 282

                    Set oAttribute = oAction2.Attributes.Add
                    oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttributeID
                    oAttribute.Fields(Attribute_fld_VALUE) = Question1
                    oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE) = ActionDate
                    Set oAttribute = oAction2.Attributes.Add
                    oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttributeID2
                    oAttribute.Fields(Attribute_fld_VALUE) = Question2
                    oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE) = ActionDate
                    Set oAttribute = oAction2.Attributes.Add
                    oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttributeID3
                    oAttribute.Fields(Attribute_fld_VALUE) = Question3
                    oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE) = ActionDate
                    Set oAttribute = oAction2.Attributes.Add
                    oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttributeID4
                    oAttribute.Fields(Attribute_fld_VALUE) = Question4
                    oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE) = ActionDate
                    Set oAttribute = oAction2.Attributes.Add
                    oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttributeID5
                    oAttribute.Fields(Attribute_fld_VALUE) = Question5
                    oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE) = ActionDate
                    
                End With
                oAction2.Save
            End If
            
            oAction2.Closedown
            Set oAction2 = Nothing
            
            Set oAction2 = New CAction
            oAction2.Init REApplication.SessionContext
'
'   Create the action for listing the event on a calendar
'
            If Question6 = True Then
                With oAction2
                    oAction2.Fields(ACTION_fld_CATEGORY) = "Task/Other"
                    oAction2.Fields(ACTION_fld_DATE) = ActionDate
                    oAction2.Fields(ACTION_fld_TYPE) = "DIY-Event to Calendar"
                    oAction2.Fields(ACTION_fld_RECORDS_ID) = oID

                    oAction2.Fields(ACTION_fld_AUTO_REMIND) = True
                    oAction2.Fields(ACTION_fld_NOTIFY_USING) = "Raiser's Edge reminders"
                    oAction2.Fields(ACTION_fld_PRIORITY) = "Normal"
                    oAction2.Fields(ACTION_fld_REMIND_FREQUENCY) = "day(s)"
                    oAction2.Fields(ACTION_fld_REMIND_VALUE) = 1
'
'     Hard coded to add kseccombe - 213 as the user to be notified....
'     Hard coded to add ematteucci - 282 as the user to be notified....
'
                    .Remindees.Add.Fields(ActionRemindee_fld_USER_ID) = 282
                    
                    Set oAttribute = oAction2.Attributes.Add
                    oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttributeID
                    oAttribute.Fields(Attribute_fld_VALUE) = Question1
                    oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE) = ActionDate
                    Set oAttribute = oAction2.Attributes.Add
                    oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttributeID2
                    oAttribute.Fields(Attribute_fld_VALUE) = Question2
                    oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE) = ActionDate
                    Set oAttribute = oAction2.Attributes.Add
                    oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttributeID3
                    oAttribute.Fields(Attribute_fld_VALUE) = Question3
                    oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE) = ActionDate
                    Set oAttribute = oAction2.Attributes.Add
                    oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttributeID6
                    oAttribute.Fields(Attribute_fld_VALUE) = Question6
                    oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE) = ActionDate
                    Set oAttribute = oAction2.Attributes.Add
                    oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttributeID7
                    oAttribute.Fields(Attribute_fld_VALUE) = Question7
                    oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE) = ActionDate
                End With
                oAction2.Save
            End If
    
            oAction2.Closedown
            Set oAction2 = Nothing
            
            Set oAction2 = New CAction
            oAction2.Init REApplication.SessionContext
'
'   Create the action if interested in having a local representative
'
            If Question8 = True Then
                With oAction2
                    oAction2.Fields(ACTION_fld_CATEGORY) = "Task/Other"
                    oAction2.Fields(ACTION_fld_DATE) = ActionDate
                    oAction2.Fields(ACTION_fld_TYPE) = "DIY-Local Representative"
                    oAction2.Fields(ACTION_fld_RECORDS_ID) = oID

                    oAction2.Fields(ACTION_fld_AUTO_REMIND) = True
                    oAction2.Fields(ACTION_fld_NOTIFY_USING) = "Raiser's Edge reminders"
                    oAction2.Fields(ACTION_fld_PRIORITY) = "Normal"
                    oAction2.Fields(ACTION_fld_REMIND_FREQUENCY) = "day(s)"
                    oAction2.Fields(ACTION_fld_REMIND_VALUE) = 1
'
'     Hard coded to add kseccombe - 213 as the user to be notified....
'     Hard coded to add ematteucci - 282 as the user to be notified....
'
                    .Remindees.Add.Fields(ActionRemindee_fld_USER_ID) = 282
                    Set oAttribute = oAction2.Attributes.Add
                    oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttributeID
                    oAttribute.Fields(Attribute_fld_VALUE) = Question1
                    oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE) = ActionDate
                    Set oAttribute = oAction2.Attributes.Add
                    oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttributeID2
                    oAttribute.Fields(Attribute_fld_VALUE) = Question2
                    oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE) = ActionDate
                    Set oAttribute = oAction2.Attributes.Add
                    oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttributeID3
                    oAttribute.Fields(Attribute_fld_VALUE) = Question3
                    oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE) = ActionDate
                    Set oAttribute = oAction2.Attributes.Add
                    oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttributeID8
                    oAttribute.Fields(Attribute_fld_VALUE) = Question8
                    oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE) = ActionDate
                End With
                oAction2.Save
            End If

            oAction2.Closedown
            Set oAction2 = Nothing
            
            Set oAction2 = New CAction
            oAction2.Init REApplication.SessionContext
'
'   Create the action to send thank you note
'
            If AddTYNote = True Then
                With oAction2
                    oAction2.Fields(ACTION_fld_CATEGORY) = "Mailing"
                    oAction2.Fields(ACTION_fld_DATE) = ActionDate
                    oAction2.Fields(ACTION_fld_TYPE) = "DIY-Thank You Note"
                    oAction2.Fields(ACTION_fld_RECORDS_ID) = oID

                    oAction2.Fields(ACTION_fld_AUTO_REMIND) = True
                    oAction2.Fields(ACTION_fld_NOTIFY_USING) = "Raiser's Edge reminders"
                    oAction2.Fields(ACTION_fld_PRIORITY) = "Normal"
                    oAction2.Fields(ACTION_fld_REMIND_FREQUENCY) = "day(s)"
                    oAction2.Fields(ACTION_fld_REMIND_VALUE) = 1
'
'     Hard coded to add kseccombe - 213 as the user to be notified....
'     Hard coded to add ematteucci - 282 as the user to be notified....
'
                    .Remindees.Add.Fields(ActionRemindee_fld_USER_ID) = 282
                End With
                oAction2.Save
            End If
            
            oAction2.Closedown
            Set oAction2 = Nothing
            Set oAction = Nothing
        
        oCons.Closedown
        Set oService = Nothing
        Set oAttributeServer = Nothing
        Set oAttribute = Nothing
        Set oCons = Nothing
    End If

End Sub

Private Sub AddPatient(ByVal oCon As Object)

    Dim oCCode As CConstituentCode
    Dim FoundIt As Boolean
    FoundIt = False
    
    For Each oCCode In oCon.ConstituentCodes
        If (oCCode.Fields(CONSTITUENT_CODE_fld_LONG_DESC_READONLY) = "Patient/Survivor") Then
            FoundIt = True
            With oCCode
                .Fields(CONSTITUENT_CODE_fld_DATE_TO) = ""
            End With
        End If
    Next oCCode
    
    If FoundIt = False Then
        With oCons.ConstituentCodes.Add
            .Fields(CONSTITUENT_CODE_fld_CODE) = "PS"
            .Fields(CONSTITUENT_CODE_fld_DATE_FROM) = Format(Date, "MM/DD/YYYY")
        End With
    End If
    
    oCon.Save
    Set oCCode = Nothing
    
End Sub
Public Sub UpdatePCConnection(oRow As IBBQueryRow)
    Const ConsID = 1

    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
    
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init REApplication.SessionContext
      
        Dim oID As Long
        
        oID = oRow.Field(ConsID)
        oCons.Load oID
'
'   Get the attribute ID
'
        Dim oService As REServices
        Set oService = New REServices
        oService.Init REApplication.SessionContext
           
        Dim oAttributeServer As CAttributeTypeServer
        Set oAttributeServer = oService.CreateServiceObject(bbsoAttributeTypeServer)
        oAttributeServer.Init REApplication.SessionContext
        
        Dim lAttributeID As Long
        Dim CaregiverID As Long
        Dim HProID As Long
        Dim FriendsID As Long
        Dim ResearcherID As Long
        
            
        lAttributeID = oAttributeServer.GetAttributeTypeID("PC_Connection", bbAttributeRecordType_CONSTITUENT)
        CaregiverID = oAttributeServer.GetAttributeTypeID("Caregiver", bbAttributeRecordType_CONSTITUENT)
        HProID = oAttributeServer.GetAttributeTypeID("Healthcare Professional", bbAttributeRecordType_CONSTITUENT)
        FriendsID = oAttributeServer.GetAttributeTypeID("Caregiver/Family Member/Friend", bbAttributeRecordType_CONSTITUENT)
        ResearcherID = oAttributeServer.GetAttributeTypeID("Researcher", bbAttributeRecordType_CONSTITUENT)
        
        Dim oAttribute As IBBAttribute
        Dim oAttribute2 As IBBAttribute
        Dim oCCode As CConstituentCode
        Dim FoundIt As Boolean
'
        For Each oAttribute In oCons.Attributes
            Debug.Print oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) & " " & lAttributeID
'
'Flag the caregiver
'
            If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttributeID And _
                oAttribute.Fields(Attribute_fld_VALUE) = "Caregiver" Then
                FoundIt = False
                For Each oAttribute2 In oCons.Attributes
                    If oAttribute2.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = CaregiverID Then
                        FoundIt = True
                    End If
                Next oAttribute2
                If FoundIt = False Then
                    Set oAttribute2 = oCons.Attributes.Add
                    oAttribute2.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = CaregiverID
                    oAttribute2.Fields(Attribute_fld_VALUE) = True
                    oAttribute2.Fields(Attribute_fld_ATTRIBUTEDATE) = Format(Date, "MM/DD/YYYY")
                    oAttribute2.Fields(Attribute_fld_COMMENTS) = "PC Connection"
                    oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE) = Format(Date, "MM/DD/YYYY")
                    On Error Resume Next
                    oCons.Save
                End If
            End If
'
'Flag the family
'
            If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttributeID And _
                oAttribute.Fields(Attribute_fld_VALUE) = "Family or friend of a patient/survivor" Then
                FoundIt = False
                For Each oAttribute2 In oCons.Attributes
                    If oAttribute2.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = FriendsID Then
                        FoundIt = True
                    End If
                Next oAttribute2
                If FoundIt = False Then
                    Set oAttribute2 = oCons.Attributes.Add
                    oAttribute2.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = FriendsID
                    oAttribute2.Fields(Attribute_fld_VALUE) = True
                    oAttribute2.Fields(Attribute_fld_ATTRIBUTEDATE) = Format(Date, "MM/DD/YYYY")
                    oAttribute2.Fields(Attribute_fld_COMMENTS) = "PC Connection"
                    oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE) = Format(Date, "MM/DD/YYYY")
                    On Error Resume Next
                    oCons.Save
                End If
            End If
'
'Flag the HP in PanCAN
'
            If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttributeID And _
                oAttribute.Fields(Attribute_fld_VALUE) = "Healthcare professional in the pancreatic cancer field" Then
                
                FoundIt = False
                For Each oCCode In oCons.ConstituentCodes
                    If (oCCode.Fields(CONSTITUENT_CODE_fld_LONG_DESC_READONLY) = "Healthcare Professional") Then
                        FoundIt = True
                        With oCCode
                            .Fields(CONSTITUENT_CODE_fld_DATE_TO) = ""
                        End With
                    End If
                Next oCCode
                
                If FoundIt = False Then
                    With oCons.ConstituentCodes.Add
                        .Fields(CONSTITUENT_CODE_fld_CODE) = "Healthcare Professional"
                        .Fields(CONSTITUENT_CODE_fld_DATE_FROM) = Format(Date, "MM/DD/YYYY")
                    End With
                    On Error Resume Next
                    oCons.Save
                End If
'
                FoundIt = False
                
                For Each oAttribute2 In oCons.Attributes
                    If oAttribute2.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = HProID Then
                        FoundIt = True
                    End If
                Next oAttribute2
                If FoundIt = False Then
                    Set oAttribute2 = oCons.Attributes.Add
                    oAttribute2.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = HProID
                    oAttribute2.Fields(Attribute_fld_VALUE) = "Pancreatic Cancer Field"
                    oAttribute2.Fields(Attribute_fld_ATTRIBUTEDATE) = Format(Date, "MM/DD/YYYY")
                    oAttribute2.Fields(Attribute_fld_COMMENTS) = "PC Connection"
                    oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE) = Format(Date, "MM/DD/YYYY")
                    On Error Resume Next
                    oCons.Save
                End If
'
                oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE) = Format(Date, "MM/DD/YYYY")
                On Error Resume Next
                oCons.Save
            End If
'
'Flag the researcher
'
            If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttributeID And _
                oAttribute.Fields(Attribute_fld_VALUE) = "Pancreatic cancer researcher" Then
                FoundIt = False
                For Each oAttribute2 In oCons.Attributes
                    If oAttribute2.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = ResearcherID Then
                        FoundIt = True
                    End If
                Next oAttribute2
                If FoundIt = False Then
                    Set oAttribute2 = oCons.Attributes.Add
                    oAttribute2.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = ResearcherID
                    oAttribute2.Fields(Attribute_fld_VALUE) = "YES"
                    oAttribute2.Fields(Attribute_fld_ATTRIBUTEDATE) = Format(Date, "MM/DD/YYYY")
                    oAttribute2.Fields(Attribute_fld_COMMENTS) = "PC Connection"
                    oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE) = Format(Date, "MM/DD/YYYY")
                    On Error Resume Next
                    oCons.Save
                End If
            End If
'
'Flag the patients
'
            If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttributeID And _
                oAttribute.Fields(Attribute_fld_VALUE) = "Patient/Survivor" Then

                FoundIt = False

                For Each oCCode In oCons.ConstituentCodes
                    If (oCCode.Fields(CONSTITUENT_CODE_fld_LONG_DESC_READONLY) = "Patient/Survivor") Then
                        FoundIt = True
                        With oCCode
                            .Fields(CONSTITUENT_CODE_fld_DATE_TO) = ""
                        End With
                    End If
                Next oCCode

                If FoundIt = False Then
                    With oCons.ConstituentCodes.Add
                        .Fields(CONSTITUENT_CODE_fld_CODE) = "Patient/Survivor"
                        .Fields(CONSTITUENT_CODE_fld_DATE_FROM) = Format(Date, "MM/DD/YYYY")
                    End With
                End If

                oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE) = Format(Date, "MM/DD/YYYY")
                On Error Resume Next
                oCons.Save
            End If
        Next oAttribute
        Set oCCode = Nothing
        Set oAttribute = Nothing
        oCons.Closedown
        Set oCons = Nothing
    End If

End Sub
Public Sub AddTeamNametoGift(oRow As IBBQueryRow)

    Const PartID = 1

    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
    
        Dim oParticipant As CParticipant
        Set oParticipant = New CParticipant
        oParticipant.Init REApplication.SessionContext
        
        Dim oService As REServices
        Set oService = New REServices
        oService.Init REApplication.SessionContext
           
        Dim oAttributeServer As CAttributeTypeServer
        Set oAttributeServer = oService.CreateServiceObject(bbsoAttributeTypeServer)
        oAttributeServer.Init REApplication.SessionContext
        
        Dim lAttributeID As Long
        Dim lAttributeID2 As Long
        
        lAttributeID = oAttributeServer.GetAttributeTypeID("TR team name", bbAttributeRecordType_PARTICIPANT)
        lAttributeID2 = oAttributeServer.GetAttributeTypeID("Luminate Online Team Name", bbAttributeRecordType_GIFT)
        
        
        Dim oID As Long
        oID = oRow.Field(PartID)
        oParticipant.Load oID
        
        Dim oDonation As CParticipantDonation
        Dim oGift As CGift
        
        Dim GiftID As Long
        
        Dim oAttribute  As IBBAttribute
        Dim oAttribute2 As IBBAttribute
        
        Dim TeamName As String
        TeamName = ""
        
        Debug.Print oParticipant.Fields(Participants_fld_Name) & " " & lAttributeID
        
        For Each oAttribute In oParticipant.Attributes
            Debug.Print "finding participant attribute " & oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID)
            If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttributeID Then
                TeamName = oAttribute.Fields(Attribute_fld_VALUE)
                Debug.Print "Found it: " & TeamName
            End If
        Next oAttribute
        
        Set oGift = New CGift
        oGift.Init RE7.SessionContext
        
        For Each oDonation In oParticipant.Donations
'
            GiftID = oDonation.Fields(LinkedGift_fld_GiftID)
            Debug.Print GiftID
            oGift.Load GiftID
            
            With oGift
                Set oAttribute2 = oGift.Attributes.Add
                oAttribute2.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttributeID2
                oAttribute2.Fields(Attribute_fld_VALUE) = TeamName
                oAttribute2.Fields(Attribute_fld_ATTRIBUTEDATE) = Format(Date, "MM/DD/YYYY")
                oAttribute2.Fields(Attribute_fld_COMMENTS) = ""
                On Error Resume Next
                oGift.Save
            End With

        Next oDonation
            
        Set oDonation = Nothing
        Set oGift = Nothing
        Set oService = Nothing
        Set oAttributeServer = Nothing
        Set oAttribute = Nothing
        Set oAttribute2 = Nothing
            
        oParticipant.Closedown
        Set oParticipant = Nothing
    End If
End Sub
Public Sub DeleteSolicitor(oRow As IBBQueryRow)
    Const ConsID = 1

    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
    
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init REApplication.SessionContext
        
        Dim oID As Long
        
        oID = oRow.Field(ConsID)
        oCons.Load oID
'
        Dim oSolicitor As CAssignedSolicitor
        
        'loop through the constituent's assigned solicitors
        For Each oSolicitor In oCons.Relations.AssignedSolicitors
            If oSolicitor.Fields(ASSIGNEDSOLICITOR_fld_DATE_FROM) = "7/1/2016" And _
                 oSolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) = "Primary Relationship Manager" Then
                 Debug.Print oSolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_NAME)
'                 oCons.Relations.AssignedSolicitors.Remove oSolicitor
'                'Save parent object to save the child collection
'                oCons.Save
            End If
        Next oSolicitor
        
        'clean up
        Set oSolicitor = Nothing
'
        oCons.Closedown
        Set oCons = Nothing
    End If

End Sub
Public Sub RemoveEventParticipantion(oRow As IBBQueryRow)
    Const ConsID = 1

    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
    
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init REApplication.SessionContext
        
        Dim oID As Long
        
        oID = oRow.Field(ConsID)
        oCons.Load oID
'
        Dim lCount As Long
        
        For lCount = 1 To oCons.Participants.Count
            If oCons.Participants.Item(lCount).EventObject.Fields(SPECIAL_EVENT_fld_NAME) = "2016 Leadership Breakfast Boston" Then
                Debug.Print oCons.Participants.Item(lCount).EventObject.Fields(SPECIAL_EVENT_fld_NAME)
                oCons.Participants.Item(lCount).Delete
                oCons.Save
            End If
        Next lCount
'
        oCons.Closedown
        Set oCons = Nothing
    End If

End Sub
Public Sub QueryPiPNSolicitor(oRow As IBBQueryRow)
    Dim Const_ID As Long
    Const_ID = 1
    
    If oRow.BOF Then
        Rem MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        Rem MsgBox "End processing"
    Else
        Dim DevSolicitor As String
        Dim SecondaryRelManager As String
        Dim PrimaryRelManager As String
        Dim OtherSolicitor As String
        
        DevSolicitor = ""
        SecondaryRelManager = ""
        PrimaryRelManager = ""
        OtherSolicitor = ""
    
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init RE7.SessionContext
        
        Dim lID As Long
        lID = oRow.Field(Const_ID)
        oCons.Load lID
        
        Debug.Print oCons.Fields(RECORDS_fld_PRIMARY_ADDRESSEE)
        
        Dim oSolicitor As CSolicitorActions
        Dim oASolicitor As CAssignedSolicitor
        Dim oSolID As CActionSolicitor
        
        'loop through the constituent's assigned solicitors
        For Each oASolicitor In oCons.Relations.AssignedSolicitors
            If oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_DATE_TO) = "" Then
                If (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) = "Development Solicitor") Then
                    If Trim(DevSolicitor) <> "" Then
                        DevSolicitor = Trim(DevSolicitor) & ", " & oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_NAME)
                    Else
                        DevSolicitor = oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_NAME)
                    End If
                End If
                If (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) = "Primary Relationship Manager") Then
                    If Trim(PrimaryRelManager) <> "" Then
                        PrimaryRelManager = Trim(PrimaryRelManager) & ", " & oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_NAME)
                    Else
                        PrimaryRelManager = oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_NAME)
                    End If
                End If
                If (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) = "Secondary Relationship Manager") Then
                    If Trim(SecondaryRelManager) <> "" Then
                        SecondaryRelManager = Trim(SecondaryRelManager) & ", " & oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_NAME)
                    Else
                        SecondaryRelManager = oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_NAME)
                    End If

                End If
            End If
        Next oASolicitor
'
'   Get the attribute ID
'
        Dim oService As REServices
        Set oService = New REServices
        oService.Init REApplication.SessionContext

        Dim oAttributeServer As CAttributeTypeServer
        Set oAttributeServer = oService.CreateServiceObject(bbsoAttributeTypeServer)
        oAttributeServer.Init REApplication.SessionContext

        Dim lAttID1 As Long
        Dim lAttID2 As Long
        Dim lAttID3 As Long
        Dim lAttID4 As Long
        Dim lAttID5 As Long
        Dim lAttID6 As Long
        Dim lAttID7 As Long
        Dim lAttID8 As Long
        Dim lAttID9 As Long
        Dim lAttID10 As Long

        lAttID1 = oAttributeServer.GetAttributeTypeID("CORP Partners in Progress - $1,000-2,499", bbAttributeRecordType_CONSTITUENT)
        lAttID2 = oAttributeServer.GetAttributeTypeID("CORP Partners in Progress - $2,500-4,999", bbAttributeRecordType_CONSTITUENT)
        lAttID3 = oAttributeServer.GetAttributeTypeID("CORP Partners in Progress - $5,000-9,999", bbAttributeRecordType_CONSTITUENT)
        lAttID4 = oAttributeServer.GetAttributeTypeID("CORP Partners in Progress - $10,000-24,999", bbAttributeRecordType_CONSTITUENT)
        lAttID5 = oAttributeServer.GetAttributeTypeID("CORP Partners in Progress - $25,000+", bbAttributeRecordType_CONSTITUENT)
        lAttID6 = oAttributeServer.GetAttributeTypeID("Partners in Progress - $1,000-2,499", bbAttributeRecordType_CONSTITUENT)
        lAttID7 = oAttributeServer.GetAttributeTypeID("Partners in Progress - $2,500-4,999", bbAttributeRecordType_CONSTITUENT)
        lAttID8 = oAttributeServer.GetAttributeTypeID("Partners in Progress - $5,000-9,999", bbAttributeRecordType_CONSTITUENT)
        lAttID9 = oAttributeServer.GetAttributeTypeID("Partners in Progress - $10,000-24,999", bbAttributeRecordType_CONSTITUENT)
        lAttID10 = oAttributeServer.GetAttributeTypeID("Partners in Progress - $25,000+", bbAttributeRecordType_CONSTITUENT)

        Dim oAttribute As IBBAttribute

        Dim NoteType As String
        NoteType = ""

        Dim NoteDate As String
        NoteDate = ""

        For Each oAttribute In oCons.Attributes

            Debug.Print oAttribute.Fields(Attribute_fld_ATTRIBUTES_ID)

            If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttID1 Then
                If NoteType = "" Then
                    NoteType = "CORP Partners in Progress - $1,000-2,499"
                Else
                    NoteType = NoteType & ", CORP Partners in Progress - $1,000-2,499"
                End If
                If NoteDate = "" Then
                    NoteDate = oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE)
                Else
                    NoteDate = NoteDate & ", " & oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE)
                End If
            End If
            If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttID2 Then
                If NoteType = "" Then
                    NoteType = "CORP Partners in Progress - $2,500-4,999"
                Else
                    NoteType = NoteType & ", CORP Partners in Progress - $2,500-4,999"
                End If
                If NoteDate = "" Then
                    NoteDate = oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE)
                Else
                    NoteDate = NoteDate & ", " & oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE)
                End If
            End If
            If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttID3 Then
                If NoteType = "" Then
                    NoteType = "CORP Partners in Progress - $5,000-9,999"
                Else
                    NoteType = NoteType & ", CORP Partners in Progress - $5,000-9,999"
                End If
                If NoteDate = "" Then
                    NoteDate = oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE)
                Else
                    NoteDate = NoteDate & ", " & oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE)
                End If
            End If
            If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttID4 Then
                If NoteType = "" Then
                    NoteType = "CORP Partners in Progress - $10,000-24,999"
                Else
                    NoteType = NoteType & ", CORP Partners in Progress - $10,000-24,999"
                End If
                If NoteDate = "" Then
                    NoteDate = oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE)
                Else
                    NoteDate = NoteDate & ", " & oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE)
                End If
            End If
            If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttID5 Then
            Debug.Print oAttribute.Fields(Attribute_fld_ATTRIBUTES_ID)
                If NoteType = "" Then
                    NoteType = "CORP Partners in Progress - $25,000+"
                Else
                    NoteType = NoteType & ", CORP Partners in Progress - $25,000+"
                End If
                If NoteDate = "" Then
                    NoteDate = oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE)
                Else
                    NoteDate = NoteDate & ", " & oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE)
                End If
            End If
            If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttID6 Then
                If NoteType = "" Then
                    NoteType = "Partners in Progress - $1,000-2,499"
                Else
                    NoteType = NoteType & ", Partners in Progress - $1,000-2,499"
                End If
                If NoteDate = "" Then
                    NoteDate = oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE)
                Else
                    NoteDate = NoteDate & ", " & oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE)
                End If
            End If
            If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttID7 Then
                If NoteType = "" Then
                    NoteType = "Partners in Progress - $2,500-4,999"
                Else
                    NoteType = NoteType & ", Partners in Progress - $2,500-4,999"
                End If
                If NoteDate = "" Then
                    NoteDate = oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE)
                Else
                    NoteDate = NoteDate & ", " & oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE)
                End If
            End If
            If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttID8 Then
                If NoteType = "" Then
                    NoteType = "Partners in Progress - $5,000-9,999"
                Else
                    NoteType = NoteType & ", Partners in Progress - $5,000-9,999"
                End If
                If NoteDate = "" Then
                    NoteDate = oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE)
                Else
                    NoteDate = NoteDate & ", " & oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE)
                End If
            End If
            If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttID9 Then
                If NoteType = "" Then
                    NoteType = "Partners in Progress - $10,000-24,999"
                Else
                    NoteType = NoteType & ", Partners in Progress - $10,000-24,999"
                End If
                If NoteDate = "" Then
                    NoteDate = oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE)
                Else
                    NoteDate = NoteDate & ", " & oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE)
                End If
            End If
            If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttID10 Then
                If NoteType = "" Then
                    NoteType = "Partners in Progress - $25,000+"
                Else
                    NoteType = NoteType & ", CPartners in Progress - $25,000+"
                End If
                If NoteDate = "" Then
                    NoteDate = oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE)
                Else
                    NoteDate = NoteDate & ", " & oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE)
                End If
            End If
        Next oAttribute
        
        oRow.Field("PiP Membership") = NoteType
        oRow.Field("PiP Date") = NoteDate
        oRow.Field("Development Solicitor") = DevSolicitor
        oRow.Field("Primary Relationship Manager") = PrimaryRelManager
        oRow.Field("Secondary Relationship Manager") = SSecondaryRelManager
              
        Set oASolicitor = Nothing
        Set oSolID = Nothing
        
        oCons.Closedown
        Set oCons = Nothing
        Set oService = Nothing
        Set oAttributeServer = Nothing
        Set oAttribute = Nothing
'
    End If
End Sub

'
' One time scipt to update the old pc_connection
'
Public Sub UpdatePC_Connection(oRow As IBBQueryRow)
    Const ConsID = 1

    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
    
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init REApplication.SessionContext
        
        Dim oID As Long
        
        oID = oRow.Field(ConsID)
        oCons.Load oID
'
        Dim oService As REServices
        Set oService = New REServices
        oService.Init REApplication.SessionContext
           
        Dim oAttributeServer As CAttributeTypeServer
        Set oAttributeServer = oService.CreateServiceObject(bbsoAttributeTypeServer)
        oAttributeServer.Init REApplication.SessionContext
        
        Dim lAttributeID As Long
        
        lAttributeID = oAttributeServer.GetAttributeTypeID("PC_Connection", bbAttributeRecordType_CONSTITUENT)
        
        Set oService = Nothing
        Set oAttributeServer = Nothing

        Dim New_PC_Connection As String
        Dim AddIt As Boolean
        
        New_PC_Connection = oRow.Field("Comments")
        If New_PC_Connection = "Healthcare professional in the pancreatic cancer f" Then
            New_PC_Connection = "Healthcare professional in the pancreatic cancer field"
        End If
        
        AddIt = True
        Dim oAttribute As IBBAttribute
        
        For Each oAttribute In oCons.Attributes
            If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttributeID Then
                oAttribute.Fields(Attribute_fld_VALUE) = Trim(New_PC_Connection)
                AddIt = False
            End If
        Next oAttribute
        If AddIt = True Then
            Set oAttribute = oCons.Attributes.Add
            With oCons
                oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttributeID
                oAttribute.Fields(Attribute_fld_VALUE) = Trim(New_PC_Connection)
            End With
        End If
        
'        On Error Resume Next
        oCons.Save
        
'
        Set oAttribute = Nothing
        oCons.Closedown
        Set oCons = Nothing
    End If

End Sub
Public Sub Fix_OffLine_TeamRaiser(oRow As IBBQueryRow)
    Const Gift_ID = 1
    
    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
        Dim oGift As CGift
        Set oGift = New CGift
        oGift.Init RE7.SessionContext
                
        Dim lID As Long
        'get the id
        lID = oRow.Field(Gift_ID)
        
        oGift.Load lID
        With oGift
            .Fields(GIFT_fld_Acknowledge_Flag) = "Not Acknowledged"
            .Fields(GIFT_fld_Acknowledge_Date) = ""
            .Fields(GIFT_fld_Letter_Code) = "CO Ack Letter"
            .Fields(GIFT_fld_Receipt_Flag) = "Receipted"
            On Error Resume Next
            .Save
        End With
          

        oGift.Closedown
        
        Set oGift = Nothing
    End If

End Sub
Public Sub Query_AddInterestCode(oRow As IBBQueryRow)
    Const Const_ID = 1
    
    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init RE7.SessionContext
        
        Dim lID As Long
        'get the id
        lID = oRow.Field(Const_ID)
        oCons.Load lID
        
        Dim oProspect As CProspect
        Set oProspect = oCons.Prospect

        With oProspect.Interests.Add
            .Fields(PHILANTHROPY_fld_INTEREST_CODE) = "Heart"
        End With
        
        On Error Resume Next
        oCons.Save
        
        oCons.Closedown
        Set oCons = Nothing
    End If
    
End Sub

Public Sub End_Solicitor_Rel(oRow As IBBQueryRow)
    
    Dim Const_ID As Long
    
    Dim NoNextAction As Boolean
    
    Const_ID = 1
    
    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init RE7.SessionContext
        
        Dim lID As Long
        Dim sID As Long
        Dim sType As String
        Dim NewSol As Boolean
        
        lID = oRow.Field(Const_ID)
        sID = oRow.Field(3)
        sType = oRow.Field(4)
        oCons.Load lID
        
        Debug.Print oCons.Fields(RECORDS_fld_PRIMARY_ADDRESSEE)

        Dim oASolicitor As CAssignedSolicitor
        
        NewSol = False
        
        sType = "Secondary Relationship Manager"
        
        'loop through the constituent's assigned solicitors
        For Each oASolicitor In oCons.Relations.AssignedSolicitors
'            If oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_ID) = sID And _
'                oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) = sType Then
'                If oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_DATE_TO) = "" Then
'                    Debug.Print "     " & oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_ID) & " " & oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_NAME) & " " & sID
'                    With oCons
''                        oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_DATE_TO) = Date
''                        .Save
'                    End With
'                End If
'            End If
            
            'Check if the new develoment solicitor is already assigned
            If oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_ID) = 660634 And _
                oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) = sType And _
                oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_DATE_TO) = "" Then
                    Debug.Print "     " & oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_ID) & " " & oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_NAME) & " " & sID
                    NewSol = True
            End If
        Next oASolicitor
        
        Set oASolicitor = Nothing
        
        'If the new development solicitor is already assigned do not add.
        If NewSol = False Then
            Dim oSol As CAssignedSolicitor2
            Set oSol = New CAssignedSolicitor2
            
'            With oSol
'                .Init REApplication.SessionContext
'                .Fields(ASSIGNEDSOLICITOR2_fld_CONSTIT_ID) = lID
'                .Fields(ASSIGNEDSOLICITOR2_fld_SOLICITOR_ID) = 660634
'                .Fields(ASSIGNEDSOLICITOR2_fld_SOLICITOR_TYPE) = sType
'                .Fields(ASSIGNEDSOLICITOR2_fld_DATE_FROM) = Date
'                .Fields(ASSIGNEDSOLICITOR2_fld_NOTES) = "Solicitor reassignment - from D Manross.  Requested by Alex "
'                .Save
'                .Closedown
'            End With
            
            Set oSol = Nothing
        End If
        
'        Dim oProspect As CProspect
'        Set oProspect = oCons.Prospect
'
'        With oCons
'            oCons.Prospect.Fields(PROSPECT_fld_STATUS) = "8 - Archive - Temporary"
'            .Save
'        End With
'
'        Set oProspect = Nothing
        
        oCons.Closedown
        Set oCons = Nothing
    
    End If
End Sub

Public Sub Reassign_Solicitor_Rel(oRow As IBBQueryRow)
    
    Dim Const_ID As Long
    
    Dim NoNextAction As Boolean
    
    Const_ID = 1
    
    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init RE7.SessionContext
        
        Dim lID As Long
        Dim sID As Long
        Dim sType As String
        lID = oRow.Field(Const_ID)
        sID = oRow.Field(3)
        sType = oRow.Field(4)
        oCons.Load lID
        
        Debug.Print oCons.Fields(RECORDS_fld_PRIMARY_ADDRESSEE)
      
'
'  The code below to reassigned the prospects
'

        Dim sID2 As Long
        sID2 = 1583473
        
        Dim oSol As CAssignedSolicitor2
        Set oSol = New CAssignedSolicitor2
        
        With oSol
            .Init REApplication.SessionContext
            .Fields(ASSIGNEDSOLICITOR2_fld_CONSTIT_ID) = lID
            .Fields(ASSIGNEDSOLICITOR2_fld_SOLICITOR_ID) = sID2
            .Fields(ASSIGNEDSOLICITOR2_fld_SOLICITOR_TYPE) = sType
            .Fields(ASSIGNEDSOLICITOR2_fld_DATE_FROM) = "10/24/2017"
            .Save
        End With
        
        Set oSol = Nothing

        oCons.Closedown
        Set oCons = Nothing
    
    End If
End Sub


Public Sub Query_ConstributionsBreakdown(oRow As IBBQueryRow)
    Const Gift_ID = 1
        
    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
    
        Set REService = New REServices
        REService.Init REApplication.SessionContext
        
        Set moCodeTablesServer = REService.CreateServiceObject(bbsoCodeTablesServer)
        moCodeTablesServer.Init REApplication.SessionContext
        
        Set moAttributetypeserver = REService.CreateServiceObject(bbsoAttributeTypeServer)
        moAttributetypeserver.Init REApplication.SessionContext
        
        Dim IDDCCR As Long
        Dim PharmaID As Long
        
        IDDCCR = moAttributetypeserver.GetAttributeTypeID("DCCR", bbAttributeRecordType_CONSTITUENT)
        PharmaID = moAttributetypeserver.GetAttributeTypeID("Industry", bbAttributeRecordType_CONSTITUENT)
        
        REService.Closedown
        Set REService = Nothing
        Set moCodeTablesServer = Nothing
        Set moAttributetypeserver = Nothing
    
        Dim Pharma As Boolean
        Dim Foundation As Boolean
        Dim PassThrough As Boolean
        Dim RevenueCategory As String
        Dim Pharma2 As Boolean
        Dim ProgRev As Boolean
        
        Pharma = False
        Pharma2 = False
        
        Foundation = False
        PassThrough = False
        RevenueCategory = ""
        ProgRev = False
        
        If oRow.Field("Program Revenue") = "Yes" Then
            ProgRev = True
        End If
        
        Dim oGift As CGift
        Set oGift = New CGift
        oGift.Init RE7.SessionContext
        
        Dim lID As Long
        lID = oRow.Field(Gift_ID)
        oGift.Load lID
        
        Dim ConsID As String
        ConsID = oRow.Field("Constituent ID")
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init RE7.SessionContext

        oCons.LoadByField uf_Record_CONSTITUENT_ID, ConsID
'
        Debug.Print oCons.Fields(RECORDS_fld_CONSTITUENT_ID) & " " & oCons.Fields(RECORDS_fld_FULL_NAME)
'
        Dim oCCode As CConstituentCode
        
        For Each oCCode In oCons.ConstituentCodes
            Debug.Print oCCode.Fields(CONSTITUENT_CODE_fld_LONG_DESC_READONLY)
            If (oCCode.Fields(CONSTITUENT_CODE_fld_LONG_DESC_READONLY) = "Family Foundation") Then
                Foundation = True
            End If
            If (oCCode.Fields(CONSTITUENT_CODE_fld_LONG_DESC_READONLY) = "Foundation-Family") Then
                Foundation = True
            End If
        Next oCCode
        
        Set oCCode = Nothing
'
        Dim oAttribute As IBBAttribute
        
        For Each oAttribute In oCons.Attributes
            If (oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = IDDCCR) Then
                If (oAttribute.Fields(Attribute_fld_VALUE) = "Industry Champion (Pharma)") Then
                    Pharma2 = True
                End If
                If (oAttribute.Fields(Attribute_fld_VALUE) = "Pass-Through Organization") Then
                    PassThrough = True
                End If
                If (oAttribute.Fields(Attribute_fld_VALUE) = "Non-Corporate Organization (Ind)") Then
                    PassThrough = True
                End If
            End If
            If (oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = PharmaID) Then
                If (oAttribute.Fields(Attribute_fld_VALUE) = "Industry Company") Then
                    Pharma = True
                End If
            End If
        Next oAttribute
        
        Set oAttribute = Nothing
        
        If ProgRev = False Then
''
'' DONOR SPECIFIC GIFTS:
''
    '' Monthly Giving
            If RevenueCategory = "" Then
                If (oGift.Fields(GIFT_fld_Amount) > 0) And ((oRow.Field("Appeal ID") = "MonthlyGiving") Or (oGift.Fields(GIFT_fld_Type) = "Recurring Gift Pay-Cash")) Then
                    RevenueCategory = "Monthly giving"
                End If
            End If
    
    '' Events Sponsorships
            If RevenueCategory = "" Then
                If (oGift.Fields(GIFT_fld_Amount) > 0) And _
                (oGift.Fields(GIFT_fld_Type) <> "Recurring Gift Pay-Cash") And (oGift.Fields(GIFT_fld_Type) <> "Gift-in-Kind") And _
                (oRow.Field("Campaign ID") = "CO EVENTS") And _
                (oRow.Field("Appeal ID") <> "MonthlyGiving") And _
                ((oRow.Field("LO Sponsor Level") <> "") Or (oRow.Field("Sphere Sponsor Description") <> "") Or (oRow.Field("Sphere Sponsor Level Desc") <> "")) Then
                    RevenueCategory = "Events Sponsorships"
                End If
            End If
    '' Events Gift-in-Kind
            If RevenueCategory = "" Then
                If (oGift.Fields(GIFT_fld_Amount) > 0) And (oRow.Field("Campaign ID") = "CO EVENTS") And _
                (oRow.Field("Appeal ID") <> "MonthlyGiving") And (oGift.Fields(GIFT_fld_Type) = "Gift-in-Kind") Then
                    RevenueCategory = "Events Gift-in-Kind"
                End If
            End If
                  
    '' Wage Hope My Way
            If RevenueCategory = "" Then
                If (oGift.Fields(GIFT_fld_Amount) > 0) And _
                (oGift.Fields(GIFT_fld_Type) <> "Recurring Gift Pay-Cash") And _
                (oRow.Field("Fund ID") <> "THPRSC") And (oRow.Field("Fund ID") <> "THPRSC2") And _
                (oRow.Field("Appeal ID") <> "MonthlyGiving") And _
                ((oRow.Field("Campaign ID") = "Community Ambassador") Or (oRow.Field("Appeal Category") = "Keep the Memory Alive")) Then
                    RevenueCategory = "Wage Hope My Way"
                End If
            End If
        
    '' Campaigns
            If RevenueCategory = "" Then
                If (oGift.Fields(GIFT_fld_Amount) > 0) And (oRow.Field("Appeal Category") <> "Keep the Memory Alive") And _
                (oGift.Fields(GIFT_fld_Type) <> "Recurring Gift Pay-Cash") And _
                (oRow.Field("Campaign ID") <> "CO EVENTS") And (oRow.Field("Campaign ID") <> "Community Ambassador") And _
                (oRow.Field("Appeal ID") <> "MonthlyGiving") And _
                ((oRow.Field("Campaign ID") = "DIRECT MKTG") Or (oRow.Field("Appeal Category") = "Direct Mail Campaign") Or (oRow.Field("Appeal Category") = "Newsletter")) Then
                    RevenueCategory = "Campaign"
                End If
            End If
            
    '' Cause Marketing
            If RevenueCategory = "" Then
                If (oGift.Fields(GIFT_fld_Amount) > 0) And (oRow.Field("Appeal ID") <> "MonthlyGiving") And _
                (oGift.Fields(GIFT_fld_Type) <> "Recurring Gift Pay-Cash") And _
                (oRow.Field("Appeal Category") <> "Direct Mail Campaign") And (oRow.Field("Appeal Category") <> "Newsletter") And (oRow.Field("Appeal Category") <> "Keep the Memory Alive") And _
                ((oRow.Field("Campaign ID") = "CAUSE") Or (oRow.Field("Fund ID") = "THPRSC") Or (oRow.Field("Fund ID") = "THPRSC2")) Then
                    RevenueCategory = "Cause Marketing"
                End If
            End If
            
'' UNATTACHED GIFTS:
''
    '' Biopharmaceutical
            If RevenueCategory = "" Then
                If (oGift.Fields(GIFT_fld_Amount) > 0) And (oRow.Field("Appeal ID") <> "MonthlyGiving") And _
                (oGift.Fields(GIFT_fld_Type) <> "Recurring Gift Pay-Cash") And _
                (oRow.Field("Appeal Category") <> "Direct Mail Campaign") And (oRow.Field("Appeal Category") <> "Newsletter") And (oRow.Field("Appeal Category") <> "Keep the Memory Alive") And _
                (oRow.Field("Fund ID") <> "THPRSC") And (oRow.Field("Fund ID") <> "THPRSC2") And _
                (oRow.Field("Campaign ID") <> "CAUSE") And (oRow.Field("Campaign ID") <> "CO EVENTS") And (oRow.Field("Campaign ID") <> "Community Ambassador") And (oRow.Field("Campaign ID") <> "DIRECT MKTG") And Pharma = True Then
                    RevenueCategory = "Biopharmaceutical"
                End If
            End If
            
    '' Planned Giving
            If RevenueCategory = "" Then
                If (oGift.Fields(GIFT_fld_Amount) > 0) And (oGift.Fields(GIFT_fld_Type) <> "Recurring Gift Pay-Cash") And _
                (oRow.Field("Appeal Category") <> "Direct Mail Campaign") And (oRow.Field("Appeal Category") <> "Newsletter") And (oRow.Field("Appeal Category") <> "Keep the Memory Alive") And _
                (oRow.Field("Campaign ID") <> "CAUSE") And (oRow.Field("Campaign ID") <> "CO EVENTS") And (oRow.Field("Campaign ID") <> "Community Ambassador") And (oRow.Field("Campaign ID") <> "DIRECT MKTG") And _
                (oRow.Field("Fund ID") <> "THPRSC") And (oRow.Field("Fund ID") <> "THPRSC2") And (Pharma = False) And (oRow.Field("Appeal ID") = "PlannedGiving") Then
                    RevenueCategory = "Planned giving"
                End If
            End If
            
    '' Workplace Giving
            If RevenueCategory = "" Then
                If (oGift.Fields(GIFT_fld_Amount) > 0) And _
                (oGift.Fields(GIFT_fld_Type) <> "Recurring Gift Pay-Cash") And _
                (oRow.Field("Appeal Category") <> "Direct Mail Campaign") And (oRow.Field("Appeal Category") <> "Newsletter") And (oRow.Field("Appeal Category") <> "Keep the Memory Alive") And _
                (oRow.Field("Fund ID") <> "THPRSC") And (oRow.Field("Fund ID") <> "THPRSC2") And _
                (oRow.Field("Campaign ID") <> "CAUSE") And (oRow.Field("Campaign ID") <> "CO EVENTS") And (oRow.Field("Campaign ID") <> "Community Ambassador") And _
                (oRow.Field("Campaign ID") <> "DIRECT MKTG") And Pharma = False And _
                ((oRow.Field("Appeal ID") = "EmployeeGvng") Or (oRow.Field("Appeal ID") = "CHCNationlEG") Or (oRow.Field("Appeal ID") = "CHCPrivateEG")) Then
                    RevenueCategory = "Workplace giving"
                End If
            End If
    '' Other Corp Gifts
            If RevenueCategory = "" Then
                If (oGift.Fields(GIFT_fld_Amount) > 0) And _
                (oGift.Fields(GIFT_fld_Type) <> "Recurring Gift Pay-Cash") And _
                (oRow.Field("Appeal Category") <> "Direct Mail Campaign") And (oRow.Field("Appeal Category") <> "Newsletter") And (oRow.Field("Appeal Category") <> "Keep the Memory Alive") And _
                (oRow.Field("Campaign ID") <> "CAUSE") And (oRow.Field("Campaign ID") <> "CO EVENTS") And (oRow.Field("Campaign ID") <> "Community Ambassador") And _
                (oRow.Field("Campaign ID") <> "DIRECT MKTG") And _
                (oRow.Field("Appeal ID") <> "EmployeeGvng") And (oRow.Field("Appeal ID") <> "CHCNationlEG") And (oRow.Field("Appeal ID") <> "CHCPrivateEG") And (oRow.Field("Appeal ID") <> "MonthlyGiving") And (oRow.Field("Appeal ID") <> "PlannedGiving") And _
                (oRow.Field("Fund ID") <> "THPRSC") And (oRow.Field("Fund ID") <> "THPRSC2") And _
                (oRow.Field("Key Indicator") = "Organization") And PassThrough = False And Foundation = False Then
                    RevenueCategory = "Other Corp Gifts"
                End If
            End If
            
    '' Major Gifts
            If ((oGift.Fields(GIFT_fld_Amount) > 9999.99) And (oRow.Field("Appeal Category") <> "Direct Mail Campaign") And (oGift.Fields(GIFT_fld_Type) <> "Recurring Gift Pay-Cash") And _
                (oRow.Field("Appeal Category") <> "Newsletter") And (oRow.Field("Appeal Category") <> "Keep the Memory Alive") And _
                (oRow.Field("Campaign ID") <> "CAUSE") And (oRow.Field("Campaign ID") <> "CO EVENTS") And (oRow.Field("Campaign ID") <> "Community Ambassador") And _
                (oRow.Field("Campaign ID") <> "DIRECT MKTG") And (oRow.Field("Appeal ID") <> "EmployeeGvng") And (oRow.Field("Appeal ID") <> "CHCNationlEG") And _
                (oRow.Field("Appeal ID") <> "CHCPrivateEG") And (oRow.Field("Appeal ID") <> "MonthlyGiving") And (oRow.Field("Appeal ID") <> "PlannedGiving")) And _
                ((oRow.Field("Key Indicator") = "Individual") Or PassThrough = True Or Foundation = True) Then
                RevenueCategory = "Major gift"
            End If
    '' Annual Giving
            If RevenueCategory = "" Then
                If (oGift.Fields(GIFT_fld_Amount) < 10000) And (oGift.Fields(GIFT_fld_Type) <> "Recurring Gift Pay-Cash") And _
                (oRow.Field("Appeal Category") <> "Direct Mail Campaign") And (oRow.Field("Appeal Category") <> "Newsletter") And (oRow.Field("Appeal Category") <> "Keep the Memory Alive") And _
                ((oRow.Field("Campaign ID") = "ANN") Or (oRow.Field("Campaign ID") = "EWTS") Or (oRow.Field("Campaign ID") = "MajorGiftSolicitors") Or (oRow.Field("Campaign ID") = "Raise the Cure") Or (oRow.Field("Campaign ID") = "WORKSHOPS")) And _
                (oRow.Field("Appeal ID") <> "EmployeeGvng") And (oRow.Field("Appeal ID") <> "CHCNationlEG") And (oRow.Field("Appeal ID") <> "CHCPrivateEG") And (oRow.Field("Appeal ID") <> "MonthlyGiving") And (oRow.Field("Appeal ID") <> "PlannedGiving") And _
                ((oRow.Field("Key Indicator") = "Individual") Or PassThrough = True Or Foundation = True) Then
                    RevenueCategory = "Annual giving"
                End If
            End If
    
        End If

''
        oRow.Field("Pass-Through") = PassThrough
        oRow.Field("Pharma") = Pharma
        oRow.Field("Foundation") = Foundation
        oRow.Field("Revenue Category") = RevenueCategory
        oRow.Field("Pharma 2") = Pharma2
'
        
        oGift.Closedown
        Set oGift = Nothing
        oCons.Closedown
        Set oCons = Nothing

    End If
End Sub
Public Sub Query_PrimaryDonor(oRow As IBBQueryRow)
        
    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
    
        Dim oGift As CGift
        Set oGift = New CGift
        oGift.Init RE7.SessionContext
                
        Dim lID As String
        'get the id
        lID = oRow.Field("Gift Import ID")
        
'        Debug.Print oRow.Field("Name") & " " & oRow.Field("Gift ID"), oRow.Field("Gift Date")
        
        oGift.LoadByField gufIMPORT_ID, lID
        
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init RE7.SessionContext
        
        Dim lid2 As String
        lid2 = oGift.Fields(GIFT_fld_Constit_ID)
        
        oCons.Load lid2
        
'        Debug.Print oRow.Field("Specific Record") & "   " & oGift.Fields(GIFT_fld_Constit_ID)
        
        Dim PID As Long
        PID = oRow.Field("Specific Record")
        
        If PID <> oGift.Fields(GIFT_fld_Constit_ID) Then
        
            oRow.Field("Primary Donor ID") = oGift.Fields(GIFT_fld_Constit_ID)
            oRow.Field("Primary Donor Name") = oGift.Fields(GIFT_fld_Constituent_Name)
            
            If oCons.Fields(RECORDS_fld_KEY_INDICATOR) = 2 Then
                oRow.Field("Primary Donor Indicator") = "Individual"
            Else
                oRow.Field("Primary Donor Indicator") = "Organization"
            End If
            
        End If

        oCons.Closedown
        Set oCons = Nothing

        oGift.Closedown
        Set oGift = Nothing
        
    End If

End Sub

Public Sub CleanHCP(oRow As IBBQueryRow)
Const ConsID = 1

    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
     
        Dim oService As REServices
        Set oService = New REServices
        oService.Init REApplication.SessionContext
        
        Dim oAttributeServer As CAttributeTypeServer
        Set oAttributeServer = oService.CreateServiceObject(bbsoAttributeTypeServer)
        oAttributeServer.Init REApplication.SessionContext
        
        Dim HCPID As Long
        
        HCPID = oAttributeServer.GetAttributeTypeID("Healthcare Professional", bbAttributeRecordType_CONSTITUENT)
         
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init REApplication.SessionContext
        
        Dim oID As Long
        
        oID = oRow.Field(ConsID)
        oCons.Load oID
        
        Dim oAttribute As IBBAttribute
        
        For Each oAttribute In oCons.Attributes
            Debug.Print oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) & " " & lAttributeID
            If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = HCPID Then
                oCons.Attributes.Remove oAttribute
            End If
        Next oAttribute
        
'        On Error Resume Next
'        oCons.Save
        
        Dim oCCode As CConstituentCode
        
        Dim FoundIt As Boolean
        FoundIt = False
        
        For Each oCCode In oCons.ConstituentCodes
            If (oCCode.Fields(CONSTITUENT_CODE_fld_LONG_DESC_READONLY) = "Healthcare Professional") Then
                FoundIt = True
                Debug.Print "Found it: " & oCCode.Fields(CONSTITUENT_CODE_fld_CODE)
                
                With oCCode
                    .Fields(CONSTITUENT_CODE_fld_CODE) = "General Constituent"
                End With
'                On Error Resume Next
'                oCons.Save
            End If
        Next oCCode
        
        Set oCCode = Nothing
        oCons.Closedown
        
        oService.Closedown
        Set oService = Nothing
        Set oAttributeServer = Nothing
        Set oAttribute = Nothing
        Set oCons = Nothing
    End If

End Sub
Public Sub RemoveAttribute(oRow As IBBQueryRow)
Const ConsID = 1

    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
     
        Dim oService As REServices
        Set oService = New REServices
        oService.Init REApplication.SessionContext
        
        Dim oAttributeServer As CAttributeTypeServer
        Set oAttributeServer = oService.CreateServiceObject(bbsoAttributeTypeServer)
        oAttributeServer.Init REApplication.SessionContext
        
        Dim AttrID As Long
        
        AttrID = oAttributeServer.GetAttributeTypeID("DCCR", bbAttributeRecordType_CONSTITUENT)
         
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init REApplication.SessionContext
        
        Dim oID As Long
        
        oID = oRow.Field(ConsID)
        oCons.Load oID
        
        Dim oAttribute As IBBAttribute
        
        For Each oAttribute In oCons.Attributes
            Debug.Print oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) & " " & lAttributeID
            If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = AttrID And _
                oAttribute.Fields(Attribute_fld_VALUE) = "Industry Champion (Pharma)" Then
                oCons.Attributes.Remove oAttribute
            End If
        Next oAttribute
        
'        On Error Resume Next
'        oCons.Save
        
        oService.Closedown
        Set oService = Nothing
        Set oAttributeServer = Nothing
        Set oAttribute = Nothing
        Set oCons = Nothing
    End If

End Sub
Public Sub AddOriginCode(oRow As IBBQueryRow)
    Const ConsID = 1

    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
        
        Dim oService As REServices
        Set oService = New REServices
        oService.Init REApplication.SessionContext
        
        Dim oAttributeServer As CAttributeTypeServer
        Set oAttributeServer = oService.CreateServiceObject(bbsoAttributeTypeServer)
        oAttributeServer.Init REApplication.SessionContext
        
        Dim ResearcherID As Long
        Dim CTFID As Long
        Dim GranteeID As Long
        Dim SMABID As Long
        Dim SourceGen As Long
        
        ResearcherID = oAttributeServer.GetAttributeTypeID("Researcher", bbAttributeRecordType_CONSTITUENT)
        CTFID = oAttributeServer.GetAttributeTypeID("Clinical Trial Finder Account", bbAttributeRecordType_CONSTITUENT)
        GranteeID = oAttributeServer.GetAttributeTypeID("Grant Recipient", bbAttributeRecordType_CONSTITUENT)
        SMABID = oAttributeServer.GetAttributeTypeID("SMAB Committee", bbAttributeRecordType_CONSTITUENT)
        SourceGen = oAttributeServer.GetAttributeTypeID("Source/General", bbAttributeRecordType_CONSTITUENT)

'        Set oAttributeServer = Nothing
'
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init REApplication.SessionContext
        
        Dim oID As Long
        
        oID = oRow.Field(ConsID)
        oCons.Load oID
        
        Debug.Print oCons.Fields(RECORDS_fld_FULL_NAME)
        
        Dim Origin_Code As String
        Origin_Code = "General Public"
        
        Dim FGift As Date
        Dim FAction As Date
        Dim GenSource As Date
        Dim HCPSource As Date
        Dim FEvent As Date
        Dim FAppeal As Date
        Dim FNote As Date
        Dim FirstDate As Date
        Dim FirstEventID As String
        
        FirstDate = Date
        FGift = "01/01/3000"
        FAction = "01/01/3000"
        GenSource = "01/01/3000"
        HCPSource = "01/01/3000"
        FEvent = "01/01/3000"
        FAppeal = "01/01/3000"
        FNote = "01/01/3000"

        
        If oRow.Field("First Gift Date") <> "" Then FGift = oRow.Field("First Gift Date")
        If oRow.Field("First Action Date") <> "" Then FAction = oRow.Field("First Action Date")
        If oRow.Field("GenSource Date") <> "" Then GenSource = oRow.Field("GenSource Date")
        If oRow.Field("HCP Source Date") <> "" Then HCPSource = oRow.Field("HCPSource Date")
        
        Debug.Print FGift
        Debug.Print FAction
        Debug.Print GenSource
        Debug.Print HCPSource
'
'   Check the constituent code
'
        Dim oCCode As CConstituentCode
        
        For Each oCCode In oCons.ConstituentCodes
            If oCCode.Fields(CONSTITUENT_CODE_fld_LONG_DESC_READONLY) = "Scientific & Medical Advisory Board" Then
                If FirstDate > oCCode.Fields(CONSTITUENT_CODE_fld_DATE_FROM) Then
                    FirstDate = oCCode.Fields(CONSTITUENT_CODE_fld_DATE_FROM)
                    Origin_Code = "Researcher"
                End If
            End If
            If oCCode.Fields(CONSTITUENT_CODE_fld_LONG_DESC_READONLY) = "Key Volunteer Position" Then
                If FirstDate > oCCode.Fields(CONSTITUENT_CODE_fld_DATE_FROM) Then
                    FirstDate = oCCode.Fields(CONSTITUENT_CODE_fld_DATE_FROM)
                    Origin_Code = "Volunteer"
                End If
            End If
        Next oCCode

        Set oCCode = Nothing
'
'
'
        Dim oAlias As CAlias
        Dim oNote As IBBNotepad
        
        Dim AliasName As String
        Dim PTAlias As Boolean
         
        PTAlias = False
        AliasName = ""
        
        For Each oAlias In oCons.Aliases
            If oAlias.Fields(ALIAS_fld_ALIAS_TYPE) = "Patient Central" Then
                PTAlias = True
                AliasName = oAlias.Fields(ALIAS_fld_KEY_NAME)
                Exit For
            End If
        Next oAlias

'        Debug.Print "Alias: " & AliasName
        
        If PTAlias = True Then
            For Each oNote In oCons.Notepads
                If UCase(oNote.Fields(NOTEPAD_fld_Author)) = UCase(AliasName) Then
                    If FirstDate < oNote.Fields(NOTEPAD_fld_DateAdded) Then
                        FirstDate = oNote.Fields(NOTEPAD_fld_NotepadDate)
                        Origin_Code = "Patient Services"
                    Debug.Print "Found it..." & oNote.Fields(NOTEPAD_fld_NotepadDate)
                    End If
                End If
            Next oNote
        End If

        Set oNote = Nothing
        Set oAlias = Nothing
'
'   Check the constituent attribute
'
        Dim oAttribute As IBBAttribute
        
        For Each oAttribute In oCons.Attributes
            If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = CTFID And _
                oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE) <> "" And _
                oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE) < FirstDate Then
                Origin_Code = "Patient Services"
                FirstDate = oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE)
            End If
            If (oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = ResearcherID Or _
                oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = GranteeID Or _
                oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = SMABID) And _
                oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE) <> "" And _
                oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE) < FirstDate Then
                Origin_Code = "Researcher"
                FirstDate = oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE)
            End If
            If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = SourceGen And _
                oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE) <> "" And _
                oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE) < FirstDate Then
                If oAttribute.Fields(Attribute_fld_VALUE) = "FB Fundraiser" Then
                    Origin_Code = "P2P Participant"
                End If
                If oAttribute.Fields(Attribute_fld_VALUE) = "FB Donor" Then
                    Origin_Code = "P2P Donor"
                End If
                If oAttribute.Fields(Attribute_fld_VALUE) = "TempurPedic Rest Test Participants 2016" Then
                    Origin_Code = "Email Submission"
                End If
                FirstDate = oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE)
            End If
        Next oAttribute
'
'   Loop to event participation records and find the event
'
        Dim FirstEventType As String
        
        Dim oParticipant As CParticipant
        
        FirstEventID = ""
        FirstEventType = ""

        For Each oParticipant In oCons.Participants
            With oParticipant
                If oParticipant.Fields(Participants_fld_DateAdded) < FirstDate Then
                    FirstEventID = oParticipant.Fields(Participants_fld_EventID)
                    FirstDate = Format(oParticipant.Fields(Participants_fld_DateAdded), "MM/DD/YYYY")
                    If oParticipant.Fields(Participants_fld_DatePaid) <> "" Then
                        FirstEventType = "Registrant"
                        Debug.Print "Date Paid: " & oParticipant.Fields(Participants_fld_DatePaid)
                    End If
                End If
            End With
        Next
        
        Debug.Print FirstEventID
        
        If FirstEventID <> "" Then
            Dim oEvent As CSpecialEvent
            Set oEvent = New CSpecialEvent
            oEvent.Init REApplication.SessionContext
            
            oEvent.Load FirstEventID
            
            Origin_Code = "Awareness Event"
            
            Debug.Print oEvent.Fields(SPECIAL_EVENT_fld_TYPEID)
            
            If oEvent.Fields(SPECIAL_EVENT_fld_TYPEID) = "PurpleStride" Or _
                oEvent.Fields(SPECIAL_EVENT_fld_TYPEID) = "PurpleRide" Or _
                oEvent.Fields(SPECIAL_EVENT_fld_TYPEID) = "Individual" Or _
                oEvent.Fields(SPECIAL_EVENT_fld_TYPEID) = "PurpleBowl/Link" Or _
                oEvent.Fields(SPECIAL_EVENT_fld_TYPEID) = "Marathon Team" Or _
                oEvent.Fields(SPECIAL_EVENT_fld_TYPEID) = "PurpleSwim" Or _
                oEvent.Fields(SPECIAL_EVENT_fld_TYPEID) = "Walk/Run" Then
                Origin_Code = "P2P Donor"
                If FirstEventType = "Registrant" Then
                    Origin_Code = "P2P Participant"
                End If
            End If
            
            If oEvent.Fields(SPECIAL_EVENT_fld_TYPEID) = "Leadership Breakfast" Or oEvent.Fields(SPECIAL_EVENT_fld_TYPEID) = "DCR Donor Reception" Or _
                oEvent.Fields(SPECIAL_EVENT_fld_TYPEID) = "Reception/Gala" Or oEvent.Fields(SPECIAL_EVENT_fld_TYPEID) = "PanCAN Internal Event" Or _
                oEvent.Fields(SPECIAL_EVENT_fld_TYPEID) = "Leadership Training" Then
                Origin_Code = "Cultivation Event"
            End If
            
            If oEvent.Fields(SPECIAL_EVENT_fld_TYPEID) = "Seminar/Lecture" Or _
                oEvent.Fields(SPECIAL_EVENT_fld_TYPEID) = "Educational Event-Webinar" Then
                Origin_Code = "Patient Services"
            End If
            If oEvent.Fields(SPECIAL_EVENT_fld_TYPEID) = "Awareness Event" Or _
               oEvent.Fields(SPECIAL_EVENT_fld_TYPEID) = "PurpleLight" Then
                Origin_Code = "Awareness Event"
            End If
            
            If InStr(UCase(oEvent.Fields(SPECIAL_EVENT_fld_DESCRIPTION)), "WAGE HOPE MY WAY") > 0 Then
                Origin_Code = "P2P Donor"
                If FirstEventType = "Registrant" Then
                    Origin_Code = "P2P Participant"
                End If
            End If

            If InStr(UCase(oEvent.Fields(SPECIAL_EVENT_fld_DESCRIPTION)), "ADVOCACY DAY") > 0 Then
                Origin_Code = "Advocacy"
            End If
        
            oEvent.Closedown
            Set oEvent = Nothing
            
        End If
'
'  Check First Gift
'
        If FirstDate > FGift Then
            FirstDate = FGift
            Origin_Code = "Donor"
            If InStr(UCase(oRow.Field("First Gift Appeal Description")), "STRIDE") > 0 Then
                Origin_Code = "P2P Participant"
                If InStr(UCase(oRow.Field("LO Donation Type")), "GIFT") > 0 Then
                    Origin_Code = "P2P Donor"
                End If
                If InStr(UCase(oRow.Field("Sphere Donation Type")), "DONATION") > 0 Then
                    Origin_Code = "P2P Donor"
                End If
            End If
            If InStr(UCase(oRow.Field("First Gift Appeal Description")), "WAGE HOPE MY WAY") > 0 Or _
                InStr(UCase(oRow.Field("First Gift Appeal Description")), "SWIM") > 0 Or _
                InStr(UCase(oRow.Field("First Gift Appeal Description")), "BOWL") > 0 Then
                Origin_Code = "P2P Participant"
            End If
            If InStr(UCase(oRow.Field("First Gift Appeal Description")), "LIGHT") > 0 Then
                Origin_Code = "Awareness Event"
            End If
            If InStr(UCase(oRow.Field("First Gift Appeal Description")), "ADVOCACY DAY") > 0 Then
                Origin_Code = "Advocacy"
            End If
        End If
        
        If FirstDate > FEvent Then
            FirstDate = FEvent
        End If
        If FirstDate > FAppeal Then
            FirstDate = FAppeal
        End If
'
'  Find first action not that is not origin code.
'
        Dim oFirstAction As CAction
        
        Dim FirstActionDate As Date
        Dim FirstActionType As String
        Dim FirstActionCateg As String
        
        FirstActionDate = Date
        FirstActionType = ""
        FirstActionCateg = ""
        
        For Each oFirstAction In oCons.Actions
            If oFirstAction.Fields(ACTION_fld_TYPE) <> "Origin Code" And _
                oFirstAction.Fields(ACTION_fld_DATE) < FirstActionDate Then
                FirstActionDate = oFirstAction.Fields(ACTION_fld_DATE)
                FirstActionType = oFirstAction.Fields(ACTION_fld_TYPE)
                FirstActionCateg = oFirstAction.Fields(ACTION_fld_CATEGORY)
            End If
        Next oFirstAction
        
'        Debug.Print FirstActionDate
'        Debug.Print FirstActionType
'        Debug.Print FirstActionCateg
'
        If FirstDate > FirstActionDate Then
            If FirstActionCateg = "Advocacy" Then
                Origin_Code = "Advocacy"
                FirstDate = FirstActionDate
            End If
            If InStr(UCase(FirstActionType), "PALS") > 0 Or _
                InStr(UCase(FirstActionType), "PTS-INT") > 0 Or _
                InStr(UCase(FirstActionType), "PACKET") > 0 Then
                Origin_Code = "Patient Services"
                FirstDate = FirstActionDate
            End If
            If InStr(UCase(FirstActionType), "FIRSTGIVING") > 0 Then
                Origin_Code = "P2P Participant"
                FirstDate = FirstActionDate
            End If
            If InStr(UCase(FirstActionType), "TEAM HOPE PAGE CREATOR") > 0 Or _
                InStr(UCase(FirstActionType), "TEAM HOPE PAGE CREATION") > 0 Or _
                InStr(UCase(FirstActionType), "WAGE HOPE SIGN-UP") > 0 Then
                Origin_Code = "P2P Participant"
                FirstDate = FirstActionDate
            End If
            If InStr(UCase(FirstActionType), "RSA") > 0 Then
                Origin_Code = "Researcher"
                FirstDate = FirstActionDate
            End If
        End If
        
        Set oFirstAction = Nothing
'
'        If FirstDate > FAction Then
'            If oRow.Field("First Action Category") = "Advocacy" Then
'                Origin_Code = "Advocacy"
'                FirstDate = FAction
'            End If
'            If InStr(UCase(oRow.Field("First Action Type")), "PALS") > 0 Or _
'                InStr(UCase(oRow.Field("First Action Type")), "PTS-INT") > 0 Or _
'                InStr(UCase(oRow.Field("First Action Type")), "PACKET") > 0 Then
'                Origin_Code = "Patient Services"
'                FirstDate = FAction
'            End If
'            If InStr(UCase(oRow.Field("First Action Type")), "FIRSTGIVING") > 0 Then
'                Origin_Code = "P2P Participant"
'                FirstDate = FAction
'            End If
'            If InStr(UCase(oRow.Field("First Action Type")), "TEAM HOPE PAGE CREATOR") > 0 Then
'                Origin_Code = "P2P Participant"
'                FirstDate = FAction
'            End If
'            If InStr(UCase(oRow.Field("First Action Type")), "RSA") > 0 Then
'                Origin_Code = "Researcher"
'                FirstDate = FAction
'            End If
'        End If
'

Debug.Print "First Date: " & FirstDate & " GenSource: " & GenSource

        If FirstDate = GenSource Then
            If (oRow.Field("GenSource Desc") = "PT Registry") Then
                Origin_Code = "Patient Services"
                FirstDate = GenSource
            End If
            If (oRow.Field("GenSource Desc") = "FB Fundraiser") Then
                Origin_Code = "P2P Participant"
                FirstDate = GenSource
            End If
'            If (oRow.Field("GenSource Desc") = "Social Media Lead Ad") Then
'                Origin_Code = "P2P Participant"
'                FirstDate = GenSource
'            End If
            If (oRow.Field("GenSource Desc") = "FB Doonor") Then
                Origin_Code = "P2P Donor"
                FirstDate = GenSource
            End If
            If (oRow.Field("GenSource Desc") = "Clinical Trial Finder User") Then
                Origin_Code = "Patient Services"
                FirstDate = GenSource
            End If
            If (oRow.Field("GenSource Desc") = "CTF Email Capture") Then
                Origin_Code = "Patient Services"
                FirstDate = GenSource
            End If
            If (oRow.Field("GenSource Desc") = "CTMS") Then
                Origin_Code = "Patient Services"
                FirstDate = GenSource
            End If
        End If
        If FirstDate > HCPSource Then
            If InStr(UCase(oRow.Field("HCP Desc")), "GI ASCO") > 0 Then
                Origin_Source = "Patient Services"
                FirstDate = HCPSource
            End If
        End If
'        If FirstDate > HCPSource Then
'            If InStr(UCase(oRow.Field("HCP Desc")), "Clinical Trial Finder User") > 0 Then
'                Origin_Source = "Patient Services"
'                FirstDate = HCPSource
'            End If
'        End If
'
'        Debug.Print "First Date: " & FirstDate
        
        If FirstDate = Date Then
            FirstDate = Format(oCons.Fields(RECORDS_fld_DATE_ADDED), "MM/DD/YYYY")
        End If
                
'
'
'
        Debug.Print "Origin Code: " & Origin_Code
        Debug.Print "Added by: " & oCons.Fields(RECORDS_fld_ADDED_BY)
        
'
'       Account all constituents added by 19673_relointegrationuser and REImport users
'
        If Origin_Code = "General Public" And oCons.Fields(RECORDS_fld_ADDED_BY) = 351 Or _
            Origin_Code = "General Public" And oCons.Fields(RECORDS_fld_ADDED_BY) = 447 Or _
            Origin_Code = "General Public" And oCons.Fields(RECORDS_fld_ADDED_BY) = 348 Then
            Origin_Code = "Email Submission"
        End If
        If FirstDate = FirstActionDate Then
            FirstDate = FirstDate - 1
        End If
        
        oRow.Field("First Date") = CStr(FirstDate)
        oRow.Field("Origin Code") = Origin_Code
'
' Add or update the Origin Code Action Type
'
        Dim FoundAction As Boolean
        FoundAction = False
        
        Dim oAction As CAction
        
        For Each oAction In oCons.Actions
            If oAction.Fields(ACTION_fld_TYPE) = "Origin Code" Then
                With oAction
                    oAction.Fields(ACTION_fld_TYPE) = "Origin Code"
                    oAction.Fields(ACTION_fld_DATE) = FirstDate
                    oAction.Fields(ACTION_fld_WORD_DOC_NAME) = Origin_Code
                    On Error Resume Next
                    oAction.Save
                End With
                FoundAction = True
            End If
        Next oAction
        
        If FoundAction = False Then
'            Debug.Print "Saving new action..."
            Set oAction = New CAction
            oAction.Init REApplication.SessionContext

            With oAction
                .Fields(ACTION_fld_CATEGORY) = "Task/Other"
                .Fields(ACTION_fld_TYPE) = "Origin Code"
                .Fields(ACTION_fld_DATE) = FirstDate
                .Fields(ACTION_fld_COMPLETED) = True
                .Fields(ACTION_fld_COMPLETED_DATE) = FirstDate
                .Fields(ACTION_fld_RECORDS_ID) = oCons.Fields(RECORDS_fld_ID)
                .Fields(ACTION_fld_WORD_DOC_NAME) = Origin_Code

                On Error Resume Next
                oAction.Save
            End With
        End If
        
        Set oAction = Nothing
            
        CSpecialEvent.Closedown
        Set CSpecialEvent = Nothing
            
        oCons.Closedown
        Set oCons = Nothing

        oService.Closedown
        Set oService = Nothing
        Set oAttributeServer = Nothing
        Set oAttribute = Nothing

    End If

End Sub
Public Sub DeleteDupOriginCode(oRow As IBBQueryRow)
    Const ConsID = 1

    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
        Dim oService As REServices
        Set oService = New REServices
        oService.Init REApplication.SessionContext
            
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init REApplication.SessionContext
        
        Dim oID As Long
        
        oID = oRow.Field(ConsID)
        oCons.Load oID
        
        Dim oAction As CAction
        Dim FirstOriginCodeDate As Date
        Dim FirstOriginType As String
        Dim OriginCodeCounter As Integer
        
        FirstOriginCodeDate = "01/01/9999"
        FirstOriginType = ""
        OriginCodeCounter = 0
        
        For Each oAction In oCons.Actions
            If oAction.Fields(ACTION_fld_TYPE) = "Origin Code" And _
                oAction.Fields(ACTION_fld_DATE) < FirstOriginCodeDate Or _
                oAction.Fields(ACTION_fld_DATE) = FirstOriginCodeDate Then
                    FirstOriginCodeDate = oAction.Fields(ACTION_fld_DATE)
                    If FirstOriginType = "" And oAction.Fields(ACTION_fld_WORD_DOC_NAME) <> "Email Submission" And _
                        oAction.Fields(ACTION_fld_WORD_DOC_NAME) <> "General Public" Then
                        FirstOriginType = oAction.Fields(ACTION_fld_WORD_DOC_NAME)
                    End If
                    Debug.Print "Found: " & oAction.Fields(ACTION_fld_DATE) & " " & FirstOriginType
            End If
        Next oAction
        Set oAction = Nothing

        Dim oAction2 As CAction
        For Each oAction In oCons.Actions
            If oAction.Fields(ACTION_fld_TYPE) = "Origin Code" Then
                Debug.Print oAction.Fields(ACTION_fld_WORD_DOC_NAME)
                If oAction.Fields(ACTION_fld_WORD_DOC_NAME) = FirstOriginType Then
                    OriginCodeCounter = OriginCodeCounter + 1
                    Debug.Print OriginCodeCounter
                End If
                If oAction.Fields(ACTION_fld_DATE) > FirstOriginCodeDate Or OriginCodeCounter > 1 Or _
                   (oAction.Fields(ACTION_fld_DATE) = FirstOriginCodeDate And FirstOriginType <> oAction.Fields(ACTION_fld_WORD_DOC_NAME)) Then
                        With oAction
                            Debug.Print "Deleting: " & oAction.Fields(ACTION_fld_DATE) & " " & oAction.Fields(ACTION_fld_WORD_DOC_NAME)
                            oAction.Delete
                            On Error Resume Next
                            oCons.Save
                            Exit For
                        End With
                End If
            End If
        Next oAction
        Set oAction2 = Nothing

        oCons.Closedown
        Set oCons = Nothing
    End If

End Sub
Public Sub Fix_OffCCGift(oRow As IBBQueryRow)
    Const Gift_ID = 1
    
    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
        Dim oGift As CGift
        Set oGift = New CGift
        oGift.Init RE7.SessionContext
                
        Dim lID As Long
        'get the id
        lID = oRow.Field(Gift_ID)
        
        Dim CCNum As String
        Dim CCType As String
        Dim CCTypeDesc As String
        oGift.Load lID
        
        CCNum = Right(oGift.Fields(GIFT_fld_Check_Number), 4)
        CCType = UCase(Left(oGift.Fields(GIFT_fld_Check_Number), 1))
        CCTypeDesc = ""
        
        Select Case CCType
        Case "M"
             CCTypeDesc = "Mastercard"
        Case "D"
             CCTypeDesc = "Discover"
        Case "V"
             CCTypeDesc = "Visa"
        Case "A"
             CCTypeDesc = "American Express"
        End Select
        
        If CCTypeDesc <> "" And Len(oGift.Fields(GIFT_fld_Check_Number)) = 5 Then
'        If CCTypeDesc <> "" Then
            With oGift
                Debug.Print .Fields(GIFT_fld_Payment_Type) & "   " & Len(oGift.Fields(GIFT_fld_Check_Number))
                .Fields(GIFT_fld_Payment_Type) = "Credit Card"
                .Fields(GIFT_fld_Credit_Type) = CCTypeDesc
                .Fields(GIFT_fld_Expires_On) = "1/2020"
                .Fields(GIFT_fld_Credit_Card_Number) = "????????????" & CCNum
                
                On Error Resume Next
                .Save
            End With
        End If
          

        oGift.Closedown
        Set oGift = Nothing
    End If

End Sub
'Public Sub GetConstitFromUser(ByVal lRecID As String)
'
'   Dim oSession As IBBSessionContext
'   Dim oRecords As CRecords
'   Dim oRecord As CRecord
'   Dim sSQL As String
'
'   Set oSession = REApplication.SessionContext
'   'select the constituent ID of the user using the session context
'
'   sSQL = "ID in (Select ConstituentID from Users where User_ID = " & oSession.CurrentUserID & ")"
'   Set oRecords = New CRecords
'   oRecords.Init oSession, tvf_record_CustomWhereClause, sSQL
'
'   'if the user is linked to a constituent record, print the constituent ID
'    Debug.Print "Im Here: " & oSession.CurrentUserID
'   If oRecords.Count = 1 Then
'       For Each oRecord In oRecords
'           Debug.Print oRecord.Fields(RECORDS_fld_FULL_NAME) & " " & "-" & " " & oRecord.Fields(RECORDS_fld_CONSTITUENT_ID) & " " & oRecord.Fields(RECORDS_fld_ID)
'           oRecord.Closedown
'           Set oRecord = Nothing
'       Next oRecord
'   End If
'
'   oRecords.Closedown
'   Set oRecords = Nothing
'End Sub

'Public Sub GetConstitID(oRec As IBBDataObject)
'   If TypeOf oRec Is CRecord Then
'       Dim oConstit As CRecord
'       Set oConstit = oRec
'
'       MsgBox oConstit.Fields(RECORDS_fld_CONSTITUENT_ID)
'
'       Set oConstit = Nothing
'   Else
'       MsgBox "This macro must be fired from a constituent record"
'   End If
'End Sub
'
'Public Sub AddMeASPatientCentralAssociate()
'
'    Dim oSession As IBBSessionContext
'    Dim oRecords As CRecords
'    Dim oRecord As CRecord
'    Dim sSQL As String
'    Dim sRecID As Long
'
'    Set oSession = REApplication.SessionContext
'    'select the constituent ID of the user using the session context
'
'    sSQL = "ID in (Select ConstituentID from Users where User_ID = " & oSession.CurrentUserID & ")"
'    Set oRecords = New CRecords
'    oRecords.Init oSession, tvf_record_CustomWhereClause, sSQL
'
'    'if the user is linked to a constituent record, print the constituent ID
'    sRecID = 0
'    If oRecords.Count = 1 Then
'        For Each oRecord In oRecords
'            Debug.Print oRecord.Fields(RECORDS_fld_FULL_NAME) & " " & "-" & " " & oRecord.Fields(RECORDS_fld_CONSTITUENT_ID) & " " & oRecord.Fields(RECORDS_fld_ID)
'            sRecID = oRecord.Fields(RECORDS_fld_ID)
'            oRecord.Closedown
'            Set oRecord = Nothing
'        Next oRecord
'    End If
'
'    oRecords.Closedown
'    Set oRecords = Nothing
''
'    Dim oRec As IBBDataObject
'    Dim oCons As CRecord
'    Set oCons = oRec
'
'    Debug.Print oCons.Fields(RECORDS_fld_CONSTITUENT_ID)
'
'    oCons.Closedown
'    Set oCons = Nothing
''
'
'    Dim oSol As CAssignedSolicitor2
'    Set oSol = New CAssignedSolicitor2
'
'    Set oSession = REApplication.SessionContext
'
'    Debug.Print oSession.CurrentUserID
'
'    With oSol
'        .Init REApplication.SessionContext
'        'Matches the system record ID of the constituent
'        .Fields(ASSIGNEDSOLICITOR2_fld_CONSTIT_ID) = 1553943
'
'        'Matches the system record ID of the solicitor
'        .Fields(ASSIGNEDSOLICITOR2_fld_SOLICITOR_ID) = sRecID
'        .Fields(ASSIGNEDSOLICITOR2_fld_DATE_FROM) = Date
'        .Fields(ASSIGNEDSOLICITOR2_fld_SOLICITOR_TYPE) = "Patient Central Associate"
'        .Save
'        .Closedown
'    End With
'
'    Set oSol = Nothing
'
'End Sub
Public Sub AddMeAsPCA(oRec As IBBDataObject)
    If TypeOf oRec Is CRecord Then
'
'  Find the Constituent ID
'
        Dim ConstitID As String
        Dim oConstit As CRecord
        Set oConstit = oRec
        
        ConstitID = oConstit.Fields(RECORDS_fld_ID)
        
        Dim oPALSSol As CAssignedSolicitor
        
        For Each oPALSSol In oConstit.Relations.AssignedSolicitors
            Debug.Print oPALSSol.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_NAME)
            If oPALSSol.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) = "Patient Central Associate" And _
                oPALSSol.Fields(ASSIGNEDSOLICITOR_fld_DATE_TO) = "" Then
                Debug.Print "Saving.."
                With oPALSSol
                    .Fields(ASSIGNEDSOLICITOR_fld_DATE_TO) = Date
                    On Error Resume Next
                    oConstit.Save
                End With
            End If
        Next
        
        Set oPALSSol = Nothing
        Set oConstit = Nothing
'
'  Find the user's system record ID (Solicitor ID)
'
        Dim oSession As IBBSessionContext
        Dim oRecords As CRecords
        Dim oRecord As CRecord
        Dim sSQL As String
        Dim sRecID As Long
        
        Set oSession = REApplication.SessionContext
        'select the constituent ID of the user using the session context
        
        sSQL = "ID in (Select ConstituentID from Users where User_ID = " & oSession.CurrentUserID & ")"
        Set oRecords = New CRecords
        oRecords.Init oSession, tvf_record_CustomWhereClause, sSQL
        
        'if the user is linked to a constituent record, print the constituent ID
        sRecID = 0
        If oRecords.Count = 1 Then
            For Each oRecord In oRecords
'                Debug.Print oRecord.Fields(RECORDS_fld_FULL_NAME) & " " & "-" & " " & oRecord.Fields(RECORDS_fld_CONSTITUENT_ID) & " " & oRecord.Fields(RECORDS_fld_ID)
                sRecID = oRecord.Fields(RECORDS_fld_ID)
                oRecord.Closedown
                Set oRecord = Nothing
            Next oRecord
        End If
        
        oRecords.Closedown
        Set oRecords = Nothing
'
'   Use the Record ID and Solicitor ID to add the solicitor relationship
'
        
        Dim oSol As CAssignedSolicitor2
        Set oSol = New CAssignedSolicitor2
        
        Set oSession = REApplication.SessionContext
        
'        Debug.Print oSession.CurrentUserID
        
        With oSol
            .Init REApplication.SessionContext
            'Matches the system record ID of the constituent
            .Fields(ASSIGNEDSOLICITOR2_fld_CONSTIT_ID) = ConstitID
        
            'Matches the system record ID of the solicitor
            .Fields(ASSIGNEDSOLICITOR2_fld_SOLICITOR_ID) = sRecID
            .Fields(ASSIGNEDSOLICITOR2_fld_DATE_FROM) = Date
            .Fields(ASSIGNEDSOLICITOR2_fld_SOLICITOR_TYPE) = "Patient Central Associate"
            On Error Resume Next
            .Save
            .Closedown
        End With
        
        MsgBox "You are now assigned as the PCA of this constituent.  Please make sure to refresh your screen"
        
        Set oSol = Nothing
    Else
        MsgBox "This macro must be fired from a constituent record"
    End If
End Sub

Public Sub GetConstitUsingAlias()
   Dim oRecords As CRecords
   Dim oRecord As CRecord
   
   Dim sSQL As String
   Dim sConstit As String

   sSQL = "dbo.RECORDS.ID in (SELECT dbo.ALIASNAME.RECORDS_ID " & _
       "FROM dbo.ALIASNAME WHERE ((dbo.ALIASNAME.KEY_NAME) = '2213505') AND (dbo.ALIASNAME.ALIAS_TYPE = '10835'))"

   Set oRecords = New CRecords

   oRecords.Init REApplication.SessionContext, tvf_record_CustomWhereClause, sSQL

   MsgBox oRecords.Count

   sConstit = ""

   For Each oRecord In oRecords
'       If oRecord.Fields(RECORDS_fld_KEY_INDICATOR) = bbki_IND Then
           sConstit = sConstit & oRecord.Fields(RECORDS_fld_ID) & oRecord.Fields(RECORDS_fld_LAST_NAME) & vbCrLf
'       Else
'           sConstit = sConstit & oRecord.Fields(RECORDS_fld_FIRST_NAME) & " " & oRecord.Fields(RECORDS_fld_LAST_NAME) & vbCrLf
'       End If
       oRecord.Closedown
   Next oRecord

   Set oRecord = Nothing
   MsgBox sConstit
   oRecords.Closedown
   Set oRecords = Nothing
End Sub
Public Sub AddGiftFundraiser(oRow As IBBQueryRow)
    Const GiftID = 1

    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
        Dim oService As REServices
        Set oService = New REServices
        oService.Init REApplication.SessionContext

        Dim oAttributeServer As CAttributeTypeServer
        Set oAttributeServer = oService.CreateServiceObject(bbsoAttributeTypeServer)
        oAttributeServer.Init REApplication.SessionContext
        
        Dim FLOID As Long
                
        FLOID = oAttributeServer.GetAttributeTypeID("Fundraiser's LO ConstID", bbAttributeRecordType_GIFT)
            
        Dim oGift As CGift
        Set oGift = New CGift
        oGift.Init RE7.SessionContext
                
        Dim lID As Long
        'get the id
        lID = oRow.Field(GiftID)
        oGift.Load lID
        
        Dim oAttribute As IBBAttribute
        Dim FID As Long
        FID = 0
'        Debug.Print FLOID
        
        For Each oAttribute In oGift.Attributes
            If (oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = FLOID) Then
                FID = oAttribute.Fields(Attribute_fld_VALUE)
'                Debug.Print FID
'
                Dim oRecords As CRecords
                Dim oRecord As CRecord
                
                Dim sSQL As String
                Dim sConstit As String
                
                sSQL = "dbo.RECORDS.ID in (SELECT dbo.ALIASNAME.RECORDS_ID " & _
                   "FROM dbo.ALIASNAME WHERE ((dbo.ALIASNAME.KEY_NAME) = '" & FID & "') AND (dbo.ALIASNAME.ALIAS_TYPE = '10835'))"
                
                Set oRecords = New CRecords
                
                oRecords.Init REApplication.SessionContext, tvf_record_CustomWhereClause, sSQL
                
                Debug.Print oRecords.Count
                
                sConstit = ""
                
                For Each oRecord In oRecords
                   FID = oRecord.Fields(RECORDS_fld_ID)
                   oRecord.Closedown
                Next oRecord
                
'                Debug.Print "Const ID: " & FID
                Set oRecord = Nothing
                oRecords.Closedown
                Set oRecords = Nothing
'
            End If
        Next oAttribute
        
        Dim oSol As CAssignedSolicitor2
        
        If FID > 0 Then
'            Debug.Print oGift.Fields(GIFT_fld_ID)
            With oGift.Solicitors.Add
                .Fields(RECORDSOLICITOR_fld_Amount) = oGift.Fields(GIFT_fld_Amount)
                .Fields(RECORDSOLICITOR_fld_SolicitorId) = FID
'                On Error Resume Next
                oGift.Save
            End With
        End If
        
        
        
        Set oSol = Nothing
        
        oService.Closedown
        Set oAttributeServer = Nothing
        Set oAttribute = Nothing
        Set oService = Nothing
        Set oGift = Nothing
     End If
End Sub

Public Sub ReMoveExtraUser(oRow As IBBQueryRow)
    Const ActionID = 1

    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
        Dim oService As REServices
        Set oService = New REServices
        oService.Init REApplication.SessionContext

        Dim oAction As CAction
        Set oAction = New CAction
        oAction.Init REApplication.SessionContext

        Dim lID As Long
        lID = oRow.Field(ActionID)

        oAction.Load lID
        
        Dim oRemindee As CActionRemindee
        
        For Each oRemindee In oAction.Remindees
            If oRemindee.Fields(ActionRemindee_fld_NAME) = "eeun" Then
                With oRemindee
                    oAction.Remindees.Remove oRemindee.Fields(ActionRemindee_fld_ID)
                    oAction.Remindees.Add.Fields(ActionRemindee_fld_USER_ID) = 368
                End With
            End If
        Next
        
        oAction.Save

        Set oRemindee = Nothing
        
        oAction.Closedown
        Set oAction = Nothing
        
        oService.Closedown
        Set oService = Nothing
        
    End If
End Sub
Public Sub End_PAC_Solicitor_Rel(oRow As IBBQueryRow)
    
    Dim Const_ID As Long
    
    Dim NoNextAction As Boolean
    
    Const_ID = 1
    
    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init RE7.SessionContext
        
        Dim lID As Long
        Dim sID As Long
        Dim sType As String
        lID = oRow.Field(Const_ID)
        sID = oRow.Field(3)
        sType = oRow.Field(4)
        oCons.Load lID
        
        Debug.Print oCons.Fields(RECORDS_fld_PRIMARY_ADDRESSEE)

        Dim oASolicitor As CAssignedSolicitor
        
        'loop through the constituent's assigned solicitors
        For Each oASolicitor In oCons.Relations.AssignedSolicitors
            If oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_ID) = sID And _
                oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) = sType Then
                If oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_DATE_TO) = "" Then
                    Debug.Print "     " & oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_ID) & " " & oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_NAME) & " " & sID
                    With oCons
                        oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_DATE_TO) = "09/24/2017"
                        .Save
                    End With
                End If
            End If
        Next oASolicitor
        
        Set oASolicitor = Nothing
        
        oCons.Closedown
        Set oCons = Nothing
    
    End If
End Sub

Public Sub GetGiftGLJournalReference(oRow As IBBQueryRow)
    Const GiftID = 1

    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
        Dim oService As REServices
        Set oService = New REServices
        oService.Init REApplication.SessionContext
            
        Dim oGift As CGift
        Set oGift = New CGift
        oGift.Init RE7.SessionContext
                
        Dim lID As Long
        'get the id
        lID = oRow.Field(GiftID)
        oGift.Load lID
        
        oRow.Field("GL Reference") = oGift.Fields(GIFT_fld_GLJournalReference)
        
        oService.Closedown
        Set oService = Nothing
        Set oGift = Nothing
     End If
End Sub

Public Sub ChangeCCodetoGeneral(oRow As IBBQueryRow)
    Const ConsID = 1

    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
    
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init REApplication.SessionContext
        
        Dim oID As Long
        
        oID = oRow.Field(ConsID)
        oCons.Load oID
'
        Dim oCCode As CConstituentCode
    
        For Each oCCode In oCons.ConstituentCodes
            Debug.Print oCCode.Fields(CONSTITUENT_CODE_fld_CODE) & " " & oCCode.Fields(CONSTITUENT_CODE_fld_DATE_TO)
            If oCCode.Fields(CONSTITUENT_CODE_fld_DATE_TO) = "" Then
                With oCCode
                    .Fields(CONSTITUENT_CODE_fld_CODE) = "General Constituent"
                End With
                oCons.Save
            End If
            Set oCCode = Nothing
        Next oCCode
       
'
        oCons.Closedown
        Set oCons = Nothing
    End If

End Sub
Public Sub FBFundraisersFollowUp(oRow As IBBQueryRow)
    Const ConsID = 1

    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
        Dim oService As REServices
        Set oService = New REServices
        oService.Init REApplication.SessionContext

        Dim oAttributeServer As CAttributeTypeServer
        Set oAttributeServer = oService.CreateServiceObject(bbsoAttributeTypeServer)
        oAttributeServer.Init REApplication.SessionContext
        
        Dim FBID As Long
                
        FBID = oAttributeServer.GetAttributeTypeID("Source/General", bbAttributeRecordType_CONSTITUENT)


'
'  Find the Constituent ID
'

        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init REApplication.SessionContext
        
        Dim oID As Long
        Dim AddNew As Boolean
        Dim ConstitID As String
        
        oID = oRow.Field(ConsID)
        oCons.Load oID
'
        Dim oAttribute As IBBAttribute

        For Each oAttribute In oCons.Attributes
            If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = FBID And _
                oAttribute.Fields(Attribute_fld_VALUE) = "FB Donor" Then
                With oCons
                    .Attributes.Remove oAttribute
                    .Save
                End With
            End If
        Next oAttribute
                
        ConstitID = oCons.Fields(RECORDS_fld_ID)
        
        AddNew = True

        Dim oPRMSol As CAssignedSolicitor
        Dim SolType As String
        Dim SolID As Long
        
        SolType = "Primary Relationship Manager"

        'Matches the system record ID of the solicitor
        'Use Elaine record's system ID  1541716
        SolID = 1541716

        For Each oPRMSol In oCons.Relations.AssignedSolicitors
            Debug.Print oPRMSol.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_NAME)
            If oPRMSol.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_ID) = SolID And _
                oPRMSol.Fields(ASSIGNEDSOLICITOR_fld_DATE_TO) = "" Then
                AddNew = False
                Exit For
            End If
            
            If oPRMSol.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) = "Primary Relationship Manager" And _
                oPRMSol.Fields(ASSIGNEDSOLICITOR_fld_DATE_TO) = "" Then
                SolType = "Secondary Relationship Manager"
            End If
        Next

        Set oPRMSol = Nothing
'
'   Use the Record ID and Solicitor ID to add the solicitor relationship
'
        If AddNew = True Then
            Dim oSol As CAssignedSolicitor2
            Set oSol = New CAssignedSolicitor2
        
            Set oSession = REApplication.SessionContext
        
            With oSol
                .Init REApplication.SessionContext
                'Matches the system record ID of the constituent
                .Fields(ASSIGNEDSOLICITOR2_fld_CONSTIT_ID) = ConstitID

                
                .Fields(ASSIGNEDSOLICITOR2_fld_SOLICITOR_ID) = SolID
                .Fields(ASSIGNEDSOLICITOR2_fld_DATE_FROM) = Date
                .Fields(ASSIGNEDSOLICITOR2_fld_SOLICITOR_TYPE) = SolType
'                On Error Resume Next
                .Save
                .Closedown
                Debug.Print "Adding it"
            End With
                    
            Set oSol = Nothing
        End If
        
        oService.Closedown
        Set oAttributeServer = Nothing
        Set oAttribute = Nothing
        Set oService = Nothing
        
        oCons.Closedown
        Set oCons = Nothing

    End If
End Sub
Public Sub AddUSCountrytoPrefAddress(oRow As IBBQueryRow)
    Const ConsID = 1

    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
    
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init REApplication.SessionContext
        
        Dim oID As Long
        
        oID = oRow.Field(ConsID)
        oCons.Load oID
'
        Dim oAddress As CConstitAddress
         
        For Each oAddress In oCons.Addresses
            If oAddress.Fields(CONSTIT_ADDRESS_fld_PREFERRED) = "-1" Then
                If oAddress.Fields(CONSTIT_ADDRESS_fld_ADDRESS_BLOCK) <> "PO Box" Then
                   With oAddress
                        .Fields(CONSTIT_ADDRESS_fld_COUNTRY) = "United States"
                    
                        On Error Resume Next
                        oCons.Save
                   End With
                End If
            End If
        Next oAddress
'
        oCons.Closedown
        Set oCons = Nothing
    End If

End Sub

Public Sub GratefulPatientGiftAdjustments(oRow As IBBQueryRow)
    Const Gift_ID = 1
    
    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
        Dim oGift As CGift
        Set oGift = New CGift
        oGift.Init RE7.SessionContext
                
        Dim GPTributeType As String
        GPTributeType = Trim(oRow.Field(3))
                
        If GPTributeType <> "" Or GPTributeType <> "Other" Then
            Dim lID As Long
            'get the id
            lID = oRow.Field(Gift_ID)
            
            oGift.Load lID
    
            Dim oTrbRec As CGiftTribute
          
            For Each oTrbRec In oGift.Tributes
                With oTrbRec
                    Debug.Print oTrbRec.Fields(GIFTTRIBUTE_fld_Tribute_Type) & " " & oTrbRec.Fields(GIFTTRIBUTE_fld_Tribute_Description)
                    oTrbRec.Fields(GIFTTRIBUTE_fld_Tribute_Type) = GPTributeType
                    On Error Resume Next
                    oGift.Save
                End With
            Next oTrbRec
            
            Set oTrbRec = Nothing
        End If

        oGift.Closedown
        
        Set oGift = Nothing
    End If

End Sub

Public Sub AdjustRecurringGifts(oRow As IBBQueryRow)
    Const Gift_ID = 1
    
    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
        Dim oGift As CGift
        Set oGift = New CGift
        oGift.Init RE7.SessionContext
                
        Dim lID As Long
        'get the id
        lID = oRow.Field(Gift_ID)
        
        oGift.Load lID
            
        If oGift.Fields(GIFT_fld_Type) = "Recurring Gift" Or oGift.Fields(GIFT_fld_Type) = "Recurring Gift Pay-Cash" Then
            If oGift.Fields(GIFT_fld_Post_Status) <> "Posted" Then
                With oGift
'                    .Fields(GIFT_fld_Appeal) = "Monthly Donation"
                    .Fields(ADJUSTMENT_fld_Appeal) = Chr(34) & "Monthly Donation" & Chr(34)
                    .Fields(GIFT_fld_Campaign) = "Annual"
                    .Fields(GIFT_fld_Fund) = "General Operating Donations - Unrestricted"
                    On Error Resume Next
                    oGift.Save
                End With
            End If
        End If
'
            If oGift.Fields(GIFT_fld_Post_Status) = "Posted" Then
                Dim oAdjustment As IBBAdjustment
                Dim oAdjustmentServer As CAdjustmentServer
    
                Set oAdjustmentServer = New CAdjustmentServer
    
                With oAdjustmentServer
                    .Init REApplication.SessionContext, oGift
    '                On Error Resume Next
                    Set oAdjustment = .AddAdjustment()
                    With oAdjustment
                    Debug.Print "creating adjustment on " & oGift.Fields(GIFT_fld_ID) & " " & oGift.Fields(GIFT_fld_Amount) & " " & oGift.Fields(GIFT_fld_Date)
                        .Fields(ADJUSTMENT_fld_Date) = Date
                        .Fields(ADJUSTMENT_fld_Amount) = oGift.Fields(GIFT_fld_Amount)
                        .Fields(ADJUSTMENT_fld_Reason) = "Recurring gift adjustment"
                        .Fields(ADJUSTMENT_fld_Fund) = oGift.Fields(GIFT_fld_Fund)
                        .Fields(ADJUSTMENT_fld_Campaign) = "Annual"
                        .Fields(ADJUSTMENT_fld_Appeal) = Chr(34) & "Monthly Donation" & Chr(34)
'                        .Fields(ADJUSTMENT_fld_Appeal) = "Monthly Donation"
                        .Fields(ADJUSTMENT_fld_Package) = ""
                    End With
                    'Validate and save the adjustment
                    On Error Resume Next
                    .Validate
                    oAdjustmentServer.Save
                End With
    
                Set oAdjustmentServer = Nothing
                On Error Resume Next
                oGift.Save
            End If

'
        
        oGift.Closedown
        Set oGift = Nothing
    End If

End Sub
Public Sub FirstTimeDonor(oRow As IBBExportRow)
    Const ConsID = 16

    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
    
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init REApplication.SessionContext
        
        Dim oID As Long
        Dim GiftNum As Integer
        Dim GiftDate As Date
        GiftDate = "01/01/9999"
        GiftNum = 0
        
        oID = oRow.Field(ConsID)
        oCons.Load oID
        
        Dim oGift As CGift
        
        Debug.Print oCons.Fields(RECORDS_fld_CONSTITUENT_ID)
        
        For Each oGift In oCons.Gifts
            If oGift.Fields(GIFT_fld_Amount) > 0 Then
                If oGift.Fields(GIFT_fld_Date) <> GiftDate Then
                    GiftNum = GiftNum + 1
                    GiftDate = oGift.Fields(GIFT_fld_Date)
                End If
            End If
        Next oGift
        
        If GiftNum > 1 Then
            oRow.Field(8) = "Multiple Times"
        Else
            oRow.Field(8) = "First Time"
        End If
        
        Set oGift = Nothing
        
        oCons.Closedown
        Set oCons = Nothing
    End If

End Sub
Public Sub Mark_Caregivers(oRow As IBBQueryRow)
    Const ConsID = 1

    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
    
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init REApplication.SessionContext
        
        Dim oID As Long
        
        oID = oRow.Field(ConsID)
        oCons.Load oID
        
        Dim CGD As Date
        Dim PCD As Date
        Dim PCDS As Date
        
        CGD = "01/01/9999"
        PCD = "01/01/9999"
        PCDS = "01/01/9999"
        
        Dim oService As REServices
        Set oService = New REServices
        oService.Init REApplication.SessionContext
           
        Dim oAttributeServer As CAttributeTypeServer
        Set oAttributeServer = oService.CreateServiceObject(bbsoAttributeTypeServer)
        oAttributeServer.Init REApplication.SessionContext
        
        Dim Date1 As Long
        Dim Date2 As Long
        Dim Date3 As Long
        
        Date1 = oAttributeServer.GetAttributeTypeID("Caregiver", bbAttributeRecordType_CONSTITUENT)
        Date2 = oAttributeServer.GetAttributeTypeID("PC_Connection", bbAttributeRecordType_CONSTITUENT)
        Date3 = oAttributeServer.GetAttributeTypeID("PC_Connection (Staff)", bbAttributeRecordType_CONSTITUENT)
        
        Dim oAttribute As IBBAttribute

        Debug.Print oCons.Fields(RECORDS_fld_CONSTITUENT_ID)

        For Each oAttribute In oCons.Attributes
            If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = Date1 Then
                If Trim(oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE)) <> "" Then
                    CGD = oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE)
                End If
            End If
            If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = Date2 And _
                oAttribute.Fields(Attribute_fld_VALUE) = "Caregiver" Then
                If Trim(oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE)) <> "" Then
                    PCD = oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE)
                End If
            End If
            If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = Date3 And _
                oAttribute.Fields(Attribute_fld_VALUE) = "Caregiver" Then
                If Trim(oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE)) <> "" Then
                    PCDS = oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE)
                End If
            End If
        Next oAttribute
        
        Set oAttribute = Nothing
        
        Debug.Print CGD & " " & PCD & " " & PCDS
        If CGD > Date Then CGD = Date
        If CGD > PCD Then CGD = PCD
        If CGD > PCDS Then CGD = PCDS
        Debug.Print CGD & " " & PCD & " " & PCDS
'
        Dim oCCode As CConstituentCode
        Dim FoundIt As Boolean
        FoundIt = False
        
        For Each oCCode In oCons.ConstituentCodes
            If (oCCode.Fields(CONSTITUENT_CODE_fld_LONG_DESC_READONLY) = "Caregiver") Then
                FoundIt = True
                With oCCode
                    .Fields(CONSTITUENT_CODE_fld_DATE_TO) = ""
                End With
            End If
        Next oCCode
'
        If FoundIt = False Then
            With oCons.ConstituentCodes.Add
                .Fields(CONSTITUENT_CODE_fld_CODE) = "Caregiver"
                If CGD <= Date Then
                    .Fields(CONSTITUENT_CODE_fld_DATE_FROM) = CGD
                End If
            End With
        End If
        
        On Error Resume Next
        oCons.Save
'
        oCons.Closedown
        Set oCons = Nothing
        Set oCCode = Nothing
        
        oService.Closedown
        Set oAttributeServer = Nothing
        Set oAttribute = Nothing
        Set oService = Nothing
        
    End If

End Sub
Public Sub UnMark_Caregivers(oRow As IBBQueryRow)
    Const ConsID = 1

    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
    
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init REApplication.SessionContext
        
        Dim oID As Long
        
        oID = oRow.Field(ConsID)
        oCons.Load oID
        
        Dim CGD As Date
        Dim PCD As Date
        Dim PCDS As Date
        
        CGD = "1/1/9999"
        PCD = "1/1/9999"
        PCDS = "1/1/9999"
        
        Dim oService As REServices
        Set oService = New REServices
        oService.Init REApplication.SessionContext
           
        Dim oAttributeServer As CAttributeTypeServer
        Set oAttributeServer = oService.CreateServiceObject(bbsoAttributeTypeServer)
        oAttributeServer.Init REApplication.SessionContext
        
        Dim Date1 As Long
        Dim Date2 As Long
        Dim Date3 As Long
        
        Date1 = oAttributeServer.GetAttributeTypeID("Caregiver", bbAttributeRecordType_CONSTITUENT)
        Date2 = oAttributeServer.GetAttributeTypeID("PC_Connection", bbAttributeRecordType_CONSTITUENT)
        Date3 = oAttributeServer.GetAttributeTypeID("PC_Connection (Staff)", bbAttributeRecordType_CONSTITUENT)
        
        Dim oAttribute As IBBAttribute

        For Each oAttribute In oCons.Attributes
            If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = Date1 And _
                oAttribute.Fields(Attribute_fld_VALUE) = False Then
                If Trim(oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE)) <> "" Then
                    CGD = oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE)
                End If
            End If
            If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = Date2 And _
                oAttribute.Fields(Attribute_fld_VALUE) = "Lost a loved one to pancreatic cancer" Then
                If Trim(oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE)) <> "" Then
                    PCD = oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE)
                End If
            End If
            If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = Date3 And _
                oAttribute.Fields(Attribute_fld_VALUE) = "Lost a loved one to pancreatic cancer" Then
                If Trim(oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE)) <> "" Then
                    PCDS = oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE)
                End If
            End If
        Next oAttribute
        
        Set oAttribute = Nothing
        
        Debug.Print CGD & " " & PCD & " " & PCDS
        If CGD > Date Then CGD = Date
        If CGD > PCD Then CGD = PCD
        If CGD > PCDS Then CGD = PCDS
        Debug.Print CGD & " " & PCD & " " & PCDS
'
        Dim oCCode As CConstituentCode
        Dim FoundIt As Boolean
        FoundIt = False
        
        For Each oCCode In oCons.ConstituentCodes
            If (oCCode.Fields(CONSTITUENT_CODE_fld_LONG_DESC_READONLY) = "Caregiver") Then
                FoundIt = True
                With oCCode
                    .Fields(CONSTITUENT_CODE_fld_DATE_TO) = CGD
                End With
            End If
        Next oCCode
'
'        If FoundIt = False Then
'            With oCons.ConstituentCodes.Add
'                .Fields(CONSTITUENT_CODE_fld_CODE) = "Caregiver"
'                .Fields(CONSTITUENT_CODE_fld_DATE_FROM) = CGD
'            End With
'        End If
        
'        On Error Resume Next
        oCons.Save
'
        oCons.Closedown
        Set oCons = Nothing
        Set oCCode = Nothing
        
        oService.Closedown
        Set oAttributeServer = Nothing
        Set oAttribute = Nothing
        Set oService = Nothing
        
    End If

End Sub
Public Sub Find_RecurringGiftID(oRow As IBBQueryRow)
    Const Gift_ID = 1
    
    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
        Dim oGift As CGift
        
        Set oGift = New CGift
        oGift.Init RE7.SessionContext
                
        Dim oGift2 As CGift
        Set oGift2 = New CGift
        oGift2.Init RE7.SessionContext
        
        oGift2.Load 1732308
                
        Dim lID As Long
        'get the id
        lID = oRow.Field(Gift_ID)
        
        oGift.Load lID
        
        With oGift
            .PledgePayer.ApplyToRecurringGift oGift2
            .Save
        End With
        

        oGift.Closedown
        oGift2.Closedown
        
        Set oGift = Nothing
        Set oGift2 = Nothing
    End If

End Sub
Private Function fncIsMail(ByVal strEmail As String) As Boolean
    Const strRFC2822 = "[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!" & _
                        "#$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:" & _
                        "[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:" & _
                        "[a-z0-9-]*[a-z0-9])?"
                        
'    Const strRFC2822 = "^[_a-z0-9-]+(.[a-z0-9-]+)@[a-z0-9-]+(.[a-z0-9-]+)*(.[a-z]{2,4})$"
    
    Dim objRegEx As Object
    On Error GoTo Fin
    Set objRegEx = CreateObject("Vbscript.Regexp")
    With objRegEx
        .Pattern = strRFC2822
        .IgnoreCase = True
        fncIsMail = .Test(strEmail)
    End With
Fin:
    Set objRegEx = Nothing
    If Err.Number <> 0 Then MsgBox "Error: " & _
        Err.Number & " " & Err.Description
End Function

Public Sub Email_Validator(oRow As IBBQueryRow)
    Const Const_ID = 1
    
    If oRow.BOF Then
        MsgBox "Beging processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
    
        Dim strEmail As String
        Dim ValidEmail As Boolean
    
        strEmail = oRow.Field("Email")
                
        ValidEmail = fncIsMail(strEmail)
    
        oRow.Field("Valid Email") = ValidEmail
        
        If ValidEmail = False Then
            
            Dim oCons As CRecord
            Set oCons = New CRecord
            oCons.Init RE7.SessionContext
            
            Dim lID As Long
            
            lID = oRow.Field(Const_ID)
            oCons.Load lID

            Dim oPhone As IBBPhone
            Dim oAddress As CConstitAddress
            
            For Each oAddress In oCons.Addresses
'                Debug.Print oAddress.Fields(CONSTIT_ADDRESS_fld_CITY)
                For Each oPhone In oAddress.Phones
                    If oPhone.Fields(Phone_fld_PhoneType) = "E-mail" And _
                       oPhone.Fields(Phone_fld_Num) = strEmail Then
'                       Debug.Print strEmail & " " & oPhone.Fields(Phone_fld_Num) & " "; oPhone.Fields(Phone_fld_PhoneType)
                        With oPhone
                            .Fields(Phone_fld_Inactive) = True
                            On Error Resume Next
                            oCons.Save
                        End With
                    End If
                Next oPhone
            Next oAddress
        
            oCons.Closedown
            Set oCons = Nothing
            
        End If
             
    End If
End Sub

Private Function ProperCase(sText)
'*** Converts text to proper case e.g.  ***'
'*** surname = Surname                  ***'
'*** o'connor = O'Connor                ***'
 
    Dim a, iLen, bSpace, tmpX, tmpFull
 
    iLen = Len(sText)
    For a = 1 To iLen
    If a <> 1 Then 'just to make sure 1st character is upper and the rest lower'
        If bSpace = True Then
            tmpX = UCase(Mid(sText, a, 1))
            bSpace = False
        Else
        tmpX = LCase(Mid(sText, a, 1))
            If tmpX = " " Or tmpX = "'" Then bSpace = True
        End If
    Else
        tmpX = UCase(Mid(sText, a, 1))
    End If
    tmpFull = tmpFull & tmpX
    Next
    ProperCase = tmpFull
End Function
Public Sub FixNameCase(oRow As IBBQueryRow)
    Const ConsID = 1
    
    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init REApplication.SessionContext
        
        Dim oID As Long
        
        oID = oRow.Field(ConsID)
        oCons.Load oID
        
        Dim LName As String
        Dim FName As String
        Dim UpdateIt As Boolean
        LName = oCons.Fields(RECORDS_fld_LAST_NAME)
        FName = oCons.Fields(RECORDS_fld_FIRST_NAME)
        
        oRow.Field("Valid Email") = " "
        UpdateIt = False
        
        
        If LName = UCase(LName) Or FName = UCase(FName) Then
            oRow.Field("Valid Email") = "All CAPS"
            oRow.Field("Proper LName") = ProperCase(LName)
            oRow.Field("Proper FName") = ProperCase(FName)
            UpdateIt = True
        End If
        If LName = LCase(LName) Or FName = LCase(FName) Then
            oRow.Field("Valid Email") = "All lower"
            oRow.Field("Proper FName") = ProperCase(FName)
            oRow.Field("Proper LName") = ProperCase(LName)
            UpdateIt = True
        End If
        If UpdateIt Then
            With oCons
                .Fields(RECORDS_fld_FIRST_NAME) = ProperCase(FName)
                .Fields(RECORDS_fld_LAST_NAME) = ProperCase(LName)
                On Error Resume Next
                .Save
            End With
        End If
        
        oCons.Closedown
        Set oCons = Nothing
        
    End If
End Sub

Public Sub GetPrimaryAddSal(oRow As IBBExportRow)
    Const PConsID = 14
    Const NameLoc = 41
    Const ConsID = 42
    Const AddLoc = 43
    Const SalLoc = 44
    Const DevSol = 45
    Const PriMgr = 46
    Const SecMgr = 47
    
    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else

        If Trim(oRow.Field(NameLoc)) <> "" Then
            Dim oCons As CRecord
            Set oCons = New CRecord
            oCons.Init REApplication.SessionContext
            
            Dim oID As String
            oID = oRow.Field(ConsID)
        
            oCons.LoadByField uf_Record_CONSTITUENT_ID, oID
            
            Debug.Print oCons.Fields(RECORDS_fld_CONSTITUENT_ID), oCons.Fields(RECORDS_fld_PRIMARY_ADDRESSEE)
            
            Dim oAddressee As String
            Dim oSalutation As String
    
            oAddressee = oCons.Fields(RECORDS_fld_PRIMARY_ADDRESSEE)
            oSalutation = oCons.Fields(RECORDS_fld_PRIMARY_SALUTATION)
    
            oRow.Field(AddLoc) = oAddressee
            oRow.Field(SalLoc) = oSalutation
            
            oCons.Closedown
            Set oCons = Nothing
        End If
'
        Dim oCons2 As CRecord
        Set oCons2 = New CRecord
        oCons2.Init REApplication.SessionContext
        
        Dim oID2 As String
        oID2 = oRow.Field(PConsID)
    
        oCons2.LoadByField uf_Record_CONSTITUENT_ID, oID2

        Dim DevSolicitor As String
        Dim SecondaryRelManager As String
        Dim PrimaryRelManager As String
        Dim OtherSolicitor As String
        
        DevSolicitor = ""
        SecondaryRelManager = ""
        PrimaryRelManager = ""
        OtherSolicitor = ""
                  
        Dim oSolicitor As CSolicitorActions
        Dim oASolicitor As CAssignedSolicitor
        Dim oSolID As CActionSolicitor
        
        'loop through the constituent's assigned solicitors
        For Each oASolicitor In oCons2.Relations.AssignedSolicitors
            Debug.Print oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_NAME)
            If oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_DATE_TO) = "" Then
                If (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) = "Development Solicitor") Then
                    DevSolicitor = oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_NAME)
                End If
                If (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) = "Primary Relationship Manager") Then
                    PrimaryRelManager = oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_NAME)
                End If
                If (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) = "Secondary Relationship Manager") Then
                    SecondaryRelManager = oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_NAME)
                End If
                If (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) <> "Development Solicitor") And _
                    (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) <> "Primary Relationship Manager") And _
                    (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) <> "Secondary Relationship Manager") Then
                    OtherSolicitor = oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_NAME)
                End If
            End If
        Next oASolicitor

        oRow.Field(DevSol) = DevSolicitor
        oRow.Field(PriMgr) = PrimaryRelManager
        oRow.Field(SecMgr) = SecondaryRelManager
            
        Set oASolicitor = Nothing
        Set oSolID = Nothing

        oCons2.Closedown
        Set oCons2 = Nothing
    End If
End Sub

Public Sub EmailPALSOrder(oRow As IBBQueryRow)
'    Const ConsID = 4

    If oRow.BOF Then
        MsgBox "Begin processing"
        
        PatPck = 0
        ShrSty = 0
        SpkPAL = 0
        SurNet = 0
        EduMat = 0
        
    ElseIf oRow.EOF Then
    
        MsgBox "End processing"
        
        Dim oOutlook As Outlook.Application
        Set oOutlook = New Outlook.Application
        
        Dim oMailItem As Outlook.MailItem
        Set oMailItem = oOutlook.CreateItem(olMailItem)
        
        With oMailItem
            .To = "mgarcia@pancan.org;fzelada-arenas@pancan.org;kbauer@pancan.org"
            .CC = "helpdesk@pancan.org;lcoronado@pancan.org;npangan@pancan.org"
            .Subject = "Patient Packet Daily Import"
            .Body = "Hello Mayra, " & vbNewLine & vbNewLine & _
                    "For " & Date & ", the following were imported into Raiser’s Edge:" & vbNewLine & vbNewLine & _
                    "   -  Patient Packets/SCN:  " & PatPck & vbNewLine & _
                    "   -  Share Your Story Interest:  " & ShrSty & vbNewLine & _
                    "   -  Speak to Patient Central:  " & SpkPAL & vbNewLine & _
                    "   -  Survivor Caregiver Network:  " & SurNet & vbNewLine & _
                    "   -  Educational Materials:  " & EduMat & vbNewLine & vbNewLine & _
                    "Please let me know if you have any questions." & vbNewLine & vbNewLine & _
                    "#assign to me" & vbNewLine & _
                    "#category Data Import/Export" & vbNewLine & _
                    "#worked 30m" & vbNewLine & _
                    "#close"
            .Display
'            .Send
        End With
        
        Set oMailItem = Nothing
        Set oOutlook = Nothing
        
    Else
       
        If oRow.Field("Action Type") = "PALS Patient Packet-Online" Then PatPck = PatPck + 1
        If oRow.Field("Action Type") = "Share Your Story Interest" Then ShrSty = ShrSty + 1
        If oRow.Field("Action Type") = "PALS-Speak ToPatientCentral" Then SpkPAL = SpkPAL + 1
        If oRow.Field("Action Type") = "PALS-Survivor Caregiver Network" Then SurNet = SurNet + 1
        If oRow.Field("Action Type") = "PALS HP Materials" Then EduMat = EduMat + 1
        
        Debug.Print PatPck
        Debug.Print ShrSty
        Debug.Print SpkPAL
        Debug.Print SurNet
        Debug.Print EduMat
            
    End If

End Sub

Public Sub Export_SDO_Actions(oRow As IBBExportRow)
    
    Dim Const_ID As Long
    Dim Last_Completed_Action_Date As Date
    Dim Last_Completed_Action_Type As String
    Dim Last_Contact_Report_Date As Date
    Dim Last_Contact_Report_Description As String
    Dim Next_Action_Date As Date
    Dim Next_Action_Type As String
    Dim With_Proposal As Boolean
    Dim With_Strategy_Note As Boolean
    Dim FoundIt As Boolean
    Dim FoundIt2 As Boolean
    Dim FoundIt3 As Boolean
    
    Const_ID = 1
    
    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init RE7.SessionContext
        
        Dim lID As Long
        'get the id
        lID = oRow.Field(Const_ID)
        oCons.Load lID
        
        Debug.Print oCons.Fields(RECORDS_fld_PRIMARY_ADDRESSEE)
        
        Last_Completed_Action_Date = "1/1/1900"
        Last_Contact_Report_Date = "1/1/1900"
        Next_Action_Date = Date
        
        FoundIt = False
        FoundIt2 = False
        FoundIt3 = False
'
        Dim Relations As String
        Relations = ""
        
        Dim oRels As CRelationships2
        Set oRels = New CRelationships2
        oRels.Init REApplication.SessionContext, lID
        
        Dim oRel As CRelationship2
        
        Dim SolID As Long
        
        SolID = oRow.Field(20)
        
        Debug.Print SolID
        
        For Each oRel In oRels
            Debug.Print "relationship"
            If oRel.Fields(RELATIONSHIP_FLD_Type) = "Development Solicitor" And _
                (IsNull(oRel.Fields(RELATIONSHIP_FLD_Date_To)) = True Or _
                oRel.Fields(RELATIONSHIP_FLD_Date_To) > Date) Then
                Relations = oRel.Fields(RELATIONSHIP_FLD_Name)
            End If
        Next oRel
        
        oRels.Closedown
        Set oRels = Nothing
        Set oRel = Nothing
'
        
        Dim oActions As CAction
'        Dim oNote As IBBNotepad
        Dim oSolicitor As CSolicitorActions
        
        Dim oSolic As CActionSolicitor
'        Set oAction = New Action
'        oAction.Init REApplication.SessionContext
        
        Debug.Print oCons.Actions.Count
        
        If oCons.Actions.Count > 0 Then
            For Each oActions In oCons.Actions
                Debug.Print "Actions"
                
'                Set oSolicitor = oActions.Solicitors
                For Each oSolic In oActions.Solicitors
'                    Debug.Print oSolic.Fields(ActionSolicitor_fld_ID) & " " & oSolic.Fields(ActionSolicitor_fld_IMPORT_ID) & " " & oSolic.Fields(ActionSolicitor_fld_RECORDS_ID) & " " & oSolic.Fields(ActionSolicitor_fld_NAME)
                    Debug.Print "HERE me: " & oSolID & " ..." & oSolic.Fields(ActionSolicitor_fld_RECORDS_ID)
                    If oSolic.Fields(ActionSolicitor_fld_RECORDS_ID) = SolID Then
                        Debug.Print "FOUND IT"
                        If oActions.Fields(ACTION_fld_COMPLETED) = True Then
                            If oActions.Fields(ACTION_fld_COMPLETED) = True And oActions.Fields(ACTION_fld_DATE) > Last_Completed_Action_Date Then
                                Last_Completed_Action_Date = oActions.Fields(ACTION_fld_DATE)
                                Last_Completed_Action_Type = oActions.Fields(ACTION_fld_TYPE)
                                FoundIt = True
                            End If
                        End If

                        If oActions.Fields(ACTION_fld_DATE) > Date And oActions.Fields(ACTION_fld_COMPLETED) = False And FoundIt3 = False Then
                            Next_Action_Date = oActions.Fields(ACTION_fld_DATE)
                            Next_Action_Type = oActions.Fields(ACTION_fld_TYPE)
                            FoundIt3 = True
                        End If
                    End If
                Next
            Next
        End If
        Debug.Print "here"
        
        Set oSolic = Nothing
        Set oActions = Nothing
                
        Dim oProspect As CProspect
        Dim oProposal As CProposal
        
        Dim With_Propose As Boolean
        Dim Ask_Amount As Currency
        Dim Expect_Amount As Currency
        Dim Target_Date As Date
        Dim Proposal_Pillar As String
        Dim With_Strategy As Boolean
        Dim Classification As String
        Dim Status As String
        
        With_Propose = False
        With_Strategy = False
        
        Dim tdate As Date
        
'        With oCons.Prospect
'            Debug.Print oCons.Prospect.Fields(PROSPECT_fld_CLASSIFICATION)
'
'            Classification = oCons.Prospect.Fields(PROSPECT_fld_CLASSIFICATION)
'            Status = oCons.Prospect.Fields(PROSPECT_fld_STATUS)
'
'            For Each oProposal In oCons.Prospect.Proposals
'                Debug.Print "Proposal"
'                tdate = Format(oProposal.Fields(PROPOSAL_fld_DATE_ADDED), "MM/DD/YYYY")
'                If tdate > "6/30/2005" Then
'                    Debug.Print "found it"
'                    With_Propose = True
'                    If oProposal.Fields(PROPOSAL_fld_AMOUNT_ASKED) > 0 Then
'                      Ask_Amount = oProposal.Fields(PROPOSAL_fld_AMOUNT_ASKED)
'                    End If
'                    If oProposal.Fields(PROPOSAL_fld_AMOUNT_EXPECTED) > 0 Then
'                        Expect_Amount = oProposal.Fields(PROPOSAL_fld_AMOUNT_EXPECTED)
'                    End If
''                    If oProposal.Fields(PROPOSAL_fld_DATE_RATED) <> Null Then
'                        Target_Date = oProposal.Fields(PROPOSAL_fld_DATE_RATED)
''                    End If
'                    Proposal_Pillar = oProposal.Fields(PROPOSAL_fld_PURPOSE)
'                End If
'                If Trim(oProposal.Fields(PROPOSAL_fld_NOTES)) <> "" Then
'                    With_Strategy = True
'                End If
'            Next
'
'        End With
        
        If FoundIt = True Then
            oRow.Field(3) = Last_Completed_Action_Date
            oRow.Field(4) = Last_Completed_Action_Type
        End If
        If FoundIt2 = True Then
            oRow.Field(5) = Last_Contact_Report_Date
            oRow.Field(6) = Last_Contact_Report_Description
        End If
        If FoundIt3 = True Then
            oRow.Field(7) = Next_Action_Date
            oRow.Field(8) = Next_Action_Type
        End If
        If With_Propose = True Then
            oRow.Field(9) = With_Propose
            oRow.Field(10) = Ask_Amount
            oRow.Field(11) = Expect_Amount
            oRow.Field(12) = Format(Target_Date, "MM/DD/YYYY")
            oRow.Field(13) = Proposal_Pillar
            oRow.Field(14) = With_Strategy
        End If
        oRow.Field(15) = Classification
        oRow.Field(16) = Status

        oCons.Closedown
        Set oCons = Nothing
    End If
End Sub

Public Sub Mark_Pat_Survivor(oRow As IBBQueryRow)
    Const ConsID = 1

    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
    
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init REApplication.SessionContext
        
        Dim oID As Long
        
        oID = oRow.Field(ConsID)
        oCons.Load oID
        
        Dim PSD As Date
        Dim PCD As Date
        Dim PCDS As Date
        
        PSD = Date
        PCD = "1/1/9999"
        PCDS = "1/1/9999"
        
        Dim oService As REServices
        Set oService = New REServices
        oService.Init REApplication.SessionContext
           
        Dim oAttributeServer As CAttributeTypeServer
        Set oAttributeServer = oService.CreateServiceObject(bbsoAttributeTypeServer)
        oAttributeServer.Init REApplication.SessionContext
        
        Dim Date2 As Long
        Dim Date3 As Long
        
        Date2 = oAttributeServer.GetAttributeTypeID("PC_Connection", bbAttributeRecordType_CONSTITUENT)
        Date3 = oAttributeServer.GetAttributeTypeID("PC_Connection (Staff)", bbAttributeRecordType_CONSTITUENT)
        
        Dim oAttribute As IBBAttribute

        For Each oAttribute In oCons.Attributes
            If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = Date2 And _
                oAttribute.Fields(Attribute_fld_VALUE) = "Patient/Survivor" Then
                If Trim(oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE)) <> "" Then
                    PCD = oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE)
                End If
            End If
            If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = Date3 And _
                oAttribute.Fields(Attribute_fld_VALUE) = "Patient/Survivor" Then
                If Trim(oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE)) <> "" Then
                    PCDS = oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE)
                End If
            End If
        Next oAttribute
        
        Set oAttribute = Nothing
        
        Debug.Print PCD & " " & PCDS
        If PSD > PCD Then PSD = PCD
        If PSD > PCDS Then PSD = PCDS
        Debug.Print PCD & " " & PCDS
'
        Dim oCCode As CConstituentCode
        Dim FoundIt As Boolean
        FoundIt = False
        
        For Each oCCode In oCons.ConstituentCodes
            If (oCCode.Fields(CONSTITUENT_CODE_fld_LONG_DESC_READONLY) = "Patient/Survivor") Then
                FoundIt = True
                With oCCode
                    .Fields(CONSTITUENT_CODE_fld_DATE_TO) = ""
                End With
            End If
        Next oCCode
'
        If FoundIt = False Then
            With oCons.ConstituentCodes.Add
                .Fields(CONSTITUENT_CODE_fld_CODE) = "Patient/Survivor"
                If PSD <= Date Then
                    .Fields(CONSTITUENT_CODE_fld_DATE_FROM) = PSD
                End If
            End With
        End If
        
        On Error Resume Next
        oCons.Save
'
        oCons.Closedown
        Set oCons = Nothing
        Set oCCode = Nothing
        
        oService.Closedown
        Set oAttributeServer = Nothing
        Set oAttribute = Nothing
        Set oService = Nothing
        
    End If

End Sub
Public Sub ProcessDeceasedAppend(oRow As IBBQueryRow)
    Const ConsID = 1

    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
    
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init REApplication.SessionContext
        
        Dim oID As Long
        
        oID = oRow.Field(ConsID)
        oCons.Load oID
        
        Dim oService As REServices
        Set oService = New REServices
        oService.Init REApplication.SessionContext
           
        Dim oAttributeServer As CAttributeTypeServer
        Set oAttributeServer = oService.CreateServiceObject(bbsoAttributeTypeServer)
        oAttributeServer.Init REApplication.SessionContext
        
        Dim AttID As Long
        Dim DeceasedDate As String
        
        AttID = oAttributeServer.GetAttributeTypeID("DeceasedFinder Result", bbAttributeRecordType_CONSTITUENT)
        
        Dim oAttribute As IBBAttribute
        
        DeceasedDate = ""

        For Each oAttribute In oCons.Attributes
            If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = AttID Then
                If Trim(oAttribute.Fields(Attribute_fld_VALUE)) <> "" Then
                    DeceasedDate = oAttribute.Fields(Attribute_fld_VALUE)
                End If
            End If
        Next oAttribute
        
        With oCons
            .Fields(RECORDS_fld_DECEASED) = True
            .Fields(RECORDS_fld_DECEASED_DATE) = DeceasedDate
            On Error Resume Next
            oCons.Save
        End With
        
        
        Set oAttribute = Nothing
        oCons.Closedown
        Set oCons = Nothing
    
        Set oAttributeServer = Nothing
        Set oAttribute = Nothing
        Set oService = Nothing
    
    End If

End Sub
Public Sub ChangeGivingScore(oRow As IBBQueryRow)
    Const ConsID = 1

    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
        Dim oService As REServices
        Set oService = New REServices
        oService.Init REApplication.SessionContext
            
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init REApplication.SessionContext
        
        Dim oID As Long
        
        oID = oRow.Field(ConsID)
        oCons.Load oID
        
        Debug.Print oCons.Fields(RECORDS_fld_LAST_NAME)
        Debug.Print oCons.Fields(RECORDS_fld_GIVING_SCORE_OVERRIDE)
        Debug.Print oCons.Fields(RECORDS_fld_GIVING_SCORE_STATUS)
        Debug.Print oCons.Fields(RECORDS_fld_GIVING_SCORE_Date)
        Debug.Print oCons.Fields(RECORDS_fld_USE_BB_GIVING_SCORE)
        
        With oCons
            oCons.Fields(RECORDS_fld_GIVING_SCORE_OVERRIDE) = "VIP"
            oCons.Fields(RECORDS_fld_GIVING_SCORE_OVERRIDE_DATE) = Date
            oCons.Fields(RECORDS_fld_USE_BB_GIVING_SCORE) = False
            oCons.Save
        End With
        
        oService.Closedown
        Set oService = Nothing
        Set oCons = Nothing
     End If
End Sub


Public Sub GetPledgeID(oRow As IBBQueryRow)
    Const Gift_ID = 1
    
    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
        Dim oGift As CGift
        Set oGift = New CGift
        oGift.Init RE7.SessionContext
        
        Dim lID As Long
        lID = oRow.Field(Gift_ID)
        
        Dim PledgeId As Long
        
        oGift.Load lID
        
        Debug.Print oGift.Fields(GIFT_fld_ID)
        
'        If oGift.Fields(GIFT_fld_GiftSubType) = "Pledge Payment" Then
            Dim oPledge As CGift
            Set oPledge = New CGift
            oPledge.Init REApplication.SessionContext
            
            PledgeId = oGift.PledgePayer.PledgeId(1)
            Debug.Print oGift.Fields(GIFT_fld_ID), PledgeId
            
            oRow.Field("Pledge ID") = PledgeId

            oPledge.Closedown
            Set oPledge = Nothing
'        End If
        
        oGift.Closedown
        Set oGift = Nothing
    End If
End Sub

Public Sub FixPreferredAddress(oRow As IBBQueryRow)
    Const ConsID = 1

    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
    
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init REApplication.SessionContext
        
        Dim oID As Long
        
        oID = oRow.Field(ConsID)
        oCons.Load oID
'
        Dim oAddress As CConstitAddress
        
        For Each oAddress In oCons.Addresses
            Debug.Print oAddress.Fields(CONSTIT_ADDRESS_fld_LAST_CHANGED_BY)
            If (oAddress.Fields(CONSTIT_ADDRESS_fld_DATE_TO) = "3/23/2018" Or _
                oAddress.Fields(CONSTIT_ADDRESS_fld_DATE_TO) = "3/24/2018") And _
                oAddress.Fields(CONSTIT_ADDRESS_fld_LAST_CHANGED_BY) = 348 Then
                Debug.Print "Fixing..."; oAddress.Fields(CONSTIT_ADDRESS_fld_ADDRESS_ID)
                With oAddress
                    oAddress.Fields(CONSTIT_ADDRESS_fld_DATE_TO) = ""
                    oAddress.Fields(CONSTIT_ADDRESS_fld_SENDMAIL) = True
                    oAddress.Fields(CONSTIT_ADDRESS_fld_PREFERRED) = True
                    On Error Resume Next
                    oCons.Save
                End With
            End If
        Next oAddress
        
        Set oAddress = Nothing
'
        oCons.Closedown
        Set oCons = Nothing
    End If
End Sub
Public Sub PostGiftAdjustment(oRow As IBBQueryRow)
    Const Gift_ID = 1
    
    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
        '
        ' load the gift record
        '
        Dim oGift As CGift
        Set oGift = New CGift
        oGift.Init RE7.SessionContext
        
        Dim lID As Long
        lID = oRow.Field(Gift_ID)
        
        oGift.Load lID

        If oGift.Fields(GIFT_fld_Post_Status) = "Posted" Then
            Debug.Print oGift.Fields(GIFT_fld_Amount)
'
' Search for any unposted adjustments and post it
'
            Dim oAdj As IBBAdjustment
            Dim oGLd As CGiftGLDistribution
            Dim Adj_Counter As Integer
            Dim Adj_Date As Date
            Adj_Date = oGift.Fields(GIFT_fld_Date)
            
            For Each oAdj In oGift.Adjustments
                If oAdj.Fields(ADJUSTMENT_fld_Post_Status) = "Not Posted" Then
                    If oAdj.Fields(ADJUSTMENT_fld_Date) > Adj_Date Then
                        Adj_Date = oAdj.Fields(ADJUSTMENT_fld_Date)
                    End If
                End If
            Next oAdj

    '        Debug.Print Adj_Date & " "; Adj_Counter

            Adj_Counter = 0

            For Each oAdj In oGift.Adjustments
                Adj_Counter = Adj_Counter + 1
                If oAdj.Fields(ADJUSTMENT_fld_Date) = Adj_Date And oAdj.Fields(ADJUSTMENT_fld_Post_Status) = "Not Posted" Then
                    Dim oAdjs As CAdjustmentServer
                    Set oAdjs = New CAdjustmentServer
                    oAdjs.Init REApplication.SessionContext, oGift
    '                Debug.Print oAdj.Fields(ADJUSTMENT_fld_Date) & " " & oAdj.Fields(ADJUSTMENT_fld_Post_Status) & " " & Adj_Date & " " & Adj_Counter
                    With oAdjs
                        Set oAdj = .EditAdjustment(oGift.Adjustments(Adj_Counter))
                        With oAdj
                            .Fields(ADJUSTMENT_fld_Post_Status) = "Posted"
                        End With
                        .Save
                        .Closedown
                    End With
                    Set oAdjs = Nothing
                End If
            Next oAdj
            Set oAdj = Nothing

'            On Error Resume Next
            oGift.Save

        End If
        
        oGift.Closedown
        Set oGift = Nothing

    End If
End Sub
Public Sub ExportGiftSolicitorInfo(oRow As IBBExportRow)
    Dim Const_ID As Long
    Const_ID = 13

    Const DevSol = 17
    Const PRM = 18
    Const SRM = 19
    Const PAC = 20
    Const GDate = 4
    
    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
        Dim DevSolicitor As String
        Dim SecondaryRelManager As String
        Dim PrimaryRelManager As String
        Dim OtherSolicitor As String
        Dim GiftDate As Date
        
        DevSolicitor = ""
        SecondaryRelManager = ""
        PrimaryRelManager = ""
        OtherSolicitor = ""
        GiftDate = oRow.Field(GDate)
        
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init RE7.SessionContext
        
        Dim lID As Long
        lID = oRow.Field(Const_ID)
        oCons.Load lID
        
        Debug.Print oCons.Fields(RECORDS_fld_CONSTITUENT_ID) & ": " & oCons.Fields(RECORDS_fld_PRIMARY_ADDRESSEE)
        
        Dim oSolicitor As CSolicitorActions
        Dim oASolicitor As CAssignedSolicitor
        Dim oSolID As CActionSolicitor
        
        Debug.Print "finding solicitor... "
        'loop through the constituent's assigned solicitors
        For Each oASolicitor In oCons.Relations.AssignedSolicitors
            If ((oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_DATE_FROM) <= GiftDate) And _
                (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_DATE_TO) >= GiftDate)) Or _
                ((oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_DATE_FROM) = "") And _
                (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_DATE_TO) >= GiftDate)) Or _
                ((oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_DATE_FROM) <= GiftDate) And _
                (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_DATE_TO) = "")) Then
                If (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) = "Development Solicitor") Then
                    DevSolicitor = oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_NAME)
                End If
                If (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) = "Primary Relationship Manager") Then
                    PrimaryRelManager = oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_NAME)
                End If
                If (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) = "Secondary Relationship Manager") Then
                    SecondaryRelManager = oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_NAME)
                End If
                If (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) = "Patient Central Associate") Then
                    SPatientCentralAss = oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_NAME)
                End If
                If (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) <> "Development Solicitor") And _
                    (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) <> "Primary Relationship Manager") And _
                    (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) <> "Secondary Relationship Manager") And _
                    (oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) <> "Patient Central Associate") Then
                    OtherSolicitor = oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_NAME)
                End If
            End If
        Next oASolicitor
'
        Debug.Print "storing... "
        oRow.Field(DevSol) = DevSolicitor
        oRow.Field(PRM) = PrimaryRelManager
        oRow.Field(SRM) = SecondaryRelManager
        oRow.Field(PAC) = SPatientCentralAss
              
        Set oASolicitor = Nothing
        Set oSolID = Nothing
        
        oCons.Closedown
        Set oCons = Nothing
    
    End If
End Sub
Public Sub DeleteConsAttribte(oRow As IBBQueryRow)
    Const ConsID = 1

    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
    
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init REApplication.SessionContext
        
        Dim oID As Long
        
        oID = oRow.Field(ConsID)
        oCons.Load oID
        
        Debug.Print oCons.Fields(RECORDS_fld_FIRST_NAME) & " " & oCons.Fields(RECORDS_fld_LAST_NAME)
'
        Dim oService As REServices
        Set oService = New REServices
        oService.Init REApplication.SessionContext
           
        Dim oAttributeServer As CAttributeTypeServer
        Set oAttributeServer = oService.CreateServiceObject(bbsoAttributeTypeServer)
        oAttributeServer.Init REApplication.SessionContext
        
        Dim lAttributeID As Long
    
        lAttributeID = oAttributeServer.GetAttributeTypeID("DCCR", bbAttributeRecordType_CONSTITUENT)
        
        Set oService = Nothing
        Set oAttributeServer = Nothing

        Dim oAttribute As IBBAttribute
        
        With oCons
            For Each oAttribute In oCons.Attributes
                Debug.Print oAttribute.Fields(Attribute_fld_VALUE)
                If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = lAttributeID And _
                    oAttribute.Fields(Attribute_fld_VALUE) = "Legacy Fund Participant" Then
                    Debug.Print "deleting..."
                    .Attributes.Remove oAttribute
        
'                    On Error Resume Next
                    .Save
                End If
            Next oAttribute
        End With
        
        Set oAttribute = Nothing
        oCons.Closedown
        Set oCons = Nothing
    End If

End Sub
Private Function fncIsPhone(ByVal strPhone As String) As Boolean
'    Const PhoneNum = "^([0-9]( |-)?)?(\(?[0-9]{3}\)?|[0-9]{3})( |-)?([0-9]{3}( |-)?[0-9]{4}|[a-zA-Z0-9]{7})$" & _
'                     "^\(*\+*[1-9]{0,3}\)*-*[1-9]{0,3}[-. /]*\(*[2-9]\d{2}\)*[-. /]*\d{3}[-. /]*\d{4} *e*x*t*\.* *\d{0,4}$"
    Const PhoneNum = "^\(*\+*[1-9]{0,3}\)*-*[1-9]{0,3}[-. /]*\(*[2-9]\d{2}\)*[-. /]*\d{3}[-. /]*\d{4} *e*x*t*\.* *\d{0,4}$"
    
    Dim objRegEx As Object
    On Error GoTo Fin
    Set objRegEx = CreateObject("Vbscript.Regexp")
    With objRegEx
        .Pattern = PhoneNum
        .IgnoreCase = True
        fncIsPhone = .Test(strPhone)
    End With
Fin:
    Set objRegEx = Nothing
    If Err.Number <> 0 Then MsgBox "Error: " & _
        Err.Number & " " & Err.Description
End Function

Public Sub Phone_Validator(oRow As IBBQueryRow)
    Const Const_ID = 1
    
    If oRow.BOF Then
        MsgBox "Beging processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
    
        Dim strPhone As String
        Dim ValidPhone As Boolean
    
        strPhone = Trim(oRow.Field("Home"))
                
        ValidPhone = fncIsPhone(strPhone)
    
        oRow.Field("Valid Phone") = ValidPhone
        
        If ValidPhone = False Then
            
            Dim oCons As CRecord
            Set oCons = New CRecord
            oCons.Init RE7.SessionContext
            
            Dim lID As Long
            
            lID = oRow.Field(Const_ID)
            oCons.Load lID

            Dim oPhone As IBBPhone
            Dim oAddress As CConstitAddress
            
            For Each oAddress In oCons.Addresses

                For Each oPhone In oAddress.Phones
                    If oPhone.Fields(Phone_fld_PhoneType) = "Home" And _
                       oPhone.Fields(Phone_fld_Num) = strPhone Then
'                       Debug.Print strPhone & " " & oPhone.Fields(Phone_fld_Num) & " "; oPhone.Fields(Phone_fld_PhoneType)
'                        With oPhone
'                            .Fields(Phone_fld_Inactive) = True
'                            On Error Resume Next
'                            oCons.Save
'                        End With
                    End If
                Next oPhone
            Next oAddress
        
            oCons.Closedown
            Set oCons = Nothing
            
        End If
             
    End If
End Sub
Public Sub UpdateRegFee(oRow As IBBQueryRow)
    Const Part_ID = 1
    
    If oRow.BOF Then
        MsgBox "Beging processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
        Dim oPart As CParticipant
        Set oPart = New CParticipant
        oPart.Init REApplication.SessionContext
        
        Dim lID As Long
        Dim RegDate As Date
        Dim UnitNo As Integer
        
        
        lID = oRow.Field(Part_ID)
        
        oPart.Load lID
        
        With oPart

            Dim oFee As CParticipantFee
             'loop through each fee and remove it
            
            For Each oFee In .Fees
                If oFee.Fields(ParticipantFees_fld_EventPricesID) = 5874 Then
                    Debug.Print oFee.Fields(ParticipantFees_fld_EventPricesID)
                    Debug.Print oFee.Fields(ParticipantFees_fld_ImportID)
                    Debug.Print oFee.Fields(ParticipantFees_fld_Unit)
                    Debug.Print oFee.Fields(ParticipantFees_fld_NumUnits)
                    Debug.Print oFee.Fields(ParticipantFees_fld_GiftAmount)
                    Debug.Print oFee.Fields(ParticipantFees_fld_ReceiptAmount)
                    RegDate = oFee.Fields(ParticipantFees_fld_FeeDate)
                    UnitNo = oFee.Fields(ParticipantFees_fld_NumUnits)
                    Debug.Print oFee.Fields(ParticipantFees_fld_ParticipantsID)
                    Debug.Print oFee.Fields(ParticipantFees_fld_ID)
                    .Fees.Remove oFee
                End If
            Next oFee
            
            With oPart.Fees.Add
                .Fields(ParticipantFees_fld_EventPricesID) = 5880
                .Fields(ParticipantFees_fld_NumUnits) = UnitNo
'                .Fields(ParticipantFees_fld_Unit) = "Free Registration Week"
'                .Fields(ParticipantFees_fld_GiftAmount) = 0
'                .Fields(ParticipantFees_fld_ReceiptAmount) = 0
                .Fields(ParticipantFees_fld_FeeDate) = RegDate
                .Fields(ParticipantFees_fld_Comments) = "Updated from public registration"
            End With

            'Save the participant
            oPart.Save

        End With
        
        'clean up
        oPart.Closedown
        Set oPart = Nothing
        Set oFee = Nothing

    End If

End Sub
Public Function GetAttributeID()

'
'   Get the attribute ID
'
    Dim oService As REServices
    Set oService = New REServices
    oService.Init REApplication.SessionContext
       
    Dim oAttributeServer As CAttributeTypeServer
    Set oAttributeServer = oService.CreateServiceObject(bbsoAttributeTypeServer)
    oAttributeServer.Init REApplication.SessionContext
    
    Dim AID1 As Long
    Dim AID2 As Long
    Dim AID3 As Long
    Dim AID4 As Long
    Dim AID5 As Long
    Dim AID6 As Long
    Dim AID7 As Long
    Dim AID8 As Long
    
    AID1 = oAttributeServer.GetAttributeTypeID("Grant Recipient", bbAttributeRecordType_CONSTITUENT)
    AID2 = oAttributeServer.GetAttributeTypeID("Researcher", bbAttributeRecordType_CONSTITUENT)
    AID3 = oAttributeServer.GetAttributeTypeID("PT Registry", bbAttributeRecordType_CONSTITUENT)
    AID4 = oAttributeServer.GetAttributeTypeID("DCCR", bbAttributeRecordType_CONSTITUENT)
    AID5 = oAttributeServer.GetAttributeTypeID("Industry", bbAttributeRecordType_CONSTITUENT)
    AID6 = oAttributeServer.GetAttributeTypeID("PC_Connection", bbAttributeRecordType_CONSTITUENT)
    AID7 = oAttributeServer.GetAttributeTypeID("Volunteer", bbAttributeRecordType_CONSTITUENT)
    AID8 = oAttributeServer.GetAttributeTypeID("Source/General", bbAttributeRecordType_CONSTITUENT)

    Debug.Print "ID1: " & AID1
    Debug.Print "ID2: " & AID2
    Debug.Print "ID3: " & AID3
    Debug.Print "ID4: " & AID4
    Debug.Print "ID5: " & AID5
    Debug.Print "ID6: " & AID6
    Debug.Print "ID7: " & AID7
    Debug.Print "ID8: " & AID8
        
    oService.Closedown
    Set oService = Nothing
    Set oAttributeServer = Nothing
    Set oAttribute = Nothing
      
End Function
Public Sub FoundationCleanUp(oRow As IBBQueryRow)
    Const ConsID = 1

    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
    
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init REApplication.SessionContext
'
        Dim oService As REServices
        Set oService = New REServices
        oService.Init REApplication.SessionContext
       
        Dim oAttributeServer As CAttributeTypeServer
        Set oAttributeServer = oService.CreateServiceObject(bbsoAttributeTypeServer)
        oAttributeServer.Init REApplication.SessionContext
    
        Dim FDID As Long

        FDID = oAttributeServer.GetAttributeTypeID("Foundation Category", bbAttributeRecordType_CONSTITUENT)
        Debug.Print FDID
'
        Dim oID As Long
        Dim FDN_Att As String
        FDN_Att = ""
        Dim FoundIt As Boolean
        
        oID = oRow.Field(ConsID)
        oCons.Load oID
'
        Dim oCCode As CConstituentCode
        Dim oAttribute As IBBAttribute
        
        For Each oCCode In oCons.ConstituentCodes
            FDN_Att = ""
            If oCCode.Fields(CONSTITUENT_CODE_fld_LONG_DESC_READONLY) = "Family Foundation" Or _
               oCCode.Fields(CONSTITUENT_CODE_fld_LONG_DESC_READONLY) = "Foundation-Family" Then
                FDN_Att = "Family Foundation"
            End If
            If oCCode.Fields(CONSTITUENT_CODE_fld_LONG_DESC_READONLY) = "Corporate Foundation" Or _
            oCCode.Fields(CONSTITUENT_CODE_fld_LONG_DESC_READONLY) = "Foundation-Corporate" Then
                FDN_Att = "Corporate Foundation"
            End If
            If oCCode.Fields(CONSTITUENT_CODE_fld_LONG_DESC_READONLY) = "Other Private Foundation" Then
                FDN_Att = "Other Private Foundation"
            End If
            If oCCode.Fields(CONSTITUENT_CODE_fld_LONG_DESC_READONLY) = "Public Foundation" Then
                FDN_Att = "Public Foundation"
            End If
            
            If FDN_Att <> "" Then
                FoundIt = False
            
                With oCons
                    For Each oAttribute In oCons.Attributes
                        If oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = FDID Then
                            oAttribute.Fields(Attribute_fld_VALUE) = FDN_Att
                            On Error Resume Next
                            oCons.Save
                        End If
                    Next oAttribute
            
                    If FoundIt = False Then
                        Set oAttribute = oCons.Attributes.Add
                        oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = FDID
                        oAttribute.Fields(Attribute_fld_VALUE) = FDN_Att
                        
                        On Error Resume Next
                        oCons.Save
                    End If
                End With
            End If
            
        Next oCCode

        Set oAttribute = Nothing
        Set oCCode = Nothing
'
        oService.Closedown
        Set oService = Nothing
        Set oAttributeServer = Nothing
        Set oAttribute = Nothing
        
        oCons.Closedown
        Set oCons = Nothing
        
    End If

End Sub
Public Sub AddDCCR_NonCorp(oRow As IBBQueryRow)
    Const ConsID = 1

    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
    
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init REApplication.SessionContext
'
        Dim oService As REServices
        Set oService = New REServices
        oService.Init REApplication.SessionContext
       
        Dim oAttributeServer As CAttributeTypeServer
        Set oAttributeServer = oService.CreateServiceObject(bbsoAttributeTypeServer)
        oAttributeServer.Init REApplication.SessionContext
    
        Dim DCCRID As Long
        DCCRID = oAttributeServer.GetAttributeTypeID("DCCR", bbAttributeRecordType_CONSTITUENT)

'
        Dim oID As Long
        Dim FDN_Att As String
        FDN_Att = ""
        Dim FoundIt As Boolean
        
        oID = oRow.Field(ConsID)
        oCons.Load oID
'
        Dim oAttribute As IBBAttribute
            
        With oCons
            Set oAttribute = oCons.Attributes.Add
            oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = DCCRID
            oAttribute.Fields(Attribute_fld_VALUE) = "Non-Corporate Organization (Ind)"
            oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE) = Date
            
            On Error Resume Next
            oCons.Save
        End With

        Set oAttribute = Nothing
'
        oService.Closedown
        Set oService = Nothing
        Set oAttributeServer = Nothing
        Set oAttribute = Nothing
        
        oCons.Closedown
        Set oCons = Nothing
        
    End If

End Sub
Public Sub FixSpecificAppeal(oRow As IBBQueryRow)
Const ConsID = 1

    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
    
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init REApplication.SessionContext
        
        Dim oID As Long
        
        oID = oRow.Field(ConsID)
        oCons.Load oID
        
        Dim oAppeal As CConstitAppeal
        
        For Each oAppeal In oCons.Appeals
            Debug.Print oAppeal.Fields(CONSTITUENT_APPEALS_fld_Appeal)
            If (oAppeal.Fields(CONSTITUENT_APPEALS_fld_Appeal) = "2018 Grateful Patient Program" Or _
                oAppeal.Fields(CONSTITUENT_APPEALS_fld_Appeal) = "2018 Lapsed Donors" Or _
                oAppeal.Fields(CONSTITUENT_APPEALS_fld_Appeal) = "2018 Anniversary") Then
                Debug.Print oAppeal.Fields(CONSTITUENT_APPEALS_fld_DATE)
                If oAppeal.Fields(CONSTITUENT_APPEALS_fld_DATE) = "5/6/2018" Then
                    Debug.Print "Fixing..."
                    With oAppeal
                        oAppeal.Fields(CONSTITUENT_APPEALS_fld_DATE) = "6/6/2018"
                    End With
'                    On Error Resume Next
                    oCons.Save
                End If
            End If
        Next oAppeal
        
        oCons.Closedown
        Set oCons = Nothing
    End If

End Sub
Public Sub AddDCCR_Pass_Through_Org(oRow As IBBQueryRow)
    Const ConsID = 1

    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
    
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init REApplication.SessionContext
'
        Dim oService As REServices
        Set oService = New REServices
        oService.Init REApplication.SessionContext
       
        Dim oAttributeServer As CAttributeTypeServer
        Set oAttributeServer = oService.CreateServiceObject(bbsoAttributeTypeServer)
        oAttributeServer.Init REApplication.SessionContext
    
        Dim DCCRID As Long
        DCCRID = oAttributeServer.GetAttributeTypeID("DCCR", bbAttributeRecordType_CONSTITUENT)

'
        Dim oID As Long
        Dim FDN_Att As String
        FDN_Att = ""
        Dim FoundIt As Boolean
        
        oID = oRow.Field(ConsID)
        oCons.Load oID
'
        Dim oAttribute As IBBAttribute
            
        With oCons
            Set oAttribute = oCons.Attributes.Add
            oAttribute.Fields(Attribute_fld_ATTRIBUTETYPES_ID) = DCCRID
            oAttribute.Fields(Attribute_fld_VALUE) = "Pass-Through Organization"
            oAttribute.Fields(Attribute_fld_ATTRIBUTEDATE) = Date
            
            On Error Resume Next
            oCons.Save
        End With

        Set oAttribute = Nothing
'
        oService.Closedown
        Set oService = Nothing
        Set oAttributeServer = Nothing
        Set oAttribute = Nothing
        
        oCons.Closedown
        Set oCons = Nothing
        
    End If

End Sub
Public Sub CleanUp_FormerSolicitor_Rel(oRow As IBBQueryRow)
    
    Dim Const_ID As Long
    
    Const_ID = 1
    
    If oRow.BOF Then
        MsgBox "Begin processing"
    ElseIf oRow.EOF Then
        MsgBox "End processing"
    Else
        Dim oCons As CRecord
        Set oCons = New CRecord
        oCons.Init RE7.SessionContext
        
        Dim lID As Long
        Dim sID As Long
        Dim sType As String
        
        lID = oRow.Field(Const_ID)
        oCons.Load lID
        
        Debug.Print oCons.Fields(RECORDS_fld_PRIMARY_ADDRESSEE)

        Dim oASolicitor As CAssignedSolicitor
        
        'loop through the constituent's assigned solicitors
        For Each oASolicitor In oCons.Relations.AssignedSolicitors
            If oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_DATE_TO) <> "" Then
                sType = ""
                If oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) = "Development Solicitor" Then
                    sType = "Previous Solicitor"
                End If
                If oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) = "Primary Relationship Manager" Or _
                oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) = "Secondary Relationship Manager" Then
                    sType = "Previous Relationship Manager"
                End If
                If sType <> "" Then
                    With oCons
                        oASolicitor.Fields(ASSIGNEDSOLICITOR_fld_SOLICITOR_TYPE) = sType
                        .Save
                    End With
                End If
            End If
        Next oASolicitor
        
        Set oASolicitor = Nothing
        
        oCons.Closedown
        Set oCons = Nothing
    
    End If
End Sub


