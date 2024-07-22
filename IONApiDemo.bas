Attribute VB_Name = "IONApiDemo"
Option Explicit
''' Requires: VBA-JSON (https://github.com/VBA-tools/VBA-JSON)
''' Requires: Reference "Microsoft Scripting Runtime"

Public Sub TestMICall()
    Dim APIProgram, APITransaction As String
    Dim APIParams As Dictionary
    
    Set APIParams = New Dictionary
    
    Debug.Print ExecuteMI("MNS150MI", "SelUsers", APIParams)
    
End Sub

Public Function ExecuteMI(APIProgram As String, APITransaction As String, APIParams As Dictionary)
    
    On Error GoTo Error_Handler
    
    Dim BearerTokenString, BaseUrl, MIUrl, MIResult As String
    Dim IONAuthDetails As Dictionary
    
    Set IONAuthDetails = GetIONAuthDetails
    
    BearerTokenString = "Bearer " & IONAuthDetails("BearerToken")
    BaseUrl = IONAuthDetails("IONUrl") & "/" & IONAuthDetails("TenantInformation") & "/"
    MIUrl = BaseUrl & "M3/m3api-rest/execute/" & APIProgram & "/" & APITransaction
              
    With CreateObject("MSXML2.XMLHTTP")
        .Open "GET", MIUrl, False
        .setRequestHeader "Content-type", "application/x-www-form-urlencoded"
        .setRequestHeader "Accept", "application/json"
        .setRequestHeader "Accept-Language", "en_US"
        .setRequestHeader "Authorization", BearerTokenString
        .send
        MIResult = .ResponseText
    End With
    
    ExecuteMI = MIResult

Error_Handler_Exit:
    On Error Resume Next
    Exit Function
    
Error_Handler:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: Modul1/ExecuteMI" & vbCrLf & _
           "Error Description: " & Err.Description, vbCritical, _
           "An Error has Occurred!"
    Resume Error_Handler_Exit
End Function

Public Function GetIONAuthDetails() As Dictionary
    
    On Error GoTo Error_Handler
    Dim TokenUrl, BearerTokenJson, BearerToken As String
    Dim Credentials, IONAuthDetails, BearerTokenResonse As Dictionary
    
    Set Credentials = ReadIONCredentialsFromFile
    Set IONAuthDetails = New Dictionary
    
    TokenUrl = Credentials("AccessTokenBaseUrl") + Credentials("AccessTokenBaseUrlSuffix")
    
    With CreateObject("MSXML2.XMLHTTP")
         .Open "POST", TokenUrl, False
         .setRequestHeader "Content-type", "application/x-www-form-urlencoded"
         .setRequestHeader "Accept", "application/json"
         .setRequestHeader "Accept-Language", "en_US"
         .send "grant_type=password&client_id=" & Credentials("ClientID") & "&client_secret=" & Credentials("ClientSecret") & "&username=" & Credentials("Username") & "&password=" & Credentials("Password")
         BearerTokenJson = .ResponseText
    End With
    
    ''' Crude check if valid response is returned
    If Left(BearerTokenJson, 17) <> "{""access_token"":""" Then
        Err.Raise vbObjectError + 1000, "", "No bearer token received"
    End If
    
    Set BearerTokenResonse = JsonConverter.ParseJson(BearerTokenJson)
    
    IONAuthDetails("BearerToken") = BearerTokenResonse("access_token")
    IONAuthDetails("IONUrl") = Credentials("IONUrl")
    IONAuthDetails("TenantInformation") = Credentials("TenantInformation")
    
    'Debug.Print IONAuthDetails("BearerToken")
    'Debug.Print IONAuthDetails("IONUrl")
    
    Set GetIONAuthDetails = IONAuthDetails
    
Error_Handler_Exit:
    On Error Resume Next
    Exit Function
    
Error_Handler:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: Modul1/GetIONAuthDetails" & vbCrLf & _
           "Error Description: " & Err.Description, vbCritical, _
           "An Error has Occurred!"
    Resume Error_Handler_Exit

End Function

Public Function ReadIONCredentialsFromFile(Optional CredFileFullPath As String = "unset") As Dictionary
   
    On Error GoTo Error_Handler
    Dim FSO As New FileSystemObject
    Dim JsonTS As TextStream
    Dim JsonText As String
    Dim Parsed As Dictionary
    Dim CredFileFieldTranslation, Credentials As Dictionary
    Dim i As Integer

    Set CredFileFieldTranslation = New Scripting.Dictionary
    Set Credentials = New Scripting.Dictionary
    
    If CredFileFullPath = "unset" Then
        CredFileFullPath = Application.ActiveWorkbook.Path & "\" & "CredFile.ionapi"
    End If
    
    'Debug.Print CredFileFullPath
    
    ''' Finding worksheet path will not work if located on sharepoint
    If Left(CredFileFullPath, 8) = "https://" Then
        Err.Raise vbObjectError + 1000, "", "Cannot find credfile, please do not place this folder on Ondrive"
    End If
    
    ''' Field name mapping
    CredFileFieldTranslation("ti") = "TenantInformation"
    'CredFileFieldTranslation("cn") =
    'CredFileFieldTranslation("dt") =
    CredFileFieldTranslation("ci") = "ClientID"
    CredFileFieldTranslation("cs") = "ClientSecret"
    CredFileFieldTranslation("iu") = "IONUrl"
    CredFileFieldTranslation("pu") = "AccessTokenBaseUrl"
    'CredFileFieldTranslation("oa") =
    CredFileFieldTranslation("ot") = "AccessTokenBaseUrlSuffix"
    'CredFileFieldTranslation("or") =
    'CredFileFieldTranslation("ev") =
    'CredFileFieldTranslation("v") =
    CredFileFieldTranslation("saak") = "Username"
    CredFileFieldTranslation("sask") = "Password"
    
    ''' Read .json file
    Set JsonTS = FSO.OpenTextFile(CredFileFullPath, ForReading)
    JsonText = JsonTS.ReadAll
    JsonTS.Close
    Set Parsed = JsonConverter.ParseJson(JsonText)
    
    ''' Fill necessary fields into dictionary variable with translated keys
    i = 0
    For i = 0 To Parsed.Count - 1
        If CredFileFieldTranslation.Exists(Parsed.Keys(i)) Then
            Credentials(CredFileFieldTranslation(Parsed.Keys(i))) = Parsed.Items(i)
        End If
    Next
    
    Set ReadIONCredentialsFromFile = Credentials
    
Error_Handler_Exit:
    On Error Resume Next
    Exit Function
    
Error_Handler:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: Modul1/ReadIONCredentialsFromFile" & vbCrLf & _
           "Error Description: " & Err.Description, vbCritical, _
           "An Error has Occurred!"
    Resume Error_Handler_Exit

End Function
