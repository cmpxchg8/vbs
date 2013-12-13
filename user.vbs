'
'  The MIT License:
'
'  Copyright (c) May 2012 Kevin Devine
'
'  Permission is hereby granted,  free of charge,  to any person obtaining a copy
'  of this software and associated documentation files (the "Software"),  to deal
'  in the Software without restriction,  including without limitation the rights
'  to use,  copy,  modify,  merge,  publish,  distribute,  sublicense,  and/or sell
'  copies of the Software,  and to permit persons to whom the Software is
'  furnished to do so,  subject to the following conditions:
'
'  The above copyright notice and this permission notice shall be included in
'  all copies or substantial portions of the Software.
'
'  THE SOFTWARE IS PROVIDED "AS IS",  WITHOUT WARRANTY OF ANY KIND,  EXPRESS OR
'  IMPLIED,  INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,  DAMAGES OR OTHER
'  LIABILITY,  WHETHER IN AN ACTION OF CONTRACT,  TORT OR OTHERWISE,  ARISING FROM,
'  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
'  THE SOFTWARE.
'
  
  Option Explicit
  
  Const DEBUG_MODE = False
  
  ' ADS_AUTHENTICATION_ENUM
  Const ADS_SECURE_AUTHENTICATION   = &H01
  Const ADS_USE_ENCRYPTION          = &H02
  Const ADS_USE_SSL                 = &H02
  Const ADS_READONLY_SERVER         = &H04
  Const ADS_PROMPT_CREDENTIALS      = &H08
  Const ADS_NO_AUTHENTICATION       = &H010
  Const ADS_FAST_BIND               = &H020
  Const ADS_USE_SIGNING             = &H040
  Const ADS_USE_SEALING             = &H080
  Const ADS_USE_DELEGATION          = &H0100
  Const ADS_SERVER_BIND             = &H0200
  Const ADS_NO_REFERRAL_CHASING     = &H0400
  Const ADS_AUTH_RESERVED           = &H080000000 
  
  
  ' ADS_OPTION_ENUM
  Const ADS_OPTION_SERVERNAME                  = 0
  Const ADS_OPTION_REFERRALS                   = 1
  Const ADS_OPTION_PAGE_SIZE                   = 2
  Const ADS_OPTION_SECURITY_MASK               = 3
  Const ADS_OPTION_MUTUAL_AUTH_STATUS          = 4
  Const ADS_OPTION_QUOTA                       = 5
  Const ADS_OPTION_PASSWORD_PORTNUMBER         = 6
  Const ADS_OPTION_PASSWORD_METHOD             = 7
  Const ADS_OPTION_ACCUMULATIVE_MODIFICATION   = 8
  Const ADS_OPTION_SKIP_SID_LOOKUP             = 9
  
  ' ADS_CHASE_REFERRALS_ENUM
  Const ADS_CHASE_REFERRALS_NEVER         = &H000
  Const ADS_CHASE_REFERRALS_SUBORDINATE   = &H020
  Const ADS_CHASE_REFERRALS_EXTERNAL      = &H040
  Const ADS_CHASE_REFERRALS_ALWAYS        = &H060

  Const ADS_SCOPE_BASE       = 0
  Const ADS_SCOPE_ONELEVEL   = 1
  Const ADS_SCOPE_SUBTREE    = 2
  
  Const SAM_GROUP_OBJECT              = &H10000000
  Const SAM_NON_SECURITY_GROUP_OBJECT = &H10000001
  Const SAM_ALIAS_OBJECT              = &H20000000
  Const SAM_USER_OBJECT               = &H30000000
  Const SAM_NORMAL_USER_ACCOUNT       = &H30000000 
  Const SAM_MACHINE_OBJECT            = &H30000001
  Const SAM_TRUST_ACCOUNT             = &H30000002
  
  ' UserAccountControl attributes
  Const ADS_UF_SCRIPT                          = &H000000001   ' The logon script is executed. 
  Const ADS_UF_ACCOUNTDISABLE                  = &H000000002   ' The user account is disabled. 
  Const ADS_UF_HOMEDIR_REQUIRED                = &H000000008   ' The home directory is required. 
  Const ADS_UF_LOCKOUT                         = &H000000010   ' The account is currently locked out. 
  Const ADS_UF_PASSWD_NOTREQD                  = &H000000020   ' No password is required. 

  Const ADS_UF_PASSWD_CANT_CHANGE              = &H000000040   ' The user cannot change the password. 
  ' Note  You cannot assign the permission settings of PASSWD_CANT_CHANGE by directly modifying the UserAccountControl attribute. 
  ' For more information and a code example that shows how to prevent a user from changing the password, see User Cannot Change Password.

  Const ADS_UF_ENCRYPTED_TEXT_PASSWORD_ALLOWED = &H000000080   'The user can send an encrypted password. 

  Const ADS_UF_TEMP_DUPLICATE_ACCOUNT          = &H000000100   'This is an account for users whose primary account is in another domain. 
  ' This account provides user access to this domain, but not to any domain that trusts this domain. Also known as a local user account. 

  '0x00000200 ADS_UF_NORMAL_ACCOUNT  This is a default account type that represents a typical user. 
  '0x00000800 ADS_UF_INTERDOMAIN_TRUST_ACCOUNT  This is a permit to trust account for a system domain that trusts other domains. 
  '0x00001000 ADS_UF_WORKSTATION_TRUST_ACCOUNT  This is a computer account for a computer that is a member of this domain. 
  '0x00002000 ADS_UF_SERVER_TRUST_ACCOUNT  This is a computer account for a system backup domain controller that is a member of this domain. 
  '0x00004000 N/A Not used. 
  '0x00008000 N/A Not used. 
  Const ADS_UF_DONT_EXPIRE_PASSWD = &H000010000 'The password for this account will never expire. 
  '0x00020000 ADS_UF_MNS_LOGON_ACCOUNT  This is an MNS logon account. 
  '0x00040000 ADS_UF_SMARTCARD_REQUIRED  The user must log on using a smart card. 
  '0x00080000 ADS_UF_TRUSTED_FOR_DELEGATION  The service account (user or computer account), under which a service runs, is trusted for Kerberos delegation. Any such service can impersonate a client requesting the service. 
  '0x00100000 ADS_UF_NOT_DELEGATED  The security context of the user will not be delegated to a service even if the service account is set as trusted for Kerberos delegation. 
  '0x00200000 ADS_UF_USE_DES_KEY_ONLY  Restrict this principal to use only Data Encryption Standard (DES) encryption types for keys. 
  '0x00400000 ADS_UF_DONT_REQUIRE_PREAUTH  This account does not require Kerberos pre-authentication for logon. 
  Const ADS_UF_PASSWORD_EXPIRED                = &H000800000   'The user password has expired. 
  'This flag is created by the system using data from the Pwd-Last-Set attribute and the domain policy. 
  '0x01000000 ADS_UF_TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION  The account is enabled for delegation. This is a security-sensitive setting; accounts with this option enabled should be strictly controlled. This setting enables a service running under the account to assume a client identity and authenticate as that user to other remote servers on the network. 

  '
  ' global variables
  '
  Dim host      : host     = vbNullString 
  Dim username  : username = vbNullString
  Dim password  : password = vbNullString
  Dim maxPwdAge : maxPwdAge = 0
  
  Dim bReports  : bReports  = False
  Dim bGroups   : bGroups   = False
  Dim bExchange : bExchange = False
  
  Call Main(WScript.Arguments.Count, WScript.Arguments)
  Call WScript.Quit()
  
  '
  ' required by Main sub routine to parse cmd line arguments
  '
  Function GetParameter(ByVal argc, ByVal argv, ByRef index)
    If (index + 1 < argc) Then
      index = index + 1
      GetParameter = argv(index)
    Else
      printf(Array("\n  [-] Error: No parameter specified to %s\n", argv(index)))
      WScript.Quit()
    End If
  End Function
  
  '
  ' display basic usage information
  '
  Sub usage()
    printf(Array("\n  %s", String(80, "*")))
    printf(Array("\n  * %-76s *", " "))
    printf(Array("\n  * %-76s *", "User v1.0 Copyright (c) 2012 Kevin Devine"))
    printf(Array("\n  * %-76s *", "Display information about user in active directory"))
    printf(Array("\n  * %-76s *", " "))
    'printf(Array("\n  * %-76s *", "Use /? switch for more information"))
    'printf(Array("\n  * %-76s *", " "))
    printf(Array("\n  %s\n\n", String(80, "*")))
    
    printf(Array("\n  USAGE: User -r -g <domain\msid> or <domain\firstname lastname> or <domain\employee id>\n"))
    printf(Array("\n        -r will display information about people that report directly to the user."))
    printf(Array("\n        -g will display the global groups user is part of."))
    printf(Array("\n        -o will display exchange/outlook info\n"))
    printf(Array("\n  Example: User ms\joe bloggs\n\n"))
    printf(Array("\n  Send comments, questions to author\n\n"))
  End Sub
  
  '
  ' :|
  ' 
  Function IsValidOption(ByVal arg)
    If (arg = "-o" Or arg = "-r" Or arg = "-g" Or arg = "-?" Or arg = "/?" Or arg = "-h" Or arg = "--help") Then
      IsValidOption = True
    Else
      IsValidOption = False
    End If
  End Function
  
  '
  ' main sub routine
  '
  Sub Main(ByVal argc, ByVal argv)
    
    Dim domain, firstname, lastname, ntid_or_employeeid, employeeid, displayname, dn_list
    
    Dim idx, arg, params
    
    For idx = 0 To argc - 1
      arg = argv(idx)
          
      If (arg = "-r") Then  ' display directReport values
        bReports = True
      ElseIf (arg = "-g") Then ' display memberOf values
        bGroups = True
      ElseIf (arg = "-o") Then
        bExchange = True
      ElseIf (arg = "-?" Or arg = "/?" Or arg = "-h" Or arg = "--help") Then
        usage()
        Exit Sub
      Else
        firstname = arg
        params = 1
        If (idx + 1 < argc) Then
          arg = argv(idx + 1)
          If (Not IsValidOption(arg)) Then
            idx = idx + 1
            lastname = arg
            params = 2
          End If
        End If
      End If
    Next
    
    '
    ' if 2 arguments supplied, assume it's first and last name
    '
    If (params = 1) Then      
      ntid_or_employeeid = firstname
      domain = GetDomain(ntid_or_employeeid)
      
      ' search firstname by what user has provided
      dprintf(Array("\n  [+] Searching for %s", ntid_or_employeeid))
      dn_list = SearchByNTIDOrEmployeeID(domain, ntid_or_employeeid)
      
      ProcessList(dn_list)
    ElseIf (params = 2) Then    
      domain = GetDomain(firstname)
      
      ' search by "givenName" and "sn" using input provided
      dprintf(Array("\n  [+] Searching for %s %s", firstname, lastname))
      dn_list = SearchByFirstAndLastName(domain, firstname, lastname)
    
      If (VarType(dn_list) <> vbArray + vbVariant) Then
        
        ' if no results returned, search by "displayName"
        dprintf(Array("\n  [+] Searching for %s, %s", lastname, firstname))
        dn_list = SearchByDisplayName(domain, firstname, lastname)
        
        If (VarType(dn_list) <> vbArray + vbVariant) Then
          
          ' the person might have an initial in their name
          ' search by displayName but wildcard on end of first name.
          dprintf(Array("\n  [+] Searching for %s, %s*", lastname, firstname))
          dn_list = SearchByDisplayName(domain, firstname & "*", lastname)
          
          If (VarType(dn_list) <> vbArray + vbVariant) Then
            
            ' if still nothing found, as a last resort, search by "givenName" and "sn" again 
            ' but this time only using first 3 letters of first name
            '
            ' this is to handle cases where someone called "Samuel" can be entered in database
            ' as "Sam" or "Sammy" ..other examples would be "Marsh" for "Marshall"
            '
            dprintf(Array("\n  [+] Searching for %s* %s\n", Left(firstname, 3), lastname))
            dn_list = SearchByFirstAndLastName(domain, Left(firstname, 3) & "*", lastname)
          End If
        End If
      End If
      
      ProcessList(dn_list)
    Else
      usage()
    End If    
  End Sub
  
  '
  ' process a list of distinguished names
  '
  Function ProcessList(ByVal dn_list)
    If (VarType(dn_list) = vbArray + vbVariant) Then
      If (UBound(dn_list) > 1) Then
        printf(Array("\n[+] Found %i records\n\n", UBound(dn_list)))
          
        printf(Array("   %-25s\t%-15s\t%-10s\t%-20s\t%-25s\t%s\n", "Full Name", "Employee ID", "MS ID", "Phone", "E-Mail", "Post Code"))
        printf(Array("   %-25s\t%-15s\t%-10s\t%-20s\t%-25s\t%s\n", String(25,"*"), String(12,"*"), String(10,"*"), String(20,"*"), String(25,"*"), String(10,"*")))
    
        Dim idx
        For idx = 0 To UBound(dn_list) - 1 
          Call DumpUserBasic(idx, dn_list(idx))
        Next
      Else
        DumpUser(dn_list(0))
      End If      
    Else
      printf(Array("\nNothing found\n"))
    End If
  End Function

  '
  '
  '
  Function GetDomain(ByRef input)
    ' domain\ntid format?
    If (InStr(input, "\") <> 0) Then
      Dim aInput : aInput = Split(input, "\")
          
      GetDomain = aInput(0)
      input     = aInput(1)
    Else
      GetDomain = CreateObject("WScript.Network").UserDomain
    End If
  End Function
    
  '
  '
  '
  Function DumpUserBasic(ByVal idx, ByVal dn)
    On Error Resume Next
    
    Dim ldap
    Set ldap = GetObject("LDAP:")
    
    If (VarType(ldap) <> vbObject) Then
      Exit Function
    End If
    
    Dim user
    Set user = ldap.OpenDSObject(dn, vbNullString, vbNullString, ADS_SERVER_BIND Or ADS_SECURE_AUTHENTICATION Or ADS_USE_SEALING Or ADS_USE_SIGNING)
    
    If (VarType(user) <> vbObject) Then
      Exit Function
    End If
    
    Dim fullName : fullName = GetValue(user, "givenName") & " " & GetValue(user,"sn")
    Dim id       : id       = GetValue(user, "employeeID")
    Dim ntid     : ntid     = GetValue(user, "sAMAccountName")
    Dim mail     : mail     = GetValue(user, "mail")
    Dim phone    : phone    = GetValue(user, "telephoneNumber")
    Dim created  : created  = GetValue(user, "whenCreated")
    Dim postCode : postCode = GetValue(user, "postalCode")
      
    printf(Array("[%2i]: %-25s\t%-15s\t%-10s\t%-20s\t%-25s\t%s\n", (idx + 1), fullName, id, UCase(ntid), phone, mail, postCode))   
    
    Set user = Nothing
    Set ldap = Nothing
  End Function
 
Const ONE_HUNDRED_NANOSECOND    = .000000100
Const SECONDS_IN_DAY            = 86400

'
' 
' 
Function DumpUser(ByVal dn)
  On Error Resume Next
    
  Dim ldap
  Set ldap = GetObject("LDAP:")
    
  If (VarType(ldap) <> vbObject) Then
    Exit Function
  End If
    
  Dim user
  Set user = ldap.OpenDSObject(dn, vbNullString, vbNullString, ADS_SERVER_BIND Or ADS_SECURE_AUTHENTICATION Or ADS_USE_SEALING Or ADS_USE_SIGNING)
    
  If (VarType(user) <> vbObject) Then
    Exit Function
  End If
  
  Dim ID          ' employeeID
  Dim NTID        ' sAMAccountName
  Dim Display     ' displayName
  Dim First       ' givenName
  Dim Last        ' sn
  Dim EMail       ' mail
  Dim Created     ' whenCreated
  Dim Phone       ' telephoneNumber
  Dim Office      ' physicalDeliveryOfficeName
  Dim PostCode    ' postalCode
  Dim Description ' description
  Dim Company     ' company
  Dim Country     ' co
  Dim Department  ' department
  Dim HomePath    ' homeDirectory
  Dim HomeDrive   ' homeDrive
  Dim Manager     ' manager
  Dim SipID       ' msRTCSIP-PrimaryUserAddress
  Dim NickName    ' mailNickName
  Dim MSID        ' uht-Migration-SAMAccountName
  Dim UNIXShell   ' loginShell
  Dim UNIXHome    ' unixHomeDirectory
  Dim OCUser      ' msRTCSIP-UserEnabled
  Dim Expired     ' userAccountControl
  Dim Locked      ' ""
  Dim Disabled    ' ""
  Dim WillExpire  ' ""
  Dim Employee    ' employeeType
  Dim HomePage    ' wWWHomePage
  
  Dim PwdLastSet  ' pwdLastSet
  Dim PwdExpires  ' pwdLastSet + maxPwdAge
  
  Dim ScriptPath  ' scriptPath
  
  
  DN          = GetValue(user, "distinguishedName")
  ID          = GetValue(user, "employeeID")
  NTID        = GetValue(user, "sAMAccountName")
  Display     = GetValue(user, "displayName")
  First       = GetValue(user, "givenName")
  Last        = GetValue(user, "sn")
  EMail       = GetValue(user, "mail")
  Created     = GetValue(user, "whenCreated")
  Phone       = GetValue(user, "telephoneNumber")
  Office      = GetValue(user, "physicalDeliveryOfficeName")
  PostCode    = GetValue(user, "postalCode")
  Description = GetValue(user, "description")
  ScriptPath  = GetValue(user, "scriptPath")
 
  UNIXShell   = GetValue(user, "loginShell")
  If (UNIXShell <> "<unlisted>") Then
    UNIXShell = UCase(UNIXShell)
  End If
  
  UNIXHome    = GetValue(user, "unixHomeDirectory")
  If (UNIXHome <> "<unlisted>") Then
    UNIXHome = UCase(UNIXHome)
  End If
  
  MSID        = GetValue(user, "uht-Migration-SAMAccountName")
  
  If (VarType(Description) <> vbString) Then
    Dim t : t = Description(0)
    Description = t
  End If
  
  Company     = GetValue(user, "company")
  Country     = GetValue(user, "co")
  Department  = GetValue(user, "department")
  HomePath    = GetValue(user, "homeDirectory")
  
  If (HomePath <> "<unlisted>") Then
    HomePath = UCase(HomePath)
  End If
  
  HomeDrive   = GetValue(user, "homeDrive")
  
  If (HomeDrive <> "<unlisted>") Then
    HomeDrive = UCase(HomeDrive)
  End If
  
  ' check for SIP account
  '
  SipID       = GetValue(user, "msRTCSIP-PrimaryUserAddress")
  
  If (InStr(SipID, "sip:") <> 0) Then
    SipID = Right(SipID, Len(SipID) - 4)
    
    ' Is the user using Personal Communicator or Office Communicator?
    ' thanks to Sam Frost
    '
    OCUser = GetValue(user, "msRTCSIP-UserEnabled")
    
    If (OCUser = "False") Then
      SipID = SipID & " ( Cisco Personal Communicator )"
    Else
      SipID = SipID & " ( Microsoft Office Communicator ) "
    End If
  End If
  
  NickName    = GetValue(user, "mailNickName")
  Manager     = GetValue(user, "Manager")
  
  If (Manager <> "<unlisted>") Then
    Dim obj_man  : Set obj_man = GetObject("LDAP://" & Manager)
    Dim man_mail : man_mail = obj_man.mail
    
    If (VarType(man_mail) <> vbString) Then
      man_mail   = "<unlisted>"
    End If
    
    Dim man_first   : man_first   = obj_man.givenName
    Dim man_last    : man_last    = obj_man.sn
    Dim man_display : man_display = obj_man.displayName
    
    If ( (VarType(man_first) <> vbString Or VarType(man_last) <> vbString) And VarType(man_display) = vbString) Then
      Manager = man_display  
    Else
      Manager = man_first & " " & man_last
    End If
    
    Manager = Manager & " ( " & man_mail & " )" 
  End If
  
  wscript.echo maxPwdAge
  
  ' check UserAccountControl attributes
  Dim accountAttr : accountAttr = GetValue(user, "userAccountControl")
  
  If (accountAttr <> "<unlisted>") Then
    accountAttr = CInt(accountAttr)
    
    ' password expired?
    If (accountAttr And ADS_UF_PASSWORD_EXPIRED) Then
      Expired = "Yes"
    Else
      Expired = "No"
    End If
    
    ' will it expire?
    If (accountAttr And ADS_UF_DONT_EXPIRE_PASSWD) Then
      WillExpire = "No"
    Else
      WillExpire = "Yes"
    End If
    
    ' account locked?
    If (accountAttr And ADS_UF_LOCKOUT) Then
      Locked = "Yes"
    Else
      Locked = "No"
    End If
    
    ' account disabled?
    If (accountAttr And ADS_UF_ACCOUNTDISABLE) Then
      Disabled = "Yes"
    Else
      Disabled = "No"
    End If
  End If

  HomePage = GetValue(user, "wWWHomePage")
  
  Employee = GetValue(user, "employeeType")
  
  If (Employee = "C") Then
    Employee = "Contractor"
  ElseIf (Employee = "E") Then
    Employee = "Full Time Employee"
  End If
  
  If (Disabled = "Yes") Then
    Employee = "Previously a " & Employee
  End If
  
  printf(Array("\n"))
  printf(Array("\n  %-15s : %s",  "Full Name",    (First & " " & Last) & " (" & Employee & ")"))
  printf(Array("\n  %-15s : %s",  "Employee ID",   ID))
  printf(Array("\n  %-15s : %s",  "Logon ID",      UCase(NTID)))

  printf(Array("\n\n  %-15s : %s","E-Mail",        EMail))
  
  If (HomePage <> "<unlisted>") Then
    printf(Array("\n  %-15s : %s", "HomePage", HomePage))
  End If
  
  printf(Array("\n  %-15s : %s",  "Manager",       Manager))
  
  printf(Array("\n\n  %-15s : %s",  "SIP Account",   SipID))
  printf(Array("\n  %-15s : %s",    "Phone",         Phone))
  
  If (MSID <> "<unlisted>") Then
    printf(Array("\n  %-15s : %s",  "MS ID",         MSID))
  End If
  
  If (HomeDrive <> "<unlisted>") Then
    printf(Array("\n\n  %-15s : %s",  "Home Drive",    HomeDrive))
  End If
  
  If (HomePath <> "<unlisted>") Then
    printf(Array("\n  %-15s : %s",    "Home Directory", HomePath))
  End If
  
  If (ScriptPath <> "<unlisted>") Then
    printf(Array("\n  %-15s : %s",    "Script Path", ScriptPath))
  End If
  
  PwdLastSet = user.PasswordLastChanged
  PwdExpires = PwdLastSet + maxPwdAge
  
  printf(Array("\n\n  %-15s : %s", "Acc. Locked", Locked))
  printf(Array("\n  %-15s : %s", "Acc. Disabled", Disabled))
  
  printf(Array("\n\n  %-15s : %s", "Pass. Updated", DateValue(PwdLastSet) & " " & TimeValue(PwdLastSet) & " (" & Int((Now - PwdLastSet)) & " days ago)"))
  
  If (WillExpire = "Yes") Then
    Dim MaxPwdSecs, MaxPwdDays
    
    MaxPwdSecs = Abs(maxPwdAge.HighPart * 2^32 + maxPwdAge.LowPart) * ONE_HUNDRED_NANOSECOND
    MaxPwdDays = Int(MaxPwdSecs / SECONDS_IN_DAY)
        
    Dim days2expire 
    If (Now >= (PwdLastSet + MaxPwdDays)) Then
      days2expire = " (Already expired, will need to change at next logon)"
    Else
      Dim days : days = Int( (PwdLastSet + MaxPwdDays) - Now)
      days2expire = " (In " & days & " day"
      If (days > 1) Then
        days2expire = days2expire & "s)"
      Else
        days2expire = days2expire & ")"
      End If
    End If
    
    printf(Array("\n  %-15s : %s", "Pass. Expires", DateValue(PwdLastSet + MaxPwdDays) & " " & TimeValue(PwdLastSet) & days2expire))
  Else
    printf(Array("\n  %-15s : %s", "Pass. Expired", Expired))
  End If
  
  If (UNIXShell <> "<unlisted>") Then
    printf(Array("\n\n  %-15s : %s","UNIX Shell",    UNIXShell))
    printf(Array("\n  %-15s : %s",  "UNIX Home",     UNIXHome))
  End If
  
  printf(Array("\n\n  %-15s : %s",  "Company",       Company))
  printf(Array("\n  %-15s : %s",    "Department",    Department))
  printf(Array("\n  %-15s : %s",    "Post Code",     PostCode))
  
  printf(Array("\n\n  %-15s : %s",  "Created",       Created))
  printf(Array("\n  %-15s : %s\n",  "Description",   Description))
  
  '
  ' reports requested?
  '
  If (bReports) Then
    ProcessReports(user)
  End If
  
  '
  ' 
  '
  If (bGroups) Then
    Call ProcessGroups(user, First)
  End If
  
  '
  '
  '
  If (bExchange) Then
    Call ProcessExchange(user, First)
  End If
  
End Function

' Frost, Sam
'
' publicDelegates - This attribute stores the user that was configured as a Delegate. 
' (Who is a Delegate of my mailbox) 
'
' publicDelegatesBL – This attribute stores which mailbox this user is a Delegate of. 
' (What mailbox am I a Delegate of) 

Function ProcessExchange(ByVal user, ByVal First)
  On Error Resume Next
  
  Dim secretaries, managers, server, sec_count, man_count
  
  '
  '
  '
  Err.Clear
  secretaries = user.Get("publicDelegates")
  
  If (Err.Number <> 0 Or VarType(secretaries) = vbEmpty) Then
    sec_count = 0
  Else
    sec_count = UBound(secretaries) + 1
  End If
  
  '
  '
  '
  Err.Clear
  managers = user.Get("publicDelegatesBL")
  
  If (Err.Number <> 0 Or VarType(managers) = vbEmpty) Then
    man_count = 0
  Else
    man_count = UBound(managers) + 1
  End If
  
  'printf(Array("\n  [+] User delegates for %i mail boxes and this mailbox has %i delegates\n", man_count, sec_count))
  
  If (sec_count > 0) Then
    printf(Array("\n  [+] %s has delegate access to the following mailboxes.\n", First))
    DumpUsers(secretaries)
  End If 
  
  printf(Array("\n"))
  
  If (man_count > 0) Then
    printf(Array("\n  [+] Users with delegate access to %s's mailbox.\n", First))
    DumpUsers(managers)
  End If
  
  On Error GoTo 0
End Function

Function DumpUsers(ByVal dn_list)
  Dim user, aUsers, MaxMail, MaxName, dn
  ReDim aUsers(0)
  
  For Each dn In dn_list
    
    Err.Clear
    Dim user_obj : Set user_obj = GetObject("LDAP://" & dn)
    
    If (Err.Number = 0) Then
      Dim user_info : Set user_info = New SReport
      
      user_info.FullName   = GetValue(user_obj, "givenName") & " " & GetValue(user_obj, "sn")
      user_info.EmployeeID = GetValue(user_obj, "employeeID")
      user_info.NTID       = GetValue(user_obj, "sAMAccountName")
      user_info.Email      = GetValue(user_obj, "mail")
      user_info.Phone      = GetValue(user_obj, "telephoneNumber")
      user_info.PostCode   = GetValue(user_obj, "postalCode")
      
      If (Len(user_info.FullName) > MaxName) Then MaxName = Len(user_info.FullName)
      If (Len(user_info.Email)    > MaxMail) Then MaxMail = Len(user_info.Email)
        
      ' add to main list
      Set aUsers(UBound(aUsers)) = user_info
      ReDim Preserve aUsers(UBound(aUsers) + 1)         
    
      Set user_obj = Nothing
      Set user_info = Nothing
    End If
  Next
 
  Dim i
  
  printf(Array("\n  %-*s\t%-15s\t%-10s\t%20s\t%*s\t%s", MaxName, "Full Name", "Employee ID", "NTID", "Phone", MaxMail, "E-Mail", "PostCode"))
  printf(Array("\n  %s\t%-15s\t%-10s\t%20s\t%s\t%s\n", String(MaxName, "*"), String(15,"*"), String(10,"*"), String(20,"*"), String(MaxMail,"*"), String(10,"*")))
  
  For i = 0 To UBound(aUsers) - 1
    printf(Array("  %-*s\t%-15s\t%-10s\t%20s\t%*s\t%s\n", MaxName, aUsers(i).FullName, aUsers(i).EmployeeID, aUsers(i).NTID, aUsers(i).Phone, MaxMail, aUsers(i).EMail, aUsers(i).PostCode))
  Next
  
End Function

Class SReport
  Public FullName
  Public EmployeeID
  Public NTID
  Public Phone
  Public Email
  Public PostCode
End Class

'
' display people that report to this user
'
Function ProcessReports(ByVal user)
  On Error Resume Next
  
  Dim aReports, MaxName, MaxMail
  ReDim aReports(0)
  
  MaxName = 25
  MaxMail = 25
  
  Err.Clear
  Dim report_count
  Dim report_list : report_list = user.Get("directReports")
  
  
  If (Err.Number <> 0) Then
    report_count = 0
  Else
    report_count = UBound(report_list) + 1
  End If
    
  printf(Array("\n  [+] Number of employees reporting to this person: %i\n", report_count))
    
  If (report_count <> 0) Then
    Dim report_dn
    
    For Each report_dn In report_list
      ' get info for this report from Active Directory
      Err.Clear
      Dim report_obj  : Set report_obj = GetObject("LDAP://" & report_dn)
        
      If (Err.Number = 0) Then
        
        Dim report_info : Set report_info = New SReport
        
        report_info.FullName   = GetValue(report_obj, "givenName") & " " & GetValue(report_obj, "sn")
        report_info.EmployeeID = GetValue(report_obj, "employeeID")
        report_info.NTID       = GetValue(report_obj, "sAMAccountName")
        report_info.Email      = GetValue(report_obj, "mail")
        report_info.Phone      = GetValue(report_obj, "telephoneNumber")
        report_info.PostCode   = GetValue(report_obj, "postalCode")
        
        If (Len(report_info.FullName) > MaxName) Then MaxName = Len(report_info.FullName)
        If (Len(report_info.Email)    > MaxMAil) Then MaxMail = Len(report_info.Email)
          
        ' add to main list
        Set aReports(UBound(aReports)) = report_info
        ReDim Preserve aReports(UBound(aReports) + 1)         
      
        Set report_obj = Nothing
        Set report_info = Nothing
      End If    
    Next
    
    Dim i
    
    printf(Array("\n  %-*s\t%-15s\t%-10s\t%20s\t%*s\t%s", MaxName, "Full Name", "Employee ID", "NTID", "Phone", MaxMail, "E-Mail", "PostCode"))
    printf(Array("\n  %-*s\t%-15s\t%-10s\t%20s\t%*s\t%s\n", MaxName, String(MaxName, "*"), String(15,"*"), String(10,"*"), String(20,"*"), MaxMail, String(25,"*"), String(10,"*")))
    
    For i = 0 To UBound(aReports) - 1
      printf(Array("  %-*s\t%-15s\t%-10s\t%20s\t%*s\t%s\n", MaxName, aReports(i).FullName, aReports(i).EmployeeID, aReports(i).NTID, aReports(i).Phone, MaxMail, aReports(i).EMail, aReports(i).PostCode))
    Next
  End If
  
  On Error GoTo 0
End Function

'
'
'
Function ProcessGroups(ByVal user)
  On Error Resume Next
  
  Dim aGroups, MaxNTID, MaxDesc
  ReDim aGroups(0)
    
  MaxNTID = 10
  MaxDesc = 10
  
  Err.Clear
  
  Dim group_count
  Dim group_list : group_list = user.Get("memberOf")
    
  If (Err.Number <> 0) Then
    group_count = 0
  Else
    group_count = UBound(group_list) + 1
  End If
  
  printf(Array("\n  [+] Number of Global Groups this person is part of: %i\n", group_count))
    
  If (group_count <> 0) Then
    Dim group_dn
    For Each group_dn In group_list
      
      ' get info for this group from Active Directory
      Dim group_obj  : Set group_obj = GetObject("LDAP://" & group_dn)
        
      If (Err.Number = 0) Then
        
        ' create new group object
        Dim group_info : Set group_info = New CGroup
      
        ' check the lengths
        Dim ntid_str : ntid_str = group_obj.sAMAccountName
        Dim desc_str : desc_str = group_obj.description
        
        If (Len(ntid_str) > MaxNTID) Then MaxNTID = Len(ntid_str)
        If (Len(desc_str) > MaxDesc) Then MaxDesc = Len(desc_str)
        
        ' assign values to group object
        group_info.DN          = group_obj.distinguishedName
        group_info.NTID        = group_obj.sAMAccountName
        group_info.Description = group_obj.description
        group_info.Info        = group_obj.info
        
        ' add to main list
        Set aGroups(UBound(aGroups)) = group_info
        ReDim Preserve aGroups(UBound(aGroups) + 1)
      
        Set group_obj = Nothing
        'Set group_info = Nothing
      Else

      End If
    Next
      
    Call DumpGroups(aGroups, MaxNTID, MaxDesc)
  End If 
  
  On Error GoTo 0
End Function

  '
  ' return value of property or "<unlisted>"
  '
Function GetValue(ByVal user, ByVal propname)
  On Error Resume Next
  
  GetValue = CStr(user.Get(propname))
  
  If (VarType(GetValue) <> vbString) Then
    GetValue = "<unlisted>"
  End If

  On Error GoTo 0
End Function
 

'
' dump group information for our user
'
Sub DumpGroups(ByVal aGroups, ByVal vNTID, ByVal vDesc)
    
    Dim group_format
    Call sprintf(group_format, Array("  %%-%is\t%%-%is", vNTID, vDesc))
    group_format = group_format & "\t%-10s\n"
    
    printf(Array("\n"))
    printf(Array(group_format, "Name", "Description", "Info"))
    printf(Array(group_format, String(vNTID,"*"), String(vDesc,"*"), String(10,"*")))
  
    Dim group, idx
    For idx = 0 To UBound(aGroups) - 1 
      Set group = aGroups(idx)
      printf(Array(group_format, group.NTID, group.Description, group.Info))
    Next
End Sub
 
' *************************************************************
' class for group
' *************************************************************
Class CGroup
  Public DN             ' distinguishedName
  Public Creation       ' whenCreated
  
  Private m_NTID        ' sAMAccountName
  Private m_Description ' description
  Private m_Info        ' info

  Private Sub Class_Initialize
    
  End Sub
  
  Public Sub SetInfo(ByRef rs)   
    Dim desc : desc = GetValue(rs, "description")
    
    If (VarType(desc) <> vbString) Then
      Description = desc(0)
    End If
    
    NTID        = GetValue(rs, "sAMAccountName")
    Info        = GetValue(rs, "info")
    Creation    = GetValue(rs, "whenCreated")
  End Sub
  
  ' 
  ' set description
  '
  Public Property Let NTID(ByVal sNTID)
    If (Len(sNTID) > 0) Then
      m_NTID = sNTID
      
      ' remove line feeds / carriage returns
      m_NTID = Replace(Trim(m_NTID), vbCrLf, " ")
      
      ' remove tabs
      m_NTID = Replace(Trim(m_NTID), vbTab,  " ")
    End If
  End Property
  
  ' 
  ' get NTID
  '
  Public Property Get NTID()
    If (Len(m_NTID) = 0) Then
      NTID = "<unlisted>"
    Else
      NTID = m_NTID
    End If
  End Property
  
  ' 
  ' set description
  '
  Public Property Let Description(ByVal sDesc)
    If (Len(sDesc) > 0) Then
      m_Description = sDesc
      ' remove line feeds / carriage returns
      m_Description = Replace(Trim(m_Description), vbCrLf, " ")
      ' remove tabs
      m_Description = Replace(Trim(m_Description), vbTab,  " ")
    End If
  End Property
  
  ' 
  ' get description
  '
  Public Property Get Description()
    If (Len(m_Description) = 0) Then
      Description = "<unlisted>"
    Else
      Description = m_Description
    End If
  End Property
  
  '
  ' set info
  '
  Public Property Let Info(ByVal sInfo)
    If (Len(sInfo) > 0) Then
      m_Info = sInfo
      
      ' remove line feeds / carriage returns
      m_Info = Replace(Trim(m_Info), vbCrLf, " ")
      
      ' remove tabs
      m_Info = Replace(Trim(m_Info), vbTab,  " ")
    End If
  End Property
  
  '
  ' get info
  '
  Public Property Get Info()
    If (Len(m_Info) = 0) Then
      Info = "<unlisted>"
    Else
      Info = m_Info
    End If
  End Property
  
  Private Function GetValue(ByRef rs, ByVal fieldName)    
    If (IsNull(rs.Fields(fieldName))) Then
      GetValue = "<unlisted>"
    Else
      GetValue = rs.Fields(fieldName).Value
    End If  
  End Function
  
End Class

  '=======================================================================================
  ' The last logon time returned by AD is expressed as the number of 100 nanosecond intervals
  ' since 12:00 AM January 1, 1601.  This algorithm - pinched off the Internet from R. Mueller
  ' - converts this to a conventional time and date.
  '=======================================================================================
  Function LastLogonDate( byVal LastLogon, byVal TimeOffset )
    Const TwoToThePower32  = 4294967296
    Const PeriodsPerMinute = 600000000
    Const StartDate        = #1/1/1601#
    Const MinutesPerDay    = 1440
    Dim objDate
    Dim lngHigh
    Dim lngLow
    If IsNull( LastLogon ) Then
      LastLogonDate = #1/1/1601#
    Else
      ' Vbscript cannot do arithmetic with 64-bit integers, so we convert it to an object then
      ' use the in-built HightPart and LowPart methods to split it into two 32-bit components.
      Set objDate = LastLogon
      lngHigh = objDate.HighPart
      lngLow = objDate.LowPart
      ' The 32nd bit becomes the MSB of the lower half.
      ' If this bit is 1, then the lower half is interpreted as a negative number
      ' because vbscript stores integers using two's complement - so it is intepreted as -2^32
      ' instead of +2^32  - so the overall value is 2 x 2^32 = 2^64 too low.  So add one to the
      ' second 32-bit block to compensate for this.
      If lngLow < 0 Then
        lngHigh = lngHigh + 1
      End If
      If (lngHigh = 0) And (lngLow = 0 ) Then
        LastLogonDate = #1/1/1601#
      Else
        LastLogonDate = StartDate + (((lngHigh * TwoToThePower32) + lngLow)/PeriodsPerMinute - (TimeOffset * 60) ) / MinutesPerDay
      End If
    End If
  End Function

  '
  ' bind to LDAP server and return defaultNamingContext / distinguishedName
  '
  Function GetDomainDN(ByVal vDomain)
    Dim ldap
    Set ldap = GetObject("LDAP://" & vDomain & "/RootDSE")
    
    If (VarType(ldap) <> vbObject) Then
      Exit Function
    End If
    
    GetDomainDN = ldap.Get("defaultNamingContext")
    Set ldap = Nothing
    
    Set ldap = GetObject("LDAP://" & GetDomainDN)
    Set maxPwdAge = ldap.Get("maxPwdAge")
    
    Set ldap = Nothing
  End Function
  
  '
  ' search for user object using first and last name as criteria
  '
  Function SearchByNTIDOrEmployeeID(ByVal vDomain, ByVal vParam)
    Dim domainDN
    
    domainDN = GetDomainDN(vDomain)
    
    If (VarType(domainDN) = vbString) Then      
      Dim strBase, strAttributes, strFilter
      strBase = "<LDAP://" & domainDN & ">"
      strAttributes = "distinguishedName"
      
      If (IsNumeric(vParam)) Then
        strFilter = "(&(objectCategory=User)(employeeID=" &  vParam & "))"
      Else
        strFilter = "(&(objectCategory=User)(sAMAccountName=" &  vParam & "))"
      End If
      
      Dim connection, command
      
      Set connection = CreateObject("ADODB.Connection") 
      Set command    = CreateObject("ADODB.Command")
      
      connection.Provider = "ADsDSOOBject" 
      connection.Open "Active Directory Provider"
      
      Set command.ActiveConnection = connection
      
      command.Properties("Searchscope")     = ADS_SCOPE_SUBTREE
      command.Properties("Chase referrals") = ADS_CHASE_REFERRALS_ALWAYS
      
      command.CommandText = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
      
      Dim rs, list
      Set rs = command.Execute()
      
      If (Not rs.EOF) Then
        ReDim list(0)
        
        Do Until rs.EOF
          list(UBound(list)) = "LDAP://" & rs.Fields("distinguishedName")
          rs.MoveNext
          ReDim Preserve list(UBound(list) + 1)
        Loop
        SearchByNTIDOrEmployeeID = list
      End If
      
      Set connection = Nothing
      Set command = Nothing
    End If   
  End Function  
  
  '
  ' search for user object using first and last name as criteria
  '
  Function SearchByFirstAndLastName(ByVal vDomain, ByVal vFirst, ByVal vLast)
    Dim domainDN
    
    domainDN = GetDomainDN(vDomain)
    
    If (VarType(domainDN) = vbString) Then      
      Dim strBase, strAttributes, strFilter
      strBase = "<LDAP://" & domainDN & ">"
      strAttributes = "distinguishedName"
      
      strFilter = "(&(objectCategory=User)(sn=" &  vLast & ")(givenName=" & vFirst & "))"
      
      Dim connection, command
      
      Set connection = CreateObject("ADODB.Connection") 
      Set command    = CreateObject("ADODB.Command")
      
      connection.Provider = "ADsDSOOBject" 
      connection.Open "Active Directory Provider"
      
      Set command.ActiveConnection = connection
      
      command.Properties("Searchscope")     = ADS_SCOPE_SUBTREE
      command.Properties("Chase referrals") = ADS_CHASE_REFERRALS_ALWAYS
      command.Properties("Sort On")         = "givenName"
      
      command.CommandText = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
      
      Dim rs, list
      Set rs = command.Execute()
      
      If (Not rs.EOF) Then
        ReDim list(0)
        
        Do Until rs.EOF
          list(UBound(list)) = "LDAP://" & rs.Fields("distinguishedName")
          rs.MoveNext
          ReDim Preserve list(UBound(list) + 1)
        Loop
        SearchByFirstAndLastName = list
      End If
      
      Set connection = Nothing
      Set command = Nothing
    End If   
  End Function  
  
  '
  ' search for user object using first and last name as criteria
  '
  Function SearchByDisplayName(ByVal vDomain, ByVal vFirst, ByVal vLast)
    Dim domainDN
    
    domainDN = GetDomainDN(vDomain)
    
    If (VarType(domainDN) = vbString) Then      
      Dim strBase, strAttributes, strFilter
      strBase = "<LDAP://" & domainDN & ">"
      strAttributes = "distinguishedName"
      
      strFilter = "(&(objectCategory=User)(displayName=" &  vLast & ", " & vFirst & "))"
      
      Dim connection, command
      
      Set connection = CreateObject("ADODB.Connection") 
      Set command    = CreateObject("ADODB.Command")
      
      connection.Provider = "ADsDSOOBject" 
      connection.Open "Active Directory Provider"
      
      Set command.ActiveConnection = connection
      
      command.Properties("Searchscope")     = ADS_SCOPE_SUBTREE
      command.Properties("Chase referrals") = ADS_CHASE_REFERRALS_ALWAYS
      
      command.CommandText = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
      
      Dim rs, list
      Set rs = command.Execute()
      
      If (Not rs.EOF) Then
        ReDim list(0)
        
        Do Until rs.EOF
          list(UBound(list)) = "LDAP://" & rs.Fields("distinguishedName")
          rs.MoveNext
          ReDim Preserve list(UBound(list) + 1)
        Loop
        SearchByDisplayName = list
      End If
      
      Set connection = Nothing
      Set command = Nothing
    End If   
  End Function  
  
  '
  ' bind to LDAP server and obtain a distinguishedName for user, group or computer
  '
  Function GetObjectDN(ByVal vDomain, ByVal vType, ByVal vValue)
    Dim domainDN
    
    domainDN = GetDomainDN(vDomain)
    
    If (VarType(domainDN) = vbString) Then      
      Dim strBase, strAttributes, strFilter
      strBase = "<LDAP://" & domainDN & ">"
      strAttributes = "distinguishedName"
      
      ' search for group, user, machine individually or toegether
      Select Case vType
        Case SAM_GROUP_OBJECT
          strFilter = "(&(objectCategory=Group)(cn=" &  vValue & "))"
        Case SAM_USER_OBJECT
          strFilter = "(&(objectCategory=User)(cn=" &  vValue & "))"
        Case SAM_MACHINE_OBJECT
          strFilter = "(&(objectCategory=Computer)(cn=" &  vValue & "))"
        Case Else
          strFilter = "(&(|(ObjectCategory=User)(ObjectCategory=Group)(ObjectCategory=Computer))(cn=" & vValue & "))"
      End Select
      
      Dim connection, command
      
      Set connection = CreateObject("ADODB.Connection") 
      Set command    = CreateObject("ADODB.Command")
      
      connection.Provider = "ADsDSOOBject" 
      connection.Open "Active Directory Provider"
      
      Set command.ActiveConnection = connection
      
      command.Properties("Searchscope")     = ADS_SCOPE_SUBTREE
      command.Properties("Chase referrals") = ADS_CHASE_REFERRALS_ALWAYS
      
      command.CommandText = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
      
      Dim rs
      Set rs = command.Execute()
      
      If (Not rs.EOF) Then
        Do Until rs.EOF 
          GetObjectDN = "LDAP://" & rs.Fields("distinguishedName")
          rs.MoveNext
        Loop
      End If
      
      Set rs = Nothing
      Set connection = Nothing
      Set command = Nothing
    End If
  End Function

  '
  ' debug print
  '
  Function dprintf(ByVal args)
    If (DEBUG_MODE) Then
      Call printf(args)
    End If
  End Function
  
  '
  ' regular printf
  '
  Function printf(ByVal args)
    Dim s
    Call sprintf(s, args)
    WScript.StdOut.Write(s)
  End Function
   
  '
  ' write printf strings to file
  '
  Function fprintf(ByVal fd, ByRef args)
    Dim s
    Call sprintf(s, args)
    
    fd.Write(s)
  End Function

  ' ##################################################################################
  '
  ' minimal printf emulation For VBScript
  ' WARNING: Not a full implementation or bug free..misuse and you'll easily break it.
  '
  ' supported specifiers
  '
  ' d or i   = decimal or integer
  ' s        = string
  ' x        = lowercase hex byte
  ' X        = uppercase hex byte
  ' o        = octal byte
  ' %        = percentage sign
  '
  ' supported flags
  '
  ' -        = left justify
  ' #        = precede with 0, 0x or 0X if radix
  ' 0        = pad with zeros
  '
  ' supported width
  '
  ' *        = specified In argument preceding value
  ' (number) = number of spaces or zeros to pad
  '
  ' supported escape sequences
  '
  ' n        = line feed
  ' t        = tab
  ' r        = carriage return
  ' \        = backslash
  '
  ' Copyright (c) 2012 - Kevin Devine
  '
  ' #################################################################################
  Function sprintf(ByRef s1, ByRef args)
    
    Const RIGHT_JUSTIFY = 0
    Const LEFT_JUSTIFY  = 1
    
    ' array and not empty?
    If (IsArray(args) And UBound(args) <> -1) Then
      Dim s2

      s1 = args(0)
       
      'Call WScript.StdOut.WriteLine(s1)
      
      ' ensure atleast 1 parameter before parsing
      If (UBound(args) >= 0) Then
        Dim fmt, i, index, pad_len
        fmt = s1
        index = 1

        For i = 1 To Len(fmt)

          Dim justify, zero_pad, lower_case, arg, show_radix

          justify    = RIGHT_JUSTIFY    ' default
          zero_pad   = False
          lower_case = False
          show_radix = False
          pad_len    = 0
          arg        = ""
           
          ' ############################################# SPECIFIER?
          ' Check if beginning of specifier
          '
          Dim c : c = Mid(fmt, i, 1)
           
          If (c = "%") Then
          
            ' skip percentage sign
            i = i + 1
               
            ' ############################################# step 1
            ' Check For supported flags
            '
            c = Mid(fmt, i, 1)
               
            ' show radix?
            If (c = "#") Then
              show_radix = True
              i = i + 1
                 
            ' left justify?
            ElseIf (c = "-") Then
              justify = LEFT_JUSTIFY
              i = i + 1
               
            ' pad with zeros instead of spaces?             
            ElseIf (c = "0") Then
              zero_pad = True
              i = i + 1
              
            ' space?  
            ElseIf (c = " ") Then
              arg = arg & " "
            End If
              
            ' ############################################# step 2
            ' Check For padding width
            '
            c = Mid(fmt, i, 1)
            
            ' specified In arguments?             
            If (c = "*") Then
              pad_len = args(index)
              index = index + 1
              i = i + 1
            
            ' if it's a number presume it to be specified length
            ElseIf (IsNumeric(c)) Then
              Do While True
                c = Mid(fmt, i, 1)

                ' convert string to binary
                If (IsNumeric(c)) Then
                  pad_len = pad_len * 10
                  pad_len = pad_len + (c - CInt("0"))
                  i = i + 1
                Else
                  Exit Do
                End If
              Loop   
            End If
               
            ' ############################################# step 3
            ' What specifier?
            '
            c = Mid(fmt, i, 1)

            Select Case c
              Case "s"
                arg = arg & args(index)
              Case "d"
                arg = arg & CStr(args(index))
              Case "i"
                arg = arg & CStr(args(index))
              Case "x"
                If (show_radix) Then
                  arg = arg & "0x"
                End If
                arg = arg & LCase(Hex(args(index)))
              Case "X"
                If (show_radix) Then
                  arg = "0X"
                End If
                  arg = arg & UCase(Hex(args(index)))
              Case "o"
                If (show_radix) Then
                  arg = "0"
                End If
                  arg = arg & Oct(args(index))
              Case "%"
                arg = "%"
              Case Else
                WScript.Echo "Unrecognized specifier: " & c
            End Select
            
            ' ############################################# step 4
            ' process padding length and justifcation
            '
            If (pad_len <> 0) Then
              If (pad_len > Len(arg)) Then
                pad_len = pad_len - Len(arg)
              Else
                pad_len = 0
              End If
            End If

            ' pad with zeros or spaces?
            Dim pad_str
            
            If (zero_pad) Then
              pad_str = String(pad_len, "0")
            Else
              pad_str = Space(pad_len)
            End If

            ' justify left or right?
            If (justify = LEFT_JUSTIFY) Then
              s2 = s2 & arg & pad_str
            Else
              s2 = s2 & pad_str & arg
            End If

            ' display percent?
            If (arg <> "%") Then
              index = index + 1
            End If
            
          ' ############################################# ESCAPE SEQUENCE?
          ' 
          '
          ElseIf (c = "\") Then
             
            ' get the byte
            Dim esc : esc = Mid(fmt, i + 1, 1)
             
            ' newline?
            If (esc = "n") Then
              s2 = s2 & vbCrLf
              i = i + 1
            
            ' carriage return?           
            ElseIf (esc = "r") Then
              s2 = s2 & vbCr
              i = i + 1
            
            ' horizontal tab?
            ElseIf (esc = "t") Then
              s2 = s2 & vbTab
              i = i + 1
            
            ' backslash?
            ElseIf (esc = "\") Then
              s2 = s2 & "\"
              i = i + 1
            
            Else
              ' unrecognized ..
              s2 = s2 & "\"
            End If
          ' ############################################# IGNORE
          '
          Else         
            s2 = s2 & c
          End If
        Next
        s1 = s2
      End If
    Else
      WScript.Echo "No Array provided"
    End If
  End Function
  
