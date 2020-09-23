Attribute VB_Name = "Module1"
' Constants used within our API calls. Refer to the MSDN for more
' information on how/what these constants are used for.

' Memory constants used through various memory API calls.
Public Const GMEM_MOVEABLE = &H2
Public Const LMEM_FIXED = &H0
Public Const LMEM_ZEROINIT = &H40
Public Const LPTR = (LMEM_FIXED + LMEM_ZEROINIT)
Public Const GENERIC_READ = &H80000000
Public Const GENERIC_ALL = &H10000000
Public Const GENERIC_EXECUTE = &H20000000
Public Const GENERIC_WRITE = &H40000000

' The file/security API call constants.
' Refer to the MSDN for more information on how/what these constants
' are used for.
Public Const DACL_SECURITY_INFORMATION = &H4
Public Const SECURITY_DESCRIPTOR_REVISION = 1
Public Const SECURITY_DESCRIPTOR_MIN_LENGTH = 20
Public Const SD_SIZE = (65536 + SECURITY_DESCRIPTOR_MIN_LENGTH)
Public Const ACL_REVISION2 = 2
Public Const ACL_REVISION = 2
Public Const MAXDWORD = &HFFFFFFFF
Public Const SidTypeUser = 1
Public Const AclSizeInformation = 2

'  The following are the inherit flags that go into the AceFlags field
'  of an Ace header.

Public Const OBJECT_INHERIT_ACE = &H1
Public Const CONTAINER_INHERIT_ACE = &H2
Public Const NO_PROPAGATE_INHERIT_ACE = &H4
Public Const INHERIT_ONLY_ACE = &H8
Public Const INHERITED_ACE = &H10
Public Const VALID_INHERIT_FLAGS = &H1F
Public Const DELETE = &H10000

' Structures used by our API calls.
' Refer to the MSDN for more information on how/what these
' structures are used for.
Type ACE_HEADER
   AceType As Byte
   AceFlags As Byte
   AceSize As Integer
End Type


Public Type ACCESS_DENIED_ACE
  Header As ACE_HEADER
  Mask As Long
  SidStart As Long
End Type

Type ACCESS_ALLOWED_ACE
   Header As ACE_HEADER
   Mask As Long
   SidStart As Long
End Type

Type ACL
   AclRevision As Byte
   Sbz1 As Byte
   AclSize As Integer
   AceCount As Integer
   Sbz2 As Integer
End Type

Type ACL_SIZE_INFORMATION
   AceCount As Long
   AclBytesInUse As Long
   AclBytesFree As Long
End Type

Type SECURITY_DESCRIPTOR
   Revision As Byte
   Sbz1 As Byte
   Control As Long
   Owner As Long
   Group As Long
   sACL As ACL
   Dacl As ACL
End Type

' API calls used within this sample. Refer to the MSDN for more
' information on how/what these APIs do.

Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function LookupAccountName Lib "advapi32.dll" Alias "LookupAccountNameA" (lpSystemName As String, ByVal lpAccountName As String, sid As Any, cbSid As Long, ByVal ReferencedDomainName As String, cbReferencedDomainName As Long, peUse As Long) As Long
Declare Function InitializeSecurityDescriptor Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal dwRevision As Long) As Long
Declare Function GetSecurityDescriptorDacl Lib "advapi32.dll" (pSecurityDescriptor As Byte, lpbDaclPresent As Long, pDacl As Long, lpbDaclDefaulted As Long) As Long
Declare Function GetFileSecurityN Lib "advapi32.dll" Alias "GetFileSecurityA" (ByVal lpFileName As String, ByVal RequestedInformation As Long, ByVal pSecurityDescriptor As Long, ByVal nLength As Long, lpnLengthNeeded As Long) As Long
Declare Function GetFileSecurity Lib "advapi32.dll" Alias "GetFileSecurityA" (ByVal lpFileName As String, ByVal RequestedInformation As Long, pSecurityDescriptor As Byte, ByVal nLength As Long, lpnLengthNeeded As Long) As Long
Declare Function GetAclInformation Lib "advapi32.dll" (ByVal pAcl As Long, pAclInformation As Any, ByVal nAclInformationLength As Long, ByVal dwAclInformationClass As Long) As Long
Public Declare Function EqualSid Lib "advapi32.dll" (pSid1 As Byte, ByVal pSid2 As Long) As Long
Declare Function GetLengthSid Lib "advapi32.dll" (pSid As Any) As Long
Declare Function InitializeAcl Lib "advapi32.dll" (pAcl As Byte, ByVal nAclLength As Long, ByVal dwAclRevision As Long) As Long
Declare Function GetAce Lib "advapi32.dll" (ByVal pAcl As Long, ByVal dwAceIndex As Long, pace As Any) As Long
Declare Function AddAce Lib "advapi32.dll" (ByVal pAcl As Long, ByVal dwAceRevision As Long, ByVal dwStartingAceIndex As Long, ByVal pAceList As Long, ByVal nAceListLength As Long) As Long
Declare Function AddAccessAllowedAce Lib "advapi32.dll" (pAcl As Byte, ByVal dwAceRevision As Long, ByVal AccessMask As Long, pSid As Byte) As Long
Public Declare Function AddAccessDeniedAce Lib "advapi32.dll" (pAcl As Byte, ByVal dwAceRevision As Long, ByVal AccessMask As Long, pSid As Byte) As Long
Declare Function SetSecurityDescriptorDacl Lib "advapi32.dll" (pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal bDaclPresent As Long, pDacl As Byte, ByVal bDaclDefaulted As Long) As Long
Declare Function SetFileSecurity Lib "advapi32.dll" Alias "SetFileSecurityA" (ByVal lpFileName As String, ByVal SecurityInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Public Sub SetAccess(sUserName As String, sFileName As String, lMask As Long)
   Dim lResult As Long            ' Result of various API calls.
   Dim I As Integer               ' Used in looping.
   Dim bUserSid(255) As Byte      ' This will contain your SID.
   Dim bTempSid(255) As Byte      ' This will contain the Sid of each ACE in the ACL .
   Dim sSystemName As String      ' Name of this computer system.

   Dim lSystemNameLength As Long  ' Length of string that contains
                                  ' the name of this system.

   Dim lLengthUserName As Long    ' Max length of user name.

   'Dim sUserName As String * 255  ' String to hold the current user
                                  ' name.


   Dim lUserSID As Long           ' Used to hold the SID of the
                                  ' current user.

   Dim lTempSid As Long            ' Used to hold the SID of each ACE in the ACL
   Dim lUserSIDSize As Long          ' Size of the SID.
   Dim sDomainName As String * 255   ' Domain the user belongs to.
   Dim lDomainNameLength As Long     ' Length of domain name needed.

   Dim lSIDType As Long              ' The type of SID info we are
                                     ' getting back.

   Dim sFileSD As SECURITY_DESCRIPTOR   ' SD of the file we want.

   Dim bSDBuf() As Byte           ' Buffer that holds the security
                                  ' descriptor for this file.

   Dim lFileSDSize As Long           ' Size of the File SD.
   Dim lSizeNeeded As Long           ' Size needed for SD for file.


   Dim sNewSD As SECURITY_DESCRIPTOR ' New security descriptor.

   Dim sACL As ACL                   ' Used in grabbing the DACL from
                                     ' the File SD.

   Dim lDaclPresent As Long          ' Used in grabbing the DACL from
                                     ' the File SD.

   Dim lDaclDefaulted As Long        ' Used in grabbing the DACL from
                                     ' the File SD.

   Dim sACLInfo As ACL_SIZE_INFORMATION  ' Used in grabbing the ACL
                                         ' from the File SD.

   Dim lACLSize As Long           ' Size of the ACL structure used
                                  ' to get the ACL from the File SD.

   Dim pAcl As Long               ' Current ACL for this file.
   Dim lNewACLSize As Long        ' Size of new ACL to create.
   Dim bNewACL() As Byte          ' Buffer to hold new ACL.

   Dim sCurrentACE As ACCESS_ALLOWED_ACE    ' Current ACE.
   Dim pCurrentAce As Long                  ' Our current ACE.

   Dim nRecordNumber As Long

   ' Get the SID of the user. (Refer to the MSDN for more information on SIDs
   ' and their function/purpose in the operating system.) Get the SID of this
   ' user by using the LookupAccountName API. In order to use the SID
   ' of the current user account, call the LookupAccountName API
   ' twice. The first time is to get the required sizes of the SID
   ' and the DomainName string. The second call is to actually get
   ' the desired information.

   lResult = LookupAccountName(vbNullString, sUserName, _
      bUserSid(0), 255, sDomainName, lDomainNameLength, _
      lSIDType)

   ' Now set the sDomainName string buffer to its proper size before
   ' calling the API again.
   sDomainName = Space(lDomainNameLength)

   ' Call the LookupAccountName again to get the actual SID for user.
   lResult = LookupAccountName(vbNullString, sUserName, _
      bUserSid(0), 255, sDomainName, lDomainNameLength, _
      lSIDType)

   ' Return value of zero means the call to LookupAccountName failed;
   ' test for this before you continue.
     If (lResult = 0) Then
        MsgBox "Error: Unable to Lookup the Current User Account: " _
           & sUserName
        Exit Sub
     End If

   ' You now have the SID for the user who is logged on.
   ' The SID is of interest since it will get the security descriptor
   ' for the file that the user is interested in.
   ' The GetFileSecurity API will retrieve the Security Descriptor
   ' for the file. However, you must call this API twice: once to get
   ' the proper size for the Security Descriptor and once to get the
   ' actual Security Descriptor information.

   lResult = GetFileSecurityN(sFileName, DACL_SECURITY_INFORMATION, _
      0, 0, lSizeNeeded)

   ' Redimension the Security Descriptor buffer to the proper size.
   ReDim bSDBuf(lSizeNeeded)

   ' Now get the actual Security Descriptor for the file.
   lResult = GetFileSecurity(sFileName, DACL_SECURITY_INFORMATION, _
      bSDBuf(0), lSizeNeeded, lSizeNeeded)

   ' A return code of zero means the call failed; test for this
   ' before continuing.
   If (lResult = 0) Then
      MsgBox "Error: Unable to Get the File Security Descriptor"
      Exit Sub
   End If

   ' Call InitializeSecurityDescriptor to build a new SD for the
   ' file.
   lResult = InitializeSecurityDescriptor(sNewSD, _
      SECURITY_DESCRIPTOR_REVISION)

   ' A return code of zero means the call failed; test for this
   ' before continuing.
   If (lResult = 0) Then
      MsgBox "Error: Unable to Initialize New Security Descriptor"
      Exit Sub
   End If

   ' You now have the file's SD and a new Security Descriptor
   ' that will replace the current one. Next, pull the DACL from
   ' the SD. To do so, call the GetSecurityDescriptorDacl API
   ' function.

   lResult = GetSecurityDescriptorDacl(bSDBuf(0), lDaclPresent, _
      pAcl, lDaclDefaulted)

   ' A return code of zero means the call failed; test for this
   ' before continuing.
   If (lResult = 0) Then
      MsgBox "Error: Unable to Get DACL from File Security " _
         & "Descriptor"
      Exit Sub
   End If

   ' You have the file's SD, and want to now pull the ACL from the
   ' SD. To do so, call the GetACLInformation API function.
   ' See if ACL exists for this file before getting the ACL
   ' information.
   If (lDaclPresent = False) Then
      MsgBox "Error: No ACL Information Available for this File"
      Exit Sub
   End If

   ' Attempt to get the ACL from the file's Security Descriptor.
   lResult = GetAclInformation(pAcl, sACLInfo, Len(sACLInfo), 2&)

   ' A return code of zero means the call failed; test for this
   ' before continuing.
   If (lResult = 0) Then
      MsgBox "Error: Unable to Get ACL from File Security Descriptor"
      Exit Sub
   End If

   ' Now that you have the ACL information, compute the new ACL size
   ' requirements.
   lNewACLSize = sACLInfo.AclBytesInUse + (Len(sCurrentACE) + _
      GetLengthSid(bUserSid(0))) * 2 - 4

   ' Resize our new ACL buffer to its proper size.
   ReDim bNewACL(lNewACLSize)

   ' Use the InitializeAcl API function call to initialize the new
   ' ACL.
   lResult = InitializeAcl(bNewACL(0), lNewACLSize, ACL_REVISION)

   ' A return code of zero means the call failed; test for this
   ' before continuing.
   If (lResult = 0) Then
      MsgBox "Error: Unable to Initialize New ACL"
      Exit Sub
   End If

   ' If a DACL is present, copy it to a new DACL.
   If (lDaclPresent) Then

      ' Copy the ACEs from the file to the new ACL.
      If (sACLInfo.AceCount > 0) Then

         ' Grab each ACE and stuff them into the new ACL.
         nRecordNumber = 0
         For I = 0 To (sACLInfo.AceCount - 1)

            ' Attempt to grab the next ACE.
            lResult = GetAce(pAcl, I, pCurrentAce)

            ' Make sure you have the current ACE under question.
            If (lResult = 0) Then
               MsgBox "Error: Unable to Obtain ACE (" & I & ")"
               Exit Sub
            End If

            ' You have a pointer to the ACE. Place it
            ' into a structure, so you can get at its size.
            CopyMemory sCurrentACE, pCurrentAce, LenB(sCurrentACE)

            'Skip adding the ACE to the ACL if this is same usersid
            lTempSid = pCurrentAce + 8
            If EqualSid(bUserSid(0), lTempSid) = 0 Then

                ' Now that you have the ACE, add it to the new ACL.
                lResult = AddAce(VarPtr(bNewACL(0)), ACL_REVISION, _
                  MAXDWORD, pCurrentAce, _
                  sCurrentACE.Header.AceSize)

                 ' Make sure you have the current ACE under question.
                 If (lResult = 0) Then
                   MsgBox "Error: Unable to Add ACE to New ACL"
                    Exit Sub
                 End If
                 nRecordNumber = nRecordNumber + 1
            End If

         Next I

         ' You have now rebuilt a new ACL and want to add it to
         ' the newly created DACL.
         lResult = AddAccessAllowedAce(bNewACL(0), ACL_REVISION, _
            lMask, bUserSid(0))

         ' Make sure added the ACL to the DACL.
         If (lResult = 0) Then
            MsgBox "Error: Unable to Add ACL to DACL"
            Exit Sub
         End If

         'If it's directory, we need to add inheritance staff.
         If GetAttr(sFileName) And vbDirectory Then

            ' Attempt to grab the next ACE which is what we just added.
            lResult = GetAce(VarPtr(bNewACL(0)), nRecordNumber, pCurrentAce)

            ' Make sure you have the current ACE under question.
            If (lResult = 0) Then
               MsgBox "Error: Unable to Obtain ACE (" & I & ")"
               Exit Sub
            End If
            ' You have a pointer to the ACE. Place it
            ' into a structure, so you can get at its size.
            CopyMemory sCurrentACE, pCurrentAce, LenB(sCurrentACE)
            sCurrentACE.Header.AceFlags = OBJECT_INHERIT_ACE + INHERIT_ONLY_ACE
            CopyMemory ByVal pCurrentAce, VarPtr(sCurrentACE), LenB(sCurrentACE)

            'add another ACE for files
            lResult = AddAccessAllowedAce(bNewACL(0), ACL_REVISION, _
               lMask, bUserSid(0))

            ' Make sure added the ACL to the DACL.
            If (lResult = 0) Then
               MsgBox "Error: Unable to Add ACL to DACL"
               Exit Sub
            End If

            ' Attempt to grab the next ACE.
            lResult = GetAce(VarPtr(bNewACL(0)), nRecordNumber + 1, pCurrentAce)

            ' Make sure you have the current ACE under question.
            If (lResult = 0) Then
               MsgBox "Error: Unable to Obtain ACE (" & I & ")"
               Exit Sub
            End If

            CopyMemory sCurrentACE, pCurrentAce, LenB(sCurrentACE)
            sCurrentACE.Header.AceFlags = CONTAINER_INHERIT_ACE
            CopyMemory ByVal pCurrentAce, VarPtr(sCurrentACE), LenB(sCurrentACE)
        End If


         ' Set the file's Security Descriptor to the new DACL.
         lResult = SetSecurityDescriptorDacl(sNewSD, 1, _
            bNewACL(0), 0)

         ' Make sure you set the SD to the new DACL.
         If (lResult = 0) Then
            MsgBox "Error: " & _
                "Unable to Set New DACL to Security Descriptor"
            Exit Sub
         End If

         ' The final step is to add the Security Descriptor back to
         ' the file!
         lResult = SetFileSecurity(sFileName, _
            DACL_SECURITY_INFORMATION, sNewSD)

         ' Make sure you added the Security Descriptor to the file!
         If (lResult = 0) Then
            MsgBox "Error: Unable to Set New Security Descriptor " _
               & " to File : " & sFileName
            MsgBox Err.LastDllError
         Else
            MsgBox "Updated Security Descriptor on File: " _
               & sFileName
         End If

      End If

   End If

End Sub


