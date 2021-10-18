$folder = "C:\Program Files (x86)\Microsoft Intune Management Extension\Content\Incoming"

$xmlContents = @'
<Objs Version="1.1.0.1" xmlns="http://schemas.microsoft.com/powershell/2004/04">
  <Obj RefId="0">
    <TN RefId="0">
      <T>System.Security.AccessControl.DirectorySecurity</T>
      <T>System.Security.AccessControl.FileSystemSecurity</T>
      <T>System.Security.AccessControl.NativeObjectSecurity</T>
      <T>System.Security.AccessControl.CommonObjectSecurity</T>
      <T>System.Security.AccessControl.ObjectSecurity</T>
      <T>System.Object</T>
    </TN>
    <ToString>System.Security.AccessControl.DirectorySecurity</ToString>
    <Props>
      <S N="AccessRightType">System.Security.AccessControl.FileSystemRights</S>
      <S N="AccessRuleType">System.Security.AccessControl.FileSystemAccessRule</S>
      <S N="AuditRuleType">System.Security.AccessControl.FileSystemAuditRule</S>
      <B N="AreAccessRulesProtected">false</B>
      <B N="AreAuditRulesProtected">false</B>
      <B N="AreAccessRulesCanonical">true</B>
      <B N="AreAuditRulesCanonical">true</B>
    </Props>
    <MS>
      <S N="PSPath">Microsoft.PowerShell.Core\FileSystem::C:\Program Files (x86)\Microsoft Intune Management Extension\Content\Incoming</S>
      <S N="PSParentPath">Microsoft.PowerShell.Core\FileSystem::C:\Program Files (x86)\Microsoft Intune Management Extension\Content</S>
      <S N="PSChildName">Incoming</S>
      <Obj N="PSDrive" RefId="1">
        <TN RefId="1">
          <T>System.Management.Automation.PSDriveInfo</T>
          <T>System.Object</T>
        </TN>
        <ToString>C</ToString>
        <Props>
          <S N="CurrentLocation">temp</S>
          <S N="Name">C</S>
          <S N="Provider">Microsoft.PowerShell.Core\FileSystem</S>
          <S N="Root">C:\</S>
          <S N="Description">Acer</S>
          <Nil N="MaximumSize" />
          <Obj N="Credential" RefId="2">
            <TN RefId="2">
              <T>System.Management.Automation.PSCredential</T>
              <T>System.Object</T>
            </TN>
            <ToString>System.Management.Automation.PSCredential</ToString>
            <Props>
              <Nil N="UserName" />
              <Nil N="Password" />
            </Props>
          </Obj>
          <Nil N="DisplayRoot" />
        </Props>
        <MS>
          <U64 N="Used">235710025728</U64>
          <U64 N="Free">19153707008</U64>
        </MS>
      </Obj>
      <Obj N="PSProvider" RefId="3">
        <TN RefId="3">
          <T>System.Management.Automation.ProviderInfo</T>
          <T>System.Object</T>
        </TN>
        <ToString>Microsoft.PowerShell.Core\FileSystem</ToString>
        <Props>
          <S N="ImplementingType">Microsoft.PowerShell.Commands.FileSystemProvider</S>
          <S N="HelpFile">System.Management.Automation.dll-Help.xml</S>
          <S N="Name">FileSystem</S>
          <S N="PSSnapIn">Microsoft.PowerShell.Core</S>
          <S N="ModuleName">Microsoft.PowerShell.Core</S>
          <Nil N="Module" />
          <S N="Description"></S>
          <S N="Capabilities">Filter, ShouldProcess, Credentials</S>
          <S N="Home">C:\Users\markstan</S>
          <Obj N="Drives" RefId="4">
            <TN RefId="4">
              <T>System.Collections.ObjectModel.Collection`1[[System.Management.Automation.PSDriveInfo, System.Management.Automation, Version=3.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35]]</T>
              <T>System.Object</T>
            </TN>
            <LST>
              <Ref RefId="1" />
              <S>D</S>
            </LST>
          </Obj>
        </Props>
      </Obj>
      <Nil N="CentralAccessPolicyId" />
      <Nil N="CentralAccessPolicyName" />
      <S N="Path">Microsoft.PowerShell.Core\FileSystem::C:\Program Files (x86)\Microsoft Intune Management Extension\Content\Incoming</S>
      <S N="Owner">BUILTIN\Administrators</S>
      <S N="Group">NT AUTHORITY\SYSTEM</S>
      <Obj N="Access" RefId="5">
        <TN RefId="5">
          <T>System.Security.AccessControl.AuthorizationRuleCollection</T>
          <T>System.Collections.ReadOnlyCollectionBase</T>
          <T>System.Object</T>
        </TN>
        <IE>
          <S>System.Security.AccessControl.FileSystemAccessRule</S>
          <S>System.Security.AccessControl.FileSystemAccessRule</S>
          <S>System.Security.AccessControl.FileSystemAccessRule</S>
          <S>System.Security.AccessControl.FileSystemAccessRule</S>
          <S>System.Security.AccessControl.FileSystemAccessRule</S>
          <S>System.Security.AccessControl.FileSystemAccessRule</S>
        </IE>
      </Obj>
      <S N="Sddl">O:BAG:SYD:AI(A;OICI;FA;;;SY)(A;;0x100001;;;LS)(A;;0x1200a9;;;NS)(A;OICIIO;FA;;;NS)(A;;FA;;;BA)(A;;0x1200a9;;;BU)</S>
      <S N="AccessToString">NT AUTHORITY\SYSTEM Allow  FullControl_x000A_NT AUTHORITY\LOCAL SERVICE Allow  ReadData, Synchronize_x000A_NT AUTHORITY\NETWORK SERVICE Allow  ReadAndExecute, Synchronize_x000A_NT AUTHORITY\NETWORK SERVICE Allow  FullControl_x000A_BUILTIN\Administrators Allow  FullControl_x000A_BUILTIN\Users Allow  ReadAndExecute, Synchronize</S>
      <S N="AuditToString"></S>
    </MS>
  </Obj>
</Objs>


'@

$xmlContents | Out-File .\acl.xml
$IncomingAcls = Import-Clixml .\acl.xml

# disable inheritence - https://stackoverflow.com/questions/31721221/disable-inheritance-and-manually-apply-permissions-when-creating-a-folder-in-pow
$IncomingAcls.SetAccessRuleProtection($true,$false)

Set-Acl -Path $folder -AclObject $IncomingAcls