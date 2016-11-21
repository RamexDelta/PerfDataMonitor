Imports System
Imports System.Reflection
Imports System.Runtime.InteropServices

' General Information about an assembly is controlled through the following 
' set of attributes. Change these attribute values to modify the information
' associated with an assembly.

' Review the values of the assembly attributes

<Assembly: AssemblyTitle("EFB Performance Data Monitor")>
<Assembly: AssemblyDescription("")>
<Assembly: AssemblyCompany("Aer Lingus Flight Ops Engineering")>
<Assembly: AssemblyProduct("EFB Performance Data Monitor")>
<Assembly: AssemblyCopyright("Copyright ©  2016")> 
<Assembly: AssemblyTrademark("")> 

<Assembly: ComVisible(False)>

'The following GUID is for the ID of the typelib if this project is exposed to COM
<Assembly: Guid("61cde08f-0de6-4402-a346-892cc31695dc")>

' Version information for an assembly consists of the following four values:
'
'      Major Version
'      Minor Version 
'      Build Number
'      Revision
'
' You can specify all the values or you can default the Build and Revision Numbers 
' by using the '*' as shown below:
' <Assembly: AssemblyVersion("1.0.*")> 

'Rn Notes:
'Third number is number of days since 1-Jan-2000
'Fourth number is number of seconds since midnight, divided by 2
'See 826777 on StackOverflow

<Assembly: AssemblyVersion("1.1.*")>
<Assembly: AssemblyFileVersion("1.0.0.0")>

'Version 1.0: Initial Release. Installed on DEO and DVG,F/O Side
'Version 1.1: Additional erro handling and logging added
