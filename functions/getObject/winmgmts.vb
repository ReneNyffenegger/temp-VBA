' option explicit

sub main() ' {

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
 & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")


Set colOSes = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
For Each objOS in colOSes
  debug.print "Computer Name: " & objOS.CSName

  debug.print "Operating System"
  debug.print "  Caption: " & objOS.Caption 'Name
  debug.print "  Version: " & objOS.Version 'Version & build
  debug.print "  BuildNumber: " & objOS.BuildNumber 'Build
  debug.print "  BuildType: " & objOS.BuildType
  debug.print "  OSProductSuite: " & objOS.OSProductsuite 'OS Product suite
  debug.print "  OSArchitecture: " & objOS.OSArchitecture
  debug.print "  OSType: " & objOS.OSType
  debug.print "  OtherTypeDescription: " & objOS.OtherTypeDescription
  debug.print "  ServicePackMajorVersion: " & objOS.ServicePackMajorVersion & "." & _
   objOS.ServicePackMinorVersion

Next

debug.print "Processors"

Set colCompSys = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
For Each objCS in colCompSys
  debug.print "  NumberOfProcessors: " & objCS.NumberOfProcessors
  debug.print "  NumberOfLogicalProcessors: " & objCS.NumberOfLogicalProcessors
  debug.print "  PCSystemType: " & objCS.PCSystemType
Next

Set colProcessors = objWMIService.ExecQuery("Select * from Win32_Processor")
For Each objProcessor in colProcessors
  debug.print "  Manufacturer: " & objProcessor.Manufacturer
  debug.print "  Name: " & objProcessor.Name
  debug.print "  Description: " & objProcessor.Description
  debug.print "  ProcessorID: " & objProcessor.ProcessorID
  debug.print "  Architecture: " & objProcessor.Architecture
  debug.print "  AddressWidth: " & objProcessor.AddressWidth
  debug.print "  NumberOfCores: " & objProcessor.NumberOfCores
  debug.print "  DataWidth: " & objProcessor.DataWidth
  debug.print "  Family: " & objProcessor.Family
  debug.print "  MaximumClockSpeed: " & objProcessor.MaxClockSpeed
Next

end sub ' }
