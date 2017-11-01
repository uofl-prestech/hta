ComputerName = "."
Set wmiServices  = GetObject ( _
    "winmgmts:{impersonationLevel=Impersonate}!//" & ComputerName)
' Get physical disk drive
Set wmiDiskDrives =  wmiServices.ExecQuery ( "SELECT Caption, DeviceID FROM Win32_DiskDrive")

For Each wmiDiskDrive In wmiDiskDrives
    WScript.Echo "Disk drive Caption: " & wmiDiskDrive.Caption & VbNewLine & "DeviceID: " & " (" & wmiDiskDrive.DeviceID & ")"

    'Use the disk drive device id to
    ' find associated partition
    query = "ASSOCIATORS OF {Win32_DiskDrive.DeviceID='" _
        & wmiDiskDrive.DeviceID & "'} WHERE AssocClass = Win32_DiskDriveToDiskPartition"    
    Set wmiDiskPartitions = wmiServices.ExecQuery(query)

    For Each wmiDiskPartition In wmiDiskPartitions
        'Use partition device id to find logical disk
        Set wmiLogicalDisks = wmiServices.ExecQuery _
            ("ASSOCIATORS OF {Win32_DiskPartition.DeviceID='" _
             & wmiDiskPartition.DeviceID & "'} WHERE AssocClass = Win32_LogicalDiskToPartition") 

        For Each wmiLogicalDisk In wmiLogicalDisks
            WScript.Echo "Drive letter associated" _
                & " with disk drive = " _ 
                & wmiDiskDrive.Caption _
                & wmiDiskDrive.DeviceID _
                & VbNewLine & " Partition = " _
                & wmiDiskPartition.DeviceID _
                & VbNewLine & " is " _
                & wmiLogicalDisk.DeviceID
        Next      
    Next
Next