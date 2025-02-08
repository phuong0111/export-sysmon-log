```bash
# Basic export for last 24 hours
Get-WinEvent -LogName "Microsoft-Windows-Sysmon/Operational" `
    -FilterHashTable @{
        LogName='Microsoft-Windows-Sysmon/Operational'
        StartTime=(Get-Date).AddHours(-24)
        EndTime=(Get-Date)
    } | Export-Csv -Path "C:\logs\sysmon_last24h.csv" -NoTypeInformation

# Export with specific date range
Get-WinEvent -LogName "Microsoft-Windows-Sysmon/Operational" `
    -FilterHashTable @{
        LogName='Microsoft-Windows-Sysmon/Operational'
        StartTime='2024-02-05T00:00:00'
        EndTime='2024-02-06T23:59:59'
    } | Export-Csv -Path "C:\logs\sysmon_custom_range.csv" -NoTypeInformation

# Export with specific EventIDs (e.g., process creation events)
Get-WinEvent -LogName "Microsoft-Windows-Sysmon/Operational" `
    -FilterHashTable @{
        LogName='Microsoft-Windows-Sysmon/Operational'
        StartTime=(Get-Date).AddHours(-24)
        EndTime=(Get-Date)
        ID=1  # Process creation events
    } | Export-Csv -Path "C:\logs\sysmon_process_creation.csv" -NoTypeInformation

# Export to XML format (preserves more detail)
Get-WinEvent -LogName "Microsoft-Windows-Sysmon/Operational" `
    -FilterHashTable @{
        LogName='Microsoft-Windows-Sysmon/Operational'
        StartTime=(Get-Date).AddHours(-24)
        EndTime=(Get-Date)
    } | Export-Clixml -Path "C:\logs\sysmon_detailed.xml"

# Export with custom fields selection
Get-WinEvent -LogName "Microsoft-Windows-Sysmon/Operational" `
    -FilterHashTable @{
        LogName='Microsoft-Windows-Sysmon/Operational'
        StartTime=(Get-Date).AddHours(-24)
        EndTime=(Get-Date)
    } | Select-Object TimeCreated, Id, Message | 
    Export-Csv -Path "C:\logs\sysmon_custom_fields.csv" -NoTypeInformation

# Export with filtering for specific processes
Get-WinEvent -LogName "Microsoft-Windows-Sysmon/Operational" `
    -FilterHashTable @{
        LogName='Microsoft-Windows-Sysmon/Operational'
        StartTime=(Get-Date).AddHours(-24)
        EndTime=(Get-Date)
    } | Where-Object {$_.Message -like "*cmd.exe*"} |
    Export-Csv -Path "C:\logs\sysmon_cmd_events.csv" -NoTypeInformation

Get-WinEvent -FilterHashTable @{ LogName = "Microsoft-Windows-Sysmon/Operational" }
Get-WinEvent -LogName "Microsoft-Windows-Sysmon/Operational" -FilterHashTable @{LogName='Microsoft-Windows-Sysmon/Operational';StartTime=(Get-Date).AddSeconds(-3600);EndTime=(Get-Date)} | Export-Csv -Path "C:\logs\sysmon_last1h.csv" -NoTypeInformation
```

```bash
<Sysmon schemaversion="4.70">
  <EventFiltering>
    <RuleGroup name="Privilege Elevation Detection" groupRelation="or">
      <!-- Process Creation with elevated privileges -->
      <ProcessCreate onmatch="include">
        <IntegrityLevel>High</IntegrityLevel>
        <ParentImage condition="is not">C:\Windows\System32\consent.exe</ParentImage>
      </ProcessCreate>

      <!-- Token Manipulation -->
      <ProcessAccess onmatch="include">
        <GrantedAccess>0x1F0FFF|0x1F1FFF|0x1F2FFF</GrantedAccess>
      </ProcessAccess>

      <!-- Registry auto-elevation paths -->
      <RegistryEvent onmatch="include">
        <TargetObject condition="contains">Software\Classes\ms-settings\shell\open</TargetObject>
        <TargetObject condition="contains">Software\Classes\Folder\shell\open</TargetObject>
        <TargetObject condition="contains">Software\Microsoft\Windows\CurrentVersion\Run</TargetObject>
        <TargetObject condition="contains">Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags</TargetObject>
      </RegistryEvent>
    </RuleGroup>
  </EventFiltering>
</Sysmon>
```