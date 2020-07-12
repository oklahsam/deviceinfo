Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

# Check for dependencies
write-host "Checking dependencies...."
if (get-installedmodule -name "proxx.snmp" -ErrorAction Ignore) { } else {
    write-host "`n`rPowershell SNMP module not found. Installing now...."
    Start-Process -FilePath powershell.exe -ArgumentList { install-packageprovider -name nuget -MinimumVersion 2.8.5.201 -force ; install-module "proxx.snmp" -AllowClobber -force } -verb RunAs -Wait
}
import-module proxx.snmp
$rsat_dhcp = Get-WindowsCapability -name RSAT.DHCP* -online
$rsat_AD = Get-WindowsCapability -name RSAT.ActiveDirectory* -online
if ( $rsat_dhcp.state -eq "NotPresent" ) {
    write-host "`n`rRSAT DHCP module not found. Installing now...."
    $rsat_dhcp | add-WindowsCapability -online
}
if ( $rsat_AD.state -eq "NotPresent" ) {
    write-host "`n`rRSAT Active Directory module not found. Installing now...."
    $rsat_ad | add-WindowsCapability -online
}
import-module activedirectory

# Hide PowerShell Console
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();
[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'
$consolePtr = [Console.Window]::GetConsoleWindow()
[Console.Window]::ShowWindow($consolePtr, 0)

# for talking across runspaces.
$sync = [Hashtable]::Synchronized(@{})



# ********** Modify these variables ************

$switches = "x.x.x.x","y.y.y.y","z.z.z.z"
$sync.community = "public"

# **********************************************



# Uses SNMP to get switch vendors
$sync.switches = foreach ( $switch in $switches ) {
    [pscustomobject]@{
        Vendor = ((invoke-snmpget -community $sync.community -ip $switch -OID 1.3.6.1.2.1.1.1.0).value -split " ")[0]
        Address = $switch
    }   
}
# Get list of computers of AD
$computers = get-adcomputer -filter * | select-object name | Sort-Object
$dc = (get-addomaincontroller).name

# Set up main window controls
$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '830,438'
$Form.text                       = "Device Info"
$Form.TopMost                    = $false
$Form.MaximizeBox                = $false
$Form.MinimizeBox                = $false
$Form.FormBorderStyle            = 'FixedDialog'

$ListBox1                        = New-Object system.Windows.Forms.ListBox
$ListBox1.text                   = "listBox"
$ListBox1.width                  = 150
$ListBox1.height                 = 415
$ListBox1.location               = New-Object System.Drawing.Point(10,11)

$button1                         = New-Object System.Windows.Forms.Button
$button1.text                     = "Start"
$button1.width                    = 60
$button1.height                   = 25
$button1.location                 = New-Object System.Drawing.Point(17,135)
$button1.Font                     = 'Microsoft Sans Serif,10'

$Label1                          = New-Object system.Windows.Forms.Label
$Label1.text                     = "Canonical Name: "
$Label1.AutoSize                 = $true
$Label1.width                    = 25
$Label1.height                   = 10
$Label1.location                 = New-Object System.Drawing.Point(17,10)
$Label1.Font                     = 'Microsoft Sans Serif,10'

$Label2                          = New-Object system.Windows.Forms.Label
$Label2.text                     = "OS: "
$Label2.AutoSize                 = $true
$Label2.width                    = 25
$Label2.height                   = 10
$Label2.location                 = New-Object System.Drawing.Point(95,33)
$Label2.Font                     = 'Microsoft Sans Serif,10'

$Label3                          = New-Object system.Windows.Forms.Label
$Label3.text                     = "OS Version: "
$Label3.AutoSize                 = $true
$Label3.width                    = 25
$Label3.height                   = 10
$Label3.location                 = New-Object System.Drawing.Point(48,54)
$Label3.Font                     = 'Microsoft Sans Serif,10'

$Label4                          = New-Object system.Windows.Forms.Label
$Label4.text                     = "Last Logon Date: "
$Label4.AutoSize                 = $true
$Label4.width                    = 25
$Label4.height                   = 10
$Label4.location                 = New-Object System.Drawing.Point(19,77)
$Label4.Font                     = 'Microsoft Sans Serif,10'

$Label5                          = New-Object system.Windows.Forms.Label
$Label5.text                     = "IPv4 Address: "
$Label5.AutoSize                 = $true
$Label5.width                    = 25
$Label5.height                   = 10
$Label5.location                 = New-Object System.Drawing.Point(36,98)
$Label5.Font                     = 'Microsoft Sans Serif,10'

$Label6                          = New-Object system.Windows.Forms.linkLabel
$Label6.text                     = "MAC Address: "
$Label6.AutoSize                 = $true
$Label6.width                    = 25
$label6.linkcolor                = "Black"
$label6.linkbehavior             = "NeverUnderline"
$Label6.height                   = 10
$Label6.location                 = New-Object System.Drawing.Point(33,119)
$Label6.Font                     = 'Microsoft Sans Serif,10'

$Label7                          = New-Object system.Windows.Forms.Label
$Label7.text                     = "Online: "
$Label7.AutoSize                 = $true
$Label7.width                    = 25
$Label7.height                   = 10
$Label7.location                 = New-Object System.Drawing.Point(76,139)
$Label7.Font                     = 'Microsoft Sans Serif,10'

$TextBox2                        = New-Object system.Windows.Forms.TextBox
$TextBox2.multiline              = $true
$TextBox2.width                  = 600
$TextBox2.height                 = 210
$TextBox2.text                   = ""
$TextBox2.location               = New-Object System.Drawing.Point(17,165)
$TextBox2.Font                   = 'Consolas,10'
$TextBox2.ReadOnly               = $true

$TextBox1                        = New-Object system.Windows.Forms.TextBox
$TextBox1.multiline              = $true
$TextBox1.width                  = 600
$TextBox1.height                 = 210
$TextBox1.text                   = ""
$TextBox1.location               = New-Object System.Drawing.Point(17,165)
$TextBox1.Font                   = 'Consolas,10'
$TextBox1.ReadOnly               = $true

$RadioButton1                    = New-Object system.Windows.Forms.RadioButton
$RadioButton1.text               = "IP Address"
$RadioButton1.AutoSize           = $true
$RadioButton1.width              = 104
$RadioButton1.height             = 20
$RadioButton1.location           = New-Object System.Drawing.Point(17,40)
$RadioButton1.Font               = 'Microsoft Sans Serif,10'
$RadioButton1.Checked            = $true

$RadioButton2                    = New-Object system.Windows.Forms.RadioButton
$RadioButton2.text               = "MAC Address"
$RadioButton2.AutoSize           = $true
$RadioButton2.width              = 104
$RadioButton2.height             = 20
$RadioButton2.location           = New-Object System.Drawing.Point(17,60)
$RadioButton2.Font               = 'Microsoft Sans Serif,10'

$tabcontrol1 = new-object 'system.windows.forms.tabcontrol'
$tabcontrol1.Alignment = 'top'
$tabcontrol1.Location = '170, 10'
$tabcontrol1.Multiline = $True
$tabcontrol1.Name = 'tabcontrol1'
$tabcontrol1.SelectedIndex = 0
$tabcontrol1.Size = '650, 410'
$tabcontrol1.TabIndex = 0

$TextBox3                        = New-Object system.Windows.Forms.TextBox
$TextBox3.multiline              = $false
$TextBox3.width                  = 150
$TextBox3.height                 = 25
$TextBox3.text                   = ""
$TextBox3.location               = New-Object System.Drawing.Point(17,10)
$TextBox3.Font                   = 'Microsoft Sans Serif,10'

$tabpage1 = new-object 'system.windows.forms.tabpage'
$tabpage1.Location = '42, 4'
$tabpage1.Padding = '3, 3, 3, 3'
$tabpage1.Size = '583, 442'
$tabpage1.TabIndex = 0
$tabpage1.Text = 'Info'
$tabpage1.font = 'Microsoft Sans Serif,10'
$tabpage1.UseVisualStyleBackColor = $True

$tabpage2 = new-object 'system.windows.forms.tabpage'
$tabpage2.Location = '42, 4'
$tabpage2.Padding = '3, 3, 3, 3'
$tabpage2.Size = '583, 442'
$tabpage2.TabIndex = 0
$tabpage2.Text = 'Port Lookup'
$tabpage2.font = 'Microsoft Sans Serif,10'
$tabpage2.UseVisualStyleBackColor = $True

$tabcontrol1.Controls.Addrange(@($tabpage1,$tabpage2))

$sync.startbutton = $button1
$sync.adcomp = $ListBox1
$sync.caname = $Label1
$sync.os = $Label2
$sync.osver = $Label3
$sync.lastlog = $Label4
$sync.ip = $Label5
$sync.mac = $label6
$sync.online = $Label7
$sync.userlog = $TextBox1
$sync.portconfig = $TextBox2
$sync.portconfigb = $textbox1
$sync.iphostlookup = $radiobutton1
$sync.maclookup = $RadioButton2
$sync.input = $TextBox3

$Form.controls.AddRange(@($sync.adcomp,$tabcontrol1,$tabcontrol2))
$tabpage1.controls.AddRange(@($sync.caname,$sync.os,$sync.osver,$sync.lastlog,$sync.ip,$sync.mac,$sync.online,$sync.portconfig))
$tabpage2.controls.addrange(@($sync.portconfigb,$sync.iphostlookup,$sync.maclookup,$sync.input,$sync.startbutton))

foreach ($line in $computers) { [void] $sync.adcomp.Items.Add($line.name) }

# Button to start IP/MAC search
$sync.startbutton.add_click({
    $script:switchport.runspace.dispose()
    $script:switchport.dispose()
    $script:switchport = [PowerShell]::Create().AddScript({
        $sync.button = $true
        invoke-expression $sync.portrun
        $sync.button = $false
        $script:switchport.runspace.dispose()
        $script:switchport.dispose()
    })
    $runspace = [RunspaceFactory]::CreateRunspace()
    $runspace.ApartmentState = "STA"
    $runspace.ThreadOptions = "ReuseThread"
    $runspace.Open()
    $runspace.SessionStateProxy.SetVariable("sync", $sync)
    $script:switchport.Runspace = $runspace
    $script:switchport.BeginInvoke()
}) 

# Initiates device lookup when selection changes
$sync.adcomp.add_selectedindexchanged({
    $script:connection.runspace.dispose()
    $script:connection.dispose()
    $script:connection = [PowerShell]::Create().AddScript({
        Invoke-Expression $sync.infotab
        $script:connection.runspace.dispose()
        $script:connection.dispose()
    })
    $runspace = [RunspaceFactory]::CreateRunspace()
    $runspace.ApartmentState = "STA"
    $runspace.ThreadOptions = "ReuseThread"
    $runspace.Open()
    $runspace.SessionStateProxy.SetVariable("sync", $sync)
    $script:connection.Runspace = $runspace
    $script:connection.BeginInvoke()
})

# Clears IP/MAC text box when selection changed
$sync.maclookup.add_checkedchanged({
    $sync.input.text = ""
})

# Copy MAC Address to clipboard when clicked
$sync.mac.add_click({
    $sync.mac.text -replace "MAC Address: ","" | clip
})

# Function that runs initially when searching by MAC/IP Address
function PORTRUN {
    $sync.portconfigb.text = "Searching for device...."
    $sync.location.clear()
    $sync.port.clear()
    $sync.description.clear()
    if ($sync.iphostlookup.checked -eq $true) {
        $sync.ipv4address = $sync.input.text
        foreach ($line in $sync.switches) { 
            $sync.macaddress = (invoke-snmpwalk -community $sync.community -ip $line.address -OID "1.3.6.1.2.1.4.22.1.2" | where-object { $_.oid -match ($sync.ipv4address + "$") }).value 
            if ($sync.macaddress) { break }
        }
    } 
    if ($sync.maclookup.checked -eq $true) {
        $sync.macaddress = $sync.input.text
    }
    if ( $sync.macaddress ) {
        invoke-expression $sync.portinfo
    } else {
        $sync.portconfigb.text = "No MAC Address found. Cannot search for device."
    }
}
$sync.portrun = get-content FUNCTION:\PORTRUN

# Function that finds device info and IP/MAC Address
function INFOTAB {
    $sync.caname.text = "Canonical Name: "
    $sync.os.text = "OS: "
    $sync.osver.text = "OS Version: "
    $sync.lastlog.text = "Last Logon Date: "
    $sync.ip.text = "IPv4 Address: "
    $sync.online.text = "Online: "
    $sync.mac.text = "MAC Address: "
    $sync.macaddress = ""
    $sync.portconfig.text = "Looking up IP Address...."
    $sync.comp = $sync.adcomp.selecteditem
    $comp = $sync.comp
    $sync.ad = get-adcomputer -filter {name -like $comp} -properties *
    $sync.caname.text = "Canonical Name: " + $sync.ad.canonicalname
    $sync.os.text = "OS: " + $sync.ad.operatingsystem
    $sync.osver.text = "OS Version: " + $sync.ad.operatingsystemversion
    $sync.lastlog.text = "Last Logon Date: " + $sync.ad.lastlogondate
    $sync.ipv4address = if ($sync.ad.ipv4address) { $sync.ad.ipv4address } else { "Not available" }
    $sync.ip.text = "IPv4 Address: " + $sync.ipv4address
    $sync.portconfig.text = "Finding MAC Address...."
    $scope = ((get-DhcpServerv4Scope -ComputerName $dc).scopeid).ipaddresstostring
    foreach ( $id in $scope ) {
        $sync.macaddress = ((Get-DhcpServerv4Lease -EA SilentlyContinue -computername $dc -scopeid $id | where-object { $_.ipaddress -eq $sync.ad.ipv4address }).clientid).toupper() -replace "-"," "
        if (-not ([string]::IsNullOrEmpty($sync.macaddress))) { break }
    }
    $sync.macaddress = if ($sync.macaddress) { 
        $sync.macaddress 
    } else { 
        $sync.portconfig.text += "."
        foreach ($line in $sync.switches) { 
            $sync.portconfig.text += "."
            $sync.macaddress = (invoke-snmpwalk -community $sync.community -ip $line.address -OID "1.3.6.1.2.1.4.22.1.2" | where-object { $_.oid -match ($sync.ipv4address + "$") }).value 
            if ( $sync.macaddress ) { break }
        }
        if ( $sync.macaddress ) { $sync.macaddress } else { "Not available" }
    }
    $sync.mac.text = "MAC Address: " + $sync.macaddress
    if ($sync.ad.ipv4address) {
        $sync.portconfig.text = "Testing Connection...."
        $test = (test-connection $sync.comp -count 1 -quiet)
        $sync.online.text = "Online: " + $test
    } else {
        $sync.online.text = "Online: False"
    }
    $sync.portconfig.text = "Searching for device...."
    if ( $sync.macaddress -ne "Not Available" ) {
        invoke-expression $sync.portinfo
    } else {
        $sync.portconfig.text = "No MAC Address found. Cannot search for device."
    }

}
$sync.infotab = get-content FUNCTION:\INFOTAB

# Function that takes IP/MAC Address and searches the list of switches using SNMP
function PORTINFO {
    $sync.macaddressb = (($sync.macaddress -replace "-","" -replace ":","" -replace " ","").toupper()) | where-object {$_.trim() -ne ""}
    $sync.result = foreach ($line in $sync.switches) {
        if ($line.vendor -ne "Cisco"){ 
            $sync.macoid = ( invoke-snmpwalk -community $sync.community -ip $line.address -OID ".1.3.6.1.2.1.17.4.3.1.1" | where-object { ($_.value -replace " ","") -eq $sync.macaddressb }).OID -replace ".4.3.1.1",".4.3.1.2"
            switch ($sync.button) {
                $false { $sync.portconfig.text += "." }
                $true { $sync.portconfigb.text += "." }
            }
        } else {
            $sync.macoidb.clear()
            $vlan = (invoke-snmpwalk -community $sync.community -ip $line.address -OID ".1.3.6.1.4.1.9.9.128.1.1.1.1.3").oid
            $sync.VLANs = foreach ($line2 in $vlan) {$line2.substring(30) -replace "\.","" -replace ".$",""}
            foreach ( $vlan in $sync.vlans ) {
                $cisco_community = $sync.community + "@" + $vlan
                $sync.macoidb = (invoke-snmpwalk -community $cisco_community -ip $line.address -OID ".1.3.6.1.2.1.17.4.3.1.1" | where-object { ($_.value -replace " ","") -eq $sync.macaddressb }).OID -replace ".4.3.1.1",".4.3.1.2"
                switch ($sync.button) {
                    $false { $sync.portconfig.text += "." }
                    $true { $sync.portconfigb.text += "." }
                }
                if ($sync.macoidb) {
                    $community = $cisco_community
                    $sync.macoid = $sync.macoidb
                    break
                }
            }
        }
        if ($line.vendor -ne 'Cisco') { $community = $sync.community}
        $sync.bridge = (invoke-snmpget -community $community -ip $line.address -OID $sync.macoid ).value
        $sync.ifindex = (invoke-snmpget -community $community -ip $line.address -OID (".1.3.6.1.2.1.17.1.4.1.2." + $sync.bridge)).value
        [pscustomobject]@{
            Switch = (invoke-snmpget -community $community -ip $line.address -OID "1.3.6.1.2.1.1.5.0" ).value
            Port = if ( $sync.ifindex ) { (invoke-snmpget -community $community -ip $line.address -OID (".1.3.6.1.2.1.31.1.1.1.1." + $sync.ifindex) ).value } else { "Unable to locate device" }
            Description =  if ( $sync.ifindex ) { (invoke-snmpget -community $community -ip $line.address -OID (".1.3.6.1.2.1.31.1.1.1.18." + $sync.ifindex)).value }
        }
    }
    switch ($sync.button) {
        $false { $sync.portconfig.text = $sync.result | format-table | out-string }
        $true { $sync.portconfigb.text = $sync.result | format-table | out-string }
    }
    $sync.result = ""
    $sync.macaddress = ""
    $sync.macaddressb = ""
    $sync.macoid = ""
    $sync.macoidb = ""
    $sync.bridge = ""
    $sync.ifindex = ""

}
$sync.portinfo = get-content FUNCTION:\PORTINFO

$sync.button = $false

[void]$Form.ShowDialog()

$script:connection.runspace.dispose()
$script:connection.dispose()
$script:switchport.runspace.dispose()
$script:switchport.dispose()

stop-process $pid
