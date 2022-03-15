#load libraries
[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null;
[reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null;

#instantiate GUI objects
$dropList = New-Object System.Windows.Forms.ComboBox;
$execute = New-Object System.Windows.Forms.Button;
$vScrollBar = New-Object System.Windows.Forms.VScrollBar;
$csvFile = New-Object System.Windows.Forms.Button;
$outputLabel = New-Object System.Windows.Forms.Label;
$optionsLabel = New-Object System.Windows.Forms.Label;
$loadLabel = New-Object System.Windows.Forms.Label;
$textBox = New-Object System.Windows.Forms.TextBox;
$importCheck = New-Object System.Windows.Forms.CheckBox;
$exportCheck = New-Object System.Windows.Forms.CheckBox;
$loadServers = New-Object System.Windows.Forms.Button;
$loadServersCSV = New-Object System.Windows.Forms.Button;
$deployRoles = New-Object System.Windows.Forms.Button;
$serverList = New-Object System.Windows.Forms.ComboBox;


#global variables instantiation
$global:csv;
$global:csvSet = $false;
$global:inputSet = $false;
$global:csvInput;
$global:allServers = @();
$global:wsp;
$global:cookieContainer;
$global:creds;
$global:serversCSV;


[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

#build GUI
function Show-Window {
$BAMApp = New-Object System.Windows.Forms.Form;
$BAMApp.Text = "BAM Migration Utility";
$BAMApp.Name = "BAMApp";
$BAMApp.ClientSize = New-Object System.Drawing.Size(664, 300);

$dropList.Size = New-Object System.Drawing.Size(188, 210);
$dropList.Name = "dropList";
$dropList.Items.Add("removed.fqdn")| Out-Null;
$dropList.Items.Add("removed.fqdn")| Out-Null;
$dropList.Text = "removed.fqdn";
$dropList.Location = New-Object System.Drawing.Point(12, 12);
$dropList.AllowDrop = $True;

$serverList.Size = New-Object System.Drawing.Size(188, 210);
$serverList.Name = "serverList";
$serverList.Text = "Please Load Servers";
$serverList.Location = New-Object System.Drawing.Point(440, 160);
$serverList.AllowDrop = $True;

$csvFile.Name = "csvFile";
$csvFile.Size = New-Object System.Drawing.Size(190, 85);
$csvFile.Text = "Load CSV File";
$csvFile.Location = New-Object System.Drawing.Point(453, 20);
$csvFile.add_Click({Choose-File});

$loadServers.Name = "LoadServers";
$loadServers.Size = New-Object System.Drawing.Size(120, 30);
$loadServers.Text = "From BAM";
$loadServers.Location = New-Object System.Drawing.Point(300, 160);
$loadServers.add_Click({Load-Servers});

$loadServersCSV.Name = "loadServersCSV";
$loadServersCSV.Size = New-Object System.Drawing.Size(120, 30);
$loadServersCSV.Text = "From CSV";
$loadServersCSV.Location = New-Object System.Drawing.Point(300, 200);
$loadServersCSV.add_Click({Load-ServersCSV});

$deployRoles.Name = "deployRoles";
$deployRoles.Size = New-Object System.Drawing.Size(120, 30);
$deployRoles.Text = "Deploy Roles";
$deployRoles.Location = New-Object System.Drawing.Point(300, 240);
$deployRoles.add_Click({Deploy-Roles});
$deployRoles.Enabled = $false;

$outputLabel.Size = New-Object System.Drawing.Size(69, 23);
$outputLabel.Text = "Output";
$outputLabel.Location = New-Object System.Drawing.Point(145, 105);
$outputLabel.Name = "output";

$optionsLabel.Size = New-Object System.Drawing.Size(69, 23);
$optionsLabel.Text = "Options";
$optionsLabel.Location = New-Object System.Drawing.Point(320, 25);
$optionsLabel.Name = "option";

$loadLabel.Size = New-Object System.Drawing.Size(80, 23);
$loadLabel.Text = "Load Servers";
$loadLabel.Location = New-Object System.Drawing.Point(330, 140);
$loadLabel.Name = "option";

$textBox.Multiline = $True
$textBox.Size = New-Object System.Drawing.Size(200, 140);
$textBox.Name = "textBox";
$textBox.Location = New-Object System.Drawing.Point(70, 130);
$textBox.ScrollBars = "Vertical";

$exportCheck.Name = "export"
$exportCheck.Size = New-Object System.Drawing.Size(104, 24);
$exportCheck.Text = "Export"
$exportCheck.Location = New-Object System.Drawing.Point(320, 40);
$exportCheck.Add_Click({Switch-Export});

$importCheck.Name = "import"
$importCheck.Size = New-Object System.Drawing.Size(104, 24);
$importCheck.Text = "Import"
$importCheck.Location = New-Object System.Drawing.Point(320, 60);
$importCheck.Add_Click({Switch-Import});

$execute.Size = New-Object System.Drawing.Size(138, 44);
$execute.Text = "Get Options";
$execute.Add_Click({Execute});
$execute.Location = New-Object System.Drawing.Point (28, 50);
$execute.Enabled = $false;

$BAMApp.Controls.Add($dropList);
$BAMApp.Controls.Add($execute);
$BAMApp.Controls.Add($csvFile);
$BAMApp.Controls.Add($outputLabel);
$BAMApp.Controls.Add($textBox)
$BAMApp.Controls.Add($importCheck);
$BAMApp.Controls.Add($exportCheck);
$BAMApp.Controls.Add($loadServers);
$BAMApp.Controls.Add($loadServersCSV);
$BAMApp.Controls.Add($deployRoles);
$BAMApp.Controls.Add($serverList);
$BAMApp.Controls.Add($optionsLabel);
$BAMApp.Controls.Add($loadLabel)

$BAMApp.ShowDialog()| Out-Null;
}


#begin execution for options

function Execute {
    #start of export
    if ($exportCheck.Checked -eq $true) {
        $import = Import-CSV -Path $global:csvInput;
        if ($global:csvSet -eq $false) { 
            [System.Windows.Messagebox]::Show("Please choose export CSV")
        } else {
            Log-In;
            $returnedNetworks = @();
            $returnedRanges = @();
            $options = @();
            $fullOptions = @();
            foreach ($entry in $import) {
               $subnet = $entry.subnet;
               $addition = @{"Subnet" = $subnet};
               $network = $global:wsp.searchByObjectTypes($subnet,"IP4Network", 0, 9999);
               $range = $global:wsp.getEntities($network.id, "DHCP4Range", 0, 9999);
               try {
                  $options = $global:wsp.getDeploymentOptions($network.id, "DHCPV4ClientOption", -1);
                  $options += $global:wsp.getDeploymentOptions($network.id, "DHCPServiceOption", -1);
               } catch { Write-Host "No deployment options found."; }
                 # try {
                     #$options += $global:wsp.getDeploymentOptions($range.id, "DHCPV4ClientOption", -1);
                     #} catch { Write-Host "No options attached to range."; }
                   foreach ($option in $options) {
                       if ($option.name -eq "tftp-server-name") {
                            $option.name = "next-server";
                            $option.type = "DHCPService";
                       } elseif ($option.name -eq "boot-file-name") {
                            $option.name = "filename";
                            $option.type = "DHCPService";
                       }
                       Add-Member -InputObject $option -NotePropertyMembers $addition -PassThru;
                   }
               $fullOptions += $options;
            }
            #foreach ($result in $returnedNetworks) {
             #   $Nid = $result.id;
              #  $range = $SrcProxy.getEntities($Nid, "DHCP4Range", 0, 9999);
               # $returnedRanges += $range;
               # $lineItem = $srcProxy.getDeploymentOptions($id, "DHCPV4ClientOption",-1);
               # $fullOptions += $lineItem;
            #} 
            $fullOptions | Select Subnet, id, name, properties, type, value | Export-CSV -Path $global:csv -NoTypeInformation;
            $global:wsp.logout();
         }
    }

    #end of export

    #start of import
    elseif ($importCheck.Checked -eq $true) {
        $import = Import-CSV -Path $global:csvInput;
        $members = Get-Member -InputObject $import[0];
        if ($members.Name.Contains("properties")) {
            Log-In;
            $returnedNetworks = @();
            $returnedRanges = @();
            $options = @();
            foreach ($entry in $import) {
                $subnet = $entry.subnet;
                $network = $global:wsp.searchByObjectTypes($subnet,'IP4Network', 0, 9999);
                $parentId = $network.id;
                $name = $entry.name;
                $properties = $entry.properties;
                #remove superfluous properties
                if ($properties.Contains("server=")) {
                    $splitProps = $properties.split("|");
                    foreach ($element in $splitProps) {
                        if ($element.Contains("server")) {
                            $properties = $properties.Replace($element + "|", '');
                        }
                    }
                }
                $type = $entry.type;
                $value = $entry.value;
                $addition = @{"Subnet" = $subnet};
                Write-Host $subnet " " $parentId " " $name " " $properties " " $type " " $value;
                if ($type.equals("DHCPClient")) {
                    $global:wsp.addDHCPClientDeploymentOption($parentId, $name, $value, $properties);
                } elseif($type.equals("DHCPService")) {
                    $global:wsp.addDHCPServiceDeploymentOption($parentId, $name, $value, $properties);
                }
             }
             $options += $option;
             $global:wsp.logout();
        } else {
            [System.Windows.Messagebox]::Show("Incorrect CSV type loaded. Please load CSV with exported options."); 
        }
    }
    #end of import
}

#end execution

#set files
function Choose-File {
    if (($importCheck.Checked -eq $false) -and ($exportCheck.Checked -eq $false)) {
        $fileIn = New-Object system.windows.forms.openfiledialog;
        $fileIn.ShowDialog();
        $global:csvInput = $fileIn.FileName;
        $execute.Enabled = $true;
        $global:inputSet = $true;
        $textBox.Text = "Subnets Loaded `r`n`r`n";
        $import = Import-CSV -Path $global:csvInput;
        foreach ($item in $import) {
            $textBox.Text += $item.subnet + "`r`n";
            Write-Host $item.subnet;
        }
        [System.Windows.Messagebox]::Show("Please choose Import/Export");
    } elseif (($dropList.Text -eq "removed.fqdn") -or ($dropList.Text -eq "removed.fqdn")) {
            switch ($exportCheck.Checked) {
            $true { 
                $file = New-Object system.windows.forms.savefiledialog;
                $file.ShowDialog();
                $global:csv = $file.FileName;
                $global:csvSet = $true;
                $execute.Enabled = $true;
                break;
            }
            $false {
                $file = New-Object system.windows.forms.openfiledialog;
                $file.ShowDialog();
                $global:csv = $file.FileName;
                $execute.Enabled = $true;
                break;
            }
        }        
    } else { 
        [System.Windows.Messagebox]::Show("Please choose a Bluecat instance");
    }
}

#import option selected
function Switch-Import {
    if ($global:inputSet -eq $false) {
        [System.Windows.Messagebox]::Show("Please load a CSV file");
        $importCheck.Checked = $false;
    } else {
        $dropList.Text = "removed.fqdn";
        $csvFile.Enabled = $false;
        $exportCheck.Checked = $false;
        $execute.Text = "Load Options";
    }
}

#export option selected
function Switch-Export {
    if ($global:inputSet -eq $false) {
        [System.Windows.Messagebox]::Show("Please load a CSV file");
        $exportCheck.Checked = $false;
    } else {
        if ($csvFile.Enabled -eq $false) {
            $csvFile.Enabled = $true;
        }
        $csvFile.Text = "Choose CSV Output File";
        $importCheck.Checked = $false;
        $execute.Text = "Get Options";
    }
}

#get list of servers from BAM
function Get-Servers {
    Log-In;
    $i = 0;
    while ($i -le 320) {
        $currentServer = $global:wsp.getEntities(860, "Server", $i, 1);
        $name = $currentServer.name;
            #populate list of servers
            if ($name.Contains("corp") -or $name.Contains("ipam")) { $global:allServers += $currentServer; }
            else {
                Write-Host $name " skipped.";
            }
        $i = $i + 1;
    }
    #export server list to CSV file for subsequent uses
    $global:allServers | Select id, name, properties | Export-CSV -Path C:\servers.csv -NoTypeInformation;
    #log out
    $global:wsp.logout();
}

#load list of servers from BAM
function Load-Servers {
     Remove-Item "C:\servers.csv";
     $dropList.Text = "removed.fqdn";
     Get-Servers;
     $serverList.Text = $global:allServers[0].name;
     foreach ($server in $global:allServers) {
        $serverList.Items.Add($server.name)| Out-Null;
     }
     $deployRoles.Enabled = $true;
}

#load list of servers from CSV
function Load-ServersCSV {
    $dropList.Text = "removed.fqdn";
    $global:serversCSV = Import-CSV -Path C:\servers.csv;
    foreach ($server in $global:serversCSV) {
        $global:allServers += $server;
        $serverList.Items.Add($server.name) | Out-Null;
    }
    $serverList.Text = $global:allServers[0].name;
    $deployRoles.Enabled = $true;
}

#deploy roles to selected server
function Deploy-Roles {
    if ($global:inputSet -eq $false) {
        [System.Windows.Messagebox]::Show("Please load a CSV file");
    } else {
        $dropList.Text = "removed.fqdn";
        Log-In;
        $import = Import-CSV -Path $global:csvInput;
        $members = Get-Member -InputObject $import[0];
        #ensure CSV file contains subnets
        if ($members.Name.Contains("Subnet")) {
             $options = @();
             $selectedServer = $serverList.Text;
             #set correct server names based on selection
             if ($selectedServer.Contains("ipam01")) {
                $server1 = $selectedServer
                $server2 = $selectedServer.Replace("ipam01", "ipam02");
             } elseif ($selectedServer.Contains("ipam02")) {
                $server2 = $selectedServer;
                $server1 = $selectedServer.Replace("ipam02", "ipam01");
             } elseif ($selectedServer.Contains("dhcp01")) {
                $server1 = $selectedServer
                $server2 = $selectedServer.Replace("dhcp01", "dhcp02");
             } elseif ($selectedServer.Contains("dhcp02")) {
                $server2 = $selectedServer;
                $server1 = $selectedServer.Replace("dhcp02", "dhcp01");
             }
             #loop through CSV and add DHCP deployment role to subnets
             foreach ($entry in $import) {
                $subnet = $entry.subnet;
                $network = $global:wsp.searchByObjectTypes($subnet,'IP4Network', 0, 9999);
                $selection = $global:allServers | ? {$_.name -eq $server1 };
                $selection2 = $global:allServers | ? {$_.name -eq $server2 };
                $interface = $global:wsp.getEntities($selection.id, "NetworkServerInterface", 0 , 10);
                $interface2 = $global:wsp.getEntities($selection2.id, "NetworkServerInterface", 0, 10);
                $int2id = $interface2.id;
                $wsp.addDHCPDeploymentRole($network.id, $interface.id, "MASTER", "secondaryServerInterfaceId="+$int2id);
            }
            #deploy servers
            $global:wsp.deployServer($selection.id);
            $global:wsp.deployServer($selection2.id);
            #log out
            $global:wsp.logout();
        } else {
            [System.Windows.Messagebox]::Show("Incorrect CSV type loaded. Please load CSV with subnets.");
        }
    }
}

#log in to selected server's API
function Log-In {
    $baseURL = $dropList.Text;
        #set correct API URL based on selection
        if ($dropList.Text -eq "removed.fqdn") {
            $path = "https://$baseURL/Services/API?wsdl";
            Write-Host "Proteus selected.";    
            } else {
               $path = "http://$baseURL/Services/API?wsdl";
               Write-Host "BAM selected.";
            }        
    $global:cookieContainer = New-Object System.Net.CookieContainer;
    $global:wsp = New-WebServiceProxy -Uri $path;
    $global:wsp.CookieContainer = $global:cookieContainer;
    $global:wsp.Url = $path;
    $global:creds = Get-Credential;
    $global:wsp.login($global:creds.UserName, $global:creds.GetNetworkCredential().Password);
}

#show GUI
Show-Window;

