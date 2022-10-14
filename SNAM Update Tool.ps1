#© Copyright 2022, Jacob Searcy, All rights reserved.

function Get-SNAMInfo {

    [CmdletBinding()]

    Param (
        [String]$ComputerName,
        [String]$SerialNumber,
        [Switch]$Swap,
        [Switch]$Inventory,
        [String]$user,
        [String]$pass,
        [String]$instance_name
    )

    Begin {
        $global:consumables_ids = New-Object System.Collections.ArrayList
        $global:consumables_list = New-Object System.Collections.ArrayList
        If (-not($SerialNumber)) {
            If (Test-Connection $ComputerName -Count 1 -Quiet) {
                $SerialNumber = (Get-WmiObject Win32_BIOS -ComputerName $ComputerName).SerialNumber
            }else{
                #Create Serial Number Entry Box
                $offline_form = New-Object System.Windows.Forms.Form
                $offline_form.Size = New-Object System.Drawing.Size(250,160)
                $offline_form.Text = "Offline"
                $offline_form.StartPosition = 'CenterScreen'

                $offline_label1 = New-Object System.Windows.Forms.Label
                $offline_label1.Text = "$ComputerName is Offline"
                $offline_label1.ForeColor = 'RED'
                $offline_label1.Location = New-Object System.Drawing.Point(5,8)
                $offline_label1.AutoSize = $true
                $offline_form.Controls.Add($offline_label1)

                $offline_label2 = New-Object System.Windows.Forms.Label
                $offline_label2.Text = "Manually Enter Serial Number:"
                $offline_label2.Location = New-Object System.Drawing.Point(5,32)
                $offline_label2.Size = New-Object System.Drawing.Size(200,20)
                $offline_form.Controls.Add($offline_label2)

                $offline_box = New-Object System.Windows.Forms.TextBox
                $offline_box.Location = New-Object System.Drawing.Point(5,58)
                $offline_box.Size = New-Object System.Drawing.Size(223,21)
                $offline_form.Controls.Add($offline_box)

                $OKButton = New-Object System.Windows.Forms.Button
                $OKButton.Location = New-Object System.Drawing.Point(155,90)
                $OKButton.Size = New-Object System.Drawing.Size(75,23)
                $OKButton.Text = "OK"
                $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
                $offline_form.Controls.Add($OKButton)
                $offline_form.AcceptButton = $OKButton

                $CancelButton = New-Object System.Windows.Forms.Button
                $CancelButton.Location = New-Object System.Drawing.Point(78,90)
                $CancelButton.Size = New-Object System.Drawing.Size(75,23)
                $CancelButton.Text = "Cancel"
                $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                $offline_form.Controls.Add($CancelButton)

                $offline_form.TopMost = $true
                $offline_result = $offline_form.ShowDialog()

                If ($offline_result -eq [System.Windows.Forms.DialogResult]::OK) {
                    $global:SerialNumber = $offline_box.Text
                }
                If ($offline_result -eq [System.Windows.Forms.DialogResult]::Cancel) {
                    exit
                }
            }
        }
        $global:SerialNumber = $SerialNumber
        #Get User Name of Tech Running Script
        $UserName = $((Get-WmiObject -Class win32_computersystem).UserName).Split("\")[1]
        $ScriptUser = Get-ADUser $UserName
        If ($ScriptUser.Surname -match "_") {
            $ScriptUser = "$($ScriptUser.GivenName) $($($ScriptUser.Surname).Split('_')[0])"
        }else{
            $ScriptUser = "$($ScriptUser.GivenName) $($ScriptUser.Surname)"
        }
        $SNAM_Path = "$PSScriptRoot\Files\SNAM Updates.csv"
        #ServiceNow API GET Variables
        $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $user, $pass)))
        $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
        $headers.Add('Authorization',('Basic {0}' -f $base64AuthInfo))
        $headers.Add('Accept','application/json')
        $headers.Add('Content-Type','application/json')
        $method = "get"
        #Get Script User
        $uri = "https://$instance_name/api/now/table/sys_user?sysparm_query=name%3D$ScriptUser"
        $scriptuser_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
        $global:ScriptUser = $scriptuser_response.result.name
        $global:ScriptUser_ID = $scriptuser_response.result.sys_id
        #Get State Options
        $uri = "https://$instance_name/api/now/table/sys_choice?sysparm_query=name%3Dalm_hardware^element%3Dinstall_status"
        $StateOptions = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
        #Get IT sys_id
        $IT_depart = "IT - Helpdesk / Desktop"
        $uri = "https://$instance_name/api/now/table/cmn_department?sysparm_query=name%3D$IT_depart"
        $IT_department = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
        $IT_department_id = $($IT_department.result | Where-Object {$_.id -eq 8247}).sys_id

        $global:install_status_selection = "1"
    }

    Process {
    #Get ServiceNow Information
        $continue = "Yes"
        While ($continue -eq "Yes") {
            #GET SNAM Record(s) for Non-Retired $SerialNumber and/or $ComputerName
            $install_status_filter = $($StateOptions.result | Where-Object {$_.label -eq 'Retired'}).value
            If ($ComputerName) {
                $uri = "https://$instance_name/api/now/table/alm_hardware?sysparm_query=serial_number%3D$SerialNumber^ORu_computer_name%3D$ComputerName^install_status!=$install_status_filter"
            }else{
                $uri = "https://$instance_name/api/now/table/alm_hardware?sysparm_query=serial_number%3D$SerialNumber^install_status!=$install_status_filter"
            }
            $serial_hw_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
        
            $total_results = $serial_hw_response.result.sys_id.count
            $entry = @{}
            for ($i = 1; $i -le $total_results; $i++) {
                #Hardware Table
                $comp_hw_response = $serial_hw_response.result[$i-1]
                #Hardware Table ID
                $comp_sys_id = $comp_hw_response.sys_id
                #Configuration Item ID
                $comp_ci_id = $comp_hw_response.ci.value
                If ($comp_ci_id -notmatch '\w') {
                    $po_number = $comp_hw_response.po_number
                    #Add Entry to Error Document
                $SNAM_Error_Path = "$PSScriptRoot\Files\SNAM Errors.csv"
                $error.Clear()
                $UserName = $((Get-WmiObject -Class win32_computersystem).UserName).Split("\")[1]
                If ($error) {
                    $UserName = "sysklindstrom"
                }
                New-Object -TypeName PSCustomObject -Property @{
                    "User" = $UserName
                    "Computer ID" = $comp_sys_id
                    "Parent ID" = $null
                    "Serial Number" = $SerialNumber
                    "Retired Serial Numbers" = $null
                    "Serials with Wrong Names" = $null
                    "Duplicate Serial Numbers" = $null
                    "Comments" = "One record doesn't have a configuration item"
                } | Select-Object "User","Computer ID","Parent ID","Serial Number","Retired Serial Numbers","Serials with Wrong Names","Duplicate Serial Numbers","Comments" | Export-Csv -Path $SNAM_Error_Path -NoTypeInformation -Append -Encoding ASCII
                    #Warning Box
                    $answer = [System.Windows.Forms.MessageBox]::Show("A record matching $ComputerName or $SerialNumber does not have a configuration item. Would you like to mark it for deletion?", "Duplicate Record", 4)
                    If ($answer -eq "Yes") {
                        $method = 'patch'
                        $uri = "https://$instance_name/api/now/table/alm_hardware/$comp_sys_id"
                        $on_order = $($StateOptions.result | Where-Object {$_.label -eq 'On order'}).value
                        $body = "{`"serial_number`":`"DELETE`",`"install_status`":`"$on_order`",`"managed_by`":`"$IT_department_id`"}"
                        $remove_entry = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body
                        $method = 'get'
                    }
                }
                #Configuration Item Table
                $comp_ci_link = $comp_hw_response.ci.link
                $uri = $comp_ci_link
                If ($comp_ci_link -notmatch '\w') {
                    $comp_ci_response = $null
                }else{
                    $comp_ci_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                }
                #Asset Table ID
                If ($comp_ci_response -ne $null) {
                    $asset_sys_id = $comp_ci_response.result.asset.value
                    $asset_link = $comp_ci_response.result.asset.link
                }else{
                    $asset_sys_id = $null
                    $asset_link = $null
                }
                #Consumables Table
                If ($asset_sys_id -ne $null) {
                    $uri = "https://$instance_name/api/now/table/alm_consumable?sysparm_query=parent%3D$asset_sys_id"
                    $comp_consumables_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                }else{
                    If ($($comp_ci_response.result.asset.value) -ne $null) {
                    $uri = "https://$instance_name/api/now/table/alm_consumable?sysparm_query=parent%3D$($comp_ci_response.result.asset.value)"
                    $comp_consumables_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                    }
                }

            #Hardware Table Results
            #$comp_hw_response.result
                #Computer Name
                $Name = $comp_hw_response.u_computer_name
                #Display Name
                $DisplayName = $comp_hw_response.display_name
                #Model Category
                $ModelCategory = $comp_hw_response.model_category
                if ($comp_hw_response.model_category.value -ne $null) {
                    $model_category_id = $comp_hw_response.model_category.value
                    $uri = $comp_hw_response.model_category.link
                    $ModelCategory_query = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                    $ModelCategory = $ModelCategory_query.result.Name
                }
                #Model
                $Model = $comp_hw_response.model
                if ($comp_hw_response.model.value -ne $null) {
                    $model_id = $comp_hw_response.model.value
                    $uri = $comp_hw_response.model.link
                    $Model_query = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                    $Model = $Model_query.result.Name
                }
                #Configuration Item
                $ConfigurationItem = $comp_hw_response.ci
                if ($comp_hw_response.ci.value -ne $null) {
                    $ci_id = $comp_hw_response.ci.value
                    $uri = $comp_hw_response.ci.link
                    $ConfigurationItem_query = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                    $ConfigurationItem = $ConfigurationItem_query.result.name
                }
                #Asset Tag
                $AssetTag = $comp_hw_response.asset_tag
                #Serial Number
                $SNAM_Serial = $comp_hw_response.serial_number
                #State
                $InstallStatus = $comp_hw_response.install_status
                $install_status = $($StateOptions.result | Where-Object {$_.value -eq $InstallStatus}).label
                #Substate
                $SubState = $comp_hw_response.substatus
                #Function
                $Function = $comp_hw_response.u_device_function
                #Assigned To
                $AssignedTo = $comp_hw_response.assigned_to
                if ($comp_hw_response.assigned_to.value -ne $null) {
                    $assignedto_id = $comp_hw_response.assigned_to.value
                    $uri = $comp_hw_response.assigned_to.link
                    $AssignedTo_query = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                    $AssignedTo = $AssignedTo_query.result.name
                }
                #Responsible Department
                $Management = $comp_hw_response.managed_by
                if ($comp_hw_response.managed_by.value -ne $null) {
                    $management_id = $comp_hw_response.managed_by.value
                    $uri = $comp_hw_response.managed_by.link
                    $Management_query = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                    $Management = $Management_query.result.name
                    $Management_costcenter = $Management_query.result.id
                }
                #Parent
                $Parent = $comp_hw_response.parent
                if ($comp_hw_response.parent.value -ne $null) {
                    $parent_id = $comp_hw_response.parent.value
                    $uri = $comp_hw_response.parent.link
                    $global:comp_parent_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                }
                #Last Physical Inventory
                $InventoryDate = $comp_hw_response.u_last_physical_inventory
                #Last Inventoried By
                $UpdatedBy = $comp_hw_response.u_last_inventory_by
                if ($comp_hw_response.u_last_inventory_by.value -ne $null) {
                    $update_id = $comp_hw_response.u_last_inventory_by.value
                    $uri = $comp_hw_response.u_last_inventory_by.link
                    $UpdatedBy_query = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                    $UpdatedBy = $UpdatedBy_query.result.name
                }
                #Machine Model Number
                $MachineModel = $comp_hw_response.u_machine_model_number
                #IP Address
                $IPAddress = $comp_hw_response.u_ip_address
                #Location
                $Location = $comp_hw_response.location
                if ($comp_hw_response.location.value -ne $null) {
                    $location_id = $comp_hw_response.location.value
                    $uri = $comp_hw_response.location.link
                    $Location_query = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                    $Location = $Location_query.result.name
                }
                #Department
                $Department = $comp_hw_response.department
                if ($comp_hw_response.department.value -ne $null) {
                    $department_id = $comp_hw_response.department.value
                    $uri = $comp_hw_response.department.link
                    $Department_query = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                    $Department = $Department_query.result.name
                    $Department_costcenter = $Department_query.result.id
                }
                #Assigned Date
                $AssignedDate = $comp_hw_response.assigned
                #Installed Date
                $InstallDate = $comp_hw_response.install_date
                #Comments
                $Comments = $comp_hw_response.comments
                #Cost Center
                if ($Department_costcenter -eq $Management_costcenter) {
                    $CostCenter = $Management_costcenter
                }else{
                    $CostCenter = echo $Management_costcenter $Department_costcenter
                }
                #PO Number
                $PO = $comp_hw_response.po_number
                If ($PO -notmatch '\w') {
                    If ($po_number -match '\w') {
                        $PO = $po_number
                    }
                }
                #Warranty Date
                $WarrantyDate = $comp_hw_response.warranty_expiration
                #StockRoom
                $StockRoom = $comp_hw_response.stockroom
                if ($comp_hw_response.stockroom.value -ne $null) {
                    $stockroom_id = $comp_hw_response.stockroom.value
                    $uri = $comp_hw_response.stockroom.link
                    $StockRoom_query = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                    $StockRoom = $StockRoom_query.result.name
                }
                #sys_id
                $hw_sys_id = $comp_hw_response.sys_id

            #Configuration Item Table Results
            #$comp_ci_response.result
                If ($comp_ci_response -ne $null) {
                    #Most Recent Discovery
                    $DiscoveryDate = $comp_ci_response.result.last_discovered
                    #Discovery Source
                    $DiscoverySource = $comp_ci_response.result.discovery_source
                    #sys_id
                    $ci_sys_id = $comp_ci_response.result.sys_id
                }else{
                    $DiscoveryDate = $null
                    $DiscoverySource = $null
                    $ci_sys_id = $null
                }

            #Consumables Table Results
            #$comp_consumables_response.result
                #Name
                $ConsumableName = $comp_consumables_response.result.display_name
                $consumable_sys_id = $comp_consumables_response.result.sys_id
                $consumable_quantity = $comp_consumables_response.result.quantity
            #Parent Table Results
            #comp_parent_response.result
                If ($Parent -match '/w') {
                    #Parent Display Name
                    $ParentName = $comp_parent_response.result.display_name
                    #Parent Model Category
                    $ParentModelCategory = $comp_parent_response.result.model_category
                    if ($comp_parent_response.result.model_category.value -ne $null) {
                        $parent_model_category_id = $comp_parent_response.result.model_category.value
                        $uri = $comp_parent_response.result.model_category.link
                        $ParentModelCategory_query = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                        $ParentModelCategory = $ParentModelCategory_query.result.name
                    }
                    #Parent Model
                    $ParentModel = $comp_parent_response.result.model
                    if ($comp_parent_response.result.model.value -ne $null) {
                        $parent_model_id = $comp_parent_response.result.model.value
                        $uri = $comp_parent_response.result.model.link
                        $ParentModel_query = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                        $ParentModel = $ParentModel_query.result.name
                    }
                    #Parent Asset Tag
                    $ParentAssetTag = $comp_parent_response.result.asset_tag
                    #Parent State
                    $ParentInstallStatus = $comp_parent_response.result.install_status
                    $parent_install_status = $($StateOptions.result | Where-Object {$_.value -eq $ParentInstallStatus}).label
                    #Parent SubState
                    $ParentSubState = $comp_parent_response.result.supported_by
                    #Parent Class
                    $ParentClassName = $comp_parent_response.result.sys_class_name
                    #Parent Last Inventory By
                    $ParentUpdatedBy = $comp_parent_response.result.u_last_inventory_by
                    if ($comp_parent_response.result.u_last_inventory_by.value -ne $null) {
                        $parent_updatedby_id = $comp_parent_response.result.u_last_inventory_by.value
                        $uri = $comp_parent_response.result.u_last_inventory_by.link
                        $ParentUpdatedBy_query = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                        $ParentUpdatedBy = $ParentUpdatedBy_query.result.name
                    }
                    #Parent Responsible Department
                    $ParentManagedBy = $comp_parent_response.result.managed_by
                    if ($comp_parent_response.result.managed_by.value -ne $null) {
                        $parent_managedby_id = $comp_parent_response.result.managed_by.value
                        $uri = $comp_parent_response.result.managed_by.link
                        $ParentManagedBy_query = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                        $ParentManagedBy = $ParentManagedBy_query.result.name
                        $ParentManagedByID = $parent_managedby_id
                    }
                    #Parent sys_id
                    $ParentSysId = $comp_parent_response.result.sys_id
                    #Parent Department
                    $ParentDepartment = $comp_parent_response.result.department
                    if ($comp_parent_response.result.department.value -ne $null) {
                        $parent_department_id = $comp_parent_response.result.department.value
                        $uri = $comp_parent_response.result.department.link
                        $ParentDepartment_query = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                        $ParentDepartment = $ParentDepartment_query.result.name
                    }
                    #Parent Assigned To
                    $ParentAssignedTo = $comp_parent_response.result.assigned_to
                    #Parent Comments
                    $ParentComments = $comp_parent_response.result.comments
                    #Parent Configuration Item
                    $ParentCI = $comp_parent_response.result.ci
                    #Parent Serial Number
                    $ParentSerialNumber = $comp_parent_response.result.serial_number
                    #Parent Last Physical Inventory
                    $ParentInventoryDate = $comp_parent_response.result.u_last_physical_inventory
                    #Parent Location
                    $ParentLocation = $comp_parent_response.result.location
                    if ($comp_parent_response.result.location.value -ne $null) {
                        $parent_location_id = $comp_parent_response.result.location.value
                        $uri = $comp_parent_response.result.location.link
                        $ParentLocation_query = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                        $ParentLocation = $ParentLocation_query.result.name
                    }
                    #Parent Consumables
                    $uri = "https://$instance_name/api/now/table/alm_consumable?sysparm_query=parent%3D$ParentSysId"
                    $parent_assets_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                    $ParentAssets = $parent_assets_response.result.display_name
                    $ParentAssetsID = $parent_assets_response.result.sys_id
                    $ParentAssetsQuantity = $parent_assets_response.result.quantity
                }
                
                #Create Hash with Relevant Information
                $entry[$i] = @{
                    name = $Name
                    display_name = $DisplayName
                    ci = @{
                        name = $ConfigurationItem
                        id = $ci_sys_id
                        link = $comp_ci_link
                    }
                    serial_number = $SNAM_Serial
                    asset_tag = $AssetTag
                    model = @{
                        model = $Model
                        category = $ModelCategory
                        machine_model_number = $MachineModel
                    }
                    ip = $IPAddress
                    install_status = $InstallStatus
                    sub_state = $SubState
                    function = $Function
                    stock_room = $StockRoom
                    assignment = @{
                        assigned_to = $AssignedTo
                        assigned_to_sys_id = $assignedto_id
                        responsible_department = $Management
                        responsible_department_sys_id = $management_id
                        department = $Department
                        department_sys_id = $department_id
                        cost_center = $CostCenter
                    }
                    po = $PO
                    location = $Location
                    location_sys_id = $location_id
                    warranty = $WarrantyDate
                    assigned_date = $AssignedDate
                    installed_date = $InstallDate
                    inventory_date = $InventoryDate
                    inventory_user = $UpdatedBy
                    comments = $Comments
                    discovery = @{
                        date = $DiscoveryDate
                        source = $DiscoverySource
                    }
                    consumables = @{
                        name = $ConsumableName
                        sys_id = $consumable_sys_id
                        quantity = $consumable_quantity
                    }
                    sys_id = $hw_sys_id
                    asset_id = $asset_sys_id
                    parent = @{
                        name = $ParentName
                        ci = $ParentCI
                        serial_number = $ParentSerialNumber
                        asset_tag = $ParentAssetTag
                        model = @{
                            name = $ParentModel
                            category = $ParentModelCategory
                        }
                        class_name = $ParentClassName
                        state = $ParentInstallStatus
                        sub_state = $ParentSubState
                        assignment = @{
                            responsible_department = $ParentManagedBy
                            department = $ParentDepartment
                            assigned_to = $ParentAssignedTo
                            department_id = $ParentManagedByID
                        }
                        location = $ParentLocation
                        inventory_date = $ParentInventoryDate
                        inventory_user = $ParentUpdatedBy
                        comments = $ParentComments
                        assets = @{
                            name = $ParentAssets
                            sys_id = $ParentAssetsID
                            quantity = $ParentAssetsQuantity
                        }
                        sys_id = $ParentSysId
                    }
                }

            Clear-Variable -Name Name,DisplayName,ConfigurationItem,ci_sys_id,comp_ci_link,snam_serial,assettag,model,modelcategory,machinemodel,ipaddress,installstatus,substate,function,stockroom,assignedto,management,department,costcenter,po,location,warrantydate,assigneddate,installdate,inventorydate,updatedby,comments,discoverydate,discoverysource,consumable,parent* -ErrorAction SilentlyContinue
            }
        
            $other_entry = New-Object System.Collections.ArrayList
            $recent_entry = @{}
            $discovery_dates = @()
            $inventory_dates = @()
            $number = @()
            $index = @()

            for ($i = 1; $i -le $total_results; $i++) {
                If ($entry[$i].ci.id -eq $null) {
                    $entry.Remove($i)
                }
            }

            $old_total = $total_results
            $total_results = $entry.Count

            #Filter Matching Serial Number
            If ($total_results -gt 1) {
                for ($i = 1; $i -le $total_results; $i++) {
                    If ($entry[$i].serial_number -ne $SerialNumber -or $($entry[$i].discovery.date) -notmatch '\w') {
                        $other_entry += $entry[$i]
                        $index += $i
                    }else{
                        $discovery_dates += [DateTime]$($entry[$i].discovery.date)
                        $inventory_dates += [DateTime]$($entry[$i].inventory_date)
                        $number += $i
                    }
                }
                If ($discovery_dates.Count -eq 0) {
                    for ($i = 0; $i -le $total_results; $i++) {
                        $discovery_dates += [DateTime]$($other_entry[$i].inventory_date)
                    }
                }
                #Make a List of Discovery Dates
                $discovery_date_list = @{
                    number = $number
                    date = $discovery_dates
                    inventory = $inventory_dates
                }
                #Filter Most Recent Discovery Date
                $most_recent_date = $discovery_date_list | Foreach-Object {$_.date | Sort-Object {$_.date} | Select-Object -Last 1}
                $most_recent_inventory = $discovery_date_list | ForEach-Object {$_.inventory | Sort-Object {$_.inventory} | Select-Object -Last 1}
        
                for ($i = 1; $i -le $total_results; $i++) {
                    If ([DateTime]$entry[$i].discovery.date -eq $most_recent_date) {
                        $recent_discover = $entry[$i]
                    }
                    If ([DateTime]$entry[$i].inventory_date -eq $most_recent_inventory) {
                        $recent_inventory = $entry[$i]
                    }
                }
                If ($recent_discover -eq $recent_inventory) {
                    $recent_entry = $recent_discover
                }else{
                    $method = 'patch'
                    $uri = "https://$instance_name/api/now/table/alm_hardware/$($recent_inventory.sys_id)"
                    $body = "{`"ci`":`"$recent_discover.ci.id`"}"
                    $change_ci = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body
                    $recent_entry = $recent_inventory
                    $method = 'get'
                }
                If ($recent_entry -notmatch '\w') {
                    $recent_entry = $entry[0]
                }
                for ($i = 1; $i -le $total_results; $i++) {
                    If ($entry[$i].serial_number -eq $recent_entry.serial_number -and $entry[$i].sys_id -ne $recent_entry.sys_id) {
                        If ($recent_entry.po -notmatch '\w' -and $entry[$i].po -match '\w') {
                            $recent_entry.po = $entry[$i].po
                        }
                        #Add Duplicate Entry to Errors Log
                        $SNAM_Error_Path = "$PSScriptRoot\Files\SNAM Errors.csv"
                        $error.Clear()
                        $UserName = $((Get-WmiObject -Class win32_computersystem).UserName).Split("\")[1]
                        If ($error) {
                            $UserName = "sysklindstrom"
                        }
                        New-Object -TypeName PSCustomObject -Property @{
                            "User" = $UserName
                            "Computer ID" = $recent_entry.sys_id
                            "Parent ID" = $null
                            "Serial Number" = $recent_entry.serial_number
                            "Retired Serial Numbers" = $null
                            "Serials with Wrong Names" = $null
                            "Duplicate Serial Numbers" = $null
                            "Comments" = "This Serial Number contains a duplicate record"
                        } | Select-Object "User","Computer ID","Parent ID","Serial Number","Retired Serial Numbers","Serials with Wrong Names","Duplicate Serial Numbers","Comments" | Export-Csv -Path $SNAM_Error_Path -NoTypeInformation -Append -Encoding ASCII
                        #Warning Box
                        $answer = [System.Windows.Forms.MessageBox]::Show("$SerialNumber has a duplicate record. Would you like to mark it for deletion?", "Duplicate Record", 4)
                        If ($answer -eq "Yes") {
                            $method = 'patch'
                            $uri = "https://$instance_name/api/now/table/alm_hardware/$($entry[$i].sys_id)"
                            $on_order = $($StateOptions.result | Where-Object {$_.label -eq 'On order'}).value
                            $body = "{`"serial_number`":`"DELETE`",`"install_status`":`"$on_order`",`"managed_by`":`"$IT_department_id`"}"
                            $remove_entry = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body
                            $uri = $entry[$i].ci.link
                            $body = "{`"serial_number`":`"DELETE`"}"
                            $remove_ci_entry = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body
                            $method = 'get'
                        }
                    }
                }
            }else{
                for ($i = 1; $i -le $old_total; $i++) {
                    If ($entry[$i] -ne $null) {
                        $recent_entry = $entry[$i]
                    }
                }
            }
        
            If ($recent_entry -notmatch '\w') {
                Add-Type -AssemblyName System.Windows.Forms
                Add-Type -AssemblyName System.Drawing

                $create_form = New-Object System.Windows.Forms.Form
                $create_form.Text = "No Results"
                $create_form.Size = New-Object System.Drawing.Size(377,200)
                $create_form.StartPosition = 'CenterScreen'
            
                $error_label = New-Object System.Windows.Forms.Label
                $error_label.Location = New-Object System.Drawing.Point(5,5)
                $error_label.AutoSize = $true
                $error_label.Text = "No Records found for Serial Number: $SerialNumber"
                $create_form.Controls.Add($error_label)

                $verify_label = New-Object System.Windows.Forms.Label
                $verify_label.Location = New-Object System.Drawing.Point(5,35)
                $verify_label.AutoSize = $true
                $verify_label.Text = "Please verify that the Computer Name and Serial Number are correct."
                $create_form.Controls.Add($verify_label)

                $Name_label = New-Object System.Windows.Forms.Label
                $Name_label.Location = New-Object System.Drawing.Point(5,65)
                $Name_label.AutoSize = $true
                $Name_label.Text = "Computer Name:"
                $create_form.Controls.Add($Name_label)

                $Serial_label = New-Object System.Windows.Forms.Label
                $Serial_label.Location = New-Object System.Drawing.Point(14,95)
                $Serial_label.AutoSize = $true
                $Serial_label.Text = "Serial Number:"
                $create_form.Controls.Add($Serial_label)

                $verify_Name_box = New-Object System.Windows.Forms.TextBox
                $verify_Name_box.Location = New-Object System.Drawing.Point($($Name_label.Location.X + $Name_label.Size.Width),$($Name_label.Location.Y - 1))
                $verify_Name_box.Size = New-Object System.Drawing.Size(255,25)
                $verify_Name_box.Text = $ComputerName
                $create_form.Controls.Add($verify_Name_box)

                $verify_Serial_box = New-Object System.Windows.Forms.TextBox
                $verify_Serial_box.Location = New-Object System.Drawing.Point($($Name_label.Location.X + $Name_label.Size.Width),$($Serial_label.Location.Y - 1))
                $verify_Serial_box.Size = New-Object System.Drawing.Size(255,25)
                $verify_Serial_box.Text = $SerialNumber
                $create_form.Controls.Add($verify_Serial_box)

                $OKButton = New-Object System.Windows.Forms.Button
                $OKButton.Location = New-Object System.Drawing.Point(256,125)
                $OKButton.Size = New-Object System.Drawing.Size(95,23)
                $OKButton.Text = "Search"
                $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
                $create_form.Controls.Add($OKButton)

                $NewButton = New-Object System.Windows.Forms.Button
                $NewButton.Location = New-Object System.Drawing.Point(158,125)
                $NewButton.Size = New-Object System.Drawing.Size(95,23)
                $NewButton.Text = "Create Record"
            
                $create_record = {
                    $global:ComputerName = $verify_Name_box.Text
                    $global:SerialNumber = $verify_Serial_box.Text
                    #Create Model Search Box
                    If (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet) {
                        $model = (Get-WmiObject -Class Win32_ComputerSystemProduct -ComputerName $ComputerName).Version
                    }else{
                    
                    }
                    $search = $model
                    $uri = "https://$instance_name/api/now/table/cmdb_model?sysparm_query=nameLIKE$search"
                    $search_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                    $model_list = @()
                    foreach ($search_result in $($search_response.result)) {
                        $row = "" | Select Name,SysID
                        $row.Name = $search_result.name
                        $row.SysID = $search_result.sys_id
                        $model_list += $row
                    }
                    $global:model_list = $model_list | Sort-Object -Property Name

                    Add-Type -AssemblyName System.Windows.Forms
                    Add-Type -AssemblyName System.Drawing

                    $search_form = New-Object System.Windows.Forms.Form
                    $search_form.Text = "Model Search:"
                    $search_form.AutoSize = $true
                    $search_form.StartPosition = 'CenterScreen'

                    $search_label = New-Object System.Windows.Forms.Label
                    $search_label.Text = "Search: (Model: $($model))"
                    $search_label.Location = New-Object System.Drawing.Point(5,8)
                    $search_label.Size = New-Object System.Drawing.Size(400,20)
                    $search_form.Controls.Add($search_label)

                    $results_label = New-Object System.Windows.Forms.Label
                    $results_label.Text = "Results:"
                    $results_label.Location = New-Object System.Drawing.Point(5,70)
                    $results_label.Size = New-Object System.Drawing.Size(50,21)
                    $search_form.Controls.Add($results_label)

                    $selectionBox = New-Object System.Windows.Forms.ListBox
                    $selectionBox.Location = New-Object System.Drawing.Point(5,91)
                    $selectionBox.AutoSize = $true
                    $selectionBox.MinimumSize = New-Object System.Drawing.Size(417,200)
                    $selectionBox.MaximumSize = New-Object System.Drawing.Size(0,200)
                    $selectionBox.ScrollAlwaysVisible = $true
                    $selectionBox.Items.Clear()
                    If ($model_list.name -notmatch '\w') {
                        $selectionBox.Items.Add("No Results Found")
                    }else{
                        foreach ($item in $model_list) {
                            [void] $selectionBox.Items.Add($item.Name)
                        }
                    }
                    $search_form.TopMost = $true
                    $search_form.Controls.Add($selectionBox)

                    $search_bar = New-Object System.Windows.Forms.TextBox
                    $search_bar.Location = New-Object System.Drawing.Size(5,28)
                    $search_bar.Size = New-Object System.Drawing.Size(400,21)
                    $search_form.Controls.Add($search_bar)
                    $search_bar.Text = $model

                    $search_button = New-Object System.Windows.Forms.Button
                    $search_button.Location = New-Object System.Drawing.Point(408,27)
                    $search_button.Size = New-Object System.Drawing.Size(21,21)
                    $search_button.BackgroundImage = $image
                    $search_button.BackgroundImageLayout = 'Zoom'

                    $search_trigger = {
                        $search = $search_bar.Text
                        $selectionBox.Items.Clear()
                        $selectionBox.Items.Add("Processing . . .")
                        $uri = "https://$instance_name/api/now/table/cmdb_model?sysparm_query=nameLIKE$search"
                        $search_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                        $model_list = @()
                        foreach ($search_result in $($search_response.result)) {
                            $row = "" | Select Name,Email,Department,SysID
                            $row.Name = $search_result.name
                            $row.SysID = $search_result.sys_id
                            $model_list += $row
                        }
                        $global:model_list = $model_list | Sort-Object -Property Name
                        $selectionBox.Items.Clear()
                        If ($model_list.name -notmatch '\w') {
                            $selectionBox.Items.Add("No Results Found")
                        }else{
                            foreach ($item in $model_list) {
                                [void] $selectionBox.Items.Add($item.Name)
                            }
                        }
                    }

                    $search_button.add_click($search_trigger)
                    $search_form.Controls.Add($search_button)

                    $search_ok_button = New-Object System.Windows.Forms.Button
                    $search_ok_button.Location = New-Object System.Drawing.Point(347, $($($selectionBox.Location.Y) + $($selectionBox.Size.Height) + 10))
                    $search_ok_button.Size = New-Object System.Drawing.Size(75,23)
                    $search_ok_button.Text = 'Confirm'
                    $search_ok_button.DialogResult = [System.Windows.Forms.DialogResult]::OK
                    $search_form.Controls.Add($search_ok_button)

                    $search_accept = {
                        $search_form.AcceptButton = $search_button
                        $selectionBox.SelectedItem = $null
                    }
                    $form_accept = {
                        $search_form.AcceptButton = $search_ok_button
                    }

                    $search_bar.add_MouseDown($search_accept)
                    $selectionBox.add_MouseDown($form_accept)

                    $search_confirm = $search_form.ShowDialog()

                    If ($search_confirm -eq [System.Windows.Forms.DialogResult]::OK) {
                        if ($model_list.Count -gt 1) {
                            $global:model_id = $model_list.SysID[$($selectionBox.SelectedIndex)]
                        }else{
                            $global:model_id = $model_list.SysID
                        }

                        $uri = "https://$instance_name/api/now/table/cmdb_model_category?sysparm_query=name%3DPC Hardware"
                        $model_category_id = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                        $model_category_id = $model_category_id.result.sys_id

                        $install_status = $($StateOptions.result | Where-Object {$_.label -eq "In stock"}).value

                        $uri = "https://$instance_name/api/now/table/alm_stockroom?sysparm_query=name%3DDesktop - Raleigh - Glass Room"
                        $stockroom_id = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                        $stockroom_id = $stockroom_id.result.sys_id

                        $uri = "https://$instance_name/api/now/table/cmn_department?sysparm_query=name%3DInformation Services"
                        $department_id = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                        $department_id = $department_id.result.sys_id
                        If ($department_id.Count -gt 1) {
                            $department_id = $department_id[0]
                        }

                        $method = "post"
                        $uri = "https://$instance_name/api/now/table/alm_hardware"
                        $body = "{`"serial_number`":`"$SerialNumber`",`"model_category`":`"$model_category_id`",`"model`":`"$model_id`",`"install_status`":`"$install_status`",`"stockroom`":`"$stockroom_id`",`"managed_by`":`"$department_id`"}"
                        $create_record_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body
                        $method = "get"
                        $create_form.Close()
                        Add-Type -AssemblyName System.Windows.Forms
                        Add-Type -AssemblyName System.Drawing

                        $font = New-Object System.Drawing.Font("Microsoft Sans Serif",10,[System.Drawing.FontStyle]::Regular)

                        $created_form = New-Object System.Windows.Forms.Form
                        $created_form.Text = "$SerialNumber"
                        $created_form.Size = New-Object System.Drawing.Size(250,100)
                        $created_form.StartPosition = 'CenterScreen'
            
                        $created_label = New-Object System.Windows.Forms.Label
                        $created_label.Location = New-Object System.Drawing.Point(69,8)
                        $created_label.AutoSize = $true
                        $created_label.Font = $font
                        $created_label.Text = "Record Created!"
                        $created_form.Controls.Add($created_label)

                        $OKButton = New-Object System.Windows.Forms.Button
                        $OKButton.Location = New-Object System.Drawing.Point(80,35)
                        $OKButton.Size = New-Object System.Drawing.Point(75,23)
                        $OKButton.Text = 'OK'
                        $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
                        $created_form.Controls.Add($OKButton)
                        $created_form.TopMost = $true
                        $created_form.ShowDialog()
                    }
                }

                $NewButton.add_click($create_record)
                $create_form.Controls.Add($NewButton)
                $create_form.TopMost = $true
                $create_results = $create_form.ShowDialog()

                If ($create_results -eq [System.Windows.Forms.DialogResult]::OK) {
                    $ComputerName = $verify_Name_box.Text
                    $SerialNumber = $verify_Serial_box.Text
                    $continue = "Yes"
                }

            }else{
                $continue = "No"
            }
        }

    #Check for Consumables on Child and Parent Item
        If ($recent_entry.parent.sys_id -match '\w') {
            If ($recent_entry.consumables.sys_id -match '\w') {
                If ($recent_entry.parent.assets.sys_id -notmatch '\w') {
                    $recent_entry.parent.assets = $recent_entry.consumables
                    #Move All Consumables from Child to Parent
                    $method = "post"
                    $uri = "https://$instance_name/api/wmh/consumables/parentSwap"
                    $new_parent_id = $($recent_entry.parent.sys_id)
                    $old_parent_id = $($recent_entry.sys_id)
                    $body = "{`"newParent`":`"$new_parent_id`",`"oldParent`":`"$old_parent_id`",`"newUser`":`"`"}"
                    $move_to_parent_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body
                    $method = "get"
                }else{
                    #Return All Consumables to Stock
                    $method = "post"
                    $cons_stock_id = $recent_entry.sys_id
                    $uri = "https://$instance_name/api/wmh/consumables/returnAllToStock"
                    $body = "{`"parent`":`"$cons_stock_id`",`"stockroom`":`"$stockroom_id`"}"
                    $consumables_stock_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body
                    $method = "get"
                    $recent_entry.consumables = $null
                }
            }
        }

    #Create Most Recent Discovery Display Box
        $search_form = $null

        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing

        $recent_form = New-Object System.Windows.Forms.Form
        $recent_form.Text = "$($recent_entry.name) SNAM Results"
        $recent_form.AutoSize = $true
        $recent_form.StartPosition = 'CenterScreen'

        $font_bold = New-Object System.Drawing.Font("Microsoft Sans Serif",8.25,[System.Drawing.FontStyle]::Bold)
        $image = [System.Drawing.Image]::FromFile("$PSScriptRoot\Files\search_icon.jpeg")

        #Header Labels
        $infolabel1 = New-Object System.Windows.Forms.Label
        $infolabel1.Location = New-Object System.Drawing.Point(0,0)
        $infolabel1.AutoSize = $true
        $infolabel1.Text = "
 Computer Name:

     Serial Number:

                   Model:

       Install Status:

           IP Address:

            Comments:"
        $recent_form.Controls.Add($infolabel1)

        #ComputerName Box
        $ComputerName_box = New-Object System.Windows.Forms.TextBox
        $ComputerName_box.Location = New-Object System.Drawing.Point($($($infolabel1.Size.Width)),10)
        $ComputerName_box.Size = New-Object System.Drawing.Size(225,21)
        $ComputerName_box.Text = $($recent_entry.name)
        $ComputerName_box.Font = $font_bold
        $recent_form.Controls.Add($ComputerName_box)

        If (-not($ComputerName)) {
            $global:ComputerName = $($recent_entry.name)
            $ComputerName = $global:ComputerName
        }

        #Header Information
        $infolabel2 = New-Object System.Windows.Forms.Label
        $infolabel2.Location = New-Object System.Drawing.Point($($($infolabel1.Size.Width)),0)
        $infolabel2.AutoSize = $true
        $infolabel2.MaximumSize = New-Object System.Drawing.Point(255,0)
        $infolabel2.Font = $font_bold
        $infolabel2.Text = "


$($recent_entry.serial_number)

$($recent_entry.model.model)

$install_status

$($recent_entry.ip)

$($recent_entry.comments)"
        $recent_form.Controls.Add($infolabel2)

        #Header Labels Column 2
        $infolabel3 = New-Object System.Windows.Forms.Label
        $infolabel3.Location = New-Object System.Drawing.Point (355,0)
        $infolabel3.AutoSize = $true
        $infolabel3.Text = "
          Warranty Date:

          Assigned Date:

           Installed Date:

    Last Inventory By:

         Inventory Date:

Last Discovery Date:

                             PO:"
        $recent_form.Controls.Add($infolabel3)

        #PO Box
        $PO_box = New-Object System.Windows.Forms.TextBox
        $PO_box.Location = New-Object System.Drawing.Point($($($infolabel3.Location.X) + $($infolabel3.Size.Width)),156)
        $PO_box.Size = New-Object System.Drawing.Size(145,21)
        $PO_box.Text = $($recent_entry.po)
        $PO_box.Font = $font_bold
        $recent_form.Controls.Add($PO_box)

        If (-not($ComputerName)) {
            $global:ComputerName = $($recent_entry.name)
            $ComputerName = $global:ComputerName
        }

        #Header Info Column 2
        $infolabel4 = New-Object System.Windows.Forms.Label
        $infolabel4.Location = New-Object System.Drawing.Point($($($infolabel3.Location.X) + $($infolabel3.Size.Width)),0)
        $infolabel4.AutoSize = $true
        $infolabel4.Font = $font_bold
        $infolabel4.Text = "
$($recent_entry.warranty)

$($recent_entry.assigned_date)

$($recent_entry.installed_date)

$($recent_entry.inventory_user)

$($recent_entry.inventory_date)

$($recent_entry.discovery.date)"

        $recent_form.Controls.Add($infolabel4)

        #If ($($recent_entry.comments) -match '\w') {
        If ($infolabel2.Size.Height -gt $infolabel3.Size.Height) {
            $functionlabel_y = $($infolabel2.Size.Height)
        }else{
            $functionlabel_y = $($infolabel3.Size.Height)
        }

        #Function Label
        $functionlabel = New-Object System.Windows.Forms.Label
        $functionlabel.Location = New-Object System.Drawing.Point(0,$($functionlabel_y + 20))
        $functionlabel.Autosize = $true
        $functionlabel.Text = "                          Function:"
        $recent_form.Controls.Add($functionlabel)

        #Function Item
        $functionitem = New-Object System.Windows.Forms.Label
        $functionitem.Location = New-Object System.Drawing.Point(137,$($functionlabel.Location.Y))
        $functionitem.Autosize = $true
        $functionitem.MaximumSize = New-Object System.Drawing.Point(188,0)
        $functionitem.Text = "$($recent_entry.function)"
        $functionitem.Font = $font_bold
        $recent_form.Controls.Add($functionitem)

        #Assigned To Label
        $assigned_to_label = New-Object System.Windows.Forms.Label
        $assigned_to_label.Location = New-Object System.Drawing.Point(0,$($($functionlabel.Location.Y) + $($functionitem.Size.Height) + 8))
        $assigned_to_label.Autosize = $true
        $assigned_to_label.Text = "                    Assigned To:"
        $recent_form.Controls.Add($assigned_to_label)

        #Assigned To Item
        $assigned_to_item = New-Object System.Windows.Forms.Label
        $assigned_to_item.Location = New-Object System.Drawing.Point(137,$($assigned_to_label.Location.Y))
        $assigned_to_item.Autosize = $true
        $assigned_to_item.MaximumSize = New-Object System.Drawing.Point(188,0)
        $assigned_to_item.Text = "$($recent_entry.assignment.assigned_to)"
        $assigned_to_item.Font = $font_bold
        $recent_form.Controls.Add($assigned_to_item)

        #Responsible Department Label
        $responsible_department_label = New-Object System.Windows.Forms.Label
        $responsible_department_label.Location = New-Object System.Drawing.Point(0,$($($assigned_to_label.Location.Y) + $($assigned_to_item.Size.Height) + 8))
        $responsible_department_label.Autosize = $true
        $responsible_department_label.Text = "Responsible Department:"
        $recent_form.Controls.Add($responsible_department_label)

        #Responsible Department Item
        $responsible_department_item = New-Object System.Windows.Forms.Label
        $responsible_department_item.Location = New-Object System.Drawing.Point(137,$($responsible_department_label.Location.Y))
        $responsible_department_item.Autosize = $true
        $responsible_department_item.MaximumSize = New-Object System.Drawing.Point(188,0)
        $responsible_department_item.Text = "$($recent_entry.assignment.responsible_department)"
        $responsible_department_item.Font = $font_bold
        If ($($recent_entry.assignment.responsible_department) -match '\*IA') {
            $responsible_department_item.ForeColor = 'RED'
        }
        $recent_form.Controls.Add($responsible_department_item)

        #Department Label
        $departmentlabel = New-Object System.Windows.Forms.Label
        $departmentlabel.Location = New-Object System.Drawing.Point(0,$($($responsible_department_label.Location.Y) + $($responsible_department_item.Size.Height) + 8))
        $departmentlabel.Autosize = $true
        $departmentlabel.Text = "                      Department:"
        $recent_form.Controls.Add($departmentlabel)

        #Department Item
        $department_item = New-Object System.Windows.Forms.Label
        $department_item.Location = New-Object System.Drawing.Point(137,$($departmentlabel.Location.Y))
        $department_item.Autosize = $true
        $department_item.MaximumSize = New-Object System.Drawing.Point(188,0)
        $department_item.Text = "$($recent_entry.assignment.department)"
        $department_item.Font = $font_bold
        If ($($recent_entry.assignment.department) -match '\*IA') {
            $department_item.ForeColor = 'RED'
        }
        $recent_form.Controls.Add($department_item)

        #Parent Serial Label
        $parentlabel = New-Object System.Windows.Forms.Label
        $parentlabel.Location = New-Object System.Drawing.Point(0,$($($departmentlabel.Location.Y) + $($department_item.Size.Height) + 8))
        $parentlabel.Autosize = $true
        $parentlabel.Text = "      Parent Serial Number:"
        $recent_form.Controls.Add($parentlabel)

        #Parent Serial Number Item
        $parent_item = New-Object System.Windows.Forms.Label
        $parent_item.Location = New-Object System.Drawing.Point(137,$($parentlabel.Location.Y))
        $parent_item.Autosize = $true
        $parent_item.MaximumSize = New-Object System.Drawing.Point(188,0)
        $parent_item.Text = "$($recent_entry.parent.serial_number)"
        $parent_item.Font = $font_bold
        $recent_form.Controls.Add($parent_item)

        #Location Label
        $locationlabel = New-Object System.Windows.Forms.Label
        $locationlabel.Location = New-Object System.Drawing.Point(0,$($($parentlabel.Location.Y) + $($parent_item.Size.Height) + 8))
        $locationlabel.Autosize = $true
        $locationlabel.Text = "                           Location:"
        $recent_form.Controls.Add($locationlabel)

        #Location Item
        $location_item = New-Object System.Windows.Forms.Label
        $location_item.Location = New-Object System.Drawing.Point(137,$($locationlabel.Location.Y))
        $location_item.Autosize = $true
        $location_item.MaximumSize = New-Object System.Drawing.Point(188,0)
        $location_item.Text = "$($recent_entry.location)"
        $location_item.Font = $font_bold
        $recent_form.Controls.Add($location_item)

        #Consumables Label
        $consumableslabel = New-Object System.Windows.Forms.Label
        $consumableslabel.Location = New-Object System.Drawing.Point(0,$($($locationlabel.Location.Y) + $($location_item.Size.Height) + 8))
        $consumableslabel.Autosize = $true
        $consumableslabel.Text = "                  Consumables:"
        $recent_form.Controls.Add($consumableslabel)

        #Consumables Item
        $consumables_item = New-Object System.Windows.Forms.Label
        $consumables_item.Location = New-Object System.Drawing.Point(0,$($($consumableslabel.Location.Y) + $($consumableslabel.Size.Height)))
        $consumables_item.Autosize = $true
        $consumables_item.MaximumSize = New-Object System.Drawing.Point(325,0)

        $consumables_item_text = @()
        If ($($recent_entry.parent.serial_number) -match '\w') {
            $record_consumables = $recent_entry.parent.assets
            for ($i = 1; $i -le $($record_consumables.name.count); $i++) {
                for ($j = 1; $j -le $record_consumables.quantity[$i-1]; $j++) {
                    $consumables_item_text += $record_consumables.name[$i-1]
                }
            }
        }
        $record_consumables = $recent_entry.consumables
        for ($i = 1; $i -le $($record_consumables.name.count); $i++) {
            for ($j = 1; $j -le $record_consumables.quantity[$i-1]; $j++) {
                $consumables_item_text += $record_consumables.name[$i-1]
            }
        }

        $consumables_item.Text = $($($consumables_item_text | Out-String).Trim())
        $consumables_item.Font = $font_bold
        $recent_form.Controls.Add($consumables_item)

        #Function Drop-Down Menu
        $functionBox = New-Object System.Windows.Forms.ComboBox
        $functionBox.Location = New-Object System.Drawing.Size(415,$($($functionlabel.Location.Y) - 2))
        $functionBox.Size = New-Object System.Drawing.Size(200,21)
        $functionBox.DropDownHeight = 100

        $uri = "https://$instance_name/api/now/table/sys_choice?sysparm_query=name%3Dalm_hardware^element%3Du_device_function"
        $FunctionOptions = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
        $range = $FunctionOptions.result.label

        $functionBox.Items.AddRange(@($range))

        #Assigned To Search Bar
        $assigned_to_box = New-Object System.Windows.Forms.TextBox
        $assigned_to_box.Location = New-Object System.Drawing.Size(415,$($($assigned_to_label.Location.Y) - 2))
        $assigned_to_box.Size = New-Object System.Drawing.Size(200,21)
        $recent_form.Controls.Add($assigned_to_box)
        <#
        $assigned_to_click = {
            If ($assigned_to_box.ReadOnly -eq $true) {
                $assigned_to_box.Text = $null
            }else{
                Clear-Variable -Name currentUserName -ErrorAction SilentlyContinue
                Clear-Variable -Name lastUser -ErrorAction SilentlyContinue
                Clear-Variable -Name currentUserFirst -ErrorAction SilentlyContinue
                Clear-Variable -Name currentUserLast -ErrorAction SilentlyContinue
                Clear-Variable -Name currentUser -ErrorAction SilentlyContinue
                $currentUserName = (Get-WmiObject -Class win32_computersystem -ComputerName $ComputerName -ErrorAction SilentlyContinue).UserName
                If ($currentUserName -eq $null) {
                    $currentUserName = (Get-WmiObject -Class win32_process -ComputerName $ComputerName -ErrorAction SilentlyContinue | Where-Object name -Match explorer -ErrorAction SilentlyContinue)
                    If ($currentUserName -ne $null) {
                        $currentUserName = $currentUserName.getowner().user
                    }
                }else{
                    $currentUserName = $currentUserName.Split("\")[1]
                }
                If ($currentUserName -ne $null) {
                    $currentUserFirst = (Get-ADUser $currentUserName).GivenName
                    $currentUserLast = (Get-ADUser $currentUserName).Surname
                    If ($currentUserLast -match "_") {
                        $currentUserLast = $currentUserLast.Split("_")[0]
                    }
                    $currentUser = "$currentUserFirst $currentUserLast"
                }
                If ($currentUser -ne $null) {
                    $assigned_to_box.Text = "$currentUser"
                }
            }
        }
        $assigned_to_box.add_click($assigned_to_click)
        #>
        #Responsible Department Search Bar
        $responsible_department_box = New-Object System.Windows.Forms.TextBox
        $responsible_department_box.Location = New-Object System.Drawing.Size(415,$($($responsible_department_label.Location.Y) - 2))
        $responsible_department_box.Size = New-Object System.Drawing.Size(200,21)
        $recent_form.Controls.Add($responsible_department_box)

        #Department Search Bar
        $department_box = New-Object System.Windows.Forms.TextBox
        $department_box.Location = New-Object System.Drawing.Size(415,$($($departmentlabel.Location.Y) - 2))
        $department_box.Size = New-Object System.Drawing.Size(200,21)
        $department_box.ReadOnly = $true
        $recent_form.Controls.Add($department_box)

        #Parent Serial Number Search Bar
        $parent_box = New-Object System.Windows.Forms.TextBox
        $parent_box.Location = New-Object System.Drawing.Size(415,$($($parentlabel.Location.Y) -2))
        $parent_box.Size = New-Object System.Drawing.Size(200,21)
        $recent_form.Controls.Add($parent_box)

        #Change Location Search Bar
        $location_box = New-Object System.Windows.Forms.TextBox
        $location_box.Location = New-Object System.Drawing.Size(415,$($($locationlabel.Location.Y) - 2))
        $location_box.Size = New-Object System.Drawing.Size(200,21)
        $recent_form.Controls.Add($location_box)

        #Change Consumables Search Bar
        $consumables_box = New-Object System.Windows.Forms.TextBox
        $consumables_box.Location = New-Object System.Drawing.Size(415,$($($consumableslabel.Location.Y) - 2))
        $consumables_box.Size = New-Object System.Drawing.Size(200,21)
        $recent_form.Controls.Add($consumables_box)

        #Change Function Label
        $change_function_label = New-Object System.Windows.Forms.Label
        $change_function_label.Location = New-Object System.Drawing.Point(355,$($functionlabel.Location.Y))
        $change_function_label.Autosize = $true
        $change_function_label.Text = "Change?"
        $recent_form.Controls.Add($change_function_label)

        #Change Assigned To Label
        $change_assigned_to_label = New-Object System.Windows.Forms.Label
        $change_assigned_to_label.Location = New-Object System.Drawing.Point(355,$($assigned_to_label.Location.Y))
        $change_assigned_to_label.Autosize = $true
        $change_assigned_to_label.Text = "Change?"
        $recent_form.Controls.Add($change_assigned_to_label)
                        
        #Change Responsible Department Label
        $change_responsible_department_label = New-Object System.Windows.Forms.Label
        $change_responsible_department_label.Location = New-Object System.Drawing.Point(355,$($responsible_department_label.Location.Y))
        $change_responsible_department_label.Autosize = $true
        $change_responsible_department_label.Text = "Change?"
        $recent_form.Controls.Add($change_responsible_department_label)

        #Change Department Label
        $change_department_label = New-Object System.Windows.Forms.Label
        $change_department_label.Location = New-Object System.Drawing.Point(355,$($departmentlabel.Location.Y))
        $change_department_label.Autosize = $true
        $change_department_label.Text = "Change?"
        $recent_form.Controls.Add($change_department_label)

        #Change Parent Serial Number Label
        $change_parent_label = New-Object System.Windows.Forms.Label
        $change_parent_label.Location = New-Object System.Drawing.Point(355,$($parentlabel.Location.Y))
        $change_parent_label.Autosize = $true
        $change_parent_label.Text = "Change?"
        $recent_form.Controls.Add($change_parent_label)

        #Change Location Label
        $change_location_label = New-Object System.Windows.Forms.Label
        $change_location_label.Location = New-Object System.Drawing.Point(355,$($locationlabel.Location.Y))
        $change_location_label.AutoSize = $true
        $change_location_label.Text = "Change?"
        $recent_form.controls.Add($change_location_label)

        #Change Consumables Label
        $change_consumables_label = New-Object System.Windows.Forms.Label
        $change_consumables_label.Location = New-Object System.Drawing.Point(355,$($consumableslabel.Location.Y))
        $change_consumables_label.AutoSize = $true
        $change_consumables_label.Text = "Change?"
        $recent_form.controls.Add($change_consumables_label)

        #Assigned-To Search Button
        $assigned_to_search = New-Object System.Windows.Forms.Button
        $assigned_to_search.Location = New-Object System.Drawing.Point($($($assigned_to_box.Location.X) + $($assigned_to_box.Size.Width) + 3),$($($assigned_to_box.Location.Y) - 1))
        $assigned_to_search.Size = New-Object System.Drawing.Size(21,21)
        $assigned_to_search.BackgroundImage = $image
        $assigned_to_search.BackgroundImageLayout = 'Zoom'

        $assigned_trigger = {
            $search = $assigned_to_box.Text
            $uri = "https://$instance_name/api/now/table/sys_user?sysparm_query=nameLIKE$search"
            $search_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
            $user_list = @()
            foreach ($search_result in $($search_response.result)) {
                If ($($search_result.email) -notmatch '\w') {
                    $email = "Empty"
                }else{
                    $email = $search_result.email
                }
                $row = "" | Select Name,Email,Department,SysID
                $row.Name = $search_result.name
                $row.Email = $email
                $row.Department = $search_result.u_department_number
                $row.SysID = $search_result.sys_id
                $user_list += $row
            }
            $global:user_list = $user_list | Sort-Object -Property Name

            Add-Type -AssemblyName System.Windows.Forms
            Add-Type -AssemblyName System.Drawing

            $search_form = New-Object System.Windows.Forms.Form
            $search_form.Text = "Assigned To Search:"
            $search_form.AutoSize = $true
            $search_form.StartPosition = 'CenterScreen'

            $search_label = New-Object System.Windows.Forms.Label
            $search_label.Text = "Search:"
            $search_label.Location = New-Object System.Drawing.Point(5,8)
            $search_label.Size = New-Object System.Drawing.Size(50,20)
            $search_form.Controls.Add($search_label)

            $results_label = New-Object System.Windows.Forms.Label
            $results_label.Text = "Results:"
            $results_label.Location = New-Object System.Drawing.Point(5,70)
            $results_label.Size = New-Object System.Drawing.Size(50,21)
            $search_form.Controls.Add($results_label)

            $selectionBox = New-Object System.Windows.Forms.ListBox
            $selectionBox.Location = New-Object System.Drawing.Point(5,91)
            $selectionBox.AutoSize = $true
            $selectionBox.MinimumSize = New-Object System.Drawing.Size(417,200)
            $selectionBox.MaximumSize = New-Object System.Drawing.Size(0,200)
            $selectionBox.ScrollAlwaysVisible = $true
            $selectionBox.Items.Clear()
            If ($user_list.name -notmatch '\w') {
                $selectionBox.Items.Add("No Results Found")
            }else{
                foreach ($item in $user_list) {
                    [void] $selectionBox.Items.Add("$($item.name) ($($item.email))")
                }
            }
            $search_form.TopMost = $true
            $search_form.Controls.Add($selectionBox)

            $search_bar = New-Object System.Windows.Forms.TextBox
            $search_bar.Location = New-Object System.Drawing.Size(5,28)
            $search_bar.Size = New-Object System.Drawing.Size(400,21)
            $search_form.Controls.Add($search_bar)
            $search_bar.Text = $assigned_to_box.Text

            $search_button = New-Object System.Windows.Forms.Button
            $search_button.Location = New-Object System.Drawing.Point(408,27)
            $search_button.Size = New-Object System.Drawing.Size(21,21)
            $search_button.BackgroundImage = $image
            $search_button.BackgroundImageLayout = 'Zoom'

            $search_trigger = {
                $search = $search_bar.Text
                $selectionBox.Items.Clear()
                $selectionBox.Items.Add("Processing . . .")
                $uri = "https://$instance_name/api/now/table/sys_user?sysparm_query=nameLIKE$search"
                $search_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                $user_list = @()
                foreach ($search_result in $($search_response.result)) {
                    If ($($search_result.email) -notmatch '\w') {
                        $email = "Empty"
                    }else{
                        $email = $search_result.email
                    }
                    $row = "" | Select Name,Email,Department,SysID
                    $row.Name = $search_result.name
                    $row.Email = $email
                    $row.Department = $search_result.u_department_number
                    $row.SysID = $search_result.sys_id
                    $user_list += $row
                }
                $global:user_list = $user_list | Sort-Object -Property Name
                $selectionBox.Items.Clear()
                If ($user_list.name -notmatch '\w') {
                    $selectionBox.Items.Add("No Results Found")
                }else{
                    foreach ($item in $user_list) {
                        [void] $selectionBox.Items.Add("$($item.name) ($($item.email))")
                    }
                }
            }

            $search_button.add_click($search_trigger)
            $search_form.Controls.Add($search_button)

            $search_ok_button = New-Object System.Windows.Forms.Button
            $search_ok_button.Location = New-Object System.Drawing.Point(347, $($($selectionBox.Location.Y) + $($selectionBox.Size.Height) + 10))
            $search_ok_button.Size = New-Object System.Drawing.Size(75,23)
            $search_ok_button.Text = 'Confirm'
            $search_ok_button.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $search_form.Controls.Add($search_ok_button)

            $search_accept = {
                $search_form.AcceptButton = $search_button
                $selectionBox.SelectedItem = $null
            }
            $form_accept = {
                $search_form.AcceptButton = $search_ok_button
            }

            $search_bar.add_MouseDown($search_accept)
            $selectionBox.add_MouseDown($form_accept)

            $search_confirm = $search_form.ShowDialog()

            if ($search_confirm -eq [System.Windows.Forms.DialogResult]::OK) {
                $assigned_to_box.Text = $selectionBox.SelectedItem
                if ($user_list.Count -gt 1) {
                    $global:assigned_to_selection_id = $user_list.SysID[$($selectionBox.SelectedIndex)]
                    $department_number = $user_list.department[$($selectionBox.SelectedIndex)]
                }else{
                    $global:assigned_to_selection_id = $user_list.SysID
                    $department_number = $user_list.department
                }
                $uri = "https://$instance_name/api/now/table/cmn_department?sysparm_query=id%3D$department_number"
                $department_search_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                $department_box.Text = $department_search_response.result.name
                $responsible_department_box.Text = ""
                $global:department_selection_id = $department_search_response.result.sys_id
                $global:responsible_department_selection_id = $department_search_response.result.sys_id
            }

        }

        $assigned_to_search.add_click($assigned_trigger)
        $recent_form.Controls.Add($assigned_to_search)

        #Responsible Department Search Button
        $responsible_department_search = New-Object System.Windows.Forms.Button
        $responsible_department_search.Location = New-Object System.Drawing.Point($($($responsible_department_box.Location.X) + $($responsible_department_box.Size.Width) + 3),$($($responsible_department_box.Location.Y) - 1))
        $responsible_department_search.Size = New-Object System.Drawing.Size(21,21)
        $responsible_department_search.BackgroundImage = $image
        $responsible_department_search.BackgroundImageLayout = 'Zoom'

        $department_trigger = {
            $search = $responsible_department_box.Text
            $uri = "https://$instance_name/api/now/table/cmn_department?sysparm_query=nameLIKE$search"
            $search_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
            $responsible_department_list = @()
            foreach ($search_result in $($search_response.result)) {
                If ($($search_result.business_unit) -notmatch '\w') {
                    $business_unit = "Empty"
                }else{
                    $uri = $search_result.business_unit.link
                    $business_unit = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                    $business_unit = $business_unit.result.name
                }
                $row = "" | Select Name,BusinessUnit,ID,SysID
                $row.Name = $search_result.name
                $row.BusinessUnit = $business_unit
                $row.ID = $search_response.id
                $row.SysID = $search_result.sys_id
                $responsible_department_list += $row
            }
            $global:responsible_department_list = $responsible_department_list | Sort-Object -Property Name

            Add-Type -AssemblyName System.Windows.Forms
            Add-Type -AssemblyName System.Drawing

            $search_form = New-Object System.Windows.Forms.Form
            $search_form.Text = "Department Search:"
            $search_form.AutoSize = $true
            $search_form.StartPosition = 'CenterScreen'

            $search_label = New-Object System.Windows.Forms.Label
            $search_label.Text = "Search:"
            $search_label.Location = New-Object System.Drawing.Point(5,8)
            $search_label.Size = New-Object System.Drawing.Size(50,20)
            $search_form.Controls.Add($search_label)

            $results_label = New-Object System.Windows.Forms.Label
            $results_label.Text = "Results:"
            $results_label.Location = New-Object System.Drawing.Point(5,70)
            $results_label.Size = New-Object System.Drawing.Size(50,21)
            $search_form.Controls.Add($results_label)

            $selectionBox = New-Object System.Windows.Forms.ListBox
            $selectionBox.Location = New-Object System.Drawing.Point(5,91)
            $selectionBox.AutoSize = $true
            $selectionBox.MinimumSize = New-Object System.Drawing.Size(417,200)
            $selectionBox.MaximumSize = New-Object System.Drawing.Size(0,200)
            $selectionBox.ScrollAlwaysVisible = $true
            $selectionBox.Items.Clear()
            if ($responsible_department_list.name -notmatch '\w') {
                $selectionBox.Items.Add("No Results Found")
            }else{
                foreach ($item in $responsible_department_list) {
                    [void] $selectionBox.Items.Add("$($item.name) ($($item.businessunit))")
                }
            }
            $search_form.TopMost = $true
            $search_form.Controls.Add($selectionBox)

            $search_bar = New-Object System.Windows.Forms.TextBox
            $search_bar.Location = New-Object System.Drawing.Size(5,28)
            $search_bar.Size = New-Object System.Drawing.Size(400,21)
            $search_form.Controls.Add($search_bar)
            $search_bar.Text = $responsible_department_box.Text

            $search_button = New-Object System.Windows.Forms.Button
            $search_button.Location = New-Object System.Drawing.Point(408,27)
            $search_button.Size = New-Object System.Drawing.Size(21,21)
            $search_button.BackgroundImage = $image
            $search_button.BackgroundImageLayout = 'Zoom'

            $search_trigger = {
                $selectionBox.Items.Clear()
                $selectionBox.Items.Add("Processing . . .")
                $search = $search_bar.Text
                $uri = "https://$instance_name/api/now/table/cmn_department?sysparm_query=nameLIKE$search"
                $search_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                $responsible_department_list = @()
                foreach ($search_result in $($search_response.result)) {
                    If ($($search_result.business_unit) -notmatch '\w') {
                        $business_unit = "Empty"
                    }else{
                        $uri = $search_result.business_unit.link
                        $business_unit = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                        $business_unit = $business_unit.result.name
                    }
                    $row = "" | Select Name,BusinessUnit,ID,SysID
                    $row.Name = $search_result.name
                    $row.BusinessUnit = $business_unit
                    $row.ID = $search_response.id
                    $row.SysID = $search_result.sys_id
                    $responsible_department_list += $row
                }
                $global:responsible_department_list = $responsible_department_list | Sort-Object -Property Name
                $selectionBox.Items.Clear()
                if ($responsible_department_list.name -notmatch '\w') {
                    $selectionBox.Items.Add("No Results Found")
                }else{
                    foreach ($item in $responsible_department_list) {
                        [void] $selectionBox.Items.Add("$($item.name) ($($item.businessunit))")
                    }
                }
            }
            $search_button.add_Click($search_trigger)
            $search_form.Controls.Add($search_button)

            $search_ok_button = New-Object System.Windows.Forms.Button
            $search_ok_button.Location = New-Object System.Drawing.Point(347, $($($selectionBox.Location.Y) + $($selectionBox.Size.Height) + 10))
            $search_ok_button.Size = New-Object System.Drawing.Size(75,23)
            $search_ok_button.Text = 'Confirm'
            $search_ok_button.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $search_form.Controls.Add($search_ok_button)

            $search_accept = {
                $search_form.AcceptButton = $search_button
                $selectionBox.SelectedItem = $null
            }
            $form_accept = {
                $search_form.AcceptButton = $search_ok_button
            }

            $search_bar.add_MouseDown($search_accept)
            $selectionBox.add_MouseDown($form_accept)

            $search_confirm = $search_form.ShowDialog()

            if ($search_confirm -eq [System.Windows.Forms.DialogResult]::OK) {
                $responsible_department_box.Text = $selectionBox.SelectedItem
                $department_box.Text = $selectionBox.SelectedItem
                if ($responsible_department_list.Count -gt 1) {
                    $global:responsible_department_selection_id = $responsible_department_list.SysID[$($selectionBox.SelectedIndex)]
                    $global:department_selection_id = $responsible_department_list.SysID[$($selectionBox.SelectedIndex)]
                }else{
                    $global:responsible_department_selection_id = $responsible_department_list.SysID
                    $global:department_selection_id = $responsible_department_list.SysID
                }
            }
        }

        $responsible_department_search.add_click($department_trigger)
        $recent_form.Controls.Add($responsible_department_search)


        #Parent Serial Number Search Button
        $parent_serial_search = New-Object System.Windows.Forms.Button
        $parent_serial_search.Location = New-Object System.Drawing.Point($($($parent_box.Location.X) + $($parent_box.Size.Width) + 3),$($($parent_box.Location.Y) - 1))
        $parent_serial_search.Size = New-Object System.Drawing.Size(21,21)
        $parent_serial_search.BackgroundImage = $image
        $parent_serial_search.BackgroundImageLayout = 'Zoom'

        $parent_trigger = {
            $search = $parent_box.Text
            $uri = "https://$instance_name/api/now/table/alm_asset?sysparm_query=serial_number%3D$search"
            $search_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri

            Add-Type -AssemblyName System.Windows.Forms
            Add-Type -AssemblyName System.Drawing

            $search_form = New-Object System.Windows.Forms.Form
            $search_form.Text = "Parent Serial Number Search:"
            $search_form.AutoSize = $true
            $search_form.StartPosition = 'CenterScreen'

            $search_label = New-Object System.Windows.Forms.Label
            $search_label.Text = "Search (Serial Number):"
            $search_label.Location = New-Object System.Drawing.Point(5,8)
            $search_label.Size = New-Object System.Drawing.Size(150,20)
            $search_form.Controls.Add($search_label)

            $results_label = New-Object System.Windows.Forms.Label
            $results_label.Text = "Results:"
            $results_label.Location = New-Object System.Drawing.Point(5,70)
            $results_label.Size = New-Object System.Drawing.Size(50,21)
            $search_form.Controls.Add($results_label)

            $selectionBox = New-Object System.Windows.Forms.ListBox
            $selectionBox.Location = New-Object System.Drawing.Point(5,91)
            $selectionBox.AutoSize = $true
            $selectionBox.MinimumSize = New-Object System.Drawing.Size(417,200)
            $selectionBox.MaximumSize = New-Object System.Drawing.Size(0,200)
            $selectionBox.ScrollAlwaysVisible = $true
            $selectionBox.Items.Clear()
            $parent_list = @()
            foreach ($parent_result in $search_response.result) {
                $row = "" | Select SerialNumber,SysID
                $row.SerialNumber = $parent_result.serial_number
                $row.SysID = $parent_result.sys_id
                $parent_list += $row
            }
            $global:parent_list = $parent_list | Sort-Object -Property SerialNumber
            if ($parent_list.SerialNumber -notmatch '\w') {
                $selectionBox.Items.Add("No Results Found")
                $global:search = $search
            }else{
                foreach ($search_result in $($parent_list.SerialNumber)) {
                    [void] $selectionBox.Items.Add($search_result)
                }
            }
            $search_form.TopMost = $true
            $search_form.Controls.Add($selectionBox)

            $search_bar = New-Object System.Windows.Forms.TextBox
            $search_bar.Location = New-Object System.Drawing.Size(5,28)
            $search_bar.Size = New-Object System.Drawing.Size(400,21)
            $search_form.Controls.Add($search_bar)
            $search_bar.Text = $parent_box.Text

            $search_button = New-Object System.Windows.Forms.Button
            $search_button.Location = New-Object System.Drawing.Point(408,27)
            $search_button.Size = New-Object System.Drawing.Size(21,21)
            $search_button.BackgroundImage = $image
            $search_button.BackgroundImageLayout = 'Zoom'

            $search_trigger = {
                $search = $search_bar.Text
                $selectionBox.Items.Clear()
                $selectionBox.Items.Add("Processing . . .")
                $uri = "https://$instance_name/api/now/table/alm_asset?sysparm_query=serial_number%3D$search"
                $search_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                $selectionBox.Items.Clear()
                $parent_list = @()
                foreach ($parent_result in $search_response.result) {
                    $row = "" | Select SerialNumber,SysID
                    $row.SerialNumber = $parent_result.serial_number
                    $row.SysID = $parent_result.sys_id
                    $parent_list += $row
                }
                $global:parent_list = $parent_list | Sort-Object -Property SerialNumber
                if ($parent_list.SerialNumber -notmatch '\w') {
                    $selectionBox.Items.Add("No Results Found")
                    $global:search = $search
                }else{
                    foreach ($search_result in $($parent_list.SerialNumber)) {
                        [void] $selectionBox.Items.Add($search_result)
                    }
                }
            }

            $search_button.add_click($search_trigger)
            $search_form.Controls.Add($search_button)

            $search_ok_button = New-Object System.Windows.Forms.Button
            $search_ok_button.Location = New-Object System.Drawing.Point(347, $($($selectionBox.Location.Y) + $($selectionBox.Size.Height) + 10))
            $search_ok_button.Size = New-Object System.Drawing.Size(75,23)
            $search_ok_button.Text = 'Confirm'
            $search_ok_button.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $search_form.Controls.Add($search_ok_button)

            $create_new_button = New-Object System.Windows.Forms.Button
            $create_new_button.Location = New-Object System.Drawing.Point(270,$($($selectionBox.Location.Y) + $($selectionBox.Size.Height) + 10))
            $create_new_button.Size = New-Object System.Drawing.Size(75,23)
            $create_new_button.Text = 'Create New'
            
            #Create a New Parent Asset
            $create_parent = {
                #Confirm New Parent Creation
                $search = $search_bar.Text
                $answer = [System.Windows.Forms.MessageBox]::Show("Create new parent item with this serial number: $($search)?" , "Create New Asset" , 4)
                If ($answer -eq "Yes") {
                    #Create Parent Form Box
                    Add-Type -AssemblyName System.Windows.Forms
                    Add-Type -AssemblyName System.Drawing

                    $parent_form = New-Object System.Windows.Forms.Form
                    $parent_form.Text = "Create New Parent Item"
                    $parent_form.AutoSize = $true
                    $parent_form.StartPosition = 'CenterScreen'

                    $serial_number_label = New-Object System.Windows.Forms.Label
                    $serial_number_label.Text = "Serial Number:"
                    $serial_number_label.Location = New-Object System.Drawing.Point(5,8)
                    $serial_number_label.AutoSize = $true
                    $parent_form.Controls.Add($serial_number_label)

                    $category_label = New-Object System.Windows.Forms.Label
                    $category_label.Text = "Model Category:"
                    $category_label.Location = New-Object System.Drawing.Point(5,$($serial_number_label.Location.Y + $serial_number_label.Size.Height + 10))
                    $category_label.AutoSize = $true
                    $parent_form.Controls.Add($category_label)

                    $model_label = New-Object System.Windows.Forms.Label
                    $model_label.Text = "Model:"
                    $model_label.Location = New-Object System.Drawing.Point(5,$($category_label.Location.Y + $category_label.Size.Height + 10))
                    $model_label.AutoSize = $true
                    $parent_form.Controls.Add($model_label)

                    $responsible_department_label = New-Object System.Windows.Forms.Label
                    $responsible_department_label.Text = "Responsible Department:"
                    $responsible_department_label.Location = New-Object System.Drawing.Point(5,$($model_label.Location.Y + $model_label.Size.Height + 10))
                    $responsible_department_label.AutoSize = $true
                    $parent_form.Controls.Add($responsible_department_label)

                    $location_label = New-Object System.Windows.Forms.Label
                    $location_label.Text = "Location:"
                    $location_label.Location = New-Object System.Drawing.Point(5,$($responsible_department_label.Location.Y + $responsible_department_label.Size.Height + 10))
                    $location_label.AutoSize = $true
                    $parent_form.Controls.Add($location_label)
                    
                    $comments_label = New-Object System.Windows.Forms.Label
                    $comments_label.Text = "Comments:"
                    $comments_label.Location = New-Object System.Drawing.Point(5,$($location_label.Location.Y + $location_label.Size.Height + 10))
                    $comments_label.AutoSize = $true
                    $parent_form.Controls.Add($comments_label)

                    $serial_number_box = New-Object System.Windows.Forms.TextBox
                    $serial_number_box.Location = New-Object System.Drawing.Point($($responsible_department_label.Size.Width + 10),$($serial_number_label.Location.Y - 2))
                    $serial_number_box.Size = New-Object System.Drawing.Size(300,18)
                    $serial_number_box.Text = $search
                    $parent_form.Controls.Add($serial_number_box)

                    $category_box = New-Object System.Windows.Forms.ComboBox
                    $category_box.Location = New-Object System.Drawing.Point($($responsible_department_label.Size.Width + 10),$($category_label.Location.Y - 2))
                    $category_box.Size = New-Object System.Drawing.Size(300,18)
                    $category_box.DropDownHeight = 100
                    $range = "Computer Wall Mount","Work Station on Wheels"
                    $category_box.Items.AddRange(@($range))

                    $model_box = New-Object System.Windows.Forms.ComboBox
                    $model_box.Location = New-Object System.Drawing.Point($($responsible_department_label.Size.Width + 10),$($model_label.Location.Y - 2))
                    $model_box.Size = New-Object System.Drawing.Size(300,18)
                    $model_box.DropDownHeight = 150
                    #Update Model Options Based on Category Selection
                    $update_model = {
                        $model_box.Items.Clear()
                        $uri = "https://$instance_name/api/now/table/cmdb_model_category?sysparm_query=name%3D$($category_box.SelectedItem)"
                        $model_id = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                        $model_id = $model_id.result.sys_id
                        $uri = "https://$instance_name/api/now/table/cmdb_model?sysparm_query=cmdb_model_category%3D$($model_id)"
                        $model_range = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                        $model_range = $model_range.result.display_name | Sort-Object
                        foreach ($item in $model_range) {
                            $model_box.Items.Add($item)
                        }
                    }
                    $category_box.add_SelectedIndexChanged($update_model)
                    $parent_form.Controls.Add($category_box)

                    $parent_form.Controls.Add($model_box)

                    $parent_responsible_department_box = New-Object System.Windows.Forms.TextBox
                    $parent_responsible_department_box.Location = New-Object System.Drawing.Point($($responsible_department_label.Size.Width + 10),$($responsible_department_label.Location.Y - 2))
                    $parent_responsible_department_box.Size = New-Object System.Drawing.Size(300,18)
                    If ($responsible_department_box.Text -notmatch '\w') {
                        $parent_responsible_department_box.Text = "$($recent_entry.assignment.responsible_department)"
                        $managed_by_update = $recent_entry.assignment.department_id
                    }else{
                        $parent_responsible_department_box.Text = $responsible_department_box.Text
                        $managed_by_update = $responsible_department_selection_id
                    }

                    $parent_department_trigger = {
                        $search = $parent_responsible_department_box.Text
                        $uri = "https://$instance_name/api/now/table/cmn_department?sysparm_query=nameLIKE$search"
                        $search_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                        $parent_responsible_department_list = @()
                        foreach ($search_result in $($search_response.result)) {
                            If ($($search_result.business_unit) -notmatch '\w') {
                                $business_unit = "Empty"
                            }else{
                                $uri = $search_result.business_unit.link
                                $business_unit = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                                $business_unit = $business_unit.result.name
                            }
                            $row = "" | Select Name,BusinessUnit,ID,SysID
                            $row.Name = $search_result.name
                            $row.BusinessUnit = $business_unit
                            $row.ID = $search_response.id
                            $row.SysID = $search_result.sys_id
                            $parent_responsible_department_list += $row
                        }
                        $parent_responsible_department_list = $parent_responsible_department_list | Sort-Object -Property Name

                        Add-Type -AssemblyName System.Windows.Forms
                        Add-Type -AssemblyName System.Drawing

                        $search_form = New-Object System.Windows.Forms.Form
                        $search_form.Text = "Department Search:"
                        $search_form.AutoSize = $true
                        $search_form.StartPosition = 'CenterScreen'

                        $search_label = New-Object System.Windows.Forms.Label
                        $search_label.Text = "Search:"
                        $search_label.Location = New-Object System.Drawing.Point(5,8)
                        $search_label.Size = New-Object System.Drawing.Size(50,20)
                        $search_form.Controls.Add($search_label)

                        $results_label = New-Object System.Windows.Forms.Label
                        $results_label.Text = "Results:"
                        $results_label.Location = New-Object System.Drawing.Point(5,70)
                        $results_label.Size = New-Object System.Drawing.Size(50,21)
                        $search_form.Controls.Add($results_label)

                        $selectionBox = New-Object System.Windows.Forms.ListBox
                        $selectionBox.Location = New-Object System.Drawing.Point(5,91)
                        $selectionBox.AutoSize = $true
                        $selectionBox.MinimumSize = New-Object System.Drawing.Size(417,200)
                        $selectionBox.MaximumSize = New-Object System.Drawing.Size(0,200)
                        $selectionBox.ScrollAlwaysVisible = $true
                        $selectionBox.Items.Clear()
                        if ($parent_responsible_department_list.name -notmatch '\w') {
                            $selectionBox.Items.Add("No Results Found")
                        }else{
                            foreach ($item in $parent_responsible_department_list) {
                                [void] $selectionBox.Items.Add("$($item.name) ($($item.businessunit))")
                            }
                        }
                        $search_form.TopMost = $true

                        $search_bar = New-Object System.Windows.Forms.TextBox
                        $search_bar.Location = New-Object System.Drawing.Size(5,28)
                        $search_bar.Size = New-Object System.Drawing.Size(400,21)
                        $search_form.Controls.Add($search_bar)
                        $search_bar.Text = $parent_responsible_department_box.Text

                        $search_button = New-Object System.Windows.Forms.Button
                        $search_button.Location = New-Object System.Drawing.Point(408,27)
                        $search_button.Size = New-Object System.Drawing.Size(21,21)
                        $search_button.BackgroundImage = $image
                        $search_button.BackgroundImageLayout = 'Zoom'

                        $search_trigger = {
                            $selectionBox.Items.Clear()
                            $selectionBox.Items.Add("Processing . . .")
                            $search = $search_bar.Text
                            $uri = "https://$instance_name/api/now/table/cmn_department?sysparm_query=nameLIKE$search"
                            $search_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                            $parent_responsible_department_list = @()
                            foreach ($search_result in $($search_response.result)) {
                                If ($($search_result.business_unit) -notmatch '\w') {
                                    $business_unit = "Empty"
                                }else{
                                    $uri = $search_result.business_unit.link
                                    $business_unit = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                                    $business_unit = $business_unit.result.name
                                }
                                $row = "" | Select Name,BusinessUnit,ID,SysID
                                $row.Name = $search_result.name
                                $row.BusinessUnit = $business_unit
                                $row.ID = $search_response.id
                                $row.SysID = $search_result.sys_id
                                $parent_responsible_department_list += $row
                            }
                            $parent_responsible_department_list = $parent_responsible_department_list | Sort-Object -Property Name
                            $selectionBox.Items.Clear()
                            if ($parent_responsible_department_list.name -notmatch '\w') {
                                $selectionBox.Items.Add("No Results Found")
                            }else{
                                foreach ($item in $parent_responsible_department_list) {
                                    [void] $selectionBox.Items.Add("$($item.name) ($($item.businessunit))")
                                }
                            }
                            $global:new_responsible_department_list = $parent_responsible_department_list
                        }
                        $search_button.add_Click($search_trigger)
                        $search_form.Controls.Add($search_button)

                        $save_selection = {
                            if ($new_responsible_department_list -ne $null) {
                                $parent_responsible_department_list = $new_responsible_department_list
                            }
                            if ($parent_responsible_department_list.Count -gt 1) {
                                $global:update = $($parent_responsible_department_list.SysID[$($selectionBox.SelectedIndex)])
                            }else{
                                $global:update = $parent_responsible_department_list.SysID
                            }
                        }

                        $selectionBox.add_SelectedIndexChanged($save_selection)
                        $search_form.Controls.Add($selectionBox)

                        $search_ok_button = New-Object System.Windows.Forms.Button
                        $search_ok_button.Location = New-Object System.Drawing.Point(347, $($($selectionBox.Location.Y) + $($selectionBox.Size.Height) + 10))
                        $search_ok_button.Size = New-Object System.Drawing.Size(75,23)
                        $search_ok_button.Text = 'Confirm'
                        $search_ok_button.DialogResult = [System.Windows.Forms.DialogResult]::OK
                        $search_form.Controls.Add($search_ok_button)

                        $search_accept = {
                            $search_form.AcceptButton = $search_button
                            $selectionBox.SelectedItem = $null
                        }
                        $form_accept = {
                            $search_form.AcceptButton = $search_ok_button
                        }

                        $search_bar.add_MouseDown($search_accept)
                        $selectionBox.add_MouseDown($form_accept)

                        $search_confirm = $search_form.ShowDialog()

                        if ($search_confirm -eq [System.Windows.Forms.DialogResult]::OK) {
                            Clear-Variable -Name new_responsible_department_list -ErrorAction SilentlyContinue
                            $parent_responsible_department_box.Text = $selectionBox.SelectedItem
                        }
                    }
                    $parent_responsible_department_box.add_click($parent_department_trigger)
                    $parent_form.Controls.Add($parent_responsible_department_box)

                    $parent_location_box = New-Object System.Windows.Forms.TextBox
                    $parent_location_box.Location = New-Object System.Drawing.Point($($responsible_department_label.Size.Width + 10),$($location_label.Location.Y - 2))
                    $parent_location_box.Size = New-Object System.Drawing.Size(300,18)
                    If ($location_box.Text -notmatch '\w') {
                        $parent_location_box.Text = "$($recent_entry.location)"
                    }else{
                        $parent_location_box.Text = $location_box.Text
                    }
                    $change_parent_location = {
                        $parent_location_box.Clear()
                        Edit-SNAMLocation -SerialNumber $global:SerialNumber
                        $parent_location_box.Text = $SNAMLocation
                    }
                    $parent_location_box.add_click($change_parent_location)
                    $parent_form.Controls.Add($parent_location_box)

                    $comments_box = New-Object System.Windows.Forms.TextBox
                    $comments_box.Location = New-Object System.Drawing.Point($($responsible_department_label.Size.Width + 10),$($comments_label.Location.Y - 2))
                    $comments_box.MultiLine = $true
                    $comments_box.Size = New-Object System.Drawing.Size(300,75)
                    $parent_form.Controls.Add($comments_box)

                    $parent_accept_button = New-Object System.Windows.Forms.Button
                    $parent_accept_button.Location = New-Object System.Drawing.Point($($parent_form.Size.Width - 93),$($comments_box.Location.Y + $comments_box.Size.Height + 5))
                    $parent_accept_button.Size = New-Object System.Drawing.Size(75,23)
                    $parent_accept_button.Text = 'Confirm'

                    $parent_confirm_click = {
                        #Mark required fields if not completed
                        If ($serial_number_box.Text -notmatch '\w' -or $category_box.Text -notmatch '\w' -or $model_box.Text -notmatch '\w' -or $parent_responsible_department_box.Text -notmatch '\w' -or $parent_location_box.Text -notmatch '\w') {
                            $parent_error_label = New-Object System.Windows.Forms.Label
                            $parent_error_label.Location = New-Object System.Drawing.Point(5,$($comments_box.Location.Y + $comments_box.Size.Height + 7))
                            $parent_error_label.AutoSize = $true
                            $parent_error_label.ForeColor = 'RED'
                            $parent_error_label.Text = "Please enter all required data"
                            $serial_number_label.ForeColor = 'RED'
                            $category_label.ForeColor = 'RED'
                            $model_label.ForeColor = 'RED'
                            $responsible_department_label.ForeColor = 'RED'
                            $location_label.ForeColor = 'RED'
                            $parent_form.Controls.Add($parent_error_label)
                        }else{
                            #Create Parent Item
                            $serial_number_update = $serial_number_box.Text
                            $uri = "https://$instance_name/api/now/table/cmdb_model_category?sysparm_query=name%3D$($category_box.SelectedItem)"
                            $model_id = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                            $model_category_update = $model_id.result.sys_id
                            $uri = "https://$instance_name/api/now/table/cmdb_model?sysparm_query=display_name%3D$($model_box.Text)"
                            $model_range = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                            $model_update = $model_range.result.sys_id
                            $install_status_update = $($StateOptions.result | Where-Object {$_.label -eq "In use"}).value
                            If ($update -notmatch '\w') {
                                If ($responsible_department_box.Text -notmatch '\w') {
                                    $update = $($recent_entry.assignment.responsible_department_sys_id)
                                }else{
                                    $update = $responsible_department_selection_id
                                }
                            }
                            $managed_by_update = $update
                            $uri = "https://$instance_name/api/now/table/cmn_location?sysparm_query=name%3D$($parent_location_box.Text)"
                            $location_update = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                            $location_update = $location_update.result.sys_id
                            $date_update = $(Get-Date -Format yyyy-MM-dd)
                            $method = "post"
                            if ($category_box.Text -eq "Computer Wall Mount") {
                                $uri = "https://$instance_name/api/now/table/u_computer_wall_mounts"
                            }else{
                                $uri = "https://$instance_name/api/now/table/u_workstation_on_wheels"
                            }
                            if ($comments_box.Text -match '\w') {
                                $comments_update = $comments_box.Text
                                $comments_update.Replace([Environment]::NewLine,"\n")
                                $body = "{`"serial_number`":`"$serial_number_update`",`"model_category`":`"$model_category_update`",`"model`":`"$model_update`",`"install_status`":`"$install_status_update`",`"managed_by`":`"$managed_by_update`",`"department`":`"$managed_by_update`",`"location`":`"$location_update`",`"u_last_physical_inventory`":`"$date_update`",`"comments`":`"$comments_update`"}"
                            }else{
                                $body = "{`"serial_number`":`"$serial_number_update`",`"managed_by`":`"$managed_by_update`",`"department`":`"$managed_by_update`",`"location`":`"$location_update`",`"u_last_physical_inventory`":`"$date_update`"}"
                            }
                            $create_parent_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body
                            $method = "get"
                            $parent_box.Text = $serial_number_update
                            $global:parent_selection = $serial_number_update
                            $global:parent_selection_id = $create_parent_response.result.sys_id
                        #Parent Created Confirmation
                            Add-Type -AssemblyName System.Windows.Forms
                            Add-Type -AssemblyName System.Drawing
                            $create_form = New-Object System.Windows.Forms.Form
                            $create_form.Text = "Confirm"
                            $create_form.Size = New-Object System.Drawing.Size(304,150)
                            $create_form.StartPosition = 'CenterScreen'

                            $create_label = New-Object System.Windows.Forms.Label
                            $create_label.Location = New-Object System.Drawing.Point($((304/2)-($($create_label.Size.Width)/2)),20)
                            $create_label.AutoSize = $true
                            $create_label.Text = "$serial_number_update has been created"
                            $create_form.Controls.Add($create_label)

                            $create_button = New-Object System.Windows.Forms.Button
                            $create_button.Location = New-Object System.Drawing.Point(114,80)
                            $create_button.Size = New-Object System.Drawing.Size(75,23)
                            $create_button.Text = 'OK'
                            $create_button.DialogResult = [System.Windows.Forms.DialogResult]::OK
                            $create_form.Controls.Add($create_button)
                            $create_form.TopMost = $true
                            $create_confirm = $create_form.ShowDialog()
                            If ($create_confirm -eq [System.Windows.Forms.DialogResult]::OK) {
                                $parent_form.Close()
                                $search_form.Close()
                            }
                        }
                    }
                    $clear_error = {
                        #Clear error
                        $serial_number_label.ForeColor = 'BLACK'
                        $category_label.ForeColor = 'BLACK'
                        $model_label.ForeColor = 'BLACK'
                        $responsible_department_label.ForeColor = 'BLACK'
                        $location_label.ForeColor = 'BLACK'
                    }
                    $serial_number_box.add_click($clear_error)
                    $category_box.add_click($clear_error)
                    $model_box.add_click($clear_error)
                    $parent_responsible_department_box.add_click($clear_error)
                    $location_box.add_click($clear_error)
                    $parent_form.add_click($clear_error)

                    $parent_accept_button.add_click($parent_confirm_click)
                    $parent_form.Controls.Add($parent_accept_button)

                    $parent_cancel_button = New-Object System.Windows.Forms.Button
                    $parent_cancel_button.Location = New-Object System.Drawing.Point($($parent_accept_button.Location.X - 80),$($parent_accept_button.Location.Y))
                    $parent_cancel_button.Size = New-Object System.Drawing.Size(75,23)
                    $parent_cancel_button.Text = 'Cancel'
                    $parent_cancel_button.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                    $parent_form.Controls.Add($parent_cancel_button)

                    $parent_form.TopMost = $true
                    $parent_confirm = $parent_form.ShowDialog()
                }
            }

            $create_new_button.add_click($create_parent)
            $search_form.Controls.Add($create_new_button)

            $search_accept = {
                $search_form.AcceptButton = $search_button
                $selectionBox.SelectedItem = $null
            }
            $form_accept = {
                $search_form.AcceptButton = $search_ok_button
            }

            $search_bar.add_MouseDown($search_accept)
            $selectionBox.add_MouseDown($form_accept)

            $search_confirm = $search_form.ShowDialog()

            if ($search_confirm -eq [System.Windows.Forms.DialogResult]::OK) {
                $parent_box.Text = $selectionBox.SelectedItem
                if ($parent_list.Count -gt 1) {
                    $global:parent_selection_id = $parent_list.SysId[$($selectionBox.SelectedIndex)]
                }else{
                    $global:parent_selection_id = $parent_list.SysId
                }
            }
        }
        $parent_serial_search.add_click($parent_trigger)
        $recent_form.Controls.Add($parent_serial_search)

        #Remove Parent Button
        If ($($recent_entry.parent.serial_number) -ne $null) {
            $remove_parent_button = New-Object System.Windows.Forms.Button
            $remove_parent_button.Location = New-Object System.Drawing.Point($($parent_serial_search.Location.X + $parent_serial_search.Size.Width + 3),$($parent_serial_search.Location.Y))
            $remove_parent_button.Size = New-Object System.Drawing.Size(98,23)
            $remove_parent_button.Text = "Remove Parent"

            $remove_parent = {
                $answer = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to remove $($recent_entry.serial_number) from $($recent_entry.parent.serial_number)?" , "Remove Parent Asset" , 4)
                If ($answer -eq "Yes") {
                    $method = "patch"
                    $uri = "https://$instance_name/api/now/table/alm_hardware/$($recent_entry.sys_id)"
                    $body = '{"parent":""}'
                    $global:parent_patch = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body
                    $method = "get"
                    $parent_item.Text = ""
                    $recent_form.Controls.Add($parent_item)
                }
            }
            $remove_parent_button.add_click($remove_parent)
            $recent_form.Controls.Add($remove_parent_button)
        }

        #Location Search Button
        $location_search = New-Object System.Windows.Forms.Button
        $location_search.Location = New-Object System.Drawing.Point($($($location_box.Location.X) + $($location_box.Size.Width) + 3),$($($location_box.Location.Y) - 1))
        $location_search.Size = New-Object System.Drawing.Size(21,21)
        $location_search.BackgroundImage = $image
        $location_search.BackgroundImageLayout = 'Zoom'

        $location_search_trigger = {
            $search = $location_box.Text
            $uri = "https://$instance_name/api/now/table/cmn_location?sysparm_query=nameLIKE$search"
            $search_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri

            Clear-Variable -Name IPAddress -ErrorAction SilentlyContinue
            Clear-Variable -Name ComputerIP -ErrorAction SilentlyContinue
            Clear-Variable -Name IPScope -ErrorAction SilentlyContinue
            Clear-Variable -Name DHCPlocation -ErrorAction SilentlyContinue
            $IPAddress = [Net.Dns]::GetHostAddresses("$ComputerName") | Where-Object { $_.AddressFamily -eq 'InterNetwork' } | Select-Object -Expand IPAddressToString
            $ComputerIP = $IPAddress -split ".", 0, "simplematch"
            $a = echo $ComputerIP[0]
            $b = echo $ComputerIP[1]
            $c = echo $ComputerIP[2]
            $d = "0"
            $e = "."
            $IPScope = echo "$a$e$b$e$c$e$d"
            $DHCPlocation = (Get-DhcpServerv4Scope -ComputerName "dhcpw01.wakemed.org" -ScopeId $IPScope -ErrorAction SilentlyContinue).Name

            Add-Type -AssemblyName System.Windows.Forms
            Add-Type -AssemblyName System.Drawing

            $search_form = New-Object System.Windows.Forms.Form
            $search_form.Text = "Location Search:"
            $search_form.AutoSize = $true
            $search_form.StartPosition = 'CenterScreen'

            $IP_label = New-Object System.Windows.Forms.Label
            $IP_label.Text = "IP Address: $IPAddress"
            $IP_label.Location = New-Object System.Drawing.Point(5,8)
            $IP_label.AutoSize = $true
            $IP_label.MinimumSize = New-Object System.Drawing.Size(70,20)
            $IP_label.MaximumSize = New-Object System.Drawing.Size(0,20)
            $search_form.Controls.Add($IP_label)

            $Location_label = New-Object System.Windows.Forms.Label
            $Location_label.Text = "DHCP Location: $DHCPlocation"
            $Location_label.Location = New-Object System.Drawing.Point($($IP_label.Location.X + $IP_label.Size.Width + 20),8)
            $Location_label.AutoSize = $true
            $Location_label.MinimumSize = New-Object System.Drawing.Size(100,20)
            $Location_label.MaximumSize = New-Object System.Drawing.Size(0,20)
            $search_form.Controls.Add($Location_label)

            $search_label = New-Object System.Windows.Forms.Label
            $search_label.Text = "Search:"
            $search_label.Location = New-Object System.Drawing.Point(5,38)
            $search_label.Size = New-Object System.Drawing.Size(50,20)
            $search_form.Controls.Add($search_label)

            $results_label = New-Object System.Windows.Forms.Label
            $results_label.Text = "Results:"
            $results_label.Location = New-Object System.Drawing.Point(5,100)
            $results_label.Size = New-Object System.Drawing.Size(50,21)
            $search_form.Controls.Add($results_label)

            $selectionBox = New-Object System.Windows.Forms.ListBox
            $selectionBox.Location = New-Object System.Drawing.Point(5,121)
            $selectionBox.AutoSize = $true
            $selectionBox.MinimumSize = New-Object System.Drawing.Size(417,200)
            $selectionBox.MaximumSize = New-Object System.Drawing.Size(0,200)
            $selectionBox.ScrollAlwaysVisible = $true
            $selectionBox.Items.Clear()
            $location_results = $search_response.result.name | Sort-Object
            If ($location_results -notmatch '\w') {
                $selectionBox.Items.Add("No Results Found")
            }else{
                foreach ($item in $location_results) {
                    [void] $selectionBox.Items.Add("$item")
                }
            }
            $search_form.TopMost = $true
            $search_form.Controls.Add($selectionBox)

            $search_bar = New-Object System.Windows.Forms.TextBox
            $search_bar.Location = New-Object System.Drawing.Size(5,58)
            $search_bar.Size = New-Object System.Drawing.Size(400,21)
            $search_form.Controls.Add($search_bar)
            $search_bar.Text = $location_box.Text

            $search_button = New-Object System.Windows.Forms.Button
            $search_button.Location = New-Object System.Drawing.Point(408,57)
            $search_button.Size = New-Object System.Drawing.Size(21,21)
            $search_button.BackgroundImage = $image
            $search_button.BackgroundImageLayout = 'Zoom'

            $search_trigger = {
                $search = $search_bar.Text
                $selectionBox.Items.Clear()
                $selectionBox.Items.Add("Processing . . .")
                $uri = "https://$instance_name/api/now/table/cmn_location?sysparm_query=nameLIKE$search"
                $search_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                $selectionBox.Items.Clear()
                $location_results = $search_response.result.name | Sort-Object
                If ($location_results -notmatch '\w') {
                    $selectionBox.Items.Add("No Results Found")
                }else{
                    foreach ($item in $location_results) {
                        [void] $selectionBox.Items.Add("$item")
                    }
                }
            }

            $search_button.add_click($search_trigger)
            $search_form.Controls.Add($search_button)

            $search_ok_button = New-Object System.Windows.Forms.Button
            $search_ok_button.Location = New-Object System.Drawing.Point(347, $($($selectionBox.Location.Y) + $($selectionBox.Size.Height) + 10))
            $search_ok_button.Size = New-Object System.Drawing.Size(75,23)
            $search_ok_button.Text = 'Confirm'
            $search_ok_button.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $search_form.Controls.Add($search_ok_button)

            $search_accept = {
                $search_form.AcceptButton = $search_button
                $selectionBox.SelectedItem = $null
            }
            $form_accept = {
                $search_form.AcceptButton = $search_ok_button
            }

            $location_dropdown_button = New-Object System.Windows.Forms.Button
            $location_dropdown_button.Location = New-Object System.Drawing.Point(267, $($($selectionBox.Location.Y) + $($selectionBox.Size.Height) + 10))
            $location_dropdown_button.Size = New-Object System.Drawing.Size(75,23)
            $location_dropdown_button.Text = 'DropDowns'

            $location_click = {
                Edit-SNAMLocation -SerialNumber $global:SerialNumber
                $location_box.Text = $Selected_Location
                $search_form.Close()
            }

            $location_dropdown_button.add_click($location_click)
            $search_form.Controls.Add($location_dropdown_button)

            $search_bar.add_MouseDown($search_accept)
            $selectionBox.add_MouseDown($form_accept)

            $search_confirm = $search_form.ShowDialog()

            if ($search_confirm -eq [System.Windows.Forms.DialogResult]::OK) {
                $location_box.Text = $selectionBox.SelectedItem
            }
        }

        $location_search.add_click($location_search_trigger)
        $recent_form.Controls.Add($location_search)

    #Consumables Search
        $consumables_click = {
            $global:consumables_list = New-Object System.Collections.ArrayList
            If ($($recent_entry.asset_id) -ne $null) {
                $uri = "https://$instance_name/api/now/table/alm_consumable?sysparm_query=parent%3D$($recent_entry.asset_id)"
                $comp_consumables_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
            }else{
                If ($($comp_ci_response.result.asset.value) -ne $null) {
                    $uri = "https://$instance_name/api/now/table/alm_consumable?sysparm_query=parent%3D$($comp_ci_response.result.asset.value)"
                    $comp_consumables_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                }
            }
            If ($parent_box.Text -match '\w') {
                $uri = "https://$instance_name/api/now/table/alm_consumable?sysparm_query=parent%3D$($parent_selection_id)"
                $parent_consumables_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
            }else{
                If ($($recent_entry.parent.serial_number) -match '\w') {
                    $uri = "https://$instance_name/api/now/table/alm_consumable?sysparm_query=parent%3D$($recent_entry.parent.sys_id)"
                    $parent_consumables_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                }
            }

            Add-Type -AssemblyName System.Windows.Forms
            Add-Type -AssemblyName System.Drawing

            $search_form = New-Object System.Windows.Forms.Form
            $search_form.Text = "Consumables Search:"
            $search_form.AutoSize = $true
            $search_form.StartPosition = 'CenterScreen'

            $search_label = New-Object System.Windows.Forms.Label
            $search_label.Text = "Search:"
            $search_label.Location = New-Object System.Drawing.Point(5,8)
            $search_label.Size = New-Object System.Drawing.Size(50,20)
            $search_form.Controls.Add($search_label)

            $results_label = New-Object System.Windows.Forms.Label
            $results_label.Text = "Results:"
            $results_label.Location = New-Object System.Drawing.Point(5,70)
            $results_label.Size = New-Object System.Drawing.Size(50,21)
            $search_form.Controls.Add($results_label)

            $selectionBox = New-Object System.Windows.Forms.ListBox
            $selectionBox.Location = New-Object System.Drawing.Point(5,91)
            $selectionBox.AutoSize = $true
            $selectionBox.MinimumSize = New-Object System.Drawing.Size(417,200)
            $selectionBox.MaximumSize = New-Object System.Drawing.Size(0,200)
            $selectionBox.ScrollAlwaysVisible = $true
            $selectionBox.Items.Clear()
            foreach ($consumable_result in $comp_consumables_response.result) {
                $row = "" | Select Name,Model,SysID
                $row.Name = $consumable_result.display_name
                $row.Model = $consumable_result.model.value
                for ($j = 1; $j -le $consumable_result.quantity; $j++) {
                    $global:consumables_list += $row
                }
            }
            If ($parent_consumables_response.result -ne $null) {
                foreach ($consumable_result in $parent_consumables_response.result) {
                    $row = "" | Select Name,Model,SysID
                    $row.Name = $consumable_result.display_name
                    $row.Model = $consumable_result.model.value
                    for ($j = 1; $j -le $consumable_result.quantity; $j++) {
                        $global:consumables_list += $row
                    }
                }
            }
            $global:consumables_list = $global:consumables_list | Sort-Object -Property Name
            if ($global:consumables_list.Name -notmatch '\w') {
                [void]$selectionBox.Items.Add("No Consumables Currently Assigned to $SerialNumber")
            }else{
                foreach ($consumable_item in $($global:consumables_list.Name)) {
                    [void]$selectionBox.Items.Add($consumable_item)
                }
            }
            $search_form.TopMost = $true
            $search_form.Controls.Add($selectionBox)

            $search_bar = New-Object System.Windows.Forms.TextBox
            $search_bar.Location = New-Object System.Drawing.Size(5,28)
            $search_bar.Size = New-Object System.Drawing.Size(393,21)
            $search_form.Controls.Add($search_bar)

            $search_button = New-Object System.Windows.Forms.Button
            $search_button.Location = New-Object System.Drawing.Point(401,27)
            $search_button.Size = New-Object System.Drawing.Size(21,21)
            $search_button.BackgroundImage = $image
            $search_button.BackgroundImageLayout = 'Zoom'

            $search_trigger = {
                $search = $search_bar.Text
                $selectionBox.Items.Clear()
                $selectionBox.Items.Add("Processing . . .")

                $uri = "https://$instance_name/api/now/table/cmdb_model?sysparm_target=alm_consumable&sysparm_query=display_nameLIKE$search"
                $comp_consumables_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                $selectionBox.Items.Clear()
                $global:consumables_list = New-Object System.Collections.ArrayList
                foreach ($consumable_result in $comp_consumables_response.result) {
                    $row = "" | Select Name,Model
                    $row.Name = $consumable_result.display_name
                    $row.Model = $consumable_result.sys_id
                    $global:consumables_list += $row
                }
                <#
                If ($parent_consumables_response.result -ne $null) {
                    foreach ($consumable_result in $parent_consumables_response.result) {
                        $row = "" | Select Name,Model,SysID
                        $row.Name = $consumable_result.display_name
                        $row.Model = $consumable_result.model.value
                        $global:consumables_list += $row
                    }
                }
                #>
                if ($global:consumables_list.Count -gt 1) {
                    $global:consumables_list = $global:consumables_list | Sort-Object -Property Name
                }else{
                    $global:consumables_list = $global:consumables_list
                }
                if ($global:consumables_list.Name -notmatch '\w') {
                    [void]$selectionBox.Items.Add("No Results Found")
                }else{
                    foreach ($consumable_item in $($global:consumables_list.Name)) {
                        [void]$selectionBox.Items.Add($consumable_item)
                    }
                }
            }

            $search_button.add_click($search_trigger)
            $search_form.Controls.Add($search_button)

            $selection_add_button = New-Object System.Windows.Forms.Button
            $selection_add_button.Location = New-Object System.Drawing.Point(342, $($($selectionBox.Location.Y) + $($selectionBox.Size.Height) + 10))
            $selection_add_button.Size = New-Object System.Drawing.Size(80,23)
            $selection_add_button.Text = 'Add Item'
            $search_form.Controls.Add($selection_add_button)

            $selection_remove_button = New-Object System.Windows.Forms.Button
            $selection_remove_button.Location = New-Object System.Drawing.Point($($($selection_add_button.Location.X) - $($selection_add_button.Size.Width) - 5),$($selection_add_button.Location.Y))
            $selection_remove_button.Size = New-Object System.Drawing.Size(80,23)
            $selection_remove_button.Text = 'Remove Item'
            $search_form.Controls.Add($selection_remove_button)

            $selectionBox2 = New-Object System.Windows.Forms.ListBox
            $selectionBox2.Location = New-Object System.Drawing.Point(5,$($($selection_add_button.Location.Y) + $($selection_add_button.Size.Height) + 10))
            $selectionBox2.AutoSize = $true
            $selectionBox2.MinimumSize = New-Object System.Drawing.Size(417,200)
            $selectionBox2.MaximumSize = New-Object System.Drawing.Size(0,200)
            $selectionBox2.ScrollAlwaysVisible = $true
            $search_form.Controls.Add($selectionBox2)

            $add_label = New-Object System.Windows.Forms.Label
            $add_label.Text = "Items to Add:"
            $add_label.Location = New-Object System.Drawing.Point(5,$($($selectionBox2.Location.Y) - 21))
            $add_label.Size = New-Object System.Drawing.Size(80,21)
            $search_form.Controls.Add($add_label)

            $new_list = @()
            $add_selection = {
                If ($selectionBox.SelectedItem -ne $null) {
                    if ($global:consumables_list.Count -gt 1) {
                        $global:consumables_ids.Add($($global:consumables_list.Model)[$selectionBox.SelectedIndex])
                    }else{
                        $global:consumables_ids.Add($global:consumables_list.Model)
                    }
                    $selectionBox2.Items.Add($selectionBox.SelectedItem)
                }
            }
            $selection_add_button.add_click($add_selection)

            $remove_selection = {
                If ($selectionBox2.SelectedItem -ne $null) {
                    $global:consumables_ids.RemoveAt($($selectionBox2.SelectedIndex))
                    $selectionBox2.Items.Remove($selectionBox2.SelectedItem)
                }
            }
            $selection_remove_button.add_click($remove_selection)

            $search_ok_button = New-Object System.Windows.Forms.Button
            $search_ok_button.Location = New-Object System.Drawing.Point(347, $($($selectionBox2.Location.Y) + $($selectionBox2.Size.Height) + 8))
            $search_ok_button.Size = New-Object System.Drawing.Size(75,23)
            $search_ok_button.Text = 'Confirm'
            $search_ok_button.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $search_form.Controls.Add($search_ok_button)

            $search_accept = {
                $search_form.AcceptButton = $search_button
                $selectionBox.SelectedItem = $null
            }
            $add_accept = {
                $search_form.AcceptButton = $selection_add_button
            }
            $remove_accept = {
                $search_form.AcceptButton = $selection_remove_button
            }
            $form_accept = {
                $search_form.AcceptButton = $search_ok_button
            }

            $search_bar.add_MouseDown($search_accept)
            $selectionBox.add_MouseDown($add_accept)
            $selectionBox2.add_MouseDown($remove_accept)

            $peripherals_label = New-Object System.Windows.Forms.Label
            $peripherals_label.Location = New-Object System.Drawing.Point($($selectionBox.Location.X + $selectionBox.Size.Width + 10),8)
            $peripherals_label.Size = New-Object System.Drawing.Size(300,20)
            $peripherals_label.Text = "$($ComputerName)'s currently connected peripheral(s):"
            $search_form.Controls.Add($peripherals_label)

            $peripherals = Get-Peripherals -ComputerName $ComputerName

            $peripherals_box = New-Object System.Windows.Forms.TextBox
            $peripherals_box.Location = New-Object System.Drawing.Point($($peripherals_label.Location.X),28)
            $peripherals_box.AutoSize = $true
            $peripherals_box.MinimumSize = New-Object System.Drawing.Size(417,260)
            $peripherals_box.MaximumSize = New-Object System.Drawing.Size(0,260)
            $peripherals_box.ScrollBars = 'Vertical'
            $peripherals_box.Multiline = $true
            $peripherals_box.Text = $($peripherals | Out-String).Trim()
            $peripherals_box.ReadOnly = $true
            $search_form.Controls.Add($peripherals_box)

            $monitors_label = New-Object System.Windows.Forms.Label
            $monitors_label.Location = New-Object System.Drawing.Point($($peripherals_label.Location.X),$($peripherals_box.Location.Y + $peripherals_box.Size.Height + 12))
            $monitors_label.Size = New-Object System.Drawing.Size(300,20)
            $monitors_label.Text = "$($ComputerName)'s currently connected moniotor(s):"
            $search_form.Controls.Add($monitors_label)

            $monitors = Get-Monitors -ComputerName $ComputerName

            $monitors_box = New-Object System.Windows.Forms.TextBox
            $monitors_box.Location = New-Object System.Drawing.Point($($peripherals_label.Location.X),$($monitors_label.Location.Y + $monitors_label.Size.Height))
            $monitors_box.AutoSize = $true
            $monitors_box.MinimumSize = New-Object System.Drawing.Size(417,248)
            $monitors_box.MaximumSize = New-Object System.Drawing.Size(0,248)
            $monitors_box.ScrollBars = 'Vertical'
            $monitors_box.Multiline = $true
            $monitors_box.Text = $($monitors | Out-String).Trim()
            $monitors_box.ReadOnly = $true
            $search_form.Controls.Add($monitors_box)

            $search_confirm = $search_form.ShowDialog()

            if ($search_confirm -eq [System.Windows.Forms.DialogResult]::OK) {
                $consumables_box.Text = $($($($selectionBox2.Items) | Out-String).Trim())
            }            
        }

        #Confirm Button
        $OKButton = New-Object System.Windows.Forms.Button
        $OKButton.Location = New-Object System.Drawing.Point($($($parent_serial_search.Location.X) + $($parent_serial_search.Size.Width) - 80), $($($consumables_box.Location.Y) + $($consumables_box.Size.Height) + 10))
        $OKButton.Size = New-Object System.Drawing.Size(75,23)
        $OKButton.Text = 'Confirm'
        $recent_form.Controls.Add($OKButton)

        #Don't Retire Button
        If ($Swap) {
            $RetireButton = New-Object System.Windows.Forms.Button
            $RetireButton.Location = New-Object System.Drawing.Point($($OKButton.Location.X - $RetireButton.Size.Width - 2), $OKButton.Location.Y)
            $RetireButton.Size = New-Object System.Drawing.Size(75,23)
            $RetireButton.Text = 'Skip Retire'
            $recent_form.Controls.Add($RetireButton)

            $skip_retire = {
                $answer = [System.Windows.Forms.MessageBox]::Show("Are you sure you don't want to retire the old computer?                      (This should only be done if the computer is in-warranty and/or will be reused)" , "Do Not Retire Old Computer" , 4)
                If ($answer -eq "Yes") {
                    $global:install_status_selection = "Skip Retirement"
                    [System.Windows.Forms.MessageBox]::Show("Old Computer will be returned to stock","Skip Retirement")
                }else{
                    $global:install_status_selection = "1"
                }
            }
            $RetireButton.add_click($skip_retire)
        }

        $assigned_to_accept_trigger = {
            $recent_form.AcceptButton = $assigned_to_search
        }
        $department_accept_trigger = {
            $recent_form.AcceptButton = $responsible_department_search
        }
        $parent_accept_trigger = {
            $recent_form.AcceptButton = $parent_serial_search
        }
        $location_accept_trigger = {
            $recent_form.AcceptButton = $location_search
        }
        $recent_form_accept = {
            $recent_form.AcceptButton = $OKButton
        }

        $assigned_to_box.add_MouseDown($assigned_to_accept_trigger)
        $responsible_department_box.add_MouseDown($department_accept_trigger)
        $parent_box.add_MouseDown($parent_accept_trigger)
        $location_box.add_MouseDown($location_accept_trigger)
        $consumables_box.add_click($consumables_click)
        $recent_form.add_MouseDown($recent_form_accept)
        $ComputerName_box.add_MouseDown($recent_form_accept)

        $function_changed = {
            if ($functionBox.SelectedItem -eq "Dedicated" -or $functionBox.SelectedItem -eq "Loaner") {
                $responsible_department_box.ReadOnly = $true
                $responsible_department_box.Text = $null
                $assigned_to_box.ReadOnly = $false
                $responsible_department_search.Visible = $false
                $assigned_to_search.Visible = $true
            }
            if ($functionBox.SelectedItem -eq "Kiosk" -or $functionBox.SelectedItem -eq "Shared Generic" -or $functionBox.SelectedItem -eq "Shared Non Generic" -or $functionBox.SelectedItem -eq "Test Machine" -or $functionBox.SelectedItem -eq "Vendor Machine") {
                $assigned_to_box.ReadOnly = $true
                $assigned_to_box.Text = $null
                $responsible_department_box.ReadOnly = $false
                $assigned_to_search.Visible = $false
                $responsible_department_search.Visible = $true
            }
        }
        $functionBox.add_SelectedIndexChanged($function_changed)
        $recent_form.Controls.Add($functionBox)
        
        $confirmation_trigger = {
            $global:ComputerName_selection = $ComputerName_box.Text
            $global:Computer_PO = $PO_box.Text
            Add-Type -AssemblyName System.Windows.Forms
            Add-Type -AssemblyName System.Drawing

            $confirmation_form = New-Object System.Windows.Forms.Form
            $confirmation_form.Text = "These Updates will be Added to the New Computer's Record"
            $confirmation_form.AutoSize = $true
            $confirmation_form.StartPosition = 'CenterScreen'

            $info_label1 = New-Object System.Windows.Forms.Label
            $info_label1.Location = New-Object System.Drawing.Point(0,0)
            $info_label1.AutoSize = $true
            $info_label1
            $info_label1.Text = "
Computer Name:

Install Status:

Assigned Date:

Installed Date:

Last Inventory By:

Inventory Date:

Comments:"
            $confirmation_form.Controls.Add($info_label1)

            If ($Swap) {
                $assigned_date = "$(Get-Date -Format yyyy-MM-dd) 08:00:00"
                $installed_date = $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
                $global:harddrive_serial = $(Get-WmiObject win32_physicalmedia -ComputerName $ComputerName).SerialNumber
            }else{
                $assigned_date = $null
                $installed_date = $null
            }
            $inventory_date = $(Get-Date -Format yyyy-MM-dd)

            #Header Information
            $info_label2 = New-Object System.Windows.Forms.Label
            $info_label2.Location = New-Object System.Drawing.Point($($($info_label1.Size.Width)),0)
            $info_label2.AutoSize = $true
            $info_label2.MaximumSize = New-Object System.Drawing.Point(255,0)
            $info_label2.Font = $font_bold
            $info_label2.Text = "
$ComputerName_selection

In use

$assigned_date

$installed_date

$ScriptUser

$inventory_date
"
            $confirmation_form.Controls.Add($info_label2)

            #Header Labels Column 2
            $info_label3 = New-Object System.Windows.Forms.Label
            $info_label3.Location = New-Object System.Drawing.Point (355,0)
            $info_label3.AutoSize = $true
            $info_label3.Text = "
Function:

Assigned To:

Responsible Department:

Department:

Parent Serial Number:

Location:

PO:

Consumables:"
            $confirmation_form.Controls.Add($info_label3)

            $function_selection = @()
            $assigned_to_selection = @()
            $responsible_department_selection = @()
            $department_selection = @()
            $parent_selection = @()
            $location_selection = @()
            $consumables_selection = @()
            if ($functionBox.Text -notmatch '\w') {
                $function_selection = $($recent_entry.function)
            }else{
                $function_selection = $($functionBox.SelectedItem)
            }
            if ($function_selection -eq "Dedicated" -or $function_selection -eq "Loaner") {
                Clear-Variable -Name responsible_department_selection -ErrorAction SilentlyContinue
                Clear-Variable -Name responsible_department_selection_id -ErrorAction SilentlyContinue
                if ($assigned_to_box.Text -notmatch '\w') {
                    $assigned_to_selection = $($recent_entry.assignment.assigned_to)
                    $assigned_to_selection_id = $($recent_entry.assignment.assigned_to_sys_id)
                }else{
                    $assigned_to_selection = $($assigned_to_box.Text)
                }
            }else{
                Clear-Variable -Name assigned_to_selection -ErrorAction SilentlyContinue
                Clear-Variable -Name assigned_to_selection_id -ErrorAction SilentlyContinue
                if ($responsible_department_box.Text -notmatch '\w') {
                    $responsible_department_selection = $($recent_entry.assignment.responsible_department)
                    $responsible_department_selection_id = $($recent_entry.assignment.responsible_department_sys_id)
                }else{
                    $responsible_department_selection = $($responsible_department_box.Text)
                }
            }
            if ($department_box.Text -notmatch '\w') {
                $department_selection = $($recent_entry.assignment.department)
                $department_selection_id = $($recent_entry.assignment.department_sys_id)
            }else{
                $department_selection = $($department_box.Text)
            }
            if ($parent_patch -ne $null) {
                $parent_selection = $null
                $global:parent_selection_id = $null
            }else{
                if ($parent_box.Text -notmatch '\w') {
                    $parent_selection = $($recent_entry.parent.serial_number)
                    $global:parent_selection_id = $($recent_entry.parent.sys_id)
                }else{
                    $parent_selection = $($parent_box.Text)
                }
            }
            if ($location_box.Text -notmatch '\w') {
                $location_selection = $($recent_entry.location)
                $location_selection_id = $($recent_entry.location_sys_id)
            }else{
                $location_selection = $($location_box.Text)
                $uri = "https://$instance_name/api/now/table/cmn_location?sysparm_query=name%3D$location_selection"
                $location_selection_id = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                $location_selection_id = $location_selection_id.result.sys_id
            }
            if ($consumables_box.Text -notmatch '\w') {
                $consumables_selection = $consumables_item.Text
                #Get Consumables sys-ids
                $uri = "https://$instance_name/api/now/table/alm_consumable?sysparm_query=parent%3D$($recent_entry.asset_id)"
                $comp_consumables_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                If ($parent_box.Text -match '\w') {
                    $uri = "https://$instance_name/api/now/table/alm_consumable?sysparm_query=parent%3D$($parent_selection_id)"
                    $parent_consumables_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                }else{
                    If ($($recent_entry.parent.serial_number) -match '\w') {
                        $uri = "https://$instance_name/api/now/table/alm_consumable?sysparm_query=parent%3D$($recent_entry.parent.sys_id)"
                        $parent_consumables_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                    }
                }
                foreach ($consumable_result in $comp_consumables_response.result) {
                    $global:consumables_ids += $consumable_result.model.value
                }
                If ($parent_consumables_response.result -ne $null) {
                    foreach ($consumable_result in $parent_consumables_response.result) { 
                        $global:consumables_ids += $consumable_result.model.value
                    }
                }
            }else{
                $consumables_selection = $consumables_box.Text
            }
            if ($parent_selection_id -eq $null -and $parent_selection -match '\w') {
                $parent_selection_id = $recent_entry.parent.sys_id
            }

            #Header Info Column 2
            $info_label4 = New-Object System.Windows.Forms.Label
            $info_label4.Location = New-Object System.Drawing.Point($($($info_label3.Location.X) + $($info_label3.Size.Width)),0)
            $info_label4.AutoSize = $true
            $info_label4.Font = $font_bold
            $info_label4.Text = "
$function_selection

$assigned_to_selection

$responsible_department_selection

$department_selection

$parent_selection

$location_selection

$Computer_PO

$consumables_selection"
            $confirmation_form.Controls.Add($info_label4)

            $confirmation_ok_button = New-Object System.Windows.Forms.Button
            $confirmation_ok_button.Location = New-Object System.Drawing.Point($($($confirmation_form.Size.Width) - 85),$($($confirmation_form.Size.Height) - 33))
            $confirmation_ok_button.Size = New-Object System.Drawing.Size(75,23)
            $confirmation_ok_button.Text = 'Confirm'

            If ($ComputerName_selection -eq $ComputerName) {
                $ComputerName_selection = $null
            }

            $final_confirmation = {
                #Save info to .csv file
                New-Object -TypeName PSCustomObject -Property @{
                    "Computer Name" = $ComputerName
                    "New Computer Name" = $ComputerName_selection
                    "Serial Number" = $global:SerialNumber
                    "Install Status" = $global:install_status_selection
                    "Assigned Date" = $assigned_date
                    "Installed Date" = $installed_date
                    "Last Inventory By" = $ScriptUser
                    "Last Inventory By ID" = $ScriptUser_ID
                    "Inventory Date" = $inventory_date
                    "Function" = $function_selection
                    "Assigned To" = $assigned_to_selection
                    "Assigned To ID" = $assigned_to_selection_id
                    "Responsible Department" = $responsible_department_selection
                    "Responsible Department ID" = $responsible_department_selection_id
                    "Department" = $department_selection
                    "Department ID" = $department_selection_id
                    "Parent Serial Number" = $parent_selection
                    "Parent ID" = $global:parent_selection_id
                    "Location" = $location_selection
                    "Location ID" = $location_selection_id
                    "PO" = $Computer_PO
                    "HardDrive" = $harddrive_serial
                    "Consumables" = $consumables_selection
                    "Consumables IDs" = $($($consumables_ids | Out-String).Trim())
                    "Comments" = $Comments_box.Text
                } | Select-Object "Computer Name","New Computer Name","Serial Number","Install Status","Assigned Date","Installed Date","Last Inventory By","Last Inventory By ID","Inventory Date","Function","Assigned To","Assigned To ID","Responsible Department","Responsible Department ID","Department","Department ID","Parent Serial Number","Parent ID","Location","Location ID","PO","HardDrive","Consumables","Consumables IDs","Comments" | Export-Csv -Path $SNAM_Path -NoTypeInformation -Append -Encoding ASCII
                if ($entry.Count -gt 1) {
                    $form.Close()
                }
                $confirmation_form.DialogResult = [System.Windows.Forms.DialogResult]::OK
                $recent_form.DialogResult = [System.Windows.Forms.DialogResult]::OK
            }

            $confirmation_ok_button.add_Click($final_confirmation)
            $confirmation_form.Controls.Add($confirmation_ok_button)
            $confirmation_form.AcceptButton = $confirmation_ok_button

            $confirmation_cancel_button = New-Object System.Windows.Forms.Button
            $confirmation_cancel_button.Location = New-Object System.Drawing.Point($($($confirmation_ok_button.Location.X) - 80),$($confirmation_ok_button.Location.Y))
            $confirmation_cancel_button.Size = New-Object System.Drawing.Size(75,23)
            $confirmation_cancel_button.Text = 'Cancel'
            $confirmation_cancel_button.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $confirmation_form.Controls.Add($confirmation_cancel_button)

            #Comments Box
            $Comments_box = New-Object System.Windows.Forms.TextBox
            $Comments_box.Location = New-Object System.Drawing.Point(3,$($info_label1.Size.Height))
            $Comments_box.AutoSize = $true
            $Comments_box.MinimumSize = New-Object System.Drawing.Point(345,100)
            $Comments_box.MaximumSize = New-Object System.Drawing.Point(345,0)
            $Comments_box.Multiline = $true
            $Comments_box.AcceptsReturn = $true
            If ($Inventory -and $assigned_to_selection -match '\w') {
                Clear-Variable -Name currentUserName -ErrorAction SilentlyContinue
                Clear-Variable -Name lastUser -ErrorAction SilentlyContinue
                Clear-Variable -Name currentUserFirst -ErrorAction SilentlyContinue
                Clear-Variable -Name currentUserLast -ErrorAction SilentlyContinue
                Clear-Variable -Name currentUser -ErrorAction SilentlyContinue
                $currentUserName = (Get-WmiObject -Class win32_computersystem -ComputerName $ComputerName -ErrorAction SilentlyContinue).UserName
                If ($currentUserName -eq $null) {
                    $currentUserName = (Get-WmiObject -Class win32_process -ComputerName $ComputerName -ErrorAction SilentlyContinue | Where-Object name -Match explorer -ErrorAction SilentlyContinue)
                    If ($currentUserName -ne $null) {
                        $currentUserName = $currentUserName.getowner().user
                    }
                }else{
                    $currentUserName = $currentUserName.Split("\")[1]
                }
                If ($currentUserName -ne $null) {
                    $currentUserFirst = (Get-ADUser $currentUserName).GivenName
                    $currentUserLast = (Get-ADUser $currentUserName).Surname
                    If ($currentUserLast -match "_") {
                        $currentUserLast = $currentUserLast.Split("_")[0]
                    }
                    $currentUser = "$currentUserFirst $currentUserLast"
                    If ($currentUser -ne $($recent_entry.assignment.assigned_to)) {
                        $Comments_box.Text = "Asset Assignment Needs Review. Last Logged On User: $currentUser"
                    }
                }
            }
            $confirmation_form.Controls.Add($Comments_box)

            $confirmation_form.ShowDialog()
        }

        $OKButton.add_click($confirmation_trigger)
        
        If ($entry.Count -gt 1) {
        #Create Other Results Display Box
            Clear-Variable -Name number_i -ErrorAction SilentlyContinue
            $otherlabel1 = @{}
            $otherlabel2 = @{}
            $otherlabel3 = @{}
            $otherlabel4 = @{}
            $otherlabel5 = @{}
            $otherlabel6 = @{}
            $other_consumables_item = @{}

            Add-Type -AssemblyName System.Windows.Forms
            Add-Type -AssemblyName System.Drawing

            $form = New-Object System.Windows.Forms.Form
            $form.Text = "Other, Older Results"
            $form.AutoSize = $true
            $form.StartPosition = 'Manual'
            $form.Location = New-Object System.Drawing.Point(1295,5)

            $font_bold = New-Object System.Drawing.Font("Microsoft Sans Serif",8.25,[System.Drawing.FontStyle]::Bold)

            for ($i = 1; $i -le $total_results; $i++) {
                if ($entry[$i] -ne $recent_entry) {
                    if ($($entry[$i].install_status) -eq 1) {
                        $install_status = "In Use"
                    }else{
                        if ($($entry[$i].install_status) -eq 7) {
                            $install_status = "Retired"
                        }else{
                            if ($($entry[$i].install_status) -eq 2) {
                                $install_status = "On order"
                            }else{
                                if ($($entry[$i].install_status) -eq 6) {
                                    $install_status = "In Stock"
                                }else{
                                    if ($($entry[$i].install_status) -eq 8) {
                                        $install_status = "Missing"
                                    }else{
                                        $install_status = $($entry[$i].install_status)
                                    }
                                }
                            }
                        }
                    }

                    if ($number_i -eq $null) {
                        $start_height = 0
                    }

                    #Header Labels
                    $otherlabel1[$i] = New-Object System.Windows.Forms.Label
                    $otherlabel1[$i].Location = New-Object System.Drawing.Point(0,$start_height)
                    $otherlabel1[$i].AutoSize = $true
                    $otherlabel1[$i]
                    $otherlabel1[$i].Text = "
Computer Name:

Serial Number:

Model:

Install Status:

IP Address:

Comments:"
                    $form.Controls.Add($otherlabel1[$i])

                    #Header Information
                    $otherlabel2[$i] = New-Object System.Windows.Forms.Label
                    $otherlabel2[$i].Location = New-Object System.Drawing.Point($($($otherlabel1[$i].Size.Width)),$start_height)
                    $otherlabel2[$i].AutoSize = $true
                    $otherlabel2[$i].MaximumSize = New-Object System.Drawing.Point(255,0)
                    $otherlabel2[$i].Font = $font_bold
                    $otherlabel2[$i].Text = "
$($entry[$i].name)

$($entry[$i].serial_number)

$($entry[$i].model.model)

$install_status

$($entry[$i].ip)

$($entry[$i].comments)"
                    $form.Controls.Add($otherlabel2[$i])

                    #Header Labels Column 2
                    $otherlabel3[$i] = New-Object System.Windows.Forms.Label
                    $otherlabel3[$i].Location = New-Object System.Drawing.Point (355,$start_height)
                    $otherlabel3[$i].AutoSize = $true
                    $otherlabel3[$i].Text = "
Warranty Date:

Assigned Date:

Installed Date:

Last Inventory By:

Inventory Date:

Last Discovery Date:

Disovered By:"
                    $form.Controls.Add($otherlabel3[$i])

                    #Header Info Column 2
                    $otherlabel4[$i] = New-Object System.Windows.Forms.Label
                    $otherlabel4[$i].Location = New-Object System.Drawing.Point($($($otherlabel3[$i].Location.X) + $($otherlabel3[$i].Size.Width)),$start_height)
                    $otherlabel4[$i].AutoSize = $true
                    $otherlabel4[$i].Font = $font_bold
                    $otherlabel4[$i].Text = "
$($entry[$i].warranty)

$($entry[$i].assigned_date)

$($entry[$i].installed_date)

$ScriptUser

$($entry[$i].inventory_date)

$($entry[$i].discovery.date)

$($entry[$i].disovery.source)"
                    $form.Controls.Add($otherlabel4[$i])

                    If ($($entry[$i].comments) -match '\w') {
                        $functionlabel_y = $($($otherlabel2[$i].Location.Y) + $($otherlabel2[$i].Size.Height))
                    }else{
                        $functionlabel_y = $($($otherlabel3[$i].Location.Y) + $($otherlabel3[$i].Size.Height))
                    }

                    #Header Labels Column 3
                    $otherlabel5[$i] = New-Object System.Windows.Forms.Label
                    $otherlabel5[$i].Location = New-Object System.Drawing.Point (0,$($functionlabel_y))
                    $otherlabel5[$i].AutoSize = $true
                    $otherlabel5[$i].Text = "
Function:

Assigned To:

Responsible Department:

Department:

Parent Serial:

Location:

Consumables:"
                    $form.Controls.Add($otherlabel5[$i])

                    #Header Info Column 3
                    $otherlabel6[$i] = New-Object System.Windows.Forms.Label
                    $otherlabel6[$i].Location = New-Object System.Drawing.Point($($($otherlabel5[$i].Location.X) + $($otherlabel5[$i].Size.Width)),$($otherlabel5[$i].Location.Y))
                    $otherlabel6[$i].AutoSize = $true
                    $otherlabel6[$i].Font = $font_bold
                    $otherlabel6[$i].Text = "
$($entry[$i].function)

$($entry[$i].assignment.assigned_to)

$($entry[$i].assignment.responsible_department)

$($entry[$i].assignment.department)

$($entry[$i].parent.serial_number)

$($entry[$i].discovery.location)"
                    $form.Controls.Add($otherlabel6[$i])

                    #Consumables Item
                    $other_consumables_item[$i] = New-Object System.Windows.Forms.Label
                    $other_consumables_item[$i].Location = New-Object System.Drawing.Point(0,$($($otherlabel5[$i].Location.Y) + $($otherlabel5[$i].Size.Height)))
                    $other_consumables_item[$i].Autosize = $true
                    $other_consumables_item[$i].MaximumSize = New-Object System.Drawing.Point(325,0)
                    $other_consumables_item[$i].Text = "$($($($entry[$i].consumables.name) | Out-String).Trim())"
                    $other_consumables_item[$i].Font = $font_bold
                    $form.Controls.Add($other_consumables_item[$i])

                    $global:start_height = $($($other_consumables_item[$i].Location.Y) + $($other_consumables_item[$i].Size.Height) + 10)
                    $global:number_i = $i
                }
            }

            $form.Show()
        }

        $recent_form_result = $recent_form.ShowDialog()
    }

    End {

    }
}

#helper functions
function Edit-SNAMLocation {
    <#
    .Synopsis
        Prompts user for location information and updates Monday.com with input.
    .Description
        This Function opens a series of selection boxes that allows the user to 
        enter the ServiceNow Location information of a specified computer. The 
        results are imported into the Location column of the Working Board on 
        Monday.com for the specified computer.
    .Parameter out
        None
    .Example
        PS> Edit-SNAMLocation -SerialNumber SERIALNUMBER
    .Notes
        Author: Jacob Searcy
        Date: 5/25/2021
        Update: 7/20/2021
    #>

    [CmdletBinding()]

    Param (
        [String]$ComputerName,
        [String]$SerialNumber,
        [String]$user,
        [String]$pass,
        [String]$instance_name
    )

    Begin {
        $date_path = '$PSScriptRoot\Files\SNAM date.txt'
        if (-not($SerialNumber)) {
            If (Test-Connection $ComputerName -Count 1 -Quiet) {
                $SerialNumber = (Get-WmiObject Win32_BIOS -ComputerName $ComputerName).SerialNumber
            }else{
                Write-Host "$ComputerName is Offline" -ForegroundColor Red
                Write-Host $null
                $answer = "No"
                While ($answer -eq "No") {
                    $SerialNumber = Read-Host "Manually Enter Serial Number"
                    $answer = [System.Windows.Forms.MessageBox]::Show("You Entered: $SerialNumber Is that correct?" , "Confirm Serial Number" , 4)
                }
            }
        }
        #Create Loading Box
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing

        $loading_form = New-Object System.Windows.Forms.Form
        $loading_form.Text = 'SNAM Location Box'
        $loading_form.Size = New-Object System.Drawing.Size(600,455)
        $loading_form.StartPosition = 'CenterScreen'

        $font_body = New-Object System.Drawing.Font("Microsoft Sans Serif",10,[System.Drawing.FontStyle]::Regular)
        $font_body_bold = New-Object System.Drawing.Font("Microsoft Sans Serif",10,[System.Drawing.FontStyle]::Bold)
        $font_header = New-Object System.Drawing.Font("Microsoft Sans Serif",15,[System.Drawing.FontStyle]::Bold)

        $loading1 = New-Object System.Windows.Forms.Label
        $loading1.Location = New-Object System.Drawing.Point(10,15)
        $loading1.AutoSize = $true
        $loading1.Text = 'Querying ServiceNow for'
        $loading1.Font = $font_body

        $loading2 = New-Object System.Windows.Forms.Label
        $loading2.Location = New-Object System.Drawing.Point(165,15)
        $loading2.AutoSize = $true
        $loading2.Text = "$SerialNumber . . ."
        $loading2.Font = $font_body_bold

        $loading3 = New-Object System.Windows.Forms.Label
        $loading3.Location = New-Object System.Drawing.Point(10,40)
        $loading3.AutoSize = $true
        $loading3.Text = 'Building ServiceNow Location Selection Boxes . . .'
        $loading3.Font = $font_body

        $instructions1 = New-Object System.Windows.Forms.Label
        $instructions1.Location = New-Object System.Drawing.Point(230,90)
        $instructions1.AutoSize = $true
        $instructions1.Text = 'Instructions'
        $instructions1.Font = $font_header

        $instructions2 = New-Object System.Windows.Forms.Label
        $instructions2.Location = New-Object System.Drawing.Point(30,125)
        $instructions2.AutoSize = $true
        $instructions2.Text = 'Use the drop-down menus to select a location. Once you find the location you need'
        $instructions2.Font = $font_body

        $instructions3 = New-Object System.Windows.Forms.Label
        $instructions3.Location = New-Object System.Drawing.Point(30,150)
        $instructions3.AutoSize = $true
        $instructions3.Text = 'Select'
        $instructions3.Font = $font_body

        $instructions4 = New-Object System.Windows.Forms.Label
        $instructions4.Location = New-Object System.Drawing.Point(70,150)
        $instructions4.AutoSize = $true
        $instructions4.Text = 'Confirm Selection'
        $instructions4.Font = $font_body
        $instructions4.ForeColor = "blue"

        $instructions5 = New-Object System.Windows.Forms.Label
        $instructions5.Location = New-Object System.Drawing.Point(180,150)
        $instructions5.AutoSize = $true
        $instructions5.Text = 'from the box that follows the box with the location you need.'
        $instructions5.Font = $font_body

        $loading_form.Controls.Add($loading1)
        $loading_form.Controls.Add($loading2)
        $loading_form.Controls.Add($loading3)
        $loading_form.Controls.Add($instructions1)
        $loading_form.Controls.Add($instructions2)
        $loading_form.Controls.Add($instructions3)
        $loading_form.Controls.Add($instructions4)
        $loading_form.Controls.Add($instructions5)

        $loading_form.Topmost = $true

        $loading_form.Show()
    }

    Process {
        Start-Sleep -Seconds 1
        #SNAM API
        $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $user, $pass)))
        $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
        $headers.Add('Authorization',('Basic {0}' -f $base64AuthInfo))
        $headers.Add('Accept','application/json')
        $method = "get"
        #Hardware Table
        $uri = "https://$instance_name/api/now/table/alm_hardware?sysparm_query=serial_number%3D$SerialNumber&sysparm_limit=1"
        $global:comp_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
        $location_id = $comp_response.result.location.value
        #Location Record
        $uri = "https://$instance_name/api/now/table/cmn_location/$location_id"
        $current_location_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
        $current_location = $current_location_response.result.name

        #Locations Table
        $uri = "https://$instance_name/api/now/table/cmn_location?sysparm_fields=name%2Cparent%2Csys_id%2Csys_updated_on"
        $global:response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri

        $last_updated_date = ($response.result.sys_updated_on | Sort-Object -Descending)[0]
        $script_date = (Get-Content -Path $date_path)[0]

        If ($last_updated_date -gt $script_date) {
            $tier1 = $response.result | Where-Object {$_.parent.value -eq $null}
            $tier1_name = ($response.result | Where-Object {$_.parent.value -eq $null}).name
            $tier1_id = ($response.result | Where-Object {$_.parent.value -eq $null}).sys_id

            $count = 0
            $continue = "Yes"

            While ($continue -eq "Yes") {
                if ($tier1_id.Count -gt 1) {
                    $tier2 = foreach ($item in $tier1_id) {$response.result | Where-Object {$_.parent.value -eq $item}}
                }else{
                    if ($tier1_id.Count -eq 1) {
                        $tier2 = $response.result | Where-Object {$_.parent.value -eq $tier1_id}
                    }else{
                        $continue = "No"
                    }
                }
                if ($tier2 -ne $null) {
                    $tier1 = $tier2
                    $tier1_name = $tier1.name
                    $tier1_id = $tier1.sys_id
                    $count = $count + 1
                    Clear-Variable -Name tier2 -ErrorAction SilentlyContinue
                }else{
                    $continue = "No"
                }
            }

            $box_count = $count
            Get-Date -Format "yyyy-MM-dd HH:mm:ss" | Out-File -FilePath '$PSScriptRoot\Files\SNAM date.txt'
            $box_count | Out-File -FilePath '$PSScriptRoot\Files\SNAM date.txt' -Append
        }else{
            [int]$box_count = (Get-Content -Path $date_path)[1]
        }

        #Locations Table
        $uri = "https://$instance_name/api/now/table/cmn_location?sysparm_fields=name%2Cparent%2Csys_id"
        $global:response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri

        $tier1 = $response.result | Where-Object {$_.parent.value -eq $null}
        $tier2 = $response.result | Where-Object {$_.parent.value -eq $tier1.sys_id}

        $loading_form.Close()

        #Create Drop-Down Box
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing

        $form = New-Object System.Windows.Forms.Form
        $form.Text = 'SNAM Location Box'
        $form.Size = New-Object System.Drawing.Size(600,$($box_count * 65))
        $form.StartPosition = 'CenterScreen'

        $font_body = New-Object System.Drawing.Font("Microsoft Sans Serif",10,[System.Drawing.FontStyle]::Regular)
        $font_body_bold = New-Object System.Drawing.Font("Microsoft Sans Serif",10,[System.Drawing.FontStyle]::Bold)
        $font_header = New-Object System.Drawing.Font("Microsoft Sans Serif",15,[System.Drawing.FontStyle]::Bold)

        $header = New-Object System.Windows.Forms.Label
        $header.Location = New-Object System.Drawing.Point(115,15)
        $header.Size = New-Object System.Drawing.Size(480,20)
        $header.Text = 'Use the drop-down menus to enter ServiceNow Location information'

        $OKButton = New-Object System.Windows.Forms.Button
        $OKButton.Location = New-Object System.Drawing.Point(505,$(($box_count * 65) - 75))
        $OKButton.Size = New-Object System.Drawing.Size(75,23)
        $OKButton.Text = 'Cancel'
        $OKButton.add_Click($handler_OK_Button_Click)
        $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.AcceptButton = $OKButton
        $OKButton.Visible = $false

        $previouslabel = New-Object System.Windows.Forms.Label
        $previouslabel.Location = New-Object System.Drawing.Point(175,$(((($box_count * 65) - 75) - 100) - 22))
        $previouslabel.Size = New-Object System.Drawing.Size(450,20)
        $previouslabel.Text = 'Current SNAM Location. Click to Select it.'

        $previousBox = New-Object System.Windows.Forms.Button
        $previousBox.Location = New-Object System.Drawing.Size(10,$((($box_count * 65) - 75) - 100))
        $previousBox.Size = New-Object System.Drawing.Size(560,85)
        $previousBox.Text = $current_location
        #$previousBox.DialogResult = [System.Windows.Forms.DialogResult]::Yes

        $previous_selection = {
            $global:SNAMLocation = $previousBox.Text
            $confirm = [System.Windows.Forms.MessageBox]::Show("You have selected $SNAMLocation. Confirm?" , "Confirm Location" , 4)
            If ($confirm -eq "Yes") {
                $global:Selected_Location = $SNAMLocation
                $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
            }
        }
        $previousbox.add_click($previous_selection)

        $form.Topmost = $true

        $form.Controls.Add($header)
        $form.Controls.Add($OKButton)
        $form.Controls.Add($previouslabel)
        $form.Controls.Add($previousBox)

        #Drop-Down Box Builder
        $box = @{}
        $status = @{}
        for ($i = 1; $i -le $box_count; $i++) {
            $box[$i] = New-Object System.Windows.Forms.ComboBox
            $box[$i].Location = New-Object System.Drawing.Size(10,$(($i * 30) + 10))
            $box[$i].Size = New-Object System.Drawing.Size(560,30)
            $box[$i].DropDownHeight = 500
            $box[$i].Name = $i
            $form.Controls.Add($box[$i])
            $box[$i].Height = 875
            $status[$i] = $box[$i].SelectedIndex

            $SelectedIndexChanged = {
                Clear-Variable -Name number -ErrorAction SilentlyContinue
                for ($k = $box_count; $k -ge 1; $k--) {
                    if ($box[$k].SelectedIndex -ne $status[$k]) {
                        $selected_item = $box[$k].SelectedItem
                        $status[$k] = $box[$k].SelectedIndex
                        $number = $k
                        $k = 0
                    }else{
                        $box[$k].Items.Clear()
                        $box[$k].Text = $null
                        $status[$k] = $box[$k].SelectedIndex
                    }
                }
                if ($selected_item -eq "Confirm Selection") {
                    $check = "No"
                }else{
                    Clear-Variable -Name tier2ID -ErrorAction SilentlyContinue
                    Clear-Variable -Name tier3 -ErrorAction SilentlyContinue
                    $tier2ID = $($response.result | Where-Object {$_.name -eq $selected_item}).sys_id
                    $tier3 = $response.result | Where-Object {$_.parent.value -eq $tier2ID}
                    If ($tier3 -eq $null) {
                        $check = "No"
                    }else{
                        $range = $tier3.name | Sort-Object
                        $confirm = "Confirm Selection"
                        $range = echo $range $confirm
                        $check = "Yes"
                    }
                }

                If ($check -ne "Yes") {
                    If ($selected_item -eq "Confirm Selection") {
                        If ($number -eq 1) {
                            $global:SNAMLocation = "North Carolina"
                        }else{
                            $jj = $number - 1
                            $global:SNAMLocation = $box[$jj].SelectedItem
                        }
                    }else{
                        $global:SNAMLocation = $selected_item
                    }
                    $confirm = [System.Windows.Forms.MessageBox]::Show("You have selected $SNAMLocation. Confirm?" , "Confirm Location" , 4)
                    If ($confirm -eq "Yes") {
                        $global:Selected_Location = $SNAMLocation
                        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
                    }else{
                        $box[$number].DroppedDown = $true
                    }
                }


                if ($check -eq "Yes") {
                    for($j = $($number + 1); $j -le $box_count; $j++) {
                        $box[$j].Items.Clear()
                    }
                    $j = $number + 1
                    $box[$j].Items.AddRange(@($range))
                    $box[$j].DroppedDown = $true
                }
            }

            $box[$i].Add_SelectedIndexChanged($SelectedIndexChanged)

            if ($i -eq 1) {
                $range = $tier2.name | Sort-Object
                $confirm = "Confirm Selection"
                $range = echo $range $confirm
                $box[$i].Items.AddRange(@($range))
            }
        }

        $form.ShowDialog()
    }

    End{}
}

function Get-Peripherals{
    <#
    .Synopsis
        Retrieves the peripherals connected to the specified computer.
    .Description
        This function retrieves the connected peripherals/installed drivers of a specified computer from Device Manager.
        It formats commonly used drivers into recognizable names.
    .Parameter out
        [Array] of [Strings]
    .Example
        PS> Get-Peripherals -ComputerName COMPUTERNAME01
        Badge Reader
        Honeywell Scanner
        Signature Pad
    .Notes
        Author: Jacob Searcy
    #>

[CmdletBinding()]

    Param(
        [Parameter(mandatory=$true,ValueFromPipeline=$true)]
        [String]$ComputerName,
        [String]$filterpath = "$PSScriptRoot\Files\exclude_peripherals.txt",
        [String]$peripheralswaplistpath = "$PSScriptRoot\Files\peripheralswaplist.csv",
        [String]$sigpad = 'USB\VID_0403&PID_6001\TOPAZBSB',
        [String]$BadgeReader1 = 'HID\VID_0C27&PID_3BFA',
        [String]$BadgeReader2 = 'USB\VID_0C27&PID_3BFA'
    )

    Begin{
        $ErrorActionPreference = "SilentlyContinue"
        $error.Clear()
    }

    Process{
        if(Test-Connection $ComputerName -count 1 -quiet){
            $NameCheck = Get-WmiObject Win32_USBControllerDevice -ComputerName $ComputerName |%{[wmi]($_.Dependent)}
            Foreach ($item in $NameCheck) {
                If ($item.DeviceID -contains $sigpad) {
                    $item = 'Signature Pad'
                    $peripherals = echo $peripherals $item
                }
                If ($item.HardwareID -contains $BadgeReader1 -or $item.HardwareID -contains $BadgeReader2) {
                    $item = 'Badge Reader'
                    $peripherals = echo $peripherals $item
                }else{
                    $item = $item.Name
                    $peripherals = echo $peripherals $item
                }
            }
            $peripherals = $peripherals.Split([Environment]::NewLine)
            $peripherals = $peripherals.Split([Environment]::NewLine, [StringSplitOptions]::RemoveEmptyEntries)
            
            $peripheralswaphash = @{}
            Import-Csv -Path $peripheralswaplistpath | foreach {$peripheralswaphash.Add($_.Name, $_.Value)}
            foreach ($item in $peripherals) {
                if ($peripheralswaphash[$item] -ne $null) {
                    $peripheral = $peripheralswaphash[$item]
                    $newlist = echo $newlist $peripheral
                }else{
                    $newlist = echo $newlist $item
                }
            }
            $peripherals = $newlist
            Clear-Variable -Name newlist

            $peripherals = $peripherals | Select -Unique | Sort-Object

            $exclude = Get-Content $filterpath
            $peripheralList = $peripherals | Where {$exclude -notcontains $_}

            return $peripheralList
            Clear-Variable -Name peripherals -ErrorAction SilentlyContinue

        }else{
            return "Offline"
        }
    }

    End{
        If ($error -ne $null) {
            Add-Content -Path $errorlog -Value $null
            Add-Content -Path $errorlog -Value "=========================================================================================================="
            Add-Content -Path $errorlog -Value $errordate
            Add-Content -Path $errorlog -Value $null
            Add-Content -Path $errorlog -Value $error
        }    
    }
}

function Get-Monitors{
    <#
    .Synopsis
        Retrieves the monitors connected to the specified computer.
    .Description
        This function retrieves the connected monitors/installed drivers of a specified computer from Device Manager.
    .Parameter out
        [Array] of [Strings]
    .Example
        PS> Get-Monitors -ComputerName COMPUTERNAME01
        
    .Notes
        Author: Jacob Searcy
        Date: 4/29/2022
    #>

[CmdletBinding()]

    Param(
        [Parameter(mandatory=$true,ValueFromPipeline=$true)]
        [String]$ComputerName
    )

    Begin{
        $ErrorActionPreference = "SilentlyContinue"
        $error.Clear()
    }

    Process{
        If (Test-Connection $ComputerName -Count 1 -Quiet) {
            Get-WmiObject -Namespace root\wmi -Class WmiMonitorBasicDisplayParams -ComputerName $ComputerName | select InstanceName, @{ N="Horizonal"; E={[System.Math]::Round(($_.MaxHorizontalImageSize/2.54), 2)} }, @{ N="Vertical"; E={[System.Math]::Round(($_.MaxVerticalImageSize/2.54), 2)} }, @{N="Size"; E={[System.Math]::Round(([System.Math]::Sqrt([System.Math]::Pow($_.MaxHorizontalImageSize, 2) + [System.Math]::Pow($_.MaxVerticalImageSize, 2))/2.54),2)} }, @{N="Ratio";E={[System.Math]::Round(($_.MaxHorizontalImageSize)/($_.MaxVerticalImageSize),2)} }
        }else{
            return "Offline"
        }
    }

    End{}
}

function Update-SNAMInfo {
    <#
    .Synopsis
        
    .Description
        
    .Parameter out
        None
    .Example
        PS> Update-SNAMInfo -ComputerName COMPUTERNAME -SerialNumber SERIALNUMBER (-Swap) (-Inventory)
    .Notes
        Author: Jacob Searcy
        Date Started: 4/29/2022
        Date Finished: 7/14/2022
    #>

    [CmdletBinding()]

    Param (
        [String]$ComputerName,
        [String]$SerialNumber,
        [Switch]$Swap,
        [Switch]$Inventory,
        [String]$user,
        [String]$pass,
        [String]$instance_name
    )

    Begin {
        If (-not($SerialNumber)) {
            If ($Swap) {
                $SerialNumber = (Get-WmiObject Win32_BIOS).SerialNumber
            }else{
                If (Test-Connection $ComputerName -Count 1 -quiet) {
                    $SerialNumber = (Get-WmiObject Win32_BIOS -ComputerName $ComputerName).SerialNumber
                }else{
                    #Create Serial Number Entry Box
                    Add-Type -AssemblyName System.Windows.Forms
                    Add-Type -AssemblyName System.Drawing

                    $offline_form = New-Object System.Windows.Forms.Form
                    $offline_form.Size = New-Object System.Drawing.Size(250,160)
                    $offline_form.Text = "Offline"
                    $offline_form.StartPosition = 'CenterScreen'

                    $offline_label1 = New-Object System.Windows.Forms.Label
                    $offline_label1.Text = "$ComputerName is Offline"
                    $offline_label1.ForeColor = 'RED'
                    $offline_label1.Location = New-Object System.Drawing.Point(5,8)
                    $offline_label1.AutoSize = $true
                    $offline_form.Controls.Add($offline_label1)

                    $offline_label2 = New-Object System.Windows.Forms.Label
                    $offline_label2.Text = "Manually Enter Serial Number:"
                    $offline_label2.Location = New-Object System.Drawing.Point(5,32)
                    $offline_label2.Size = New-Object System.Drawing.Size(200,20)
                    $offline_form.Controls.Add($offline_label2)

                    $offline_box = New-Object System.Windows.Forms.TextBox
                    $offline_box.Location = New-Object System.Drawing.Point(5,58)
                    $offline_box.Size = New-Object System.Drawing.Size(223,21)
                    $offline_form.Controls.Add($offline_box)

                    $OKButton = New-Object System.Windows.Forms.Button
                    $OKButton.Location = New-Object System.Drawing.Point(155,90)
                    $OKButton.Size = New-Object System.Drawing.Size(75,23)
                    $OKButton.Text = "OK"
                    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
                    $offline_form.Controls.Add($OKButton)
                    $offline_form.AcceptButton = $OKButton

                    $CancelButton = New-Object System.Windows.Forms.Button
                    $CancelButton.Location = New-Object System.Drawing.Point(78,90)
                    $CancelButton.Size = New-Object System.Drawing.Size(75,23)
                    $CancelButton.Text = "Cancel"
                    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
                    $offline_form.Controls.Add($CancelButton)

                    $offline_form.TopMost = $true
                    $offline_result = $offline_form.ShowDialog()

                    If ($offline_result -eq [System.Windows.Forms.DialogResult]::OK) {
                        $global:SerialNumber = $offline_box.Text
                    }
                    If ($offline_result -eq [System.Windows.Forms.DialogResult]::Cancel) {
                        exit
                    }
                    $SerialNumber = $global:SerialNumber
                }
            }
        }
    #Import SNAM Updates
        $UserName = $((Get-WmiObject -Class win32_computersystem).UserName).Split("\")[1]
        $SNAM_Path = "$PSScriptRoot\Files\SNAM Updates.csv"
        If (-not($ComputerName)) {
            $Updates = Import-Csv -Path $SNAM_Path | Where-Object {$_."Serial Number" -eq $SerialNumber}
        }else{
            $Updates = Import-Csv -Path $SNAM_Path | Where-Object {$_."Computer Name" -eq $ComputerName}
            If ($Updates -eq $null) {
                $Updates = Import-Csv -Path $SNAM_Path | Where-Object {$_."New Computer Name" -eq $ComputerName}
            }
        }
        If ($Updates.Count -gt 1) {
            $number = $($test.Count - 1)
            $Updates = $Updates[$number]
        }
    #ServiceNow API GET Variables
        $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $user, $pass)))
        $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
        $headers.Add('Authorization',('Basic {0}' -f $base64AuthInfo))
        $headers.Add('Accept','application/json')
        $headers.Add('Content-Type','application/json')
        $method = "get"
    #GET Install Status Options
        $uri = "https://$instance_name/api/now/table/sys_choice?sysparm_query=name%3Dalm_hardware^element%3Dinstall_status"
        $StateOptions = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
    #Get Stockroom ID
        $uri = "https://$instance_name/api/now/table/alm_stockroom?sysparm_query=name%3DInventory Stockroom"
        $stockroom_id = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
        $stockroom_id = $stockroom_id.result.sys_id
    #Get IT sys_id
        $IT_depart = "IT - Helpdesk / Desktop"
        $uri = "https://$instance_name/api/now/table/cmn_department?sysparm_query=name%3D$IT_depart"
        $IT_department = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
        $IT_department_id = $($IT_department.result | Where-Object {$_.id -eq 8247}).sys_id

        $continue = "Yes"

        While ($continue -eq "Yes") {
        #GET SNAM Record(s) for $SerialNumber
            $uri = "https://$instance_name/api/now/table/alm_hardware?sysparm_query=serial_number%3D$SerialNumber"
            $serial_hw_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri

            If ($($serial_hw_response.result) -eq $null) {
                Add-Type -AssemblyName System.Windows.Forms
                Add-Type -AssemblyName System.Drawing

                $create_form = New-Object System.Windows.Forms.Form
                $create_form.Text = "No Results"
                $create_form.Size = New-Object System.Drawing.Size(377,200)
                $create_form.StartPosition = 'CenterScreen'
            
                $error_label = New-Object System.Windows.Forms.Label
                $error_label.Location = New-Object System.Drawing.Point(5,5)
                $error_label.AutoSize = $true
                $error_label.Text = "No Records found for Serial Number: $SerialNumber"
                $create_form.Controls.Add($error_label)

                $verify_label = New-Object System.Windows.Forms.Label
                $verify_label.Location = New-Object System.Drawing.Point(5,35)
                $verify_label.AutoSize = $true
                $verify_label.Text = "Please verify that the Computer Name and Serial Number are correct."
                $create_form.Controls.Add($verify_label)

                $Name_label = New-Object System.Windows.Forms.Label
                $Name_label.Location = New-Object System.Drawing.Point(5,65)
                $Name_label.AutoSize = $true
                $Name_label.Text = "Computer Name:"
                $create_form.Controls.Add($Name_label)

                $Serial_label = New-Object System.Windows.Forms.Label
                $Serial_label.Location = New-Object System.Drawing.Point(14,95)
                $Serial_label.AutoSize = $true
                $Serial_label.Text = "Serial Number:"
                $create_form.Controls.Add($Serial_label)

                $verify_Name_box = New-Object System.Windows.Forms.TextBox
                $verify_Name_box.Location = New-Object System.Drawing.Point($($Name_label.Location.X + $Name_label.Size.Width),$($Name_label.Location.Y - 1))
                $verify_Name_box.Size = New-Object System.Drawing.Size(255,25)
                $verify_Name_box.Text = $ComputerName
                $create_form.Controls.Add($verify_Name_box)

                $verify_Serial_box = New-Object System.Windows.Forms.TextBox
                $verify_Serial_box.Location = New-Object System.Drawing.Point($($Name_label.Location.X + $Name_label.Size.Width),$($Serial_label.Location.Y - 1))
                $verify_Serial_box.Size = New-Object System.Drawing.Size(255,25)
                $verify_Serial_box.Text = $SerialNumber
                $create_form.Controls.Add($verify_Serial_box)

                $OKButton = New-Object System.Windows.Forms.Button
                $OKButton.Location = New-Object System.Drawing.Point(256,125)
                $OKButton.Size = New-Object System.Drawing.Size(95,23)
                $OKButton.Text = "Search"
                $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
                $create_form.Controls.Add($OKButton)

                $NewButton = New-Object System.Windows.Forms.Button
                $NewButton.Location = New-Object System.Drawing.Point(158,125)
                $NewButton.Size = New-Object System.Drawing.Size(95,23)
                $NewButton.Text = "Create Record"
            
                $create_record = {
                    $global:ComputerName = $verify_Name_box.Text
                    $global:SerialNumber = $verify_Serial_box.Text
                    #Create Model Search Box
                    If (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet) {
                        $model = (Get-WmiObject -Class Win32_ComputerSystemProduct -ComputerName $ComputerName).Version
                    }else{
                    
                    }
                    $search = $model
                    $uri = "https://$instance_name/api/now/table/cmdb_model?sysparm_query=nameLIKE$search"
                    $search_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                    $model_list = @()
                    foreach ($search_result in $($search_response.result)) {
                        $row = "" | Select Name,SysID
                        $row.Name = $search_result.name
                        $row.SysID = $search_result.sys_id
                        $model_list += $row
                    }
                    $global:model_list = $model_list | Sort-Object -Property Name

                    Add-Type -AssemblyName System.Windows.Forms
                    Add-Type -AssemblyName System.Drawing

                    $search_form = New-Object System.Windows.Forms.Form
                    $search_form.Text = "Model Search:"
                    $search_form.AutoSize = $true
                    $search_form.StartPosition = 'CenterScreen'

                    $search_label = New-Object System.Windows.Forms.Label
                    $search_label.Text = "Search: (Model: $($model))"
                    $search_label.Location = New-Object System.Drawing.Point(5,8)
                    $search_label.Size = New-Object System.Drawing.Size(400,20)
                    $search_form.Controls.Add($search_label)

                    $results_label = New-Object System.Windows.Forms.Label
                    $results_label.Text = "Results:"
                    $results_label.Location = New-Object System.Drawing.Point(5,70)
                    $results_label.Size = New-Object System.Drawing.Size(50,21)
                    $search_form.Controls.Add($results_label)

                    $selectionBox = New-Object System.Windows.Forms.ListBox
                    $selectionBox.Location = New-Object System.Drawing.Point(5,91)
                    $selectionBox.AutoSize = $true
                    $selectionBox.MinimumSize = New-Object System.Drawing.Size(417,200)
                    $selectionBox.MaximumSize = New-Object System.Drawing.Size(0,200)
                    $selectionBox.ScrollAlwaysVisible = $true
                    $selectionBox.Items.Clear()
                    If ($model_list.name -notmatch '\w') {
                        $selectionBox.Items.Add("No Results Found")
                    }else{
                        foreach ($item in $model_list) {
                            [void] $selectionBox.Items.Add($item.Name)
                        }
                    }
                    $search_form.TopMost = $true
                    $search_form.Controls.Add($selectionBox)

                    $search_bar = New-Object System.Windows.Forms.TextBox
                    $search_bar.Location = New-Object System.Drawing.Size(5,28)
                    $search_bar.Size = New-Object System.Drawing.Size(400,21)
                    $search_form.Controls.Add($search_bar)
                    $search_bar.Text = $model

                    $search_button = New-Object System.Windows.Forms.Button
                    $search_button.Location = New-Object System.Drawing.Point(408,27)
                    $search_button.Size = New-Object System.Drawing.Size(21,21)
                    $search_button.BackgroundImage = $image
                    $search_button.BackgroundImageLayout = 'Zoom'

                    $search_trigger = {
                        $search = $search_bar.Text
                        $selectionBox.Items.Clear()
                        $selectionBox.Items.Add("Processing . . .")
                        $uri = "https://$instance_name/api/now/table/cmdb_model?sysparm_query=nameLIKE$search"
                        $search_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                        $model_list = @()
                        foreach ($search_result in $($search_response.result)) {
                            $row = "" | Select Name,Email,Department,SysID
                            $row.Name = $search_result.name
                            $row.SysID = $search_result.sys_id
                            $model_list += $row
                        }
                        $global:model_list = $model_list | Sort-Object -Property Name
                        $selectionBox.Items.Clear()
                        If ($model_list.name -notmatch '\w') {
                            $selectionBox.Items.Add("No Results Found")
                        }else{
                            foreach ($item in $model_list) {
                                [void] $selectionBox.Items.Add($item.Name)
                            }
                        }
                    }

                    $search_button.add_click($search_trigger)
                    $search_form.Controls.Add($search_button)

                    $search_ok_button = New-Object System.Windows.Forms.Button
                    $search_ok_button.Location = New-Object System.Drawing.Point(347, $($($selectionBox.Location.Y) + $($selectionBox.Size.Height) + 10))
                    $search_ok_button.Size = New-Object System.Drawing.Size(75,23)
                    $search_ok_button.Text = 'Confirm'
                    $search_ok_button.DialogResult = [System.Windows.Forms.DialogResult]::OK
                    $search_form.Controls.Add($search_ok_button)

                    $search_accept = {
                        $search_form.AcceptButton = $search_button
                        $selectionBox.SelectedItem = $null
                    }
                    $form_accept = {
                        $search_form.AcceptButton = $search_ok_button
                    }

                    $search_bar.add_MouseDown($search_accept)
                    $selectionBox.add_MouseDown($form_accept)

                    $search_confirm = $search_form.ShowDialog()

                    If ($search_confirm -eq [System.Windows.Forms.DialogResult]::OK) {
                        if ($model_list.Count -gt 1) {
                            $global:model_id = $model_list.SysID[$($selectionBox.SelectedIndex)]
                        }else{
                            $global:model_id = $model_list.SysID
                        }

                        $uri = "https://$instance_name/api/now/table/cmdb_model_category?sysparm_query=name%3DPC Hardware"
                        $model_category_id = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                        $model_category_id = $model_category_id.result.sys_id

                        $install_status = $($StateOptions.result | Where-Object {$_.label -eq "In stock"}).value

                        $uri = "https://$instance_name/api/now/table/alm_stockroom?sysparm_query=name%3DDesktop - Raleigh - Glass Room"
                        $stockroom_id = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                        $stockroom_id = $stockroom_id.result.sys_id

                        $uri = "https://$instance_name/api/now/table/cmn_department?sysparm_query=name%3DInformation Services"
                        $department_id = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                        $department_id = $department_id.result.sys_id
                        If ($department_id.Count -gt 1) {
                            $department_id = $department_id[0]
                        }

                        $method = "post"
                        $uri = "https://$instance_name/api/now/table/alm_hardware"
                        $body = "{`"serial_number`":`"$SerialNumber`",`"model_category`":`"$model_category_id`",`"model`":`"$model_id`",`"install_status`":`"$install_status`",`"stockroom`":`"$stockroom_id`",`"managed_by`":`"$department_id`"}"
                        $create_record_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body
                        $method = "get"
                        $create_form.Close()
                        Add-Type -AssemblyName System.Windows.Forms
                        Add-Type -AssemblyName System.Drawing

                        $font = New-Object System.Drawing.Font("Microsoft Sans Serif",10,[System.Drawing.FontStyle]::Regular)

                        $created_form = New-Object System.Windows.Forms.Form
                        $created_form.Text = "$SerialNumber"
                        $created_form.Size = New-Object System.Drawing.Size(250,100)
                        $created_form.StartPosition = 'CenterScreen'
            
                        $created_label = New-Object System.Windows.Forms.Label
                        $created_label.Location = New-Object System.Drawing.Point(69,8)
                        $created_label.AutoSize = $true
                        $created_label.Font = $font
                        $created_label.Text = "Record Created!"
                        $created_form.Controls.Add($created_label)

                        $OKButton = New-Object System.Windows.Forms.Button
                        $OKButton.Location = New-Object System.Drawing.Point(80,35)
                        $OKButton.Size = New-Object System.Drawing.Point(75,23)
                        $OKButton.Text = 'OK'
                        $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
                        $created_form.Controls.Add($OKButton)
                        $created_form.TopMost = $true
                        $created_form.ShowDialog()
                    }
                }

                $NewButton.add_click($create_record)
                $create_form.Controls.Add($NewButton)
                $create_form.TopMost = $true
                $create_results = $create_form.ShowDialog()

                If ($create_results -eq [System.Windows.Forms.DialogResult]::OK) {
                    $ComputerName = $verify_Name_box.Text
                    $SerialNumber = $verify_Serial_box.Text
                    $continue = "Yes"
                }

            }else{
                $continue = "No"
            }
        }

        If (-not($ComputerName)) {
            If ($Updates.'New Computer Name' -match '\w') {
                $ComputerName = $Updates.'New Computer Name'
            }else{
                $ComputerName = $Updates.'Computer Name'
            }
        }
<#
    #GET SNAM Record(s) for $ComputerName and NOT $SerialNumber
        $install_status_filter = $($StateOptions.result | Where-Object {$_.label -eq 'Retired'}).value
        $uri = "https://$instance_name/api/now/table/alm_hardware?sysparm_query=u_computer_name%3D$ComputerName^serial_number!=$SerialNumber^install_status!=$install_status_filter"
        $computername_hw_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
    #Change Name to Match Serial Number for Records with Wrong Name
        $check_serials = $computername_hw_response.result.serial_number
        If ($check_serials -ne $null) {
            If ($Swap) {
                foreach ($item in $check_serials) {
                    If ($item -eq $Updates.'Serial Number') {
                        $check_serials = $check_serials | Where-Object {$_ -ne $item} | Select-Object -Unique
                    }
                }
            }
            If ($check_serials -ne $null) {
                foreach ($item in $check_serials) {
                    $uri = "https://$instance_name/api/now/table/alm_hardware?sysparm_query=serial_number%3D$item"
                    $change_name = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                    $uri = $change_name.result.ci.link
                    $method = "patch"
                    $body_1 = '{"name":"'
                    $body_2 = '"}'
                    $body = "$body_1$item$body_2"
                    If ($uri.Count -gt 1) {
                        foreach ($thing in $uri) {
                            $update_name = Invoke-RestMethod -Headers $headers -Method $method -Uri $thing -Body $body
                        }
                    }else{
                        $update_name = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body
                    }
                }
            }
        }
        #>
        $method = "get"
    #GET Most Recently Discovered Asset
        $entry = @()
        $total_results = $($serial_hw_response.result.Count)
        for ($i = 1; $i -le $total_results; $i++) {
            #Hardware Table
            $comp_hw_response = $serial_hw_response.result[$i-1]
            #Hardware Table ID
            $comp_sys_id = $comp_hw_response.sys_id
            #Configuration Item ID
            If ($comp_hw_response.ci.value -eq $null) {
                $comp_ci_id = $null
                $comp_ci_link = $null
                $DiscoveryDate = $null
            }else{
                $comp_ci_id = $comp_hw_response.ci.value
                #Configuration Item Table
                $comp_ci_link = $comp_hw_response.ci.link
                $uri = $comp_ci_link
                $error.Clear()
                $comp_ci_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                If ($error) {
                    $comp_ci_id = $null
                    $comp_ci_link = $null
                    $DiscoveryDate = $null
                }else{
                    #Most Recent Discovery
                    $DiscoveryDate = $comp_ci_response.result.last_discovered
                }
            }
            $row = "" | Select SysID,Discovery
            $row.SysID = $comp_sys_id
            $row.Discovery = $DiscoveryDate
            $entry += $row
        }

        If ($total_results -gt 1) {
            $discovery_dates = @()
            $number = @()
            #Get Most Recently Discovered Item
            for ($i = 1; $i -le $total_results; $i++) {
                If ($($entry.discovery)[$i-1] -notmatch '\w') {
                    $discovery_dates += $null
                }else{
                    $discovery_dates += [DateTime]$($entry.discovery)[$i-1]
                }
                $number += $i
            }
            $discovery_date_list = @{
                number = $number
                date = $discovery_dates
            }
        
            $most_recent_date = $discovery_date_list | Foreach-Object {$_.date | Sort-Object {$_.date} | Select-Object -Last 1}

            $other_entries = @()
            for ($i = 1; $i -le $total_results; $i++) {
                If ($entry.discovery -match '\w' -and $($entry.discovery)[$i-1] -notmatch '\w') {
                    $other_entry = $($entry.SysID)[$i-1]
                    $other_entries += $($entry.serial_number)[$i-1]
                    #Add Duplicate Entry to Errors Log
                    $SNAM_Error_Path = "$PSScriptRoot\Files\SNAM Errors.csv"
                    $error.Clear()
                    $UserName = $((Get-WmiObject -Class win32_computersystem).UserName).Split("\")[1]
                    If ($error) {
                        $UserName = "sysklindstrom"
                    }
                    New-Object -TypeName PSCustomObject -Property @{
                        "User" = $UserName
                        "Computer ID" = $($entry.sys_id)[$i-1]
                        "Parent ID" = $null
                        "Serial Number" = $($entry.serial_number)[$i-1]
                        "Retired Serial Numbers" = $null
                        "Serials with Wrong Names" = $null
                        "Duplicate Serial Numbers" = $null
                        "Comments" = "This Serial Number contains a duplicate record"
                    } | Select-Object "User","Computer ID","Parent ID","Serial Number","Retired Serial Numbers","Serials with Wrong Names","Duplicate Serial Numbers","Comments" | Export-Csv -Path $SNAM_Error_Path -NoTypeInformation -Append -Encoding ASCII
                    #Warning Box
                    $answer = [System.Windows.Forms.MessageBox]::Show("$SerialNumber has a duplicate record. Would you like to mark it for deletion?", "Duplicate Record", 4)
                    If ($answer -eq "Yes") {
                        $uri = "https://$instance_name/api/now/table/alm_hardware/$other_entry"
                        $other_entry_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                        if ($($other_entry_response.result.po_number) -match '\w') {
                            $po_number = $other_entry_response.result.po_number
                        }
                        $method = "patch"
                        $on_order = $($StateOptions.result | Where-Object {$_.label -eq 'On order'}).value
                        $body = "{`"serial_number`":`"DELETE`",`"install_status`":`"$on_order`",`"managed_by`":`"$IT_department_id`"}"
                        $other_entry_update = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body
                        $uri = $other_entry_response.result.ci.link
                        $body = "{`"serial_number`":`"DELETE`"}"
                        $other_entry_ci = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body
                        $method = "get"
                    }
                }else{
                    If ([DateTime]$($entry.discovery)[$i-1] -eq $most_recent_date) {
                        $recent_entry = $($entry.SysID)[$i-1]
                    }else{
                    #Delete Duplicate Serials
                        $other_entry = $($entry.SysID)[$i-1]
                        $other_entries += $($entry.serial_number)[$i-1]
                        #Add Duplicate Entry to Errors Log
                        $SNAM_Error_Path = "$PSScriptRoot\Files\SNAM Errors.csv"
                        $error.Clear()
                        $UserName = $((Get-WmiObject -Class win32_computersystem).UserName).Split("\")[1]
                        If ($error) {
                            $UserName = "sysklindstrom"
                        }
                        New-Object -TypeName PSCustomObject -Property @{
                            "User" = $UserName
                            "Computer ID" = $($entry.sys_id)[$i-1]
                            "Parent ID" = $null
                            "Serial Number" = $($entry.serial_number)[$i-1]
                            "Retired Serial Numbers" = $null
                            "Serials with Wrong Names" = $null
                            "Duplicate Serial Numbers" = $null
                            "Comments" = "This Serial Number contains a duplicate record"
                        } | Select-Object "User","Computer ID","Parent ID","Serial Number","Retired Serial Numbers","Serials with Wrong Names","Duplicate Serial Numbers","Comments" | Export-Csv -Path $SNAM_Error_Path -NoTypeInformation -Append -Encoding ASCII
                        #Warning Box
                        $answer = [System.Windows.Forms.MessageBox]::Show("$SerialNumber has a duplicate record. Would you like to mark it for deletion?", "Duplicate Record", 4)
                        If ($answer -eq "Yes") {
                            $uri = "https://$instance_name/api/now/table/alm_hardware/$other_entry"
                            $other_entry_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                            if ($($other_entry_response.result.po_number) -match '\w') {
                                $po_number = $other_entry_response.result.po_number
                            }
                            $method = "patch"
                            $on_order = $($StateOptions.result | Where-Object {$_.label -eq 'On order'}).value
                            $body = "{`"serial_number`":`"DELETE`",`"install_status`":`"$on_order`",`"managed_by`":`"$IT_department_id`"}"
                            $other_entry_update = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body
                            $uri = $other_entry_response.result.ci.link
                            $body = "{`"serial_number`":`"DELETE`"}"
                            $other_entry_ci = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body
                            $method = "get"
                        }
                    }
                }
            }
        }else{
            $recent_entry = $entry.SysID
        }
        $uri = "https://$instance_name/api/now/table/alm_hardware/$recent_entry"
        $recent_entry_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
        If ($recent_entry_response.result.po -notmatch '\w') {
            foreach ($item in $other_entry) {
                $uri = "https://$instance_name/api/now/table/alm_hardware/$item"
                $other_entry_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                If ($($other_entry_response.result.po) -match '\w') {
                    $po_change = $($other_entry_response.result.po)
                    $method = 'patch'
                    $uri = "https://$instance_name/api/now/table/alm_hardware/$recent_entry"
                    $body = "{`"po`":`"$po_change`"}"
                    $po_update = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body
                    $method = 'get'
                }
            }
        }
    }

    Process {
        $body_2 = '"}'
    #Retire Old Computer
        If ($Swap) {
            $method = "get"
            $OldSerialNumber = $Updates.'Serial Number'
            $uri = "https://$instance_name/api/now/table/alm_hardware?sysparm_query=serial_number%3D$OldSerialNumber"
            $old_hw_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
        #GET Most Recently Discovered Asset
            $old_entry = @()
            $old_total_results = $($old_hw_response.result.Count)
            for ($i = 1; $i -le $old_total_results; $i++) {
                #Hardware Table
                $old_comp_hw_response = $old_hw_response.result[$i-1]
                #Hardware Table ID
                $old_comp_sys_id = $old_comp_hw_response.sys_id
                #Configuration Item ID
                $old_comp_ci_id = $old_comp_hw_response.ci.value
                #Configuration Item Table
                $old_comp_ci_link = $old_comp_hw_response.ci.link
                $uri = $old_comp_ci_link
                $old_comp_ci_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                #Most Recent Discovery
                $old_DiscoveryDate = $old_comp_ci_response.result.last_discovered
                #Add Result to List
                $row = "" | Select SysID,Discovery
                $row.SysID = $old_comp_sys_id
                $row.Discovery = $old_DiscoveryDate
                $old_entry += $row
            }
            #Compare Dates in List
            If ($old_total_results -gt 1) {
                $old_discovery_dates = @()
                $old_number = @()
                #Get Most Recently Discovered Item
                for ($i = 1; $i -le $old_total_results; $i++) {
                    $old_discovery_dates += [DateTime]$($old_entry.discovery)[$i-1]
                    $old_number += $i
                }
                $old_discovery_date_list = @{
                    number = $oldnumber
                    date = $old_discovery_dates
                }
                $old_most_recent_date = $old_discovery_date_list | Foreach-Object {$_.date | Sort-Object {$_.date} | Select-Object -Last 1}

                for ($i = 1; $i -le $old_total_results; $i++) {
                    If ([DateTime]$($old_entry.discovery)[$i-1] -eq $old_most_recent_date) {
                        $old_recent_entry = $($old_entry.SysID)[$i-1]
                    }else{
                        #Return All Consumables to Stock
                        $method = "post"
                        $old_other_entry = $($old_entry.SysID)[$i-1]
                        $uri = "https://$instance_name/api/wmh/consumables/returnAllToStock"
                        $body = "{`"parent`":`"$old_other_entry`",`"stockroom`":`"$stockroom_id`"}"
                        $consumables_stock_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body
                        $method = "get"
                        #Remove from Parent if Assigned to One
                        $uri = "https://$instance_name/api/now/table/alm_hardware/$old_other_entry"
                        $parent_check = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                        If ($parent_check.result.parent -ne $null) {
                            $method = "patch"
                            $body = '{"parent":""}'
                            $remove_old_parent = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body
                        }
                    }
                }
            }else{
                $old_recent_entry = $old_entry.SysID
            }
            #Transfer All Consumables to New Asset
            $method = "post"
            $uri = "https://$instance_name/api/wmh/consumables/parentSwap"
            if ($Updates.'Assigned To ID' -match '\w') {
                $new_user_id = $Updates.'Assigned To ID'
            }else{
                $new_user_id = $null
            }
            $body = "{`"newParent`":`"$recent_entry`",`"oldParent`":`"$old_recent_entry`",`"newUser`":`"$new_user_id`"}"
            $response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body
            $method = "patch"
            #Retire All Old Serials
            foreach ($item in $old_hw_response.result.sys_id) {
                #Remove from Parent if Assigned to One
                $uri = "https://$instance_name/api/now/table/alm_hardware/$item"
                $method = "get"
                $parent_check = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                $method = "patch"
                If ($parent_check.result.parent -ne $null) {
                    $method = "patch"
                    $body = '{"parent":""}'
                    $remove_old_parent = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body
                }
                #Retire Computer
                $uri = "https://$instance_name/api/now/table/alm_hardware/$item"
                $body_1 = '{"install_status":"'
                If ($Updates.'Install Status' -eq "1") {
                    $state_update = ($StateOptions.result | Where-Object {$_.label -eq "Retired"}).value
                    $body = "$body_1$state_update$body_2"
                #Return to Stock if Selected
                }else{
                    $state_update = ($StateOptions.result | Where-Object {$_.label -eq "In Stock"}).value
                    $inventory_by = $Updates.'Last Inventory By ID'
                    $date = Get-Date -Format yyyy-MM-dd
                    $body = "{`"install_status`":`"$state_update`",`"stockroom`":`"$stockroom_id`",`"u_last_inventory_by`":`"$inventory_by`",`"u_last_physical_inventory`":`"$date`"}"
                }
                $retire_update = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body
            }
        
            #Assign HardDrive
            $harddrive_serial = $Updates.HardDrive
            $method = "get"
            $uri = "https://$instance_name/api/now/table/cmdb_ci?sysparm_view=Harddrives&sysparm_query=serial_number%3D$OldSerialNumber^sys_class_name!=cmdb_ci_pc_hardware"
            $harddrive_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
            $harddrive_id = $harddrive_response.result.sys_id
            $method = "patch"
            If ($harddrive_id.Count -gt 1) {
                foreach ($item in $harddrive_id) {
                    $uri = "https://$instance_name/api/now/table/u_cmdb_ci_harddrives/$item"
                    #$body = "{`"u_hd_serial_number`":`"$harddrive_serial`"}"
                    $body = "{`"u_stockroom`":`"$stockroom_id`"}"
                    $harddrive_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body
                }
            }else{
                $uri = "https://$instance_name/api/now/table/u_cmdb_ci_harddrives/$harddrive_id"
                #$body = "{`"u_hd_serial_number`":`"$harddrive_serial`"}"
                $body = "{`"u_stockroom`":`"$stockroom_id`"}"
                $harddrive_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body
            }
        
        }

    #Update Record
        $method = "patch"
        $parent_removed = $false
        #Remove Parent if Needed
        If ($Updates.'Parent ID' -notmatch '\w' -and $recent_entry_response.result.parent -ne $null) {
            $uri = "https://$instance_name/api/now/table/alm_hardware/$recent_entry"
            $body = "{`"parent`":`"`"}"
            $remove_parent_patch = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body
            $parent_removed = $true
        }
        #Change Name
        If ($Updates.'New Computer Name' -match '\w' -or $($recent_entry_response.result.u_computer_name) -cne $($Updates.'Computer Name').ToUpper()) {
            $uri = $recent_entry_response.result.ci.link
            If ($Updates.'New Computer Name' -match '\w') {
                $NewName = $Updates.'New Computer Name'
            }else{
                $NewName = $Updates.'Computer Name'
            }
            $NewName = $NewName.ToUpper()
            $body_1 = '{"name":"'
            $body = "$body_1$NewName$body_2"
            $name_patch = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body
        }

        $uri = "https://$instance_name/api/now/table/alm_hardware/$recent_entry"
        #Change Install Status to "In Use"
        $body = "{`"install_status`":`"1`"}"
        $install_status_patch = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body
        
        #Change Function
        if ($Updates.Function -match '\w') {
            $function_update = $Updates.Function
            $body_1 = '{"u_device_function":"'
            $body ="$body_1$function_update$body_2"
            $function_patch = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body
        }

        #Change SubState
        If ($Updates.Function -eq "Dedicated") {
            $substate_update = "Assigned"
        }else{
            $substate_update = ""
        }
        $body_1 = '{"substatus":"'
        $body = "$body_1$substate_update$body_2"
        $substate_patch = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body

        #Clear Assigned To/Department
        If ($Updates.Function -eq "Dedicated" -or $Updates.Function -eq "Loaner") {
            $body = "{`"managed_by`":`"`",`"department`":`"`"}"
            $department_patch = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body
        }else{
            $body = "{`"assigned_to`":`"`"}"
            $assigned_patch = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body
        }

        #Change Assigned Date
        if ($Updates.'Assigned Date' -match '\w') {
            $assigned_date = $Updates.'Assigned Date'
            $body_1 = '{"assigned":"'
            $body ="$body_1$assigned_date$body_2"
            $assigned_date_patch = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body
        }
        #Change Installed Date
        if ($Updates.'Installed Date' -match '\w') {
            $installed_date = $Updates.'Installed Date'
            $body_1 = '{"install_date":"'
            $body ="$body_1$installed_date$body_2"
            $installed_date_patch = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body
        }
        #Change Inventory Date
        if ($Updates.'Inventory Date' -match '\w') {
            $inventory_date = $Updates.'Inventory Date'
            $body_1 = '{"u_last_physical_inventory":"'
            $body ="$body_1$inventory_date$body_2"
            $inventory_date_patch = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body
        }
        #Change Inventory By
        if ($Updates.'Last Inventory By ID' -match '\w') {
            $inventory_by = $Updates.'Last Inventory By ID'
            $body_1 = '{"u_last_inventory_by":"'
            $body = "$body_1$inventory_by$body_2"
            $inventory_by_patch = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body
        }
        #Change Assigned To
        if ($Updates.'Assigned To ID' -match '\w') {
            $assigned_to_id = $Updates.'Assigned To ID'
            $body_1 = '{"assigned_to":"'
            $body ="$body_1$assigned_to_id$body_2"
            $assigned_to_patch = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body
        }
        #Change Responsible Department
        if ($Updates.'Responsible Department ID' -match '\w') {
            $responsible_department_id = $Updates.'Responsible Department ID'
            $body_1 = '{"managed_by":"'
            $body ="$body_1$responsible_department_id$body_2"
            $responsible_department_patch = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body
            $method = "get"
            $uri = "https://$instance_name/api/now/table/alm_hardware/$recent_entry"
            $department_check = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
            $method = "patch"
            If ($department_check.result.department.value -ne $responsible_department_id) {
                $body_1 = '{"department":"'
                $body = "$body_1$responsible_department_id$body_2"
                $department_patch = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body
            }
        }
        #Change Location
        if ($Updates.'Location ID' -match '\w') {
            $location_id = $Updates.'Location ID'
            $body_1 = '{"location":"'
            $body ="$body_1$location_id$body_2"
            $parent_patch = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body
        }
        #Change Warranty Date
        $answer = Invoke-WebRequest -uri https://csp.lenovo.com/ibapp/il/WarrantyStatus.jsp -Method POST -Body @{serial="$SerialNumber"}
        $focustext = $answer.parsedhtml.body.innertext
        $regex = "[0-9]{4}-[0-9]{2}-[0-9]{2}"
        $expireDates = $focustext | select-string -Pattern $regex -AllMatches | % { $_.Matches } | % { $_.Value }
        $warrantyExpiration = ($expireDates | measure -Maximum).Maximum
        if ($warrantyExpiration -match '\w') {
            $body_1 = '{"warranty_expiration":"'
            $body = "$body_1$warrantyExpiration$body_2"
            $warranty_patch = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body
        }
        #Change Comments
        if ($Updates.Comments -match '\w') {
            $body_1 = '{"comments":"'
            $comments = $Updates.Comments
            If ($po_comment -ne $null) {
                $comments = $po_comment + $comments
            }
            $comments = $comments.Replace([Environment]::NewLine,"\n")
            $body ="$body_1$comments$body_2"
            $comments_patch = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body
        }
        If ($Inventory) {
            #Change Work Notes
            $work_notes = "Inventory Project Update"
            $body_1 = '{"work_notes":"'
            $body = "$body_1$work_notes$body_2"
            $work_notes_patch = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body
        }
        #Change Parent Serial Number
        if ($Updates.'Parent ID' -match '\w') {
            $parent_id = $Updates.'Parent ID'
            $asset_sys_id = $parent_id
            $body_1 = '{"parent":"'
            $body ="$body_1$parent_id$body_2"
            $parent_patch = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body
        #Update Parent Information
            $uri = "https://$instance_name/api/now/table/alm_asset/$parent_id"
            #Change Inventory Date
            if ($Updates.'Inventory Date' -match '\w') {
                $inventory_date = $Updates.'Inventory Date'
                $body_1 = '{"u_last_physical_inventory":"'
                $body ="$body_1$inventory_date$body_2"
                $inventory_date_patch = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body
            }
            #Change Inventory By
            if ($Updates.'Last Inventory By ID' -match '\w') {
                $inventory_by = $Updates.'Last Inventory By ID'
                $body_1 = '{"u_last_inventory_by":"'
                $body = "$body_1$inventory_by$body_2"
                $inventory_by_patch = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body
            }
            #Change Responsible Department
            if ($Updates.'Responsible Department ID' -match '\w') {
                $responsible_department_id = $Updates.'Responsible Department ID'
                $body_1 = '{"managed_by":"'
                $body = "$body_1$responsible_department_id$body_2"
                $responsible_department_patch = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body
                $method = "get"
                $uri = "https://$instance_name/api/now/table/alm_asset/$parent_id"
                $department_check = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                $method = "patch"
                If ($department_check.result.department.value -ne $responsible_department_id) {
                    $body_1 = '{"department":"'
                    $body = "$body_1$responsible_department_id$body_2"
                    $department_patch = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body
                }
            }
            #Change Location
            if ($Updates.'Location ID' -match '\w') {
                $location_id = $Updates.'Location ID'
                $body_1 = '{"location":"'
                $body ="$body_1$location_id$body_2"
                $parent_patch = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body
            }
            #Change Work Notes
            $work_notes = "Updated by: $($Updates.'Last Inventory By')"
            $body_1 = '{"work_notes":"'
            $body = "$body_1$work_notes$body_2"
            $work_notes_patch = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body
            #Move Consumables
            $method = "get"
            $uri = "https://$instance_name/api/now/table/alm_consumable?sysparm_query=parent%3D$($serial_hw_response.result.sys_id)"
            $current_consumables_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
            If ($current_consumables_response.result -ne $null) {
                $method = "post"
                $uri = "https://$instance_name/api/wmh/consumables/parentSwap"
                $body = "{`"newParent`":`"$parent_id`",`"oldParent`":`"$($serial_hw_response.result.sys_id)`",`"newUser`":`"`"}"
                $move_to_parent_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body
                $method = "patch"
            }
        }else{
            $asset_sys_id = $install_status_patch.result.sys_id
        }

        #Change Consumables
        if ($Updates.'Consumables IDs' -match '\w') {
            #Get Consumable Names and Models from Update List
            $method = "get"
            $consumables = New-Object System.Collections.ArrayList
            $consumables_list = $Updates.'Consumables IDs'
            $consumables_list = $consumables_list.Split([Environment]::NewLine, [StringSplitOptions]::RemoveEmptyEntries)
            foreach ($search in $consumables_list) {
                $uri = "https://$instance_name/api/now/table/alm_consumable?sysparm_query=model%3D$search^stockroom%3D$stockroom_id"
                $consumables_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                If ($($consumables_response.result) -eq $null) {
                    $uri = "https://$instance_name/api/now/table/cmdb_model_category?sysparm_query=name%3DConsumable"
                    $category = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                    $category = $category.result.sys_id
                    $stock_status = ($StateOptions | Where-Object {$_.label -eq "In stock"}).value
                    $uri = "https://$instance_name/api/now/table/cmn_cost_center?sysparm_query=name%3D8247"
                    $cost_center = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                    $cost_center = $cost_center.result.sys_id
                    $uri = "https://$instance_name/api/now/table/alm_consumable"
                    $method = "post"
                    $body = "{`"model_category`":`"$category`",`"model`":`"$search`",`"quantity`":`"10`",`"install_status`":`"6`",`"stockroom`":`"$stockroom_id`",`"cost_center`":`"$cost_center`"}"
                    $add_stock = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body
                    $method = "get"
                    $uri = "https://$instance_name/api/now/table/alm_consumable?sysparm_query=model%3D$search^stockroom%3D$stockroom_id"
                    $consumables_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                }
                $row = "" | Select Name,Model,SysID
                if ($consumables_response.result.Count -gt 1) {
                    $row.Name = $consumables_response.result.display_name[0]
                    $row.Model = $consumables_response.result.model.value[0]
                    $row.SysID = $consumables_response.result.sys_id[0]
                }else{
                    $row.Name = $consumables_response.result.display_name
                    $row.Model = $consumables_response.result.model.value
                    $row.SysId = $consumables_response.result.sys_id
                }
                [System.Collections.ArrayList]$consumables += $row
            }
        #Get Currently Assigned Consumables
            $uri = "https://$instance_name/api/now/table/alm_consumable?sysparm_query=parent%3D$asset_sys_id"
            $current_consumables_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
            $current_consumables = @() 
            foreach ($item in $current_consumables_response.result) {
                $row = "" | Select Name,Model,Quantity,SysID
                $row.Name = $item.display_name
                $row.Model = $item.model.value
                $row.Quantity = $item.quantity
                $row.SysID = $item.sys_id
                $current_consumables += $row
            }

            $quantity = @{}
            for ($j = 0; $j -lt $current_consumables.Count; $j++) {
                $quantity[$j] = $current_consumables.Quantity[$j]
            }

            for ($i = 0; $i -le $current_consumables.Count; $i++) {
                If ($current_consumables.Count -le 1) {
                    $compare = $current_consumables.Model
                    $compare_id = $current_consumables.SysID
                    $compare_qty = $current_consumables.Quantity
                }else{
                    $compare = $current_consumables.Model[$i]
                    $compare_id = $current_consumables.SysID[$i]
                    $compare_qty = $current_consumables.Quantity[$i]
                }
                If ($consumables.Model -notcontains $compare) {
                    #Return $compare_qty quantity of $compare to Stock
                    $method = "post"
                    $uri = "https://$instance_name/api/wmh/consumables/returnOneToStock"
                    $body = "{`"stockroom`":`"$stockroom_id`",`"consumable`":`"$compare_id`"}"
                    $response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body
                    $method = "get"
                }else{
                    $qty = ($consumables.Model | Where-Object {$_ -eq $compare}).Count
                    $current_qty = ($current_consumables | Where-Object {$_.Model -eq $compare}).Quantity
                    If ($current_qty.Count -gt 1) {
                        $sum = 0
                        $current_qty | Foreach {$sum += $_}
                        $current_qty = $sum
                    }
                    $difference = $current_qty - $qty
                    If ($difference -gt 0) {
                        #Return $compare to Stock
                        $method = "post"
                        $uri = "https://$instance_name/api/wmh/consumables/returnOneToStock"
                        $body = "{`"stockroom`":`"$stockroom_id`",`"consumable`":`"$compare_id`"}"
                        $response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body
                        <#
                        #Add $qty $compare's Back to Record
                        If ($Updates.'Assigned To ID' -match '\w') {
                            $user_sys_id = $Updates.'Assigned To ID'
                        }else{
                            $user_sys_id = ""
                        }
                        $method = "post"
                        $uri = "https://$instance_name/api/wmh/consumables/assignConsumable"
                        $body = "{`"consumable`":`"$compare_id`",`"asset`":`"$asset_sys_id`",`"qty`":`"$qty`",`"user`":`"$user_sys_id`"}"
                        $response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body                        
                        $method = "get"
                        #>
                    }
                }
            }
            #Get Currently Assigned Consumables After Changes Have Been Made
            $uri = "https://$instance_name/api/now/table/alm_consumable?sysparm_query=parent%3D$asset_sys_id"
            $current_consumables_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
            $current_consumables = @() 
            foreach ($item in $current_consumables_response.result) {
                $row = "" | Select Name,Model,Quantity,SysID
                $row.Name = $item.display_name
                $row.Model = $item.model.value
                $row.Quantity = $item.quantity
                $row.SysID = $item.sys_id
                $current_consumables += $row
            }

            $quantity = @{}
            for ($j = 0; $j -lt $current_consumables.Count; $j++) {
                $quantity[$j] = $current_consumables.Quantity[$j]
            }

            for ($i = 0; $i -le $consumables.Count; $i++) {
                for ($k = 0; $k -le $current_consumables.Count; $k++) {
                    If ($consumables.Count -le 1) {
                        $compare = $consumables.Model
                    }else{
                        $compare = $consumables.Model[$i]
                    }
                    If ($current_consumables.Count -le 1) {
                        $contrast = $current_consumables.Model
                    }else{
                        $contrast = $current_consumables[$k].Model
                    }
                    If ($compare -eq $contrast) {
                        If ($quantity[$k] -gt 0) {
                            $consumables.RemoveAt($i)
                            $quantity[$k] = $quantity[$k] - 1
                            $i = -1
                            $k = $current_consumables.Count
                        }
                    }
                }
            }
        #Assign Remaining Consumables
            if ($consumables -ne $null) {
                $method = "post"
                $uri = "https://$instance_name/api/wmh/consumables/assignConsumable"
                If ($Updates.'Assigned To ID' -match '\w') {
                    $user_sys_id = $Updates.'Assigned To ID'
                }else{
                    $user_sys_id = ""
                }
                foreach ($consumable_item in $consumables) {
                    $body = "{`"consumable`":`"$($consumable_item.SysID)`",`"asset`":`"$asset_sys_id`",`"qty`":`"1`",`"user`":`"$user_sys_id`"}"
                    $response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri -Body $body
                }
            }
        }
    }

    End {
        $method = "get"
        $uri = "https://$instance_name/api/now/table/alm_hardware/$recent_entry"
        $results_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri

        #Create Final Results Box
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing

        $updates_form = New-Object System.Windows.Forms.Form
        $updates_form.Text = "SNAM Updates"
        $updates_form.AutoSize = $true
        $updates_form.StartPosition = 'CenterScreen'

        $font_bold = New-Object System.Drawing.Font("Microsoft Sans Serif",8.25,[System.Drawing.FontStyle]::Bold)

        #Asset Header
        $computer_header = New-Object System.Windows.Forms.Label
        $computer_header.Location = New-Object System.Drawing.Point(0,0)
        $computer_header.AutoSize = $true
        $computer_header.Font = $font_bold
        $computer_header.Text = "
-----------------------------------------
          Computer Updates
-----------------------------------------"
        $updates_form.Controls.Add($computer_header)

        $infolabel1 = New-Object System.Windows.Forms.Label
        $infolabel1.Location = New-Object System.Drawing.Point(0,$($computer_header.Size.Height + $computer_header.Location.Y))
        $infolabel1.AutoSize = $true
        $infolabel1.Text = "
Computer Name:

Serial Number:

Model:

Install Status:

Function:

Assigned Date:

Installed Date:

Comments:"
        $updates_form.Controls.Add($infolabel1)

        $uri = $($results_response.result.model.link)
        $model_results = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
        $state_results = ($StateOptions.result | Where-Object {$_.value -eq $($results_response.result.install_status)}).label

        #Header Information
        $infolabel2 = New-Object System.Windows.Forms.Label
        $infolabel2.Location = New-Object System.Drawing.Point($($($infolabel1.Size.Width)),$($computer_header.Size.Height + $computer_header.Location.Y))
        $infolabel2.AutoSize = $true
        $infolabel2.MaximumSize = New-Object System.Drawing.Point(255,0)
        $infolabel2.Font = $font_bold
        $infolabel2.Text = "
$($results_response.result.u_computer_name)

$($results_response.result.serial_number)

$($model_results.result.name)

$state_results

$($results_response.result.u_device_function)

$($results_response.result.assigned)

$($results_response.result.install_date)

$($results_response.result.comments)"
        $updates_form.Controls.Add($infolabel2)
        
        #Header Labels Column 2
        $infolabel3 = New-Object System.Windows.Forms.Label
        $infolabel3.Location = New-Object System.Drawing.Point (355,$($computer_header.Size.Height + $computer_header.Location.Y))
        $infolabel3.AutoSize = $true
        $infolabel3.Text = "
Assigned To:

Responsible Department:

Location:

Warranty Date:

PO:

IP Address:

Last Inventory Date:

Consumables:"
        $updates_form.Controls.Add($infolabel3)

        $uri = $($results_response.result.managed_by.link)
        If ($uri -ne $null) {
            $responsible_department_result = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
        }else{
            $responsible_department_result = $null
        }
        $uri = $($results_response.result.location.link)
        $location_result = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
        $uri = $($results_response.result.assigned_to.link)
        If ($uri -ne $null) {
            $assigned_to_confirm = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
            $assigned_to_confirm = $assigned_to_confirm.result.name
        }else{
            $assigned_to_confirm = $null
        }

        #Header Info Column 2
        $infolabel4 = New-Object System.Windows.Forms.Label
        $infolabel4.Location = New-Object System.Drawing.Point($($($infolabel3.Location.X) + $($infolabel3.Size.Width)),$($computer_header.Size.Height + $computer_header.Location.Y))
        $infolabel4.AutoSize = $true
        $infolabel4.Font = $font_bold
        $infolabel4.Text = "
$assigned_to_confirm

$($responsible_department_result.result.name)

$($location_result.result.name)

$($results_response.result.warranty_expiration)

$($results_response.result.po_number)

$($results_response.result.ip_address)

$($results_response.result.u_last_physical_inventory)"
        $updates_form.Controls.Add($infolabel4)

        $uri = "https://$instance_name/api/now/table/alm_consumable?sysparm_query=parent%3D$recent_entry"
        $consumables_result = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri

        #Consumables Label
        $consumables_label = New-Object System.Windows.Forms.Label
        $consumables_label.Location = New-Object System.Drawing.Point($($infolabel3.Location.X),$($infolabel3.Location.Y + $infolabel3.Size.Height))
        $consumables_label.AutoSize = $true
        $consumables_label.Font = $font_bold
        $consumables_label.Text = "$($($consumables_result.result.display_name | Out-String).Trim())"
        $updates_form.Controls.Add($consumables_label)

        if ($parent_removed -ne $true -and $results_response.result.parent -ne $null) {
            $uri = $($results_response.result.parent.link)
            $parent_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri

            #Parent Label
            $parent_header = New-Object System.Windows.Forms.Label
            $parent_header.Location = New-Object System.Drawing.Point(0,$($infolabel2.Location.Y + $infolabel2.Size.Height + 10))
            $parent_header.AutoSize = $true
            $parent_header.Font = $font_bold
            $parent_header.Text = "
-----------------------------------------
            Parent Updates
-----------------------------------------"
            $updates_form.Controls.Add($parent_header)

            $infolabel5 = New-Object System.Windows.Forms.Label
            $infolabel5.Location = New-Object System.Drawing.Point(0,$($parent_header.Location.Y + $parent_header.Size.Height))
            $infolabel5.AutoSize = $true
            $infolabel5.Text = "
Serial Number:

Model Category:

Model:

Comments:"
            $updates_form.Controls.Add($infolabel5)

            $uri = $($parent_response.result.model_category.link)
            $model_category_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
            $uri = $($parent_response.result.model.link)
            $model_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri

            $infolabel6 = New-Object System.Windows.Forms.Label
            $infolabel6.Location = New-Object System.Drawing.Point($($($infolabel5.Size.Width)),$($infolabel5.Location.Y))
            $infolabel6.AutoSize = $true
            $infolabel6.Font = $font_bold
            $infolabel6.Text = "
$($parent_response.result.serial_number)

$($model_category_response.result.name)

$($model_response.result.name)

$($parent_response.result.comments)"
            $updates_form.Controls.Add($infolabel6)

            $infolabel7 = New-Object System.Windows.Forms.Label
            $infolabel7.Location = New-Object System.Drawing.Point(355,$($infolabel5.Location.Y))
            $infolabel7.AutoSize = $true
            $infolabel7.Text = "
Responsible Department:

Location:

Last Physical Inventory:

Consumables:"
            $updates_form.Controls.Add($infolabel7)

            $uri = $($parent_response.result.managed_by.link)
            $parent_mangaed_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
            $uri = $($parent_response.result.location.link)
            $parent_location_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri

            $infolabel8 = New-Object System.Windows.Forms.Label
            $infolabel8.Location = New-Object System.Drawing.Point($($infolabel7.Location.X + $infolabel7.Size.Width),$($infolabel7.Location.Y))
            $infolabel8.AutoSize = $true
            $infolabel8.Font = $font_bold
            $infolabel8.Text = "
$($parent_mangaed_response.result.name)

$($parent_location_response.result.name)

$($parent_response.result.u_last_physical_inventory)"
            $updates_form.Controls.Add($infolabel8)

            $uri = "https://$instance_name/api/now/table/alm_consumable?sysparm_query=parent%3D$($parent_response.result.sys_id)"
            $parent_consumables_result = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
            
            $parent_consumables_label = New-Object System.Windows.Forms.Label
            $parent_consumables_label.Location = New-Object System.Drawing.Point($($infolabel7.Location.X),$($infolabel7.Location.Y + $infolabel7.Size.Height))
            $parent_consumables_label.AutoSize = $true
            $parent_consumables_label.Font = $font_bold
            $parent_consumables_label.Text = "$($($parent_consumables_result.result.display_name | Out-String).Trim())"
            $updates_form.Controls.Add($parent_consumables_label)
        }

        If ($Swap) {
            If ($infolabel8 -ne $null) {
                $header_height = $($parent_consumables_label.Location.Y + $parent_consumables_label.Size.Height + 10)
            }else{
                $header_height = $($consumables_label.Location.Y + $consumables_label.Size.Height + 10)
            }
            $retired_header = New-Object System.Windows.Forms.Label
            $retired_header.Location = New-Object System.Drawing.Point(0,$header_height)
            $retired_header.AutoSize = $true
            $retired_header.Font = $font_bold
            $retired_header.Text = "
-----------------------------------------
          Retired Computer(s)
-----------------------------------------
$($($($retire_update.result.serial_number) | Out-String).Trim())
"
            $updates_form.Controls.Add($retired_header)
        }

        If ($check_serials -ne $null) {
            If ($retired_header -ne $null) {
                $check_header_width = $($retired_header.Location.X + $retired_header.Size.Width + 10)
                $check_header_height = $($retired_header.Location.Y)
            }else{
                If ($infolabel8 -ne $null) {
                    $check_header_width = 0
                    $check_header_height = $($parent_consumables_label.Location.Y + $parent_consumables_label.Size.Height + 10)
                }else{
                    $check_header_width = 0
                    $check_header_height = $($consumables_label.Location.Y + $consumables_label.Size.Height + 10)
                }
            }
            $check_serials_header = New-Object System.Windows.Forms.Label
            $check_serials_header.Location = New-Object System.Drawing.Point($check_header_width,$check_header_height)
            $check_serials_header.AutoSize = $true
            $check_serials_header.Font = $font_bold
            $check_serials_header.Text = "
-------------------------------------------------
 Computer(s) With Incorrect Name(s)
-------------------------------------------------
$($($check_serials | Out-String).Trim())
"
            $updates_form.Controls.Add($check_serials_header)
        }

        If ($other_entries -ne $null) {
            $other_entries_serials = @()
            foreach ($item in $other_entries) {
                $uri = "https://$instance_name/api/now/table/alm_hardware/"
                $other_entries_response = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                $other_entries_serials += $($other_entries_response.result.serial_number)
            }

            If ($check_serials_header -ne $null) {
                $other_entries_width = $($check_serials_header.Location.X + $check_serials_header.Size.Width + 10)
                $other_entries_height = $($check_serials_header.Location.Y)
            }else{
                If ($retired_header -ne $null) {
                    $other_entries_width = $($retired_header.Size.Width + 10)
                    $other_entries_height = $($retired_header.Location.Y)
                }else{
                    If ($infolabel8 -ne $null) {
                        $other_entries_width = 0
                        $other_entries_height = $($parent_consumables_label.Location.Y + $parent_consumables_label.Size.Height + 10)
                    }else{
                        $other_entries_width = 0
                        $other_entries_height = $($consumables_label.Location.Y + $consumables_label.Size.Height + 10)
                    }
                }
            }
            $other_entries_header = New-Object System.Windows.Forms.Label
            $other_entries_header.Location = New-Object System.Drawing.Point($other_entries_width,$other_entries_height)
            $other_entries_header.AutoSize = $true
            $other_entries_header.Font = $font_bold
            $other_entries_header.Text = "
-------------------------------------------
    Duplicate Serial Number(s)
-------------------------------------------
$($($other_entries_serials | Out-String).Trim())
"
            $updates_form.Controls.Add($other_entries_header)
        }

        #Confirm Button
        $OKButton = New-Object System.Windows.Forms.Button
        $OKButton.Location = New-Object System.Drawing.Point($($updates_form.Size.Width - 90),$($updates_form.Size.Height - 40))
        $OKButton.Size = New-Object System.Drawing.Size(75,23)
        $OKButton.Text = 'Confirm'
        $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $updates_form.Controls.Add($OKButton)

        #Mark Errors Button
        $ErrorsButton = New-Object System.Windows.Forms.Button
        $ErrorsButton.Location = New-Object System.Drawing.Point($($OKButton.Location.X - 80),$($OKButton.Location.Y))
        $ErrorsButton.Size = New-Object System.Drawing.Size(75,23)
        $ErrorsButton.Text = 'Mark Errors'

        $mark_errors = {
            Add-Type -AssemblyName System.Windows.Forms
            Add-Type -AssemblyName System.Drawing

            $errors_form = New-Object System.Windows.Forms.Form
            $errors_form.Text = "Error Comments"
            $errors_form.Size = New-Object System.Drawing.Size(427,250)
            $errors_form.StartPosition = 'CenterScreen'
            
            $errors_label = New-Object System.Windows.Forms.Label
            $errors_label.Location = New-Object System.Drawing.Point(5,12)
            $errors_label.AutoSize = $true
            $errors_label.Text = "Describe the error:"
            $errors_form.Controls.Add($errors_label)

            $errors_box = New-Object System.Windows.Forms.TextBox
            $errors_box.Location = New-Object System.Drawing.Point(5,32)
            $errors_box.Size = New-Object System.Drawing.Size(400,150)
            $errors_box.Multiline = $true
            $errors_form.Controls.Add($errors_box)

            $OKButton = New-Object System.Windows.Forms.Button
            $OKButton.Location = New-Object System.Drawing.Point(213,$($errors_box.Size.Height + 35))
            $OKButton.Size = New-Object System.Drawing.Size(95,23)
            $OKButton.Text = "Confirm"
            $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $errors_form.Controls.Add($OKButton)
            $errors_form.AcceptButton = $OKButton

            $CancelButton = New-Object System.Windows.Forms.Button
            $CancelButton.Location = New-Object System.Drawing.Point(310,$($errors_box.Size.Height + 35))
            $CancelButton.Size = New-Object System.Drawing.Size(95,23)
            $CancelButton.Text = "Cancel"
            $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $errors_form.Controls.Add($CancelButton)
            
            $errors_form.TopMost = $true
            $errors_confirm = $errors_form.ShowDialog()
            If ($errors_confirm -eq [System.Windows.Forms.DialogResult]::OK) {
                $SNAM_Error_Path = "$PSScriptRoot\Files\SNAM Errors.csv"
                $error.Clear()
                $UserName = $((Get-WmiObject -Class win32_computersystem).UserName).Split("\")[1]
                If ($error) {
                    $UserName = "sysklindstrom"
                }
                $uri = "https://$instance_name/api/now/table/alm_hardware/$recent_entry"
                $method = 'get'
                $recent_entry_serial = Invoke-RestMethod -Headers $headers -Method $method -Uri $uri
                New-Object -TypeName PSCustomObject -Property @{
                    "User" = $UserName
                    "Computer ID" = $recent_entry
                    "Parent ID" = $($results_response.result.parent.value)
                    "Serial Number" = $($recent_entry_serial.result.serial_number)
                    "Retired Serial Numbers" = $($($($retire_update.result.serial_number) | Out-String).Trim())
                    "Serials with Wrong Names" = $($($check_serials | Out-String).Trim())
                    "Duplicate Serial Numbers" = $($($other_entries_serials | Out-String).Trim())
                    "Comments" = $($($errors_box.Text | Out-String).Trim())
                } | Select-Object "User","Computer ID","Parent ID","Serial Number","Retired Serial Numbers","Serials with Wrong Names","Duplicate Serial Numbers","Comments" | Export-Csv -Path $SNAM_Error_Path -NoTypeInformation -Append -Encoding ASCII

                $updates_form.Close()
            }
        }
        $ErrorsButton.add_click($mark_errors)
        $updates_form.Controls.Add($ErrorsButton)

        $updates_form.ShowDialog()
    }
}

#Get ServiceNow API Credentials
$creds = Get-Credential -Message "Enter your ServiceNow REST API credentials."
$user = $creds.UserName
$pass = $creds.GetNetworkCredential().Password

#Get ServiceNow Instance Name
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$domain_form = New-Object System.Windows.Forms.Form
$domain_form.Size = New-Object System.Drawing.Size(301,128)
$domain_form.Text = "Instance Name"
$domain_form.StartPosition = 'CenterScreen'
$domain_form.TopMost = $true

$header_label = New-Object System.Windows.Forms.Label
$header_label.Text = "Enter your ServiceNow Instance Name:"
$header_label.Location = New-Object System.Drawing.Point(5,8)
$header_label.AutoSize = $true
$domain_form.Controls.Add($header_label)

$domain_box = New-Object System.Windows.Forms.TextBox
$domain_box.Location = New-Object System.Drawing.Point(5,30)
$domain_box.Size = New-Object System.Drawing.Size(180,20)
$domain_box.TextAlign = 'Right'
$domain_form.Controls.Add($domain_box)

$domain_label = New-Object System.Windows.Forms.Label
$domain_label.Text = ".service-now.com/"
$domain_label.Location = New-Object System.Drawing.Point(184,34)
$domain_label.AutoSize = $true
$domain_form.Controls.Add($domain_label)

$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Point(103,60)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = "OK"
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$domain_form.Controls.Add($OKButton)
$domain_form.AcceptButton = $OKButton

$domain_result = $domain_form.ShowDialog()
If ($domain_result -eq [System.Windows.Forms.DialogResult]::OK) {
    $instance_name = $domain_box.Text
    $instance_name = "$($instance_name).service-now.com"
}

$try = "Yes"
$message = "good"
While ($try -eq "Yes") {
    #Computer Info Box
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $create_form = New-Object System.Windows.Forms.Form
    $create_form.Text = "SNAM Swap Updates Tool"
    $create_form.Size = New-Object System.Drawing.Size(387,210)
    $create_form.StartPosition = 'CenterScreen'
            
    $error_label = New-Object System.Windows.Forms.Label
    $error_label.Location = New-Object System.Drawing.Point(5,12)
    $error_label.AutoSize = $true
    $error_label.Text = "Enter Computer Name, Old Serial Number, and New Serial Number:"
    If ($message -eq "error") {
        $error_label.ForeColor = "red"
    }
    $create_form.Controls.Add($error_label)

    $Name_label = New-Object System.Windows.Forms.Label
    $Name_label.Location = New-Object System.Drawing.Point(19,45)
    $Name_label.AutoSize = $true
    $Name_label.Text = "Computer Name:"
    $create_form.Controls.Add($Name_label)

    $Serial_label = New-Object System.Windows.Forms.Label
    $Serial_label.Location = New-Object System.Drawing.Point(8,75)
    $Serial_label.AutoSize = $true
    $Serial_label.Text = "Old Serial Number:"
    $create_form.Controls.Add($Serial_label)

    $new_Serial_label = New-Object System.Windows.Forms.Label
    $new_Serial_label.Location = New-Object System.Drawing.Point(5,105)
    $new_Serial_label.AutoSize = $true
    $new_Serial_label.Text = "New Serial Number:"
    $create_form.Controls.Add($new_Serial_label)

    $verify_Name_box = New-Object System.Windows.Forms.TextBox
    $verify_Name_box.Location = New-Object System.Drawing.Point($($new_Serial_label.Location.X + $new_Serial_label.Size.Width),$($Name_label.Location.Y - 1))
    $verify_Name_box.Size = New-Object System.Drawing.Size(255,25)
    $create_form.Controls.Add($verify_Name_box)

    $verify_Serial_box = New-Object System.Windows.Forms.TextBox
    $verify_Serial_box.Location = New-Object System.Drawing.Point($($new_Serial_label.Location.X + $new_Serial_label.Size.Width),$($Serial_label.Location.Y - 1))
    $verify_Serial_box.Size = New-Object System.Drawing.Size(255,25)
    $create_form.Controls.Add($verify_Serial_box)

    $new_Serial_box = New-Object System.Windows.Forms.TextBox
    $new_Serial_box.Location = New-Object System.Drawing.Point($($new_Serial_label.Location.X + $new_Serial_label.Size.Width),$($new_Serial_label.Location.Y - 1))
    $new_Serial_box.Size = New-Object System.Drawing.Size(255,25)
    $create_form.Controls.Add($new_Serial_box)

    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Point(172,140)
    $OKButton.Size = New-Object System.Drawing.Size(95,23)
    $OKButton.Text = "Search"
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $create_form.Controls.Add($OKButton)
    $create_form.AcceptButton = $OKButton

    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Point(270,140)
    $CancelButton.Size = New-Object System.Drawing.Size(95,23)
    $CancelButton.Text = "Exit"
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $create_form.Controls.Add($CancelButton)

    $search_result = $create_form.ShowDialog()

    If ($search_result -eq [System.Windows.Forms.DialogResult]::OK) {
        $ComputerName = $verify_Name_box.Text
        $OldSerialNumber = $verify_Serial_box.Text
        $NewSerialNumber = $new_Serial_box.Text
        If ($ComputerName -notmatch '\w' -or $OldSerialNumber -notmatch '\w' -or $NewSerialNumber -notmatch '\w') {
            $message = "error"
        }else{
            Get-SNAMInfo -ComputerName $ComputerName -SerialNumber $OldSerialNumber -Swap -user $user -pass $pass -instance_name $instance_name
            Update-SNAMInfo -ComputerName $ComputerName -SerialNumber $NewSerialNumber -Swap -user $user -pass $pass -instance_name $instance_name
        }
    }
    If ($search_result -eq [System.Windows.Forms.DialogResult]::Cancel) {
        $try = "No"
    }
}