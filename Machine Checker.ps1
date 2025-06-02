################################################################################
#                          INITIALIZATION                                      #
################################################################################

Add-Type -AssemblyName System.Windows.Forms

# Paths to CSV files
$equipmentCsvPath = "C:\Users\JR\Documents\Work\Programming Projects\LaborChecklist\Mapping CSVs\Machines_and_Acronyms.csv"
$machineClassCsvPath = "C:\Users\JR\Documents\Work\Programming Projects\LaborChecklist\Mapping CSVs\Machine_PubNum_Mappings.csv"

# Load data
try {
    $machineClassCodes = Import-Csv -Path $machineClassCsvPath | Group-Object Acronym -AsHashTable -AsString
}
catch {
    # Show error in a message box instead of console
    [System.Windows.Forms.MessageBox]::Show(
        "Failed to load CSV files: $_",
        "Critical Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
    return
}

################################################################################
#                        LOGGING FUNCTIONS                                     #
################################################################################

function Write-DebugLog {
    param (
        [string]$Message,
        [string]$Category = "General"
    )
    
    # Format message with timestamp and category
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
    $logMessage = "[$timestamp] [$Category] $Message"
    
    # Output to console (this always works)
    Write-Host $logMessage
    
    # Try to log to a file in a location with proper permissions
    try {
        $logFolder = "$env:USERPROFILE\AppData\Local\Temp\MachineCheckerLogs"
        
        # Create the directory if it doesn't exist
        if (-not (Test-Path $logFolder)) {
            New-Item -Path $logFolder -ItemType Directory -Force | Out-Null
        }
        
        $logFile = Join-Path $logFolder "debug_log.txt"
        Add-Content -Path $logFile -Value $logMessage -ErrorAction SilentlyContinue
    }
    catch {
        # Silently continue if we can't write to the log file
        # At least we've already written to the console
    }
}

################################################################################
#                      CSV UTILITY FUNCTIONS                                   #
################################################################################

# Function to determine columns directly from CSV
function Get-CsvColumns {
    param (
        [string]$CsvPath
    )
    
    Write-DebugLog "Getting columns from CSV: $CsvPath" -Category "Columns"
    
    try {
        # First check if file exists
        if (-not (Test-Path -Path $CsvPath)) {
            Write-DebugLog "File not found: $CsvPath" -Category "Columns"
            return @()
        }
        
        # Read all the content of the CSV
        $csvContent = Get-Content -Path $CsvPath -ErrorAction Stop
        
        if ($csvContent.Count -eq 0) {
            Write-DebugLog "File is empty: $CsvPath" -Category "Columns"
            return @()
        }
        
        # Get the first line which contains the headers
        $headerLine = $csvContent[0]
        
        if ([string]::IsNullOrWhiteSpace($headerLine)) {
            Write-DebugLog "First line is blank: $CsvPath" -Category "Columns"
            return @()
        }
        
        # Directly parse the header by splitting on commas
        $headers = $headerLine -split ',' | ForEach-Object { $_.Trim('"').Trim() }
        Write-DebugLog "Parsed headers directly: $($headers -join ', ')" -Category "Columns"
        
        # If that worked, return the headers
        if ($headers.Count -gt 0) {
            return $headers
        }
        
        # If everything fails, try directly loading with Import-Csv
        $importedCsv = Import-Csv -Path $CsvPath -ErrorAction Stop
        if ($importedCsv -ne $null) {
            $properties = $importedCsv[0].PSObject.Properties.Name
            Write-DebugLog "Got headers from Import-Csv: $($properties -join ', ')" -Category "Columns"
            return $properties
        }
        
        # If all methods fail, return empty array
        return @()
    }
    catch {
        Write-DebugLog "Error reading CSV headers: $_" -Category "Columns"
        return @()
    }
}

################################################################################
#                MACHINE CONFIGURATION DIALOG FUNCTIONS                        #
################################################################################

function Show-MachineConfigDialog {
    param (
        [bool]$ExistingMachine = $false,
        [hashtable]$InitialValues = @{},
        [string]$MMO = "",
        [string]$MachineAcronym = "",
        [string]$ClassCode = "",        # Added parameter for class code
        [string[]]$MMOHeaders = @(),
        [string]$LookupTablePath = ""
    )
    
    Write-DebugLog "Starting Show-MachineConfigDialog" -Category "ConfigDialog"
    Write-DebugLog "ExistingMachine: $ExistingMachine, MMO: $MMO, MachineAcronym: $MachineAcronym, ClassCode: $ClassCode" -Category "ConfigDialog"
    Write-DebugLog "MMOHeaders: $($MMOHeaders -join ', ')" -Category "ConfigDialog"
    Write-DebugLog "LookupTablePath: $LookupTablePath" -Category "ConfigDialog"
    Write-DebugLog "Initial Values: $(($InitialValues.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join ', ')" -Category "ConfigDialog"
    
    # Try to extract MMO from LookupTablePath if MMO is empty
    if ([string]::IsNullOrWhiteSpace($MMO) -and -not [string]::IsNullOrWhiteSpace($LookupTablePath)) {
        # Extract MMO from path if possible
        if ($LookupTablePath -match "(MMO-\d+-\d+)") {
            $MMO = $matches[1]
            Write-DebugLog "Extracted MMO from lookup table path: $MMO" -Category "ConfigDialog"
        }
    }
    
    $configForm = New-Object System.Windows.Forms.Form
    $configForm.Text = "Machine Configuration"
    $configForm.Size = New-Object System.Drawing.Size(600, 450)  # Wider to accommodate original value notes
    $configForm.StartPosition = "CenterScreen"
    $configForm.FormBorderStyle = "FixedDialog"
    $configForm.MaximizeBox = $false
    $configForm.MinimizeBox = $false
    
    # Add instructional label
    $lblInstructions = New-Object System.Windows.Forms.Label
    $lblInstructions.Location = New-Object System.Drawing.Point(20, 20)
    $lblInstructions.Size = New-Object System.Drawing.Size(560, 40)
    $lblInstructions.Text = "Please configure the machine parameters. These values will be used to find the best matching maintenance requirements."
    $configForm.Controls.Add($lblInstructions)
    
    # Get original values from ListView item if available
    $origDaysPerWeek = $null
    $origToursPerDay = $null
    
    if ($ExistingMachine) {
        # Try to find the machine in the ListView
        $selectedItem = $null
        foreach ($item in $listView.Items) {
            if ($item.SubItems[0].Text -eq $MachineAcronym -and 
                $item.SubItems[1].Text -eq $machineNumber) {
                $selectedItem = $item
                break
            }
        }
        
        # Check if we have original values stored in the Tag
        if ($selectedItem -ne $null -and $selectedItem.Tag -is [hashtable]) {
            if ($selectedItem.Tag.ContainsKey("OriginalDaysPerWeek")) {
                $origDaysPerWeek = $selectedItem.Tag.OriginalDaysPerWeek
            }
            if ($selectedItem.Tag.ContainsKey("OriginalToursPerDay")) {
                $origToursPerDay = $selectedItem.Tag.OriginalToursPerDay
            }
        }
        # If no stored original values but we have values in InitialValues, treat those as original
        elseif ($InitialValues.ContainsKey("Operation (days/wk)") -or $InitialValues.ContainsKey("Tours/Day")) {
            if ($InitialValues.ContainsKey("Operation (days/wk)")) {
                $origDaysPerWeek = $InitialValues["Operation (days/wk)"]
            }
            if ($InitialValues.ContainsKey("Tours/Day")) {
                $origToursPerDay = $InitialValues["Tours/Day"]
            }
        }
    }
    
    # If we have original values, display them in the UI
    if ($origDaysPerWeek -ne $null -or $origToursPerDay -ne $null) {
        $origValuesText = "Original calculated values: "
        if ($origDaysPerWeek -ne $null) {
            $origValuesText += "Days/Week = $origDaysPerWeek"
        }
        if ($origDaysPerWeek -ne $null -and $origToursPerDay -ne $null) {
            $origValuesText += ", "
        }
        if ($origToursPerDay -ne $null) {
            $origValuesText += "Tours/Day = $origToursPerDay"
        }
        
        $origValuesLabel = New-Object System.Windows.Forms.Label
        $origValuesLabel.Location = New-Object System.Drawing.Point(20, 60)
        $origValuesLabel.Size = New-Object System.Drawing.Size(560, 20)
        $origValuesLabel.Text = $origValuesText
        $origValuesLabel.ForeColor = [System.Drawing.Color]::Blue
        $configForm.Controls.Add($origValuesLabel)
        
        # Add note about standardization
        $noteLabel = New-Object System.Windows.Forms.Label
        $noteLabel.Location = New-Object System.Drawing.Point(20, 80)
        $noteLabel.Size = New-Object System.Drawing.Size(560, 20)
        $noteLabel.Text = "Note: Values will be standardized to match maintenance requirements."
        $noteLabel.ForeColor = [System.Drawing.Color]::DarkGray
        $configForm.Controls.Add($noteLabel)
        
        # Adjust starting position for other controls
        $yPos = 110
    } else {
        $yPos = 70  # Normal starting position if no original values
    }
    
    $labelWidth = 140
    $controlWidth = 180  # Wider control for class code dropdown
    $controlHeight = 25
    $spacing = 10
    
    # Add Class Code label and dropdown
    $lblClassCode = New-Object System.Windows.Forms.Label
    $lblClassCode.Location = New-Object System.Drawing.Point(20, $yPos)
    $lblClassCode.Size = New-Object System.Drawing.Size($labelWidth, $controlHeight)
    $lblClassCode.Text = "Class Code:"
    $configForm.Controls.Add($lblClassCode)
    
    $cmbClassCode = New-Object System.Windows.Forms.ComboBox
    $cmbClassCode.Location = New-Object System.Drawing.Point(($labelWidth + 30), $yPos)
    $cmbClassCode.Size = New-Object System.Drawing.Size($controlWidth, $controlHeight)
    $cmbClassCode.DropDownStyle = "DropDownList"
    $configForm.Controls.Add($cmbClassCode)
    
    # Get valid class codes for this machine acronym
    if (-not [string]::IsNullOrWhiteSpace($MachineAcronym) -and $machineClassCodes.ContainsKey($MachineAcronym)) {
        $validCodes = $machineClassCodes[$MachineAcronym]."Class Code" | Sort-Object -Unique
        foreach ($code in $validCodes) {
            $cmbClassCode.Items.Add($code) | Out-Null
        }
        
        # Set selected class code if provided
        if (-not [string]::IsNullOrWhiteSpace($ClassCode) -and $validCodes -contains $ClassCode) {
            $index = $cmbClassCode.Items.IndexOf($ClassCode)
            if ($index -ge 0) {
                $cmbClassCode.SelectedIndex = $index
            } elseif ($cmbClassCode.Items.Count -gt 0) {
                $cmbClassCode.SelectedIndex = 0
            }
        } elseif ($cmbClassCode.Items.Count -gt 0) {
            $cmbClassCode.SelectedIndex = 0
        }
    }
    
    # Move to next position for other parameters
    $yPos += $controlHeight + $spacing
    
    # Store controls in a hashtable
    $inputControls = @{
        "ClassCode" = $cmbClassCode
    }
    
    # Update MMO label when class code changes
    $mmoLabel = New-Object System.Windows.Forms.Label
    $mmoLabel.Location = New-Object System.Drawing.Point(($labelWidth + 30), ($yPos - 5))
    $mmoLabel.AutoSize = $true
    if (-not [string]::IsNullOrWhiteSpace($MMO)) {
        $mmoLabel.Text = "MMO: $MMO"
    } else {
        $mmoLabel.Text = "MMO: Select a class code"
    }
    $configForm.Controls.Add($mmoLabel)
    
    # Add handler to update MMO when class code changes
    $cmbClassCode.Add_SelectedIndexChanged({
        $selectedCode = $cmbClassCode.SelectedItem
        if (-not [string]::IsNullOrWhiteSpace($selectedCode) -and $machineClassCodes.ContainsKey($MachineAcronym)) {
            $mmoEntry = $machineClassCodes[$MachineAcronym] | 
                Where-Object { $_.'Class Code' -eq $selectedCode } | 
                Select-Object -First 1
            
            if ($mmoEntry -and $mmoEntry.'Pub Num') {
                $mmoLabel.Text = "MMO: $($mmoEntry.'Pub Num')"
                
                # Update the MMO parameter for use in relevance parameters
                $global:UpdatedMMO = $mmoEntry.'Pub Num'
                
                # Try to find the lookup table for this MMO
                $baseDir = "C:\Users\JR\Documents\Work\Programming Projects\LaborChecklist\Machine Labor Rubrics"
                
                # Try alternate path if first one doesn't exist
                if (-not (Test-Path $baseDir)) {
                    $baseDir = "C:\Users\JR\Documents\Work\Programming Projects\LaborChecklist\Machine Labor Rubrics"
                }
                
                $mmoDirectories = Get-ChildItem $baseDir -Directory | Where-Object { $_.Name -like "*$($mmoEntry.'Pub Num')*" }
                
                if ($mmoDirectories.Count -gt 0) {
                    # Use the first matching directory
                    $mmoDirectory = $mmoDirectories[0].FullName
                    
                    # Look for labor lookup file in that directory
                    $lookupFiles = Get-ChildItem $mmoDirectory -Filter "*-Labor-Lookup.csv"
                    if ($lookupFiles.Count -gt 0) {
                        $global:UpdatedLookupTablePath = $lookupFiles[0].FullName
                        Write-DebugLog "Updated lookup table path: $($global:UpdatedLookupTablePath)" -Category "ConfigDialog"
                        
                        # Re-get parameters based on the new lookup table
                        UpdateParametersForClassCode
                    }
                }
            } else {
                $mmoLabel.Text = "MMO: N/A"
                $global:UpdatedMMO = ""
            }
        } else {
            $mmoLabel.Text = "MMO: Select a class code"
            $global:UpdatedMMO = ""
        }
    })
    
    # Move down for a separator
    $yPos += 20
    
    # Create a separator
    $separator = New-Object System.Windows.Forms.Label
    $separator.Location = New-Object System.Drawing.Point(20, $yPos)
    $separator.Size = New-Object System.Drawing.Size(560, 2)
    $separator.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    $configForm.Controls.Add($separator)
    
    $yPos += 10
    
    # Get relevant parameters for this MMO/machine type
    $global:UpdatedMMO = $MMO
    $global:UpdatedLookupTablePath = $LookupTablePath
    
    # Function to update parameters when class code changes
    function global:UpdateParametersForClassCode {
        # Remove existing parameter controls
        $parameterControls = $configForm.Controls | Where-Object {
            $_.Tag -eq "Parameter" -or $_.Tag -eq "ParameterLabel"
        }
        foreach ($ctrl in $parameterControls) {
            $configForm.Controls.Remove($ctrl)
        }
        
        # Get new parameters from updated MMO
        Write-DebugLog "Getting relevant parameters for updated MMO: $($global:UpdatedMMO)" -Category "ConfigDialog"
        $parameters = Get-RelevantParameters -MMO $global:UpdatedMMO -MachineAcronym $MachineAcronym `
                                          -InitialValues $InitialValues -ExistingMachine $ExistingMachine `
                                          -MMOHeaders $MMOHeaders -LookupTablePath $global:UpdatedLookupTablePath
                                          
        # Add each parameter to the form
        $currentY = $yPos
        
        foreach ($param in $parameters) {
            Write-DebugLog "Creating control for parameter: $($param.Name)" -Category "ConfigDialog"
            
            # Create label
            $label = New-Object System.Windows.Forms.Label
            $label.Location = New-Object System.Drawing.Point(20, $currentY)
            $label.Size = New-Object System.Drawing.Size($labelWidth, $controlHeight)
            $label.Text = "$($param.Name):"
            $label.Tag = "ParameterLabel"
            $configForm.Controls.Add($label)
            
            # Create input control
            if ($param.Type -eq "ComboBox") {
                $control = New-Object System.Windows.Forms.ComboBox
                $control.DropDownStyle = "DropDownList"
                
                Write-DebugLog "Adding ComboBox values for $($param.Name): $($param.Values -join ', ')" -Category "ConfigDialog"
                foreach ($value in $param.Values) {
                    $control.Items.Add($value) | Out-Null
                }
                
                # Set value from initial values or default
                if ($ExistingMachine -and $InitialValues.ContainsKey($param.Name) -and -not [string]::IsNullOrWhiteSpace($InitialValues[$param.Name])) {
                    $valueToSelect = $InitialValues[$param.Name]
                    $indexToSelect = $control.Items.IndexOf($valueToSelect)
                    
                    Write-DebugLog "Setting ComboBox from initial value: $valueToSelect (index: $indexToSelect)" -Category "ConfigDialog"
                    
                    if ($indexToSelect -ge 0) {
                        $control.SelectedIndex = $indexToSelect
                    } else {
                        if ($control.Items.Count -gt 0) {
                            $control.SelectedIndex = 0
                            Write-DebugLog "Value not found in ComboBox, setting index 0" -Category "ConfigDialog"
                        } else {
                            Write-DebugLog "WARNING: ComboBox has no items!" -Category "ConfigDialog"
                        }
                    }
                } else {
                    # Set default value
                    $defaultIndex = $control.Items.IndexOf($param.Default)
                    
                    Write-DebugLog "Setting ComboBox from default value: $($param.Default) (index: $defaultIndex)" -Category "ConfigDialog"
                    
                    if ($defaultIndex -ge 0) {
                        $control.SelectedIndex = $defaultIndex
                    } else {
                        if ($control.Items.Count -gt 0) {
                            $control.SelectedIndex = 0
                            Write-DebugLog "Default not found in ComboBox, setting index 0" -Category "ConfigDialog"
                        } else {
                            Write-DebugLog "WARNING: ComboBox has no items!" -Category "ConfigDialog"
                        }
                    }
                }
                
                # If this is a parameter with original values, add a note about adjustment if needed
                if (($param.Name -eq "Operation (days/wk)" -and $origDaysPerWeek -ne $null) -or 
                    ($param.Name -eq "Tours/Day" -and $origToursPerDay -ne $null)) {
                    
                    # Get the original value
                    $originalValue = if ($param.Name -eq "Operation (days/wk)") { $origDaysPerWeek } else { $origToursPerDay }
                    $defaultValue = if ($control.SelectedItem -ne $null) { $control.SelectedItem } else { $param.Default }
                    
                    # Check if there's a significant difference
                    $diff = [Math]::Abs([double]$originalValue - [double]$defaultValue)
                    
                    if ($diff -gt 0.2) {  # Only show note if difference is significant
                        $noteLabel = New-Object System.Windows.Forms.Label
                        $noteLabel.Location = New-Object System.Drawing.Point(($labelWidth + $controlWidth + 40), $currentY)
                        $noteLabel.Size = New-Object System.Drawing.Size(200, $controlHeight)
                        $noteLabel.Text = "(Original: $originalValue)"
                        $noteLabel.ForeColor = [System.Drawing.Color]::DarkOrange
                        $noteLabel.Tag = "ParameterLabel"  # For cleanup purposes
                        $configForm.Controls.Add($noteLabel)
                    }
                }
            } else {
                $control = New-Object System.Windows.Forms.TextBox
                if ($ExistingMachine -and $InitialValues.ContainsKey($param.Name) -and -not [string]::IsNullOrWhiteSpace($InitialValues[$param.Name])) {
                    $control.Text = $InitialValues[$param.Name]
                    Write-DebugLog "Setting TextBox from initial value: $($InitialValues[$param.Name])" -Category "ConfigDialog"
                } else {
                    $control.Text = $param.Default
                    Write-DebugLog "Setting TextBox from default value: $($param.Default)" -Category "ConfigDialog"
                }
            }
            
            $control.Location = New-Object System.Drawing.Point(($labelWidth + 30), $currentY)
            $control.Size = New-Object System.Drawing.Size($controlWidth, $controlHeight)
            $control.Tag = "Parameter"
            $configForm.Controls.Add($control)
            
            # Store control reference
            $inputControls[$param.Name] = $control
            
            # Move to next row
            $currentY += $controlHeight + $spacing
        }
        
        # Resize form to fit controls
        if ($inputControls.Count -gt 1) {  # Count includes ClassCode
            $formHeight = $currentY + 60   # Add space for buttons
            $configForm.ClientSize = New-Object System.Drawing.Size($configForm.ClientSize.Width, $formHeight)
        }
    }
    
    # Get initial parameters
    UpdateParametersForClassCode
    
    # Add OK and Cancel buttons
    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "OK"
    $btnOK.Size = New-Object System.Drawing.Size(80, 30)
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $configForm.Controls.Add($btnOK)
    $configForm.AcceptButton = $btnOK
    
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Size = New-Object System.Drawing.Size(80, 30)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $configForm.Controls.Add($btnCancel)
    $configForm.CancelButton = $btnCancel
    
    # Set button positions when form loads to avoid Location property not found errors
    $configForm.Add_Shown({
        $btnOK.Location = New-Object System.Drawing.Point(($configForm.ClientSize.Width - 180), ($configForm.ClientSize.Height - 40))
        $btnCancel.Location = New-Object System.Drawing.Point(($configForm.ClientSize.Width - 90), ($configForm.ClientSize.Height - 40))
    })
    
    # Show the form and get result
    Write-DebugLog "Showing configuration dialog" -Category "ConfigDialog"
    $result = $configForm.ShowDialog()
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        # Collect values into a hashtable
        $config = @{}
        
        foreach ($paramName in $inputControls.Keys) {
            $control = $inputControls[$paramName]
            
            if ($control -is [System.Windows.Forms.ComboBox]) {
                $config[$paramName] = $control.SelectedItem
                Write-DebugLog "Collected config value from ComboBox: $paramName = $($control.SelectedItem)" -Category "ConfigDialog"
            } else {
                $config[$paramName] = $control.Text
                Write-DebugLog "Collected config value from TextBox: $paramName = $($control.Text)" -Category "ConfigDialog"
            }
        }
        
        # Add the MMO to the config
        $config["MMO"] = $global:UpdatedMMO
        
        # Add the original values if available
        if ($origDaysPerWeek -ne $null) {
            $config["OriginalDaysPerWeek"] = $origDaysPerWeek
        }
        if ($origToursPerDay -ne $null) {
            $config["OriginalToursPerDay"] = $origToursPerDay
        }
        
        # Flag if values were adjusted
        $daysAdjusted = $origDaysPerWeek -ne $null -and $config.ContainsKey("Operation (days/wk)") -and 
                       ($config["Operation (days/wk)"] -ne $origDaysPerWeek)
                       
        $toursAdjusted = $origToursPerDay -ne $null -and $config.ContainsKey("Tours/Day") -and 
                        ($config["Tours/Day"] -ne $origToursPerDay)
                        
        $config["ValuesAdjusted"] = $daysAdjusted -or $toursAdjusted
        
        Write-DebugLog "Returning configuration with $(($config.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join ', ')" -Category "ConfigDialog"

        function Ensure-StaffingTableEntry {
            param (
                [string]$MMO,
                [string]$MachineID
            )
            
            Write-DebugLog "Starting Ensure-StaffingTableEntry for MMO: $MMO, Machine ID: $MachineID" -Category "StaffingTable"
            
            try {
                # Find the MMO directory
                $baseDir = "C:\Users\JR\Documents\Work\Programming Projects\LaborChecklist\Machine Labor Rubrics"
                
                # Check if base directory exists
                if (-not (Test-Path $baseDir)) {
                    $baseDir = "C:\Users\JR\Documents\Work\Programming Projects\LaborChecklist\Machine Labor Rubrics"
                    
                    # If still doesn't exist, try to create it
                    if (-not (Test-Path $baseDir)) {
                        New-Item -Path $baseDir -ItemType Directory -Force | Out-Null
                        Write-DebugLog "Created base directory: $baseDir" -Category "StaffingTable"
                    }
                }
                
                # Find the MMO directory
                $mmoDirectory = Get-ChildItem $baseDir -Directory |
                    Where-Object { $_.Name -like "*$MMO*" -and $_.Name -like "*-$ClassCode" } |
                    Select-Object -First 1 -ExpandProperty FullName
                
                if (-not $mmoDirectory) {
                    $mmoDirectoryName = "$MMO-$MachineAcronym-$ClassCode"
                    $mmoDirectory = Join-Path $baseDir $mmoDirectoryName
                    New-Item -Path $mmoDirectory -ItemType Directory -Force | Out-Null
                }
                
                # Define the staffing table CSV path
                $staffingFileName = "$MMO-$ClassCode-Staffing-Table.csv"
                $staffingFilePath = Join-Path $mmoDirectory $staffingFileName
                
                Write-DebugLog "Staffing table file path: $staffingFilePath" -Category "StaffingTable"
                
                # Create a DataTable for the staffing data
                $staffingTable = New-Object System.Data.DataTable
                
                # Add Machine ID column
                $staffingTable.Columns.Add("Machine ID", [string]) | Out-Null
                
                # Check if the file already exists
                if (Test-Path $staffingFilePath) {
                    # Load existing data
                    try {
                        $existingData = Import-Csv $staffingFilePath
                        
                        if ($existingData -and $existingData.Count -gt 0) {
                            # Check schema from existing file
                            $existingColumns = $existingData[0].PSObject.Properties.Name
                            
                            # Add all columns from existing file
                            foreach ($column in $existingColumns) {
                                if (-not $staffingTable.Columns.Contains($column)) {
                                    $staffingTable.Columns.Add($column, [string]) | Out-Null
                                }
                            }
                            
                            # Add all existing rows
                            foreach ($item in $existingData) {
                                $newRow = $staffingTable.NewRow()
                                
                                foreach ($prop in $item.PSObject.Properties) {
                                    if ($staffingTable.Columns.Contains($prop.Name)) {
                                        $newRow[$prop.Name] = $prop.Value
                                    }
                                }
                                
                                $staffingTable.Rows.Add($newRow)
                            }
                            
                            Write-DebugLog "Loaded existing staffing table with $($staffingTable.Rows.Count) rows" -Category "StaffingTable"
                        }
                    }
                    catch {
                        Write-DebugLog "Error loading existing staffing table: $_" -Category "StaffingTable"
                        # Continue with empty table if there was an error
                    }
                }
                else {
                    # Add standard columns for a new file
                    $staffingTable.Columns.Add("Operation (days/wk)", [string]) | Out-Null
                    $staffingTable.Columns.Add("Tours/Day", [string]) | Out-Null
                    $staffingTable.Columns.Add("MM7", [string]) | Out-Null
                    $staffingTable.Columns.Add("MPE9", [string]) | Out-Null
                    $staffingTable.Columns.Add("ET10", [string]) | Out-Null
                    $staffingTable.Columns.Add("Total (hrs/yr)", [string]) | Out-Null
                    $staffingTable.Columns.Add("Operational Maintenance (hrs/yr)", [string]) | Out-Null
                }
                
                # Check if this machine already exists in the table
                $machineExists = $false
                foreach ($row in $staffingTable.Rows) {
                    if ($row["Machine ID"] -eq $MachineID) {
                        $machineExists = $true
                        break
                    }
                }
                
                # If machine doesn't exist, add a row for it
                if (-not $machineExists) {
                    $newRow = $staffingTable.NewRow()
                    $newRow["Machine ID"] = $MachineID
                    $staffingTable.Rows.Add($newRow)
                    
                    # Save the updated table
                    $staffingTable | Export-Csv -Path $staffingFilePath -NoTypeInformation
                    Write-DebugLog "Added row for Machine ID: $MachineID to staffing table" -Category "StaffingTable"
                }
                
                return $true
            }
            catch {
                Write-DebugLog "Error ensuring staffing table entry: $_" -Category "StaffingTable"
                return $false
            }
        }

        # Test the staffing table functions directly (for debugging)
        function Test-StaffingTable {
            param (
                [string]$MMO,
                [string]$MachineID
            )
            
            Write-Host "Testing staffing table functions for MMO: $MMO, Machine ID: $MachineID"
            
            # Create a test entry
            $result = Ensure-StaffingTableEntry -MMO $MMO -MachineID $MachineID
            Write-Host "Ensure-StaffingTableEntry result: $result"
            
            # Load the staffing table
            $data = Load-StaffingTable -MMO $MMO -MachineID $MachineID
        }

        return $config

    } else {
        Write-DebugLog "User cancelled dialog" -Category "ConfigDialog"
        return $null
    }
}

################################################################################
#                    DROPDOWN CREATION FUNCTIONS                               #
################################################################################

function Get-RelevantParameters {
    param (
        [string]$MMO,
        [string]$MachineAcronym,
        [hashtable]$InitialValues = @{},
        [bool]$ExistingMachine = $false,
        [string[]]$MMOHeaders = @(),
        [string]$LookupTablePath = ""
    )
    
    Write-DebugLog "Starting Get-RelevantParameters for MMO: $MMO, Machine: $MachineAcronym" -Category "ConfigDialog"
    Write-DebugLog "Initial values: $(($InitialValues.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join ', ')" -Category "ConfigDialog"
    Write-DebugLog "Lookup table path: '$LookupTablePath'" -Category "ConfigDialog"
    
    # Find the lookup table if not provided
    if ([string]::IsNullOrWhiteSpace($LookupTablePath) -and -not [string]::IsNullOrWhiteSpace($MMO)) {
        $baseDir = "C:\Users\JR\Documents\Work\Programming Projects\LaborChecklist\Machine Labor Rubrics"
        
        # First try to find a directory that matches the MMO
        $mmoDirectories = Get-ChildItem $baseDir -Directory | Where-Object { $_.Name -like "*$MMO*" }
        
        if ($mmoDirectories.Count -gt 0) {
            # Use the first matching directory
            $mmoDirectory = $mmoDirectories[0].FullName
            Write-DebugLog "Found MMO directory: $mmoDirectory" -Category "ConfigDialog"
            
            # Extract MMO number from directory name if we don't have it already
            if ([string]::IsNullOrWhiteSpace($MMO)) {
                # Extract MMO pattern (like MMO-090-16) from directory name
                if ($mmoDirectories[0].Name -match "(MMO-\d+-\d+)") {
                    $extractedMMO = $matches[1]
                    $MMO = $extractedMMO
                    Write-DebugLog "Extracted MMO from directory name: $MMO" -Category "ConfigDialog"
                }
            }
            
            # Look for labor lookup file in that directory
            $lookupFiles = Get-ChildItem $mmoDirectory -Filter "*-Labor-Lookup.csv"
            if ($lookupFiles.Count -gt 0) {
                $LookupTablePath = $lookupFiles[0].FullName
                Write-DebugLog "Found lookup table: $LookupTablePath" -Category "ConfigDialog"
            } else {
                Write-DebugLog "No lookup table found in directory: $mmoDirectory" -Category "ConfigDialog"
                # Return empty parameters array if no lookup table
                return @()
            }
        }
    }
    
    # Define our standard output columns
    $outputColumns = @("Total (hrs/yr)", "Operational Maintenance (hrs/yr)")
    
    # Load lookup table if available to get columns and values
    $lookupData = $null
    $columnNames = @()
    
    # Safely check if lookup table exists
    if (-not [string]::IsNullOrWhiteSpace($LookupTablePath) -and (Test-Path $LookupTablePath)) {
        try {
            $lookupData = Import-Csv $LookupTablePath
            Write-DebugLog "Loaded lookup table: $LookupTablePath with $($lookupData.Count) rows" -Category "ConfigDialog"
            
            # Get column names from the lookup table
            $columnNames = $lookupData[0].PSObject.Properties.Name
            Write-DebugLog "Found columns: $($columnNames -join ', ')" -Category "ConfigDialog"
        }
        catch {
            Write-DebugLog "Error loading lookup table: $_" -Category "ConfigDialog"
            $lookupData = $null
        }
    } else {
        Write-DebugLog "No valid lookup table available at path: $LookupTablePath" -Category "ConfigDialog"
        # Return empty parameters array
        return @()
    }
    
    # Use provided headers if available
    if ($MMOHeaders.Count -gt 0) {
        $columnNames = $MMOHeaders
        Write-DebugLog "Using provided headers: $($columnNames -join ', ')" -Category "ConfigDialog"
    }
    
    # Initialize parameters array
    $parameters = @()
    
    # If we have lookup data, create parameters for all input columns
    if ($lookupData -and $columnNames.Count -gt 0) {
        # Filter out output columns to get input columns
        $inputColumns = $columnNames | Where-Object { $_ -notin $outputColumns }
        Write-DebugLog "Input columns for configuration: $($inputColumns -join ', ')" -Category "ConfigDialog"
        
        # Create a parameter for each input column
        foreach ($column in $inputColumns) {
            # Get unique values for this column
            $uniqueValues = $lookupData | Select-Object -ExpandProperty $column -Unique | 
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | 
                Sort-Object
            
            if ($uniqueValues.Count -gt 0) {
                # Determine default value
                $defaultValue = $uniqueValues[0] # Default to first value
                
                # Check if this column has a typical default
                if ($column -eq "Operation (days/wk)" -and $uniqueValues -contains "6") {
                    $defaultValue = "6"
                }
                elseif ($column -eq "Tours/Day" -and $uniqueValues -contains "2") {
                    $defaultValue = "2"
                }
                
                # Set existing value if we have one
                if ($ExistingMachine -and $InitialValues.ContainsKey($column)) {
                    $defaultValue = $InitialValues[$column]
                }
                
                # Create parameter definition
                $parameters += @{
                    Name = $column
                    Type = "ComboBox"
                    Values = $uniqueValues
                    Default = $defaultValue
                }
                
                Write-DebugLog "Added parameter for $column with values: $($uniqueValues -join ', ')" -Category "ConfigDialog"
            }
        }
    }
    
    Write-DebugLog "Returning $($parameters.Count) parameters" -Category "ConfigDialog"
    return $parameters
}

################################################################################
#                     DATA MATCHING FUNCTIONS                                  #
################################################################################

# Function to highlight matched cells in DataGridView
function Highlight-MatchedCells {
    param (
        [System.Windows.Forms.DataGridView]$DataGridView,
        [hashtable]$MachineMetrics,
        [hashtable]$ColumnMapping
    )
    
    if ($null -eq $DataGridView -or $null -eq $MachineMetrics -or $DataGridView.Rows.Count -eq 0) {
        Write-DebugLog "Cannot highlight cells - missing required parameters or no rows" -Category "ViewDetails"
        return
    }
    
    # Define highlight colors
    $inputHighlightColor = [System.Drawing.Color]::FromArgb(255, 255, 200)  # Light yellow
    $totalHighlightColor = [System.Drawing.Color]::FromArgb(200, 255, 200)  # Light green
    
    # Log all machine metrics for debugging
    Write-DebugLog "Trying to match with metrics: $(($MachineMetrics.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join ', ')" -Category "ViewDetails"
    
    # Get column names from the DataGridView
    $columnNames = @()
    foreach ($col in $DataGridView.Columns) {
        $columnNames += $col.Name
    }
    Write-DebugLog "DataGridView columns: $($columnNames -join ', ')" -Category "ViewDetails"
    
    # For each row in the DataGridView, calculate a match score
    $bestRow = -1
    $bestScore = 0
    
    for ($rowIndex = 0; $rowIndex -lt $DataGridView.Rows.Count; $rowIndex++) {
        $row = $DataGridView.Rows[$rowIndex]
        $rowScore = 0
        $matchCount = 0
        $totalParams = 0
        
        # Check each metric against row values
        foreach ($paramName in $MachineMetrics.Keys) {
            # Skip non-parameter metrics
            if ($paramName -eq "MachineType" -or $paramName -eq "MachineNumber") {
                continue
            }
            
            # Skip empty values
            if ([string]::IsNullOrWhiteSpace($MachineMetrics[$paramName])) {
                continue
            }
            
            $totalParams++
            
            # If DataGridView has this column, check for a match
            if ($columnNames -contains $paramName) {
                $cellValue = $row.Cells[$paramName].Value
                $metricValue = $MachineMetrics[$paramName]
                
                # Try to convert to same type for comparison
                if ($cellValue -is [string] -and $metricValue -isnot [string]) {
                    $metricValue = $metricValue.ToString()
                }
                elseif ($cellValue -is [int] -and $metricValue -is [string]) {
                    $metricValue = [int]::Parse($metricValue)
                }
                
                Write-DebugLog "Comparing column ${paramName}: Cell=$cellValue, Metric=$metricValue" -Category "ViewDetails"
                
                if ($cellValue -eq $metricValue) {
                    $matchCount++
                    $rowScore += 10  # Add 10 points for exact match
                }
            }
        }
        
        # Calculate final score as percentage of matches
        if ($totalParams -gt 0) {
            $finalScore = ($rowScore / ($totalParams * 10)) * 100
            Write-DebugLog "Row $rowIndex score: $finalScore% ($matchCount/$totalParams matches)" -Category "ViewDetails"
            
            if ($finalScore -gt $bestScore) {
                $bestScore = $finalScore
                $bestRow = $rowIndex
            }
        }
    }
    
    # Highlight the best matching row if score is high enough
    if ($bestRow -ge 0 -and $bestScore -gt 0) {
        Write-DebugLog "Best match found at row $bestRow with score $bestScore%" -Category "ViewDetails"
        
        # Highlight input columns
        foreach ($paramName in $MachineMetrics.Keys) {
            if ($paramName -ne "MachineType" -and $paramName -ne "MachineNumber" -and 
                $paramName -ne "Total (hrs/yr)" -and $paramName -ne "Operational Maintenance (hrs/yr)" -and
                $columnNames -contains $paramName) {
                $DataGridView.Rows[$bestRow].Cells[$paramName].Style.BackColor = $inputHighlightColor
            }
        }
        
        # Highlight output columns
        if ($columnNames -contains "Total (hrs/yr)") {
            $DataGridView.Rows[$bestRow].Cells["Total (hrs/yr)"].Style.BackColor = $totalHighlightColor
        }
        
        if ($columnNames -contains "Operational Maintenance (hrs/yr)") {
            $DataGridView.Rows[$bestRow].Cells["Operational Maintenance (hrs/yr)"].Style.BackColor = $totalHighlightColor
        }
    }
    else {
        Write-DebugLog "No matching row found in DataGridView with sufficient score" -Category "ViewDetails"
    }
}

################################################################################
#                  STAFFING TABLE FUNCTIONS                                    #
################################################################################

function Show-StaffingTableDialog {
    param (
        [string]$MachineID,
        [string]$MMO,
        [string]$LookupTablePath,
        [System.Data.DataTable]$ExistingData = $null,
        [string]$ClassCode = "",
        [string]$MachineAcronym = ""
    )
    
    Write-DebugLog "Starting Show-StaffingTableDialog for Machine ID: $MachineID, MMO: $MMO, Class Code: $ClassCode, Machine Acronym: $MachineAcronym" -Category "StaffingTable"
    
    # Extract machine acronym from machine ID if not provided
    if ([string]::IsNullOrWhiteSpace($MachineAcronym) -and $MachineID -match "^([A-Za-z]+)") {
        $MachineAcronym = $matches[1]
        Write-DebugLog "Extracted machine acronym from ID: $MachineAcronym" -Category "StaffingTable"
    }
    
    # Create dialog form
    $staffingForm = New-Object System.Windows.Forms.Form
    $staffingForm.Text = "Staffing Table - $MachineID"
    $staffingForm.Size = New-Object System.Drawing.Size(800, 600)
    $staffingForm.StartPosition = "CenterScreen"
    $staffingForm.FormBorderStyle = "Sizable"
    $staffingForm.MinimumSize = New-Object System.Drawing.Size(600, 400)
    
    # Create DataGridView for staffing data
    $staffingGrid = New-Object System.Windows.Forms.DataGridView
    $staffingGrid.Location = New-Object System.Drawing.Point(10, 50) # Leave space for info label
    $staffingGrid.Size = New-Object System.Drawing.Size(760, 460)
    $staffingGrid.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor
                           [System.Windows.Forms.AnchorStyles]::Bottom -bor
                           [System.Windows.Forms.AnchorStyles]::Left -bor
                           [System.Windows.Forms.AnchorStyles]::Right
    $staffingGrid.AllowUserToAddRows = $true
    $staffingGrid.AllowUserToDeleteRows = $true
    $staffingGrid.AutoSizeColumnsMode = "Fill"
    
    # Add info label at the top
    $infoLabel = New-Object System.Windows.Forms.Label
    $infoLabel.Location = New-Object System.Drawing.Point(10, 10)
    $infoLabel.Size = New-Object System.Drawing.Size(760, 30)
    $infoLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $infoLabel.Font = New-Object System.Drawing.Font($infoLabel.Font, [System.Drawing.FontStyle]::Bold)
    
    # Set info text showing MMO and Class Code 
    $infoText = "MMO: $MMO"
    if (-not [string]::IsNullOrWhiteSpace($ClassCode)) {
        $infoText += " | Class Code: $ClassCode"
    }
    if (-not [string]::IsNullOrWhiteSpace($MachineAcronym)) {
        $infoText += " | Machine Type: $MachineAcronym"
    }
    $infoLabel.Text = $infoText
    
    # Add info about saving location
    $saveInfoLabel = New-Object System.Windows.Forms.Label
    $saveInfoLabel.Location = New-Object System.Drawing.Point(10, 520)
    $saveInfoLabel.Size = New-Object System.Drawing.Size(430, 40)
    $saveInfoLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $saveInfoLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor
                           [System.Windows.Forms.AnchorStyles]::Left
    
    # Set save info text showing where the file will be saved
    $fileName = if (-not [string]::IsNullOrWhiteSpace($ClassCode)) {
        "$MMO-$ClassCode-Staffing-Table.csv"
    } else {
        "$MMO-Staffing-Table.csv"
    }
    
    $dirSearchPattern = if (-not [string]::IsNullOrWhiteSpace($ClassCode)) {
        if (-not [string]::IsNullOrWhiteSpace($MachineAcronym)) {
            "*$MMO*$MachineAcronym*-$ClassCode*"
        } else {
            "*$MMO*-$ClassCode*"
        }
    } else {
        "*$MMO*"
    }
    
    $saveInfoLabel.Text = "Will save as: $fileName`nIn directory matching: $dirSearchPattern"
    
    # Add buttons
    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Text = "Save"
    $btnSave.Size = New-Object System.Drawing.Size(100, 30)
    $btnSave.Location = New-Object System.Drawing.Point(670, 520)
    $btnSave.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor
                       [System.Windows.Forms.AnchorStyles]::Right
    
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Size = New-Object System.Drawing.Size(100, 30)
    $btnCancel.Location = New-Object System.Drawing.Point(560, 520)
    $btnCancel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor
                        [System.Windows.Forms.AnchorStyles]::Right
    
    # Create DataTable for grid
    $staffingTable = New-Object System.Data.DataTable
    
    # Determine input and output columns from lookup table if available
    $inputColumns = @()
    $outputColumns = @("Total (hrs/yr)", "Operational Maintenance (hrs/yr)")
    
    try {
        if (Test-Path $LookupTablePath) {
            Write-DebugLog "Getting columns from lookup table: $LookupTablePath" -Category "StaffingTable"
            
            # Load lookup table to determine input columns
            $lookupData = Import-Csv $LookupTablePath
            
            if ($lookupData -and $lookupData.Count -gt 0) {
                $lookupColumns = $lookupData[0].PSObject.Properties.Name
                
                # Filter out output columns to get input columns
                $inputColumns = $lookupColumns | Where-Object { $_ -notin $outputColumns }
                Write-DebugLog "Lookup table input columns: $($inputColumns -join ', ')" -Category "StaffingTable"
            }
        }
    }
    catch {
        Write-DebugLog "Error reading lookup table: $_" -Category "StaffingTable"
    }
    
    # Add Machine ID column first
    $staffingTable.Columns.Add("Machine ID", [string]) | Out-Null
    
    # Add standard Input columns from lookup table
    foreach ($column in $inputColumns) {
        $staffingTable.Columns.Add($column, [string]) | Out-Null
    }
    
    # Add the specific machine type columns
    $staffingTable.Columns.Add("MM7", [string]) | Out-Null
    $staffingTable.Columns.Add("MPE9", [string]) | Out-Null
    $staffingTable.Columns.Add("ET10", [string]) | Out-Null
    
    # Add Output columns
    foreach ($column in $outputColumns) {
        $staffingTable.Columns.Add($column, [string]) | Out-Null
    }
    
    # If existing data is provided, fill the datatable
    if ($ExistingData) {
        Write-DebugLog "Loading existing staffing data" -Category "StaffingTable"
        
        # Check if the data has a row with matching Machine ID
        $machineRow = $null
        foreach ($row in $ExistingData.Rows) {
            if ($row["Machine ID"] -eq $MachineID) {
                $machineRow = $row
                break
            }
        }
        
        if ($machineRow) {
            # Copy the matching row
            $newRow = $staffingTable.NewRow()
            
            # Copy values for all matching columns
            foreach ($column in $staffingTable.Columns) {
                $columnName = $column.ColumnName
                
                if ($ExistingData.Columns.Contains($columnName) -and $machineRow[$columnName] -ne $null) {
                    $newRow[$columnName] = $machineRow[$columnName]
                }
            }
            
            $staffingTable.Rows.Add($newRow)
            Write-DebugLog "Added existing row for Machine ID: $MachineID" -Category "StaffingTable"
        }
        else {
            # Add a new row with just the Machine ID
            $newRow = $staffingTable.NewRow()
            $newRow["Machine ID"] = $MachineID
            $staffingTable.Rows.Add($newRow)
            Write-DebugLog "Added new row with Machine ID: $MachineID" -Category "StaffingTable"
        }
    }
    else {
        # Add a new row with just the Machine ID
        $newRow = $staffingTable.NewRow()
        $newRow["Machine ID"] = $MachineID
        $staffingTable.Rows.Add($newRow)
        Write-DebugLog "Added new row with Machine ID: $MachineID" -Category "StaffingTable"
    }
    
    # Set the DataGridView's data source
    $staffingGrid.DataSource = $staffingTable
    
    # Configure the Machine ID column to be read-only
    if ($staffingGrid.Columns.Count -gt 0) {
        $staffingGrid.Columns["Machine ID"].ReadOnly = $true
    }
    
    # Add controls to form
    $staffingForm.Controls.Add($staffingGrid)
    $staffingForm.Controls.Add($btnSave)
    $staffingForm.Controls.Add($btnCancel)
    $staffingForm.Controls.Add($infoLabel)
    $staffingForm.Controls.Add($saveInfoLabel)
    
    # Event handlers
    $btnSave.Add_Click({
        # Prepare to save staffing data - PASS THE CLASS CODE AND MACHINE ACRONYM
        $saveResult = Save-StaffingTable -DataTable $staffingTable -MMO $MMO -MachineID $MachineID -ClassCode $ClassCode -MachineAcronym $MachineAcronym
        
        if ($saveResult) {
            [System.Windows.Forms.MessageBox]::Show(
                "Staffing table saved successfully!",
                "Success",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            $staffingForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $staffingForm.Close()
        }
    })
    
    $btnCancel.Add_Click({
        $staffingForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $staffingForm.Close()
    })
    
    # Show dialog
    $result = $staffingForm.ShowDialog()
    return $result
}

function Save-StaffingTable {
    param (
        [System.Data.DataTable]$DataTable,
        [string]$MMO,
        [string]$MachineID,
        [string]$ClassCode = "",
        [string]$MachineAcronym = ""
    )
    
    Write-DebugLog "Starting Save-StaffingTable for MMO: $MMO, Machine ID: $MachineID, Class Code: $ClassCode, Machine Acronym: $MachineAcronym" -Category "StaffingTable"
    
    # Extract machine acronym from machine ID if not provided
    if ([string]::IsNullOrWhiteSpace($MachineAcronym) -and $MachineID -match "^([A-Za-z]+)") {
        $MachineAcronym = $matches[1]
        Write-DebugLog "Extracted machine acronym from ID: $MachineAcronym" -Category "StaffingTable"
    }
    
    try {
        # Find the MMO directory
        $baseDir = "C:\Users\JR\Documents\Work\Programming Projects\LaborChecklist\Machine Labor Rubrics"
        
        # Check if base directory exists
        if (-not (Test-Path $baseDir)) {
            $baseDir = "C:\Users\JR\Documents\Work\Programming Projects\LaborChecklist\Machine Labor Rubrics"
            
            # If still doesn't exist, try to create it
            if (-not (Test-Path $baseDir)) {
                New-Item -Path $baseDir -ItemType Directory -Force | Out-Null
                Write-DebugLog "Created base directory: $baseDir" -Category "StaffingTable"
            }
        }
        
        # [Directory search logic - keeping original code]
        $mmoDirectory = $null
        $exactMatch = $false
        
        # Search patterns (keeping existing code)
        if (-not [string]::IsNullOrWhiteSpace($ClassCode) -and -not [string]::IsNullOrWhiteSpace($MachineAcronym)) {
            $searchPattern = "*$MMO*$MachineAcronym*-$ClassCode*"
            Write-DebugLog "Searching for directory with pattern: $searchPattern" -Category "StaffingTable"
            
            $mmoDirectories = Get-ChildItem $baseDir -Directory |
                Where-Object { $_.Name -like $searchPattern }
                
            if ($mmoDirectories.Count -gt 0) {
                $mmoDirectory = $mmoDirectories[0].FullName
                $exactMatch = $true
                Write-DebugLog "Found exact directory match: $mmoDirectory" -Category "StaffingTable"
            }
        }
        
        # [Directory creation logic - keeping original code]
        # Other logic for finding/creating directory remains the same
        if (-not $exactMatch -and -not [string]::IsNullOrWhiteSpace($ClassCode)) {
            $searchPattern = "*$MMO*-$ClassCode*"
            $mmoDirectories = Get-ChildItem $baseDir -Directory |
                Where-Object { $_.Name -like $searchPattern }
                
            if ($mmoDirectories.Count -gt 0) {
                $mmoDirectory = $mmoDirectories[0].FullName
                $exactMatch = $true
            }
        }
        
        if (-not $exactMatch) {
            $mmoDirectories = Get-ChildItem $baseDir -Directory |
                Where-Object { $_.Name -like "*$MMO*" }
                
            if ($mmoDirectories.Count -gt 0) {
                $mmoDirectory = $mmoDirectories[0].FullName
            }
        }
        
        if (-not $mmoDirectory) {
            if (-not [string]::IsNullOrWhiteSpace($MachineAcronym) -and -not [string]::IsNullOrWhiteSpace($ClassCode)) {
                $dirName = "$MMO-$MachineAcronym-$ClassCode"
                $mmoDirectory = Join-Path $baseDir $dirName
            }
            elseif (-not [string]::IsNullOrWhiteSpace($ClassCode)) {
                $dirName = "$MMO-$ClassCode"
                $mmoDirectory = Join-Path $baseDir $dirName
            }
            else {
                $mmoDirectory = Join-Path $baseDir "$MMO"
            }
            
            New-Item -Path $mmoDirectory -ItemType Directory -Force | Out-Null
        }
        
        # Define the staffing table CSV path with class code if available
        $staffingFileName = if (-not [string]::IsNullOrWhiteSpace($ClassCode)) {
            "$MMO-$ClassCode-Staffing-Table.csv"
        } else {
            "$MMO-Staffing-Table.csv"
        }
        
        $staffingFilePath = Join-Path $mmoDirectory $staffingFileName
        
        Write-DebugLog "Staffing table file path: $staffingFilePath" -Category "StaffingTable"
        
        # Check if the file already exists and load existing data
        $allMachines = @()
        
        if (Test-Path $staffingFilePath) {
            Write-DebugLog "Existing staffing table found, loading data" -Category "StaffingTable"
            try {
                # Load existing data - using custom parsing to avoid quotes
                $lines = Get-Content -Path $staffingFilePath -ErrorAction Stop
                
                if ($lines.Count -gt 0) {
                    # Parse header line
                    $headers = $lines[0].Split(',') | ForEach-Object { $_.Trim('"').Trim() }
                    
                    # Process data rows (skip header)
                    for ($i = 1; $i -lt $lines.Count; $i++) {
                        $line = $lines[$i].Trim()
                        if ([string]::IsNullOrWhiteSpace($line)) { continue }
                        
                        $values = $line.Split(',') | ForEach-Object { $_.Trim('"').Trim() }
                        
                        # Create object with properties
                        $rowObj = New-Object PSObject
                        
                        # Add each field as a property
                        for ($j = 0; $j -lt [Math]::Min($headers.Count, $values.Count); $j++) {
                            $propName = $headers[$j]
                            $propValue = $values[$j]
                            $rowObj | Add-Member -MemberType NoteProperty -Name $propName -Value $propValue
                        }
                        
                        # Add missing fields with empty values
                        for ($j = $values.Count; $j -lt $headers.Count; $j++) {
                            $propName = $headers[$j]
                            $rowObj | Add-Member -MemberType NoteProperty -Name $propName -Value ""
                        }
                        
                        # Skip the machine we're updating
                        if ($rowObj.'Machine ID' -ne $MachineID) {
                            $allMachines += $rowObj
                        }
                    }
                }
            }
            catch {
                Write-DebugLog "Error loading existing staffing data: $_" -Category "StaffingTable"
                # Continue with just our new data
            }
        }
        
        # Convert DataTable rows to PSObjects
        foreach ($row in $DataTable.Rows) {
            $rowObj = New-Object PSObject
            
            foreach ($column in $DataTable.Columns) {
                $rowObj | Add-Member -MemberType NoteProperty -Name $column.ColumnName -Value $row[$column.ColumnName]
            }
            
            $allMachines += $rowObj
        }
        
        Write-DebugLog "Saving staffing table with $($allMachines.Count) total rows" -Category "StaffingTable"
        
        # Write CSV WITHOUT quotes - custom implementation
        try {
            # Get all unique property names
            $propertyNames = @()
            foreach ($obj in $allMachines) {
                foreach ($prop in $obj.PSObject.Properties) {
                    if ($propertyNames -notcontains $prop.Name) {
                        $propertyNames += $prop.Name
                    }
                }
            }
            
            # Ensure Machine ID is the first column
            if ($propertyNames -contains "Machine ID") {
                $propertyNames = @("Machine ID") + ($propertyNames | Where-Object { $_ -ne "Machine ID" })
            }
            
            # Create header row
            $headerRow = $propertyNames -join ","
            
            # Create content with header row
            $content = @($headerRow)
            
            # Add data rows
            foreach ($obj in $allMachines) {
                $rowValues = @()
                
                foreach ($propName in $propertyNames) {
                    $value = if ($obj.PSObject.Properties.Name -contains $propName) {
                        # Clean value: replace commas with spaces to avoid CSV parsing issues
                        if ($obj.$propName -ne $null) {
                            ($obj.$propName).ToString() -replace ",", " "
                        } else {
                            ""
                        }
                    } else {
                        ""
                    }
                    
                    $rowValues += $value
                }
                
                $content += ($rowValues -join ",")
            }
            
            # Write content to file
            Set-Content -Path $staffingFilePath -Value $content
            
            Write-DebugLog "Successfully saved staffing data without quotes" -Category "StaffingTable"
        }
        catch {
            Write-DebugLog "Error writing CSV without quotes: $_" -Category "StaffingTable"
            throw
        }
        
        return $true
    }
    catch {
        Write-DebugLog "Error saving staffing table: $_" -Category "StaffingTable"
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to save staffing table: $_",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return $false
    }
}

function Load-StaffingTable {
    param (
        [string]$MMO,
        [string]$MachineID,
        [string]$ClassCode = "",
        [string]$MachineAcronym = "",
        [switch]$LoadAllMachines = $false
    )
    
    Write-DebugLog "Starting Load-StaffingTable for MMO: $MMO, Machine ID: $MachineID, Class Code: $ClassCode" -Category "StaffingTable"
    
    try {
        # Base directory path
        $baseDir = "C:\Users\JR\Documents\Work\Programming Projects\LaborChecklist\Machine Labor Rubrics"
        
        if (-not (Test-Path $baseDir)) {
            Write-DebugLog "Base directory not found: $baseDir" -Category "StaffingTable"
            return $null
        }

        # Find MMO directory (exact match)
        $mmoDirectory = Get-ChildItem $baseDir -Directory | 
            Where-Object { $_.Name -like "*$MMO*" -and $_.Name -like "*-$ClassCode" } |
            Select-Object -First 1 -ExpandProperty FullName

        if (-not $mmoDirectory) {
            Write-DebugLog "No matching MMO directory found for MMO: $MMO, Class Code: $ClassCode" -Category "StaffingTable"
            return $null
        }

        # Staffing file path (only one possible pattern)
        $staffingFilePath = Join-Path $mmoDirectory "$MMO-$ClassCode-Staffing-Table.csv"

        if (-not (Test-Path $staffingFilePath)) {
            Write-DebugLog "Staffing file not found: $staffingFilePath" -Category "StaffingTable"
            return $null
        }

        Write-DebugLog "Loading staffing data from: $staffingFilePath" -Category "StaffingTable"

        # Read CSV data
        $staffingData = Import-Csv -Path $staffingFilePath

        if (-not $staffingData -or $staffingData.Count -eq 0) {
            Write-DebugLog "No data found in staffing file" -Category "StaffingTable"
            return $null
        }

        # Create DataTable
        $staffingTable = New-Object System.Data.DataTable

        # Add columns from CSV headers
        $headers = $staffingData[0].PSObject.Properties.Name
        foreach ($header in $headers) {
            $staffingTable.Columns.Add($header, [string]) | Out-Null
        }

        # Add rows to DataTable
        foreach ($row in $staffingData) {
            $newRow = $staffingTable.NewRow()
            foreach ($header in $headers) {
                $newRow[$header] = $row.$header
            }
            $staffingTable.Rows.Add($newRow)
        }

        # Filter for specific machine if needed
        if (-not $LoadAllMachines -and -not [string]::IsNullOrWhiteSpace($MachineID)) {
            $filteredTable = $staffingTable.Clone()
            foreach ($row in $staffingTable.Rows) {
                if ($row["Machine ID"].Trim() -eq $MachineID.Trim()) {
                    $filteredTable.ImportRow($row)
                }
            }
            if ($filteredTable.Rows.Count -gt 0) {
                $staffingTable = $filteredTable
            } else {
                Write-DebugLog "No matching rows for Machine ID: $MachineID" -Category "StaffingTable"
            }
        }

        return $staffingTable
    }
    catch {
        Write-DebugLog "Error loading staffing table: $_" -Category "StaffingTable"
        Write-DebugLog "Stack trace: $($_.ScriptStackTrace)" -Category "StaffingTable"
        return $null
    }
}
function Debug-StaffingTable {
    param (
        [string]$FilePath,
        [string]$MachineID
    )
    
    Write-Host "Debugging staffing table loading from: $FilePath"
    Write-Host "Looking for Machine ID: $MachineID"
    
    if (-not (Test-Path $FilePath)) {
        Write-Host "ERROR: File does not exist!"
        return
    }
    
    try {
        $data = Import-Csv $FilePath
        Write-Host "Successfully loaded CSV with $($data.Count) rows"
        
        foreach ($row in $data) {
            $rowID = $row.'Machine ID'
            Write-Host "Row Machine ID: '$rowID'"
            
            if ($rowID -eq $MachineID) {
                Write-Host "MATCH FOUND!"
                Write-Host "Full row data:"
                $row | Format-Table | Out-String | Write-Host
            }
        }
    }
    catch {
        Write-Host "ERROR: Failed to load or process CSV: $_"
    }
}


function Debug-StaffingTableBinding {
    param (
        [string]$MMO,
        [string]$MachineID,
        [string]$ClassCode = "",
        [string]$MachineAcronym = ""
    )
    
    Write-Host "============== STAFFING TABLE BINDING DEBUG ================"
    Write-Host "Parameters:"
    Write-Host "  MMO: '$MMO'"
    Write-Host "  Machine ID: '$MachineID'"
    Write-Host "  Class Code: '$ClassCode'"
    Write-Host "  Machine Acronym: '$MachineAcronym'"
    
    $baseDir = "C:\Users\JR\Documents\Work\Programming Projects\LaborChecklist\Machine Labor Rubrics"
    Write-Host "`nChecking base directory: $baseDir"
    $dirExists = Test-Path $baseDir
    Write-Host "  Base directory exists: $dirExists"
    
    if ($dirExists) {
        # Search for all matching directories
        $mmoPattern = "*$MMO*"
        $mmoDirectories = Get-ChildItem $baseDir -Directory | Where-Object { $_.Name -like $mmoPattern }
        
        Write-Host "`nFound $($mmoDirectories.Count) directories matching pattern '$mmoPattern':"
        foreach ($dir in $mmoDirectories) {
            Write-Host "  - $($dir.Name)"
            
            # Check for staffing files in this directory
            $staffingFiles = Get-ChildItem $dir.FullName -Filter "*Staffing-Table.csv"
            Write-Host "    Contains $($staffingFiles.Count) staffing files:"
            
            foreach ($file in $staffingFiles) {
                Write-Host "      - $($file.Name)"
                
                # Try to read the file content
                try {
                    $content = Get-Content $file.FullName -ErrorAction Stop
                    Write-Host "        File has $($content.Count) lines"
                    
                    if ($content.Count -gt 0) {
                        # Check header line
                        $headerLine = $content[0]
                        Write-Host "        Header: $headerLine"
                        
                        # Check for Machine ID column
                        if ($headerLine -like "*Machine ID*") {
                            Write-Host "        Contains 'Machine ID' column: Yes"
                            
                            # Check data rows for our Machine ID
                            $matchCount = 0
                            for ($i=1; $i -lt $content.Count; $i++) {
                                $line = $content[$i]
                                if ($line -like "*$MachineID*") {
                                    $matchCount++
                                    Write-Host "        Line $i contains '$MachineID': Yes"
                                    Write-Host "        Content: $line"
                                }
                            }
                            
                            Write-Host "        Found $matchCount rows with Machine ID '$MachineID'"
                        } else {
                            Write-Host "        Contains 'Machine ID' column: No"
                        }
                    }
                }
                catch {
                    Write-Host "        ERROR reading file: $_"
                }
            }
        }
        
        # Now try to load using Load-StaffingTable
        Write-Host "`nAttempting to load staffing table with Load-StaffingTable function:"
        
        try {
            $staffingTable = Load-StaffingTable -MMO $MMO -MachineID $MachineID -ClassCode $ClassCode -MachineAcronym $MachineAcronym
            
            if ($staffingTable -ne $null) {
                Write-Host "  Function returned DataTable with $($staffingTable.Rows.Count) rows"
                Write-Host "  Columns in returned DataTable:"
                
                foreach ($column in $staffingTable.Columns) {
                    Write-Host "    - $($column.ColumnName)"
                }
                
                Write-Host "  Row data for Machine ID '$MachineID':"
                $foundMachine = $false
                
                foreach ($row in $staffingTable.Rows) {
                    if ($row["Machine ID"] -eq $MachineID) {
                        $foundMachine = $true
                        $rowData = @()
                        foreach ($column in $staffingTable.Columns) {
                            $rowData += "$($column.ColumnName)='$($row[$column])'"
                        }
                        Write-Host "    $($rowData -join ", ")"
                    }
                }
                
                if (-not $foundMachine) {
                    Write-Host "    No rows found with exact Machine ID match"
                }
            }
            else {
                Write-Host "  Function returned NULL"
                
                # Try with LoadAllMachines
                Write-Host "`nTrying with LoadAllMachines=true:"
                $allMachinesTable = Load-StaffingTable -MMO $MMO -LoadAllMachines -ClassCode $ClassCode -MachineAcronym $MachineAcronym
                
                if ($allMachinesTable -ne $null) {
                    Write-Host "  LoadAllMachines returned DataTable with $($allMachinesTable.Rows.Count) rows"
                    Write-Host "  Machine IDs in table:"
                    
                    foreach ($row in $allMachinesTable.Rows) {
                        Write-Host "    - '$($row["Machine ID"])'"
                    }
                }
                else {
                    Write-Host "  LoadAllMachines also returned NULL"
                }
            }
        }
        catch {
            Write-Host "  ERROR calling Load-StaffingTable: $_"
            Write-Host "  Stack trace: $($_.ScriptStackTrace)"
        }
    }
    
    Write-Host "`n=============== END DEBUG ================"
}

################################################################################
#                         UI FORM SETUP                                        #
################################################################################

# Create main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Machine Entry System"
$form.Size = New-Object System.Drawing.Size(600, 520)
$form.StartPosition = "CenterScreen"
$form.MinimumSize = New-Object System.Drawing.Size(600, 400)
$form.MinimumSize = New-Object System.Drawing.Size(600, 520)
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor 
               [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

# ListView setup
$listView = New-Object System.Windows.Forms.ListView
$listView.Location = New-Object System.Drawing.Point(10, 10)
$listView.Size = New-Object System.Drawing.Size(560, ($form.ClientSize.Height - 300))
$listView.View = [System.Windows.Forms.View]::Details
$listView.FullRowSelect = $true

# Add all columns needed to support all possible input parameters
$listView.Columns.Add("Acronym", 80) | Out-Null
$listView.Columns.Add("Number", 60) | Out-Null
$listView.Columns.Add("Class Code", 80) | Out-Null
$listView.Columns.Add("MMO", 60) | Out-Null
$listView.Columns.Add("Days/Week", 70) | Out-Null         # Operation (days/wk)
$listView.Columns.Add("Tours/Day", 70) | Out-Null
$listView.Columns.Add("Stackers", 70) | Out-Null
$listView.Columns.Add("Inductions", 70) | Out-Null
$listView.Columns.Add("Transports", 70) | Out-Null
$listView.Columns.Add("LIM Modules", 70) | Out-Null
$listView.Columns.Add("Machine Type", 90) | Out-Null 
$listView.Columns.Add("Site", 90) | Out-Null
$listView.Columns.Add("PSM #", 60) | Out-Null
$listView.Columns.Add("Terminal Type", 90) | Out-Null
$listView.Columns.Add("Equipment Code", 90) | Out-Null
$listView.Columns.Add("Machines", 90) | Out-Null

################################################################################
#                    CONTEXT MENU FOR ORIGINAL VALUES                          #
################################################################################

# Create context menu for ListView
$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
$listView.ContextMenuStrip = $contextMenu

# Add menu item to show original values
$showOriginalMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem("Show Original Values")
$showOriginalMenuItem.Add_Click({
    if ($listView.SelectedItems.Count -gt 0) {
        $selectedItem = $listView.SelectedItems[0]
        
        if ($selectedItem.Tag -is [hashtable] -and $selectedItem.Tag.Adjusted) {
            $origDays = $selectedItem.Tag.OriginalDaysPerWeek
            $origTours = $selectedItem.Tag.OriginalToursPerDay
            $currentDays = $selectedItem.SubItems[4].Text
            $currentTours = $selectedItem.SubItems[5].Text
            
            [System.Windows.Forms.MessageBox]::Show(
                "Machine: $($selectedItem.SubItems[0].Text) $($selectedItem.SubItems[1].Text)`n`n" +
                "Original calculated values:`n" +
                "Days/Week: $origDays`n" +
                "Tours/Day: $origTours`n`n" +
                "Current standardized values:`n" +
                "Days/Week: $currentDays`n" +
                "Tours/Day: $currentTours",
                "Original vs. Standardized Values",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        }
        else {
            [System.Windows.Forms.MessageBox]::Show(
                "This machine is using its original calculated values.",
                "Original Values",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        }
    }
})

$contextMenu.Items.Add($showOriginalMenuItem)

# Add menu item to toggle between original and standardized values
$toggleValuesMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem("Toggle Values (Original/Standard)")
$toggleValuesMenuItem.Add_Click({
    if ($listView.SelectedItems.Count -gt 0) {
        $selectedItem = $listView.SelectedItems[0]
        
        if ($selectedItem.Tag -is [hashtable] -and $selectedItem.Tag.Adjusted) {
            # If currently showing standardized values, switch to original
            if ($selectedItem.Tag.ShowingStandard -ne $false) {
                # Store current values
                $selectedItem.Tag.StandardDaysPerWeek = $selectedItem.SubItems[4].Text
                $selectedItem.Tag.StandardToursPerDay = $selectedItem.SubItems[5].Text
                
                # Show original values
                $selectedItem.SubItems[4].Text = [string]$selectedItem.Tag.OriginalDaysPerWeek
                $selectedItem.SubItems[5].Text = [string]$selectedItem.Tag.OriginalToursPerDay
                $selectedItem.Tag.ShowingStandard = $false
                
                # Change text color to indicate these are original values
                $selectedItem.SubItems[4].ForeColor = [System.Drawing.Color]::Blue
                $selectedItem.SubItems[5].ForeColor = [System.Drawing.Color]::Blue
            }
            else {
                # Switch back to standardized values
                $selectedItem.SubItems[4].Text = $selectedItem.Tag.StandardDaysPerWeek
                $selectedItem.SubItems[5].Text = $selectedItem.Tag.StandardToursPerDay
                $selectedItem.Tag.ShowingStandard = $true
                
                # Change text color back to orange for adjusted values
                $selectedItem.SubItems[4].ForeColor = [System.Drawing.Color]::DarkOrange
                $selectedItem.SubItems[5].ForeColor = [System.Drawing.Color]::DarkOrange
            }
        }
    }
})

$contextMenu.Items.Add($toggleValuesMenuItem)

# Add tooltip to explain colors
$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.SetToolTip($listView, "Orange text = adjusted values from original calculations`nBlue text = original calculated values`nRight-click for more options")

################################################################################
#                         UI FORM RESIZE LOGIC                                 #
################################################################################

$form.Add_Resize({
    # Define spacing constants
    $margin = 10
    $verticalSpacing = 5
    $buttonGroupHeight = 40
    $inputGroupHeight = 120
 
    # Calculate available height for ListView
    $listViewHeight = [int]($form.ClientSize.Height - $inputGroupHeight - $buttonGroupHeight - ($margin * 3))
    
    # Set minimum sizes
    $minListViewHeight = 100
    $minFormWidth = [Math]::Max(600, $minButtonAreaWidth)
    $minFormHeight = 400
    
    # Enforce minimum dimensions
    if ($form.ClientSize.Width -lt $minFormWidth) { 
        $form.ClientSize = New-Object System.Drawing.Size($minFormWidth, $form.ClientSize.Height)
    }
    if ($form.ClientSize.Height -lt $minFormHeight) { 
        $form.ClientSize = New-Object System.Drawing.Size($form.ClientSize.Width, $minFormHeight)
    }
    if ($listViewHeight -lt $minListViewHeight) {
        $listViewHeight = $minListViewHeight
    }
 

    $listView.Location = New-Object System.Drawing.Point([int]$margin, [int]$margin)
    $listViewWidth = [int]($form.ClientSize.Width - (2 * $margin))
    
    $listView.Size = New-Object System.Drawing.Size($listViewWidth, [int]$listViewHeight)
    
    # Dynamically resize columns with explicit casting
    $totalWidth = [int]($listView.Width - 2)
    $acronymWidth = [int]($totalWidth * 0.15)
    $numberWidth = [int]($totalWidth * 0.08)
    $codeWidth = [int]($totalWidth * 0.12)
    $mmoWidth = [int]($totalWidth * 0.08)
    $metricWidth = [int](($totalWidth - $acronymWidth - $numberWidth - $codeWidth - $mmoWidth) / 6)
    
    # Set column widths to better accommodate the metrics
    $listView.Columns[0].Width = $acronymWidth
    $listView.Columns[1].Width = $numberWidth
    $listView.Columns[2].Width = $codeWidth
    $listView.Columns[3].Width = $mmoWidth
    
    if ($listView.Columns.Count -gt 4) {
        for ($i = 4; $i -lt $listView.Columns.Count; $i++) {
            $listView.Columns[$i].Width = $metricWidth
        }
    }
 
    # --- Input Controls Positioning ---
    $inputStartY = [int]($listView.Bottom + $margin)
    $controlX = [int]$margin
    $labelWidth = [int]110
    $inputWidth = [int]150
    
    # Position input controls
    $lblAcronym.Location = New-Object System.Drawing.Point([int]$controlX, [int]$inputStartY)
    $lblAcronym.Size = New-Object System.Drawing.Size([int]$labelWidth, 20)
    $cmbAcronym.Location = New-Object System.Drawing.Point([int]($controlX + $labelWidth + $margin), [int]$inputStartY)
    $cmbAcronym.Size = New-Object System.Drawing.Size([int]$inputWidth, 20)
    
    $lblNumber.Location = New-Object System.Drawing.Point([int]$controlX, [int]($inputStartY + 30))
    $lblNumber.Size = New-Object System.Drawing.Size([int]$labelWidth, 20)
    $txtNumber.Location = New-Object System.Drawing.Point([int]($controlX + $labelWidth + $margin), [int]($inputStartY + 30))
    $txtNumber.Size = New-Object System.Drawing.Size([int]$inputWidth, 20)
    
    $lblClassCode.Location = New-Object System.Drawing.Point([int]$controlX, [int]($inputStartY + 60))
    $lblClassCode.Size = New-Object System.Drawing.Size([int]$labelWidth, 20)
    $cmbClassCode.Location = New-Object System.Drawing.Point([int]($controlX + $labelWidth + $margin), [int]($inputStartY + 60))
    $cmbClassCode.Size = New-Object System.Drawing.Size([int]$inputWidth, 20)
 
    # --- Button Positioning ---
    $buttonY = $form.ClientSize.Height - $buttonHeight - $bottomMargin
    $totalButtonsWidth = ($buttonWidth * 7) + ($buttonSpacing * 6)  # Now 7 buttons with 6 spaces
    $startX = ($form.ClientSize.Width - $totalButtonsWidth) / 2

    $btnRestoreSession.Location = New-Object System.Drawing.Point([int]$startX, [int]$buttonY)
    $btnSave.Location = New-Object System.Drawing.Point([int]($startX + $buttonWidth + $buttonSpacing), [int]$buttonY)
    $btnImport.Location = New-Object System.Drawing.Point([int]($startX + ($buttonWidth + $buttonSpacing) * 2), [int]$buttonY)
    $btnExport.Location = New-Object System.Drawing.Point([int]($startX + ($buttonWidth + $buttonSpacing) * 3), [int]$buttonY)
    $btnConfigure.Location = New-Object System.Drawing.Point([int]($startX + ($buttonWidth + $buttonSpacing) * 4), [int]$buttonY)
    $btnViewDetails.Location = New-Object System.Drawing.Point([int]($startX + ($buttonWidth + $buttonSpacing) * 5), [int]$buttonY)
    $btnGenerateReport.Location = New-Object System.Drawing.Point([int]($startX + ($buttonWidth + $buttonSpacing) * 6), [int]$buttonY)

    # Set anchoring for all controls
    $listView.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor
                      [System.Windows.Forms.AnchorStyles]::Bottom -bor
                      [System.Windows.Forms.AnchorStyles]::Left -bor
                      [System.Windows.Forms.AnchorStyles]::Right
                      
    $cmbAcronym.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
    $cmbClassCode.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
    $txtNumber.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
    $lblAcronym.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
    $lblNumber.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
    $lblClassCode.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
    
    $btnSave.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom
    $btnImport.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom
    $btnConfigure.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom
    $btnViewDetails.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom
    $btnExport.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom
})

# Form controls
$lblAcronym = New-Object System.Windows.Forms.Label -Property @{
    Location = New-Object System.Drawing.Point(10, 220)
    Size = New-Object System.Drawing.Size(100, 20)
    Text = "Machine Acronym:"
}

$cmbAcronym = New-Object System.Windows.Forms.ComboBox -Property @{
    Location = New-Object System.Drawing.Point(120, 220)
    Size = New-Object System.Drawing.Size(150, 20)
    DropDownStyle = "DropDownList"
}
$cmbAcronym.Items.AddRange(($machineClassCodes.Keys | Sort-Object))

$lblNumber = New-Object System.Windows.Forms.Label -Property @{
    Location = New-Object System.Drawing.Point(10, 250)
    Size = New-Object System.Drawing.Size(100, 20)
    Text = "Machine Number:"
}

$txtNumber = New-Object System.Windows.Forms.TextBox -Property @{
    Location = New-Object System.Drawing.Point(120, 250)
    Size = New-Object System.Drawing.Size(150, 20)
}

$lblClassCode = New-Object System.Windows.Forms.Label -Property @{
    Location = New-Object System.Drawing.Point(10, 280)
    Size = New-Object System.Drawing.Size(100, 20)
    Text = "Class Code:"
}

$cmbClassCode = New-Object System.Windows.Forms.ComboBox -Property @{
    Location = New-Object System.Drawing.Point(120, 280)
    Size = New-Object System.Drawing.Size(150, 20)
    DropDownStyle = "DropDownList"
    Enabled = $false
}

# Standard button size
$buttonWidth = 120
$buttonHeight = 30
$buttonSpacing = 10
$bottomMargin = 20

# Calculate total width needed for all 7 buttons (including Restore Session and Generate Report)
$totalButtonsWidth = ($buttonWidth * 7) + ($buttonSpacing * 6)  # 7 buttons with 6 spaces

# Calculate starting X position to center the buttons
$startX = ($form.ClientSize.Width - $totalButtonsWidth) / 2

# Define Y position for all buttons (at bottom of form)
$buttonY = $form.ClientSize.Height - $buttonHeight - $bottomMargin

# Create buttons with updated positions
$btnRestoreSession = New-Object System.Windows.Forms.Button -Property @{
    Text = "Restore Session"
    Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
    Location = New-Object System.Drawing.Point($startX, $buttonY)
    Anchor = [System.Windows.Forms.AnchorStyles]::Bottom
}

$btnSave = New-Object System.Windows.Forms.Button -Property @{
    Text = "Save"
    Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
    Location = New-Object System.Drawing.Point(($startX + $buttonWidth + $buttonSpacing), $buttonY)
    Anchor = [System.Windows.Forms.AnchorStyles]::Bottom
}

$btnImport = New-Object System.Windows.Forms.Button -Property @{
    Text = "Import CSV"
    Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
    Location = New-Object System.Drawing.Point(($startX + ($buttonWidth + $buttonSpacing) * 2), $buttonY)
    Anchor = [System.Windows.Forms.AnchorStyles]::Bottom
}

$btnExport = New-Object System.Windows.Forms.Button -Property @{
    Text = "Export to CSV"
    Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
    Location = New-Object System.Drawing.Point(($startX + ($buttonWidth + $buttonSpacing) * 3), $buttonY)
    Anchor = [System.Windows.Forms.AnchorStyles]::Bottom
}

$btnConfigure = New-Object System.Windows.Forms.Button -Property @{
    Text = "Configure"
    Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
    Location = New-Object System.Drawing.Point(($startX + ($buttonWidth + $buttonSpacing) * 4), $buttonY)
    Anchor = [System.Windows.Forms.AnchorStyles]::Bottom
    Enabled = $false
}

$btnViewDetails = New-Object System.Windows.Forms.Button -Property @{
    Text = "View Details"
    Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
    Location = New-Object System.Drawing.Point(($startX + ($buttonWidth + $buttonSpacing) * 5), $buttonY)
    Anchor = [System.Windows.Forms.AnchorStyles]::Bottom
    Enabled = $false
}

$btnGenerateReport = New-Object System.Windows.Forms.Button -Property @{
    Text = "Generate Report"
    Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
    Location = New-Object System.Drawing.Point(($startX + ($buttonWidth + $buttonSpacing) * 6), $buttonY)
    Anchor = [System.Windows.Forms.AnchorStyles]::Bottom
}

# Add all buttons to the form
$form.Controls.AddRange(@($btnRestoreSession, $btnSave, $btnImport, $btnExport, $btnConfigure, $btnViewDetails, $btnGenerateReport))


################################################################################
#                     EVENT HANDLER FUNCTIONS                                  #
################################################################################

# Event Handlers
$cmbAcronym.Add_SelectedIndexChanged({
    Write-DebugLog "Acronym selected: '$($cmbAcronym.Text)'" -Category "Dropdown"
    
    if ($machineClassCodes.ContainsKey($cmbAcronym.Text)) {
        # Get unique class codes for selected acronym to verify it's a valid acronym
        $classCodes = $machineClassCodes[$cmbAcronym.Text]."Class Code" | Sort-Object -Unique
        Write-DebugLog "Found class codes: $($classCodes -join ', ')" -Category "Dropdown"
        
        # With class code removed from main form, just focus the number field
        $txtNumber.Select()
    } else {
        Write-DebugLog "No class codes found for acronym: '$($cmbAcronym.Text)'" -Category "Dropdown"
    }
})

$listView.Add_SelectedIndexChanged({
    if ($listView.SelectedItems.Count -gt 0) {
        $selected = $listView.SelectedItems[0]
        $cmbAcronym.Text = $selected.SubItems[0].Text
        $txtNumber.Text = $selected.SubItems[1].Text
        $cmbClassCode.Text = $selected.SubItems[2].Text
        $btnConfigure.Enabled = $true
        $btnViewDetails.Enabled = $listView.SelectedItems.Count -gt 0
    }
    else {
        $btnConfigure.Enabled = $false
    }
})

# Save Button Click Event
$btnSave.Add_Click({
    # Input validation
    if ([string]::IsNullOrWhiteSpace($cmbAcronym.Text) -or 
        [string]::IsNullOrWhiteSpace($txtNumber.Text)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Machine acronym and number are required!",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    $selectedAcronym = $cmbAcronym.Text
    $machineNumber = $txtNumber.Text.Trim()

    # Check for duplicates in ListView
    $isDuplicate = $false
    foreach ($item in $listView.Items) {
        if ($item.SubItems[0].Text -eq $selectedAcronym -and 
            $item.SubItems[1].Text -eq $machineNumber) {
            $isDuplicate = $true
            break
        }
    }

    if ($isDuplicate) {
        [System.Windows.Forms.MessageBox]::Show(
            "A machine with acronym '$selectedAcronym' and number '$machineNumber' already exists!",
            "Duplicate Entry",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    # Check if valid machine acronym
    if (-not $machineClassCodes.ContainsKey($selectedAcronym)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Invalid machine acronym!",
            "Validation Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    # Create a basic ListView item with just acronym and number
    $item = New-Object System.Windows.Forms.ListViewItem($selectedAcronym)
    $item.SubItems.Add($machineNumber)
    $item.SubItems.Add("")  # Empty class code
    $item.SubItems.Add("")  # Empty MMO
    $listView.Items.Add($item)
    
    # Select the newly added item
    $item.Selected = $true
    
    # Show a simple message to instruct the user to configure
    [System.Windows.Forms.MessageBox]::Show(
        "Machine entry added. Please use the Configure button to set up the machine parameters.",
        "Machine Added",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )

    # Clear inputs
    $cmbAcronym.SelectedIndex = -1
    $txtNumber.Clear()
})

# Import Button Logic - Modified to handle directory of CSV files
$btnImport.Add_Click({
    # Use FolderBrowserDialog instead of OpenFileDialog
    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderDialog.Description = "Select Directory Containing WebEOR CSV Files"
    $folderDialog.ShowNewFolderButton = $false
    
    if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        # Initialize progress form variables
        $progressForm = New-Object System.Windows.Forms.Form
        $progressForm.Text = "Combining and Processing CSV Files..."
        $progressForm.Size = New-Object System.Drawing.Size(400, 100)
        $progressForm.StartPosition = "CenterScreen"
        $progressForm.FormBorderStyle = "FixedDialog"
        $progressForm.ControlBox = $false

        $progressBar = New-Object System.Windows.Forms.ProgressBar
        $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
        $progressBar.Maximum = 100
        $progressBar.Value = 0
        $progressBar.Size = New-Object System.Drawing.Size(360, 20)
        $progressBar.Location = New-Object System.Drawing.Point(10, 10)

        $progressLabel = New-Object System.Windows.Forms.Label
        $progressLabel.Size = New-Object System.Drawing.Size(360, 20)
        $progressLabel.Location = New-Object System.Drawing.Point(10, 35)
        $progressLabel.Text = "Scanning directory..."

        $progressForm.Controls.Add($progressBar)
        $progressForm.Controls.Add($progressLabel)
        $progressForm.Show()
        [System.Windows.Forms.Application]::DoEvents()

        try {
            # Step 1: Combine CSV files
            $progressBar.Value = 5
            $progressLabel.Text = "Finding CSV files..."
            [System.Windows.Forms.Application]::DoEvents()
            
            $csvDirectory = $folderDialog.SelectedPath
            $outputFile = "WebEOR-Combined.csv"
            $combinedFilePath = Join-Path $csvDirectory $outputFile
            $header = "Site,MType,MNo,Op No.,Sort Program,Tour,Run#,Start,End,Fed,MODS,DOIS"

            # Verify directory exists
            if (-not (Test-Path $csvDirectory)) {
                throw "Directory not found: $csvDirectory"
            }

            # Get all CSV files (excluding any existing combined file)
            $csvFiles = Get-ChildItem -Path $csvDirectory -Filter "*.csv" -File | 
                Where-Object { $_.Name -ne $outputFile }
            
            if ($csvFiles.Count -eq 0) {
                throw "No CSV files found in selected directory"
            }

            $progressBar.Value = 10
            $progressLabel.Text = "Found $($csvFiles.Count) CSV files. Combining..."
            [System.Windows.Forms.Application]::DoEvents()

            # Create a HashSet to store unique rows (automatically handles duplicates)
            $allRows = New-Object System.Collections.Generic.HashSet[string]

            # Process each file
            $processedCount = 0
            $totalRows = 0
            
            foreach ($file in $csvFiles) {
                $fileProgress = 10 + (5 * ($processedCount / $csvFiles.Count))
                $progressBar.Value = $fileProgress
                $progressLabel.Text = "Processing: $($file.Name) ($($processedCount + 1) of $($csvFiles.Count))"
                [System.Windows.Forms.Application]::DoEvents()
                
                # Get content and skip header (assuming first line is header)
                $content = Get-Content $file.FullName | Select-Object -Skip 1
                $rowCount = 0
                
                foreach ($row in $content) {
                    if (-not [string]::IsNullOrWhiteSpace($row)) {
                        $null = $allRows.Add($row)
                        $rowCount++
                    }
                }
                
                $processedCount++
                $totalRows += $rowCount
                Write-DebugLog "Added $rowCount rows from $($file.Name)" -Category "Import"
            }

            # Write combined file with header and unique rows
            $progressBar.Value = 15
            $progressLabel.Text = "Writing combined file..."
            [System.Windows.Forms.Application]::DoEvents()
            
            $header | Set-Content -Path $combinedFilePath -Encoding UTF8
            $allRows | Add-Content -Path $combinedFilePath -Encoding UTF8

            $uniqueRows = $allRows.Count
            Write-DebugLog "Combined $processedCount CSV files with $totalRows total rows. After deduplication: $uniqueRows unique rows." -Category "Import"

            # Step 2: Process the combined CSV file (using existing logic)
            $progressBar.Value = 20
            $progressLabel.Text = "Loading combined data..."
            [System.Windows.Forms.Application]::DoEvents()
            
            # Import the combined CSV
            $importedData = Import-Csv -Path $combinedFilePath
            $global:ImportedData = $importedData
            
            $progressBar.Value = 25
            $progressLabel.Text = "Processing raw data..."
            [System.Windows.Forms.Application]::DoEvents()
            
            # Transform data - handle different possible column names
            $processedData = $importedData | ForEach-Object {
                # Handle different possible column naming conventions
                $machineType = if ($_.MType) { $_.MType } 
                              elseif ($_.'Machine Type') { $_.'Machine Type' }
                              elseif ($_.Type) { $_.Type }
                              else { "UNKNOWN" }
                
                $machineNo = if ($_.MNo) { $_.MNo }
                            elseif ($_.'Machine No') { $_.'Machine No' }
                            elseif ($_.Number) { $_.Number }
                            else { "000" }
                
                $startTime = if ($_.Start) { $_.Start }
                            elseif ($_.'Start Time') { $_.'Start Time' }
                            elseif ($_.Begin) { $_.Begin }
                            else { Get-Date }
                
                $tour = if ($_.Tour) { $_.Tour }
                       elseif ($_.'Tour#') { $_.'Tour#' }
                       elseif ($_.Shift) { $_.Shift }
                       else { 1 }
                
                $fed = if ($_.Fed) { $_.Fed }
                      elseif ($_.'Fed Count') { $_.'Fed Count' }
                      elseif ($_.Volume) { $_.Volume }
                      else { 0 }
                
                $runNo = if ($_.'Run#') { $_.'Run#' }
                        elseif ($_.RunNo) { $_.RunNo }
                        elseif ($_.'Run No.') { $_.'Run No.' }
                        elseif ($_.'Op No.') { $_.'Op No.' }
                        else { "R000" }

                $machineID = "$machineType $machineNo"
                $date = if ($startTime -is [DateTime]) { $startTime } else { 
                    try { [DateTime]::Parse($startTime) } catch { Get-Date }
                }
                
                [PSCustomObject]@{
                    Machine_ID = $machineID
                    MType = $machineType
                    MNo = $machineNo
                    Date = $date.Date
                    Tour = [int]$tour
                    Fed = [int]$fed
                    Run_No = $runNo
                }
            }
            
            $progressBar.Value = 35
            $progressLabel.Text = "Analyzing date ranges..."
            [System.Windows.Forms.Application]::DoEvents()
            
            # Get date range for the entire dataset
            $allDates = $processedData | Select-Object -ExpandProperty Date -Unique | Sort-Object
            if ($allDates.Count -gt 0) {
                $startDate = $allDates[0]
                $endDate = $allDates[-1]
                $totalDays = ($endDate - $startDate).Days + 1
                $totalWeeks = if ($totalDays -gt 0) { $totalDays / 7 } else { 1 }
                
                Write-DebugLog "Date range: $startDate to $endDate ($totalDays days, $([Math]::Round($totalWeeks, 1)) weeks)" -Category "Import"
                
                # Group data by machine
                $progressBar.Value = 45
                $progressLabel.Text = "Grouping data by machine..."
                [System.Windows.Forms.Application]::DoEvents()
                
                $machineGroups = $processedData | Group-Object -Property MType, MNo
                $totalMachines = $machineGroups.Count
                $machineMetrics = @{}
                
                Write-DebugLog "Found $totalMachines unique machines in combined data" -Category "Import"
                
                # Process each machine
                for ($i = 0; $i -lt $machineGroups.Count; $i++) {
                    $group = $machineGroups[$i]
                    $progress = 45 + [Math]::Min(35 * ($i / $totalMachines), 35)
                    $progressBar.Value = $progress
                    $progressLabel.Text = "Calculating metrics for machine $($i+1) of $totalMachines..."
                    [System.Windows.Forms.Application]::DoEvents()
                    
                    # Extract machine info
                    $machineInfo = $group.Name -split ', '
                    $mType = $machineInfo[0] -replace "'", "" -replace '"', ''  # Remove quotes
                    $mNo = $machineInfo[1] -replace "'", "" -replace '"', ''    # Remove quotes
                    
                    # Process this machine's data
                    $machineData = $group.Group
                    
                    # Group by date to get daily metrics
                    $dailyMetrics = $machineData | Group-Object Date | ForEach-Object {
                        $date = $_.Name
                        $dayData = $_.Group
                        
                        # Get distinct tours for this day (excluding 0 tours)
                        $distinctTours = $dayData | 
                            Where-Object { $_.Tour -gt 0 } | 
                            Select-Object -ExpandProperty Tour -Unique
                        
                        # Sum Fed for the day
                        $dailyFed = ($dayData | Measure-Object -Property Fed -Sum).Sum
                        
                        [PSCustomObject]@{
                            Date = $date
                            DistinctTours = $distinctTours.Count
                            Fed = $dailyFed
                        }
                    }
                    
                    # Calculate metrics
                    $activeDays = @($dailyMetrics | Where-Object { $_.DistinctTours -gt 0 })
                    $avgTours = if ($activeDays.Count -gt 0) {
                        ($activeDays | Measure-Object -Property DistinctTours -Average).Average
                    } else { 0 }
                    
                    $daysOperated = $activeDays.Count
                    $avgDaysPerWeek = if ($totalWeeks -gt 0) { $daysOperated / $totalWeeks } else { 0 }
                    
                    $totalFed = ($dailyMetrics | Measure-Object -Property Fed -Sum).Sum
                    $avgDailyFed = if ($totalDays -gt 0) { $totalFed / $totalDays } else { 0 }
                    $yearlyFed = $avgDailyFed * 365
                    
                    # Store metrics for this machine
                    $machineKey = "$mType $mNo"
                    $machineMetrics[$machineKey] = @{
                        MType = $mType
                        MNo = $mNo
                        AvgToursPerDay = [Math]::Round($avgTours, 1)
                        AvgDaysPerWeek = [Math]::Round($avgDaysPerWeek, 1)
                        YearlyFed = [Math]::Round($yearlyFed)
                        ActiveDays = $daysOperated
                        TotalRecords = $machineData.Count
                    }
                    
                    Write-DebugLog "Machine ${machineKey}: $($daysOperated) active days, avg $([Math]::Round($avgTours, 1)) tours/day, $([Math]::Round($avgDaysPerWeek, 1)) days/week" -Category "Import"
                }
                
                # Now add machines to ListView with calculated metrics
                $progressBar.Value = 85
                $progressLabel.Text = "Adding machines to list..."
                [System.Windows.Forms.Application]::DoEvents()
                
                $addedCount = 0
                $skippedCount = 0
                
                foreach ($machineKey in $machineMetrics.Keys) {
                    $metrics = $machineMetrics[$machineKey]
                    $mType = $metrics.MType
                    $mNo = $metrics.MNo
                    
                    # Skip machines with insufficient data
                    if ($metrics.ActiveDays -lt 3 -or $metrics.TotalRecords -lt 5) {
                        $skippedCount++
                        Write-DebugLog "Skipped $machineKey (insufficient data: $($metrics.ActiveDays) days, $($metrics.TotalRecords) records)" -Category "Import"
                        continue
                    }
                    
                    # Check for duplicates in ListView
                    $isDuplicate = $false
                    foreach ($existingItem in $listView.Items) {
                        if ($existingItem.SubItems[0].Text -eq $mType -and $existingItem.SubItems[1].Text -eq $mNo) {
                            $isDuplicate = $true
                            break
                        }
                    }
                    
                    if ($isDuplicate) {
                        $skippedCount++
                        Write-DebugLog "Skipped $machineKey (already exists in list)" -Category "Import"
                        continue
                    }
                    
                    # Create ListView item
                    $item = New-Object System.Windows.Forms.ListViewItem($mType)
                    $item.SubItems.Add([string]$mNo)
                    $item.SubItems.Add("")  # Empty class code
                    $item.SubItems.Add("")  # Empty MMO
                    
                    # Add calculated values
                    $item.SubItems.Add([string]$metrics.AvgDaysPerWeek)  # Days/Week
                    $item.SubItems.Add([string]$metrics.AvgToursPerDay)  # Tours/Day
                    
                    # Add remaining empty columns
                    $item.SubItems.Add("")  # Stackers
                    $item.SubItems.Add("")  # Inductions
                    $item.SubItems.Add("")  # Transports  
                    $item.SubItems.Add("")  # LIM Modules
                    $item.SubItems.Add("")  # Machine Type
                    $item.SubItems.Add("")  # Site
                    $item.SubItems.Add("")  # PSM #
                    $item.SubItems.Add("")  # Terminal Type
                    $item.SubItems.Add("")  # Equipment Code
                    $item.SubItems.Add("")  # Machines
                    
                    # Store original values in Tag property
                    $item.Tag = @{
                        OriginalDaysPerWeek = $metrics.AvgDaysPerWeek
                        OriginalToursPerDay = $metrics.AvgToursPerDay
                        Adjusted = $false
                        DataQuality = @{
                            ActiveDays = $metrics.ActiveDays
                            TotalRecords = $metrics.TotalRecords
                            YearlyFed = $metrics.YearlyFed
                        }
                    }

                    $listView.Items.Add($item)
                    $addedCount++
                }
                
                $progressBar.Value = 95
                $progressLabel.Text = "Exporting data..."
                [System.Windows.Forms.Application]::DoEvents()
                
                # Export ListView data to CSV
                $exportData = @()
                foreach ($item in $listView.Items) {
                    $exportData += [PSCustomObject]@{
                        Acronym = $item.SubItems[0].Text
                        Number = $item.SubItems[1].Text
                        ClassCode = $item.SubItems[2].Text
                        MMO = $item.SubItems[3].Text
                        DaysPerWeek = $item.SubItems[4].Text
                        ToursPerDay = $item.SubItems[5].Text
                        Stackers = $item.SubItems[6].Text
                        Inductions = $item.SubItems[7].Text
                        Transports = $item.SubItems[8].Text
                        LIMModules = $item.SubItems[9].Text
                        MachineType = $item.SubItems[10].Text
                        Site = $item.SubItems[11].Text
                        PSMN = $item.SubItems[12].Text
                        TerminalType = $item.SubItems[13].Text
                        EquipmentCode = $item.SubItems[14].Text
                        Machines = $item.SubItems[15].Text
                    }
                }
                
                # Create export file name
                $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
                $exportPath = Join-Path $csvDirectory "WebEOR-Processed-$timestamp.csv"
                
                # Export the data
                $exportData | Export-Csv -Path $exportPath -NoTypeInformation
                
                $progressBar.Value = 100
                $progressLabel.Text = "Completing import..."
                [System.Windows.Forms.Application]::DoEvents()
                
                # Auto-resize columns
                $listView.Columns | ForEach-Object { $_.Width = -2 }
                
                # Show results
                $progressForm.Close()
                
                # Create summary message
                $summaryMessage = @"
Import Complete!

Combined Files: $processedCount CSV files
Total Data Rows: $totalRows (before deduplication)
Unique Data Rows: $uniqueRows (after deduplication)
Machines Found: $totalMachines
Machines Added: $addedCount
Machines Skipped: $skippedCount (duplicates or insufficient data)

Files Created:
- Combined data: $combinedFilePath
- Processed data: $exportPath

Date Range: $($startDate.ToString('yyyy-MM-dd')) to $($endDate.ToString('yyyy-MM-dd'))
Analysis Period: $totalDays days ($([Math]::Round($totalWeeks, 1)) weeks)
"@
                
                [System.Windows.Forms.MessageBox]::Show(
                    $summaryMessage,
                    "Import Success",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )
                
            } else {
                $progressForm.Close()
                [System.Windows.Forms.MessageBox]::Show(
                    "The combined data file contains no valid dates.",
                    "Error",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
            }
        }
        catch {
            if ($progressForm -and $progressForm.Visible) {
                $progressForm.Close()
            }
            
            $errorMessage = "Import failed: $_"
            Write-DebugLog $errorMessage -Category "Import"
            
            [System.Windows.Forms.MessageBox]::Show(
                $errorMessage,
                "Import Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    }
})

$btnExport.Add_Click({
    if ($listView.Items.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "No data to export!",
            "Warning",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }
    
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV files (*.csv)|*.csv"
    $saveFileDialog.Title = "Export Machine Data"
    $saveFileDialog.DefaultExt = "csv"
    
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            # Prepare data for export
            $exportData = @()
            foreach ($item in $listView.Items) {
                # Create basic object with visible columns
                $rowData = [PSCustomObject]@{
                    Acronym = $item.SubItems[0].Text
                    Number = $item.SubItems[1].Text
                    ClassCode = $item.SubItems[2].Text
                    MMO = $item.SubItems[3].Text
                }
                
                # Handle cases where not all columns exist
                if ($item.SubItems.Count -gt 4) {
                    $rowData | Add-Member -NotePropertyName "DaysPerWeek" -NotePropertyValue $item.SubItems[4].Text
                }
                if ($item.SubItems.Count -gt 5) {
                    $rowData | Add-Member -NotePropertyName "ToursPerDay" -NotePropertyValue $item.SubItems[5].Text
                }
                if ($item.SubItems.Count -gt 6) {
                    $rowData | Add-Member -NotePropertyName "Stackers" -NotePropertyValue $item.SubItems[6].Text
                }
                if ($item.SubItems.Count -gt 7) {
                    $rowData | Add-Member -NotePropertyName "Inductions" -NotePropertyValue $item.SubItems[7].Text
                }
                if ($item.SubItems.Count -gt 8) {
                    $rowData | Add-Member -NotePropertyName "Transports" -NotePropertyValue $item.SubItems[8].Text
                }
                if ($item.SubItems.Count -gt 9) {
                    $rowData | Add-Member -NotePropertyName "LIMModules" -NotePropertyValue $item.SubItems[9].Text
                }
                if ($item.SubItems.Count -gt 10) {
                    $rowData | Add-Member -NotePropertyName "MachineType" -NotePropertyValue $item.SubItems[10].Text
                }
                if ($item.SubItems.Count -gt 11) {
                    $rowData | Add-Member -NotePropertyName "Site" -NotePropertyValue $item.SubItems[11].Text
                }
                if ($item.SubItems.Count -gt 12) {
                    $rowData | Add-Member -NotePropertyName "PSMN" -NotePropertyValue $item.SubItems[12].Text
                }
                if ($item.SubItems.Count -gt 13) {
                    $rowData | Add-Member -NotePropertyName "TerminalType" -NotePropertyValue $item.SubItems[13].Text
                }
                if ($item.SubItems.Count -gt 14) {
                    $rowData | Add-Member -NotePropertyName "EquipmentCode" -NotePropertyValue $item.SubItems[14].Text
                }

                if ($item.SubItems.Count -gt 15) {
                    $rowData | Add-Member -NotePropertyName "Machines" -NotePropertyValue $item.SubItems[15].Text
                }
                
                # Add original values if available in the Tag
                if ($item.Tag -is [hashtable]) {
                    # Add original Days Per Week if available
                    if ($item.Tag.ContainsKey("OriginalDaysPerWeek")) {
                        $rowData | Add-Member -NotePropertyName "OriginalDaysPerWeek" -NotePropertyValue $item.Tag.OriginalDaysPerWeek
                    }
                    
                    # Add original Tours Per Day if available
                    if ($item.Tag.ContainsKey("OriginalToursPerDay")) {
                        $rowData | Add-Member -NotePropertyName "OriginalToursPerDay" -NotePropertyValue $item.Tag.OriginalToursPerDay
                    }
                    
                    # Add adjustment flag if available
                    if ($item.Tag.ContainsKey("Adjusted")) {
                        $rowData | Add-Member -NotePropertyName "ValuesAdjusted" -NotePropertyValue $item.Tag.Adjusted
                    }
                    
                    # Add which values are currently being shown (original or standard)
                    if ($item.Tag.ContainsKey("ShowingStandard")) {
                        $rowData | Add-Member -NotePropertyName "ShowingStandardValues" -NotePropertyValue $item.Tag.ShowingStandard
                    }
                }
                
                $exportData += $rowData
            }
            
            # Export to CSV
            $exportData | Export-Csv -Path $saveFileDialog.FileName -NoTypeInformation
            
            [System.Windows.Forms.MessageBox]::Show(
                "Data exported successfully!`n`nThe export includes both standardized and original calculated values where applicable.",
                "Success",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Failed to export data: $_",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    }
})

# Configure Button Logic
$btnConfigure.Add_Click({
    if ($listView.SelectedItems.Count -eq 0) { return }
    
    $selectedItem = $listView.SelectedItems[0]
    $acronym = $selectedItem.SubItems[0].Text
    $machineNumber = $selectedItem.SubItems[1].Text
    $classCode = $selectedItem.SubItems[2].Text
    $mmo = $selectedItem.SubItems[3].Text
    
    # Get initial values from ListView
    $initialValues = @{}
    
    # Define the mappings between column index and parameter name
    $columnMap = @{
        4 = "Operation (days/wk)"
        5 = "Tours/Day"
        6 = "Stackers"
        7 = "Inductions"
        8 = "Transports"
        9 = "LIM Modules"
        10 = "Machine Type"
        11 = "Site"
        12 = "PSM #"
        13 = "Terminal Type"
        14 = "Equipment Code"
        15 = "Machines"
    }
    
    # Add all available metrics from SubItems based on column map
    foreach ($index in $columnMap.Keys) {
        if ($selectedItem.SubItems.Count -gt $index -and -not [string]::IsNullOrWhiteSpace($selectedItem.SubItems[$index].Text)) {
            $initialValues[$columnMap[$index]] = $selectedItem.SubItems[$index].Text
            Write-DebugLog "Added initial value for $($columnMap[$index]): $($selectedItem.SubItems[$index].Text)" -Category "Configure"
        }
    }
    
    # Find the lookup table path for this MMO/machine
    $lookupTablePath = ""
    if (-not [string]::IsNullOrWhiteSpace($mmo)) {
        $baseDir = "C:\Users\JR\Documents\Work\Programming Projects\LaborChecklist\Machine Labor Rubrics"
        
        # Try alternate path if first one doesn't exist
        if (-not (Test-Path $baseDir)) {
            $baseDir = "C:\Users\JR\Documents\Work\Programming Projects\LaborChecklist\Machine Labor Rubrics"
        }
        
        # First try to find a directory that matches the MMO and class code
        $mmoDirectories = Get-ChildItem $baseDir -Directory | 
            Where-Object { $_.Name -like "*$mmo*" -and $_.Name -like "*$classCode*" }
        
        if ($mmoDirectories.Count -gt 0) {
            # Use the first matching directory
            $mmoDirectory = $mmoDirectories[0].FullName
            Write-DebugLog "Found directory matching MMO and class code: $mmoDirectory" -Category "Configure"
            
            # Look for labor lookup file in that directory
            $lookupFiles = Get-ChildItem $mmoDirectory -Filter "*-Labor-Lookup.csv"
            if ($lookupFiles.Count -gt 0) {
                $lookupTablePath = $lookupFiles[0].FullName
                Write-DebugLog "Found lookup table: $lookupTablePath" -Category "Configure"
            }
        }
        # If not found, try just the MMO
        elseif ([string]::IsNullOrWhiteSpace($lookupTablePath)) {
            $mmoDirectories = Get-ChildItem $baseDir -Directory | 
                Where-Object { $_.Name -like "*$mmo*" }
            
            if ($mmoDirectories.Count -gt 0) {
                # Use the first matching directory
                $mmoDirectory = $mmoDirectories[0].FullName
                Write-DebugLog "Found directory matching just MMO: $mmoDirectory" -Category "Configure"
                
                # Look for labor lookup file in that directory
                $lookupFiles = Get-ChildItem $mmoDirectory -Filter "*-Labor-Lookup.csv"
                if ($lookupFiles.Count -gt 0) {
                    $lookupTablePath = $lookupFiles[0].FullName
                    Write-DebugLog "Found lookup table: $lookupTablePath" -Category "Configure"
                }
            }
        }
    }
    
    # If MMO isn't available, try to find by machine acronym
    if ([string]::IsNullOrWhiteSpace($lookupTablePath) -and -not [string]::IsNullOrWhiteSpace($acronym)) {
        $baseDir = "C:\Users\JR\Documents\Work\Programming Projects\LaborChecklist\Machine Labor Rubrics"
        
        # Try alternate path if first one doesn't exist
        if (-not (Test-Path $baseDir)) {
            $baseDir = "C:\Users\JR\Documents\Work\Programming Projects\LaborChecklist\Machine Labor Rubrics"
        }
        
        # Try to find directories with this machine acronym
        $machineDirectories = Get-ChildItem $baseDir -Directory | 
            Where-Object { $_.Name -like "*$acronym*" }
        
        if ($machineDirectories.Count -gt 0) {
            # Use the first matching directory
            $machineDirectory = $machineDirectories[0].FullName
            Write-DebugLog "Found directory matching machine acronym: $machineDirectory" -Category "Configure"
            
            # Look for labor lookup file in that directory
            $lookupFiles = Get-ChildItem $machineDirectory -Filter "*-Labor-Lookup.csv"
            if ($lookupFiles.Count -gt 0) {
                $lookupTablePath = $lookupFiles[0].FullName
                Write-DebugLog "Found lookup table: $lookupTablePath" -Category "Configure"
            }
        }
    }
    
    Write-DebugLog "Using lookup table path: $lookupTablePath" -Category "Configure"
    
    # Check if we have original values in the ListView item's Tag property
    $origDaysPerWeek = $null
    $origToursPerDay = $null
    
    if ($selectedItem.Tag -is [hashtable]) {
        if ($selectedItem.Tag.ContainsKey("OriginalDaysPerWeek")) {
            $origDaysPerWeek = $selectedItem.Tag.OriginalDaysPerWeek
        }
        if ($selectedItem.Tag.ContainsKey("OriginalToursPerDay")) {
            $origToursPerDay = $selectedItem.Tag.OriginalToursPerDay
        }
    }
    # If no original values stored, but initial values exist, store them now
    elseif ($initialValues.ContainsKey("Operation (days/wk)") -or $initialValues.ContainsKey("Tours/Day")) {
        $selectedItem.Tag = @{}
        
        if ($initialValues.ContainsKey("Operation (days/wk)")) {
            $selectedItem.Tag.OriginalDaysPerWeek = $initialValues["Operation (days/wk)"]
            $origDaysPerWeek = $initialValues["Operation (days/wk)"]
        }
        
        if ($initialValues.ContainsKey("Tours/Day")) {
            $selectedItem.Tag.OriginalToursPerDay = $initialValues["Tours/Day"]
            $origToursPerDay = $initialValues["Tours/Day"]
        }
    }
    
    # Show machine configuration dialog with relevant parameters
    $machineConfig = Show-MachineConfigDialog -ExistingMachine $true -InitialValues $initialValues -MMO $mmo -MachineAcronym $acronym -ClassCode $classCode -LookupTablePath $lookupTablePath
    
    if ($machineConfig) {
        # Update ListView with new configuration
        # First, update class code and MMO if changed
        $newClassCode = $machineConfig["ClassCode"]
        $newMMO = $machineConfig["MMO"]
        
        Write-DebugLog "Configuration returned - Class Code: $newClassCode, MMO: $newMMO" -Category "Configure"
        
        if (-not [string]::IsNullOrWhiteSpace($newClassCode) -and $newClassCode -ne $classCode) {
            $selectedItem.SubItems[2].Text = $newClassCode
            Write-DebugLog "Updated class code from '$classCode' to '$newClassCode'" -Category "Configure"
        }
        
        if (-not [string]::IsNullOrWhiteSpace($newMMO) -and $newMMO -ne $mmo) {
            $selectedItem.SubItems[3].Text = $newMMO
            Write-DebugLog "Updated MMO from '$mmo' to '$newMMO'" -Category "Configure"
        }
        
        # Update all parameters using reverse column mapping
        $reverseColumnMap = @{
            "Operation (days/wk)" = 4
            "Tours/Day" = 5
            "Stackers" = 6
            "Inductions" = 7
            "Transports" = 8
            "LIM Modules" = 9
            "Machine Type" = 10
            "Site" = 11
            "PSM #" = 12
            "Terminal Type" = 13
            "Equipment Code" = 14
            "Machines" = 15
        }
        
        # Now update the appropriate columns for each configured parameter
        foreach ($paramName in $machineConfig.Keys) {
            # Skip ClassCode, MMO, OriginalDaysPerWeek, OriginalToursPerDay, and ValuesAdjusted as they're handled separately
            if ($paramName -eq "ClassCode" -or $paramName -eq "MMO" -or 
                $paramName -eq "OriginalDaysPerWeek" -or $paramName -eq "OriginalToursPerDay" -or
                $paramName -eq "ValuesAdjusted") {
                continue
            }
            
            # Find the column index for this parameter
            if ($reverseColumnMap.ContainsKey($paramName)) {
                $columnIndex = $reverseColumnMap[$paramName]
                
                # Ensure there are enough SubItems in the ListView item
                while ($selectedItem.SubItems.Count -le $columnIndex) {
                    $selectedItem.SubItems.Add("")
                }
                
                # Update the value
                $selectedItem.SubItems[$columnIndex].Text = $machineConfig[$paramName]
                Write-DebugLog "Updated $paramName to: $($machineConfig[$paramName])" -Category "Configure"
                
                # If this is a value that has been adjusted, set the color
                if (($paramName -eq "Operation (days/wk)" -and $machineConfig.ContainsKey("OriginalDaysPerWeek") -and 
                     $machineConfig[$paramName] -ne $machineConfig["OriginalDaysPerWeek"]) -or
                    ($paramName -eq "Tours/Day" -and $machineConfig.ContainsKey("OriginalToursPerDay") -and
                     $machineConfig[$paramName] -ne $machineConfig["OriginalToursPerDay"])) {
                    
                    $selectedItem.SubItems[$columnIndex].ForeColor = [System.Drawing.Color]::DarkOrange
                }
            }
        }
        
        # Store original values and flags in the Tag if values were adjusted
        if ($machineConfig.ContainsKey("ValuesAdjusted") -and $machineConfig["ValuesAdjusted"]) {
            # Initialize Tag as hashtable if it's not already
            if ($selectedItem.Tag -isnot [hashtable]) {
                $selectedItem.Tag = @{}
            }
            
            # Store or update original values
            if ($machineConfig.ContainsKey("OriginalDaysPerWeek")) {
                $selectedItem.Tag.OriginalDaysPerWeek = $machineConfig["OriginalDaysPerWeek"]
            }
            
            if ($machineConfig.ContainsKey("OriginalToursPerDay")) {
                $selectedItem.Tag.OriginalToursPerDay = $machineConfig["OriginalToursPerDay"]
            }
            
            # Set adjusted flag
            $selectedItem.Tag.Adjusted = $true
            
            # Store the standardized values for toggle feature
            $selectedItem.Tag.StandardDaysPerWeek = $machineConfig["Operation (days/wk)"]
            $selectedItem.Tag.StandardToursPerDay = $machineConfig["Tours/Day"]
            
            # Set showing standard flag
            $selectedItem.Tag.ShowingStandard = $true
        }
        
        # Display success message
        [System.Windows.Forms.MessageBox]::Show(
            "Machine configuration updated successfully.",
            "Success",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
})

# View Details Button Logic
$btnViewDetails.Add_Click({
    if ($listView.SelectedItems.Count -eq 0) { 
        Write-DebugLog "No item selected in ListView" -Category "ViewDetails"
        return 
    }
    
    $selectedItem = $listView.SelectedItems[0]
    Write-DebugLog "Selected item index: $($listView.Items.IndexOf($selectedItem))" -Category "ViewDetails"
    Write-DebugLog "Selected item subItems count: $($selectedItem.SubItems.Count)" -Category "ViewDetails"
    
    # Debug each column of the selected item
    for ($i = 0; $i -lt $selectedItem.SubItems.Count; $i++) {
        $columnName = if ($i -lt $listView.Columns.Count) { $listView.Columns[$i].Text } else { "Column $i" }
        $value = $selectedItem.SubItems[$i].Text
        Write-DebugLog "ListView column '$columnName': '$value'" -Category "ViewDetails"
    }

    $mmo = $selectedItem.SubItems[3].Text
    $acronym = $selectedItem.SubItems[0].Text
    $classCode = $selectedItem.SubItems[2].Text
    $machineNumber = $selectedItem.SubItems[1].Text
    # Create Machine ID from acronym and number
    $machineID = "$acronym $machineNumber"
    
    Write-DebugLog "Parsed values - MMO: '$mmo', Acronym: '$acronym', Class Code: '$classCode', Machine Number: '$machineNumber', Machine ID: '$machineID'" -Category "ViewDetails"
    
    # Base directory
    $baseDir = "C:\Users\JR\Documents\Work\Programming Projects\LaborChecklist\Machine Labor Rubrics"
    
    # Update paths if needed (check if they exist)
    if (-not (Test-Path $baseDir)) {
        $baseDir = "C:\Users\JR\Documents\Work\Programming Projects\LaborChecklist\Machine Labor Rubrics"
    }
    
    Write-DebugLog "Base directory: $baseDir" -Category "ViewDetails"
 
    # Find MMO directory
    $mmoDirectory = Get-ChildItem $baseDir -Directory |
        Where-Object { $_.Name -like "$mmo*" } |
        Select-Object -First 1 -ExpandProperty FullName
 
    if (-not $mmoDirectory) {
        Write-DebugLog "MMO directory not found for MMO: '$mmo'" -Category "ViewDetails"
        [System.Windows.Forms.MessageBox]::Show("MMO directory not found", "Error")
        return
    }
    
    Write-DebugLog "Found MMO directory: $mmoDirectory" -Category "ViewDetails"
    
    # Directly analyze the lookup table if available
    $mmoHeaders = @()
    $needsConfigDialog = $false

    try {
        # Find lookup table
        $lookupTablePath = $null
        $lookupFiles = Get-ChildItem $mmoDirectory -Filter "*-Labor-Lookup.csv"
        
        if ($lookupFiles.Count -gt 0) {
            $lookupTablePath = $lookupFiles[0].FullName
            Write-DebugLog "Found lookup table: $lookupTablePath" -Category "ViewDetails"
            
            # Load the lookup table and extract headers directly
            $lookupData = Import-Csv $lookupTablePath
            if ($lookupData.Count -gt 0) {
                $mmoHeaders = $lookupData[0].PSObject.Properties.Name
                Write-DebugLog "Extracted headers directly from lookup table: $($mmoHeaders -join ', ')" -Category "ViewDetails"
                
                # Check if this MMO needs more than just basic fields
                $basicColumns = @("Operation (days/wk)", "Tours/Day", "Total (hrs/yr)")
                $specialColumns = $mmoHeaders | Where-Object { $_ -notin $basicColumns }
                $needsConfigDialog = $specialColumns.Count -gt 0
                
                Write-DebugLog "Special columns: $($specialColumns -join ', ')" -Category "ViewDetails"
                Write-DebugLog "Needs config dialog: $needsConfigDialog" -Category "ViewDetails"
            }
        } else {
            Write-DebugLog "No lookup table found in directory: $mmoDirectory" -Category "ViewDetails"
        }
    }
    catch {
        Write-DebugLog "Error analyzing lookup table headers: $_" -Category "ViewDetails"
    }
    
    # Define our parameter mapping 
    $columnMap = @{
        4 = "Operation (days/wk)"
        5 = "Tours/Day"
        6 = "Stackers"
        7 = "Inductions"
        8 = "Transports"
        9 = "LIM Modules"
        10 = "Machine Type"
        11 = "Site"
        12 = "PSM #"
        13 = "Terminal Type"
        14 = "Equipment Code"
        15 = "Machines"
    }
    
    # Check if machine has any configuration - looking at all possible configuration fields
    $hasConfig = $false
    
    # Check each potential configuration column
    foreach ($index in $columnMap.Keys) {
        # If we have this column and it's not empty, we have configuration
        if (($selectedItem.SubItems.Count -gt $index) -and 
            (-not [string]::IsNullOrWhiteSpace($selectedItem.SubItems[$index].Text))) {
            $hasConfig = $true
            break
        }
    }
    
    Write-DebugLog "Machine has configuration: $hasConfig" -Category "ViewDetails"
    
    # For special columns that need configuration, show dialog if not configured
    if ((-not $hasConfig) -and $needsConfigDialog) {
        Write-DebugLog "Machine needs configuration and doesn't have it" -Category "ViewDetails"
        $configResult = [System.Windows.Forms.MessageBox]::Show(
            "This machine needs configuration for special parameters. Would you like to configure them now?",
            "Configure Machine",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )

        if ($configResult -eq [System.Windows.Forms.DialogResult]::Yes) {
            # Find lookup table to get actual values
            $lookupTablePath = Get-ChildItem $mmoDirectory -Filter "*-Labor-Lookup.csv" | 
                Select-Object -First 1 -ExpandProperty FullName
                
            if (-not $lookupTablePath) {
                Write-DebugLog "Labor lookup file not found in directory: $mmoDirectory" -Category "ViewDetails"
                [System.Windows.Forms.MessageBox]::Show("Labor lookup file not found", "Error")
                return
            }
            
            Write-DebugLog "Found lookup table: $lookupTablePath" -Category "ViewDetails"
            
            # Create initial values from existing configuration
            $initialValues = @{}
            foreach ($index in $columnMap.Keys) {
                if ($selectedItem.SubItems.Count -gt $index -and -not [string]::IsNullOrWhiteSpace($selectedItem.SubItems[$index].Text)) {
                    $initialValues[$columnMap[$index]] = $selectedItem.SubItems[$index].Text
                }
            }
            
            # Pass the headers extracted directly from the lookup table
            $machineConfig = Show-MachineConfigDialog -ExistingMachine ($initialValues.Count -gt 0) -MMO $mmo -MachineAcronym $acronym -ClassCode $classCode -MMOHeaders $mmoHeaders -LookupTablePath $lookupTablePath -InitialValues $initialValues

            if ($machineConfig) {
                Write-DebugLog "Machine configuration returned from dialog" -Category "ViewDetails"
                
                # Update the ListView item with the new configuration
                # First update Class Code and MMO if changed
                if ($machineConfig.ContainsKey("ClassCode") -and -not [string]::IsNullOrWhiteSpace($machineConfig["ClassCode"])) {
                    $selectedItem.SubItems[2].Text = $machineConfig["ClassCode"]
                }
                
                if ($machineConfig.ContainsKey("MMO") -and -not [string]::IsNullOrWhiteSpace($machineConfig["MMO"])) {
                    $selectedItem.SubItems[3].Text = $machineConfig["MMO"]
                }
                
                # Now update all other parameters using the column map
                foreach ($paramName in $machineConfig.Keys) {
                    # Skip ClassCode and MMO as they're handled separately
                    if ($paramName -eq "ClassCode" -or $paramName -eq "MMO") { continue }
                    
                    # Find the corresponding column index
                    $columnIndex = $columnMap.GetEnumerator() | Where-Object { $_.Value -eq $paramName } | Select-Object -ExpandProperty Key
                    
                    if ($columnIndex) {
                        # Ensure we have enough columns
                        while ($selectedItem.SubItems.Count -le $columnIndex) {
                            $selectedItem.SubItems.Add("")
                        }
                        
                        # Update the value
                        $selectedItem.SubItems[$columnIndex].Text = $machineConfig[$paramName]
                        Write-DebugLog "Updated $paramName to: $($machineConfig[$paramName])" -Category "ViewDetails"
                    }
                }
            }
            else {
                # User cancelled configuration, abort view details
                Write-DebugLog "User cancelled machine configuration" -Category "ViewDetails"
                return
            }
        }
        else {
            # User declined to configure, abort view details
            Write-DebugLog "User declined to configure machine" -Category "ViewDetails"
            return
        }
    }

    # Create details form
    $detailForm = New-Object System.Windows.Forms.Form
    $detailForm.Text = "MMO Details - $mmo"
    $detailForm.Size = New-Object System.Drawing.Size(800, 600)
    $detailForm.StartPosition = "CenterScreen"
    $detailForm.MinimumSize = New-Object System.Drawing.Size(600, 400)
 
    # Create tab control
    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Dock = "Fill"
    $detailForm.Controls.Add($tabControl)
    

    
    # Create tab selected event handler to ensure data loads when switching tabs
    $tabControl.Add_Selected({
        $selectedTab = $tabControl.SelectedTab
        
        if ($selectedTab -eq $staffingTab) {
            Write-DebugLog "Staffing tab selected, loading data for Machine ID: '$machineID'" -Category "StaffingTab"
            
            # Clear any existing data
            $staffingTableGrid.DataSource = $null
            $staffingTableGrid.Refresh()
            
            # Extract basic information we need
            $machineAcronym = if ($machineID -match "^([A-Za-z]+)") { $matches[1] } else { "" }
            
            Write-DebugLog "Using MMO: '$mmo', Class Code: '$classCode', Machine Acronym: '$machineAcronym'" -Category "StaffingTab"
            
            # Find the appropriate staffing table file
            $baseDir = "C:\Users\JR\Documents\Work\Programming Projects\LaborChecklist\Machine Labor Rubrics"
            if (-not (Test-Path $baseDir)) {
                $baseDir = "C:\Users\JR\Documents\Work\Programming Projects\LaborChecklist\Machine Labor Rubrics"
            }
            
            # Look for the most specific directory first
            $staffingFilePath = $null
            $searchPatterns = @(
                # Most specific pattern
                "*$mmo*$machineAcronym*-$classCode*", 
                # Less specific patterns
                "*$mmo*-$classCode*",
                "*$mmo*$machineAcronym*",
                "*$mmo*"
            )
            
            foreach ($pattern in $searchPatterns) {
                if ($staffingFilePath) { break }
                
                Write-DebugLog "Searching for directories with pattern: '$pattern'" -Category "StaffingTab"
                $matchingDirs = Get-ChildItem $baseDir -Directory | Where-Object { $_.Name -like $pattern }
                
                if ($matchingDirs -and $matchingDirs.Count -gt 0) {
                    # Search in each matching directory
                    foreach ($dir in $matchingDirs) {
                        Write-DebugLog "Checking directory: $($dir.FullName)" -Category "StaffingTab"
                        
                        # Define possible file patterns in order of specificity
                        $filePatterns = @(
                            "$mmo-$classCode-Staffing-Table.csv",
                            "$machineAcronym-$classCode-Staffing-Table.csv",
                            "$mmo-Staffing-Table.csv",
                            "*Staffing-Table.csv"
                        )
                        
                        foreach ($filePattern in $filePatterns) {
                            $matchingFiles = Get-ChildItem -Path $dir.FullName -Filter $filePattern -ErrorAction SilentlyContinue
                            
                            if ($matchingFiles -and $matchingFiles.Count -gt 0) {
                                $staffingFilePath = $matchingFiles[0].FullName
                                Write-DebugLog "Found staffing file: $staffingFilePath" -Category "StaffingTab"
                                break
                            }
                        }
                        
                        if ($staffingFilePath) { break }
                    }
                }
            }
            
            # Load data from the file if found
            if ($staffingFilePath -and (Test-Path $staffingFilePath)) {
                Write-DebugLog "Loading data from: $staffingFilePath" -Category "StaffingTab"
                
                try {
                    # ---- USE A SIMPLER, MORE DIRECT APPROACH ----
                    # Create new DataTable
                    $staffingTable = New-Object System.Data.DataTable
                    
                    # Read the raw CSV content
                    $csvContent = Get-Content -Path $staffingFilePath -Raw -ErrorAction Stop
                    
                    # Simple cleanup - normalize line endings, trim whitespace
                    $csvContent = $csvContent.Replace("`r`n", "`n").Replace("`r", "`n").Trim()
                    $lines = $csvContent -split "`n" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
                    
                    if ($lines.Count -lt 1) {
                        Write-DebugLog "File contains no data" -Category "StaffingTab"
                        throw "Staffing file is empty"
                    }
                    
                    # --- PART 1: PARSE HEADER ROW ---
                    $headerLine = $lines[0]
                    Write-DebugLog "Header line: $headerLine" -Category "StaffingTab"
                    
                    # Parse the header - looking at your debug output, headers have quotes
                    # Extract column names from "header1","header2",...
                    $headerColumns = @()
                    $headerLine = $headerLine.Trim()
                    
                    # Special handling for quoted CSV format
                    if ($headerLine.StartsWith('"') -and $headerLine.Contains('","')) {
                        # Split by "," pattern and remove surrounding quotes
                        $headerColumns = $headerLine.Split('","') | ForEach-Object {
                            $_.Replace('"', '').Trim()
                        }
                        # Fix first and last items which might still have quotes
                        if ($headerColumns.Count -gt 0) {
                            $headerColumns[0] = $headerColumns[0].TrimStart('"')
                            $headerColumns[$headerColumns.Count-1] = $headerColumns[$headerColumns.Count-1].TrimEnd('"')
                        }
                    } else {
                        # Fallback to simple comma split
                        $headerColumns = $headerLine.Split(',') | ForEach-Object { $_.Trim('"').Trim() }
                    }
                    
                    Write-DebugLog "Parsed headers: $($headerColumns -join '|')" -Category "StaffingTab"
                    
                    # Add columns to DataTable
                    foreach ($colName in $headerColumns) {
                        if (-not [string]::IsNullOrWhiteSpace($colName)) {
                            $staffingTable.Columns.Add($colName, [string]) | Out-Null
                        }
                    }
                    
                    # Find Machine ID column index
                    $machineIdColIndex = -1
                    for ($i = 0; $i -lt $headerColumns.Count; $i++) {
                        if ($headerColumns[$i] -eq "Machine ID") {
                            $machineIdColIndex = $i
                            break
                        }
                    }
                    
                    # --- PART 2: PARSE DATA ROWS ---
                    $addedMachineRow = $false

                    # Process the data rows
                    for ($i = 1; $i -lt $lines.Count; $i++) {
                        $dataLine = $lines[$i].Trim()
                        if ([string]::IsNullOrWhiteSpace($dataLine)) { continue }
                        
                        Write-DebugLog "Processing data line ${i}: $dataLine" -Category "StaffingTab"
                        
                        # Extract field values - handle both quoted and unquoted formats
                        $fieldValues = @()
                        
                        if ($dataLine.StartsWith('"') -and $dataLine.Contains('","')) {
                            # Split by "," pattern and remove surrounding quotes
                            $fieldValues = $dataLine.Split('","') | ForEach-Object {
                                $_.Replace('"', '').Trim()
                            }
                            # Fix first and last items which might still have quotes
                            if ($fieldValues.Count -gt 0) {
                                $fieldValues[0] = $fieldValues[0].TrimStart('"')
                                $fieldValues[$fieldValues.Count-1] = $fieldValues[$fieldValues.Count-1].TrimEnd('"')
                            }
                        } else {
                            # Fallback to simple comma split
                            $fieldValues = $dataLine.Split(',') | ForEach-Object { $_.Trim('"').Trim() }
                        }
                        
                        # Check if field values count matches header count
                        if ($fieldValues.Count -ne $headerColumns.Count) {
                            Write-DebugLog "Row $i has $($fieldValues.Count) fields, expected $($headerColumns.Count) - attempting to fix" -Category "StaffingTab"
                            
                            # If missing trailing empty fields, add them
                            while ($fieldValues.Count -lt $headerColumns.Count) {
                                $fieldValues += ""
                            }
                            
                            # If too many fields, truncate (shouldn't happen with proper CSV)
                            if ($fieldValues.Count -gt $headerColumns.Count) {
                                $fieldValues = $fieldValues[0..($headerColumns.Count-1)]
                            }
                        }
                        
                        # Check if this row is for our machine
                        $rowMachineId = if ($machineIdColIndex -ge 0 -and $machineIdColIndex -lt $fieldValues.Count) { 
                            $fieldValues[$machineIdColIndex] 
                        } else { 
                            "" 
                        }
                        
                        # Only include the row if it matches our exact machine ID
                        $includeRow = $false
                        if (-not [string]::IsNullOrWhiteSpace($machineID) -and -not [string]::IsNullOrWhiteSpace($rowMachineId)) {
                            # Compare with trimming and case insensitivity
                            $includeRow = $rowMachineId.Trim() -eq $machineID.Trim()
                            if ($includeRow) {
                                $addedMachineRow = $true
                                Write-DebugLog "Found exact matching machine ID: $rowMachineId" -Category "StaffingTab"
                            }
                        }
                        
                        if ($includeRow) {
                            $newRow = $staffingTable.NewRow()
                            for ($j = 0; $j -lt [Math]::Min($headerColumns.Count, $fieldValues.Count); $j++) {
                                $newRow[$headerColumns[$j]] = $fieldValues[$j]
                            }
                            $staffingTable.Rows.Add($newRow)
                            Write-DebugLog "Added row for machine: $rowMachineId" -Category "StaffingTab"
                        }
                    }

                    # --- PART 3: BIND DATA TO GRID ---
                    if ($staffingTable.Rows.Count -gt 0) {
                        Write-DebugLog "Binding grid with $($staffingTable.Rows.Count) rows, $($staffingTable.Columns.Count) columns" -Category "StaffingTab"
                        
                        # Force clear any existing binding
                        $staffingTableGrid.DataSource = $null
                        [System.Windows.Forms.Application]::DoEvents()
                        
                        # Bind the grid
                        $staffingTableGrid.DataSource = $staffingTable
                        $staffingTableGrid.Refresh()
                        $staffingTableGrid.AutoResizeColumns()
                        
                        # Force refresh UI
                        [System.Windows.Forms.Application]::DoEvents()
                        
                    }
                    else {
                        # No data for this machine - show empty grid
                        Write-DebugLog "No matching machine ID found in staffing data" -Category "StaffingTab"
                        
                        # Still bind the empty table to maintain column structure
                        $staffingTableGrid.DataSource = $staffingTable
                        $staffingTableGrid.Refresh()
                        
                        [System.Windows.Forms.MessageBox]::Show(
                            "No staffing data found for machine ID '$machineID'. Click 'Edit Staffing Data' to add data.",
                            "No Data",
                            [System.Windows.Forms.MessageBoxButtons]::OK,
                            [System.Windows.Forms.MessageBoxIcon]::Information
                        )
                    }
                }
                catch {
                    Write-DebugLog "Error processing staffing data: $_" -Category "StaffingTab"
                    Write-DebugLog "Stack trace: $($_.ScriptStackTrace)" -Category "StaffingTab"
                    
                    [System.Windows.Forms.MessageBox]::Show(
                        "Error loading staffing data: $_`n`nPlease use the 'Edit Staffing Data' button to correct any issues.",
                        "Error",
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Error
                    )
                }
            }
        }
        if ($selectedTab -eq $calculatedStaffingTab) {
            Write-DebugLog "Calculated Staffing tab selected" -Category "CalculatedStaffing"
            
            # Clear any existing data
            $calculatedStaffingGrid.DataSource = $null
            $calculatedStaffingGrid.Refresh()
            
            # Extract basic information we need
            $machineAcronym = if ($machineID -match "^([A-Za-z]+)") { $matches[1] } else { "" }
            
            Write-DebugLog "Loading calculated staffing table for MMO: $mmo, Machine ID: $machineID" -Category "CalculatedStaffing"
            
            # Find the appropriate calculated staffing table file
            $baseDir = "C:\Users\JR\Documents\Work\Programming Projects\LaborChecklist\Machine Labor Rubrics"
            if (-not (Test-Path $baseDir)) {
                $baseDir = "C:\Users\JR\Documents\Work\Programming Projects\LaborChecklist\Machine Labor Rubrics"
            }
            
            # Look for the most specific directory first
            $calculatedFilePath = $null
            $searchPatterns = @(
                # Most specific pattern
                "*$mmo*$machineAcronym*-$classCode*", 
                # Less specific patterns
                "*$mmo*-$classCode*",
                "*$mmo*$machineAcronym*",
                "*$mmo*"
            )
            
            foreach ($pattern in $searchPatterns) {
                if ($calculatedFilePath) { break }
                
                Write-DebugLog "Searching for directories with pattern: '$pattern'" -Category "CalculatedStaffing"
                $matchingDirs = Get-ChildItem $baseDir -Directory | Where-Object { $_.Name -like $pattern }
                
                if ($matchingDirs -and $matchingDirs.Count -gt 0) {
                    # Search in each matching directory
                    foreach ($dir in $matchingDirs) {
                        Write-DebugLog "Checking directory: $($dir.FullName)" -Category "CalculatedStaffing"
                        
                        # Define possible file patterns in order of specificity
                        $filePatterns = @(
                            "$mmo-$classCode-Calculated-Staffing-Table.csv",
                            "$machineAcronym-$classCode-Calculated-Staffing-Table.csv",
                            "$mmo-Calculated-Staffing-Table.csv",
                            "*Calculated-Staffing-Table.csv"
                        )
                        
                        foreach ($filePattern in $filePatterns) {
                            $matchingFiles = Get-ChildItem -Path $dir.FullName -Filter $filePattern -ErrorAction SilentlyContinue
                            
                            if ($matchingFiles -and $matchingFiles.Count -gt 0) {
                                $calculatedFilePath = $matchingFiles[0].FullName
                                Write-DebugLog "Found calculated staffing file: $calculatedFilePath" -Category "CalculatedStaffing"
                                break
                            }
                        }
                        
                        if ($calculatedFilePath) { break }
                    }
                }
            }
            
            # Load data from the file if found
            if ($calculatedFilePath -and (Test-Path $calculatedFilePath)) {
                Write-DebugLog "Reading calculated staffing data from: $calculatedFilePath" -Category "CalculatedStaffing"
                
                try {
                    # Create new DataTable
                    $calculatedTable = New-Object System.Data.DataTable
                    
                    # Read the raw CSV content
                    $csvContent = Get-Content -Path $calculatedFilePath -Raw -ErrorAction Stop
                    
                    # Simple cleanup - normalize line endings, trim whitespace
                    $csvContent = $csvContent.Replace("`r`n", "`n").Replace("`r", "`n").Trim()
                    $lines = $csvContent -split "`n" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
                    
                    if ($lines.Count -lt 1) {
                        Write-DebugLog "File contains no data" -Category "CalculatedStaffing"
                        throw "Calculated staffing file is empty"
                    }
                    
                    # Parse header row
                    $headerLine = $lines[0]
                    Write-DebugLog "Header line: $headerLine" -Category "CalculatedStaffing"
                    
                    # Parse header - simply split by comma and trim
                    $headerColumns = $headerLine.Split(',') | ForEach-Object { $_.Trim() }
                    
                    Write-DebugLog "Parsed headers: $($headerColumns -join '|')" -Category "CalculatedStaffing"
                    
                    # Add columns to DataTable
                    foreach ($colName in $headerColumns) {
                        if (-not [string]::IsNullOrWhiteSpace($colName)) {
                            $calculatedTable.Columns.Add($colName, [string]) | Out-Null
                        }
                    }
                    
                    # Find Machine ID column index
                    $machineIdColIndex = -1
                    for ($i = 0; $i -lt $headerColumns.Count; $i++) {
                        if ($headerColumns[$i] -eq "Machine ID") {
                            $machineIdColIndex = $i
                            break
                        }
                    }
                    
                    # Process data rows (skip header)
                    $addedMachineRow = $false
                    
                    for ($i = 1; $i -lt $lines.Count; $i++) {
                        $dataLine = $lines[$i].Trim()
                        if ([string]::IsNullOrWhiteSpace($dataLine)) { continue }
                        
                        Write-DebugLog "Processing data line ${i}: $dataLine" -Category "CalculatedStaffing"
                        
                        # Simply split by commas
                        $fieldValues = $dataLine.Split(',') | ForEach-Object { $_.Trim() }
                        
                        # Check if field values count matches header count
                        if ($fieldValues.Count -ne $headerColumns.Count) {
                            Write-DebugLog "Row $i has $($fieldValues.Count) fields, expected $($headerColumns.Count) - attempting to fix" -Category "CalculatedStaffing"
                            
                            # If missing trailing empty fields, add them
                            while ($fieldValues.Count -lt $headerColumns.Count) {
                                $fieldValues += ""
                            }
                            
                            # If too many fields, truncate
                            if ($fieldValues.Count -gt $headerColumns.Count) {
                                $fieldValues = $fieldValues[0..($headerColumns.Count-1)]
                            }
                        }
                        
                        # Check if this row is for our machine
                        $rowMachineId = if ($machineIdColIndex -ge 0 -and $machineIdColIndex -lt $fieldValues.Count) { 
                            $fieldValues[$machineIdColIndex] 
                        } else { 
                            "" 
                        }
                        
                        # Only include the row if it matches our exact machine ID
                        $includeRow = $false
                        if (-not [string]::IsNullOrWhiteSpace($machineID) -and -not [string]::IsNullOrWhiteSpace($rowMachineId)) {
                            # Compare with trimming and case insensitivity
                            $includeRow = $rowMachineId.Trim() -eq $machineID.Trim()
                            if ($includeRow) {
                                $addedMachineRow = $true
                                Write-DebugLog "Found matching machine ID: $rowMachineId" -Category "CalculatedStaffing"
                            }
                        }
                        
                        if ($includeRow) {
                            $newRow = $calculatedTable.NewRow()
                            for ($j = 0; $j -lt [Math]::Min($headerColumns.Count, $fieldValues.Count); $j++) {
                                $newRow[$headerColumns[$j]] = $fieldValues[$j]
                            }
                            $calculatedTable.Rows.Add($newRow)
                            Write-DebugLog "Added row for machine: $rowMachineId" -Category "CalculatedStaffing"
                        }
                    }
                    
                    # Bind data to grid
                    if ($calculatedTable.Rows.Count -gt 0) {
                        Write-DebugLog "Successfully loaded calculated staffing table with $($calculatedTable.Rows.Count) rows" -Category "CalculatedStaffing"
                        
                        # Force clear any existing binding
                        $calculatedStaffingGrid.DataSource = $null
                        [System.Windows.Forms.Application]::DoEvents()
                        
                        # Bind the grid
                        $calculatedStaffingGrid.DataSource = $calculatedTable
                        $calculatedStaffingGrid.Refresh()
                        $calculatedStaffingGrid.AutoResizeColumns()
                        
                        # Force refresh UI
                        [System.Windows.Forms.Application]::DoEvents()
                    }
                    else {
                        # No data for this machine - show empty grid
                        Write-DebugLog "No matching rows found in calculated staffing data" -Category "CalculatedStaffing"
                        
                        # Create empty table with same columns
                        $emptyTable = New-Object System.Data.DataTable
                        
                        foreach ($colName in $headerColumns) {
                            if (-not [string]::IsNullOrWhiteSpace($colName)) {
                                $emptyTable.Columns.Add($colName, [string]) | Out-Null
                            }
                        }
                        
                        # Bind the empty table
                        $calculatedStaffingGrid.DataSource = $emptyTable
                        $calculatedStaffingGrid.Refresh()
                        
                        Write-DebugLog "No calculated staffing data found - loaded empty table" -Category "CalculatedStaffing"
                    }
                }
                catch {
                    Write-DebugLog "Error loading calculated staffing data: $_" -Category "CalculatedStaffing"
                    Write-DebugLog "Stack trace: $($_.ScriptStackTrace)" -Category "CalculatedStaffing"
                    
                    [System.Windows.Forms.MessageBox]::Show(
                        "Error loading calculated staffing data: $_`n`nPlease use the 'Calculate Staffing' button to regenerate the data.",
                        "Error",
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Error
                    )
                }
            }
            else {
                Write-DebugLog "No calculated staffing file found" -Category "CalculatedStaffing"
                
                # Create empty table with basic columns
                $emptyTable = New-Object System.Data.DataTable
                $emptyTable.Columns.Add("Machine ID", [string]) | Out-Null
                $emptyTable.Columns.Add("Operation (days/wk)", [string]) | Out-Null
                $emptyTable.Columns.Add("Tours/Day", [string]) | Out-Null
                $emptyTable.Columns.Add("MM7", [string]) | Out-Null
                $emptyTable.Columns.Add("MPE9", [string]) | Out-Null
                $emptyTable.Columns.Add("ET10", [string]) | Out-Null
                $emptyTable.Columns.Add("Total (hrs/yr)", [string]) | Out-Null
                $emptyTable.Columns.Add("Operational Maintenance (hrs/yr)", [string]) | Out-Null
                
                $calculatedStaffingGrid.DataSource = $emptyTable
                $calculatedStaffingGrid.Refresh()
                
                Write-DebugLog "No calculated staffing data found - loaded empty table" -Category "CalculatedStaffing"
            }
        }

    })

    
    $btnDebugStaffing = New-Object System.Windows.Forms.Button -Property @{
    Text = "Debug Staffing"
    Size = New-Object System.Drawing.Size(120, 30)
    Location = New-Object System.Drawing.Point(140, 10)
    Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
    
    }

    # Create PDF button
    $btnOpenPDF = New-Object System.Windows.Forms.Button -Property @{
        Text = "Open PDF"
        Size = New-Object System.Drawing.Size(120, 30)
        Location = New-Object System.Drawing.Point(270, 10)
    }

    # Find PDF
    $pdfFiles = Get-ChildItem $mmoDirectory -Filter "*.pdf" | 
        Select-Object -First 1

    if ($pdfFiles) {
        $pdfPath = $pdfFiles.FullName
        $btnOpenPDF.Add_Click({ Start-Process $pdfPath })
    } else {
        $pdfPath = $null
        $btnOpenPDF.Enabled = $false
        $btnOpenPDF.Text = "PDF Not Found"
    }
 
    # Create data grids with names for easier reference
    $laborLookupGrid = New-Object System.Windows.Forms.DataGridView -Property @{
        Name = "laborLookupGrid"
        Dock = "Fill"
        AutoSizeColumnsMode = "Fill"
        ReadOnly = $true
        AllowUserToAddRows = $false
        AllowUserToDeleteRows = $false
    }
 
    $staffingTableGrid = New-Object System.Windows.Forms.DataGridView -Property @{
        Name = "staffingTableGrid"
        Dock = "Fill"
        AutoSizeColumnsMode = "Fill"
        ReadOnly = $true
        AllowUserToAddRows = $false
        AllowUserToDeleteRows = $false
    }

    $staffingTableGrid.AutoGenerateColumns = $true
    $staffingTableGrid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $staffingTableGrid.AllowUserToAddRows = $false  # Prevent adding rows in the grid
    $staffingTableGrid.AllowUserToDeleteRows = $false  # Prevent deleting rows in the grid
    $staffingTableGrid.EditMode = [System.Windows.Forms.DataGridViewEditMode]::EditOnKeystrokeOrF2  # Enable editing
    $staffingTableGrid.MultiSelect = $false  # Prevent multiple selection

    $laborTab = New-Object System.Windows.Forms.TabPage -Property @{
        Text = "Labor Lookup"
    }
    $laborTab.Controls.Add($laborLookupGrid)
 
    $staffingTab = New-Object System.Windows.Forms.TabPage -Property @{
        Text = "Staffing Table"
    }

    # Add tabs
    $calculatedStaffingTab = New-Object System.Windows.Forms.TabPage -Property @{
        Text = "Calculated Staffing Table"
    }

    # Create DataGridView for calculated staffing data
    $calculatedStaffingGrid = New-Object System.Windows.Forms.DataGridView -Property @{
        Name = "calculatedStaffingGrid"
        Dock = "Fill"
        AutoSizeColumnsMode = "Fill"
        ReadOnly = $true
        AllowUserToAddRows = $false
        AllowUserToDeleteRows = $false
        AutoGenerateColumns = $true
        SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
        MultiSelect = $false       
    }

    # Debug grid properties
    Write-DebugLog "Grid properties before calculation:" -Category "CalculatedStaffing"
    Write-DebugLog "- calculatedStaffingGrid.AutoGenerateColumns = $($calculatedStaffingGrid.AutoGenerateColumns)" -Category "CalculatedStaffing"
    Write-DebugLog "- calculatedStaffingGrid.DataSource = $($calculatedStaffingGrid.DataSource -ne $null)" -Category "CalculatedStaffing"
    Write-DebugLog "- calculatedStaffingGrid.Columns.Count = $($calculatedStaffingGrid.Columns.Count)" -Category "CalculatedStaffing"
    Write-DebugLog "- calculatedStaffingGrid.Visible = $($calculatedStaffingGrid.Visible)" -Category "CalculatedStaffing"
    Write-DebugLog "- calculatedStaffingGrid.Enabled = $($calculatedStaffingGrid.Enabled)" -Category "CalculatedStaffing"
    Write-DebugLog "- calculatedStaffingGrid.Dock = $($calculatedStaffingGrid.Dock)" -Category "CalculatedStaffing"
    Write-DebugLog "- calculatedStaffingGrid.Size = $($calculatedStaffingGrid.Size.Width)x$($calculatedStaffingGrid.Size.Height)" -Category "CalculatedStaffing"

    # Create panel for calculated staffing tab
    $calculatedStaffingPanel = New-Object System.Windows.Forms.Panel
    $calculatedStaffingPanel.Dock = "Fill"

    # Create Calculate button
    $btnCalculate = New-Object System.Windows.Forms.Button -Property @{
        Text = "Calculate Staffing"
        Size = New-Object System.Drawing.Size(120, 30)
        Location = New-Object System.Drawing.Point(10, 10)
        Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
    }

    # Create Export button for calculated data
    $btnExportCalculated = New-Object System.Windows.Forms.Button -Property @{
        Text = "Export Calculated"
        Size = New-Object System.Drawing.Size(120, 30)
        Location = New-Object System.Drawing.Point(140, 10)
        Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
    }

    # Position the DataGridView below buttons
    $calculatedStaffingGrid.Location = New-Object System.Drawing.Point(0, 50)
    $calculatedStaffingGrid.Width = 600
    $calculatedStaffingGrid.Height = 400
    $calculatedStaffingGrid.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor
                                    [System.Windows.Forms.AnchorStyles]::Bottom -bor
                                    [System.Windows.Forms.AnchorStyles]::Left -bor
                                    [System.Windows.Forms.AnchorStyles]::Right

    # Add controls to panel
    $calculatedStaffingPanel.Controls.Add($btnCalculate)
    $calculatedStaffingPanel.Controls.Add($btnExportCalculated)
    $calculatedStaffingPanel.Controls.Add($calculatedStaffingGrid)

    # Add panel to tab
    $calculatedStaffingTab.Controls.Add($calculatedStaffingPanel)

    # Handle resize events
    $calculatedStaffingPanel.Add_Resize({
        $gridWidth = $calculatedStaffingPanel.ClientSize.Width
        $gridHeight = $calculatedStaffingPanel.ClientSize.Height - 50
        $calculatedStaffingGrid.Size = New-Object System.Drawing.Size($gridWidth, $gridHeight)
    })

    # Function to calculate staffing based on labor lookup data
    function Calculate-StaffingRequirements {
        param (
            [hashtable]$MachineMetrics,
            [string]$MachineID,
            [string]$MMO,
            [string]$ClassCode
        )
        
        Write-DebugLog "Starting Calculate-StaffingRequirements for Machine ID: $MachineID" -Category "CalculatedStaffing"
        Write-DebugLog "Machine Metrics: $($MachineMetrics | ConvertTo-Json -Compress)" -Category "CalculatedStaffing"
        Write-DebugLog "MMO: $MMO, Class Code: $ClassCode" -Category "CalculatedStaffing"
        
        # Define standard column sets
        $outputColumns = @("Total (hrs/yr)", "Operational Maintenance (hrs/yr)")
        $managementColumns = @("MM7", "MPE9", "ET10")
        $metadataColumns = @("Machine ID")
        
        # Create DataTable for calculated values
        $calculatedTable = New-Object System.Data.DataTable
        
        # Add standard columns to match staffing table structure
        $calculatedTable.Columns.Add("Machine ID", [string]) | Out-Null
        
        # Get all column names from labor lookup grid
        $gridColumns = @()
        foreach ($col in $laborLookupGrid.Columns) {
            $gridColumns += $col.Name
        }
        
        # Determine input columns (everything except output columns)
        $inputColumns = $gridColumns | Where-Object { $_ -notin $outputColumns }
        Write-DebugLog "Input columns from labor lookup grid: $($inputColumns -join ', ')" -Category "CalculatedStaffing"
        
        # Add input columns to calculated table
        foreach ($column in $inputColumns) {
            if (-not $calculatedTable.Columns.Contains($column)) {
                $calculatedTable.Columns.Add($column, [string]) | Out-Null
            }
        }
        
        # Add management columns
        foreach ($column in $managementColumns) {
            if (-not $calculatedTable.Columns.Contains($column)) {
                $calculatedTable.Columns.Add($column, [string]) | Out-Null
            }
        }
        
        # Add output columns
        foreach ($column in $outputColumns) {
            if (-not $calculatedTable.Columns.Contains($column)) {
                $calculatedTable.Columns.Add($column, [string]) | Out-Null
            }
        }
        
        # Create row for this machine
        $newRow = $calculatedTable.NewRow()
        $newRow["Machine ID"] = $MachineID
        
        # --- FIND BEST MATCHING ROW USING HIGHLIGHT-MATCHEDCELLS LOGIC ---
        
        # Initialize variables for finding best match
        $bestRow = -1
        $bestScore = 0
        
        # Loop through all rows to find the best match
        for ($rowIndex = 0; $rowIndex -lt $laborLookupGrid.Rows.Count; $rowIndex++) {
            $row = $laborLookupGrid.Rows[$rowIndex]
            $rowScore = 0
            $matchCount = 0
            $totalParams = 0
            
            # Check each metric against row values
            foreach ($paramName in $MachineMetrics.Keys) {
                # Skip non-parameter metrics
                if ($paramName -eq "MachineType" -or $paramName -eq "MachineNumber") {
                    continue
                }
                
                # Skip empty values
                if ([string]::IsNullOrWhiteSpace($MachineMetrics[$paramName])) {
                    continue
                }
                
                $totalParams++
                
                # If grid has this column, check for a match
                if ($gridColumns -contains $paramName) {
                    $cellValue = $row.Cells[$paramName].Value
                    $metricValue = $MachineMetrics[$paramName]
                    
                    # Try to convert to same type for comparison
                    if ($cellValue -is [string] -and $metricValue -isnot [string]) {
                        $metricValue = $metricValue.ToString()
                    }
                    elseif ($cellValue -is [int] -and $metricValue -is [string]) {
                        try { $metricValue = [int]::Parse($metricValue) } catch { }
                    }
                    
                    if ($cellValue -eq $metricValue) {
                        $matchCount++
                        $rowScore += 10  # Add 10 points for exact match
                    }
                }
            }
            
            # Calculate final score as percentage of matches
            if ($totalParams -gt 0) {
                $finalScore = ($rowScore / ($totalParams * 10)) * 100
                Write-DebugLog "Row $rowIndex score: $finalScore% ($matchCount/$totalParams matches)" -Category "CalculatedStaffing"
                
                if ($finalScore -gt $bestScore) {
                    $bestScore = $finalScore
                    $bestRow = $rowIndex
                }
            }
        }
        
        # If we found a matching row
        if ($bestRow -ge 0 -and $bestScore -gt 0) {
            $matchedRow = $laborLookupGrid.Rows[$bestRow]
            Write-DebugLog "Using row $bestRow with score $bestScore% for calculations" -Category "CalculatedStaffing"
            
            # Extract values from the matched row for input columns
            foreach ($column in $inputColumns) {
                if ($gridColumns -contains $column) {
                    $newRow[$column] = $matchedRow.Cells[$column].Value
                    Write-DebugLog "Set $column = $($matchedRow.Cells[$column].Value)" -Category "CalculatedStaffing"
                }
            }
            
            # Extract values for output columns
            foreach ($column in $outputColumns) {
                if ($gridColumns -contains $column) {
                    $newRow[$column] = $matchedRow.Cells[$column].Value
                    Write-DebugLog "Set $column = $($matchedRow.Cells[$column].Value)" -Category "CalculatedStaffing"
                }
            }
            
            # Highlight the matched row for visual feedback
            $laborLookupGrid.ClearSelection()
            $matchedRow.Selected = $true
            if ($bestRow -ge 0) {
                $laborLookupGrid.FirstDisplayedScrollingRowIndex = $bestRow
            }
            $laborLookupGrid.Refresh()
            
            # --- FIND AND READ STAFFING FILE ---
            $machineAcronym = if ($MachineID -match "^([A-Za-z]+)") { $matches[1] } else { "" }
            $baseDir = "C:\Users\JR\Documents\Work\Programming Projects\LaborChecklist\Machine Labor Rubrics"
            
            Write-DebugLog "Looking for staffing data for Machine ID: $MachineID" -Category "CalculatedStaffing"
            
            # Get all directories that might contain staffing files
            $allDirs = Get-ChildItem $baseDir -Directory | Where-Object { $_.Name -like "*$MMO*" }
            Write-DebugLog "Found $($allDirs.Count) MMO directories to search" -Category "CalculatedStaffing"
            
            $foundFile = $false
            $staffingData = $null
            
            foreach ($dir in $allDirs) {
                Write-DebugLog "Checking directory: $($dir.FullName)" -Category "CalculatedStaffing"
                
                # Get all staffing table files in this directory
                $staffingFiles = Get-ChildItem $dir.FullName -Filter "*Staffing-Table.csv" -ErrorAction SilentlyContinue
                
                if ($staffingFiles.Count -gt 0) {
                    foreach ($file in $staffingFiles) {
                        Write-DebugLog "Found staffing file: $($file.FullName)" -Category "CalculatedStaffing"
                        
                        # Try to read the file to see if it contains our machine ID
                        try {
                            $content = Get-Content $file.FullName -ErrorAction Stop
                            
                            if ($content.Count -gt 0) {
                                # Check if any line contains our machine ID
                                $machineFound = $false
                                foreach ($line in $content) {
                                    if ($line -like "*$MachineID*") {
                                        Write-DebugLog "Machine ID found in line: $line" -Category "CalculatedStaffing"
                                        $machineFound = $true
                                        break
                                    }
                                }
                                
                                if ($machineFound) {
                                    # Process this file
                                    Write-DebugLog "Processing staffing file with matching machine ID: $($file.FullName)" -Category "CalculatedStaffing"
                                    
                                    # Parse header row
                                    $headerLine = $content[0].Trim()
                                    $headers = $headerLine.Split(',') | ForEach-Object { $_.Trim('"').Trim() }
                                    
                                    # Create DataTable to hold staffing data
                                    $staffingData = New-Object System.Data.DataTable
                                    
                                    # Add columns based on headers
                                    foreach ($header in $headers) {
                                        $staffingData.Columns.Add($header, [string]) | Out-Null
                                    }
                                    
                                    # Process data rows (skip header)
                                    for ($i = 1; $i -lt $content.Count; $i++) {
                                        $line = $content[$i].Trim()
                                        if ([string]::IsNullOrWhiteSpace($line)) { continue }
                                        
                                        # Parse CSV line
                                        $values = $line.Split(',') | ForEach-Object { $_.Trim('"').Trim() }
                                        
                                        # Create new row
                                        $staffingRow = $staffingData.NewRow()
                                        
                                        # Fill row with values
                                        for ($j = 0; $j -lt [Math]::Min($headers.Count, $values.Count); $j++) {
                                            $staffingRow[$headers[$j]] = $values[$j]
                                        }
                                        
                                        # Add row to table
                                        $staffingData.Rows.Add($staffingRow)
                                    }
                                    
                                    $foundFile = $true
                                    break
                                }
                            }
                        }
                        catch {
                            Write-DebugLog "Error reading staffing file: $_" -Category "CalculatedStaffing"
                        }
                    }
                }
                
                if ($foundFile) { break }
            }
            
            # If we found staffing data, extract management values
            if ($staffingData -ne $null -and $staffingData.Rows.Count -gt 0) {
                Write-DebugLog "Staffing data found with $($staffingData.Rows.Count) rows" -Category "CalculatedStaffing"
                
                # Find the row for our machine
                $machineRow = $null
                foreach ($row in $staffingData.Rows) {
                    if ($staffingData.Columns.Contains("Machine ID")) {
                        $rowMachineId = $row["Machine ID"].ToString().Trim()
                        if ($rowMachineId -eq $MachineID.Trim()) {
                            $machineRow = $row
                            Write-DebugLog "Found matching row in staffing data" -Category "CalculatedStaffing"
                            break
                        }
                    }
                }
                
                if ($machineRow -ne $null) {
                    # Check which management columns exist in the staffing data
                    $availableManagementColumns = @()
                    foreach ($column in $managementColumns) {
                        if ($staffingData.Columns.Contains($column)) {
                            $availableManagementColumns += $column
                        }
                    }
                    
                    Write-DebugLog "Available management columns: $($availableManagementColumns -join ', ')" -Category "CalculatedStaffing"
                    
                    if ($availableManagementColumns.Count -gt 0) {
                        try {
                            # Get the management values
                            $managementValues = @{}
                            $totalManagementValue = 0
                            
                            foreach ($column in $availableManagementColumns) {
                                if (-not [string]::IsNullOrWhiteSpace($machineRow[$column])) {
                                    $value = [double]::Parse($machineRow[$column].ToString())
                                    $managementValues[$column] = $value
                                    $totalManagementValue += $value
                                    Write-DebugLog "Got $column = $value" -Category "CalculatedStaffing"
                                }
                            }
                            
                            # If we have a total management value, calculate proportionate values
                            if ($totalManagementValue -gt 0 -and 
                                $calculatedTable.Columns.Contains("Total (hrs/yr)") -and 
                                -not [string]::IsNullOrWhiteSpace($newRow["Total (hrs/yr)"])) {
                                
                                $totalValue = [double]::Parse($newRow["Total (hrs/yr)"].ToString())
                                
                                # Calculate proportions for each management column
                                foreach ($column in $availableManagementColumns) {
                                    if ($managementValues.ContainsKey($column)) {
                                        $proportion = $managementValues[$column] / $totalManagementValue
                                        $calculatedValue = [Math]::Round($totalValue * $proportion, 2)
                                        $newRow[$column] = $calculatedValue
                                        Write-DebugLog "Calculated $column = $calculatedValue (proportion: $proportion)" -Category "CalculatedStaffing"
                                    }
                                }
                            }
                        }
                        catch {
                            Write-DebugLog "Error processing management values: $_" -Category "CalculatedStaffing"
                            # Set empty management values
                            foreach ($column in $managementColumns) {
                                if ($calculatedTable.Columns.Contains($column)) {
                                    $newRow[$column] = ""
                                }
                            }
                        }
                    }
                    else {
                        Write-DebugLog "No management columns found in staffing data" -Category "CalculatedStaffing"
                        # Set empty management values
                        foreach ($column in $managementColumns) {
                            if ($calculatedTable.Columns.Contains($column)) {
                                $newRow[$column] = ""
                            }
                        }
                    }
                }
                else {
                    Write-DebugLog "Machine not found in staffing data" -Category "CalculatedStaffing"
                    # Set empty management values
                    foreach ($column in $managementColumns) {
                        if ($calculatedTable.Columns.Contains($column)) {
                            $newRow[$column] = ""
                        }
                    }
                }
            }
            else {
                Write-DebugLog "No staffing data found - management values left blank" -Category "CalculatedStaffing"
                # Set empty management values
                foreach ($column in $managementColumns) {
                    if ($calculatedTable.Columns.Contains($column)) {
                        $newRow[$column] = ""
                    }
                }
            }
            
            # Add row to calculated table - THIS IS THE FIXED PART
            $calculatedTable.Rows.Add($newRow)
            
            # Explicitly return the DataTable - THIS IS THE FIXED PART
            Write-DebugLog "Returning DataTable with $($calculatedTable.Rows.Count) rows" -Category "CalculatedStaffing"
            return $calculatedTable
        }
        else {
            Write-DebugLog "No matching row found in labor lookup grid" -Category "CalculatedStaffing"
            return $null
        }
    }

    # Function to load calculated staffing table
    function global:Load-CalculatedStaffingTable {
        param (
            [string]$MMO,
            [string]$MachineID,
            [string]$ClassCode = "",
            [string]$MachineAcronym = "",
            [switch]$LoadAllMachines = $false
        )
        
        Write-DebugLog "Starting global:Load-CalculatedStaffingTable for MMO: $MMO, Machine ID: $MachineID, Class Code: $ClassCode" -Category "CalculatedStaffing"
        
        try {
            # Base directory path
            $baseDir = "C:\Users\JR\Documents\Work\Programming Projects\LaborChecklist\Machine Labor Rubrics"
            
            if (-not (Test-Path $baseDir)) {
                # Try the alternate path
                $baseDir = "C:\Users\JR\Documents\Work\Programming Projects\LaborChecklist\Machine Labor Rubrics"
                
                if (-not (Test-Path $baseDir)) {
                    Write-DebugLog "Base directory not found: $baseDir" -Category "CalculatedStaffing"
                    return $null
                }
            }

            # Find the most specific directory first
            $mmoDirectory = $null
            $searchPatterns = @(
                "*$MMO*$MachineAcronym*-$ClassCode*", 
                "*$MMO*-$ClassCode*",
                "*$MMO*$MachineAcronym*",
                "*$MMO*"
            )
            
            foreach ($pattern in $searchPatterns) {
                if ($mmoDirectory) { break }
                
                Write-DebugLog "Searching for directories with pattern: '$pattern'" -Category "CalculatedStaffing"
                $matchingDirs = Get-ChildItem $baseDir -Directory | Where-Object { $_.Name -like $pattern }
                
                if ($matchingDirs -and $matchingDirs.Count -gt 0) {
                    $mmoDirectory = $matchingDirs[0].FullName
                    Write-DebugLog "Found directory: $mmoDirectory" -Category "CalculatedStaffing"
                    break
                }
            }
            
            if (-not $mmoDirectory) {
                Write-DebugLog "No matching directory found for MMO: $MMO, Class Code: $ClassCode" -Category "CalculatedStaffing"
                return $null
            }

            # Define possible file patterns in order of specificity
            $filePatterns = @(
                "$MMO-$ClassCode-Calculated-Staffing-Table.csv",
                "$MachineAcronym-$ClassCode-Calculated-Staffing-Table.csv",
                "$MMO-Calculated-Staffing-Table.csv",
                "*Calculated-Staffing-Table.csv"
            )
            
            $calculatedFilePath = $null
            foreach ($filePattern in $filePatterns) {
                $matchingFiles = Get-ChildItem -Path $mmoDirectory -Filter $filePattern -ErrorAction SilentlyContinue
                
                if ($matchingFiles -and $matchingFiles.Count -gt 0) {
                    $calculatedFilePath = $matchingFiles[0].FullName
                    Write-DebugLog "Found calculated staffing file: $calculatedFilePath" -Category "CalculatedStaffing"
                    break
                }
            }
            
            if (-not $calculatedFilePath -or -not (Test-Path $calculatedFilePath)) {
                Write-DebugLog "Calculated staffing file not found" -Category "CalculatedStaffing"
                return $null
            }
            
            Write-DebugLog "Reading calculated staffing data from: $calculatedFilePath" -Category "CalculatedStaffing"
            
            # Create DataTable
            $calculatedTable = New-Object System.Data.DataTable
            
            # Read file content directly
            $content = Get-Content -Path $calculatedFilePath -ErrorAction Stop
            
            if ($content.Count -eq 0) {
                Write-DebugLog "Calculated staffing file is empty" -Category "CalculatedStaffing"
                return $null
            }
            
            # Parse header
            $headerLine = $content[0].Trim()
            Write-DebugLog "Header line: $headerLine" -Category "CalculatedStaffing"
            
            # Simply split by commas
            $headers = $headerLine.Split(',') | ForEach-Object { $_.Trim() }
            
            Write-DebugLog "Parsed headers: $($headers -join '|')" -Category "CalculatedStaffing"
            
            # Add columns to DataTable
            foreach ($header in $headers) {
                $calculatedTable.Columns.Add($header, [string]) | Out-Null
            }
            
            # Find Machine ID column index
            $machineIdColIndex = -1
            for ($i = 0; $i -lt $headers.Count; $i++) {
                if ($headers[$i] -eq "Machine ID") {
                    $machineIdColIndex = $i
                    break
                }
            }
            
            # Process data rows (skip header)
            for ($i = 1; $i -lt $content.Count; $i++) {
                $line = $content[$i].Trim()
                if ([string]::IsNullOrWhiteSpace($line)) { continue }
                
                Write-DebugLog "Processing data line ${i}: $line" -Category "CalculatedStaffing"
                
                # Simply split by commas
                $values = $line.Split(',') | ForEach-Object { $_.Trim() }
                
                # Check if this row is for our machine or we're loading all machines
                $includeRow = $LoadAllMachines
                if ($machineIdColIndex -ge 0 -and $machineIdColIndex -lt $values.Count) {
                    $rowMachineId = $values[$machineIdColIndex]
                    
                    if (-not [string]::IsNullOrWhiteSpace($MachineID) -and $rowMachineId -eq $MachineID) {
                        $includeRow = $true
                        Write-DebugLog "Found matching machine ID: $rowMachineId" -Category "CalculatedStaffing"
                    }
                }
                
                if ($includeRow) {
                    # Make sure values array has enough elements
                    while ($values.Count -lt $headers.Count) {
                        $values += ""
                    }
                    
                    # Add row to DataTable
                    $newRow = $calculatedTable.NewRow()
                    for ($j = 0; $j -lt $headers.Count; $j++) {
                        $newRow[$headers[$j]] = $values[$j]
                    }
                    $calculatedTable.Rows.Add($newRow)
                    Write-DebugLog "Added row for machine: $rowMachineId" -Category "CalculatedStaffing"
                }
            }
            
            if ($calculatedTable.Rows.Count -gt 0) {
                Write-DebugLog "Successfully loaded calculated staffing table with $($calculatedTable.Rows.Count) rows" -Category "CalculatedStaffing"
                return $calculatedTable
            } else {
                Write-DebugLog "No matching rows found in calculated staffing data" -Category "CalculatedStaffing"
                return $null
            }
        }
        catch {
            Write-DebugLog "Error loading calculated staffing table: $_" -Category "CalculatedStaffing"
            Write-DebugLog "Stack trace: $($_.ScriptStackTrace)" -Category "CalculatedStaffing"
            return $null
        }
    }

    # Function to save calculated staffing table
    function Save-CalculatedStaffingTable {
        param (
            [System.Data.DataTable]$DataTable,
            [string]$MMO,
            [string]$MachineID,
            [string]$ClassCode = "",
            [string]$MachineAcronym = ""
        )
        
        Write-DebugLog "Saving calculated staffing table for MMO: $MMO, Machine ID: $MachineID" -Category "CalculatedStaffing"
        
        try {
            # Base directory
            $baseDir = "C:\Users\JR\Documents\Work\Programming Projects\LaborChecklist\Machine Labor Rubrics"
            if (-not (Test-Path $baseDir)) {
                $baseDir = "C:\Users\JR\Documents\Work\Programming Projects\LaborChecklist\Machine Labor Rubrics"
            }
            
            # Find suitable directory (same as staffing table)
            $mmoDirectory = $null
            
            # Search patterns in order of specificity
            if (-not [string]::IsNullOrWhiteSpace($ClassCode) -and -not [string]::IsNullOrWhiteSpace($MachineAcronym)) {
                $matchingDirs = Get-ChildItem $baseDir -Directory | 
                    Where-Object { $_.Name -like "*$MMO*$MachineAcronym*-$ClassCode*" }
                
                if ($matchingDirs.Count -gt 0) {
                    $mmoDirectory = $matchingDirs[0].FullName
                }
            }
            
            if (-not $mmoDirectory -and -not [string]::IsNullOrWhiteSpace($ClassCode)) {
                $matchingDirs = Get-ChildItem $baseDir -Directory | 
                    Where-Object { $_.Name -like "*$MMO*-$ClassCode*" }
                
                if ($matchingDirs.Count -gt 0) {
                    $mmoDirectory = $matchingDirs[0].FullName
                }
            }
            
            if (-not $mmoDirectory) {
                $matchingDirs = Get-ChildItem $baseDir -Directory | 
                    Where-Object { $_.Name -like "*$MMO*" }
                
                if ($matchingDirs.Count -gt 0) {
                    $mmoDirectory = $matchingDirs[0].FullName
                }
            }
            
            # If still no directory, create one
            if (-not $mmoDirectory) {
                $dirName = if (-not [string]::IsNullOrWhiteSpace($ClassCode)) {
                    "$MMO-$ClassCode"
                } else {
                    "$MMO"
                }
                
                $mmoDirectory = Join-Path $baseDir $dirName
                New-Item -Path $mmoDirectory -ItemType Directory -Force | Out-Null
            }
            
            # Define filename for calculated staffing table
            $calculatedFileName = if (-not [string]::IsNullOrWhiteSpace($ClassCode)) {
                "$MMO-$ClassCode-Calculated-Staffing-Table.csv"
            } else {
                "$MMO-Calculated-Staffing-Table.csv"
            }
            
            $calculatedFilePath = Join-Path $mmoDirectory $calculatedFileName
            
            # Check if file exists and load existing data
            $allMachines = @()
            
            if (Test-Path $calculatedFilePath) {
                Write-DebugLog "Existing calculated file found: $calculatedFilePath" -Category "CalculatedStaffing"
                
                try {
                    # Read existing data (without quotes)
                    $lines = Get-Content -Path $calculatedFilePath | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
                    
                    if ($lines.Count -gt 0) {
                        # Get headers
                        $headers = $lines[0].Split(',') | ForEach-Object { $_.Trim() }
                        
                        # Process data rows
                        for ($i = 1; $i -lt $lines.Count; $i++) {
                            $values = $lines[$i].Split(',') | ForEach-Object { $_.Trim() }
                            
                            # Create object
                            $rowObj = New-Object PSObject
                            
                            # Add properties
                            for ($j = 0; $j -lt [Math]::Min($headers.Count, $values.Count); $j++) {
                                $rowObj | Add-Member -MemberType NoteProperty -Name $headers[$j] -Value $values[$j]
                            }
                            
                            # Skip the machine we're updating
                            if ($rowObj.'Machine ID' -ne $MachineID) {
                                $allMachines += $rowObj
                            }
                        }
                    }
                }
                catch {
                    Write-DebugLog "Error reading existing calculated file: $_" -Category "CalculatedStaffing"
                }
            }
            
            # Convert DataTable to objects
            foreach ($row in $DataTable.Rows) {
                $rowObj = New-Object PSObject
                
                foreach ($column in $DataTable.Columns) {
                    $rowObj | Add-Member -MemberType NoteProperty -Name $column.ColumnName -Value $row[$column.ColumnName]
                }
                
                $allMachines += $rowObj
            }
            
            # Save all machines without quotes
            
            # Get all property names
            $propertyNames = @()
            foreach ($machine in $allMachines) {
                foreach ($prop in $machine.PSObject.Properties) {
                    if ($propertyNames -notcontains $prop.Name) {
                        $propertyNames += $prop.Name
                    }
                }
            }
            
            # Ensure Machine ID is first
            if ($propertyNames -contains "Machine ID") {
                $propertyNames = @("Machine ID") + ($propertyNames | Where-Object { $_ -ne "Machine ID" })
            }
            
            # Create header row
            $headerRow = $propertyNames -join ","
            
            # Create content
            $content = @($headerRow)
            
            # Add data rows
            foreach ($machine in $allMachines) {
                $rowValues = @()
                
                foreach ($prop in $propertyNames) {
                    $value = if ($machine.PSObject.Properties.Name -contains $prop) {
                        if ($machine.$prop -ne $null) {
                            ($machine.$prop).ToString() -replace ",", " "
                        } else {
                            ""
                        }
                    } else {
                        ""
                    }
                    
                    $rowValues += $value
                }
                
                $content += ($rowValues -join ",")
            }
            
            # Write to file
            Set-Content -Path $calculatedFilePath -Value $content
            
            Write-DebugLog "Saved calculated staffing table with $($allMachines.Count) rows" -Category "CalculatedStaffing"
            
            return $true
        }
        catch {
            Write-DebugLog "Error saving calculated staffing table: $_" -Category "CalculatedStaffing"
            return $false
        }
    }

    $btnCalculate.Add_Click({
        Write-DebugLog "Calculate button clicked" -Category "CalculatedStaffing"
        
        # Extract machine metrics from ListView
        $machineMetrics = @{
            MachineType = $acronym
            MachineNumber = $machineNumber
        }
        
        # Add values from ListView
        foreach ($index in $columnMap.Keys) {
            if ($selectedItem.SubItems.Count -gt $index -and 
                -not [string]::IsNullOrWhiteSpace($selectedItem.SubItems[$index].Text)) {
                $paramName = $columnMap[$index]
                $machineMetrics[$paramName] = $selectedItem.SubItems[$index].Text
                Write-DebugLog "Added metric from ListView: $paramName = $($selectedItem.SubItems[$index].Text)" -Category "CalculatedStaffing"
            }
        }
        
        # Highlight cells in the labor lookup grid
        Highlight-MatchedCells -DataGridView $laborLookupGrid -MachineMetrics $machineMetrics -ColumnMapping $columnMap
        
        # Force UI update
        $laborLookupGrid.Refresh()
        [System.Windows.Forms.Application]::DoEvents()
        
        # Find the lookup table path for this MMO/machine
        $lookupTablePath = ""
        if (-not [string]::IsNullOrWhiteSpace($mmo)) {
            $baseDir = "C:\Users\JR\Documents\Work\Programming Projects\LaborChecklist\Machine Labor Rubrics"
            
            # Try alternate path if first one doesn't exist
            if (-not (Test-Path $baseDir)) {
                $baseDir = "C:\Users\JR\Documents\Work\Programming Projects\LaborChecklist\Machine Labor Rubrics"
            }
            
            # First try to find a directory that matches the MMO and class code
            $mmoDirectories = Get-ChildItem $baseDir -Directory | 
                Where-Object { $_.Name -like "*$mmo*" -and $_.Name -like "*$classCode*" }
            
            if ($mmoDirectories.Count -gt 0) {
                # Use the first matching directory
                $mmoDirectory = $mmoDirectories[0].FullName
                Write-DebugLog "Found directory matching MMO and class code: $mmoDirectory" -Category "CalculatedStaffing"
                
                # Look for labor lookup file in that directory
                $lookupFiles = Get-ChildItem $mmoDirectory -Filter "*-Labor-Lookup.csv"
                if ($lookupFiles.Count -gt 0) {
                    $lookupTablePath = $lookupFiles[0].FullName
                    Write-DebugLog "Found lookup table: $lookupTablePath" -Category "CalculatedStaffing"
                }
            }
            # If not found, try just the MMO
            elseif ([string]::IsNullOrWhiteSpace($lookupTablePath)) {
                $mmoDirectories = Get-ChildItem $baseDir -Directory | 
                    Where-Object { $_.Name -like "*$mmo*" }
                
                if ($mmoDirectories.Count -gt 0) {
                    # Use the first matching directory
                    $mmoDirectory = $mmoDirectories[0].FullName
                    Write-DebugLog "Found directory matching just MMO: $mmoDirectory" -Category "CalculatedStaffing"
                    
                    # Look for labor lookup file in that directory
                    $lookupFiles = Get-ChildItem $mmoDirectory -Filter "*-Labor-Lookup.csv"
                    if ($lookupFiles.Count -gt 0) {
                        $lookupTablePath = $lookupFiles[0].FullName
                        Write-DebugLog "Found lookup table: $lookupTablePath" -Category "CalculatedStaffing"
                    }
                }
            }
        }
        
        # Now use the lookup table to dynamically determine columns
        # Define our standard output columns
        $outputColumns = @("Total (hrs/yr)", "Operational Maintenance (hrs/yr)")
        $managementColumns = @("MM7", "MPE9", "ET10")
        $metadataColumns = @("Machine ID")
        
        # Load lookup table if available to get columns and values
        $lookupData = $null
        $inputColumnNames = @()
        $allColumnNames = @()
        
        # Safely check if lookup table exists
        if (-not [string]::IsNullOrWhiteSpace($lookupTablePath) -and (Test-Path $lookupTablePath)) {
            try {
                $lookupData = Import-Csv $lookupTablePath
                Write-DebugLog "Loaded lookup table: $lookupTablePath with $($lookupData.Count) rows" -Category "CalculatedStaffing"
                
                # Get column names from the lookup table
                $allColumnNames = $lookupData[0].PSObject.Properties.Name
                Write-DebugLog "Found columns: $($allColumnNames -join ', ')" -Category "CalculatedStaffing"
                
                # Filter out output columns to get input columns
                $inputColumnNames = $allColumnNames | Where-Object { $_ -notin $outputColumns }
                Write-DebugLog "Input columns for calculation: $($inputColumnNames -join ', ')" -Category "CalculatedStaffing"
            }
            catch {
                Write-DebugLog "Error loading lookup table: $_" -Category "CalculatedStaffing"
                $lookupData = $null
            }
        } else {
            Write-DebugLog "No valid lookup table available at path: $lookupTablePath" -Category "CalculatedStaffing"
        }
        
        # If no columns found, use a default set
        if ($inputColumnNames.Count -eq 0) {
            $inputColumnNames = @("Operation (days/wk)", "Tours/Day")
            Write-DebugLog "Using default input columns" -Category "CalculatedStaffing"
        }
        
        # Create DataTable with dynamically determined columns
        $calculatedTable = New-Object System.Data.DataTable
        
        # Add Machine ID column first
        $calculatedTable.Columns.Add("Machine ID", [string]) | Out-Null
        
        # Add Input columns from lookup table
        foreach ($column in $inputColumnNames) {
            if (-not $calculatedTable.Columns.Contains($column)) {
                $calculatedTable.Columns.Add($column, [string]) | Out-Null
            }
        }
        
        # Add Management columns
        foreach ($column in $managementColumns) {
            if (-not $calculatedTable.Columns.Contains($column)) {
                $calculatedTable.Columns.Add($column, [string]) | Out-Null
            }
        }
        
        # Add Output columns
        foreach ($column in $outputColumns) {
            if (-not $calculatedTable.Columns.Contains($column)) {
                $calculatedTable.Columns.Add($column, [string]) | Out-Null
            }
        }
        
        # Create new row
        $newRow = $calculatedTable.NewRow()
        $newRow["Machine ID"] = $machineID
        
        # Get the selected row from Labor Lookup grid (the highlighted row)
        $selectedRow = $null
        foreach ($row in $laborLookupGrid.Rows) {
            if ($row.Selected) {
                $selectedRow = $row
                break
            }
        }
        
        if ($selectedRow -eq $null) {
            # Try to find the best matching row if none is selected
            Write-DebugLog "No row selected, finding best match" -Category "CalculatedStaffing"
            $bestScore = 0
            $bestRow = -1
            
            for ($rowIndex = 0; $rowIndex -lt $laborLookupGrid.Rows.Count; $rowIndex++) {
                $row = $laborLookupGrid.Rows[$rowIndex]
                $rowScore = 0
                $matchCount = 0
                $totalParams = 0
                
                # Calculate score based on matching parameters
                foreach ($paramName in $machineMetrics.Keys) {
                    if ($paramName -eq "MachineType" -or $paramName -eq "MachineNumber") { continue }
                    if ([string]::IsNullOrWhiteSpace($machineMetrics[$paramName])) { continue }
                    
                    $totalParams++
                    
                    if ($row.Cells.Count -gt 0 -and $laborLookupGrid.Columns.Contains($paramName) -and
                        $row.Cells[$paramName] -ne $null) {
                        $cellValue = $row.Cells[$paramName].Value
                        $metricValue = $machineMetrics[$paramName]
                        
                        if ($cellValue -eq $metricValue) {
                            $matchCount++
                            $rowScore += 10
                        }
                    }
                }
                
                if ($totalParams -gt 0) {
                    $finalScore = ($rowScore / ($totalParams * 10)) * 100
                    Write-DebugLog "Row $rowIndex score: $finalScore% ($matchCount/$totalParams matches)" -Category "CalculatedStaffing"
                    if ($finalScore -gt $bestScore) {
                        $bestScore = $finalScore
                        $bestRow = $rowIndex
                    }
                }
            }
            
            if ($bestRow -ge 0) {
                $selectedRow = $laborLookupGrid.Rows[$bestRow]
                $selectedRow.Selected = $true
                if ($bestRow -ge 0) {
                    $laborLookupGrid.FirstDisplayedScrollingRowIndex = $bestRow
                }
                $laborLookupGrid.Refresh()
                Write-DebugLog "Selected best matching row: $bestRow with score $bestScore%" -Category "CalculatedStaffing"
            }
        }
        
        if ($selectedRow -eq $null) {
            [System.Windows.Forms.MessageBox]::Show(
                "Could not find a matching row in the Labor Lookup table.",
                "No Match Found",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            return
        }
        
        # Copy values from selected labor lookup row to our new row
        foreach ($column in $inputColumnNames) {
            if ($laborLookupGrid.Columns.Contains($column) -and $selectedRow.Cells[$column] -ne $null) {
                $newRow[$column] = $selectedRow.Cells[$column].Value
                Write-DebugLog "Set $column = $($selectedRow.Cells[$column].Value)" -Category "CalculatedStaffing"
            }
        }
        
        # Copy output column values
        foreach ($column in $outputColumns) {
            if ($laborLookupGrid.Columns.Contains($column) -and $selectedRow.Cells[$column] -ne $null) {
                $newRow[$column] = $selectedRow.Cells[$column].Value
                Write-DebugLog "Set $column = $($selectedRow.Cells[$column].Value)" -Category "CalculatedStaffing"
            }
        }
        
        # Load staffing table data for management hours
        try {
            # Determine the base dir
            $baseDir = "C:\Users\JR\Documents\Work\Programming Projects\LaborChecklist\Machine Labor Rubrics"
            if (-not (Test-Path $baseDir)) {
                $baseDir = "C:\Users\JR\Documents\Work\Programming Projects\LaborChecklist\Machine Labor Rubrics"
            }
            
            # Extract machine acronym
            $machineAcronym = if ($machineID -match "^([A-Za-z]+)") { $matches[1] } else { "" }
            
            # Search for staffing table using patterns similar to Load-StaffingTable
            $staffingTable = $null
            
            # Define search patterns in order of specificity
            $searchPatterns = @(
                "*$mmo*$machineAcronym*-$classCode*", 
                "*$mmo*-$classCode*",
                "*$mmo*$machineAcronym*",
                "*$mmo*"
            )
            
            $staffingFilePath = $null
            foreach ($pattern in $searchPatterns) {
                if ($staffingFilePath) { break }
                
                Write-DebugLog "Searching for directories with pattern: '$pattern'" -Category "CalculatedStaffing"
                $matchingDirs = Get-ChildItem $baseDir -Directory | Where-Object { $_.Name -like $pattern }
                
                if ($matchingDirs -and $matchingDirs.Count -gt 0) {
                    foreach ($dir in $matchingDirs) {
                        Write-DebugLog "Checking directory: $($dir.FullName)" -Category "CalculatedStaffing"
                        
                        # Define possible file patterns in order of specificity
                        $filePatterns = @(
                            "$mmo-$classCode-Staffing-Table.csv",
                            "$machineAcronym-$classCode-Staffing-Table.csv",
                            "$mmo-Staffing-Table.csv",
                            "*-Staffing-Table.csv"
                        )
                        
                        foreach ($filePattern in $filePatterns) {
                            $matchingFiles = Get-ChildItem -Path $dir.FullName -Filter $filePattern -ErrorAction SilentlyContinue
                            
                            if ($matchingFiles -and $matchingFiles.Count -gt 0) {
                                $staffingFilePath = $matchingFiles[0].FullName
                                Write-DebugLog "Found staffing file: $staffingFilePath" -Category "CalculatedStaffing"
                                break
                            }
                        }
                        
                        if ($staffingFilePath) { break }
                    }
                }
            }
            
            # Load data from staffing file if found
            if ($staffingFilePath -and (Test-Path $staffingFilePath)) {
                Write-DebugLog "Loading data from staffing file: $staffingFilePath" -Category "CalculatedStaffing"
                
                # Read the staffing table file
                $staffingContent = Get-Content -Path $staffingFilePath -Raw -ErrorAction Stop
                
                # Parse the staffing CSV (same approach as in staffing tab)
                $staffingContent = $staffingContent.Replace("`r`n", "`n").Replace("`r", "`n").Trim()
                $staffingLines = $staffingContent -split "`n" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
                
                if ($staffingLines.Count -gt 0) {
                    # Parse header
                    $headerLine = $staffingLines[0]
                    $staffingHeaders = @()
                    
                    # Handle quoted or unquoted format
                    if ($headerLine.StartsWith('"') -and $headerLine.Contains('","')) {
                        # Quoted format
                        $staffingHeaders = $headerLine.Split('","') | ForEach-Object {
                            $_.Replace('"', '').Trim()
                        }
                        # Fix first and last items
                        if ($staffingHeaders.Count -gt 0) {
                            $staffingHeaders[0] = $staffingHeaders[0].TrimStart('"')
                            $staffingHeaders[$staffingHeaders.Count-1] = $staffingHeaders[$staffingHeaders.Count-1].TrimEnd('"')
                        }
                    } else {
                        # Simple format
                        $staffingHeaders = $headerLine.Split(',') | ForEach-Object { $_.Trim('"').Trim() }
                    }
                    
                    Write-DebugLog "Staffing headers: $($staffingHeaders -join ', ')" -Category "CalculatedStaffing"
                    
                    # Find Machine ID column index
                    $machineIdColIndex = -1
                    for ($i = 0; $i -lt $staffingHeaders.Count; $i++) {
                        if ($staffingHeaders[$i] -eq "Machine ID") {
                            $machineIdColIndex = $i
                            break
                        }
                    }
                    
                    # Find management columns indices
                    $mm7Index = $mpe9Index = $et10Index = $yearlyRunIndex = -1
                    for ($i = 0; $i -lt $staffingHeaders.Count; $i++) {
                        switch ($staffingHeaders[$i]) {
                            "MM7" { $mm7Index = $i }
                            "MPE9" { $mpe9Index = $i }
                            "ET10" { $et10Index = $i }
                        }
                    }
                    
                    # Find the row for our machine
                    for ($lineIndex = 1; $lineIndex -lt $staffingLines.Count; $lineIndex++) {
                        $line = $staffingLines[$lineIndex]
                        
                        # Parse the line based on format
                        $values = @()
                        if ($line.StartsWith('"') -and $line.Contains('","')) {
                            # Quoted format
                            $values = $line.Split('","') | ForEach-Object {
                                $_.Replace('"', '').Trim()
                            }
                            # Fix first and last items
                            if ($values.Count -gt 0) {
                                $values[0] = $values[0].TrimStart('"')
                                $values[$values.Count-1] = $values[$values.Count-1].TrimEnd('"')
                            }
                        } else {
                            # Simple format
                            $values = $line.Split(',') | ForEach-Object { $_.Trim('"').Trim() }
                        }
                        
                        # Check if field values count matches header count
                        if ($values.Count -ne $staffingHeaders.Count) {
                            # Fix mismatch
                            while ($values.Count -lt $staffingHeaders.Count) {
                                $values += ""
                            }
                            if ($values.Count -gt $staffingHeaders.Count) {
                                $values = $values[0..($staffingHeaders.Count-1)]
                            }
                        }
                        
                        # Check if this is our machine
                        if ($machineIdColIndex -ge 0 -and $machineIdColIndex -lt $values.Count) {
                            $rowMachineId = $values[$machineIdColIndex].Trim()
                            if ($rowMachineId -eq $machineID.Trim()) {
                                Write-DebugLog "Found matching machine in staffing table: $rowMachineId" -Category "CalculatedStaffing"
                                
                                # Extract values
                                $mm7Value = $mpe9Value = $et10Value = $yearlyRunValue = ""
                                
                                if ($mm7Index -ge 0 -and $mm7Index -lt $values.Count) {
                                    $mm7Value = $values[$mm7Index]
                                }
                                
                                if ($mpe9Index -ge 0 -and $mpe9Index -lt $values.Count) {
                                    $mpe9Value = $values[$mpe9Index]
                                }
                                
                                if ($et10Index -ge 0 -and $et10Index -lt $values.Count) {
                                    $et10Value = $values[$et10Index]
                                }
                                
                                if ($yearlyRunIndex -ge 0 -and $yearlyRunIndex -lt $values.Count) {
                                    $yearlyRunValue = $values[$yearlyRunIndex]
                                }
                                
                                # Store values or calculate proportions
                                if (-not [string]::IsNullOrWhiteSpace($newRow["Total (hrs/yr)"])) {
                                    # Get total hours
                                    $totalHours = 0
                                    try {
                                        $totalHours = [double]::Parse($newRow["Total (hrs/yr)"].ToString())
                                    } catch {
                                        Write-DebugLog "Error parsing Total hours: $_" -Category "CalculatedStaffing"
                                    }
                                    
                                    if ($totalHours -gt 0) {
                                        # Check if we have management ratios
                                        $mm7 = $mpe9 = $et10 = 0
                                        $totalRatio = 0
                                        
                                        if (-not [string]::IsNullOrWhiteSpace($mm7Value)) {
                                            try {
                                                $mm7 = [double]::Parse($mm7Value)
                                                $totalRatio += $mm7
                                            } catch {
                                                Write-DebugLog "Error parsing MM7: $_" -Category "CalculatedStaffing"
                                            }
                                        }
                                        
                                        if (-not [string]::IsNullOrWhiteSpace($mpe9Value)) {
                                            try {
                                                $mpe9 = [double]::Parse($mpe9Value)
                                                $totalRatio += $mpe9
                                            } catch {
                                                Write-DebugLog "Error parsing MPE9: $_" -Category "CalculatedStaffing"
                                            }
                                        }
                                        
                                        if (-not [string]::IsNullOrWhiteSpace($et10Value)) {
                                            try {
                                                $et10 = [double]::Parse($et10Value)
                                                $totalRatio += $et10
                                            } catch {
                                                Write-DebugLog "Error parsing ET10: $_" -Category "CalculatedStaffing"
                                            }
                                        }
                                        
                                        if ($totalRatio -gt 0) {
                                            # Calculate proportions
                                            $newRow["MM7"] = [Math]::Round(($mm7 / $totalRatio) * $totalHours, 2)
                                            $newRow["MPE9"] = [Math]::Round(($mpe9 / $totalRatio) * $totalHours, 2)
                                            $newRow["ET10"] = [Math]::Round(($et10 / $totalRatio) * $totalHours, 2)
                                            Write-DebugLog "Calculated management hours based on ratios - MM7: $($newRow["MM7"]), MPE9: $($newRow["MPE9"]), ET10: $($newRow["ET10"])" -Category "CalculatedStaffing"
                                        } else {
                                            # Equal distribution
                                            $newRow["MM7"] = [Math]::Round($totalHours / 3, 2)
                                            $newRow["MPE9"] = [Math]::Round($totalHours / 3, 2)
                                            $newRow["ET10"] = [Math]::Round($totalHours / 3, 2)
                                            Write-DebugLog "Using equal distribution for management hours" -Category "CalculatedStaffing"
                                        }
                                    } else {
                                        # Set raw values
                                        $newRow["MM7"] = $mm7Value
                                        $newRow["MPE9"] = $mpe9Value
                                        $newRow["ET10"] = $et10Value
                                    }
                                } else {
                                    # Set raw values
                                    $newRow["MM7"] = $mm7Value
                                    $newRow["MPE9"] = $mpe9Value
                                    $newRow["ET10"] = $et10Value
                                }
                                
                                break
                            }
                        }
                    }
                }
            } else {
                Write-DebugLog "No staffing table found for this machine" -Category "CalculatedStaffing"
                
                # Default values if no staffing data found
                if ($newRow["Total (hrs/yr)"] -ne $null -and $newRow["Total (hrs/yr)"] -ne "") {
                    try {
                        $totalHours = [double]::Parse($newRow["Total (hrs/yr)"].ToString())
                        $newRow["MM7"] = [Math]::Round($totalHours / 3, 2)
                        $newRow["MPE9"] = [Math]::Round($totalHours / 3, 2)
                        $newRow["ET10"] = [Math]::Round($totalHours / 3, 2)
                        Write-DebugLog "Set default management hours (equal distribution)" -Category "CalculatedStaffing"
                    } catch {
                        Write-DebugLog "Error setting default management hours: $_" -Category "CalculatedStaffing"
                    }
                }
            }
        } catch {
            Write-DebugLog "Error loading staffing data: $_" -Category "CalculatedStaffing"
            Write-DebugLog "Stack trace: $($_.ScriptStackTrace)" -Category "CalculatedStaffing"
        }
        
        # Add the row to the table
        $calculatedTable.Rows.Add($newRow)
        Write-DebugLog "Added row to calculated table" -Category "CalculatedStaffing"
        
        # Clear current DataSource and set new one
        $calculatedStaffingGrid.DataSource = $null
        [System.Windows.Forms.Application]::DoEvents()
        
        $calculatedStaffingGrid.DataSource = $calculatedTable
        Write-DebugLog "Set DataSource to calculated table" -Category "CalculatedStaffing"
        
        # Force refresh
        $calculatedStaffingGrid.Refresh()
        $calculatedStaffingGrid.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells)
        [System.Windows.Forms.Application]::DoEvents()
        
        # Check if rendering succeeded
        if ($calculatedStaffingGrid.Rows.Count -gt 0) {
            Write-DebugLog "Successfully rendered calculated staffing table with $($calculatedTable.Rows.Count) rows" -Category "CalculatedStaffing"
            
            # Check if MM7, MPE9, ET10 values are populated
            $managementDataMissing = [string]::IsNullOrWhiteSpace($newRow["MM7"]) -or
                            [string]::IsNullOrWhiteSpace($newRow["MPE9"]) -or
                            [string]::IsNullOrWhiteSpace($newRow["ET10"])
            
            # Prompt to save
            $saveResult = [System.Windows.Forms.MessageBox]::Show(
                "Calculated staffing requirements generated successfully." + 
                $(if ($managementDataMissing) { "`n`nNOTE: MM7, MPE9, ET10 values are missing or incomplete. Please enter management data in the Staffing Table tab first." } else { "" }) +
                "`n`nData is displayed in the grid. Would you like to save these calculations?",
                "Save Calculated Staffing",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )
            
            if ($saveResult -eq [System.Windows.Forms.DialogResult]::Yes) {
                # Save the calculated table
                $saveSuccess = Save-CalculatedStaffingTable -DataTable $calculatedTable -MMO $mmo -MachineID $machineID -ClassCode $classCode -MachineAcronym $machineAcronym
                
                if ($saveSuccess) {
                    [System.Windows.Forms.MessageBox]::Show(
                        "Calculated staffing requirements saved successfully!",
                        "Save Successful",
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Information
                    )
                } else {
                    [System.Windows.Forms.MessageBox]::Show(
                        "Failed to save calculated staffing requirements.",
                        "Save Failed",
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Error
                    )
                }
            }
        } else {
            # If rendering still failed, try the tab-switching approach
            Write-DebugLog "Grid rendering failed, trying tab-switch approach" -Category "CalculatedStaffing"
            
            # Store the data in a global variable
            $global:PendingCalculatedStaffingData = $calculatedTable
            
            # Switch to another tab and back
            if ($tabControl.SelectedTab -eq $calculatedStaffingTab -and $tabControl.TabPages.Count -gt 1) {
                # Switch to another tab and back using a timer
                $currentIndex = $tabControl.SelectedIndex
                $otherIndex = ($currentIndex + 1) % $tabControl.TabPages.Count
                
                $tabControl.SelectedIndex = $otherIndex
                [System.Windows.Forms.Application]::DoEvents()
                
                $timer = New-Object System.Windows.Forms.Timer
                $timer.Interval = 200
                $timer.Add_Tick({
                    $tabControl.SelectedIndex = $currentIndex
                    $timer.Stop()
                    $timer.Dispose()
                    
                    # Try to set the data again
                    if ($global:PendingCalculatedStaffingData -ne $null) {
                        $calculatedStaffingGrid.DataSource = $global:PendingCalculatedStaffingData
                        $calculatedStaffingGrid.Refresh()
                        $calculatedStaffingGrid.AutoResizeColumns()
                        [System.Windows.Forms.Application]::DoEvents()
                    }
                })
                $timer.Start()
            }
        }
    })


    
    # Add Edit button to staffing tab
    $btnEditStaffing = New-Object System.Windows.Forms.Button -Property @{
        Text = "Edit Staffing Data"
        Size = New-Object System.Drawing.Size(120, 30)
        Location = New-Object System.Drawing.Point(10, 10)
        Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
    }
    
    # Add button click handler
    $btnEditStaffing.Add_Click({
        # Find lookup table to get input columns
        $lookupTablePath = Get-ChildItem $mmoDirectory -Filter "*-Labor-Lookup.csv" | 
            Select-Object -First 1 -ExpandProperty FullName
                
        if (-not $lookupTablePath) {
            Write-DebugLog "Labor lookup file not found in directory: $mmoDirectory" -Category "StaffingTable"
            [System.Windows.Forms.MessageBox]::Show(
                "Labor lookup file not found. Please configure the machine first.",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            return
        }
        
        # Extract machine acronym from machine ID
        $machineAcronym = ""
        if ($machineID -match "^([A-Za-z]+)") {
            $machineAcronym = $matches[1]
        }
        
        # Load existing staffing data if available
        $existingDataResult = Load-StaffingTable -MMO $mmo -MachineID $machineID -ClassCode $classCode -MachineAcronym $machineAcronym
        
        # Convert to DataTable if it's a DataRow
        $existingData = $null
        if ($existingDataResult -is [System.Data.DataRow]) {
            Write-DebugLog "Converting DataRow to DataTable for dialog" -Category "StaffingTable"
            # Create a new DataTable with the same schema
            $existingData = $existingDataResult.Table.Clone()
            # Add the row to the new table
            $existingData.Rows.Add($existingDataResult.ItemArray)
        } 
        elseif ($existingDataResult -is [System.Data.DataTable]) {
            $existingData = $existingDataResult
        }
        else {
            # Create an empty DataTable as fallback
            $existingData = New-Object System.Data.DataTable
            
            # Add basic columns
            $basicColumns = @("Machine ID", "Operation (days/wk)", "Tours/Day", "MM7", "MPE9", "ET10", 
                            "Total (hrs/yr)", "Operational Maintenance (hrs/yr)")
            foreach ($col in $basicColumns) {
                $existingData.Columns.Add($col, [string])
            }
        }
        
        # Show the staffing table dialog - PASS THE CLASS CODE AND MACHINE ACRONYM
        $staffingResult = Show-StaffingTableDialog -MachineID $machineID -MMO $mmo -ClassCode $classCode -MachineAcronym $machineAcronym -LookupTablePath $lookupTablePath -ExistingData $existingData
        
        # If dialog was successful, reload the staffing table
        if ($staffingResult -eq [System.Windows.Forms.DialogResult]::OK) {
            # Reload the staffing table data - with CLASS CODE and MACHINE ACRONYM
            $refreshedData = Load-StaffingTable -MMO $mmo -MachineID $machineID -ClassCode $classCode -MachineAcronym $machineAcronym
            
            if ($refreshedData) {
                # Clear the existing DataSource to force a refresh
                $staffingTableGrid.DataSource = $null
                # Set the new DataSource
                $staffingTableGrid.DataSource = $refreshedData
                $staffingTableGrid.Refresh()
                $staffingTableGrid.AutoResizeColumns()
                
                # Force visual update
                [System.Windows.Forms.Application]::DoEvents()
            }
            else {
                Write-DebugLog "Failed to reload staffing data after editing" -Category "StaffingTable"
                [System.Windows.Forms.MessageBox]::Show(
                    "Data was saved successfully but couldn't be reloaded. Please try closing and reopening this window.",
                    "Warning",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                )
            }
        }
    })

    # Add a Debug button to the staffing panel
    $btnDebugBinding = New-Object System.Windows.Forms.Button -Property @{
        Text = "Debug Data Binding"
        Size = New-Object System.Drawing.Size(120, 30)
        Location = New-Object System.Drawing.Point(140, 10)
        Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
    }

    $btnDebugBinding.Add_Click({
        # Extract machine acronym
        $machineAcronym = ""
        if ($machineID -match "^([A-Za-z]+)") {
            $machineAcronym = $matches[1]
        }
        
        # Call our debugging function
        Debug-StaffingTableBinding -MMO $mmo -MachineID $machineID -ClassCode $classCode -MachineAcronym $machineAcronym
        
        # Show MessageBox with results
        [System.Windows.Forms.MessageBox]::Show(
            "Debugging completed. Check the console output for details.",
            "Debug Complete",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    })
    

    
    # Add the button to a panel for proper layout
    $staffingPanel = New-Object System.Windows.Forms.Panel
    $staffingPanel.Dock = "Fill"
    $staffingPanel.Controls.Add($btnDebugBinding)

    # Add the grid to the panel but position it below the button
    $staffingTableGrid.Location = New-Object System.Drawing.Point(0, 50)
    # Use fixed initial size that will be updated when the form is shown
    $staffingTableGrid.Width = 600
    $staffingTableGrid.Height = 400
    $staffingTableGrid.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor
                            [System.Windows.Forms.AnchorStyles]::Bottom -bor
                            [System.Windows.Forms.AnchorStyles]::Left -bor
                            [System.Windows.Forms.AnchorStyles]::Right

    # Add controls to the panel
    $staffingPanel.Controls.Add($btnEditStaffing)
    $staffingPanel.Controls.Add($btnDebugBinding)
    $staffingPanel.Controls.Add($btnDebugStaffing)
    $staffingPanel.Controls.Add($btnOpenPDF)    
    $staffingPanel.Controls.Add($staffingTableGrid)

    # Add the panel to the tab
    $staffingTab.Controls.Add($staffingPanel)

    # Initialize grid size
    $detailForm.Add_Shown({
        # Initial size adjustment for the grid
        $staffingTableGrid.Width = $staffingPanel.ClientSize.Width
        $staffingTableGrid.Height = $staffingPanel.ClientSize.Height - 50
    })

    # Add a resize handler to the panel
    $staffingPanel.Add_Resize({
        # Update grid size when panel resizes
        $staffingTableGrid.Width = $staffingPanel.ClientSize.Width
        $staffingTableGrid.Height = $staffingPanel.ClientSize.Height - 50
    })

    # Add Resize event handler to panel
    $staffingPanel.Add_Resize({
        # Update grid size when panel resizes
        $gridWidth = $staffingPanel.ClientSize.Width
        $gridHeight = $staffingPanel.ClientSize.Height - 50  # Leave space for button
        $staffingTableGrid.Size = New-Object System.Drawing.Size($gridWidth, $gridHeight)
    })

    # Add button click handler
    $btnDebugStaffing.Add_Click({
        $staffingFilePath = Join-Path $mmoDirectory "$mmo-Staffing-Table.csv"
        Debug-StaffingTable -FilePath $staffingFilePath -MachineID $machineID
        
        # Also try the Load-StaffingTable function directly
        $result = Load-StaffingTable -MMO $mmo -MachineID $machineID -ClassCode $classCode
        
        if ($result -ne $null) {
            Write-Host "Load-StaffingTable returned $($result.Rows.Count) rows"
            
            # Create a new DataTable for display
            $displayTable = New-Object System.Data.DataTable
            
            # Copy schema
            foreach ($column in $result.Columns) {
                $displayTable.Columns.Add($column.ColumnName, [string]) | Out-Null
            }
            
            # Copy data
            foreach ($row in $result.Rows) {
                $newRow = $displayTable.NewRow()
                foreach ($column in $result.Columns) {
                    $newRow[$column.ColumnName] = $row[$column.ColumnName]
                }
                $displayTable.Rows.Add($newRow)
            }
            
            # Directly set the DataSource
            $staffingTableGrid.DataSource = $displayTable
            
            # Force refresh
            $staffingTableGrid.Refresh()
            $staffingTableGrid.AutoResizeColumns()
        }
        else {
            Write-Host "Load-StaffingTable returned null"
        }
    })

    # Add the debug button to the panel
    $staffingPanel.Controls.Add($btnDebugStaffing)

    # Add PDF Button
    $staffingPanel.Controls.Add($btnOpenPDF)

    # Add the panel to the tab
    $staffingTab.Controls.Add($staffingPanel)
    
    # Add tabs to the tab control
    $tabControl.TabPages.AddRange(@($laborTab, $staffingTab, $calculatedStaffingTab))
 
    # Find and load CSV files
    try {
        # Redirect console output to a null writer to prevent it showing in UI
        $originalConsoleOut = [System.Console]::Out
        $null = [System.IO.StreamWriter]::Null
        [System.Console]::SetOut([System.IO.StreamWriter]::Null)
        
        $formattedAcronym = $acronym -replace "-", "_"
        $laborLookupFiles = Get-ChildItem $mmoDirectory -Filter "*-Labor-Lookup.csv"
        $staffingTableFiles = Get-ChildItem $mmoDirectory -Filter "*-Staffing-Table.csv"
    
        # Process staffing table data first
        try {
            # Use our function to load the staffing table
            $staffingTable = Load-StaffingTable -MMO $mmo -MachineID $machineID -ClassCode $classCode
            
            if ($staffingTable -ne $null -and $staffingTable.Rows.Count -gt 0) {
                $staffingTableGrid.DataSource = $staffingTable
                $staffingTableGrid.AutoResizeColumns()
                Write-DebugLog "Loaded staffing table data with $($staffingTable.Rows.Count) rows" -Category "ViewDetails"
            }
            else {
                Write-DebugLog "No staffing table data found for this machine" -Category "ViewDetails"
            }
        }
        catch {
            Write-DebugLog "Error loading staffing table: $_" -Category "ViewDetails"
        }
    
        # Process labor lookup data
        if ($laborLookupFiles.Count -gt 0) {
            # Read the CSV into a DataTable
            try {
                # Load the lookup table and process it
                $lookupPath = $laborLookupFiles[0].FullName
                $lookupData = Import-Csv $lookupPath
                
                # Convert to DataTable format
                $laborTable = New-Object System.Data.DataTable
                if ($lookupData.Count -gt 0) {
                    $propertyNames = $lookupData[0].PSObject.Properties.Name
                    
                    foreach ($prop in $propertyNames) {
                        $null = $laborTable.Columns.Add($prop)
                    }
                    
                    foreach ($item in $lookupData) {
                        $row = $laborTable.NewRow()
                        foreach ($prop in $propertyNames) {
                            $row[$prop] = $item.$prop
                        }
                        $laborTable.Rows.Add($row)
                    }
                    
                    # Extract machine metrics from ListView
                    $machineMetrics = @{
                        MachineType = $acronym
                        MachineNumber = $machineNumber
                    }
                    
                    # Add all values from ListView based on column mapping
                    foreach ($index in $columnMap.Keys) {
                        if ($selectedItem.SubItems.Count -gt $index -and 
                            -not [string]::IsNullOrWhiteSpace($selectedItem.SubItems[$index].Text)) {
                            $paramName = $columnMap[$index]
                            $machineMetrics[$paramName] = $selectedItem.SubItems[$index].Text
                            Write-DebugLog "Added metric from ListView: $paramName = $($selectedItem.SubItems[$index].Text)" -Category "ViewDetails"
                        }
                    }
                    
                    # Bind data to grid
                    $laborLookupGrid.DataSource = $null
                    $laborLookupGrid.DataSource = $laborTable
                    
                    # Auto-resize columns for better display
                    $laborLookupGrid.AutoResizeColumns()
                    
                    # After grid is populated and visible, highlight cells based on ListView values
                    $detailForm.Add_Shown({
                        Highlight-MatchedCells -DataGridView $laborLookupGrid -MachineMetrics $machineMetrics -ColumnMapping $columnMap
                    })
                }
            }
            catch {
                Write-DebugLog "Error processing labor lookup CSV: $_" -Category "ViewDetails"
                Write-DebugLog "Stack Trace: $($_.ScriptStackTrace)" -Category "ViewDetails"
            }
        }
        
        # Restore original console output
        [System.Console]::SetOut($originalConsoleOut)
    
        if ($detailForm -ne $null) {
            $detailForm.ShowDialog()
        }
    }
    catch {
        # Restore original console output if exception occurs
        if ($originalConsoleOut -ne $null) {
            [System.Console]::SetOut($originalConsoleOut)
        }
        
        Write-DebugLog "Error occurred: $_" -Category "ViewDetails"
        Write-DebugLog "Stack Trace: $($_.ScriptStackTrace)" -Category "ViewDetails"
        [System.Windows.Forms.MessageBox]::Show(
            "Error loading data: $_",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
    finally {
        # Always restore console output
        if ($originalConsoleOut -ne $null) {
            [System.Console]::SetOut($originalConsoleOut)
        }
    }
})

# Generate Report Button Click Handler
$btnGenerateReport.Add_Click({
    # Check if there are machines in the list
    if ($listView.Items.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "No machines found. Please add machines before generating a report.",
            "No Machines",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }
    
    # Create progress form
    $progressForm = New-Object System.Windows.Forms.Form
    $progressForm.Text = "Generating Report..."
    $progressForm.Size = New-Object System.Drawing.Size(400, 100)
    $progressForm.StartPosition = "CenterScreen"
    $progressForm.FormBorderStyle = "FixedDialog"
    $progressForm.ControlBox = $false

    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
    $progressBar.Maximum = 100
    $progressBar.Value = 0
    $progressBar.Size = New-Object System.Drawing.Size(360, 20)
    $progressBar.Location = New-Object System.Drawing.Point(10, 10)

    $progressLabel = New-Object System.Windows.Forms.Label
    $progressLabel.Size = New-Object System.Drawing.Size(360, 20)
    $progressLabel.Location = New-Object System.Drawing.Point(10, 35)
    $progressLabel.Text = "Verifying machine data..."

    $progressForm.Controls.Add($progressBar)
    $progressForm.Controls.Add($progressLabel)
    $progressForm.Show()
    [System.Windows.Forms.Application]::DoEvents()
    
    try {
        # Verify all machines have staffing and calculated staffing tables
        $progressBar.Value = 10
        $progressLabel.Text = "Verifying machine staffing tables..."
        [System.Windows.Forms.Application]::DoEvents()
        
        $missingStaffingTables = @()
        $missingCalculatedTables = @()
        $baseDir = "C:\Users\JR\Documents\Work\Programming Projects\LaborChecklist\Machine Labor Rubrics"
        
        # Check alternate path if first one doesn't exist
        if (-not (Test-Path $baseDir)) {
            $baseDir = "C:\Users\JR\Documents\Work\Programming Projects\LaborChecklist\Machine Labor Rubrics"
        }
        
        # Create the Completed audits directory if it doesn't exist
        $auditDir = Join-Path -Path (Split-Path $baseDir -Parent) -ChildPath "Completed audits"
        if (-not (Test-Path $auditDir)) {
            New-Item -Path $auditDir -ItemType Directory -Force | Out-Null
            Write-DebugLog "Created Completed audits directory: $auditDir" -Category "Report"
        }
        
        $totalMachines = $listView.Items.Count
        
        # Define all ListView columns
        $listViewColumns = @{
            0 = "Acronym"
            1 = "Number"
            2 = "Class Code"
            3 = "MMO"
            4 = "Days/Week"
            5 = "Tours/Day"
            6 = "Stackers"
            7 = "Inductions"
            8 = "Transports"
            9 = "LIM Modules"
            10 = "Machine Type"
            11 = "Site"
            12 = "PSM #"
            13 = "Terminal Type" 
            14 = "Equipment Code"
            15 = "Machines"
        }
        
        # Define key management columns that are required for the summary tab
        $managementColumns = @{
            "MM7" = $true
            "MPE9" = $true
            "ET10" = $true
            "Total (hrs/yr)" = $true
            "Operational Maintenance (hrs/yr)" = $true
        }
        
        # Define mapping between ListView columns and staffing/calculated table columns
        $columnMapping = @{
            "Days/Week" = "Operation (days/wk)"
            "Tours/Day" = "Tours/Day"
            "Stackers" = "Stackers" 
            "Inductions" = "Inductions"
            "Transports" = "Transports"
            "LIM Modules" = "LIM Modules"
            "Machine Type" = "Machine Type"
            "Site" = "Site"
            "PSM #" = "PSM #"
            "Terminal Type" = "Terminal Type"
            "Equipment Code" = "Equipment Code"
            "Machines" = "Machines"
        }
        
        # Track all unique columns discovered in staffing/calculated tables
        $allStaffingColumns = @{"Machine ID" = $true}  # Always include Machine ID
        $allCalculatedColumns = @{"Machine ID" = $true}  # Always include Machine ID
        
        # Track directories with staffing files
        $directoriesWithFiles = @{}

        # Verify each machine has staffing and calculated tables
        for ($i = 0; $i -lt $totalMachines; $i++) {
            $item = $listView.Items[$i]
            $machineAcronym = $item.SubItems[0].Text
            $machineNumber = $item.SubItems[1].Text
            $classCode = $item.SubItems[2].Text
            $mmo = $item.SubItems[3].Text
            $machineID = "$machineAcronym $machineNumber"
            
            $progress = 10 + [int](20 * ($i / $totalMachines))
            $progressBar.Value = $progress
            $progressLabel.Text = "Verifying machine $machineID ($($i+1) of $totalMachines)..."
            [System.Windows.Forms.Application]::DoEvents()
            
            # Check if staffing file exists directly
            $staffingFileExists = $false
            $calculatedFileExists = $false
            
            # First, find the appropriate directory
            $matchingDir = $null
            $staffingFilePath = $null
            $calculatedFilePath = $null
            $searchPatterns = @(
                "*$mmo*$machineAcronym*-$classCode*",
                "*$mmo*-$classCode*",
                "*$mmo*$machineAcronym*",
                "*$mmo*"
            )
            
            foreach ($pattern in $searchPatterns) {
                $dirs = Get-ChildItem $baseDir -Directory | Where-Object { $_.Name -like $pattern }
                if ($dirs -and $dirs.Count -gt 0) {
                    $matchingDir = $dirs[0].FullName
                    break
                }
            }
            
            if ($matchingDir) {
                # Check for staffing file
                $staffingFilePatterns = @(
                    "$mmo-$classCode-Staffing-Table.csv",
                    "$machineAcronym-$classCode-Staffing-Table.csv",
                    "$mmo-Staffing-Table.csv",
                    "*-Staffing-Table.csv"
                )
                
                foreach ($pattern in $staffingFilePatterns) {
                    $files = Get-ChildItem -Path $matchingDir -Filter $pattern -ErrorAction SilentlyContinue
                    if ($files -and $files.Count -gt 0) {
                        $staffingFilePath = $files[0].FullName
                        
                        # Read the file to see if it contains data for this machine
                        try {
                            $content = Get-Content -Path $staffingFilePath -ErrorAction Stop
                            
                            # Properly parse the CSV to check for machine ID
                            if ($content.Count -gt 0) {
                                # Get the header row to find Machine ID column
                                $headerLine = $content[0]
                                $headers = $headerLine.Split(',') | ForEach-Object { $_.Trim() }
                                
                                # Find Machine ID column index
                                $machineIdIndex = -1
                                for ($h = 0; $h -lt $headers.Count; $h++) {
                                    if ($headers[$h] -eq "Machine ID") {
                                        $machineIdIndex = $h
                                        break
                                    }
                                }
                                
                                # Check data rows for machine ID
                                if ($machineIdIndex -ge 0) {
                                    for ($r = 1; $r -lt $content.Count; $r++) {
                                        $line = $content[$r].Trim()
                                        if ([string]::IsNullOrWhiteSpace($line)) { continue }
                                        
                                        $values = $line.Split(',') | ForEach-Object { $_.Trim() }
                                        if ($values.Count -gt $machineIdIndex) {
                                            $rowMachineId = $values[$machineIdIndex]
                                            if ($rowMachineId -eq $machineID) {
                                                $staffingFileExists = $true
                                                
                                                # Add directory to tracking
                                                if (-not $directoriesWithFiles.ContainsKey($matchingDir)) {
                                                    $directoriesWithFiles[$matchingDir] = @{
                                                        "Directory" = $matchingDir
                                                        "StaffingFile" = $staffingFilePath
                                                        "CalculatedFile" = $null
                                                        "MachineIDs" = @($machineID)
                                                    }
                                                } elseif (-not $directoriesWithFiles[$matchingDir].MachineIDs.Contains($machineID)) {
                                                    $directoriesWithFiles[$matchingDir].MachineIDs += $machineID
                                                    $directoriesWithFiles[$matchingDir].StaffingFile = $staffingFilePath
                                                }
                                                
                                                # Collect column names for report
                                                foreach ($header in $headers) {
                                                    $allStaffingColumns[$header] = $true
                                                }
                                                
                                                break
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        catch {
                            Write-DebugLog "Error reading staffing file: $_" -Category "Report"
                        }
                        
                        if ($staffingFileExists) { break }
                    }
                }
                
                # Check for calculated staffing file
                $calculatedFilePatterns = @(
                    "$mmo-$classCode-Calculated-Staffing-Table.csv",
                    "$machineAcronym-$classCode-Calculated-Staffing-Table.csv",
                    "$mmo-Calculated-Staffing-Table.csv",
                    "*-Calculated-Staffing-Table.csv"
                )
                
                foreach ($pattern in $calculatedFilePatterns) {
                    $files = Get-ChildItem -Path $matchingDir -Filter $pattern -ErrorAction SilentlyContinue
                    if ($files -and $files.Count -gt 0) {
                        $calculatedFilePath = $files[0].FullName
                        
                        # Read the file to see if it contains data for this machine
                        try {
                            $content = Get-Content -Path $calculatedFilePath -ErrorAction Stop
                            
                            # Properly parse the CSV to check for machine ID
                            if ($content.Count -gt 0) {
                                # Get the header row to find Machine ID column
                                $headerLine = $content[0]
                                $headers = $headerLine.Split(',') | ForEach-Object { $_.Trim() }
                                
                                # Find Machine ID column index
                                $machineIdIndex = -1
                                for ($h = 0; $h -lt $headers.Count; $h++) {
                                    if ($headers[$h] -eq "Machine ID") {
                                        $machineIdIndex = $h
                                        break
                                    }
                                }
                                
                                # Check data rows for machine ID
                                if ($machineIdIndex -ge 0) {
                                    for ($r = 1; $r -lt $content.Count; $r++) {
                                        $line = $content[$r].Trim()
                                        if ([string]::IsNullOrWhiteSpace($line)) { continue }
                                        
                                        $values = $line.Split(',') | ForEach-Object { $_.Trim() }
                                        if ($values.Count -gt $machineIdIndex) {
                                            $rowMachineId = $values[$machineIdIndex]
                                            if ($rowMachineId -eq $machineID) {
                                                $calculatedFileExists = $true
                                                
                                                # Update directory tracking
                                                if ($directoriesWithFiles.ContainsKey($matchingDir)) {
                                                    $directoriesWithFiles[$matchingDir].CalculatedFile = $calculatedFilePath
                                                } else {
                                                    $directoriesWithFiles[$matchingDir] = @{
                                                        "Directory" = $matchingDir
                                                        "StaffingFile" = $null
                                                        "CalculatedFile" = $calculatedFilePath
                                                        "MachineIDs" = @($machineID)
                                                    }
                                                }
                                                
                                                # Collect column names for report
                                                foreach ($header in $headers) {
                                                    $allCalculatedColumns[$header] = $true
                                                }
                                                
                                                break
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        catch {
                            Write-DebugLog "Error reading calculated staffing file: $_" -Category "Report"
                        }
                        
                        if ($calculatedFileExists) { break }
                    }
                }
            }
            
            # Add to missing lists if needed
            if (-not $staffingFileExists) {
                $missingStaffingTables += $machineID
            }
            
            if (-not $calculatedFileExists) {
                $missingCalculatedTables += $machineID
            }
        }
        
        # If any tables are missing, show warning but proceed with available data
        $includedMachines = @()
        
        if ($missingStaffingTables.Count -gt 0 -or $missingCalculatedTables.Count -gt 0) {
            $excludedMachines = ($missingStaffingTables + $missingCalculatedTables | Sort-Object -Unique)
            $totalMachineCount = $listView.Items.Count
            $excludedCount = $excludedMachines.Count
            $includedCount = $totalMachineCount - $excludedCount
            
            foreach ($item in $listView.Items) {
                $machineAcronym = $item.SubItems[0].Text
                $machineNumber = $item.SubItems[1].Text
                $machineID = "$machineAcronym $machineNumber"
                
                if (-not $excludedMachines.Contains($machineID)) {
                    $includedMachines += $machineID
                }
            }
            
            # Create a shorter message that won't overflow the screen
            $message = "WARNING: Some machines are missing required data and will be excluded from the report.`n`n"
            $message += "Machines included: $includedCount`n"
            $message += "Machines excluded: $excludedCount`n`n"
            
            # Instead of showing all excluded machines, just show a summary count by type
            if ($missingStaffingTables.Count -gt 0) {
                $message += "Missing staffing tables: $($missingStaffingTables.Count) machines`n"
            }
            
            if ($missingCalculatedTables.Count -gt 0) {
                $message += "Missing calculated staffing tables: $($missingCalculatedTables.Count) machines`n"
            }
            
            $message += "`nDo you want to proceed with generating the report for the included machines?"
            
            $result = [System.Windows.Forms.MessageBox]::Show(
                $message,
                "Incomplete Data - Proceed Anyway?",
                [System.Windows.Forms.MessageBoxButtons]::OKCancel,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            
            if ($result -eq [System.Windows.Forms.DialogResult]::Cancel) {
                $progressForm.Close()
                return
            }
            
            # Continue with available data
            $totalMachines = $includedMachines.Count
        } else {
            # All machines have complete data
            foreach ($item in $listView.Items) {
                $machineAcronym = $item.SubItems[0].Text
                $machineNumber = $item.SubItems[1].Text
                $machineID = "$machineAcronym $machineNumber"
                $includedMachines += $machineID
            }
        }
        
        # All verifications passed, prepare for report generation
        $progressBar.Value = 30
        $progressLabel.Text = "Creating data for HTML report..."
        [System.Windows.Forms.Application]::DoEvents()
        
        # Extract all unique columns as ordered lists
        $staffingColumnsList = @("Machine ID") + ($allStaffingColumns.Keys | Where-Object { $_ -ne "Machine ID" } | Sort-Object)
        $calculatedColumnsList = @("Machine ID") + ($allCalculatedColumns.Keys | Where-Object { $_ -ne "Machine ID" } | Sort-Object)
        $managementColumnsList = @("Machine ID") + ($managementColumns.Keys | Where-Object { $_ -ne "Machine ID" -and $managementColumns[$_] -eq $true } | Sort-Object)
        
        # Prepare data for summary
        $summaryData = @{}
        
        # Prepare the HTML report content
        $progressBar.Value = 50
        $progressLabel.Text = "Generating HTML report..."
        [System.Windows.Forms.Application]::DoEvents()
        
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $htmlReport = @"
<!DOCTYPE html>
<html>
<head>
    <title>Machine Staffing Report</title>
    <style>
        body { font-family: Consolas, monospace; margin: 20px; }
        h1, h2, h3, h4 { color: #333; }
        table { border-collapse: collapse; margin: 20px 0; width: 100%; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f0f0f0; }
        .numeric { text-align: right; }
        .positive { color: green; font-weight: bold; }
        .negative { color: red; font-weight: bold; }
        .footer { margin-top: 30px; color: #666; }
        .total-row { font-weight: bold; background-color: #f8f8f8; }
        
        /* Make the summary table more readable with alternating colors */
        .summary-table tr:nth-child(even) {background-color: #f9f9f9;}
        .summary-table th {
            position: sticky;
            top: 0;
            background-color: #e0e0e0;
            z-index: 10;
        }
        
        /* Group headers for the summary table */
        .group-header {
            text-align: center;
            border-bottom: none;
            background-color: #e8e8e8;
        }
        
        /* Add some space between column groups */
        .column-spacer {
            border-right: 2px solid #aaa;
        }
        
        /* Warning box for excluded machines */
        .warning-box {
            background-color: #fff8dc;
            border: 1px solid #e6d68d;
            padding: 15px;
            margin: 20px 0;
            border-radius: 5px;
        }
        .warning-box h3 {
            color: #856404;
            margin-top: 0;
        }
    </style>
</head>
<body>
    <h1>MACHINE STAFFING COMPARISON REPORT</h1>
    <p>Generated: $timestamp</p>
    <p>Number of machines included in report: $totalMachines</p>
"@

        # Add warning section if machines were excluded
        if ($missingStaffingTables.Count -gt 0 -or $missingCalculatedTables.Count -gt 0) {
            $htmlReport += @"
    <div class="warning-box">
        <h3>Some machines excluded due to missing data</h3>
        <p>The following machines were excluded from the report because they are missing required data:</p>
        <ul>
"@
            if ($missingStaffingTables.Count -gt 0) {
                $htmlReport += "<li><strong>Missing staffing tables:</strong> " + ($missingStaffingTables -join ", ") + "</li>"
            }
            
            if ($missingCalculatedTables.Count -gt 0) {
                $htmlReport += "<li><strong>Missing calculated staffing tables:</strong> " + ($missingCalculatedTables -join ", ") + "</li>"
            }
            
            $htmlReport += @"
        </ul>
    </div>
"@
        }

        # Machine Data Table - only showing included machines
        $htmlReport += "<h2>Machine Data</h2>"
        $htmlReport += "<table><tr>"
        
        # Use the original column order (left to right) as in the ListView
        for ($i = 0; $i -lt $listViewColumns.Count; $i++) {
            $col = $listViewColumns[$i]
            $htmlReport += "<th>$col</th>"
        }
        $htmlReport += "</tr>"
        
        foreach ($item in $listView.Items) {
            $machineAcronym = $item.SubItems[0].Text
            $machineNumber = $item.SubItems[1].Text
            $machineID = "$machineAcronym $machineNumber"
            
            # Only include machines that have both staffing and calculated data
            if ($includedMachines -contains $machineID) {
                $htmlReport += "<tr>"
                for ($i = 0; $i -lt $listViewColumns.Count; $i++) {
                    $cellValue = if ($i -lt $item.SubItems.Count) { $item.SubItems[$i].Text } else { "" }
                    $htmlReport += "<td>$cellValue</td>"
                }
                $htmlReport += "</tr>"
            }
        }
        $htmlReport += "</table>"
        
        # Staffing Tables Comparison by Directory
        $htmlReport += "<h2>Staffing Tables Comparison by MMO</h2>"
        
        $directoryIndex = 0
        foreach ($dirInfo in $directoriesWithFiles.Values) {
            $directoryPath = $dirInfo.Directory
            $directoryName = Split-Path $directoryPath -Leaf
            $staffingFile = $dirInfo.StaffingFile
            $calculatedFile = $dirInfo.CalculatedFile
            $machineIDs = $dirInfo.MachineIDs
            
            # Extract MMO from directory name or site information
            $mmoMatch = $directoryName -match "MMO-\d+-\d+"
            $mmoNumber = if ($mmoMatch) { $matches[0] } else { "Unknown MMO" }
            
            $directoryIndex++
            $progress = 50 + [int](30 * ($directoryIndex / $directoriesWithFiles.Count))
            $progressBar.Value = $progress
            $progressLabel.Text = "Processing MMO $mmoNumber..."
            [System.Windows.Forms.Application]::DoEvents()
            
            if ($staffingFile -and $calculatedFile) {
                $htmlReport += "<h3>Relevant Documentation: $directoryName</h3>"
                $htmlReport += "<p>Machines under this MMO: $($machineIDs.Count)</p>"
                
                # Read the staffing file - preserve column order
                $staffingData = @{}
                $originalColumnOrder = @()
                try {
                    $staffingCSV = Import-Csv -Path $staffingFile
                    $staffingHeaders = $staffingCSV[0].PSObject.Properties.Name
                    
                    # Store original column order
                    $originalColumnOrder = $staffingHeaders
                    
                    foreach ($row in $staffingCSV) {
                        $machineID = $row."Machine ID"
                        if ($machineIDs -contains $machineID) {
                            $staffingData[$machineID] = $row
                        }
                    }
                } catch {
                    Write-DebugLog "Error reading staffing CSV: $_" -Category "Report"
                }
                
                # Read the calculated file
                $calculatedData = @{}
                try {
                    $calculatedCSV = Import-Csv -Path $calculatedFile
                    $calculatedHeaders = $calculatedCSV[0].PSObject.Properties.Name
                    
                    foreach ($row in $calculatedCSV) {
                        $machineID = $row."Machine ID"
                        if ($machineIDs -contains $machineID) {
                            $calculatedData[$machineID] = $row
                        }
                    }
                } catch {
                    Write-DebugLog "Error reading calculated CSV: $_" -Category "Report"
                }
                
                # Get column headers from both files - preserving staffing file order first
                $allColumns = @("Machine ID")
                
                # First add all columns from staffing file in original order
                foreach ($header in $originalColumnOrder) {
                    if ($header -ne "Machine ID" -and -not $allColumns.Contains($header)) {
                        $allColumns += $header
                    }
                }
                
                # Then add any additional columns from calculated file
                foreach ($header in $calculatedHeaders) {
                    if ($header -ne "Machine ID" -and -not $allColumns.Contains($header)) {
                        $allColumns += $header
                    }
                }
                
                # Create a comparison table for each machine in this directory
                foreach ($machineID in $machineIDs) {
                    # Skip machines that aren't in the includedMachines list
                    if (-not $includedMachines.Contains($machineID)) {
                        continue
                    }
                    
                    $htmlReport += "<h4>Machine: $machineID</h4>"
                    $htmlReport += "<table>"
                    $htmlReport += "<tr><th>Column Name</th><th>Staffing Value</th><th>Calculated Value</th><th>Difference (S-C)</th></tr>"
                    
                    $staffingTotal = 0
                    $calculatedTotal = 0
                    
                    foreach ($column in $allColumns) {
                        if ($column -eq "Machine ID") { continue }
                        
                        $staffingValue = if ($staffingData.ContainsKey($machineID) -and 
                                            $staffingData[$machineID].PSObject.Properties.Name -contains $column) {
                            $staffingData[$machineID].$column
                        } else { "" }
                        
                        $calculatedValue = if ($calculatedData.ContainsKey($machineID) -and 
                                              $calculatedData[$machineID].PSObject.Properties.Name -contains $column) {
                            $calculatedData[$machineID].$column
                        } else { "" }
                        
                        $difference = ""
                        $diffClass = ""
                        
                        if (![string]::IsNullOrWhiteSpace($staffingValue) -and ![string]::IsNullOrWhiteSpace($calculatedValue)) {
                            $staffingNum = 0
                            $calculatedNum = 0
                            
                            if ([double]::TryParse($staffingValue, [ref]$staffingNum) -and 
                                [double]::TryParse($calculatedValue, [ref]$calculatedNum)) {
                                $diff = $staffingNum - $calculatedNum
                                $difference = $diff.ToString("0.00")
                                
                                # Track totals for summary if this is one of the columns we track
                                if ($column -eq "Total (hrs/yr)" -or $column -eq "MM7" -or $column -eq "MPE9" -or $column -eq "ET10") {
                                    if (-not $summaryData.ContainsKey($machineID)) {
                                        $summaryData[$machineID] = @{
                                            "StaffingTotal" = 0
                                            "CalculatedTotal" = 0
                                            "Difference" = 0
                                            "StaffingMM7" = 0
                                            "CalculatedMM7" = 0
                                            "DifferenceMM7" = 0
                                            "StaffingMPE9" = 0
                                            "CalculatedMPE9" = 0
                                            "DifferenceMPE9" = 0
                                            "StaffingET10" = 0
                                            "CalculatedET10" = 0
                                            "DifferenceET10" = 0
                                        }
                                    }
                                    
                                    if ($column -eq "Total (hrs/yr)") {
                                        $summaryData[$machineID]["StaffingTotal"] = $staffingNum
                                        $summaryData[$machineID]["CalculatedTotal"] = $calculatedNum
                                        $summaryData[$machineID]["Difference"] = $diff
                                    } elseif ($column -eq "MM7") {
                                        $summaryData[$machineID]["StaffingMM7"] = $staffingNum
                                        $summaryData[$machineID]["CalculatedMM7"] = $calculatedNum
                                        $summaryData[$machineID]["DifferenceMM7"] = $diff
                                    } elseif ($column -eq "MPE9") {
                                        $summaryData[$machineID]["StaffingMPE9"] = $staffingNum
                                        $summaryData[$machineID]["CalculatedMPE9"] = $calculatedNum
                                        $summaryData[$machineID]["DifferenceMPE9"] = $diff
                                    } elseif ($column -eq "ET10") {
                                        $summaryData[$machineID]["StaffingET10"] = $staffingNum
                                        $summaryData[$machineID]["CalculatedET10"] = $calculatedNum
                                        $summaryData[$machineID]["DifferenceET10"] = $diff
                                    }
                                }
                                
                                # Add styling class based on value
                                if ($diff -gt 0) {
                                    $difference = "+$difference"
                                    $diffClass = "positive"
                                } elseif ($diff -lt 0) {
                                    $diffClass = "negative"
                                }
                            }
                        }
                        
                        $htmlReport += "<tr>"
                        $htmlReport += "<td>$column</td>"
                        $htmlReport += "<td class='numeric'>$staffingValue</td>"
                        $htmlReport += "<td class='numeric'>$calculatedValue</td>"
                        $htmlReport += "<td class='numeric $diffClass'>$difference</td>"
                        $htmlReport += "</tr>"
                    }
                    
                    $htmlReport += "</table>"
                }
            }
        }
        
        # Build the summary table with grouped headers
        $htmlReport += "<h2>Summary Table</h2>"
        $htmlReport += "<table class='summary-table'>"
        
# Create grouped header row
        $htmlReport += "<tr>"
        $htmlReport += "<th rowspan='2'>Machine ID</th>"
        $htmlReport += "<th colspan='3' class='group-header'>MM7</th>"
        $htmlReport += "<th colspan='3' class='group-header'>MPE9</th>"
        $htmlReport += "<th colspan='3' class='group-header'>ET10</th>"
        $htmlReport += "<th colspan='3' class='group-header'>Total Hours</th>"
        $htmlReport += "</tr>"
        
        $htmlReport += "<tr>"
        # MM7 column headers
        $htmlReport += "<th>Staffing</th><th>Calculated</th><th class='column-spacer'>Diff</th>"
        # MPE9 column headers
        $htmlReport += "<th>Staffing</th><th>Calculated</th><th class='column-spacer'>Diff</th>"
        # ET10 column headers
        $htmlReport += "<th>Staffing</th><th>Calculated</th><th class='column-spacer'>Diff</th>"
        # Total column headers
        $htmlReport += "<th>Staffing</th><th>Calculated</th><th>Diff</th>"
        $htmlReport += "</tr>"
        
        $totalStaffingMM7 = 0
        $totalCalculatedMM7 = 0
        $totalDifferenceMM7 = 0
        
        $totalStaffingMPE9 = 0
        $totalCalculatedMPE9 = 0
        $totalDifferenceMPE9 = 0
        
        $totalStaffingET10 = 0
        $totalCalculatedET10 = 0
        $totalDifferenceET10 = 0
        
        $totalStaffing = 0
        $totalCalculated = 0
        $totalDifference = 0
        
        foreach ($machineID in $summaryData.Keys | Sort-Object) {
            $data = $summaryData[$machineID]
            
            # MM7 data
            $staffingMM7 = $data.StaffingMM7
            $calculatedMM7 = $data.CalculatedMM7
            $differenceMM7 = $data.DifferenceMM7
            
            $totalStaffingMM7 += $staffingMM7
            $totalCalculatedMM7 += $calculatedMM7
            $totalDifferenceMM7 += $differenceMM7
            
            $diffClassMM7 = if ($differenceMM7 -gt 0) { "positive" } elseif ($differenceMM7 -lt 0) { "negative" } else { "" }
            $diffValueMM7 = if ($differenceMM7 -gt 0) { "+$($differenceMM7.ToString('0.00'))" } else { $differenceMM7.ToString("0.00") }
            
            # MPE9 data
            $staffingMPE9 = $data.StaffingMPE9
            $calculatedMPE9 = $data.CalculatedMPE9
            $differenceMPE9 = $data.DifferenceMPE9
            
            $totalStaffingMPE9 += $staffingMPE9
            $totalCalculatedMPE9 += $calculatedMPE9
            $totalDifferenceMPE9 += $differenceMPE9
            
            $diffClassMPE9 = if ($differenceMPE9 -gt 0) { "positive" } elseif ($differenceMPE9 -lt 0) { "negative" } else { "" }
            $diffValueMPE9 = if ($differenceMPE9 -gt 0) { "+$($differenceMPE9.ToString('0.00'))" } else { $differenceMPE9.ToString("0.00") }
            
            # ET10 data
            $staffingET10 = $data.StaffingET10
            $calculatedET10 = $data.CalculatedET10
            $differenceET10 = $data.DifferenceET10
            
            $totalStaffingET10 += $staffingET10
            $totalCalculatedET10 += $calculatedET10
            $totalDifferenceET10 += $differenceET10
            
            $diffClassET10 = if ($differenceET10 -gt 0) { "positive" } elseif ($differenceET10 -lt 0) { "negative" } else { "" }
            $diffValueET10 = if ($differenceET10 -gt 0) { "+$($differenceET10.ToString('0.00'))" } else { $differenceET10.ToString("0.00") }
            
            # Total data
            $staffingTotal = $data.StaffingTotal
            $calculatedTotal = $data.CalculatedTotal
            $difference = $data.Difference
            
            $totalStaffing += $staffingTotal
            $totalCalculated += $calculatedTotal
            $totalDifference += $difference
            
            $diffClass = if ($difference -gt 0) { "positive" } elseif ($difference -lt 0) { "negative" } else { "" }
            $diffValue = if ($difference -gt 0) { "+$($difference.ToString('0.00'))" } else { $difference.ToString("0.00") }
            
            $htmlReport += "<tr>"
            $htmlReport += "<td>$machineID</td>"
            
            # MM7 columns
            $htmlReport += "<td class='numeric'>$($staffingMM7.ToString('0.00'))</td>"
            $htmlReport += "<td class='numeric'>$($calculatedMM7.ToString('0.00'))</td>"
            $htmlReport += "<td class='numeric $diffClassMM7 column-spacer'>$diffValueMM7</td>"
            
            # MPE9 columns
            $htmlReport += "<td class='numeric'>$($staffingMPE9.ToString('0.00'))</td>"
            $htmlReport += "<td class='numeric'>$($calculatedMPE9.ToString('0.00'))</td>"
            $htmlReport += "<td class='numeric $diffClassMPE9 column-spacer'>$diffValueMPE9</td>"
            
            # ET10 columns
            $htmlReport += "<td class='numeric'>$($staffingET10.ToString('0.00'))</td>"
            $htmlReport += "<td class='numeric'>$($calculatedET10.ToString('0.00'))</td>"
            $htmlReport += "<td class='numeric $diffClassET10 column-spacer'>$diffValueET10</td>"
            
            # Total columns
            $htmlReport += "<td class='numeric'>$($staffingTotal.ToString('0.00'))</td>"
            $htmlReport += "<td class='numeric'>$($calculatedTotal.ToString('0.00'))</td>"
            $htmlReport += "<td class='numeric $diffClass'>$diffValue</td>"
            
            $htmlReport += "</tr>"
        }
        
        # Add total row
        $totalDiffClassMM7 = if ($totalDifferenceMM7 -gt 0) { "positive" } elseif ($totalDifferenceMM7 -lt 0) { "negative" } else { "" }
        $totalDiffValueMM7 = if ($totalDifferenceMM7 -gt 0) { "+$($totalDifferenceMM7.ToString('0.00'))" } else { $totalDifferenceMM7.ToString("0.00") }
        
        $totalDiffClassMPE9 = if ($totalDifferenceMPE9 -gt 0) { "positive" } elseif ($totalDifferenceMPE9 -lt 0) { "negative" } else { "" }
        $totalDiffValueMPE9 = if ($totalDifferenceMPE9 -gt 0) { "+$($totalDifferenceMPE9.ToString('0.00'))" } else { $totalDifferenceMPE9.ToString("0.00") }
        
        $totalDiffClassET10 = if ($totalDifferenceET10 -gt 0) { "positive" } elseif ($totalDifferenceET10 -lt 0) { "negative" } else { "" }
        $totalDiffValueET10 = if ($totalDifferenceET10 -gt 0) { "+$($totalDifferenceET10.ToString('0.00'))" } else { $totalDifferenceET10.ToString("0.00") }
        
        $totalDiffClass = if ($totalDifference -gt 0) { "positive" } elseif ($totalDifference -lt 0) { "negative" } else { "" }
        $totalDiffValue = if ($totalDifference -gt 0) { "+$($totalDifference.ToString('0.00'))" } else { $totalDifference.ToString("0.00") }
        
        $htmlReport += "<tr class='total-row'>"
        $htmlReport += "<td>TOTAL</td>"
        
        # MM7 totals
        $htmlReport += "<td class='numeric'>$($totalStaffingMM7.ToString('0.00'))</td>"
        $htmlReport += "<td class='numeric'>$($totalCalculatedMM7.ToString('0.00'))</td>"
        $htmlReport += "<td class='numeric $totalDiffClassMM7 column-spacer'>$totalDiffValueMM7</td>"
        
        # MPE9 totals
        $htmlReport += "<td class='numeric'>$($totalStaffingMPE9.ToString('0.00'))</td>"
        $htmlReport += "<td class='numeric'>$($totalCalculatedMPE9.ToString('0.00'))</td>"
        $htmlReport += "<td class='numeric $totalDiffClassMPE9 column-spacer'>$totalDiffValueMPE9</td>"
        
        # ET10 totals
        $htmlReport += "<td class='numeric'>$($totalStaffingET10.ToString('0.00'))</td>"
        $htmlReport += "<td class='numeric'>$($totalCalculatedET10.ToString('0.00'))</td>"
        $htmlReport += "<td class='numeric $totalDiffClassET10 column-spacer'>$totalDiffValueET10</td>"
        
        # Overall totals
        $htmlReport += "<td class='numeric'>$($totalStaffing.ToString('0.00'))</td>"
        $htmlReport += "<td class='numeric'>$($totalCalculated.ToString('0.00'))</td>"
        $htmlReport += "<td class='numeric $totalDiffClass'>$totalDiffValue</td>"
        
        $htmlReport += "</tr>"
        
        $htmlReport += "</table>"
        
        # Add notes
        $htmlReport += @"
    <div class='footer'>
        <h3>Notes:</h3>
        <ul>
            <li>This report compares management staffing tables with calculated staffing tables</li>
            <li>Each directory containing both types of tables is analyzed separately</li>
            <li>Positive differences (green) indicate a surplus in management hours</li>
            <li>Negative differences (red) indicate a deficit in management hours</li>
        </ul>
    </div>
</body>
</html>
"@
        
    # Save the HTML report
    $progressBar.Value = 90
    $progressLabel.Text = "Saving report..."
    [System.Windows.Forms.Application]::DoEvents()
    
    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $reportPath = Join-Path -Path $auditDir -ChildPath "Machine_Staffing_Report_$timestamp.html"
    
    try {
        $htmlReport | Out-File -FilePath $reportPath -Encoding utf8
        
        $progressBar.Value = 100
        $progressLabel.Text = "Done!"
        [System.Windows.Forms.Application]::DoEvents()

        $progressForm.Close()
                    
        # Show success message with option to open grievance form
        $result = [System.Windows.Forms.MessageBox]::Show(
            "Machine Staffing Report generated successfully!`n`nReport saved to: $reportPath`n`nWould you like to open the Grievance Form to import this report?",
            "Report Generated",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            # Get the grievance form path
            $grievanceFormPath = "C:\Users\JR\Documents\Work\Programming Projects\LaborChecklist\Powershell Scripts\UI\Grievance Form.html"
            
            if (Test-Path $grievanceFormPath) {
                # Open the grievance form in default browser
                Start-Process $grievanceFormPath
            } else {
                [System.Windows.Forms.MessageBox]::Show(
                    "Grievance Form not found at: $grievanceFormPath",
                    "File Not Found",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                )
            }
        }
    }
    catch {
        if ($progressForm -and $progressForm.Visible) {
            $progressForm.Close()
        }
        
        [System.Windows.Forms.MessageBox]::Show(
            "Error saving report: $_",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        
        Write-DebugLog "Error saving report: $_" -Category "Report"
        }
    }
    catch {
        if ($progressForm -and $progressForm.Visible) {
            $progressForm.Close()
        }
        
        [System.Windows.Forms.MessageBox]::Show(
            "Error generating report: $_",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        
        Write-DebugLog "Error generating report: $_" -Category "Report"
        Write-DebugLog "Stack trace: $($_.ScriptStackTrace)" -Category "Report"
    }
})

$btnSave.Add_Click({
    # Input validation
    if ([string]::IsNullOrWhiteSpace($cmbAcronym.Text) -or 
        [string]::IsNullOrWhiteSpace($txtNumber.Text)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Machine acronym and number are required!",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    $selectedAcronym = $cmbAcronym.Text
    $machineNumber = $txtNumber.Text.Trim()

    # Check for duplicates in ListView
    $isDuplicate = $false
    foreach ($item in $listView.Items) {
        if ($item.SubItems[0].Text -eq $selectedAcronym -and 
            $item.SubItems[1].Text -eq $machineNumber) {
            $isDuplicate = $true
            break
        }
    }

    if ($isDuplicate) {
        [System.Windows.Forms.MessageBox]::Show(
            "A machine with acronym '$selectedAcronym' and number '$machineNumber' already exists!",
            "Duplicate Entry",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    # Check if valid machine acronym
    if (-not $machineClassCodes.ContainsKey($selectedAcronym)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Invalid machine acronym!",
            "Validation Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    # Create a basic ListView item with just acronym and number
    $item = New-Object System.Windows.Forms.ListViewItem($selectedAcronym)
    $item.SubItems.Add($machineNumber)
    $item.SubItems.Add("")  # Empty class code
    $item.SubItems.Add("")  # Empty MMO
    
    # Add empty values for all other parameters (12 total additional columns)
    $item.SubItems.Add("")  # Days/Week
    $item.SubItems.Add("")  # Tours/Day
    $item.SubItems.Add("")  # Stackers
    $item.SubItems.Add("")  # Inductions
    $item.SubItems.Add("")  # Transports
    $item.SubItems.Add("")  # LIM Modules
    $item.SubItems.Add("")  # Machine Type
    $item.SubItems.Add("")  # Site
    $item.SubItems.Add("")  # PSM #
    $item.SubItems.Add("")  # Terminal Type
    $item.SubItems.Add("")  # Equipment Code
    $item.SubItems.Add("")  # Machines
    
    $listView.Items.Add($item)
    
    # Select the newly added item
    $item.Selected = $true
    
    # Show a simple message to instruct the user to configure
    [System.Windows.Forms.MessageBox]::Show(
        "Machine entry added. Please use the Configure button to set up the machine parameters.",
        "Machine Added",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )

    # Clear inputs
    $cmbAcronym.SelectedIndex = -1
    $txtNumber.Clear()
})

# Restore Session Button Click Handler
$btnRestoreSession.Add_Click({
    # Show file open dialog
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "CSV files (*.csv)|*.csv"
    $openFileDialog.Title = "Select Previous Session CSV"
    
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        # Create progress form
        $progressForm = New-Object System.Windows.Forms.Form
        $progressForm.Text = "Restoring Session..."
        $progressForm.Size = New-Object System.Drawing.Size(400, 100)
        $progressForm.StartPosition = "CenterScreen"
        $progressForm.FormBorderStyle = "FixedDialog"
        $progressForm.ControlBox = $false

        $progressBar = New-Object System.Windows.Forms.ProgressBar
        $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
        $progressBar.Maximum = 100
        $progressBar.Value = 0
        $progressBar.Size = New-Object System.Drawing.Size(360, 20)
        $progressBar.Location = New-Object System.Drawing.Point(10, 10)

        $progressLabel = New-Object System.Windows.Forms.Label
        $progressLabel.Size = New-Object System.Drawing.Size(360, 20)
        $progressLabel.Location = New-Object System.Drawing.Point(10, 35)
        $progressLabel.Text = "Loading session data..."

        $progressForm.Controls.Add($progressBar)
        $progressForm.Controls.Add($progressLabel)
        $progressForm.Show()
        [System.Windows.Forms.Application]::DoEvents()
        
        try {
            # Clear existing ListView data
            $listView.Items.Clear()
            
            # Import the CSV
            $sessionData = Import-Csv -Path $openFileDialog.FileName
            
            # Determine number of items to process
            $totalItems = $sessionData.Count
            $processedItems = 0
            
            # Process each row
            foreach ($row in $sessionData) {
                # Update progress
                $processedItems++
                $percentComplete = [int](($processedItems / $totalItems) * 100)
                $progressBar.Value = $percentComplete
                $progressLabel.Text = "Restoring machine $processedItems of $totalItems..."
                [System.Windows.Forms.Application]::DoEvents()
                
                # Create ListView item
                $item = New-Object System.Windows.Forms.ListViewItem($row.Acronym)
                
                # Add additional columns
                $item.SubItems.Add($row.Number)
                $item.SubItems.Add($row.ClassCode)
                $item.SubItems.Add($row.MMO)
                $item.SubItems.Add($row.DaysPerWeek)
                $item.SubItems.Add($row.ToursPerDay)
                $item.SubItems.Add($row.Stackers)
                $item.SubItems.Add($row.Inductions)
                $item.SubItems.Add($row.Transports)
                $item.SubItems.Add($row.LIMModules)
                $item.SubItems.Add($row.MachineType)
                $item.SubItems.Add($row.Site)
                $item.SubItems.Add($row.PSMN)
                $item.SubItems.Add($row.TerminalType)
                $item.SubItems.Add($row.EquipmentCode)
                $item.SubItems.Add($row.Machines)
                
                # Create Tag hashtable to store original values and adjustment flags
                $tag = @{}
                
                # Add original values if available
                if ($row.PSObject.Properties.Name -contains "OriginalDaysPerWeek" -and 
                    -not [string]::IsNullOrWhiteSpace($row.OriginalDaysPerWeek)) {
                    $tag["OriginalDaysPerWeek"] = $row.OriginalDaysPerWeek
                }
                
                if ($row.PSObject.Properties.Name -contains "OriginalToursPerDay" -and 
                    -not [string]::IsNullOrWhiteSpace($row.OriginalToursPerDay)) {
                    $tag["OriginalToursPerDay"] = $row.OriginalToursPerDay
                }
                
                # Set adjustment flag
                $isAdjusted = $false
                if ($row.PSObject.Properties.Name -contains "ValuesAdjusted") {
                    if ($row.ValuesAdjusted -eq "True") {
                        $isAdjusted = $true
                    }
                }
                $tag["Adjusted"] = $isAdjusted
                
                # Store standardized values for toggle feature if values were adjusted
                if ($isAdjusted) {
                    $tag["StandardDaysPerWeek"] = $row.DaysPerWeek
                    $tag["StandardToursPerDay"] = $row.ToursPerDay
                    $tag["ShowingStandard"] = $true
                    
                    # Apply color to cells that have been adjusted
                    if ($tag.ContainsKey("OriginalDaysPerWeek") -and 
                        $row.DaysPerWeek -ne $tag["OriginalDaysPerWeek"]) {
                        $item.SubItems[4].ForeColor = [System.Drawing.Color]::DarkOrange
                    }
                    
                    if ($tag.ContainsKey("OriginalToursPerDay") -and 
                        $row.ToursPerDay -ne $tag["OriginalToursPerDay"]) {
                        $item.SubItems[5].ForeColor = [System.Drawing.Color]::DarkOrange
                    }
                }
                
                # Assign tag to item
                $item.Tag = $tag
                
                # Add to ListView
                $listView.Items.Add($item)
            }
            
            $progressForm.Close()
            
            # Auto-resize columns
            $listView.Columns | ForEach-Object { $_.Width = -2 }
            
            [System.Windows.Forms.MessageBox]::Show(
                "Session restored successfully with $totalItems machines!",
                "Session Restored",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        }
        catch {
            $progressForm.Close()
            [System.Windows.Forms.MessageBox]::Show(
                "Failed to restore session: $_",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            Write-DebugLog "Error restoring session: $_" -Category "RestoreSession"
        }
    }
})


################################################################################
#                        FORM DISPLAY LOGIC                                    #
################################################################################

# Add controls to form
$form.Controls.AddRange(@(
    $listView,
    $lblAcronym, $cmbAcronym,
    $lblNumber, $txtNumber,
    $btnSave, $btnImport, $btnConfigure, $btnViewDetails, $btnExport, $btnGenerateReport, $btnRestoreSession
))

$form.ShowDialog() | Out-Null
