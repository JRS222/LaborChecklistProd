<#
.SYNOPSIS
    Splits a TACS "Employee Everything" CSV export into normalized, typed CSV files.

.DESCRIPTION
    The USPS TACS "Employee Everything" report packs three different record types into
    one CSV, and re-uses columns 19+ ("Job", "D/A", "LDC", "Oper/Lu", ...) for different
    datatypes depending on the row:

      * Base rows        (Job = "Base"/"Temp")  - employee assignment / schedule header
      * Summary rows     (Job = day of week, "Paid Hours :", "TACS Hours :")
                          - hours totals per code for that day / pay period week
      * Transaction rows (Job = badge reader number, e.g. "702-01", "000-0000",
                          "888-8888", "999-9999") - clock rings and hours entries

    Some columns even hold two datatypes at once:

      * YrPPWk          = Year + Pay Period + Week  ("2023161" = 2023, PP 16, week 1)
      * Summary "LDC"   = "code: hours" composite   ("04300: 000.05")
      * Transaction "LDC" / "Oper/Lu" = ring date ("14-JUL-23") and ergot decimal
                          time ("22.73" = 22:43:48)
      * Transaction "D/A" = ring type (BT/ET/OL/IL/MV/OT) OR a 3-digit/5-digit
                          hours code ("090", "05699")

    The script detects each row's shape, decomposes every overloaded column into
    separate typed columns, converts ergot times to HH:MM:SS and TACS dates
    (DD-MMM-YY) to ISO dates (yyyy-MM-dd), and writes one CSV per record type:

      <prefix>.Base.csv
      <prefix>.Summary.csv
      <prefix>.Transactions.csv
      <prefix>.Unmatched.csv   (rows that did not fit; normally empty)

    The first 23 columns (RowNumber ... SLDC_Used) are identical across the three
    typed files, so the files can be concatenated or joined later. Every output row
    also carries a Parse_Warning column that is empty unless the row did not match
    the expected pattern for its shape.

.PARAMETER InputPath
    Path to one or more TACS CSV files. Wildcards are allowed.

.PARAMETER OutputDirectory
    Folder for the output files. Default: the input file's folder.

.PARAMETER OutputPrefix
    Prefix for the output file names. Default: the input file's base name.
    When several input files are processed, the base name is always appended.

.PARAMETER SkipBase
    Do not write the Base output file.

.PARAMETER SkipSummary
    Do not write the Summary output file.

.PARAMETER SkipTransactions
    Do not write the Transactions output file.

.EXAMPLE
    PS> .\Split-TacsReport.ps1 -InputPath ".\Combined_TACS_Report.csv"

    Splits Combined_TACS_Report.csv in place into Combined_TACS_Report.Base.csv,
    Combined_TACS_Report.Summary.csv, Combined_TACS_Report.Transactions.csv and
    Combined_TACS_Report.Unmatched.csv.

.EXAMPLE
    PS> .\Split-TacsReport.ps1 "C:\TACS\*.csv" -OutputDirectory "C:\TACS\Normalized"

    Splits every CSV in C:\TACS and writes the normalized files to C:\TACS\Normalized.

.NOTES
    Compatible with Windows PowerShell 5.1 and PowerShell 7.
    Reference tables (D/A hour codes, LDC function codes, RSC letters, designation
    codes) are taken from "Analyzing the Employee Everything Report (for CSVs)".
    Codes that are not in that reference are exported without a Description value.
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string[]] $InputPath,

    [string] $OutputDirectory,

    [string] $OutputPrefix,

    [switch] $SkipBase,
    [switch] $SkipSummary,
    [switch] $SkipTransactions
)

# ---------------------------------------------------------------------------
# Reference dictionaries (from the analysis guide).
# ---------------------------------------------------------------------------

# D/A hour codes: 5-digit form "HHHSS" (3-digit hour code + 2-digit specific code).
$script:DaCode5 = @{
    '01000' = 'Begin tour'
    '01100' = 'Move'
    '01200' = 'Out to lunch'
    '01300' = 'In from lunch'
    '01400' = 'End tour'
    '02400' = 'AWOL'
    '02409' = 'AWOL - Late'
    '03000' = 'Full LWOP hours (generated at mainframe)'
    '03100' = 'Partial LWOP hours (generated at mainframe)'
    '03200' = 'Telephone time'
    '03300' = 'Guarantee telephone time'
    '03400' = 'Beeper time'
    '03500' = 'Extra straight time'
    '03600' = 'Guarantee telephone OT'
    '03800' = 'Extra 50% (adjustments only)'
    '03900' = 'Extra Straight Time (adjustments only)'
    '04300' = 'Penalty Overtime Payment'
    '04400' = 'Military LWOP'
    '04600' = 'Donated leave - personal'
    '04700' = 'Rural free Saturday'
    '04800' = 'Holiday schedule premium'
    '04900' = 'OWCP Regular LWOP hours'
    '04999' = 'OWCP FMLA LWOP hours'
    '05000' = 'Rural carrier trips'
    '05100' = 'Rural carrier actual hours'
    '05200' = 'Work hours'
    '05300' = 'Overtime hours'
    '05400' = 'Night work premium hours'
    '05500' = 'Annual leave - Regular'
    '05501' = 'Annual leave - In lieu of SL'
    '05502' = 'Annual leave - Late'
    '05510' = 'Annual leave - Emergency (aka EAL)'
    '05599' = 'Annual leave - FMLA'
    '05600' = 'Sick leave - Regular'
    '05609' = 'Sick leave - Late'
    '05611' = 'Sick leave - Restricted'
    '05697' = 'Sick leave - Dependent Care'
    '05698' = 'Sick leave - FMLA Dependent Care'
    '05699' = 'Sick leave - FMLA Self'
    '05700' = 'Holiday work'
    '05800' = 'Holiday leave'
    '05900' = 'Part day LWOP'
    '05901' = 'Part day LWOP - In lieu of SL'
    '05902' = 'Part day LWOP - Proffered'
    '05903' = 'Part day LWOP - Personal'
    '05904' = 'Part day LWOP - Other'
    '05905' = 'Part day LWOP - Maternity'
    '05906' = 'Part day LWOP - Suspension'
    '05907' = 'Part day LWOP - Union Official'
    '05908' = 'Part day LWOP - Suspending term'
    '05909' = 'Part day LWOP - Late'
    '05995' = 'Part day LWOP - System-generated'
    '05997' = 'Part day LWOP - FMLA'
    '06000' = 'Full day LWOP'
    '06001' = 'Full day LWOP - In lieu of SL'
    '06002' = 'Full day LWOP - Proffered'
    '06003' = 'Full day LWOP - Personal'
    '06004' = 'Full day LWOP - Other'
    '06005' = 'Full day LWOP - Maternity'
    '06006' = 'Full day LWOP - Suspension'
    '06007' = 'Full day LWOP - Union Official'
    '06008' = 'Full day LWOP - Suspending term'
    '06009' = 'Full day LWOP - Late'
    '06099' = 'Full day LWOP - FMLA'
    '06100' = 'Court Leave'
    '06200' = 'Guarantee time'
    '06300' = 'TE cross-foot (also for code 035)'
    '06500' = 'Meeting time'
    '06600' = 'Convention leave'
    '06700' = 'Military leave'
    '06800' = 'Guarantee overtime'
    '06900' = 'Blood donor leave'
    '07000' = 'Stewards duty time'
    '07100' = 'Continuation of pay leave'
    '07199' = 'Continuation of pay leave - FMLA'
    '07200' = 'Sunday premium'
    '07300' = 'Out of schedule premium'
    '07400' = 'Christmas work'
    '07600' = 'Non-scheduled cross-foot'
    '07800' = 'Act of nature leave'
    '07900' = 'Veterans funeral leave'
    '08000' = 'Relocation leave'
    '08100' = 'Civil disorder leave'
    '08200' = 'Travel within schedule'
    '08300' = 'Travel outside schedule'
    '08400' = 'Union official leave'
    '08500' = 'Voting leave'
    '08600' = 'Other paid leave'
    '08800' = 'Non-bargain reschedule premium'
    '08900' = 'Postmaster Org. leave'
    '09000' = 'Higher level authorization'
    '09100' = 'OT Auth'
    '09101' = 'OT Auth - Before Scheduled Tour'
    '09102' = 'OT Auth - After Scheduled Tour'
    '09103' = 'OT Auth - Full Tour'
    '09104' = 'OT Auth - Before Scheduled Tour, Out of Schedule'
    '09105' = 'OT Auth - After Scheduled Tour, Out of Schedule'
    '09200' = 'Disallow guarantee time'
    '09300' = 'No lunch punch'
}

# 3-digit hour code forms (specific code omitted).
$script:DaCode3 = @{
    '010' = 'Begin tour'
    '011' = 'Move'
    '012' = 'Out to lunch'
    '013' = 'In from lunch'
    '014' = 'End tour'
    '024' = 'AWOL'
    '030' = 'Full LWOP hours'
    '031' = 'Partial LWOP hours'
    '032' = 'Telephone time'
    '033' = 'Guarantee telephone time'
    '034' = 'Beeper time'
    '035' = 'Extra straight time'
    '036' = 'Guarantee telephone OT'
    '038' = 'Extra 50% (adjustments only)'
    '039' = 'Extra Straight Time (adjustments only)'
    '043' = 'Penalty Overtime Payment'
    '044' = 'Military LWOP'
    '046' = 'Donated leave - personal'
    '047' = 'Rural free Saturday'
    '048' = 'Holiday schedule premium'
    '049' = 'OWCP LWOP hours'
    '050' = 'Rural carrier trips'
    '051' = 'Rural carrier actual hours'
    '052' = 'Work hours'
    '053' = 'Overtime hours'
    '054' = 'Night work premium hours'
    '055' = 'Annual leave'
    '056' = 'Sick leave'
    '057' = 'Holiday work'
    '058' = 'Holiday leave'
    '059' = 'Part day LWOP'
    '060' = 'Full day LWOP'
    '061' = 'Court Leave'
    '062' = 'Guarantee time'
    '063' = 'TE cross-foot'
    '065' = 'Meeting time'
    '066' = 'Convention leave'
    '067' = 'Military leave'
    '068' = 'Guarantee overtime'
    '069' = 'Blood donor leave'
    '070' = 'Stewards duty time'
    '071' = 'Continuation of pay leave'
    '072' = 'Sunday premium'
    '073' = 'Out of schedule premium'
    '074' = 'Christmas work'
    '076' = 'Non-scheduled cross-foot'
    '078' = 'Act of nature leave'
    '079' = 'Veterans funeral leave'
    '080' = 'Relocation leave'
    '081' = 'Civil disorder leave'
    '082' = 'Travel within schedule'
    '083' = 'Travel outside schedule'
    '084' = 'Union official leave'
    '085' = 'Voting leave'
    '086' = 'Other paid leave'
    '088' = 'Non-bargain reschedule premium'
    '089' = 'Postmaster Org. leave'
    '09000' = 'Higher level authorization'
    '09100' = 'OT Auth'
    '09200' = 'Disallow guarantee time'
    '09300' = 'No lunch punch'
}

# Clock ring / transaction types (D/A column of Transaction rows).
$script:PunchType = @{
    'BT' = 'Begin tour'
    'ET' = 'End tour'
    'OL' = 'Out to lunch'
    'IL' = 'In from lunch'
    'MV' = 'Move'
    'OT' = 'Overtime entry (system-generated)'
}

# Designation codes (D/A column of Base rows). 2-digit spellings are zero-padded
# to 3 digits before lookup.
$script:Designation = @{
    '053' = 'VMF (Supervisory)'
    '066' = 'Maint/Bldg Svcs (Supervisory)'
    '067' = 'Maint/Bldg & Equip (Supervisory)'
    '068' = 'Maint/Postal Op Equip (Supervisory)'
    '069' = 'Maint/Admin (Supervisory)'
    '080' = 'Postmasters (Supervisory)'
    '089' = 'Postal Inspectors'
    '090' = 'Managers & Supervisors (also Medical Officers & Head Nurses)'
    '119' = 'All Employees / Non-Inspectors'
    '166' = 'Maint/Bldg Svcs (Supervisory)'
    '167' = 'Maint/Bldg & Equip (Supervisory)'
    '168' = 'Operating Equipment (Supervisory)'
    '169' = 'Maint/Admin (Supervisory)'
    '179' = 'Postal Police Officers'
    '999' = '(not in reference table)'
}

# Rate Schedule Code letters (RSC column and the letter part of "X#-##").
$script:RscLetter = @{
    'A' = 'MTEC bargaining employees (CAG A)'
    'C' = 'MESC bargaining employees (CAG M)'
    'E' = 'EAS non-bargaining employees'
    'F' = 'Fourth class postmasters (CAG L)'
    'G' = 'Nurses'
    'J' = 'Management trainees'
    'K' = 'Headquarters & Rockville support services'
    'L' = 'Fourth class leave replacements'
    'M' = 'Mailhandlers'
    'N' = 'ISC craft employees'
    'P' = 'Clerks'
    'Q' = 'City Carrier'
    'S' = 'PCES'
    'T' = 'Tool and die makers'
    'U' = 'Attorneys'
    'Y' = 'Security Force'
}

# Labor Distribution Codes (LDC column of Base rows), 4-digit form.
$script:LdcCode = @{
    '0100' = 'Supervisor - Operations Programs Support'
    '0200' = 'Quality Improvement'
    '0300' = 'Industrial Engineering'
    '0400' = 'Address Management Systems'
    '0500' = 'Production Planning'
    '0600' = 'Vacant'
    '0700' = 'Environmental Management'
    '0800' = 'Administrative and Clerical'
    '0900' = 'Delivery and Retail Programs'
    '1000' = 'Supervisor - Mail Processing'
    '1100' = 'Automated Distribution - Letters & Flats'
    '1200' = 'Mechanized Distribution - Letters & Flats'
    '1300' = 'Mechanized Distribution - Other'
    '1400' = 'Manual Distribution'
    '1500' = 'Remote Bar Code System'
    '1600' = 'Vacant'
    '1700' = 'MP - Other Direct Operations'
    '1800' = 'MP - Indirect/Related'
    '1900' = 'Vacant'
    '2000' = 'Supervisor - Delivery Services'
    '2100' = 'City Delivery - Office Time'
    '2200' = 'City Delivery - Street Time'
    '2300' = 'Other City Delivery'
    '2400' = 'Special Delivery'
    '2500' = 'Rural Delivery'
    '2600' = 'Carrier Customer Support Activities'
    '2700' = 'Collections'
    '2800' = 'City Carrier - Tertiary Distribution'
    '2900' = 'Routers - Office'
    '3000' = 'Supervisor - Vehicle Services'
    '3100' = 'Administrative & Clerical'
    '3200' = 'Vehicle Maintenance'
    '3300' = 'VOMA Support'
    '3400' = 'Vehicle Operations'
    '3500' = 'Supervisor - Plant & Equipment Maintenance'
    '3600' = 'Postal Operating Equipment'
    '3700' = 'Building Systems Equipment'
    '3800' = 'Building Services (Custodial)'
    '3900' = 'Maintenance, Planning, Control and Stores'
    '4000' = 'Supervisor - Customer Services'
    '4100' = 'Unit Distribution - Automated'
    '4200' = 'Unit Distribution - Mechanized'
    '4300' = 'Unit Distribution - Manual'
    '4400' = 'Post Office Box Distribution'
    '4500' = 'Window Services'
    '4600' = 'Vending Equipment'
    '4700' = 'Miscellaneous (CAG H-L only)'
    '4800' = 'Customer Services - Administrative/Miscellaneous'
    '4900' = 'Computerized Forwarding System'
    '5000' = 'Supervisor - Finance & Planning'
    '5100' = 'Vacant'
    '5200' = 'Financial Systems'
    '5300' = 'Vacant'
    '5400' = 'Budget & Financial Analysis'
    '5500' = 'Internal Control Group'
    '5600' = 'Administrative & Clerical'
    '5700' = 'Statistical Programs'
    '5800' = 'TACS'
    '5900' = 'Vacant'
    '6000' = 'Supervisor - Human Resources'
    '6100' = 'Labor Relations Activities'
    '6200' = 'Personnel Services'
    '6300' = 'Safety and Health'
    '6400' = 'EEO'
    '6500' = 'Training Support'
    '6600' = 'Injury Compensation'
    '6700' = 'Reserved'
    '6800' = 'Limited Duty'
    '6900' = 'Rehabilitation'
    '7000' = 'Supervisor - Marketing'
    '7100' = 'Commercial Sales and Account Management'
    '7200' = 'Postal Business Centers'
    '7300' = 'Expedited Mail Service'
    '7400' = 'Retail Marketing'
    '7500' = 'Claims & Inquiries - Dead Mail Branches'
    '7600' = 'Consumer Affairs'
    '7700' = 'Accountable Paper'
    '7800' = 'Marketing - Administrative & Clerical'
    '7900' = 'Mailing Requirements and Business Mail Entry'
    '8000' = 'Postmaster or Installation Manager'
    '8100' = 'Supervisor - Administrative'
    '8200' = 'Administrative & Clerical'
    '8300' = 'Purchasing'
    '8400' = 'Information Systems'
    '8500' = 'Facilities'
    '8600' = 'Vacant'
    '8700' = 'Vacant'
    '8800' = 'Management Projects - Supervisory'
    '8900' = 'Management Projects - Non-supervisory'
    '9000' = 'Training - Operations Programs Support'
    '9100' = 'Training - Operations - Mail Processing'
    '9200' = 'Training - Operations - Delivery Services'
    '9300' = 'Training - Operations - Maintenance'
    '9400' = 'Training - Operations - Customer Services'
    '9500' = 'Training - Finance and Planning'
    '9600' = 'Training - Human Resources'
    '9700' = 'Training - Customer Services Support'
    '9800' = 'Training - Administration (Field)'
    '9900' = 'Default Code'
}

# ---------------------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------------------

# Splits one CSV line into fields (handles quoted fields, embedded commas,
# doubled quotes, and a trailing comma).
function Split-CsvLine {
    param([string] $Line)
    $fields = New-Object System.Collections.Generic.List[string]
    $sb = New-Object System.Text.StringBuilder
    $inQuotes = $false
    $len = $Line.Length
    for ($i = 0; $i -lt $len; $i++) {
        $c = $Line[$i]
        if ($inQuotes) {
            if ($c -eq '"') {
                if (($i + 1) -lt $len -and $Line[$i + 1] -eq '"') {
                    [void] $sb.Append('"'); $i++
                } else {
                    $inQuotes = $false
                }
            } else {
                [void] $sb.Append($c)
            }
        } else {
            if ($c -eq '"') { $inQuotes = $true }
            elseif ($c -eq ',') { $fields.Add($sb.ToString()); [void] $sb.Clear() }
            else { [void] $sb.Append($c) }
        }
    }
    $fields.Add($sb.ToString())
    return , $fields
}

# Converts an ergot decimal-hour value ("22.73", ".5") to HH:MM:SS.
# "0", "0.0", "0.00" and blanks are treated as "no time" and return ''.
function ConvertFrom-Ergot {
    param([string] $Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return '' }
    $d = 0.0
    if (-not [double]::TryParse($Value,
            [System.Globalization.NumberStyles]::Float,
            [System.Globalization.CultureInfo]::InvariantCulture, [ref] $d)) {
        return $Value
    }
    if ($d -eq 0.0) { return '' }
    $totalSec = [math]::Round($d * 3600)
    $hh = [math]::Floor($totalSec / 3600)
    $mm = [math]::Floor(($totalSec % 3600) / 60)
    $ss = $totalSec % 60
    return ('{0:00}:{1:00}:{2:00}' -f $hh, $mm, $ss)
}

# Converts a TACS date "DD-MMM-YY" ("14-JUL-23") to ISO "yyyy-MM-dd".
# Returns '' when the value is not a valid TACS date.
function ConvertFrom-TacsDate {
    param([string] $Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return '' }
    if ($Value -notmatch '^\d{2}-[A-Za-z]{3}-\d{2}$') { return '' }
    $dt = New-Object datetime
    if ([datetime]::TryParseExact($Value, 'dd-MMM-yy',
            [System.Globalization.CultureInfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::None, [ref] $dt)) {
        return $dt.ToString('yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
    }
    return ''
}

# Normalizes a numeric text value ("008.05" -> "8.05", "8" -> "8.00").
# Returns the input unchanged when it is not numeric; '' stays ''.
function ConvertTo-NumericText {
    param([string] $Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return '' }
    $d = 0.0
    if ([double]::TryParse($Value,
            [System.Globalization.NumberStyles]::Float,
            [System.Globalization.CultureInfo]::InvariantCulture, [ref] $d)) {
        return $d.ToString('0.00', [System.Globalization.CultureInfo]::InvariantCulture)
    }
    return $Value
}

# Writes one properly quoted CSV row.
function Write-CsvRow {
    param([System.IO.StreamWriter] $Writer, [string[]] $Values)
    $sb = New-Object System.Text.StringBuilder
    foreach ($v in $Values) {
        if ($sb.Length -gt 0) { [void] $sb.Append(',') }
        $s = $v
        if ($null -eq $s) { $s = '' }
        if ($s.IndexOf(',') -ge 0 -or $s.IndexOf('"') -ge 0 -or
            $s.IndexOf("`n") -ge 0 -or $s.IndexOf("`r") -ge 0) {
            [void] $sb.Append('"').Append($s.Replace('"', '""')).Append('"')
        } else {
            [void] $sb.Append($s)
        }
    }
    $Writer.WriteLine($sb.ToString())
}

# ---------------------------------------------------------------------------
# Output schemas
# ---------------------------------------------------------------------------

# Identity columns shared by Base, Summary and Transactions files.
$script:IdColumns = @(
    'RowNumber', 'RowType', 'YrPPWk', 'Year', 'PayPeriod', 'Week',
    'Hire_Fin', 'Organization_Name', 'Sub_Unit', 'Employee_Id', 'Last_Name',
    'FI', 'MI', 'PayLoc_FinUnit', 'Var_EAS', 'Borrowed', 'Auto_HL',
    'Annual_Lv_Bal', 'Sick_Lv_Bal', 'LWOP_Lv_Bal', 'FMLA_Hrs', 'FMLA_Used',
    'SLDC_Used'
)

$script:BaseColumns = @(
    'Job_Category', 'D_A', 'Designation_Code', 'Designation_Desc',
    'LDC_Function', 'LDC_Desc', 'Operation_Code', 'Local_Unit',
    'RSC', 'RSC_Desc', 'Lvl', 'FLSA', 'Route_No', 'Loaned_Fin',
    'Effective_Start_YrPPWk', 'Effective_End_YrPPWk', 'Begin_Tour', 'End_Tour',
    'Lunch_Amt', 'Ind_1261', 'Lunch_Ind', 'Daily_Sched_Ind', 'Time_Zone',
    'FTF', 'Duty_Fin', 'OOS', 'Day', 'Parse_Warning'
)

$script:SummaryColumns = @(
    'Job_Raw', 'Day_Of_Week', 'Hours_Type', 'Hours_Note', 'D_A', 'RSC',
    'RSC_Desc', 'Lvl', 'Code_Hours_Raw', 'Hours_Code', 'Specific_Code',
    'Hours_Code_Desc', 'Hours', 'Parse_Warning'
)

$script:TxnColumns = @(
    'EBR_Number', 'D_A', 'Punch_Type', 'Punch_Type_Desc', 'Hours_Code',
    'Specific_Code', 'Hours_Code_Desc', 'Txn_Date', 'Txn_Date_ISO', 'Txn_Time',
    'Txn_Time_HHMMSS', 'Duty_Fin', 'Operation_Code', 'Local_Unit', 'Fin_No_2',
    'Time_Amount', 'Input_Auth', 'Corr_Date', 'Corr_Date_ISO', 'Corr_Time',
    'Corr_Time_HHMMSS', 'Input_Auth_2', 'Corr_Date_2', 'Corr_Date_2_ISO',
    'Corr_Time_2', 'Corr_Time_2_HHMMSS', 'Error_Message', 'Parse_Warning'
)

# ---------------------------------------------------------------------------
# Per-file processing
# ---------------------------------------------------------------------------

function Split-TacsFile {
    param(
        [System.IO.FileInfo] $File,
        [string] $OutDirectory,
        [string] $Prefix,
        [switch] $SkipBaseFile,
        [switch] $SkipSummaryFile,
        [switch] $SkipTransactionsFile
    )

    $utf8Bom = New-Object System.Text.UTF8Encoding($true)
    $outPaths = @{
        Base         = Join-Path $OutDirectory ($Prefix + '.Base.csv')
        Summary      = Join-Path $OutDirectory ($Prefix + '.Summary.csv')
        Transactions = Join-Path $OutDirectory ($Prefix + '.Transactions.csv')
        Unmatched    = Join-Path $OutDirectory ($Prefix + '.Unmatched.csv')
    }

    $writers = @{}
    try {
        if (-not $SkipBaseFile) {
            $writers.Base = New-Object System.IO.StreamWriter($outPaths.Base, $false, $utf8Bom)
            Write-CsvRow $writers.Base ($script:IdColumns + $script:BaseColumns)
        }
        if (-not $SkipSummaryFile) {
            $writers.Summary = New-Object System.IO.StreamWriter($outPaths.Summary, $false, $utf8Bom)
            Write-CsvRow $writers.Summary ($script:IdColumns + $script:SummaryColumns)
        }
        if (-not $SkipTransactionsFile) {
            $writers.Transactions = New-Object System.IO.StreamWriter($outPaths.Transactions, $false, $utf8Bom)
            Write-CsvRow $writers.Transactions ($script:IdColumns + $script:TxnColumns)
        }
        $writers.Unmatched = New-Object System.IO.StreamWriter($outPaths.Unmatched, $false, $utf8Bom)
        Write-CsvRow $writers.Unmatched @('RowNumber', 'FieldCount', 'RawLine')

        $nBase = 0; $nSummary = 0; $nTxn = 0; $nUnmatched = 0; $nHeader = 0
        $unmatchedSamples = New-Object System.Collections.Generic.List[string]

        $lines = [System.IO.File]::ReadAllLines($File.FullName)

        for ($i = 0; $i -lt $lines.Length; $i++) {
            $line = $lines[$i]
            if ([string]::IsNullOrWhiteSpace($line)) { continue }

            $f = Split-CsvLine $line
            $rowNumber = $i + 1

            # Repeated header rows (e.g. concatenated exports) are skipped.
            $first = $f[0].Trim()
            if (($first -replace '^\uFEFF', '') -eq 'YrPPWk') { $nHeader++; continue }

            # ----- identity columns (fields 1-18) -----
            $yrppwk = $f[0].Trim()
            $year = ''; $payPeriod = ''; $week = ''
            if ($yrppwk -match '^(\d{4})(\d{2})(\d{1})$') {
                $year = $Matches[1]; $payPeriod = $Matches[2]; $week = $Matches[3]
            }
            $id = @(
                [string] $rowNumber,
                $null,   # RowType placeholder
                $yrppwk, $year, $payPeriod, $week,
                $f[1].Trim(),  $f[2].Trim(), $f[3].Trim(), $f[4].Trim(), $f[5].Trim(),
                $f[6].Trim(),  $f[7].Trim(), $f[8].Trim(),
                $f[9].Trim(),  $f[10].Trim(), $f[11].Trim(),
                $f[12].Trim(), $f[13].Trim(), $f[14].Trim(),
                $f[15].Trim(), $f[16].Trim(), $f[17].Trim()
            )

            switch ($f.Count) {
                { $_ -ge 39 -and $_ -le 40 } {
                    # ----- Base row: employee/assignment header -----
                    if (-not $writers.ContainsKey('Base')) { break }
                    $warning = ''
                    $jobCat = $f[18].Trim()
                    if ($jobCat -notin @('Base', 'Temp')) {
                        $warning = "Unexpected Job value '$jobCat'"
                    }
                    $daRaw = $f[19].Trim()
                    $desig = $daRaw
                    if ($desig -match '^\d{2}$') { $desig = '0' + $desig }
                    $desigDesc = ''
                    if ($script:Designation.ContainsKey($desig)) { $desigDesc = $script:Designation[$desig] }
                    $ldc = $f[20].Trim()
                    $ldcDesc = ''
                    if ($script:LdcCode.ContainsKey($ldc)) { $ldcDesc = $script:LdcCode[$ldc] }
                    $operRaw = $f[21].Trim()
                    $operCode = ''; $localUnit = ''
                    if ($operRaw -match '^(\d{4})-(\d{2})$') {
                        $operCode = $Matches[1]; $localUnit = $Matches[2]
                    } elseif ($operRaw -ne '') {
                        $warning = "Unexpected Oper/Lu value '$operRaw'"
                    }
                    $rsc = $f[22].Trim()
                    $rscDesc = ''
                    if ($rsc.Length -gt 0 -and $script:RscLetter.ContainsKey($rsc.Substring(0, 1))) {
                        $rscDesc = $script:RscLetter[$rsc.Substring(0, 1)]
                    }
                    $day = ''
                    if ($f.Count -eq 40) { $day = $f[39].Trim() }

                    $id[1] = 'Base'
                    $row = $id + @(
                        $jobCat, $daRaw, $desig, $desigDesc, $ldc, $ldcDesc,
                        $operCode, $localUnit, $rsc, $rscDesc, $f[23].Trim(),
                        $f[24].Trim(), $f[25].Trim(), $f[26].Trim(), $f[27].Trim(),
                        $f[28].Trim(), $f[29].Trim(), $f[30].Trim(), $f[31].Trim(),
                        $f[32].Trim(), $f[33].Trim(), $f[34].Trim(), $f[35].Trim(),
                        $f[36].Trim(), $f[37].Trim(), $f[38].Trim(), $day, $warning
                    )
                    Write-CsvRow $writers.Base $row
                    $nBase++
                    break
                }
                { $_ -ge 21 -and $_ -le 22 } {
                    # ----- Summary row: day / Paid Hours / TACS Hours totals -----
                    if (-not $writers.ContainsKey('Summary')) { break }
                    $warning = ''
                    $jobRaw = $f[18].Trim()
                    $dayOfWeek = ''; $hoursType = ''; $hoursNote = ''
                    if ($jobRaw -match '^(Sunday|Monday|Tuesday|Wednesday|Thursday|Friday|Saturday)$') {
                        $dayOfWeek = $jobRaw
                    } elseif ($jobRaw -match '^(Paid Hours|TACS Hours)\s*:?\s*(.*)$') {
                        $hoursType = $Matches[1]
                        $hoursNote = $Matches[2].Trim()
                    } else {
                        $warning = "Unrecognized summary Job value '$jobRaw'"
                    }

                    $daRaw = $f[19].Trim()
                    $rsc = ''; $lvl = ''; $rscDesc = ''
                    if ($daRaw -match '^([A-Z]\d)-(\d{2})$') {
                        $rsc = $Matches[1]; $lvl = $Matches[2]
                        if ($script:RscLetter.ContainsKey($Matches[1].Substring(0, 1))) {
                            $rscDesc = $script:RscLetter[$Matches[1].Substring(0, 1)]
                        }
                    } elseif ($daRaw -ne '') {
                        $warning = "Unexpected summary D/A value '$daRaw'"
                    }

                    $comp = $f[20].Trim()
                    $hoursCode = ''; $specific = ''; $codeDesc = ''; $hours = ''
                    if ($comp -match '^(\d{3,5}):\s*(\d+(?:\.\d+)?)$') {
                        $codeRaw = $Matches[1]
                        $hours = ConvertTo-NumericText $Matches[2]
                        if ($codeRaw.Length -eq 5) {
                            $hoursCode = $codeRaw.Substring(0, 3)
                            $specific = $codeRaw.Substring(3, 2)
                            if ($script:DaCode5.ContainsKey($codeRaw)) {
                                $codeDesc = $script:DaCode5[$codeRaw]
                            } elseif ($script:DaCode3.ContainsKey($hoursCode)) {
                                $codeDesc = $script:DaCode3[$hoursCode]
                            }
                        } elseif ($codeRaw.Length -eq 3) {
                            $hoursCode = $codeRaw
                            if ($script:DaCode3.ContainsKey($codeRaw)) {
                                $codeDesc = $script:DaCode3[$codeRaw]
                            }
                        } else {
                            $warning = "Unexpected composite code '$codeRaw'"
                        }
                    } elseif ($comp -ne '') {
                        $warning = "Unexpected summary hours composite '$comp'"
                    }

                    $id[1] = 'Summary'
                    $row = $id + @(
                        $jobRaw, $dayOfWeek, $hoursType, $hoursNote, $daRaw,
                        $rsc, $rscDesc, $lvl, $comp, $hoursCode, $specific,
                        $codeDesc, $hours, $warning
                    )
                    Write-CsvRow $writers.Summary $row
                    $nSummary++
                    break
                }
                { $_ -ge 36 -and $_ -le 37 } {
                    # ----- Transaction row: clock ring or hours entry -----
                    if (-not $writers.ContainsKey('Transactions')) { break }
                    $warning = ''
                    $ebr = $f[18].Trim()

                    $daRaw = $f[19].Trim()
                    $punchType = ''; $punchDesc = ''
                    $hoursCode = ''; $specific = ''; $codeDesc = ''
                    if ($daRaw -match '^(BT|ET|OL|IL|MV|OT)$') {
                        $punchType = $daRaw
                        if ($script:PunchType.ContainsKey($daRaw)) {
                            $punchDesc = $script:PunchType[$daRaw]
                        }
                    } elseif ($daRaw -match '^\d{5}$') {
                        $hoursCode = $daRaw.Substring(0, 3)
                        $specific = $daRaw.Substring(3, 2)
                        if ($script:DaCode5.ContainsKey($daRaw)) {
                            $codeDesc = $script:DaCode5[$daRaw]
                        } elseif ($script:DaCode3.ContainsKey($hoursCode)) {
                            $codeDesc = $script:DaCode3[$hoursCode]
                        }
                    } elseif ($daRaw -match '^\d{3}$') {
                        $hoursCode = $daRaw
                        if ($script:DaCode3.ContainsKey($daRaw)) {
                            $codeDesc = $script:DaCode3[$daRaw]
                        }
                    } else {
                        $warning = "Unrecognized transaction D/A value '$daRaw'"
                    }

                    $txnDate = $f[20].Trim()
                    $txnDateIso = ConvertFrom-TacsDate $txnDate
                    if ($txnDate -ne '' -and $txnDateIso -eq '') {
                        $warning = "Unexpected transaction date '$txnDate'"
                    }

                    $txnTime = $f[21].Trim()
                    $txnTimeHH = ConvertFrom-Ergot $txnTime

                    $dutyFin = $f[22].Trim()
                    $operRaw = $f[23].Trim()
                    $operCode = ''; $localUnit = ''
                    if ($operRaw -match '^(\d{4})-(\d{2})$') {
                        $operCode = $Matches[1]; $localUnit = $Matches[2]
                    }

                    $fin2 = $f[24].Trim()
                    $timeAmt = ConvertTo-NumericText $f[25].Trim()
                    $auth = $f[26].Trim()
                    $corrDate = $f[27].Trim()
                    $corrDateIso = ConvertFrom-TacsDate $corrDate
                    $corrTime = $f[28].Trim()
                    $corrTimeHH = ConvertFrom-Ergot $corrTime
                    $auth2 = $f[29].Trim()
                    $corrDate2 = $f[30].Trim()
                    $corrDate2Iso = ConvertFrom-TacsDate $corrDate2
                    $corrTime2 = $f[31].Trim()
                    $corrTime2HH = ConvertFrom-Ergot $corrTime2
                    $msg = $f[32].Trim()

                    $id[1] = 'Transaction'
                    $row = $id + @(
                        $ebr, $daRaw, $punchType, $punchDesc, $hoursCode,
                        $specific, $codeDesc, $txnDate, $txnDateIso, $txnTime,
                        $txnTimeHH, $dutyFin, $operCode, $localUnit, $fin2,
                        $timeAmt, $auth, $corrDate, $corrDateIso, $corrTime,
                        $corrTimeHH, $auth2, $corrDate2, $corrDate2Iso,
                        $corrTime2, $corrTime2HH, $msg, $warning
                    )
                    Write-CsvRow $writers.Transactions $row
                    $nTxn++
                    break
                }
                default {
                    # ----- Anything else: keep the raw line for review -----
                    Write-CsvRow $writers.Unmatched @(
                        [string] $rowNumber, [string] $f.Count, $line
                    )
                    if ($unmatchedSamples.Count -lt 5) { $unmatchedSamples.Add($line) }
                    $nUnmatched++
                    break
                }
            }
        }
    }
    finally {
        foreach ($key in @($writers.Keys)) {
            if ($writers[$key] -ne $null) { $writers[$key].Dispose() }
        }
    }

    if ($nHeader -eq 0) {
        Write-Warning "No header row found in $($File.Name); output may be unreliable."
    }
    if ($nUnmatched -gt 0) {
        Write-Warning "$nUnmatched row(s) did not match any known shape in $($File.Name) (see $($outPaths.Unmatched))."
        foreach ($s in $unmatchedSamples) { Write-Warning "  Sample: $s" }
    }

    $summary = [pscustomobject] @{
        File            = $File.FullName
        BaseRows        = $nBase
        SummaryRows     = $nSummary
        TransactionRows = $nTxn
        UnmatchedRows   = $nUnmatched
        BaseFile        = if ($SkipBaseFile) { $null } else { $outPaths.Base }
        SummaryFile     = if ($SkipSummaryFile) { $null } else { $outPaths.Summary }
        TransactionsFile = if ($SkipTransactionsFile) { $null } else { $outPaths.Transactions }
        UnmatchedFile   = $outPaths.Unmatched
    }

    Write-Host ("{0}: Base={1}  Summary={2}  Transactions={3}  Unmatched={4}" -f `
        $File.Name, $nBase, $nSummary, $nTxn, $nUnmatched) -ForegroundColor Cyan

    return $summary
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

$files = @()
foreach ($p in $InputPath) {
    $items = @(Get-Item -Path $p -ErrorAction SilentlyContinue)
    if ($items.Count -eq 0) {
        Write-Warning "No files match: $p"
        continue
    }
    $files += $items
}
if ($files.Count -eq 0) { throw 'No input files found.' }

$results = @()
foreach ($file in $files) {
    $outDir = $OutputDirectory
    if ([string]::IsNullOrWhiteSpace($outDir)) { $outDir = $file.DirectoryName }
    $outDir = [System.IO.Path]::GetFullPath($outDir)

    $prefix = $OutputPrefix
    if ([string]::IsNullOrWhiteSpace($prefix)) {
        $prefix = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
    } elseif ($files.Count -gt 1) {
        $prefix = $prefix + '-' + [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
    }

    if (-not (Test-Path -LiteralPath $outDir)) {
        New-Item -ItemType Directory -Path $outDir -Force | Out-Null
    }

    Write-Verbose "Processing $($file.FullName) -> $outDir (prefix: $prefix)"

    if ($PSCmdlet.ShouldProcess($file.FullName, "Split TACS report into normalized CSVs")) {
        $results += Split-TacsFile -File $file -OutDirectory $outDir -Prefix $prefix `
            -SkipBaseFile:$SkipBase -SkipSummaryFile:$SkipSummary `
            -SkipTransactionsFile:$SkipTransactions
    }
}

if ($results.Count -gt 0) {
    $totals = [pscustomobject] @{
        BaseRows        = ($results | Measure-Object -Property BaseRows -Sum).Sum
        SummaryRows     = ($results | Measure-Object -Property SummaryRows -Sum).Sum
        TransactionRows = ($results | Measure-Object -Property TransactionRows -Sum).Sum
        UnmatchedRows   = ($results | Measure-Object -Property UnmatchedRows -Sum).Sum
    }
    Write-Host ("TOTAL: Base={0}  Summary={1}  Transactions={2}  Unmatched={3}" -f `
        $totals.BaseRows, $totals.SummaryRows, $totals.TransactionRows, $totals.UnmatchedRows) `
        -ForegroundColor Green
}

$results
