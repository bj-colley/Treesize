# treesizeWeeklyRun.ps1
# Brandon Colley 5-21-2013
# Scheduled to run weekly via task manager. Script takes 1 command line parameter for path to scan.
# The script checks the given path for known folder match. This is where the reports will be stored.
# Variables are then built. The treesize exe path is needed. The main reports folder is needed.
# The report variables are then built based on previously entered information.
# tsizepro.exe is then called to build todays xml file. 
# Next, todays file is compared with last weeks report and outputs a compare excel file.

#Parameters: 
# $scanPath takes a string for the path to scan. Format-- "S:\Macomb Campus"
Param(
[string]$scanPath
)

# Case statement used to determine folder to place reports.
$folder = switch ($scanPath) 
    { 
		'S:\Macomb Campus' {"MC"}
		'S:\Quad Cities Campus' {"QC"}
		'\\ad.wiu.edu\WIU' {"WIU"}
        default {"OTHER"}
    }	

# Variables:
# ----------"Report" variables are built based on the $mainFolder path and the condition above. 
# ----------Add additional conditions above to build new folders for each $scanPath.
$treesizePath = "C:\Program Files (x86)\JAM Software\TreeSize Professional"
$mainFolder = "P:\scripts\FileShare\treesizeWeeklyRun\"

# Report Variables: No need to modify these. Built from above data combined with Get-Date
$todaysReport = $mainFolder+$folder+"\scan{0:yyyyMMdd}.xml" -f (Get-Date)
$lastWeeksReport = $mainFolder+$folder+"\scan{0:yyyyMMdd}.xml" -f (Get-Date).AddDays(-7)
$excelReport = $mainFolder+$folder+"\excelReport.xls"

cd $treesizePath

# Runs Treesize. exports report of $scanPath as xml document with timestamp to $xmlReport
.\tsizepro.exe /xml $todaysReport $scanPath

# Runs Treesize. Suppresses GUI, expands tree 2 levels deep, 
# opens & compares the latest reports, exports as excel document with timestamp.
.\tsizepro.exe /nogui /expand 2 /open $todaysReport /compare $lastWeeksReport /date /excel $excelReport

#manual run option to display xml report for compare
#.\tsizepro.exe /expand 2 /open $todaysReport /compare $lastWeeksReport

# Add archive portion if needed later	