<#------------------------------------------------------------------------------
    Jason McClary
    mcclarj@mail.amc.edu
    06 Jul 2016

    
    Description:
    Move all faxes from main folder to a monthly folder except the last few days worth
    
    Arguments:
    None
        
    Tasks:
    - Create monthly folder if not there yet
    - Name folder as Year-Month (1993-08 for August 1993)
    - Move files from that month in to that folder
    - Only file PDFs older then one week old

        
--------------------------------------------------------------------------------
                                CONSTANTS
------------------------------------------------------------------------------#>
set-variable folderToClean -option Constant -value "NewTest"


<#------------------------------------------------------------------------------
                                FUNCTIONS
------------------------------------------------------------------------------#>

    
<#------------------------------------------------------------------------------
                                    MAIN
------------------------------------------------------------------------------#>
$DestinationDir = "1993-08"

# Get all the PDFs in the folder
$files = get-childitem $folderToClean *.pdf

IF ($files.count -gt 0) {
    FOREACH ($file in $files) {
        # For each file older then 2 days...
        IF ($file.LastWriteTime -lt (Get-Date).adddays(-1).date) {

            # Use that files date to make a matching folder name (ex. 1993-08)
            $DestinationDir = $folderToClean + "\$($file.LastWriteTime.year)" + "-"
            $month = "$($file.LastWriteTime.month)"
            # Add a leading zero to the month
            IF ($month.length -lt 2) {
                $month = "0" + $month
            }
            $DestinationDir = $DestinationDir + $month

            # Now see if there is a folder for that Year/Month if not make it
            IF (!(Test-Path $DestinationDir)) {
                New-Item $DestinationDir -type directory
            }

            $DestinationFile = $DestinationDir + "\" + $file.Name
            
            # Move the file to that folder
            Move-Item $file.fullname $DestinationDir

        }
    }
}