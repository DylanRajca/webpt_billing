# Synopsis - Pull patient data from webPT reports in 'Upload' directory and parse into a billing report, then upload billing report to 'Billing Reports' directory.

# Include functions.ps1
. "betsy\webpt\webpt-billing\webpt_billing_functions.ps1"
# File Paths
$analyticsPath = "$env:USERPROFILE\Documents\WebPT Analytics"
$uploadPath = "$analyticsPath\Upload"
$billedPath = "$analyticsPath\Billing Reports"

#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
##### Main #####
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# If files exist in 'Upload' directory
if ($files = getFiles) {
    main $files
}
