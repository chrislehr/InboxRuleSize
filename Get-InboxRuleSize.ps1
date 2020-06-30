$Mailbox = read-host "Enter the email addresses of the user you want to measure rules on"

$inboxrules = Get-inboxrule -mailbox $mailbox

$RuleMeasure = @()
$charactercount = 0
$totalcharactercount = 0

foreach ($rule in $inboxrules) {
    $charactercount = 0
    
    $rule.psobject.properties | ForEach-Object {
        $charactercount += (Measure-object -character -inputobject $_.value).characters    
    }
    $totalcharactercount = $totalcharactercount + $charactercount
    $RuleMeasure += New-Object -TypeName PSObject -Property @{
        RuleName = $rule.name
        CharacterCount = $charactercount
        SizeinKB =  [math]::Round($charactercount/1024,2)
        TotalCharacterCount = $totalcharactercount
        }
 }

$RuleMeasure | ft RuleName,Charactercount,sizeinKB -AutoSize
$total =  [math]::Round(($totalcharactercount/1024),2)

write-host "Total RuleSet Size for $mailbox is $total Kb" -ForegroundColor Yellow

