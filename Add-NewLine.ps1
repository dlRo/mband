# Add a new line / blank line after just these sections (modify to your own needs) in an xml file to increase readability
param ($File)

$strings = @('</ClaimsTransformation>',
'</TechnicalProfile>',
'</ClaimType>',
'/ClaimsProvider>',
'</OrchestrationStep>',
'</SubJourney>')

$newContent = Get-Content -Path $File |
    foreach-object {
        $_
        if ($_ -match ($strings -join '|')) { ' ' }
    }
$newContent | out-file -FilePath $File -Encoding Default -Force
