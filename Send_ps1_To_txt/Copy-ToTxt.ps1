param(  
[Parameter(
    Position=0, 
    Mandatory=$true, 
    ValueFromPipeline=$true,
    ValueFromPipelineByPropertyName=$true)
]
[Alias('FullName')]
$File)
copy-item $File (((Split-Path $File -Leaf).Split("."))[0]+".txt")
