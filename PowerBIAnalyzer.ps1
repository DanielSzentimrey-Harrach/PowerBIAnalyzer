<#
    This script reads a Power BI Template (.pbit) file, and identifies which columns and/or measures are not used anywhere in the visuals or in DAX expressions.
#>
# Define a mandatory templatePath parameter
param(
    [Parameter(Mandatory=$true)]
    [string]$templatePath
)

Set-StrictMode -Version Latest

#process the template file (copy to temp location, rename to .zip, unzip, delete at the end) and get the contents of the DataModelSchema and Layout files
function GetFileContents($templatePath) {
    $randomFolderName = (New-Guid).ToString()
    $tempFolderPath = Join-Path "$env:USERPROFILE\Downloads" $randomFolderName
    $tempZipPath = Join-Path $tempFolderPath "temp.zip"

    # Create a temporary directory
    New-Item -ItemType Directory -Path $tempFolderPath -Force | Out-Null

    # Copy the template file to the temporary directory and unzip it
    Copy-Item -Path $templatePath -Destination $tempZipPath -Force
    Expand-Archive -Path $tempZipPath -DestinationPath $tempFolderPath -Force

    # Get the contents of the DataModelSchema and Layout files
    $dataModel = Get-Content -Path (Join-Path $tempFolderPath "DataModelSchema") -Encoding unicode
    $layout = Get-Content -Path (Join-Path $tempFolderPath "Report\Layout") -Encoding unicode

    # Delete the temporary directory
    Remove-Item -Path $tempFolderPath -Recurse -Force

    return $dataModel, $layout
}

#Get report details into 2 files, make sure they are treated as objects
$dataModel, $layout = GetFileContents $templatePath.Trim().Trim('"')
$dataModel = $dataModel -join " " | ConvertFrom-Json
$layout = $layout | ConvertFrom-Json

#create variables to store the list of columns and measures used in the report, as well as the various expressions from the files
$fieldList = @()
$allConfigDetails = ""
$allFilterDetails = ""

#Iterate through all tables. Only hidden tables will have an isHidden property, so filtering by its existence is acceptable
foreach ($table in ($dataModel.model.tables | Where-Object {-not (Get-Member -InputObject $_ -Name "isHidden")})) {
    #iterate through all columns and add them to the filedList collection
    foreach ($column in ($table.columns | Where-Object {-not (Get-Member -InputObject $_ -Name "isHidden")})) {
        $type = "column"
        $expression = ""
        if (Get-Member -InputObject $column -Name "type") {
            $type = $column.type
            if ($column.type -eq "calculated") {
                $expression = $column.expression
            }
        }
        $fieldList += [PSCustomObject]@{
            table = $table.name
            field = $column.name
            type = $type
            used = $false
            expression = $expression
        }
    }
    #iterate through all measures and add them to the filedList collection
    if (Get-Member -InputObject $table -Name "measures") {
        foreach ($measure in $table.measures) {
            $fieldList += [PSCustomObject]@{
                table = $table.name
                field = $measure.name
                type = "measure"
                used = $false
                expression = $measure.expression
            }
        }
    }
}

#iterate through all visuals and add the configuration details into the allConfigDetails variable
foreach ($section in $layout.sections) {
    foreach ($visualContainer in $section.visualContainers) {
        $config = ConvertFrom-Json $visualContainer.config
        $allConfigDetails += " " + (ConvertTo-Json $config.singleVisual -Depth 100)

        $filters = ConvertFrom-Json $visualContainer.filters
        foreach ($filter in $filters) {
            $allFilterDetails += " " + (ConvertTo-Json $filter -Depth 100)
        }
    }
}

#iterate through all fields and mark them as used if they appear anywhere in the configs
foreach ($field in $fieldList) {
    ### REGEX PATTERN TO MATCH FIELD IN CONFIGS
    # Table.field
    # 'Table'.field
    # FUNTION(Table.field)
    # FUNCTION('Table'.field)
    # "Property":  "field" <-- two spaces after the colon
    ###
    $pattern1 = "$([Regex]::Escape($field.table))[']?`.$([Regex]::Escape($field.field))[)]?`"|`"Property`":[ ]{1,2}`"$([Regex]::Escape($field.field))`""
    ### REGEX PATTERN TO MATCH FIELD IN FILTERS
    # "Property":  "field" <-- two spaces after the colon
    ###
    $pattern2 = "`"Property`":[ ]{1,2}`"$([Regex]::Escape($field.field))`""
    if ((Select-String -InputObject $allConfigDetails -Pattern $pattern1) -or (Select-String -InputObject $allFilterDetails -Pattern $pattern2)) {
        $field.used = $true
    }
}

#recursively iterate through all unused fields to see if they are used in any DAX expressions of used fields
do {
    $changeMade = $false
    $expressions = ($fieldList | Where-Object {$_.used}).expression -join " "
    foreach ($field in ($fieldList | Where-Object {-not $_.used})) {
        ### REGEX PATTERN TO MATCH FIELD IN DAX EXPRESSIONS
        # Table[field]
        # 'Table'[field]
        ###
        $pattern = "$([Regex]::Escape($field.table))[']?\[$([Regex]::Escape($field.field))\]"
        if (Select-String -InputObject $expressions -Pattern $pattern) {
            $field.used = $true
            $changeMade = $true
        }
    }
} while($changeMade)

#print out all unused fields
$fieldList | Where-Object {-not $_.used} | Select-Object table, field, type, used