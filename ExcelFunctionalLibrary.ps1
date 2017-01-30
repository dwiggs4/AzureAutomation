#Load the Excel Assembly, Locally or from GAC
try 
{
    Add-Type -ASSEMBLY "Microsoft.Office.Interop.Excel"  | out-null
}
catch 
{
    #If the assembly can't be found this will load the most recent version in the GAC
    [Reflection.Assembly]::LoadWithPartialname("Microsoft.Office.Interop.Excel") | out-null
}

function Open-ExcelApplication 
{
    param
    (
        [switch] $Visible,
        [switch] $HideAlerts
    ) 
    
    $app = New-Object Microsoft.Office.Interop.Excel.ApplicationClass
    $app.Visible  = $Visible
    $app.DisplayAlerts = -not $HideAlerts
    return $app
}

function Get-ExcelWorkBook 
{
    param
    (
        [parameter(Mandatory=$true,ValueFromPipeline=$true)] $inputObject,
        [switch] $Visible,
        [switch] $readonly
    )

    [Microsoft.Office.Interop.Excel.ApplicationClass]$app = $null 
    if($inputObject -is [Microsoft.Office.Interop.Excel.ApplicationClass]) 
    {
        $app = $inputObject
        $WorkBook = $app.ActiveWorkbook
    } 
    
    else 
    {
        $app = Open-ExcelApplication -Visible:$Visible  
        try 
        {
            if($inputObject.Contains("\\") -or $inputObject.Contains("//")) 
            {
                $WorkBook = $app.Workbooks.Open($inputObject,$true,[System.Boolean]$readonly)
            } 
            
            else 
            {
                $WorkBook = $app.Workbooks.Open((Resolve-path $inputObject),$true,[System.Boolean]$readonly)
            }
        } 
        
        catch 
        {
            $WorkBook = $app.Workbooks.Open((Resolve-path $inputObject),$true,[System.Boolean]$readonly)
        }
    } 

    $app.CalculateFullRebuild() 
    return $WorkBook
}

function Get-ExcelWorkSheet 
{
    param
    (
        [parameter(Mandatory=$true,ValueFromPipeline=$true)] $inputObject,
        $SheetName,
        [switch] $Visible,
        [switch] $readonly
    )
  
    if($inputObject -is [Microsoft.Office.Interop.Excel.Workbook]) 
    {
        $WorkBook = $inputObject
    } 
    
    else 
    {
        $WorkBook = Get-ExcelWorkBook $inputObject -Visible:$Visible `
                                                   -readonly:$readonly
    }
    
    if (($SheetName -eq $null) -or $SheetName -eq 0) 
    {
        $WorkBook.ActiveSheet
    } 
    
    else 
    {
        $WorkBook.WorkSheets.item($SheetName)
    } 
}

function Import-Row 
{
    param
    (
        $Row,[hashtable] $Headers =@{},
        $ColumnStart = 1,
        $ColumnCount = $Row.Value2.Count
    )
    
    $output = @{}
    for ($index=$ColumnStart;$index -le $ColumnCount;$index ++)
    {
        if($Headers.Count -eq 0)
        {
            $Key = $Index
        } 
        
        Else 
        {
            $Key = $Headers[$index]
        }
        
        $output.Add($Key,$row.Cells.Item(1,$index).Text)
    }
    return $output
}

function Release-Ref ($ref) 
{
    ([System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$ref) -gt 0)  | Out-Null
    [System.GC]::Collect() | Out-Null
    [System.GC]::WaitForPendingFinalizers() | Out-Null
}

function Close-ExcelApplication 
{
    param
    (
        [parameter(Mandatory=$true,ValueFromPipeline=$true)] $inputObject
    )
    
    if ($inputObject -is [Microsoft.Office.Interop.Excel.ApplicationClass]) 
    {
        $app = $inputObject  
    } 
    else 
    {
    $app = $inputObject.Application
    Release-Ref $inputObject
    }

    $app.ActiveWorkBook.Close($false) | Out-Null
    $app.Quit() | Out-Null
    Release-Ref $app
}

function Import-Excel 
{
    param
    (
        [parameter(Mandatory=$true,ValueFromPipeline=$true)] $inputObject,
        [Object] $SheetName,
        [switch] $Visible,
        [switch] $readonly,
        [int] $startOnLineNumber =1,
        [switch] $closeExcel,
        [switch] $asHashTable,
        [hashtable] $FieldNames =@{}
    )
    
    #Check what the input is. 
    if ($inputObject -is [Microsoft.Office.Interop.Excel.range]) 
    { 
        $range = $inputObject
    } 
    elseif ($inputObject -isnot [Microsoft.Office.Interop.Excel.Worksheet]) 
    { 
        $WorkSheet = Get-ExcelWorkSheet $inputObject -SheetName $SheetName `
                                                     -Visible:$Visible `
                                                     -readonly:$readonly  
        $range = $WorkSheet.UsedRange
    } 
    else 
    {
        $WorkSheet = $inputObject
        $range = $WorkSheet.UsedRange
    }
    
    # populate the Header 
    if ($FieldNames.Count -eq 0) 
    {
        $FieldNames = Import-Row $range.Rows.Item($startOnLineNumber++)              
    }

    for ($RowIndex=$startOnLineNumber;$RowIndex -le $range.Rows.Count;$RowIndex++) 
    {
        $output = Import-Row $range.Rows.Item($RowIndex) -Headers $FieldNames
    
        if ($asHashtAble) 
        {
            Write-Output $output
        } 
        else 
        {
            New-Object PSObject -property $output
        }
    }  

    # If we opened Excel, we should close Excel.
    if ($closeExcel) 
    {   
        $WorkSheet.Activate() | Out-Null
        Close-ExcelApplication $WorkSheet
    } 
}