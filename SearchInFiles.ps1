###########################################################################
# Searching for a string inside different types of file
# 1. PDF
# 2. Word
# 3. Excel
# 4. PowerPoint
# 5. Other types of files readable in notepad
# Copyrights by Paweł Grabiec Hornbeam Consulting, Freeware software
###########################################################################


Param(
    [Parameter(Mandatory=$True, HelpMessage="Enter a path like: c:\users\kkowalski\downloads - script will search in this folder and it's subfolders")]
    [string]$Path,
    [Parameter(Mandatory=$True, HelpMessage="It's kinda obvious, isn't it?")]
    [string]$TextToFind,
    [Parameter(Mandatory=$True, HelpMessage="Enter the file extension like: pdf, docx, xlsx, or txt")]
    [string]$FileType
)


#$ErrorActionPreference= 'silentlycontinue'

##################################
# Path validations               #
##################################

$isPathValid = Test-Path -Path $Path
if($isPathValid -eq $False){
    Write-Host ""
    Write-Host ""
    Write-Host "The path is incorrect";
    Exit;
}

$results = @()

########################################
# Search for a string in a PDF files   #
########################################
If ($FileType -eq "pdf") {

    # Include dll

    if ($MyInvocation.MyCommand.CommandType -eq "ExternalScript")
    { $ScriptPath = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition }
    else
    { $ScriptPath = Split-Path -Parent -Path ([Environment]::GetCommandLineArgs()[0])
        if (!$ScriptPath){ $ScriptPath = "." } }

    Unblock-File $ScriptPath"\lib\itextsharp.dll" 
    Add-Type -Path $ScriptPath"\lib\itextsharp.dll" 

    $pdfs = Get-Childitem $Path -Include "*.$FileType" -Recurse | Where-Object { !($_.psiscontainer) }
    $keyword = $TextToFind

    <# Search for the string #>
    Foreach($pdf in $pdfs)  {
        # prepare the pdf
        $reader = New-Object iTextSharp.text.pdf.pdfreader -ArgumentList $pdf.FullName -ErrorAction SilentlyContinue
        # for each page
        for($page = 1; $page -le $reader.NumberOfPages; $page++) {
            # set the page text
            $pageText = [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($reader,$page).Split([char]0x000A) 
            if($pageText -match $keyword) {
                $response = @{
                    keyword = $keyword
                    file = $pdf.FullName
                    page = $page
                }
                $results += New-Object PSObject -Property $response
            }
        }
        $reader.Close()
    }



########################################
# Search for a string in a Excel files #
########################################
} elseif ($FileType -eq "xls" -or $FileType -eq "xlsx" ) {
    
    <# Initiate excel document #>
    $Excel = New-Object -ComObject Excel.Application
    $Excel.visible = $False

    $SourcePath = Convert-Path $Path

    $Sources = Get-Childitem $SourcePath -Include *.xlsx,*.xls -Recurse | Where-Object { !($_.psiscontainer) }

    <# Search for the string #>
    Foreach ($Source In $Sources) {
        $Source = $Source.FullName
        $Source = Convert-Path $Source
        $Workbook = $Excel.Workbooks.Open($Source)
        ForEach ($Worksheet in @($Workbook.Sheets)) {
            # Find Method https://msdn.microsoft.com/en-us/vba/excel-vba/articles/range-find-method-excel
            $Found = $WorkSheet.Cells.Find($TextToFind) #What
            If ($Found) {
                # Address Method https://msdn.microsoft.com/en-us/vba/excel-vba/articles/range-address-property-excel
                $BeginAddress = $Found.Address(0,0,1,1)
                #Initial Found Cell
                $response = @{
                    WorkSheet = $Worksheet.Name
                    Column = $Found.Column
                    Row =$Found.Row
                    Text = $Found.Text
                    Address = $BeginAddress
                }
                $results += New-Object PSObject -Property $response
                Do {
                    $Found = $WorkSheet.Cells.FindNext($Found)
                    $Address = $Found.Address(0,0,1,1)
                    If ($Address -eq $BeginAddress) {
                        BREAK
                    }
                    $response = @{
                        WorkSheet = $Worksheet.Name
                        Column = $Found.Column
                        Row =$Found.Row
                        Text = $Found.Text
                        Address = $Address
                    }
                    $results += New-Object PSObject -Property $response                 
                } Until ($False)
            }
        }
        $workbook.close($false)
    }
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$excel)
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
    Remove-Variable excel -ErrorAction SilentlyContinue


########################################
# Search for a string in a Word files  #
########################################
} elseif ($FileType -eq "doc" -or $FileType -eq "docx" ) {

    <# Initiate word document #>
    $Word = New-Object -ComObject Word.Application
    $Word.visible = $False

    $files = Get-Childitem $Path -Include *.docx,*.doc -Recurse | Where-Object { !($_.psiscontainer) }

    <# Search for the string #>
    Foreach ($file In $files) {
        $Document = $Word.Documents.Open($file.FullName, $false, $true) 
        if($Document.content.find.execute($TextToFind)){
            $results += $file.FullName
        }
    $document.close()             
    }   

    # Quit word document
    $Word.Quit()

########################################
# Search for a string in a PPT files  #
########################################
} elseif ($FileType -eq "ppt" -or $FileType -eq "pptx" ) {

    $PPT = New-Object -ComObject PowerPoint.Application
    $PPT.Visible = $False

    $files = Get-Childitem $path -Include *.pptx,*.ppt -Recurse | Where-Object { !($_.psiscontainer) }

    <# Search for the string #>
    Foreach ($file In $files) {
        $Document = $PPT.Presentations.Open($file.FullName, $false, $true, $false)
        ForEach ($slide in $Document.Slides) {
            ForEach ($shape in $slide.shapes) {
                If ($shape.HasTextFrame) {
                    If ($shape.TextFrame.HasText) {
                        If($shape.TextFrame.TextRange.Text -Match $TextToFind) {
                            $response = @{
                                keyword = $TextToFind
                                file = $file.FullName
                                slideNo = $slide.SlideNumber
                            }
                        $results += New-Object PSObject -Property $response
                        }
                    }
                }
            }
        }
        $Document.Close() 
    }
    $PPT.PresentationClose

         
      
    # Quit PPT document
    $PPT.Quit()
    

########################################
# Search in other files                #
########################################
} else {

    $files = Get-Childitem $path -Recurse -Include "*.$FileType"

    $results = $files | select-string -Pattern $TextToFind | group path | select name

}

# Show results
Write-Host ""
Write-Host "done"
Write-Host "Copyrights by Paweł Grabiec Hornbeam Consulting, Freeware software"
$results | Out-GridView -Title "Search in Files                                                        Copyrights by Paweł Grabiec Hornbeam Consulting, Freeware software"




