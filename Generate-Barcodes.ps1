
using namespace IronSoftware.Drawing
# Requires IronBarCode NuGet Package
Add-Type -AssemblyName System.Drawing

# Load IronBarCode (assuming installed via NuGet)
Add-Type -Path "C:\Users\$env:USERNAME\.nuget\packages\ironbarcode\*\lib\net6.0\IronBarCode.dll"


function Generate-SerialNumber {
    param (
        [int]$length = 12
    )
    
    $characters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    $random = New-Object System.Random
    $serialNumber = 1..$length | ForEach-Object {
        $characters[$random.Next(0, $characters.Length)]
    }
    
    return -join $serialNumber
}

function Generate-Barcode {
    param (
        [string]$serialNumber,
        [string]$folder = "barcodes"
    )
    
    # Use script's directory for saving
    $scriptPath = $PSScriptRoot
    $folder = Join-Path $scriptPath $folder
    
    # Create folder if it doesn't exist
    if (-not (Test-Path $folder)) {
        New-Item -ItemType Directory -Path $folder | Out-Null
    }
    
    $barcodePath = Join-Path $folder "$serialNumber`_barcode.png"
    
    # Generate barcode using IronBarCode
    $barcode = [BarcodeWriter]::CreateBarcode($serialNumber, [BarcodeEncoding]::Code128)
    $barcode.ResizeTo(300, 100)
    $barcode.SetMargins(10)
    $barcode.SaveAsPng($barcodePath)
    
    return $barcodePath
}

function Insert-BarcodesToWord {
    param (
        [array]$barcodeData,
        [string]$docPath = "barcodes.docx"
    )
    
    # Use script's directory for saving document
    $scriptPath = $PSScriptRoot
    $docPath = Join-Path $scriptPath $docPath
    
    # Create Word application instance
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    
    # Create new document
    $doc = $word.Documents.Add()
    
    # Add each barcode and serial number
    foreach ($item in $barcodeData) {
        $serialNumber = $item[0]
        $imagePath = $item[1]
        
        # Add serial number
        $paragraph = $doc.Content.Paragraphs.Add()
        $paragraph.Range.Text = "Serial Number: $serialNumber"
        $paragraph.Range.InsertParagraphAfter()
        
        # Add image with error handling
        try {
            $inlineShape = $doc.InlineShapes.AddPicture($imagePath)
            $inlineShape.ScaleWidth = 50
        }
        catch {
            Write-Host "Error inserting image: $_"
        }
        
        # Add space after image
        $doc.Content.Paragraphs.Add()
    }
    
    # Save and close
    $doc.SaveAs([ref]$docPath)
    $doc.Close()
    $word.Quit()
    
    # Release COM objects
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

function Generate-MultipleBarcodes {
    param (
        [int]$numBarcodes
    )
    
    $barcodeData = @()
    
    1..$numBarcodes | ForEach-Object {
        $serialNumber = Generate-SerialNumber
        $barcodePath = Generate-Barcode -serialNumber $serialNumber
        $barcodeData += ,@($serialNumber, $barcodePath)
    }
    
    return $barcodeData
}

# Main execution
$numBarcodes = 15
$barcodeData = Generate-MultipleBarcodes -numBarcodes $numBarcodes
Insert-BarcodesToWord -barcodeData $barcodeData -docPath "barcodes.docx"

Write-Host "Generated $numBarcodes scannable barcodes and saved them to barcodes.docx"