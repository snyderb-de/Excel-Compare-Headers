# Comapre headers from selected Excel Files
# TODO: add a block to unhide all columns in each file before comparing headers

# Prompt user for files to compare
Add-Type -AssemblyName System.Windows.Forms
$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$openFileDialog.Multiselect = $true
$openFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
$dialogResult = $openFileDialog.ShowDialog()

if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
    $files = $openFileDialog.FileNames

    # Read in headers from first file
    $base_df = New-Object -ComObject Excel.Application
    $base_wb = $base_df.Workbooks.Open($files[0])
    $base_ws = $base_wb.Worksheets.Item(1)
    $base_headers = @()
    for ($i = 1; $i -le $base_ws.UsedRange.Columns.Count; $i++) {
        $header = $base_ws.Cells.Item(1, $i).Value()
        $base_headers += $header
    }
    $base_wb.Close()

    # Compare headers from other files to headers from first file
    $results = @()
    foreach ($file in $files[1..($files.Length - 1)]) {
        $df = New-Object -ComObject Excel.Application
        $wb = $df.Workbooks.Open($file)
        $ws = $wb.Worksheets.Item(1)
        $headers = @()
        for ($i = 1; $i -le $ws.UsedRange.Columns.Count; $i++) {
            $header = $ws.Cells.Item(1, $i).Value()
            $headers += $header
        }
        $wb.Close()
        if ($headers -ne $base_headers) {
            $result = "Headers in $file do not match headers in base file."
            Write-Host $result
            $results += $result
        }
    }

    # Write results to text file
    $outputFilePath = "output.txt"
    $results | Out-File -FilePath $outputFilePath

    Write-Host "Results written to $outputFilePath."
} else {
    Write-Host "No files selected."
}
