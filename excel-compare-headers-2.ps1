# Prompt user for base file
Add-Type -AssemblyName System.Windows.Forms
$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$openFileDialog.Filter = "Excel Files (*.xls;*.xlsx;*.xlsm;*.xlsb)|*.xls;*.xlsx;*.xlsm;*.xlsb"
$openFileDialog.Title = "Please choose the base file with the correct headers"
$dialogResult = $openFileDialog.ShowDialog()

if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
    $base_file = $openFileDialog.FileName

    # Prompt user for files to compare
    $openFileDialog.Multiselect = $true
    $openFileDialog.FileName = ""
    $openFileDialog.Title = "Select all the files to compare to the base file"
    $dialogResult = $openFileDialog.ShowDialog()

    if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
        $compare_files = $openFileDialog.FileNames

        # Get output file path in the same directory as compare files
        $output_file = Join-Path (Split-Path $compare_files[0]) "output.txt"

        # Create output file
        $output = New-Object System.Collections.ArrayList

        # Read in headers from base file and add to output
        $base_df = New-Object -ComObject Excel.Application
        $base_wb = $base_df.Workbooks.Open($base_file)
        $base_ws = $base_wb.Worksheets.Item(1)
        $base_headers = @()
        $output.Add("Base File Headers")
        $output.Add("Column Letter`tHeader Value")
        for ($i = 1; $i -le $base_ws.UsedRange.Columns.Count; $i++) {
            $header = $base_ws.Cells.Item(1, $i).Value()
            $column_letter = [char](64 + $i)
            $output.Add("$column_letter`t$header")
            $base_headers += $header
        }
        $base_wb.Close()
        }
        }

        # Compare headers from other files to headers from base file
        $all_match = $true
        foreach ($file in $compare_files) {
            $df = New-Object -ComObject Excel.Application
            $wb = $df.Workbooks.Open($file)
            $ws = $wb.Worksheets.Item(1)
            $headers = @()
            $output.Add("")
            $output.Add("$file Headers")
            $output.Add("Column Letter`tHeader Value")
            for ($i = 1; $i -le $ws.UsedRange.Columns.Count; $i++) {
                $header = $ws.Cells.Item(1, $i).Value()
                $column_letter = [char](64 + $i)
                $headers += $header
                $output.Add("$column_letter`t$header")
            }
            if ($headers -ne $base_headers) {
                $all_match = $false
                $output.Add("")
                $output.Add("$file Headers that do not match Base File")
                $output.Add("Column Letter`tHeader Value")
                for ($i = 1; $i -le $ws.UsedRange.Columns.Count; $i++) {
                    $header = $ws.Cells.Item(1, $i).Value()
                    $column_letter = [char](64 + $i)
                    if ($header -ne $base_headers[$i-1]) {
                        $output.Add("$column_letter`t$header")
                    }
                }
            }
            $wb.Close()
        }

        # Write results to text file
        $output | Out-File -FilePath $output_file
