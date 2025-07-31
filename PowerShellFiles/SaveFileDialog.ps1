# Add-Type -AssemblyName System.Windows.Forms

Function fileSave {
    $saveFileDialog = [System.Windows.Forms.SaveFileDialog]@{
        CheckPathExists  = $true
        CreatePrompt     = $true
        OverwritePrompt  = $true
        InitialDirectory = [Environment]::GetFolderPath('MyDocuments')
        FileName         = 'NewFile'
        Title            = 'Choose directory to save the output file'
        Filter           = "Text documents (.txt)|*.txt"
    }

    # Show save file dialog box
   # if($saveFileDialog.ShowDialog() -eq 'Ok') {
     #   New-Item -Path $saveFileDialog.FileName -ItemType File -Force
    #}
    $saveFileDialog.ShowDialog()
}

fileSave