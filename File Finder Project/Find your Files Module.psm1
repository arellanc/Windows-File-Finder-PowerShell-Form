Function Find-Files{

[CmdletBinding()]

<#
.SYNOPSIS
    Searches for files in a specified directory and its subdirectories based on name and optional file type.

.DESCRIPTION
    The `Find-Files` function allows users to search for files by specifying a part of the file name.
    An optional file type can be selected to narrow down the search results.
    The function searches recursively within the user's home directory and displays results in a user-friendly GUI.
    Users can also interact with file types using a dynamic dropdown for easy selection.

.PARAMETER FileName
    Specifies the name or part of the name of the file(s) to search for.
    This parameter accepts partial matches and is case-insensitive.

.PARAMETER FileType
    (Optional) Specifies the type of file to search for. This parameter accepts standard file extensions
    (e.g., `.txt`, `.pdf`, `.docx`) and narrows down the results.

.EXAMPLE
    Example 1
    ---------
    Search for files with "report" in their name in any format:
    ```powershell
    .\Find-Files -FileName "report"
    ```
    This will return all files containing "report" in their name within the home directory.

.EXAMPLE
    Example 2
    ---------
    Search for PDF files with "summary" in their name:
    ```powershell
    .\Find-Files -FileName "summary" -FileType ".pdf"
    ```
    This filters results to only `.pdf` files matching the given name.

.NOTES
    Author: CJ Arellano
    Version: 1.0
    Dependencies: PowerShell 5.0 or higher, Windows Forms .NET assembly.

.INPUTS
    [string] FileName
    [string] FileType

.OUTPUTS
    System.Object[]
        Returns an array of objects representing the matching files, including their name,
        last modified time, and directory.

.LINK
    GitHub: https://github.com/arellanc
    Author Website: https://mywork.cjsdevhive.tech
#>


# Define Comments for File Types
$File_Types = @'
Office Documents:
- .doc  - Microsoft Word document (older version)
- .docx - Microsoft Word document (XML format)
- .xls  - Microsoft Excel spreadsheet (older version)
- .xlsx - Microsoft Excel spreadsheet (XML format)
- .ppt  - Microsoft PowerPoint presentation (older version)
- .pptx - Microsoft PowerPoint presentation (XML format)
- .odt  - OpenDocument Text (OpenOffice/LibreOffice)
- .ods  - OpenDocument Spreadsheet (OpenOffice/LibreOffice)
- .odp  - OpenDocument Presentation (OpenOffice/LibreOffice)
- .rtf  - Rich Text Format
- .txt  - Plain Text File
- .csv  - Comma-Separated Values
- .tsv  - Tab-Separated Values
- .html/.htm - HyperText Markup Language document
- .pdf  - Portable Document Format

Image and Media File Types:
- .jpg/.jpeg - JPEG Image
- .png - Portable Network Graphics (image format)
- .gif - Graphics Interchange Format (image format)
- .bmp - Bitmap Image
- .tiff/.tif - Tagged Image File Format
- .svg - Scalable Vector Graphics
- .webp - WebP Image Format
- .mp3 - MPEG Audio Layer 3 (audio file)
- .wav - Waveform Audio File Format (audio)
- .mp4 - MPEG-4 Video File
- .mov - QuickTime Movie (video file)
- .avi - Audio Video Interleave (video file)

Compression and Archive File Types:
- .zip - ZIP archive file
- .rar - WinRAR archive file
- .tar - Tape Archive file (common on Unix/Linux)
- .gz  - Gzip compressed file
- .7z  - 7-Zip compressed file

System and Executable Files:
- .exe - Executable File (Windows program)
- .bat - Batch File (Windows script)
- .ps1 PowerShell File (Windows script)
- .dll - Dynamic Link Library
- .sh  - Shell Script (Unix/Linux)
- .bin - Binary File
- .iso - ISO Disk Image
- .img - Disk Image File

Database and Data Files:
- .sql - SQL Database File
- .db  - Database File
- .mdb - Microsoft Access Database
- .sqlite - SQLite Database

Other:
- .json - JavaScript Object Notation (data format)
- .xml  - eXtensible Markup Language (data format)
- .yml/.yaml - YAML Ain't Markup Language (data format)
- .log  - Log File (text file for logs)
- .ini  - Initialization File (configuration)
- .md   - Markdown File (text file for documentation)
'@

# Define Array for file types in combo box drop-down list
$File_Types_Array = @(
'.doc  Microsoft Word document (older version)',
'.docx Microsoft Word document (XML format)',
'.xls  Microsoft Excel spreadsheet (older version)',
'.xlsx Microsoft Excel spreadsheet (XML format)',
'.ppt  Microsoft PowerPoint presentation (older version)',
'.pptx Microsoft PowerPoint presentation (XML format)',
'.odt  OpenDocument Text (OpenOffice/LibreOffice)',
'.ods  OpenDocument Spreadsheet (OpenOffice/LibreOffice)',
'.odp  OpenDocument Presentation (OpenOffice/LibreOffice)',
'.rtf  Rich Text Format',
'.txt  Plain Text File',
'.csv  Comma-Separated Values',
'.tsv  Tab-Separated Values',
'.html/.htm HyperText Markup Language document',
'.pdf  Portable Document Format',
'.jpg/.jpeg JPEG Image',
'.png Portable Network Graphics (image format)',
'.gif Graphics Interchange Format (image format)',
'.bmp Bitmap Image',
'.tiff/.tif Tagged Image File Format',
'.svg Scalable Vector Graphics',
'.webp WebP Image Format',
'.mp3 MPEG Audio Layer 3 (audio file)',
'.wav Waveform Audio File Format (audio)',
'.mp4 MPEG-4 Video File',
'.mov QuickTime Movie (video file)',
'.avi Audio Video Interleave (video file)',
'.zip ZIP archive file',
'.rar WinRAR archive file',
'.tar Tape Archive file (common on Unix/Linux)',
'.gz  Gzip compressed file',
'.7z  7-Zip compressed file',
'.exe Executable File (Windows program)',
'.bat Batch File (Windows script)',
'.ps1 PowerShell File (Windows script)',
'.dll Dynamic Link Library',
'.sh  Shell Script (Unix/Linux)',
'.bin Binary File',
'.iso ISO Disk Image',
'.img Disk Image File',
'.sql SQL Database File',
'.db  Database File',
'.mdb Microsoft Access Database',
'.sqlite SQLite Database',
'.json JavaScript Object Notation (data format)',
'.xml  eXtensible Markup Language (data format)',
".yml/.yaml YAML Ain't Markup Language (data format)",
'.log  Log File (text file for logs)',
'.ini  Initialization File (configuration)',
'.md   Markdown File (text file for documentation)'
)

# Load .NET assemblies for GUI
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$Default_Text_Color = "#ffffff"

# Create PowerShell Form
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "File Finder"
$Form.Size = New-Object System.Drawing.Size(900, 900)
$Form.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)  # Dark modern background
$Form.StartPosition = "CenterScreen"
$Form.AutoScroll = $true
$Form.ShowIcon = $true
$Form.Icon = "C:\Users\$env:UserName\Documents\File Finder Project\File Finder Icon.ico"

# Create an App Menu Strip
$menuStrip = New-Object System.Windows.Forms.MenuStrip

# Create 'File' menu item
$fileMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$fileMenu.Text = "File"

###########################################################################################

$exitItem = New-Object System.Windows.Forms.ToolStripMenuItem
$exitItem.Text = "Exit"
$exitItem.Add_Click({ $form.Close() })

$fileMenu.DropDownItems.AddRange(@($exitItem))

###########################################################################################
# Create 'Help' menu item
$helpMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$helpMenu.Text = "Help"

# Add sub-items to 'Help' menu
$aboutItem = New-Object System.Windows.Forms.ToolStripMenuItem
$aboutItem.Text = "About"
$aboutItem.Add_Click({ [System.Windows.Forms.MessageBox]::Show("`File Finder created by CJ Arellano

If you have any issues, questions or suggestions don't hesitate to reach out via GitHub, or my website!

- https://github.com/arellanc
- https://mywork.cjsdevhive.tech
") })

$helpMenu.DropDownItems.Add($aboutItem)

###########################################################################################
# Add main menu items to the MenuStrip
$menuStrip.Items.AddRange(@($fileMenu, $helpMenu))

# Add MenuStrip to the Form
$Form.Controls.Add($menuStrip)
$Form.MainMenuStrip = $menuStrip

# Create Input Text Box to Enter Data
$Input_TextBox = New-Object System.Windows.Forms.TextBox
$Input_TextBox.Size = New-Object System.Drawing.Size(250, 40)
$Input_TextBox.Location = New-Object System.Drawing.Point(10, 50)

# Create a Label
$Label = New-Object System.Windows.Forms.Label
$Label.Text = "Enter Item Name"
$Label.AutoSize = $true
$Label.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($Default_Text_Color)
$Label.Location = New-Object System.Drawing.Point(275, 50)

# Creates a Label
$Optional_Label = New-Object System.Windows.Forms.Label
$Optional_Label.Text = "File Type"
$Optional_Label.AutoSize = $true
$Optional_Label.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($Default_Text_Color)
$Optional_Label.Location = New-Object System.Drawing.Point("370", "90")

# Creates a Label
$Author_Label = New-Object System.Windows.Forms.Label
$Author_Label.Text = "File Finder Created by CJ Arellano"
$Author_Label.AutoSize = $true
$Author_Label.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($Default_Text_Color)
$Author_Label.Location = New-Object System.Drawing.Point("10", "830")

# Add Tooltip for the TextBox
$Input_TextBox_Tooltip = New-Object System.Windows.Forms.ToolTip
$Input_TextBox_Tooltip.SetToolTip($Input_TextBox, "Enter the file name or part of it.")

# Create a ComboBox (Drop-down list)
$Combo_Box = New-Object System.Windows.Forms.ComboBox
$Combo_Box.Location = New-Object System.Drawing.Point(10, 90)
$Combo_Box.Size = New-Object System.Drawing.Size(350, 30)
$Combo_Box.Items.AddRange($File_Types_Array)
$Combo_Box.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown

# Add Tooltip for the ComboBox
$Combo_Box_Tooltip = New-Object System.Windows.Forms.ToolTip
$Combo_Box_Tooltip.SetToolTip($Combo_Box, "Select a file type from the list.")

# Flag to prevent recursive updates
$InTextChangedEvent = $false

# Handle TextChanged event for dynamic filtering
$Combo_Box.Add_TextChanged({
    if (-not $InTextChangedEvent) {
        $InTextChangedEvent = $true

        # Get the current text in the ComboBox
        $currentText = $Combo_Box.Text

        # Filter the file types array based on the entered text
        $filteredItems = $File_Types_Array | Where-Object { $_ -like "*$currentText*" }

        # Update the ComboBox items
        $Combo_Box.Items.Clear()
        $Combo_Box.Items.AddRange($filteredItems)

        # Preserve the entered text and caret position
        $Combo_Box.Text = $currentText
        $Combo_Box.SelectionStart = $currentText.Length

        $InTextChangedEvent = $false
    }
})

# Handle SelectedIndexChanged event for selecting from the drop-down
$Combo_Box.Add_SelectedIndexChanged({
    if (-not $InTextChangedEvent) {
        $InTextChangedEvent = $true
        $Combo_Box.Text = $Combo_Box.SelectedItem
        $InTextChangedEvent = $false
    }
})


# Create a Button
$Button = New-Object System.Windows.Forms.Button
$Button.Text = "Search"
$Button.AutoSize = $true
$Button.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($Default_Text_Color)
$Button.Location = New-Object System.Drawing.Point(10, 135)

# Add Tooltip for the Button
$Button_Tooltip = New-Object System.Windows.Forms.ToolTip
$Button_Tooltip.SetToolTip($Button, "Click to search for matching files.")


# Create a Multiline Output Text Box
$Output_TextBox = New-Object System.Windows.Forms.TextBox
$Output_TextBox.Size = New-Object System.Drawing.Size(835, 600)
$Output_TextBox.Location = New-Object System.Drawing.Point(10, 200)
$Output_TextBox.Multiline = $true
$Output_TextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$Output_TextBox.ReadOnly = $true
$Output_TextBox.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 45)
$Output_TextBox.ForeColor = [System.Drawing.Color]::White

# Creates a Loading Progress Bar for your form.
$ProgressBar = New-Object System.Windows.Forms.ProgressBar
$ProgressBar.Location = New-Object System.Drawing.Point("10", "170")
$ProgressBar.Size = New-Object System.Drawing.Size(835,20)
$ProgressBar.Minimum = 0
$ProgressBar.Maximum = 100
$ProgressBar.Step = 10

<#For Progress bar to load add these as your actions are processed
$ProgressBar.Value = 20
$ProgressBar.Refresh()
$ProgressBar.Value = 40
$ProgressBar.Refresh()
$ProgressBar.Value = 60
$ProgressBar.Refresh()
$ProgressBar.Value = 80
$ProgressBar.Refresh()
$ProgressBar.Value = 100
$ProgressBar.Refresh()
#>


# Button Click Event
$Button.Add_Click({
    # Initialize an empty list
    $List = @()

    # Clear the output text box
    $Output_TextBox.Clear()

    # If the file name input box is not empty and the file type drop down is empty, search for any file type
    IF ($Input_TextBox.Text -ne "" -and $Combo_Box.Text -eq "") {
        try {
            # Retrieve matching files
            $List += Get-ChildItem -Recurse -Path "C:\Users\$env:UserName" |
                Where-Object { $_.Name -match $Input_TextBox.Text } |
                Select-Object -Property @{Name='Name'; Expression={$_.Name}},
                               @{Name='LastWriteTime'; Expression={$_.LastWriteTime}},
                               @{Name='Directory'; Expression={$_.Directory}};

                               $ProgressBar.Value = 20;
                               $ProgressBar.Refresh();
                               $ProgressBar.Value = 40;
                               $ProgressBar.Refresh();
                               $ProgressBar.Value = 60;
                               $ProgressBar.Refresh();
                               $ProgressBar.Value = 80;
                               $ProgressBar.Refresh()

            # Process and display results
            If ($List.Count -gt 0) {
                # Display found files in the output text box
                $List | ForEach-Object {
                    $Output_TextBox.AppendText("Name: $($_.Name)`r`n")
                    $Output_TextBox.AppendText("LastWriteTime: $($_.LastWriteTime.ToString('g'))`r`n")
                    $Output_TextBox.AppendText("Directory: $($_.Directory)`r`n")
                    $Output_TextBox.AppendText("`r`n")  # Add a blank line between entries
                    ;$ProgressBar.Value = 100; $ProgressBar.Refresh()
                }
                # Display found files in Out-GridView
                <#$List | Out-GridView -Title "Found Files"#>
            } Else {
                [System.Windows.Forms.MessageBox]::Show("No matching files found.", "No Results", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
        } Catch {
            [System.Windows.Forms.MessageBox]::Show("An error occurred while searching. Please try again.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    } ElseIf ($Combo_Box.Text -ne "" -and $Input_TextBox.Text -ne "") {
        try {
            # Retrieve matching files with a specified file type
            $List += Get-ChildItem -Recurse -Path "C:\Users\$env:UserName" |
                Where-Object {
                    $_.Name -match $Input_TextBox.Text -and $_.Extension -match $Combo_Box.SelectedItem.Split(" ")[0]
                } |
                Select-Object -Property @{Name='Name'; Expression={$_.Name}},
                               @{Name='LastWriteTime'; Expression={$_.LastWriteTime}},
                               @{Name='Directory'; Expression={$_.Directory}};

                               $ProgressBar.Value = 20;
                               $ProgressBar.Refresh();
                               $ProgressBar.Value = 40;
                               $ProgressBar.Refresh();
                               $ProgressBar.Value = 60;
                               $ProgressBar.Refresh();
                               $ProgressBar.Value = 80;
                               $ProgressBar.Refresh()

            # Process and display results
            If ($List.Count -gt 0) {
                # Display found files in the output text box
                $List | ForEach-Object {
                    $Output_TextBox.AppendText("Name: $($_.Name)`r`n")
                    $Output_TextBox.AppendText("LastWriteTime: $($_.LastWriteTime.ToString('g'))`r`n")
                    $Output_TextBox.AppendText("Directory: $($_.Directory)`r`n")
                    $Output_TextBox.AppendText("`r`n")  # Add a blank line between entries
                    ;$ProgressBar.Value = 100; $ProgressBar.Refresh()
                }
                # Display found files in Out-GridView
                <#$List | Out-GridView -Title "Found Files"#>
            } Else {
                [System.Windows.Forms.MessageBox]::Show("No matching files found.", "No Results", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
        } Catch {
            [System.Windows.Forms.MessageBox]::Show("An error occurred while searching. Please try again.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }Elseif($Input_TextBox.Text -eq "" -and $Combo_Box.Text -ne ""){
    
    try {
            # Retrieve matching files
            $List += Get-ChildItem -Recurse -Path "C:\Users\$env:UserName" |
                Where-Object { $_.Extension -match $Combo_Box.SelectedItem.Split(" ")[0]} |
                Select-Object -Property @{Name='Name'; Expression={$_.Name}},
                               @{Name='LastWriteTime'; Expression={$_.LastWriteTime}},
                               @{Name='Directory'; Expression={$_.Directory}};

                               $ProgressBar.Value = 20;
                               $ProgressBar.Refresh();
                               $ProgressBar.Value = 40;
                               $ProgressBar.Refresh();
                               $ProgressBar.Value = 60;
                               $ProgressBar.Refresh();
                               $ProgressBar.Value = 80;
                               $ProgressBar.Refresh()

            # Process and display results
            If ($List.Count -gt 0) {
                # Display found files in the output text box
                $List | ForEach-Object {
                    $Output_TextBox.AppendText("Name: $($_.Name)`r`n")
                    $Output_TextBox.AppendText("LastWriteTime: $($_.LastWriteTime.ToString('g'))`r`n")
                    $Output_TextBox.AppendText("Directory: $($_.Directory)`r`n")
                    $Output_TextBox.AppendText("`r`n")  # Add a blank line between entries
                    ;$ProgressBar.Value = 100; $ProgressBar.Refresh()
                }
                # Display found files in Out-GridView
                <#$List | Out-GridView -Title "Found Files"#>
            } Else {
                [System.Windows.Forms.MessageBox]::Show("No matching files found.", "No Results", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
        } Catch {
            [System.Windows.Forms.MessageBox]::Show("An error occurred while searching. Please try again.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    
    }Else {
        [System.Windows.Forms.MessageBox]::Show("Please enter an item name and select a file type.", "Input Missing", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Exclamation)
    }
})


# Sets Search Button to be executed when the "Enter" Key is pressed.
$Form.AcceptButton = $Button

# Add controls to the form
$Form.Controls.AddRange(@(
    $Input_TextBox,
    $Combo_Box,
    $Button,
    $Label,
    $Author_Label,
    $Optional_Label,
    $Output_TextBox,
    $ProgressBar
))

# Display the form
$Form.ShowDialog()

}