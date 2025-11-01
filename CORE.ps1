Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Master Installer"
$form.Size = New-Object System.Drawing.Size(500,350)
$form.StartPosition = "CenterScreen"

# Output TextBox (multiline)
$txtOutput = New-Object System.Windows.Forms.TextBox
$txtOutput.Multiline = $true
$txtOutput.ScrollBars = 'Vertical'
$txtOutput.ReadOnly = $true
$txtOutput.WordWrap = $true
$txtOutput.Size = New-Object System.Drawing.Size(460,100)
$txtOutput.Location = New-Object System.Drawing.Point(10,220)
$form.Controls.Add($txtOutput)

# Office Button
$btnOffice = New-Object System.Windows.Forms.Button
$btnOffice.Text = "Install Office 365"
$btnOffice.Size = New-Object System.Drawing.Size(150,50)
$btnOffice.Location = New-Object System.Drawing.Point(50,50)
$btnOffice.Add_Click({
    $txtOutput.AppendText("Starting Office 365 installation..." + [Environment]::NewLine)

    # Run installation in a separate process and capture output
    $process = Start-Process -FilePath "powershell.exe" -ArgumentList "-ExecutionPolicy Bypass -Command `"Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/MrWes3000/SPARK/main/Install_Office.bat' -OutFile '$env:TEMP\Install_Office.bat'; Start-Process '$env:TEMP\Install_Office.bat' -Wait`"" -NoNewWindow -PassThru -Wait

    $txtOutput.AppendText("Office installation finished." + [Environment]::NewLine)
})

# Adobe Button
$btnAdobe = New-Object System.Windows.Forms.Button
$btnAdobe.Text = "Install Adobe Reader"
$btnAdobe.Size = New-Object System.Drawing.Size(150,50)
$btnAdobe.Location = New-Object System.Drawing.Point(250,50)
$btnAdobe.Add_Click({
    $txtOutput.AppendText("Starting Adobe Reader installation..." + [Environment]::NewLine)
    Start-Process -FilePath "powershell.exe" -ArgumentList "-ExecutionPolicy Bypass -Command `"Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/MrWes3000/SPARK/main/Install_Adobe.bat' -OutFile '$env:TEMP\Install_Adobe.bat'; Start-Process '$env:TEMP\Install_Adobe.bat' -Wait`"" -NoNewWindow -Wait
    $txtOutput.AppendText("Adobe Reader installation finished." + [Environment]::NewLine)
})

# Add buttons to form
$form.Controls.Add($btnOffice)
$form.Controls.Add($btnAdobe)

# Show form
$form.Add_Shown({$form.Activate()})
[void]$form.ShowDialog()
