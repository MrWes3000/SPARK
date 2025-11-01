Add-Type -AssemblyName System.Windows.Forms

# Create Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Master Installer"
$form.Size = New-Object System.Drawing.Size(400,250)
$form.StartPosition = "CenterScreen"

# Office Button
$btnOffice = New-Object System.Windows.Forms.Button
$btnOffice.Text = "Install Office 365"
$btnOffice.Size = New-Object System.Drawing.Size(150,50)
$btnOffice.Location = New-Object System.Drawing.Point(50,50)
$btnOffice.Add_Click({
    # Run Office installer from GitHub
    powershell -ExecutionPolicy Bypass -Command "Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/MrWes3000/SPARK/main/Install_Office.bat' -OutFile '$env:TEMP\Install_Office.bat'; Start-Process '$env:TEMP\Install_Office.bat' -Wait"
    [System.Windows.Forms.MessageBox]::Show("Office installation started!")
})

# Adobe Button
$btnAdobe = New-Object System.Windows.Forms.Button
$btnAdobe.Text = "Install Adobe Reader"
$btnAdobe.Size = New-Object System.Drawing.Size(150,50)
$btnAdobe.Location = New-Object System.Drawing.Point(50,120)
$btnAdobe.Add_Click({
    powershell -ExecutionPolicy Bypass -Command "Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/MrWes3000/SPARK/main/Install_Adobe.bat' -OutFile '$env:TEMP\Install_Adobe.bat'; Start-Process '$env:TEMP\Install_Adobe.bat' -Wait"
    [System.Windows.Forms.MessageBox]::Show("Adobe installation started!")
})

# Add buttons to form
$form.Controls.Add($btnOffice)
$form.Controls.Add($btnAdobe)

# Show form
$form.Add_Shown({$form.Activate()})
[void]$form.ShowDialog()
