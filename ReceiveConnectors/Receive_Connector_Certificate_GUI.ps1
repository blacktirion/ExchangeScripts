Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
Add-Type -AssemblyName System.Windows.Forms

# Set DPI awareness to PerMonitorV2 to handle multiple DPI settings
Add-Type -TypeDefinition @"
    using System;
    using System.Runtime.InteropServices;

    public static class DPIAwareness {
        [DllImport("shcore.dll")]
        private static extern int SetProcessDpiAwareness(int value);

        private const int PROCESS_PER_MONITOR_DPI_AWARE = 2;

        public static void SetPerMonitorDpiAware() {
            SetProcessDpiAwareness(PROCESS_PER_MONITOR_DPI_AWARE);
        }
    }
"@

[DPIAwareness]::SetPerMonitorDpiAware()

# Define the Windows Forms
$form = New-Object Windows.Forms.Form
$form.Text = "Configure Receive Connector Certificate"
$form.Size = New-Object Drawing.Size(1000, 700)  # Increased the height to accommodate the Receive Connector Binding section
$form.StartPosition = "CenterScreen"
$form.Topmost = $true

# Prompt for the server name
$serverNameLabel = New-Object Windows.Forms.Label
$serverNameLabel.Text = "Enter the Server Name:"
$serverNameLabel.Location = New-Object Drawing.Point(30, 30)
$serverNameLabel.AutoSize = $true

$serverNameTextBox = New-Object Windows.Forms.TextBox
$serverNameTextBox.Location = New-Object Drawing.Point(230, 30)  # Adjusted the X position to add more space
$serverNameTextBox.Width = 200

# Find button to search for receive connectors and certificates
$findButton = New-Object Windows.Forms.Button
$findButton.Text = "Find"
$findButton.Location = New-Object Drawing.Point(460, 28)  # Adjusted the X position to add more space

$findButton.Add_Click({
    Clear-SelectedInformation
    Get-ExchangeCertificates
    Get-ReceiveConnectors
})

# Fetch certificates from the Exchange Server
$certificatesDropdown = New-Object Windows.Forms.ComboBox
$certificatesDropdown.Location = New-Object Drawing.Point(30, 130)  # Increased the Y position to add more space
$certificatesDropdown.Width = 920  # Increased the width to accommodate the longer certificate attributes

function Get-ExchangeCertificates {
    $serverName = $serverNameTextBox.Text
    $certificates = Get-ExchangeCertificate -Server $serverName | Select-Object Subject, FriendlyName, NotAfter, Thumbprint, Issuer
    $certificatesDropdown.Items.Clear()
    foreach ($cert in $certificates) {
        $certName = "{0} ({1})" -f $cert.FriendlyName, $cert.Subject
        $certAttributesList[$certName] = @{
            'Subject'      = $cert.Subject
            'FriendlyName' = $cert.FriendlyName
            'NotAfter'     = $cert.NotAfter
            'Thumbprint'   = $cert.Thumbprint
            'Issuer'       = $cert.Issuer
        }
        $certificatesDropdown.Items.Add($certName)
    }
}

# Label for selected certificate details
$selectedCertLabel = New-Object Windows.Forms.Label
$selectedCertLabel.Text = "Selected Certificate Details:"
$selectedCertLabel.Location = New-Object Drawing.Point(30, 180)  # Increased the Y position to add more space
$selectedCertLabel.AutoSize = $true

$selectedCertValueLabel = New-Object Windows.Forms.Label
$selectedCertValueLabel.Text = ""
$selectedCertValueLabel.Location = New-Object Drawing.Point(30, 210)  # Increased the Y position to add more space
$selectedCertValueLabel.AutoSize = $true

function Show-SelectedCertificateDetails {
    $selectedCertName = $certificatesDropdown.SelectedItem
    $selectedCertAttributes = $certAttributesList[$selectedCertName]
    $certDetails = @()
    foreach ($attribute in $selectedCertAttributes.GetEnumerator()) {
        $certDetails += "{0}: {1}" -f $attribute.Key, $attribute.Value
    }
    $selectedCertValueLabel.Text = $certDetails -join "`n"
}

# Fetch the receive connectors and bindings on the specified server
$receiveConnectorsDropdown = New-Object Windows.Forms.ComboBox
$receiveConnectorsDropdown.Location = New-Object Drawing.Point(30, 340)  # Increased the Y position to add more space
$receiveConnectorsDropdown.Width = 920  # Increased the width to accommodate the longer connector names

function Get-ReceiveConnectors {
    $serverName = $serverNameTextBox.Text
    $receiveConnectors = Get-ReceiveConnector -Server $serverName

    $receiveConnectorsDropdown.Items.Clear()
    foreach ($connector in $receiveConnectors) {
        $receiveConnectorsDropdown.Items.Add($connector.Name)
        $bindingsHash[$connector.Name] = $connector.Bindings
        $tlsCertHash[$connector.Name] = $connector.TlsCertificateName
    }
}

# Label for "Select Exchange Certificate" above the certificate drop down
$certificatesTitleLabel = New-Object Windows.Forms.Label
$certificatesTitleLabel.Text = "Select Exchange Certificate:"
$certificatesTitleLabel.Location = New-Object Drawing.Point(30, 100)  # Increased the Y position to add more space
$certificatesTitleLabel.AutoSize = $true

# Label for "Select Receive Connector" above the receive connector drop-down
$receiveConnectorsTitleLabel = New-Object Windows.Forms.Label
$receiveConnectorsTitleLabel.Text = "Select Receive Connector:"
$receiveConnectorsTitleLabel.Location = New-Object Drawing.Point(30, 310)  # Increased the Y position to add more space
$receiveConnectorsTitleLabel.AutoSize = $true

# Label for Receive Connector's TlsCertificateName attribute
$receiveConnectorTlsCertLabel = New-Object Windows.Forms.Label
$receiveConnectorTlsCertLabel.Text = "Receive Connector TlsCertificateName:"
$receiveConnectorTlsCertLabel.Location = New-Object Drawing.Point(30, 400)  # Increased the Y position to add more space
$receiveConnectorTlsCertLabel.AutoSize = $true

$receiveConnectorTlsCertValueLabel = New-Object Windows.Forms.Label
$receiveConnectorTlsCertValueLabel.Text = ""
$receiveConnectorTlsCertValueLabel.Location = New-Object Drawing.Point(30, 430)  # Increased the Y position to add more space
$receiveConnectorTlsCertValueLabel.Width = 920  # Increased the width to accommodate the longer attribute
$receiveConnectorTlsCertValueLabel.AutoSize = $true

function Show-SelectedReceiveConnectorDetails {
    $selectedConnector = $receiveConnectorsDropdown.SelectedItem
    $receiveConnectorTlsCertValueLabel.Text = $tlsCertHash[$selectedConnector] -replace '(?m)^$','NONE'

    # Get the bindings for the selected receive connector
    $bindings = $bindingsHash[$selectedConnector]
    $receiveConnectorBindingsValueLabel.Text = $bindings -join "`n"
}

# Label for Receive Connector's Bindings attribute
$receiveConnectorBindingsLabel = New-Object Windows.Forms.Label
$receiveConnectorBindingsLabel.Text = "Receive Connector Bindings:"
$receiveConnectorBindingsLabel.Location = New-Object Drawing.Point(30, 480)  # Increased the Y position
$receiveConnectorBindingsLabel.AutoSize = $true

$receiveConnectorBindingsValueLabel = New-Object Windows.Forms.Label
$receiveConnectorBindingsValueLabel.Text = ""
$receiveConnectorBindingsValueLabel.Location = New-Object Drawing.Point(30, 510)  # Increased the Y position
$receiveConnectorBindingsValueLabel.Width = 920
$receiveConnectorBindingsValueLabel.Height = 50  # Increased the height to accommodate the Receive Connector Binding section
$receiveConnectorBindingsValueLabel.AutoSize = $true

# Apply button
$applyCertificateButton = New-Object Windows.Forms.Button
$applyCertificateButton.Text = "Apply Certificate"
$applyCertificateButton.Location = New-Object Drawing.Point(370, 600)  # Increased the Y position
$applyCertificateButton.Width = 180  # Increased the width to accommodate the text

$applyCertificateButton.Add_Click({
    $selectedServer = $serverNameTextBox.Text
    $selectedConnector = $receiveConnectorsDropdown.SelectedItem

    # Extract the certificate attributes from the selected certificate details
    $selectedCertName = $certificatesDropdown.SelectedItem
    $selectedCertAttributes = $certAttributesList[$selectedCertName]
    $certName = "<i>{0}<s>{1}" -f $selectedCertAttributes['Issuer'], $selectedCertAttributes['Subject']

    try {
        # Apply the certificate to the selected receive connector
        Set-ReceiveConnector -Identity "$selectedServer\$selectedConnector" -TlsCertificateName $certName

        # Show a success message
        [System.Windows.Forms.MessageBox]::Show("Certificate applied successfully!", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
    catch {
        # Show an error message if the operation encountered an error
        [System.Windows.Forms.MessageBox]::Show("An error occurred while applying the certificate.`nError: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# Function to clear the selected certificate and receive connector information
function Clear-SelectedInformation {
    $selectedCertValueLabel.Text = ""
    $receiveConnectorTlsCertValueLabel.Text = ""
    $receiveConnectorBindingsValueLabel.Text = ""
}

# Add controls to the form
$form.Controls.Add($serverNameLabel)
$form.Controls.Add($serverNameTextBox)
$form.Controls.Add($findButton)
$form.Controls.Add($certificatesTitleLabel)
$form.Controls.Add($certificatesDropdown)
$form.Controls.Add($selectedCertLabel)
$form.Controls.Add($selectedCertValueLabel)
$form.Controls.Add($receiveConnectorsTitleLabel)
$form.Controls.Add($receiveConnectorsDropdown)
$form.Controls.Add($receiveConnectorTlsCertLabel)
$form.Controls.Add($receiveConnectorTlsCertValueLabel)
$form.Controls.Add($receiveConnectorBindingsLabel)
$form.Controls.Add($receiveConnectorBindingsValueLabel)
$form.Controls.Add($applyCertificateButton)

# Event handlers
$certificatesDropdown.Add_SelectedIndexChanged({ Show-SelectedCertificateDetails })
$receiveConnectorsDropdown.Add_SelectedIndexChanged({ Show-SelectedReceiveConnectorDetails })

# Hash table to store receive connector names and bindings
$bindingsHash = @{}
$tlsCertHash = @{}
$certAttributesList = @{}

# Event handler for form closing
$form.Add_FormClosed({ Clear-SelectedInformation })

# Run the form
$form.ShowDialog()
