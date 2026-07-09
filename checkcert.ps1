$CERTIFICATE_PATH = "signing_cert.pfx"
if ((Test-Path $CERTIFICATE_PATH) -and $env:CERT_PASSWORD) {
    $SECURE_PASSWORD = ConvertTo-SecureString $env:CERT_PASSWORD -AsPlainText -Force
    $THUMBPRINT = (Get-PfxData -FilePath $CERTIFICATE_PATH -Password $SECURE_PASSWORD).EndEntityCertificates[0].Thumbprint
} else {
    $THUMBPRINT = (Get-ChildItem Cert:\CurrentUser\My |
        Where-Object { $_.Subject -eq "CN=Figoint_Temporary" } |
        Sort-Object NotAfter -Descending |
        Select-Object -First 1).Thumbprint
}
Write-Host "Certificate Thumbprint: $THUMBPRINT"
