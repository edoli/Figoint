$cert = New-SelfSignedCertificate `
    -Subject "CN=EdoliAddIn_Temporary" `
    -KeyAlgorithm RSA `
    -KeyLength 2048 `
    -NotAfter (Get-Date).AddYears(5) `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -KeyUsage DigitalSignature `
    -Type CodeSigningCert

$CertPassword = Read-Host -Prompt "Enter password for PFX file" -AsSecureString
Export-PfxCertificate -Cert $cert -FilePath "signing_cert.pfx" -Password $CertPassword
[Convert]::ToBase64String([IO.File]::ReadAllBytes("signing_cert.pfx")) | Out-File -Encoding ascii CERT_PFX_BASE64.txt
