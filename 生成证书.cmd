@echo off
chcp 65001 >nul
echo 正在生成代码签名证书并导出 PFX，请稍候...

powershell -ExecutionPolicy Bypass -Command ^
"$cert = New-SelfSignedCertificate -Type CodeSigning -Subject 'CN=Word Plugin Test' -KeyUsage DigitalSignature -NotAfter (Get-Date).AddYears(50) -CertStoreLocation 'Cert:\CurrentUser\My' -KeyExportPolicy Exportable -KeyAlgorithm RSA -KeyLength 2048; ^
$dir = 'D:\source\repos\李艇的办公助手'; ^
if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }; ^
Export-PfxCertificate -Cert $cert -FilePath (Join-Path $dir '办公助手_TemporaryKey.pfx') -Password (ConvertTo-SecureString -String '123456' -Force -AsPlainText); ^
Write-Host '✅ 导出成功！文件位于：' (Join-Path $dir '办公助手_TemporaryKey.pfx') -ForegroundColor Green; ^
Read-Host '按 Enter 键退出'"

pause