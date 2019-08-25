$base_dir = Split-Path -Parent $PSScriptRoot
$dist_dir = "$base_dir\dist"
$source_dir = "$base_dir\src"

# Clean

Get-ChildItem $dist_dir -Recurse -Force | Remove-Item

# Concatenate

Get-Content $source_dir\*.vbs | Set-Content $dist_dir\StandardUtils.vbs
