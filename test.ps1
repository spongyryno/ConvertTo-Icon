#======================================================================================================================================================================================================
#======================================================================================================================================================================================================
$loc = if ($PsScriptRoot) {$PsScriptRoot} else {$pwd.Path}

.\ConvertTo-Icon.ps1 -SourcePath smile.png -TargetPath smile1.ico -Force -Formats '16,32,48,64'
.\ConvertTo-Icon.ps1 -SourcePath smile.png -TargetPath smile2.ico -Force -Formats '16 4bpp BMP,32 4bpp BMP,48 4bpp BMP,64 4bpp BMP'
.\ConvertTo-Icon.ps1 -SourcePath smile.png -TargetPath smile3.ico -Force -Formats '64 24bpp'
.\ConvertTo-Icon.ps1 -SourcePath smile.png -TargetPath smile4.ico -Force -Formats '64 32bpp'
.\ConvertTo-Icon.ps1 -SourcePath smile.png -TargetPath smile5.ico -Force -Formats '64 24bpp BMP'
.\ConvertTo-Icon.ps1 -SourcePath smile.png -TargetPath smile6.ico -Force -Formats '64 32bpp BMP'
.\ConvertTo-Icon.ps1 -SourcePath smile.png -TargetPath smile7.ico -Force -Formats '16x16 PNG, 16x16 BMP, 32x32, 256 PNG'

start smile7.ico

return

