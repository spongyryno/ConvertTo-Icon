#======================================================================================================================================================================================================
#
# Copyright (C) SpongySoft. All rights reserved.
#
#======================================================================================================================================================================================================

<#
	.SYNOPSIS
	Converts an image file, ideally with transparency (like a PNG), to an Icon file (.ico)

	.DESCRIPTION
	Converts an image file, ideally with transparency (like a PNG), to an Icon file (.ico). It can control which different icon sizes and bit depths are embedded in the ICO file

	.PARAMETER SourcePath
	The path to the input image file

	.PARAMETER TargetPath
	The path to the target icon file, or the target folder if you're using pipeline input

	.PARAMETER Formats
	A comma-delimited string (or string array) consisting of the formats desired in the output. Valid formats include 3 things:

		* The dimensions
		* The bit-depth (bits per pixel)
		* The desired internal format (BMP or PNG)

	The dimensions can be a single number, 16, 32, 48, 64,128 or 256, or it can be the same number twice with an 'x'
	in between, such as "32x32".

	The bit depth is optional, and will default to 32bpp. It must be followed by "bpp".

	The format can be omitted, in which case both BMP and PNG will be generated, and the smaller of the two will be used.

	Note that a 24bpp PNG icon will have zero transparency.

	Examples of valid formats include:

		32
		64 24bpp BMP
		16x16,32x32,64x64 8bpp
		16x16 PNG, 16x16 BMP, 32x32

	.INPUTS
	You can pass in paths to input file names. PNGs work best, since they have embedded transparency, but other
	formats that System.Drawing can load should work as well.

	.OUTPUTS
	Icon files.

	.EXAMPLE
	ConvertTo-Ico -SourcePath MyLogo.png -TargetPath MyLogoIcon.ico -Formats 16,32,64

	.EXAMPLE
	Get-ChildItem -Path . -Filter *.png | ConvertTo-Ico -TargetPath . -Formats 16,32,64,256

	.LINK
	None.
#>

[CmdletBinding()]
param(
	[string][Parameter(ValueFromPipeline)]$SourcePath,
	[string]$TargetPath,
	[string[]]$Formats = @('16x16 32bpp,32x32 32bpp,48x48 32bpp,64x64 32bpp,128x128 32bpp,256x256 32bpp'),
	[switch]$Force
)


#======================================================================================================================================================================================================
#======================================================================================================================================================================================================
begin
{

	#==================================================================================================================================================================================================
	# Get the location of "this" script
	#==================================================================================================================================================================================================
	$loc = if ($PsScriptRoot) {$PsScriptRoot} else {$pwd.Path}


	#==================================================================================================================================================================================================
	# Create a new 1-bit-per-pixel AND mask for an existing 32-bit-per-pixel ARGB Bitmap based on the alpha channel
	#
	# The 1's in the AND mask will represent opaque, where the icon's pixel should be shown, and 0's are transparent,
	# where the background should show through
	#==================================================================================================================================================================================================
	function New-AndMask
	{
		param([Drawing.Bitmap]$Bitmap)

		if ($bitmap.PixelFormat -ne [Drawing.Imaging.PixelFormat]::Format32bppArgb)
		{
			throw [NotSupportedException]::new("The pixel format ""$($bitmap.PixelFormat)"" is not supported (only Format32bppArgb).")
			return
		}

		#
		# Here, we're going to use a color matrix to transform the input colors so that
		# all of the colors are an INVERSE of the alpha. In other words, we ignore all input
		# RGB, and we make the output have WHITE (255,255,255) where the alpha is 0% (opaque),
		# and black (0,0,0) where the alpha is 100% (transparent)
		#
		$colorMatrix = [Drawing.Imaging.ColorMatrix]::new()

		# source R
		$colorMatrix[0,0] = 0.0  # target R
		$colorMatrix[0,1] = 0.0  # target G
		$colorMatrix[0,2] = 0.0  # target B
		$colorMatrix[0,3] = 0.0  # target A
		$colorMatrix[0,4] = 0.0  #

		# source G
		$colorMatrix[1,0] = 0.0  # target R
		$colorMatrix[1,1] = 0.0  # target G
		$colorMatrix[1,2] = 0.0  # target B
		$colorMatrix[1,3] = 0.0  # target A
		$colorMatrix[1,4] = 0.0  #

		# source B
		$colorMatrix[2,0] = 0.0  # target R
		$colorMatrix[2,1] = 0.0  # target G
		$colorMatrix[2,2] = 0.0  # target B
		$colorMatrix[2,3] = 0.0  # target A
		$colorMatrix[2,4] = 0.0  #

		# source A
		$colorMatrix[3,0] = -1.0 # target R
		$colorMatrix[3,1] = -1.0 # target G
		$colorMatrix[3,2] = -1.0 # target B
		$colorMatrix[3,3] = 0.0  # target A
		$colorMatrix[3,4] = 0.0  #

		# Translate
		$colorMatrix[4,0] = 1.0 # target R
		$colorMatrix[4,1] = 1.0 # target G
		$colorMatrix[4,2] = 1.0 # target B
		$colorMatrix[4,3] = 1.0 # target A
		$colorMatrix[4,4] = 1.0 #

		#
		# Set the matrix
		#
		$attributes = [Drawing.Imaging.ImageAttributes]::new()
		$attributes.SetColorMatrix($colorMatrix)

		#
		# create the new image
		#
		$canvas = [Drawing.Bitmap]::new($Bitmap.Width, $Bitmap.Height, [Drawing.Imaging.PixelFormat]::Format32bppArgb)
		$canvas.SetResolution(96.0, 96.0)

		#
		# Draw into the new image using the color matrix
		#
		$graphics = [Drawing.Graphics]::FromImage($canvas)
		$graphics.DrawImage($Bitmap, [Drawing.Rectangle]::new(0, 0, $Bitmap.Width, $Bitmap.Height), 0, 0, $Bitmap.Width, $Bitmap.Height, [Drawing.GraphicsUnit]::Pixel, $attributes)

		#
		# Now, create the 1-bpp AND MASK
		#
		$andMask = $canvas.Clone([Drawing.Rectangle]::new(0, 0, $Bitmap.Width, $Bitmap.Height), [Drawing.Imaging.PixelFormat]::Format1bppIndexed)

		# debugging
		if ($false)
		{
			$filename = 'alpha_{0}.png' -f $Bitmap.Width
			$filepath = Join-Path -Path $pwd.Path -ChildPath $filename
			$andMask.Save($filepath, [Drawing.Imaging.ImageFormat]::Png)
		}

		return $andMask
	}


	#==================================================================================================================================================================================================
	# Given a Bitmap, create a new square version of it in the specified format and size
	#==================================================================================================================================================================================================
	function New-SquareBitmap
	{
		param([Drawing.Bitmap]$Bitmap, [int]$Dimension, [Drawing.Imaging.PixelFormat]$Format)

		# create the target bitmap
		$canvas = [Drawing.Bitmap]::new($Dimension, $Dimension, $Format)
		$canvas.SetResolution(96.0, 96.0)

		# create and configure the graphics object
		$graphics = [Drawing.Graphics]::FromImage($canvas)
		$graphics.CompositingMode = [Drawing.Drawing2D.CompositingMode]::SourceOver
		$graphics.CompositingQuality = [Drawing.Drawing2D.CompositingQuality]::HighQuality
		$graphics.PixelOffsetMode = [Drawing.Drawing2D.PixelOffsetMode]::HighQuality

		# fill it with ugly
		$graphics.Clear([Drawing.Color]::FromArgb(0,255,0,255))

		# copy over the source, but account for the case where the source image is **NOT** perfectly square
		$w = $Bitmap.Width
		$h = $Bitmap.Height
		$d = [Math]::Max($w, $h)

		$dx = $Dimension * ($d-$w) / (2*$d)
		$dy = $Dimension * ($d-$h) / (2*$d)
		$dw = $Dimension * $w / $d
		$dh = $Dimension * $h / $d

		$graphics.DrawImage($Bitmap, $dx, $dy, $dw, $dh)

		# clean up
		$graphics.Dispose()
		$graphics = $null

		# return the object
		return $canvas
	}


	#==================================================================================================================================================================================================
	# Debugging aid
	#==================================================================================================================================================================================================
	function line
	{
		param([object]$Object)
		Write-Host -Fore White -Back Magenta -NoNewLine ("{0}: {1}" -f $MyInvocation.ScriptName, $MyInvocation.ScriptLineNumber);Write-Host " $Object"
	}


	#==================================================================================================================================================================================================
	# Function to write a struct to a stream
	#==================================================================================================================================================================================================
	function Write-StructToStream
	{
		param([IO.FileStream]$Stream, [object]$Object)

		$size = [Runtime.InteropServices.Marshal]::SizeOf($object)
		$bytes = [byte[]]::new($size)
		$handle = [Runtime.InteropServices.GCHandle]::Alloc($bytes, [Runtime.InteropServices.GCHandleType]::Pinned)
		[Runtime.InteropServices.Marshal]::StructureToPtr($object, $handle.AddrOfPinnedObject(), $false)
		$handle.Free()
		$stream.Write($bytes, 0, $size)

		return $size
	}


	#==================================================================================================================================================================================================
	# Function to write a struct to a buffer
	#==================================================================================================================================================================================================
	function Write-StructToBuffer
	{
		param([byte[]]$Buffer, [int]$Offset, [object]$Object)

		$size = [Runtime.InteropServices.Marshal]::SizeOf($object)
		$bytes = [byte[]]::new($size)
		$handle = [Runtime.InteropServices.GCHandle]::Alloc($bytes, [Runtime.InteropServices.GCHandleType]::Pinned)
		[Runtime.InteropServices.Marshal]::StructureToPtr($object, $handle.AddrOfPinnedObject(), $false)
		$handle.Free()
		[Array]::Copy($bytes, 0, $Buffer, $Offset, $size)

		return $size
	}


	#==================================================================================================================================================================================================
	# Function to read a struct from a buffer
	#==================================================================================================================================================================================================
	function Read-StructFromBuffer
	{
		param([byte[]]$Buffer, [int]$Offset, [Type]$Type)

		$size = [Runtime.InteropServices.Marshal]::SizeOf([Type]$type)
		$bytes = [byte[]]::new($size)
		[Array]::Copy($buffer, $offset, $bytes, 0, $size)
		$handle = [Runtime.InteropServices.GCHandle]::Alloc($bytes, [Runtime.InteropServices.GCHandleType]::Pinned)
		$object = [Runtime.InteropServices.Marshal]::PtrToStructure($handle.AddrOfPinnedObject(), [Type]$Type)
		$handle.Free()

		return $object
	}


	#==================================================================================================================================================================================================
	# Function to read a struct from a stream
	#==================================================================================================================================================================================================
	function Read-StructFromStream
	{
		param([IO.Stream]$Stream, [Type]$Type)

		$size = [Runtime.InteropServices.Marshal]::SizeOf([Type]$type)
		$bytes = [byte[]]::new($size)
		$null = $stream.Read($bytes, 0, $size)
		$handle = [Runtime.InteropServices.GCHandle]::Alloc($bytes, [Runtime.InteropServices.GCHandleType]::Pinned)
		$object = [Runtime.InteropServices.Marshal]::PtrToStructure($handle.AddrOfPinnedObject(), [Type]$Type)
		$handle.Free()

		return $object
	}


	#==================================================================================================================================================================================================
	# Round up to the next multiple of 4
	#==================================================================================================================================================================================================
	function PAD4
	{
		param([uint]$i)
		$i = ($i+3) -band 0xFFFFFFFC
		return $i
	}


	#==================================================================================================================================================================================================
	# Create a mapping of the different Drawing.Imaging formats to how many BPP they use
	#==================================================================================================================================================================================================
	$FormatBppMap = @{
		[Drawing.Imaging.PixelFormat]::Format1bppIndexed	= 1
		[Drawing.Imaging.PixelFormat]::Format4bppIndexed	= 4
		[Drawing.Imaging.PixelFormat]::Format8bppIndexed	= 8
		[Drawing.Imaging.PixelFormat]::Format16bppArgb1555	= 16
		[Drawing.Imaging.PixelFormat]::Format16bppGrayScale	= 16
		[Drawing.Imaging.PixelFormat]::Format16bppRgb555	= 16
		[Drawing.Imaging.PixelFormat]::Format16bppRgb565	= 16
		[Drawing.Imaging.PixelFormat]::Format24bppRgb		= 24
		[Drawing.Imaging.PixelFormat]::Format32bppArgb		= 32
		[Drawing.Imaging.PixelFormat]::Format32bppPArgb		= 32
		[Drawing.Imaging.PixelFormat]::Format32bppRgb		= 32
		[Drawing.Imaging.PixelFormat]::Format48bppRgb		= 48
		[Drawing.Imaging.PixelFormat]::Format64bppArgb		= 64
		[Drawing.Imaging.PixelFormat]::Format64bppPArgb		= 64
	}

	#==================================================================================================================================================================================================
	# Add the types to the runspace
	#==================================================================================================================================================================================================
	Add-Type -TypeDefinition @"
		namespace SpongySoft.Utilities
		{
			using System;
			using System.Runtime.InteropServices;

			[StructLayout(LayoutKind.Sequential, CharSet=CharSet.Ansi, Pack=1)]
			public struct IconDir
			{
				public ushort	idReserved;	// Reserved (must be 0)
				public ushort	idType;		// Resource Type (1 for icons)
				public ushort	idCount;	// How many images?
			}

			[StructLayout(LayoutKind.Sequential, CharSet=CharSet.Ansi, Pack=1)]
			public struct IconDirEntry
			{
				public byte		bWidth;			// Width, in pixels, of the image
				public byte		bHeight;		// Height, in pixels, of the image
				public byte		bColorCount;	// Number of colors in image (0 if >=8bpp)
				public byte		bReserved;		// Reserved ( must be 0)
				public ushort	wPlanes;		// Color Planes
				public ushort	wBitCount;		// Bits per pixel
				public uint		dwBytesInRes;	// How many bytes in this resource?
				public uint		dwImageOffset;	// Where in the file is this image?
			}

			[StructLayout(LayoutKind.Sequential, CharSet=CharSet.Ansi, Pack=1)]
			public struct BITMAPFILEHEADER
			{
				public ushort	bfType;			// 0x4D42, or 'MB'
				public uint		bfSize;
				public ushort	bfReserved1;
				public ushort	bfReserved2;
				public uint		bfOffBits;
			}

			[StructLayout(LayoutKind.Sequential, CharSet=CharSet.Ansi, Pack=1)]
			public struct BITMAPINFOHEADER
			{
				public uint		biSize;
				public int		biWidth;
				public int		biHeight;
				public ushort	biPlanes;
				public ushort	biBitCount;
				public uint		biCompression;
				public uint		biSizeImage;
				public int		biXPelsPerMeter;
				public int		biYPelsPerMeter;
				public uint		biClrUsed;
				public uint		biClrImportant;
			}

			[StructLayout(LayoutKind.Sequential, CharSet=CharSet.Ansi, Pack=1)]
			public struct RGBQUAD
			{
				public byte	b;
				public byte	g;
				public byte	r;
				public byte	a;
			}

			[StructLayout(LayoutKind.Sequential, CharSet=CharSet.Ansi, Pack=1)]
			public struct BITMAPFILE
			{
				public BITMAPFILEHEADER	bmfh;
				public BITMAPINFOHEADER	bmih;
				//public RGBQUAD[]		pal;
			}

			[StructLayout(LayoutKind.Sequential, CharSet=CharSet.Ansi, Pack=1)]
			public struct ICONIMAGE
			{
				public BITMAPINFOHEADER	icHeader;	// DIB header
				//public RGBQUAD[]		icColors;	// Color table
			}
		}
"@


	#==================================================================================================================================================================================================
	# Alter the Formats string array to account for individual strings that have commas
	#==================================================================================================================================================================================================
	$Formats = @($Formats | % {$_ -split ','})


	#==================================================================================================================================================================================================
	# Get the encoders
	#==================================================================================================================================================================================================
	$encoders = [Drawing.Imaging.ImageCodecInfo]::GetImageEncoders()
	$pngEncoder = $encoders | ? {$_.FormatDescription -eq 'PNG' }
	$jpgEncoder = $encoders | ? {$_.FormatDescription -eq 'JPEG' }
	$bmpEncoder = $encoders | ? {$_.FormatDescription -eq 'BMP' }
	$gifEncoder = $encoders | ? {$_.FormatDescription -eq 'GIF' }

	#
	# jpg encoder parameters
	#
	$jpgEncoderParameters = [Drawing.Imaging.EncoderParameters]::new(1)
	$jpgEncoderParameters.Param[0] = [Drawing.Imaging.EncoderParameter]::new([Drawing.Imaging.Encoder]::Quality, 50L)
}

process
{
	#==================================================================================================================================================================================================
	# Read in input
	#==================================================================================================================================================================================================
	if (-not (Test-Path -LiteralPath $SourcePath))
	{
		throw [IO.FileNotFoundException]::new("Source file not found", $SourcePath)
		return
	}

	# resolve it to a full name
	$SourcePath = Resolve-Path -Path $SourcePath | % Path
	$SourceItem = Get-Item -Force -LiteralPath $SourcePath
	$SourcePath = $SourceItem.FullName


	# load the image
	$sourceBitmap = [Drawing.Bitmap]::new($SourcePath)

	#==================================================================================================================================================================================================
	# The target
	#==================================================================================================================================================================================================
	$outputFolder = $null
	$outputFile = $null

	if ($TargetPath)
	{
		if (Test-Path -LiteralPath $TargetPath)
		{
			$TargetItem = Get-Item -Force -LiteralPath $TargetPath

			if ($TargetItem.PsIsContainer)
			{
				$outputFolder = $TargetItem.FullName
			}`
			else
			{
				if ($MyInvocation.ExpectingInput)
				{
					throw [NotSupportedException]::new("You cannot specify a file TargetPath when using pipeline input.")
					return
				}`
				else
				{
					if ($Force)
					{
						# overwrite it
						Remove-Item -Force -LiteralPath $TargetPath -ErrorAction Stop
						$outputFolder = $TargetItem.DirectoryName
						$outputFile = $TargetItem.Name
					}`
					else
					{
						throw [NotSupportedException]::new("Target file ""$TargetPath"" exists. Pass in ""-Force"" to forcibly overwrite it.")
						return
					}
				}
			}
		}`
		else
		{
			# target path was passed in but doesn't exist... if we are processing pipeline input, treat it as a folder,
			# and if not, treat it as a file
			if ($MyInvocation.ExpectingInput)
			{
				$outputFolder = $TargetPath
			}`
			else
			{
				$outputFolder = Split-Path -Path $TargetPath -Parent
				$outputFile = Split-Path -Path $TargetPath -Leaf
			}
		}
	}`
	else
	{
		# target path was not passed in... create one
		$outputFolder = $SourceItem.DirectoryName
	}

	if ($outputFolder)
	{
		if (-not (Test-Path -LiteralPath $outputFolder))
		{
			$null = New-Item -Path $outputFolder -ItemType Directory -Force -ErrorAction Stop
		}
	}`
	else
	{
		$outputFolder = $pwd.Path
	}

	if ($outputFile)
	{
		$outputPath = Join-Path -Path $outputFolder -ChildPath $outputFile
	}`
	else
	{
		$outputFile = '{0}.ico' -f $SourceItem.BaseName
		$outputPath = Join-Path -Path $outputFolder -ChildPath $outputFile

		$index = 0

		while (Test-Path -LiteralPath $outputPath)
		{
			++$index
			$outputFile = '{0}({1}).ico' -f $SourceItem.BaseName, $index
			$outputPath = Join-Path -Path $outputFolder -ChildPath $outputFile
		}
	}

	$len = [Math]::Max($SourcePath.Length, $outputPath.Length) + "Source: ".Length

	# display the outputs
	Write-Host -Fore Gray ("="*$len)
	Write-Host -Fore DarkGray -NoNewLine "Source: ";Write-Host -Fore Blue $SourcePath
	Write-Host -Fore DarkGray -NoNewLine "Target: ";Write-Host -Fore Blue $outputPath
	Write-Host -Fore Gray ("="*$len)


	#==================================================================================================================================================================================================
	# Process the formats
	#==================================================================================================================================================================================================
	$bppToFormatMap = @{
		32 = [Drawing.Imaging.PixelFormat]::Format32bppArgb
		24 = [Drawing.Imaging.PixelFormat]::Format24bppRgb
		16 = [Drawing.Imaging.PixelFormat]::Format16bppArgb1555
		8  = [Drawing.Imaging.PixelFormat]::Format8bppIndexed
		4  = [Drawing.Imaging.PixelFormat]::Format4bppIndexed
	}

	# for debugging
	$i=0

	# the formats
	$formatsToEmbed = [Collections.ArrayList]::new()

	$formatString = $Formats | select -first 1 -skip 4
	foreach ($formatString in $Formats)
	{
		++$i

		# trim leading and trailing whitespace
		$formatString = $formatString.Trim()

		# parse the format string
		if ($formatString -match '^(?<width>\d+)x(?<height>\d+)( (?<bpp>\d+)bpp)?( (?<type>BMP|PNG))?$')
		{
			$width = [int]::Parse($matches.width)
			$height = [int]::Parse($matches.height)
			$type = 'Any'
		}`
		elseif ($formatString -match '^(?<width>\d+)( (?<bpp>\d+)bpp)?( (?<type>BMP|PNG))?$')
		{
			$width = [int]::Parse($matches.width)
			$height = $width
			$type = 'Any'
		}`
		else
		{
			throw [NotSupportedException]::new("Unknown format: ""$formatString"".")
			return
		}

		# if we have a bpp
		if ($matches.bpp)
		{
			$bpp = [int]::Parse($matches.bpp)
		}`
		else
		{
			$bpp = 32
		}

		# if we specified a type
		if ($matches.type)
		{
			$type = $matches.type
		}

		if ($width -ne $height)
		{
			throw [NotSupportedException]::new("Format $i not square: $($width)x$($height)")
			return
		}

		if (@(1,16,32,48,64,128,256) -notcontains $width)
		{
			throw [NotSupportedException]::new("Dimension $width for image $i not supported")
			return
		}

		if (@(4,8,16,24,32) -notcontains $bpp)
		{
			throw [NotSupportedException]::new("Bit depth $bpp for image $i not supported")
			return
		}

		if (-not $bppToFormatMap.Contains($bpp))
		{
			throw [NotSupportedException]::new("Pixel format $formatString ($bpp bits per pixel) not supported")
			return
		}

		$o = [PsCustomObject][ordered]@{
			Dimension = $width
			Format = $bppToFormatMap[$bpp]
			Type = $type
		}

		$null = $formatsToEmbed.Add($o)
	}


	#==================================================================================================================================================================================================
	# Create a list of sizes of the icons we'll need. For each one, we'll create both a 32 bpp ARGB version, and a 1bpp AND mask
	#==================================================================================================================================================================================================
	$dimensions = @($formatsToEmbed | % Dimension | select -Unique | sort -Desc)
	$dimensionMap = @{}
	$dimension = $dimensions | select -first 1
	foreach ($dimension in $dimensions)
	{
		$bitmap = $SourceBitmap

		# create the XOR mask bitmap
		$format = [Drawing.Imaging.PixelFormat]::Format32bppArgb
		$bitmap = New-SquareBitmap -Bitmap $sourceBitmap -Dimension $dimension -Format $format

		# create the AND mask bitmap
		$andMask = New-AndMask -Bitmap $bitmap

		$o = [PsCustomObject][ordered]@{
			FullBitmap = $bitmap
			AndMask = $andMask
		}

		$dimensionMap[$dimension] = $o
	}


	#==================================================================================================================================================================================================
	# Create the list of icon objects
	#==================================================================================================================================================================================================
	$icons = [Collections.ArrayList]::new()

	$Dimensions = $formatToEmbed.Dimension
	$Format = $formatToEmbed.Format
	$Type = $formatToEmbed.Type

	$formatToEmbed = $formatsToEmbed | select -first 1 -skip 0
	foreach ($formatToEmbed in $formatsToEmbed)
	{
		$bitmapObject = $dimensionMap[$formatToEmbed.dimension]
		$dimension = $formatToEmbed.Dimension
		$format = $formatToEmbed.Format
		$type = $formatToEmbed.Type

		#
		# Create the master bitmap
		#
		$bitmap = $bitmapObject.FullBitmap.Clone([Drawing.Rectangle]::new(0, 0, $dimension, $dimension), $format)

		#
		# Create the Icon object
		#
		$icon = [PsCustomObject]::new()
		Add-Member -InputObject $icon -MemberType NoteProperty -Name Bitmap -Value $bitmap -Force
		Add-Member -InputObject $icon -MemberType NoteProperty -Name Format -Value $formatToEmbed -Force

		#
		# Create the PNG
		#
		if (@('Any','PNG') -contains $type)
		{
			$stream = [IO.MemoryStream]::new()
			$bitmap.Save($stream, $pngEncoder, $null)
			#$stream.Close()

			Add-Member -InputObject $icon -MemberType NoteProperty -Name Png -Value $stream -Force
		}

		#
		# Create the BMP
		#
		if (@('Any','BMP') -contains $type)
		{
			$stream = [IO.MemoryStream]::new()
			$bitmap.Save($stream, $bmpEncoder, $null)
			#$stream.Close()

			Add-Member -InputObject $icon -MemberType NoteProperty -Name Bmp -Value $stream -Force
		}

		#
		# Add the icon to the list
		#
		$null = $icons.Add($icon)
	}


	#==================================================================================================================================================================================================
	# Create the ICO file
	#
	# The format of the file is:
	#
	#   1 ICONDIR structure
	#   N ICONDIRENTRY structures
	#   N data blobs
	#==================================================================================================================================================================================================

	#==================================================================================================================================================================================================
	# Go through and calculate the image bits for each icon
	#==================================================================================================================================================================================================
	$icon = $icons | select -first 1
	foreach ($icon in $icons)
	{
		$bmp = $null
		$png = $null

		if (($icon.Format.Type -eq 'PNG') -or ($icon.Format.Type -eq 'Any'))
		{
			$png = [byte[]]::new($icon.Png.Length)
			$icon.Png.Position = 0
			$null = $icon.Png.Read($png, 0, $icon.Png.Length)
		}

		if (($icon.Format.Type -eq 'BMP') -or ($icon.Format.Type -eq 'Any'))
		{
			# reset the position of the BMP file in memory
			$icon.Bmp.Position = 0

			$bitmapFileSize = [Runtime.InteropServices.Marshal]::SizeOf([type][SpongySoft.Utilities.BITMAPFILE])
			$bitmapFileHeaderSize = [Runtime.InteropServices.Marshal]::SizeOf([type][SpongySoft.Utilities.BITMAPFILEHEADER])
			$bitmapInfoHeaderSize = [Runtime.InteropServices.Marshal]::SizeOf([type][SpongySoft.Utilities.BITMAPINFOHEADER])
			$rgbQuadSize = [Runtime.InteropServices.Marshal]::SizeOf([type][SpongySoft.Utilities.RGBQUAD])

			# skip over the BITMAPFILEHEADER to get to the BITMAPINFOHEADER

			$type = [SpongySoft.Utilities.BITMAPFILEHEADER]
			$bmfh = Read-StructFromStream -Stream $icon.Bmp -Type $type
			$type = [SpongySoft.Utilities.BITMAPINFOHEADER]
			$bmih = Read-StructFromStream -Stream $icon.Bmp -Type $type

			# since in icons, the height is twice the width due to the AND mask in addition to the XOR mask
			$bmih.biHeight = $bmih.biWidth * 2

			# And Mask
			if ($true)
			{
				# get the already-made AND mask BMP object for this icon size
				$andMask = $dimensionMap[$icon.Format.Dimension].AndMask

				# create a stream for a BMP file of the AND mask
				$andMaskMemStream = [IO.MemoryStream]::new()

				# write the BMP file of the AND mask
				$andMask.Save($andMaskMemStream, $bmpEncoder, $null)

				# read the headers
				$andMaskMemStream.Position = 0

				$type = [SpongySoft.Utilities.BITMAPFILEHEADER]
				$andMaskBmfh = Read-StructFromStream -Stream $andMaskMemStream -Type $type

				$type = [SpongySoft.Utilities.BITMAPINFOHEADER]
				$andMaskBmih = Read-StructFromStream -Stream $andMaskMemStream -Type $type

				# calculate the size of the AND mask's pixels
				$andMaskMemStream.Position = $bitmapFileSize + $andMaskBmih.biClrUsed * $rgbQuadSize
				$andMaskSize = $andMaskMemStream.Length - $andMaskMemStream.Position
			}

			# Debugging... validate the right size...
			if ($false)
			{
				$actualSize = $icon.Bmp.Length
				$expectedSize = 0
				$expectedSize += $bitmapFileSize
				$expectedSize += $rgbQuadSize * $bmih.biClrUsed
				$expectedSize += $bmih.biWidth * $bmih.biWidth * $bmih.biBitCount / 8

				$calculatedAndMaskSize = $bmih.biWidth * (PAD4 ($bmih.biWidth * 1 / 8))

				$andMaskSize,$calculatedAndMaskSize
				$actualSize,$expectedSize
			}


			# the size is everything else, plus the size of the AND mask
			$size = $icon.Bmp.Length - $bitmapFileHeaderSize + $andMaskSize

			# the BMP object is just a byte array. Since we know how big it is going to be,
			# create it now
			$bmp = [byte[]]::new($size)

			# copy the BITMAPINFOHEADER
			$offset =0
			$offset += Write-StructToBuffer -Buffer $bmp -Offset $offset -Object $bmih

			# copy the palette
			$paletteSize = $bmih.biClrUsed * $rgbQuadSize
			$offset += $icon.Bmp.Read($bmp, $offset, $paletteSize)

			# copy the rest of the bitmap
			$remainingBytes = $icon.Bmp.Length - $icon.Bmp.Position
			$offset += $icon.Bmp.Read($bmp, $offset, $remainingBytes)

			# copy the and mask
			$offset += $andMaskMemStream.Read($bmp, $offset, $andMaskSize)

			$andMaskMemStream.Close()
			$andMaskMemStream.Dispose()
			$andMaskMemStream = $null
		}

		# now, determine which image we will use: BMP or PNG
		$image = $null

		if ($png -and $bmp)
		{
			if ($png.Length -lt $bmp.Length)
			{
				$image = $png
			}`
			else
			{
				$image = $bmp
			}
		}`
		elseif ($png)
		{
			$image = $png
		}`
		elseif ($bmp)
		{
			$image = $bmp
		}`
		else
		{
			throw [Exception]::new("Invalid state!")
			return
		}

		Add-Member -InputObject $icon -MemberType NoteProperty -Name ImageBytes -Value $image -Force
	}


	#
	# Open the output file
	#
	$stream = $null
	$stream = [IO.File]::Open($outputPath, [IO.FileMode]::Create, [IO.FileAccess]::Write, [IO.FileShare]::None)

	#
	# Write the ICONDIR structure
	#
	$iconDir = [SpongySoft.Utilities.IconDir]::new()
	$iconDir.idReserved = 0
	$iconDir.idType = 1
	$iconDir.idCount = $icons.Count

	$null = Write-StructToStream -Stream $stream -Object $iconDir


	#
	# Get the starting offset
	#
	$iconDirSize = [Runtime.InteropServices.Marshal]::SizeOf([SpongySoft.Utilities.IconDir]::new())
	$iconDirEntrySize = [Runtime.InteropServices.Marshal]::SizeOf([SpongySoft.Utilities.IconDirEntry]::new())
	$offset = $iconDirSize + ($iconDirEntrySize * $icons.Count)


	#
	# Write each icon dir entry
	#
	$icon = $icons | select -first 1
	foreach ($icon in $icons)
	{
		$iconDirEntry = [SpongySoft.Utilities.IconDirEntry]::new()

		$iconDirEntry.bWidth			= $icon.Bitmap.Width -band 0xFF
		$iconDirEntry.bHeight			= $icon.Bitmap.Height -band 0xFF
		$iconDirEntry.bColorCount		= $icon.Bitmap.Palette.Entries.Count
		$iconDirEntry.bReserved			= 0
		$iconDirEntry.wPlanes			= 1
		$iconDirEntry.wBitCount			= $FormatBppMap[$icon.Format.Format]
		$iconDirEntry.dwBytesInRes		= $icon.ImageBytes.Length
		$iconDirEntry.dwImageOffset		= $offset

		$null = Write-StructToStream -Stream $stream -Object $iconDirEntry

		$offset += $icon.ImageBytes.Length
	}


	#
	# Write each icon
	#
	$icon = $icons | select -first 1
	foreach ($icon in $icons)
	{
		$stream.Write($icon.ImageBytes, 0, $icon.ImageBytes.Length)
	}


	#
	# Close the file
	#
	$stream.Close()
	$stream.Dispose()
	$stream = $null


	#
	# Cleanup icon resources
	#
	$icons | ? {$_.Png} | % { $_.Png.Close();$_.Png.Dispose();$_.Png=$null }
	$icons | ? {$_.Bmp} | % { $_.Bmp.Close();$_.Bmp.Dispose();$_.Bmp=$null }
	$icons = $null
	[GC]::Collect()
}

end
{
}


