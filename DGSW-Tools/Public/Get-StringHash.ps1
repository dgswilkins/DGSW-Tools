function Get-StringHash {
    <#
    .SYNOPSIS
    Compute a hash of a string using the specified algorithm.
    .DESCRIPTION
    Supports SHA256, SHA512, SHA3_256, and SHA3_512. Assumes that the string is UTF-8 encoded.
    .PARAMETER String
    The input string to hash.
    .PARAMETER HashName
    The hash algorithm to use: SHA256, SHA512, SHA3_256, or SHA3_512.
    .EXAMPLE
    Get-StringHash -String "hello" -HashName SHA256
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [string]$String,
        [Parameter()][ValidateSet('SHA256', 'SHA512', 'SHA3_256', 'SHA3_512')]
        [string]$HashName = 'SHA256'
    )

    begin {
        if ($null -eq $String -or $String -eq '') {
            throw "No string was provided to hash."
        }
        switch ($HashName.ToUpper()) {
            'SHA256' { 
                $algo = [System.Security.Cryptography.SHA256]::Create() 
                Write-Verbose 'Using SHA256 hash algorithm'
            }
            'SHA512' { 
                $algo = [System.Security.Cryptography.SHA512]::Create() 
                Write-Verbose 'Using SHA512 hash algorithm'
            }
            'SHA3_256' {
                try {
                    $algo = [System.Security.Cryptography.SHA3_256]::Create()
                    Write-Verbose 'Using SHA3_256 hash algorithm'
                } catch {
                    throw 'SHA3_256 is not supported on this system. Please install a compatible .NET implementation (8 or later).'
                }
            }
            'SHA3_512' {
                try {
                    $algo = [System.Security.Cryptography.SHA3_512]::Create()
                    Write-Verbose 'Using SHA3_512 hash algorithm'
                } catch {
                    throw 'SHA3_512 is not supported on this system. Please install a compatible .NET implementation (8 or later).'
                }
            }
            default { throw "Unsupported hash algorithm: $HashName" }
        }
    }
    process {
        try {
            $bytes = [System.Text.Encoding]::UTF8.GetBytes($String)
            Write-Verbose "Computing hash for string: [$String]"
            $hash = $algo.ComputeHash($bytes)
            ($hash | ForEach-Object { $_.ToString('x2') }) -join ''
        } catch {
            throw $_
        }
    }
    end {
        if ($algo) { $algo.Dispose() }
    }
}
