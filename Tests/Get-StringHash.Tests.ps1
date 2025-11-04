# Pester tests for Get-StringHash.ps1

# Import the function to test
. "$PSScriptRoot\..\DGSW-Tools\Public\Get-StringHash.ps1"

Describe "Get-StringHash" {
    It "Returns correct SHA256 hash for 'hello'" {
        $result = Get-StringHash -String "hello" -HashName "SHA256"
        $expected = "2cf24dba5fb0a30e26e83b2ac5b9e29e1b161e5c1fa7425e73043362938b9824"
        $result | Should -Be $expected
    }
    It "Returns correct SHA512 hash for 'hello'" {
        $result = Get-StringHash -String "hello" -HashName "SHA512"
        $expected = "9b71d224bd62f3785d96d46ad3ea3d73319bfbc2890caadae2dff72519673ca72323c3d99ba5c11d7c7acc6e14b8c5da0c4663475c2e5c3adef46f73bcdec043"
        $result | Should -Be $expected
    }
    It "Throws for unsupported algorithm" {
        { Get-StringHash -String "hello" -HashName "MD5" } | Should -Throw
    }
    It "Throws for missing string parameter" {
        { Get-StringHash } | Should -Throw
    }
    It "Supports pipeline input" {
        $results = @("hello", "world") | Get-StringHash -HashName "SHA256"
        $expected = @(
            "2cf24dba5fb0a30e26e83b2ac5b9e29e1b161e5c1fa7425e73043362938b9824",
            "486ea46224d1bb4fb680f34f7c9ad96a8f24ec88be73ea8e5a6c65260e9cb8a7"
        )
        $results | Should -Be $expected
    }
    # Optionally test SHA3 algorithms if supported
    It "Returns a hash for SHA3_256 if supported" -Skip:([string]::IsNullOrEmpty((Get-Command Get-StringHash).ScriptBlock.ToString() -match 'SHA3_256')) {
        try {
            $result = Get-StringHash -String "hello" -HashName "SHA3_256"
            $result | Should -Match '^[0-9a-f]{64}$'
        } catch {
            $_ | Should -BeNullOrEmpty
        }
    }
    It "Returns a hash for SHA3_512 if supported" -Skip:([string]::IsNullOrEmpty((Get-Command Get-StringHash).ScriptBlock.ToString() -match 'SHA3_512')) {
        try {
            $result = Get-StringHash -String "hello" -HashName "SHA3_512"
            $result | Should -Match '^[0-9a-f]{128}$'
        } catch {
            $_ | Should -BeNullOrEmpty
        }
    }
}
