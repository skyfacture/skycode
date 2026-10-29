try {
    $swApp = New-Object -ComObject "DSComposer.Application"
    Write-Host "Connected to Composer!"
    # Try every plausible method
    $guesses = @("GetActiveDocument", "ActiveDoc", "GetDocument", "Open", "Document", "Documents", "ActiveDocument", "GetDoc", "OpenDocument")
    foreach ($guess in $guesses) {
        try {
            if ($guess -eq "Open" -or $guess -eq "OpenDocument") {
                $swDoc = $swApp.$guess("C:/Path/To/Your/File.smg")  # Replace with real .smg path
            } elseif ($guess -eq "GetDocument") {
                $swDoc = $swApp.$guess(0)  # Guess doc ID
            } else {
                $swDoc = $swApp.$guess
            }
            if ($swDoc) {
                Write-Host "$guess worked!"
                Write-Host "Views: $($swDoc.ViewCount)"
                Write-Host "File: $($swDoc.FileName)"
                break
            }
        } catch {
            Write-Host "$guess failed: $_"
        }
    }
    # If all fail, dump properties
    Write-Host "Dumping all properties:"
    $swApp | Get-Member -MemberType * | ForEach-Object { Write-Host $_.Name }
} catch {
    Write-Host "Error: $_"
}