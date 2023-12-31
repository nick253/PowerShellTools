# This script is meant to open all the files inside whichever folder that it is ran from.

function MyCommandName() { return $MyInvocation.MyCommand.Name }

foreach ($File in Get-ChildItem($PWD)) {
    # $File.LastWriteTime = $File.CreationTime
    if ($File.Name -ne $MyInvocation.MyCommand.Name){
        Write-Output($File.Name)
        # Write-Output($PSCommandPath)
        # Write-Output($MyInvocation.MyCommand.Name)
        # Write-Output($MyInvocation.ScriptName)
        Start-Process -FilePath $File # -WorkingDirectory "E:\New folder\New folder\New folder (2)"
    }
}
Write-Output("Complete")
