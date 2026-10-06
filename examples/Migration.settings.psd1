# Optional. Put this file next to Invoke-ScomMigrationStep.ps1 so nobody has to
# type these on every step. Anything given on the command line wins.
@{
    # Target management server the import/test steps connect to.
    ManagementServer = 'SCOMMS01.contoso.local'

    # ExportSource: the OLD management server. The export runs there through
    # PowerShell remoting and the result lands in SourceFolder here.
    SourceServer     = 'OLDSCOM01.contoso.local'

    # ExportSource: folders or shares with original .mp/.mpb files,
    # searched from THIS server (optional).
    SealedSearchPath = @('\\fileserver.contoso.local\SCOM\MPs')

    # Folder names (relative to the script) or absolute paths.
    SourceFolder     = 'Source'      # whole output folder of Export-ScomEnvironment.ps1 -Role Source
    TargetFolder     = 'Target'      # must contain AllMPs\

    # Labels used in messages and the Import-Batch.<TargetVersion>.ps1 file name.
    # In v3.x these do not change any compile logic.
    SourceVersion    = '2016'
    TargetVersion    = '2025'
}
