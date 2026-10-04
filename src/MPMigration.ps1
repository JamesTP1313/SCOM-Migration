<#
.SYNOPSIS
    SCOM Management Pack Batch Migration Compiler / Transpiler.

.DESCRIPTION
    Analyzes a BATCH of source SCOM Management Packs (e.g. exported from a
    SCOM 2016 environment, a mix of unsealed .xml and sealed .mp files)
    against a target repository of installed Management Packs (e.g. a SCOM
    2025 management server) and produces:

      - A symbol table of every Class, Rule, Monitor, Discovery, and Task
        known to the target repository.
      - A batch-internal dependency graph (which source MPs reference other
        source MPs in the same batch).
      - A topological import order for the whole batch, so MPs that depend
        on each other are imported in a safe sequence (with cycle detection
        -- a real, fatal SCOM condition, reported rather than silently
        "ordered").
      - Per-MP: validation of all Overrides against the target symbol table,
        a reference rewrite table mapping old (source-version) MP references
        to their best-match equivalent in the target repository, a candidate
        output MP with references rewritten/updated, and a migration
        advisory report scoring overall compatibility.
      - A combined batch summary: ordered import manifest (CSV) + a generated
        PowerShell snippet using Import-SCOMManagementPack in the correct
        order, ready for you to review and run by hand. This script never
        connects to a live SCOM server and never imports anything itself.

.PARAMETER InputPath
    One or more paths to source Management Pack files (.xml or .mp), and/or
    one or more folders containing them (each folder is searched
    recursively). Accepts a single path, a comma-separated list of paths,
    or an array. If omitted, a dialog explicitly asks whether you want to
    pick a folder or pick one/many individual files.

.PARAMETER RepositoryFolder
    Path to the folder containing the target environment's installed/sealed MPs
    (the version you are migrating TO). No default -- always prompted for if
    not supplied, since silently defaulting to the local SCOM cache path
    previously caused this to point at an incomplete repository without any
    visible confirmation. A folder with fewer than ~80 MPs triggers a warning,
    since a real SCOM install (even freshly installed) normally has 100+.

.PARAMETER SourceRepositoryFolder
    Optional. Path to a folder (e.g. a network share) containing MPs exported
    from the SOURCE environment (e.g. SCOM 2016) -- typically a broader set
    than just your -InputPath batch, such as a full export of everything
    installed on the old management group. When a reference can't be
    resolved against the batch or the target repository, this folder is
    searched as a fallback: if found here, that dependency MP is pulled into
    the batch, run through the SAME validation/rewrite pipeline as every
    other batch MP (it is NOT trusted blindly or copied in as-is), and only
    then counted as resolved. References still not found anywhere end up in
    MissingDependencies.csv as before. If omitted (and -SourceManagementServer
    is also omitted), this fallback is skipped and unresolved references are
    reported exactly as in prior versions.

    Mutually exclusive with -SourceManagementServer -- if you supply both on
    the command line, the script throws rather than silently picking one, so
    you don't end up wondering which source it actually used. Use whichever
    one you actually have: a folder if someone already exported one for you,
    or -SourceManagementServer if you have live access to the old
    environment and would rather this script pull a fresh export itself.

.PARAMETER SourceManagementServer
    Optional. SCOM management server name for the SOURCE (OLD) environment,
    e.g. a SCOM 2016 management server -- an alternative to
    -SourceRepositoryFolder for anyone who has live access to both the old
    and new management groups (the common case: if you're migrating MPs,
    you almost certainly still have the old environment reachable). When
    supplied, the script connects to it directly (same mechanism as
    -LiveImport's connection to the target), runs Get-SCOMManagementPack to
    enumerate every MP currently installed there, and exports all of them
    with Export-SCOMManagementPack to a fresh, timestamped folder under
    -OutputFolder. That export then feeds the exact same dependency
    auto-resolution pipeline -SourceRepositoryFolder always has -- nothing
    downstream changes; this only replaces "browse to a folder someone
    exported earlier" with "pull a live, current export of the old
    environment right now."

    This is READ-ONLY against the source environment -- it only enumerates
    and exports (Get-SCOMManagementPack / Export-SCOMManagementPack), never
    writes to it, and is entirely separate from -LiveImport (which writes to
    the TARGET). You can use -SourceManagementServer with or without
    -LiveImport; they connect to different management groups and are never
    connected to at the same time.

    Requires the OperationsManager PowerShell module to be able to connect
    to the source management group's version. If your console/module is
    tied to the NEW (target) SCOM version and can't reach the older source
    management group from the same session, run this script twice instead:
    once from a machine that can reach the source (with just
    -SourceManagementServer, or export manually and use
    -SourceRepositoryFolder), and once for the actual migration -- or fall
    back to a manual Export-SCOMManagementPack + -SourceRepositoryFolder,
    which always works regardless of module/version constraints.

    If the connection or enumeration fails outright, the script throws
    immediately (same philosophy as -RepositoryFolder / -LiveImport: no
    silent partial fallback) -- rerun with -SourceRepositoryFolder pointing
    at a manual export instead. If SOME MPs export successfully and others
    individually fail (e.g. one locked/corrupt MP on the source side), the
    run continues with whatever exported successfully, and every failure is
    logged by name so you know exactly what's missing from auto-resolution.

.PARAMETER SourceCredential
    Optional. PSCredential to use when connecting to -SourceManagementServer,
    for cases where the old environment is in a different domain/forest or
    otherwise needs different credentials than your current session. Ignored
    if -SourceManagementServer is not supplied. If -SourceManagementServer is
    supplied without this, the connection uses your current session's
    credentials, same as -LiveImport's connection to the target does today.

.PARAMETER OutputFolder
    Path where the candidate MPs and reports will be written. Defaults to
    Documents\SCOMCompiler under the current user profile.

.PARAMETER SourceVersion
    Label for the SCOM version the InputPath batch was authored against
    (e.g. "2016"). Used only for logging/reporting context; does not change
    matching logic.

.PARAMETER TargetVersion
    Label for the SCOM version of the RepositoryFolder (e.g. "2025"). Used only
    for logging/reporting context.

.PARAMETER AllowIDRewrite
    By default, the rewrite engine only updates reference VERSIONS, never IDs,
    because rewriting an MP ID is a higher-risk change (it changes which MP is
    actually referenced, not just which version of it). Pass this switch to
    allow ID rewrites for references matched via $ReferencePatterns.

.PARAMETER Strict
    Treat any ERROR-severity diagnostic (e.g. a missing override reference, or
    a dependency cycle in the batch) as fatal for that MP / for the batch
    ordering step. Without -Strict, the script proceeds and writes output
    anyway, with all errors/warnings shown in the console and reports.

.PARAMETER SkipSealedExtraction
    Skip attempting to load the SCOM SDK assemblies and extract XML from
    sealed .mp files. Any .mp files found will be logged and skipped rather
    than processed. Use this only if you know SDK extraction will fail and
    want to avoid the noise; normally leave this off.

.PARAMETER AuditOnly
    Changes what "compatible" means, for the case where -InputPath is an
    entire source environment (e.g. every MP exported from SCOM 2016) rather
    than a small batch you're actively migrating. Normally, if MP A
    references MP B and B is ALSO somewhere in your -InputPath batch, that
    reference is considered resolved ("BATCH_INTERNAL") without checking
    whether B is actually compatible with -RepositoryFolder (the target) --
    which is the right behavior for a real migration batch, but produces a
    misleadingly clean "100% compatible" result when the batch IS the whole
    source environment, since old MPs are trivially "compatible" with other
    old MPs from the same era. With -AuditOnly, that BATCH_INTERNAL shortcut
    is skipped entirely for scoring/missing-dependency purposes: every
    single reference is checked directly against -RepositoryFolder, so the
    result reflects genuine target compatibility, not internal self-
    consistency. Does not affect candidate file generation, the dependency
    graph, or import ordering -- only the advisory/scoring verdict and what
    lands in MissingDependencies.csv.

.PARAMETER LiveImport
    OPT-IN, OFF BY DEFAULT. Without this switch, the script behaves exactly
    as before: read-only, file-prep only, never touches a live SCOM server.

    With -LiveImport, the script connects to a real SCOM management group
    (see -ManagementServer) and ACTUALLY IMPORTS every auto-resolved
    DEPENDENCY (anything pulled in via -SourceRepositoryFolder that is NOT
    one of the MP(s) you explicitly named in -InputPath), at its genuine old
    version, exactly as exported from the source environment -- no rewrite
    applied. You are prompted to confirm before each individual live import.

    If you decline a prompt, or an import genuinely fails, that ONE
    dependency (and anything that depends on it) is skipped, and the run
    CONTINUES with everything else -- you get a complete picture of every
    dependency's outcome in one pass, instead of having to fix or accept
    one problem before discovering the next. A full summary (imported /
    already present / declined / failed / skipped-by-cascade) is shown at
    the end. Declines and failures are also saved to LiveImportDecisions.json
    in -OutputFolder, so a FUTURE run never re-prompts for or re-attempts
    something already known not to work -- delete the relevant entry from
    that file if you want to retry it later.

    The MP(s) you explicitly named in -InputPath are still written as
    candidate files to CandidateMPs\ exactly as before, in every case. In
    addition, if -LiveImport is set AND a target's own compatibility score
    reaches 95% or higher (the same threshold the Bottom Line summary calls
    "READY"), that candidate is ALSO automatically imported into the
    connected management group -- no separate confirmation prompt, since
    passing -LiveImport at all is taken as that confirmation. A target
    already present at a matching version is detected and skipped, same as
    dependencies. If the score is below 95%, or the import is attempted and
    SCOM rejects it anyway, nothing is silently assumed -- the Bottom Line
    summary states plainly what happened and the candidate file remains for
    manual review either way. This switch automates getting DEPENDENCIES in
    place first, and now also the target itself once it looks genuinely
    ready -- not a guarantee, since a 95%+ score is "every reference we
    could check resolved," not a guarantee SCOM will accept the file.

    Dependencies already present in SCOM at a matching version are detected
    via Get-SCOMManagementPack and skipped silently -- no prompt, no
    re-import -- so re-running after a partial success doesn't re-ask about
    things already settled.

.PARAMETER ManagementServer
    SCOM management server to connect to when -LiveImport is set. Defaults
    to "localhost" (i.e. you're running this script ON the management
    server itself). Prompted for if -LiveImport is set and this is not
    supplied. Ignored entirely if -LiveImport is not set.

.PARAMETER ExcludeMPIDs
    One or more MP IDs to treat as a KNOWN, INTENTIONAL gap rather than
    something to search for or prompt about. Useful once you've already
    confirmed (e.g. via earlier testing) that a particular legacy dependency
    genuinely cannot be brought forward -- its required version doesn't
    exist anywhere you have access to, or it's confirmed not needed -- and
    you don't want to keep re-discovering and re-declining the same prompt
    on every run.

    An excluded ID is skipped everywhere, consistently:
      - Step 2.5 auto-resolution never searches -SourceRepositoryFolder for it
      - -LiveImport never prompts for or attempts to import it
      - It is never pulled into the batch, so nothing downstream (override
        validation, the dependency graph, candidate rewriting) processes it
      - It still appears in MissingDependencies.csv, but explicitly marked
        as excluded-by-request rather than genuinely unresolved -- so the
        report stays honest about what's actually missing vs. what you've
        already decided not to chase.

    Excluding an MP does NOT remove the need for whatever depends on it --
    anything that required the excluded MP will still show that dependency
    as unmet. This only stops the script from re-asking about something
    you've already decided on.

.PARAMETER ForceRewriteCoreLibraries
    OPT-IN. OFF BY DEFAULT, AND FOR GOOD REASON -- READ THIS BEFORE USING IT.

    By default, references to ExactMatchOnly families (CoreLibrary,
    SystemCenter, IISCommonLibrary -- the foundational libraries every other
    MP builds on) are NEVER auto-rewritten to a newer version, even when one
    is sitting right there in -RepositoryFolder. This is deliberate: an
    earlier version of this tool DID auto-bump these, and it produced a
    candidate MP that looked clean (100% compatible, no errors) but that
    SCOM rejected outright on actual import, because the newer core library
    version was not actually compatible with what the old MP's logic
    expected. That failure mode is exactly what this default avoids.

    With -ForceRewriteCoreLibraries, that safeguard is turned OFF: the
    rewrite engine will bump CoreLibrary/SystemCenter/IISCommonLibrary
    references to the newest version found in -RepositoryFolder, the same
    way it already does for families that are known to unify safely (IIS,
    SQL, Windows Server, etc.). This trades a known, real risk (the
    rewritten candidate might still fail on actual import, for reasons this
    tool cannot detect ahead of time) for convenience on simpler MPs -- e.g.
    a small custom .Overrides MP whose only use of a core library reference
    is declarative, not deep custom logic -- where that risk is genuinely
    low. It is much riskier for MPs with substantial custom monitoring
    logic built against the OLD library's schema.

    Every reference rewritten under this switch is logged distinctly and
    flagged in the rewrite report, specifically so you always know which
    parts of a candidate file relied on this looser rule -- never silent.

    This does not guarantee a successful import. Test the result before
    trusting it, the same way you would any other rewritten candidate.

.PARAMETER CheckCatalog
    Runs a new Step 1.5 right after -RepositoryFolder is loaded: fetches
    Microsoft's official Management Pack catalog (the "Microsoft Management
    Packs" reference list on Microsoft Learn, filtered to -CatalogView) and
    fuzzy-matches every catalog entry against what's already in your
    -RepositoryFolder. Writes MPCatalog.csv to -OutputFolder listing every
    catalog entry, whether a likely local match was found, and the Download
    Center link -- this is how you find out what's available for SCOM 2025
    without hunting the Download Center by hand, and it's catalog logic
    folded in here so it runs as part of your normal migration workflow
    instead of a separate step.

    Matching is by normalized display-name token overlap, NOT by MP ID --
    the catalog page only has display names, not IDs. Treat the CSV's
    "Likely Match" / "No Local Match Found" column as a strong hint to
    eyeball, not a verdict; -CatalogMatchThreshold controls the cutoff.

    This step is purely informational and never blocks the rest of the run
    -- migration proceeds normally afterward unless -CatalogOnly is also
    set. Requires network access to learn.microsoft.com (and, if
    -DownloadFromCatalog is also used, microsoft.com) from this session;
    if that's not available from your SCOM server, run it from a machine
    that does have that access, or skip this switch entirely and keep using
    -RepositoryFolder the way you already do.

.PARAMETER CatalogOnly
    Implies -CheckCatalog, and stops the script right after that step
    completes -- no batch is loaded or processed, and -InputPath is not
    required. Use this when all you want is "what's available from
    Microsoft for my SCOM 2025 repository," independent of any specific
    migration. Same relationship to -CheckCatalog as -AuditOnly has to a
    normal run: a narrower, self-contained mode.

.PARAMETER DownloadFromCatalog
    Experimental. Only meaningful combined with -CheckCatalog or
    -CatalogOnly. For every catalog entry with "No Local Match Found" (or
    every entry, if -RepositoryFolder somehow indexed zero MPs), attempts to
    resolve the actual installer URL from its Download Center page and
    downloads it to -OutputFolder\MPCatalog\Downloads. Best-effort:
    Microsoft's Download Center page layout has changed before and may not
    match the patterns this script looks for by the time you run it --
    failures are logged with the Download Center URL so you can grab the
    file by hand instead. Never extracts or imports anything; downloaded
    files are installers only, same trust boundary as everywhere else in
    this script that stops short of touching content it hasn't validated.

.PARAMETER CatalogView
    Which SCOM version's catalog page to pull for -CheckCatalog / 
    -CatalogOnly, matching the `?view=` query parameter on the Microsoft
    Learn page. Default: sc-om-2025. The underlying page covers sc-om-2016
    through sc-om-2025 with the SAME table regardless of which one you
    pick -- this mainly matters for which docs-site version silo you're
    citing, not for materially different MP content.

.PARAMETER CatalogFilter
    Optional array of keywords to narrow the catalog check to specific MP
    families, e.g. -CatalogFilter "IIS","SQL Server","Windows Server". If
    omitted, -CheckCatalog / -CatalogOnly checks the entire Microsoft
    catalog (100+ entries), which is slower and produces a much longer CSV.

.PARAMETER CatalogMatchThreshold
    Local-match confidence cutoff (0.0-1.0) for -CheckCatalog / -CatalogOnly,
    based on normalized token overlap between the catalog's display name and
    each repository MP's display name. Default: 0.5. Lower this if real
    matches are showing up as "No Local Match Found"; raise it if you're
    seeing false "Likely Match" hits.

.PARAMETER CatalogStrictMatch
    OFF by default. Changes what counts as a local match in -CheckCatalog /
    -CatalogOnly, for the specific case where you'd rather over-flag gaps
    than soft-match to a family cousin.

    By default (soft), a catalog entry counts as a "Likely Match" if it
    shares enough distinctive tokens with ANY local MP -- so all the SQL
    Server sub-feature entries (Reporting Services, Analysis Services,
    Replication, Dashboards) will match your generic SQLServer library even
    if you don't have those specific per-feature MPs. That's reassuring
    ("you have SQL monitoring") but it can hide a real gap.

    With -CatalogStrictMatch, a match additionally requires that the catalog
    entry's OWN distinctive family/component word actually appears in the
    local MP -- e.g. "SQL Server Reporting Services" only matches a local MP
    whose name/ID actually contains "reporting", not merely "sql"+"server".
    Entries that only soft-matched a generic library flip to "No Local Match
    Found" instead, landing them in the actionable bucket (with a Download
    Center link) rather than the reassurance bucket.

    Recommended when the point of the run is building a shopping list of MPs
    you still need to obtain for SCOM 2025: over-flagging ("go confirm you
    have this") is the safer failure direction than a soft match that waves
    a genuine gap through. Note this ONLY affects the human-facing catalog
    CSV -- it has no effect on actual MP migration, whose dependency
    resolution always uses exact reference-ID lookups against the
    repository, never this fuzzy catalog matcher.

.PARAMETER ResolveStaticGroupMembers
    Opt-in. When an MP contains STATIC group membership -- an <IncludeList> of
    explicit <MonitoringObjectId> GUIDs (the "hand-jammed the servers in"
    pattern) -- those GUIDs are object instance IDs from the SOURCE (old)
    environment. They are environment-specific: the SAME server, once
    discovered by the NEW management group, gets a DIFFERENT GUID. So a static-
    membership group migrated as-is imports cleanly, scores 100%, and then
    populates EMPTY in the target -- every dashboard/override targeting it
    silently monitors nothing. This is caught and warned about always (see
    -SkipStaticMembershipCheck to suppress).

    With -ResolveStaticGroupMembers AND a source connection
    (-SourceManagementServer), the script goes further: it connects to the OLD
    environment and resolves each membership GUID back to the server's real
    name (FQDN / DisplayName / PrincipalName) by looking up the object. It
    writes StaticGroupMembership.csv to -OutputFolder listing, per group, the
    named servers that were members -- turning ~hundreds of opaque, dead-on-
    arrival GUIDs into a human-readable membership list you can act on. This is
    READ-ONLY against the source (Get-SCOMClassInstance / Get-SCOMMonitoringObject
    only) and does not modify the candidate MP.

    Deliberately does NOT rewrite the static GUID list into new-environment
    GUIDs, even though that's technically possible once servers are dual-homed.
    Reproducing a hand-maintained static list just carries the fragility
    forward (it drifts the moment someone adds a server). The resolved name
    list is far more useful as the basis for a DYNAMIC membership rule in 2025,
    or as a validation set ("did my new dynamic group capture the same
    servers?"). Converting static->dynamic is a deliberate authoring decision,
    not something this tool should do silently.

    Requires -SourceManagementServer (to resolve GUIDs against the old env). If
    -ResolveStaticGroupMembers is set without a source connection, the static-
    membership warning still fires but names can't be resolved, and the script
    says so rather than guessing.

.PARAMETER SkipStaticMembershipCheck
    Suppresses the static-group-membership detection warning entirely. Off by
    default (the check always runs) because a static-membership group that
    imports empty is exactly the kind of silent failure this tool exists to
    surface. Only set this if you've already accounted for every static group
    and don't want the warnings in your log.

.PARAMETER Manifest
    Optional. Path to MigrationManifest.csv (built from the disposition
    workbook). Only rows with Migrate = Y are kept from -InputPath; every
    other MP found under -InputPath is ignored. Matching is on the
    MatchPattern column (wildcards allowed, e.g. "Contoso.App*.Monitoring")
    against the MP ID first, then its display name. ManifestMatch.csv in
    -OutputFolder shows which rows matched which MP, and which matched nothing.

.PARAMETER SourceInventory
    Optional. Path to SourceInventory.csv written by Export-ScomEnvironment.ps1
    (or by -SourceManagementServer). Used to tell which source MPs were SEALED
    in the source. An exported .xml of a sealed MP can never be imported in its
    place (the signature is gone, so every reference to it fails), so
    those targets are marked BLOCKED until the original .mp/.mpb is supplied.

.PARAMETER NonInteractive
    No dialogs or prompts at all. Missing required paths throw instead of
    opening a picker, the low-repository-count check warns instead of asking,
    and -LiveImport dependency confirmations are auto-approved. Use this for
    unattended or remote runs.

.PARAMETER GroupConversionFile
    Optional. GroupConversion.csv from Export-ScomEnvironment.ps1 -Role Source.
    Every static group membership rule (explicit source object GUIDs, which
    would import EMPTY) whose row has Convert = Y is rewritten into a DYNAMIC
    rule on Microsoft.Windows.Computer NetbiosComputerName using the row's
    Operator/Pattern (hosted classes match on the host computer's NetBIOS name
    via HostProperty -- the same XML the console group wizard writes, so the
    groups stay editable in the console). ExcludeList GUIDs become a DoesNotMatchRegularExpression on
    the excluded names. Nested static subgroups (member class is a singleton
    group) are converted automatically, with or without this file.
    Every change is listed in GroupConversionResults.csv.

.PARAMETER AutoApproveDependencies
    With -LiveImport, import sealed dependency MPs without a per-MP prompt.

.NOTES
    Author : JT Perry
    Version: 3.40
    License: MIT (see LICENSE)
    History: see CHANGELOG.md

    Each "Step" in this script is a self-contained phase that prints its own
    banner. Step 0 ingests the batch (resolving .mp -> .xml). Steps 1-2 build
    context (repository + symbol table). Step 1.5 (new in 3.32, ONLY runs
    with -CheckCatalog or -CatalogOnly) fetches Microsoft's official MP
    catalog and gap-checks it against -RepositoryFolder -- purely
    informational, never blocks the rest of the run, except that
    -CatalogOnly stops the script right after this step. Step 2.5 (new in
    3.1) attempts to auto-resolve any batch reference that isn't satisfied by
    the batch or the target repository, by searching -SourceRepositoryFolder
    and pulling any match into the batch for full processing. Step 3 builds
    the batch-internal dependency graph and topological import order over the
    now-possibly-larger batch. Step 3.5 (new in 3.14, ONLY runs with
    -LiveImport) connects to a real SCOM management group and imports each
    auto-resolved DEPENDENCY live, with per-MP confirmation and a hard stop
    on any failure. Steps 4-8 run PER MP in that order: override validation,
    dependency graph vs. target repo, rewrite table, candidate emission, and
    advisory reporting. A final Step 9 rolls everything up into one batch
    manifest and PowerShell import snippet.

    By default (no -LiveImport), this script never connects to a live SCOM
    management server and never imports anything itself -- it only prepares
    and validates candidate MP files plus an ordered plan for you to execute
    by hand. -LiveImport is an explicit, off-by-default opt-in to a different
    mode where dependency imports happen for real, with confirmation at each
    step -- the MP(s) you're actually migrating are still never auto-imported
    either way.
#>

[CmdletBinding()]
param(
    [string[]]$InputPath,

    [string]$RepositoryFolder,

    [string]$SourceRepositoryFolder,

    [string]$SourceManagementServer,

    [System.Management.Automation.PSCredential]$SourceCredential,

    [string]$OutputFolder = "$env:USERPROFILE\Documents\SCOMCompiler",

    [string]$SourceVersion = "2016",

    [string]$TargetVersion = "2025",

    [switch]$AllowIDRewrite,

    [switch]$Strict,

    [switch]$SkipSealedExtraction,

    [switch]$AuditOnly,

    [switch]$LiveImport,

    [string]$ManagementServer,

    [string[]]$ExcludeMPIDs,

    [switch]$ForceRewriteCoreLibraries,

    [switch]$CheckCatalog,

    [switch]$CatalogOnly,

    [switch]$DownloadFromCatalog,

    [string]$CatalogView = 'sc-om-2025',

    [string[]]$CatalogFilter,

    [double]$CatalogMatchThreshold = 0.5,

    [switch]$CatalogStrictMatch,

    [switch]$ResolveStaticGroupMembers,

    [switch]$SkipStaticMembershipCheck,

    [string]$Manifest,

    [string]$SourceInventory,

    [switch]$NonInteractive,

    [switch]$AutoApproveDependencies,

    [string]$GroupConversionFile,

    # By default a group whose membership rule depends on an MP that is not
    # in the target (e.g. a group of a third-party MP's objects) is DROPPED from
    # its MP -- with everything that points at it -- so the rest of the MP
    # can import. Set this to keep the old behaviour (whole MP blocked).
    [switch]$KeepBlockedGroups,

    # InstanceMap.csv from Invoke-ScomMigrationStep.ps1 MapInstances: source object
    # GUID -> target object GUID for overrides aimed at one specific object
    # (ContextInstance). Mapped overrides are rewritten; unmapped ones removed.
    [string]$InstanceMapFile
)

#Requires -Version 5.1

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$script:BuildLabel = "v3.40 (build 2026-09-30)"

###########################################################
# LOAD UI LIBRARY
###########################################################

# WinForms is only needed for the pickers/prompts. Under -NonInteractive (or
# on a host without it) the script runs entirely from parameters.
$script:UIAvailable = $false
if (-not $NonInteractive) {
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        $script:UIAvailable = $true
    }
    catch {
        Write-Host "System.Windows.Forms is not available on this host -- running non-interactively." -ForegroundColor Yellow
        $NonInteractive = [switch]$true
    }
}

###########################################################
# XML LOADING (encoding-safe)
###########################################################
# [xml](Get-Content path) in Windows PowerShell 5.1 decodes BOM-less files as
# the ANSI code page, so any non-ASCII character in an exported MP (curly
# quotes in knowledge articles, accented names) is corrupted and then written
# back out corrupted. XmlDocument.Load() honours the file's own XML
# declaration / BOM instead, which is what SCOM itself does.
function Read-MPXml {
    param([Parameter(Mandatory)][string]$Path)
    $doc = New-Object System.Xml.XmlDocument
    $doc.PreserveWhitespace = $false
    $doc.Load($Path)
    return , $doc
}

# All element IDs defined anywhere in an MP document (classes, relationships,
# monitors, rules, discoveries, tasks, diagnostics, recoveries, views,
# folders, overrides, data sources, ...). Used to check that an override or
# category still points at something that exists in the version of the MP
# that will actually be installed in the target.
function Get-MPElementIdSet {
    param([Parameter(Mandatory)][System.Xml.XmlDocument]$Doc)
    $set = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::Ordinal)
    $nodes = $Doc.SelectNodes('//*[@ID]')
    foreach ($n in $nodes) {
        if ($n.LocalName -eq 'Reference') { continue }
        [void]$set.Add([string]$n.GetAttribute('ID'))
    }
    # The MP itself is a valid target (e.g. <Category Target="<this MP>">).
    $selfId = $Doc.SelectSingleNode('/ManagementPack/Manifest/Identity/ID')
    if ($selfId -and $selfId.InnerText) { [void]$set.Add([string]$selfId.InnerText) }
    return , $set
}

###########################################################
# UI HELPERS
###########################################################

function Get-TopMostOwner {
    # OpenFileDialog/FolderBrowserDialog/MessageBox shown via .ShowDialog()
    # or .Show() with NO owner have no guaranteed Z-order relative to other
    # app windows -- they can silently open BEHIND the console/ISE window,
    # which looks like "the script just hung." Giving every dialog this
    # invisible, always-on-top owner forces it to the foreground reliably,
    # regardless of host (console, ISE, VS Code terminal).
    $owner = New-Object System.Windows.Forms.Form
    $owner.TopMost = $true
    $owner.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $owner.ShowInTaskbar = $false
    $owner.Size = New-Object System.Drawing.Size(0, 0)
    $owner.Opacity = 0
    $owner.Show()
    $owner.Activate()
    return $owner
}

function Select-File {
    param(
        [string]$Title = "Select Management Pack",
        [string]$Filter = "Management Pack files (*.xml;*.mp)|*.xml;*.mp|XML files (*.xml)|*.xml|Sealed MP files (*.mp)|*.mp|All files (*.*)|*.*",
        [switch]$Multiselect
    )

    $owner = Get-TopMostOwner

    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = $Title
    $dialog.Filter = $Filter
    $dialog.Multiselect = $Multiselect.IsPresent

    try {
        if ($dialog.ShowDialog($owner) -eq [System.Windows.Forms.DialogResult]::OK) {
            if ($Multiselect) {
                return @($dialog.FileNames)
            }
            return $dialog.FileName
        }
        return $null
    }
    finally {
        $owner.Close()
        $owner.Dispose()
    }
}

function Select-Folder {
    param(
        [string]$Description = "Select Folder"
    )

    # NOTE: unlike OpenFileDialog, calling FolderBrowserDialog.ShowDialog($owner)
    # with an explicit IWin32Window owner is brittle under PowerShell's method
    # overload binder and can throw "Argument types do not match" even though
    # the same call is valid C#. Every working PowerShell example of this
    # dialog calls ShowDialog() with NO arguments. To still get the dialog to
    # the foreground (the actual goal), activate the invisible topmost owner
    # form first, then show the dialog without passing it as an argument --
    # this reliably brings the dialog forward without touching the brittle
    # overload at all.
    $owner = Get-TopMostOwner

    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = $Description

    try {
        $owner.Activate()
        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            return $dialog.SelectedPath
        }
        return $null
    }
    finally {
        $owner.Close()
        $owner.Dispose()
    }
}

function Select-InputPathOrPaths {
    # Explicitly asks File(s) vs. Folder instead of overloading "Cancel" as
    # a hidden second meaning -- the previous design (Cancel the file
    # dialog to get a folder dialog instead) was confusing in both ISE and
    # console hosts, since cancelling looks identical to "I changed my mind
    # and want nothing" right up until the folder dialog unexpectedly pops
    # up next.
    #
    # Returns either a single path string (folder or one file) or a string
    # array (multiple files), matching whatever Step 0's ingestion logic
    # already knows how to handle -- both Get-Item and Get-ChildItem work
    # fine against either shape downstream.

    $owner = Get-TopMostOwner

    try {
        $choice = [System.Windows.Forms.MessageBox]::Show(
            $owner,
            "Do you want to select a FOLDER of Management Packs (recommended for a batch),`nor pick one or more individual MP files?`n`nYes  = Pick a FOLDER`nNo   = Pick FILE(S)",
            "Select Source Input",
            [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
    }
    finally {
        $owner.Close()
        $owner.Dispose()
    }

    switch ($choice) {
        "Yes" {
            return Select-Folder "Select Folder of Source Management Packs"
        }
        "No" {
            return Select-File -Title "Select Source Management Pack file(s) (Ctrl/Shift-click for multiple)" -Multiselect
        }
        default {
            return $null
        }
    }
}

###########################################################
# LOGGING HELPER
###########################################################
# Centralizes console + transcript output so nothing silently
# vanishes into a bare catch{} block anymore. $script:LogFile is
# set once the OutputFolder is confirmed to exist (see INIT below).

$script:LogFile = $null

function Write-Log {
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Message,
        [ValidateSet("INFO", "WARN", "ERROR", "SUCCESS")][string]$Level = "INFO",
        [switch]$NoConsole
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$timestamp] [$Level] $Message"

    if (-not $NoConsole) {
        switch ($Level) {
            "WARN"    { Write-Host $Message -ForegroundColor Yellow }
            "ERROR"   { Write-Host $Message -ForegroundColor Red }
            "SUCCESS" { Write-Host $Message -ForegroundColor Green }
            default   { Write-Host $Message }
        }
    }

    if ($script:LogFile) {
        Add-Content -LiteralPath $script:LogFile -Value $line -Encoding UTF8
    }
}

function Write-Banner {
    param([Parameter(Mandatory)][string]$Text)

    Write-Log ""
    Write-Log "==========================================" -Level INFO
    Write-Log $Text -Level INFO
    Write-Log "==========================================" -Level INFO
}

###########################################################
# LIVE SOURCE-ENVIRONMENT CONNECTION (-SourceManagementServer)
###########################################################
# Alternative to -SourceRepositoryFolder for anyone with live access to the
# OLD environment (community feedback -- if you're migrating MPs,
# you almost certainly still have the old management group reachable, so
# asking for a hand-exported folder every time is an unnecessary step).
#
# Deliberately NOT rewired into the dependency-resolution internals below --
# doing that would mean two different code paths (live objects vs. files on
# disk) feeding Step 2.5's indexing/matching logic, which is exactly the
# kind of divergence that produces "works for folder input, silently wrong
# for live input" bugs. Instead, this connects, enumerates, and exports the
# source management group to an ordinary folder using the EXACT same
# Get-SCOMManagementPack / Export-SCOMManagementPack pattern the script
# already documents telling people to run by hand -- then hands that folder
# to $SourceRepositoryFolder. Everything downstream (indexing, Find-In
# SourceRepository, batch-internal pulling, candidate rewriting) is
# completely unchanged and already battle-tested against folder input.
#
# READ-ONLY against the source: only Get-SCOMManagementPack (enumerate) and
# Export-SCOMManagementPack (read a copy out) are ever called here. Nothing
# is imported, changed, or deleted on the source side. Entirely separate
# from -LiveImport, which is a later, different connection to the TARGET.
function Connect-AndExportSourceEnvironment {
    param(
        [Parameter(Mandatory)][string]$ServerName,
        [System.Management.Automation.PSCredential]$Credential,
        [Parameter(Mandatory)][string]$ExportFolder
    )

    Write-Banner "STEP 2.4 - LIVE SOURCE ENVIRONMENT EXPORT (-SourceManagementServer)"
    Write-Log "Connecting to the SOURCE (OLD) SCOM management group on '$ServerName'..." -Level WARN
    Write-Log "This connection is READ-ONLY -- only enumerating and exporting installed MPs. Nothing is written to the source environment."

    try {
        Import-Module OperationsManager -ErrorAction Stop
        if ($Credential) {
            New-SCOMManagementGroupConnection -ComputerName $ServerName -Credential $Credential -ErrorAction Stop
        }
        else {
            New-SCOMManagementGroupConnection -ComputerName $ServerName -ErrorAction Stop
        }
        $srcConn = @(Get-SCOMManagementGroupConnection -ErrorAction Stop | Where-Object { $_.IsActive }) | Select-Object -First 1
        $script:SourceMGName = [string]$srcConn.ManagementGroupName
        $script:SourceMSName = [string]$srcConn.ManagementServerName
        Write-Log "Connected to SOURCE: Management Group '$($srcConn.ManagementGroupName)' via '$($srcConn.ManagementServerName)'" -Level SUCCESS
    }
    catch {
        throw "Could not connect to the SOURCE SCOM management group on '$ServerName': $($_.Exception.Message). Common causes: wrong server name, no network path from this machine, insufficient rights (Read-Only Operator or higher is enough to enumerate/export), or an OperationsManager module that can't talk to that SCOM version from this session. Fix the connection and re-run, or fall back to -SourceRepositoryFolder with a manually-run 'Get-SCOMManagementPack | Export-SCOMManagementPack -Path <folder>' instead."
    }

    if (-not (Test-Path -LiteralPath $ExportFolder)) {
        New-Item -ItemType Directory -Path $ExportFolder -Force | Out-Null
    }

    Write-Log "Enumerating Management Packs installed on '$ServerName'..."
    $srcMPs = $null
    try {
        $srcMPs = @(Get-SCOMManagementPack -ErrorAction Stop)
    }
    catch {
        throw "Connected to '$ServerName' but could not enumerate its Management Packs: $($_.Exception.Message)"
    }

    if ($srcMPs.Count -eq 0) {
        throw "Connected to '$ServerName' but it reports zero installed Management Packs -- that's not a real SCOM environment result, so something is wrong with the connection/session. Nothing to export. Fix and re-run, or use -SourceRepositoryFolder instead."
    }

    Write-Log "Source management group reports $($srcMPs.Count) installed Management Pack(s). Exporting live snapshot to '$ExportFolder'..."

    $exportFailures = New-Object System.Collections.Generic.List[string]
    $exportedCount = 0
    foreach ($mp in $srcMPs) {
        try {
            $mp | Export-SCOMManagementPack -Path $ExportFolder -ErrorAction Stop
            $exportedCount++
        }
        catch {
            $exportFailures.Add("$($mp.Name) (v$($mp.Version)): $($_.Exception.Message)")
        }
    }

    # Inventory with the Sealed flag and KeyToken. The exported .xml files do
    # not record whether the MP was sealed, and that is the single most
    # important fact for whether an MP can be moved at all (see -SourceInventory).
    try {
        $invPath = Join-Path $ExportFolder "SourceInventory.csv"
        $srcMPs | ForEach-Object {
            [PSCustomObject]@{
                Name         = $_.Name
                DisplayName  = $_.DisplayName
                Version      = [string]$_.Version
                Sealed       = [bool]$_.Sealed
                KeyToken     = [string]$_.KeyToken
                TimeCreated  = $_.TimeCreated
                LastModified = $_.LastModified
            }
        } | Export-Csv -LiteralPath $invPath -NoTypeInformation -Encoding UTF8
        $script:LiveSourceInventoryPath = $invPath
        Write-Log "Source inventory (with Sealed flag) written: $invPath" -Level SUCCESS
    }
    catch {
        Write-Log "Could not write source inventory: $($_.Exception.Message)" -Level WARN
    }

    Write-Log "NOTE: Export-SCOMManagementPack writes SEALED MPs out as plain .xml. Those copies are useful for analysis only -- they can never be imported in place of the sealed original, so they are not used to auto-resolve dependencies. Supply original .mp/.mpb files for any sealed MP you need." -Level WARN

    Write-Log "Exported $exportedCount of $($srcMPs.Count) source Management Pack(s) to '$ExportFolder'." -Level $(if ($exportFailures.Count -gt 0) { "WARN" } else { "SUCCESS" })

    if ($exportFailures.Count -gt 0) {
        Write-Log "$($exportFailures.Count) source MP(s) failed to export and will NOT be available for dependency auto-resolution (everything else still will be):" -Level WARN
        foreach ($ef in $exportFailures) { Write-Log "  - $ef" -Level WARN -NoConsole }
    }

    if ($exportedCount -eq 0) {
        throw "Connected to '$ServerName' but zero Management Packs were successfully exported -- nothing usable for dependency resolution. Check permissions/disk space at '$ExportFolder' and try again, or fall back to -SourceRepositoryFolder with a manual export."
    }

    Write-Log "Live source export complete. This folder will be used exactly like a manually-supplied -SourceRepositoryFolder from here on: $ExportFolder" -Level SUCCESS

    return $ExportFolder
}

###########################################################
# STATIC GROUP MEMBERSHIP DETECTION + GUID->NAME RESOLUTION
###########################################################
# Detects the "hand-jammed the servers into a static group" pattern: a
# GroupPopulator discovery whose MembershipRule uses an explicit <IncludeList>
# of <MonitoringObjectId> GUIDs rather than a dynamic <Expression> formula.
#
# Why this matters (and why it's invisible without this check): those GUIDs are
# object instance IDs from the SOURCE environment. The SAME server, once
# discovered by the NEW management group, gets a DIFFERENT GUID. So the group
# imports cleanly, scores 100%, and populates EMPTY in the target -- a silent
# failure that every dashboard/override targeting the group inherits.
#
# Returns a list of per-group findings: group element id, its display name if
# resolvable, how many static member GUIDs, and the GUIDs themselves (for
# optional name resolution against the source environment).
function Get-StaticGroupMembership {
    param([Parameter(Mandatory)][xml]$MPXml)

    $findings = New-Object System.Collections.Generic.List[object]

    $discoveries = $null
    try { $discoveries = $MPXml.ManagementPack.Monitoring.Discoveries.Discovery } catch { }
    if (-not $discoveries) { return $findings }

    foreach ($disc in @($discoveries)) {
        # A group-population discovery uses the GroupPopulator data source.
        $ds = $null
        try { $ds = $disc.DataSource } catch { }
        if (-not $ds) { continue }
        $dsType = [string]$ds.TypeID
        if ($dsType -notlike '*GroupPopulator*') { continue }

        $groupTarget = [string]$disc.Target

        # Each MembershipRule may carry an <IncludeList> of explicit object IDs
        # (static) and/or an <Expression> (dynamic). We only flag the static
        # portion -- dynamic rules re-populate themselves in the new env.
        $memberGuids = New-Object System.Collections.Generic.List[string]
        $hasDynamic = $false
        $mrNodes = $disc.SelectNodes('.//*[local-name()="MembershipRule"]')
        foreach ($mr in @($mrNodes)) {
            $incl = $mr.SelectNodes('.//*[local-name()="MonitoringObjectId"]')
            foreach ($idNode in @($incl)) {
                $g = [string]$idNode.InnerText
                if ($g) { $memberGuids.Add($g.Trim()) }
            }
            $expr = $mr.SelectNodes('.//*[local-name()="Expression"]')
            if ($expr -and $expr.Count -gt 0) { $hasDynamic = $true }
        }

        if ($memberGuids.Count -gt 0) {
            $findings.Add([PSCustomObject]@{
                GroupTargetID  = $groupTarget
                StaticCount    = $memberGuids.Count
                DistinctCount  = @($memberGuids | Select-Object -Unique).Count
                HasDynamicToo  = $hasDynamic
                MemberGuids    = @($memberGuids | Select-Object -Unique)
            })
        }
    }

    return $findings
}

# Resolves a set of source-environment object GUIDs to server names by querying
# the OLD management group. READ-ONLY. Requires an active source connection
# (established by -SourceManagementServer earlier). Returns a hashtable
# GUID -> name (FQDN/DisplayName/PrincipalName), with unresolved GUIDs mapped to
# $null so callers can report "member no longer in source" (stale/decommissioned
# servers that were never cleaned out of the hand-jammed list).
function Resolve-SourceObjectNames {
    param(
        [Parameter(Mandatory)][string[]]$Guids,
        [Parameter(Mandatory)][string]$SourceServer,
        [System.Management.Automation.PSCredential]$Credential
    )

    $result = @{}
    foreach ($g in $Guids) { $result[$g] = $null }

    # Ensure a connection to the source. It may already be connected from the
    # earlier -SourceManagementServer export; connect (idempotently) to be safe.
    try {
        Import-Module OperationsManager -ErrorAction Stop
        if ($Credential) {
            New-SCOMManagementGroupConnection -ComputerName $SourceServer -Credential $Credential -ErrorAction Stop
        }
        else {
            New-SCOMManagementGroupConnection -ComputerName $SourceServer -ErrorAction Stop
        }
    }
    catch {
        Write-Log "Could not connect to source '$SourceServer' to resolve static group member names: $($_.Exception.Message). GUIDs will be reported unresolved." -Level WARN
        return $result
    }

    $resolved = 0
    foreach ($g in $Guids) {
        try {
            $obj = Get-SCOMMonitoringObject -Id $g -ErrorAction SilentlyContinue
            if ($obj) {
                $name = $obj.DisplayName
                if (-not $name) { $name = $obj.Name }
                if (-not $name -and $obj.Path) { $name = $obj.Path }
                if ($name) { $result[$g] = $name; $resolved++ }
            }
        }
        catch {
            # leave as $null (unresolved)
        }
    }

    Write-Log "Resolved $resolved of $($Guids.Count) static member GUID(s) to names against source '$SourceServer'. Unresolved GUIDs are likely stale/decommissioned objects no longer in the old environment." -Level $(if ($resolved -lt $Guids.Count) { 'WARN' } else { 'SUCCESS' })

    # Connecting to the source made IT the active connection. Every later
    # Get-SCOMManagementPack / Import-SCOMManagementPack call in this session
    # must go to the TARGET, so switch back explicitly.
    Use-TargetConnection
    return $result
}

###########################################################
# CONNECTION + INSTALLED-MP HELPERS (target side)
###########################################################
# The OperationsManager module keeps every connection made in the session and
# runs cmdlets against whichever one is ACTIVE. With both a source and a target
# connection open, an unguarded Get-SCOMManagementPack can silently query the
# wrong management group. All target-side calls go through these helpers.

$script:TargetServer = $null
$script:TargetConnection = $null
$script:InstalledMPCache = $null

# Makes the TARGET connection active again. Throws if it cannot -- carrying on
# against whatever connection happens to be active could mean the SOURCE.
function Use-TargetConnection {
    if (-not $script:TargetConnection) { return }
    try { $script:TargetConnection | Set-SCOMManagementGroupConnection -ErrorAction Stop }
    catch { throw "Could not switch back to the TARGET SCOM connection ($($script:TargetConnection.ManagementGroupName)): $($_.Exception.Message). Stopping rather than risk running against the source." }
}

function Set-ActiveScomConnection {
    param([Parameter(Mandatory)][string]$Server)
    try {
        $short = ($Server -split '\.')[0]
        $conn = @(Get-SCOMManagementGroupConnection -ErrorAction Stop) | Where-Object {
            $_.ManagementServerName -eq $Server -or (([string]$_.ManagementServerName) -split '\.')[0] -eq $short
        } | Select-Object -First 1
        if ($conn) {
            $conn | Set-SCOMManagementGroupConnection -ErrorAction Stop
            return $true
        }
        Write-Log "No open SCOM connection matches '$Server'." -Level WARN -NoConsole
    }
    catch {
        Write-Log "Could not switch active SCOM connection to '$Server': $($_.Exception.Message)" -Level WARN
    }
    return $false
}

function Update-InstalledMPCache {
    Use-TargetConnection
    $script:InstalledMPCache = @{}
    foreach ($m in @(Get-SCOMManagementPack -ErrorAction Stop)) {
        $script:InstalledMPCache[[string]$m.Name] = $m
    }
    Write-Log "Target management group currently has $($script:InstalledMPCache.Count) MP(s) installed." -NoConsole
}

function Get-InstalledMP {
    param([Parameter(Mandatory)][string]$Name)
    if ($null -eq $script:InstalledMPCache) { Update-InstalledMPCache }
    if ($script:InstalledMPCache.ContainsKey($Name)) { return $script:InstalledMPCache[$Name] }
    return $null
}

function Get-ExceptionChainText {
    param([Parameter(Mandatory)]$ErrorRecord)
    $lines = New-Object System.Collections.Generic.List[string]
    $ex = $ErrorRecord.Exception
    $depth = 0
    while ($ex -and $depth -lt 8) {
        $msg = if ($null -ne $ex.Message) { $ex.Message.Trim() } else { "(no message)" }
        $lines.Add("[$depth] $($ex.GetType().Name): $msg")
        $ex = $ex.InnerException
        $depth++
    }
    return , $lines
}


###########################################################
# INIT
###########################################################

Clear-Host

# Per-batch-MP working state lives in $BatchResults (populated in Steps 4-8).
# Repository / symbol table / target-side state is shared across the whole
# batch and built once in Steps 1-2, same as the original single-file design.
$Repository   = @{}
$VersionMap   = @{}
$Warnings     = @()
$BatchSource  = @{}   # MPID -> [PSCustomObject] batch source MP entry (see Step 0)
$BatchResults = @{}   # MPID -> per-MP results bag, populated through Steps 4-8
$script:StaticGroupReport = New-Object System.Collections.Generic.List[object]  # rows for StaticGroupMembership.csv

###########################################################
# UI FALLBACKS
###########################################################

if ($NonInteractive) {
    if (-not $InputPath -and -not $CatalogOnly) { throw "-NonInteractive: -InputPath is required." }
    if (-not $RepositoryFolder) { throw "-NonInteractive: -RepositoryFolder is required." }
}

if (-not $InputPath -and -not $CatalogOnly) {
    $step1Owner = Get-TopMostOwner
    try {
        [System.Windows.Forms.MessageBox]::Show(
            $step1Owner,
            "STEP 1 of 3: WHAT DO YOU WANT TO MIGRATE?`n`nSelect the specific Management Pack file(s), or a folder of them, that you want to MIGRATE INTO SCOM $TargetVersion.`n`nThis is your starting point -- e.g. one custom MP, or a small batch you've pulled out to work on.",
            "Step 1 of 3 - Select What To Migrate",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    }
    finally {
        $step1Owner.Close()
        $step1Owner.Dispose()
    }

    Write-Host "Select Source MP(s)..." -ForegroundColor Yellow
    $InputPath = Select-InputPathOrPaths
}

if (-not $RepositoryFolder) {
    $step2Owner = Get-TopMostOwner
    try {
        [System.Windows.Forms.MessageBox]::Show(
            $step2Owner,
            "STEP 2 of 3: WHAT'S ALREADY IN SCOM $TargetVersion (the destination)?`n`nSelect a folder containing a COMPLETE export of every Management Pack currently installed in your SCOM $TargetVersion environment -- the system you're migrating INTO.`n`nDon't have one yet? From a machine connected to SCOM 2025, run:`n  Get-SCOMManagementPack | ForEach-Object { `$_ | Export-SCOMManagementPack -Path 'C:\Temp\Target_AllMPs' }`nThen select that folder next.`n`nThis must be a COMPLETE export, not just a few MPs -- an incomplete folder here causes wrong results throughout the rest of this tool.",
            "Step 2 of 3 - Select SCOM $TargetVersion (Destination) Repository",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    }
    finally {
        $step2Owner.Close()
        $step2Owner.Dispose()
    }

    Write-Host "Select Repository Folder..." -ForegroundColor Yellow
    $RepositoryFolder = Select-Folder "STEP 2 of 3: Select the SCOM $TargetVersion (destination) complete MP repository folder"
}

# Sanity check: a real SCOM install -- even a fresh one -- has 50+ core
# library MPs alone. A folder with far fewer than that is almost certainly
# an incomplete scan (e.g. the default Health Service State runtime cache,
# which only holds whatever the management server happens to have cached
# locally, not the full repository), and using it will silently produce
# wrong "missing" / "found" verdicts and version mismatches throughout the
# rest of this run. Catch it here, loudly, before any of that happens.
if ($RepositoryFolder -and (Test-Path -LiteralPath $RepositoryFolder)) {
    $repoXmlCount = @(Get-ChildItem -LiteralPath $RepositoryFolder -Filter *.xml -Recurse -ErrorAction SilentlyContinue).Count
    Write-Host "Repository folder '$RepositoryFolder' contains $repoXmlCount .xml file(s)." -ForegroundColor Yellow

    if ($repoXmlCount -lt 80 -and $NonInteractive) {
        Write-Host "WARNING: '$RepositoryFolder' only contains $repoXmlCount MP file(s). A complete SCOM repository export normally has 100+. Results will be unreliable if this is not a complete export. Continuing because -NonInteractive is set." -ForegroundColor Yellow
    }
    elseif ($repoXmlCount -lt 80) {
        $lowCountOwner = Get-TopMostOwner
        try {
            $proceedAnyway = [System.Windows.Forms.MessageBox]::Show(
                $lowCountOwner,
                "WARNING: '$RepositoryFolder' only contains $repoXmlCount MP file(s).`n`nA real SCOM repository -- even freshly installed -- normally has 100+ MPs (the core libraries alone are 50+). This folder is very likely incomplete, such as the default Health Service State runtime cache rather than a full export.`n`nUsing an incomplete repository will cause wrong 'missing'/'found' results and version mismatches throughout this run.`n`nRECOMMENDED: Cancel, then run from a live SCOM connection:`n  Get-SCOMManagementPack | ForEach-Object { `$_ | Export-SCOMManagementPack -Path <folder> }`nand point -RepositoryFolder at that export instead.`n`nContinue anyway with this folder?",
                "Repository Folder Looks Incomplete",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
        }
        finally {
            $lowCountOwner.Close()
            $lowCountOwner.Dispose()
        }

        if ($proceedAnyway -ne "Yes") {
            throw "Aborted: repository folder '$RepositoryFolder' looks incomplete ($repoXmlCount MPs). Re-run with a complete repository export -- see the warning message above for the command to generate one."
        }

        Write-Log "User chose to proceed despite a low repository MP count ($repoXmlCount). Results may be unreliable." -Level WARN
    }
}

if ($SourceRepositoryFolder -and $SourceManagementServer) {
    throw "Both -SourceRepositoryFolder and -SourceManagementServer were supplied -- pick one. (Silently preferring one over the other is exactly the kind of hidden default this tool avoids elsewhere, e.g. -RepositoryFolder has no default for the same reason.) Use -SourceRepositoryFolder if you already have a folder export of the old environment, or -SourceManagementServer if you'd rather this script connect and export a live snapshot itself."
}

if (-not $NonInteractive -and -not $CatalogOnly -and -not $SourceRepositoryFolder -and -not $SourceManagementServer) {
    # This one is genuinely optional (unlike InputPath/RepositoryFolder/
    # OutputFolder above), so it gets an explicit choice instead of just
    # popping a folder browser -- silently skipping this on Cancel left no
    # visible trace earlier, which is exactly the confusion this fixes.
    #
    # Skipped entirely in -CatalogOnly mode: a catalog check never touches
    # the old source environment, so prompting for it there is pure noise.
    #
    # Three-way choice (YesNoCancel), per community feedback that if
    # someone is running this migration at all, they most likely still have
    # live access to BOTH the old and new environments -- so "connect
    # directly" should be offered on equal footing with "browse to a folder
    # someone already exported," not left as a command-line-only option.
    $promptOwner = Get-TopMostOwner
    try {
        $sourceChoice = [System.Windows.Forms.MessageBox]::Show(
            $promptOwner,
            "STEP 3 of 3 (OPTIONAL): WHAT WAS IN THE OLD SCOM $SourceVersion (the source)?`n`nIf the MP(s) you're migrating depend on other MPs that aren't in SCOM $TargetVersion yet (e.g. a shared library), this tool can automatically find and pull them in from the OLD SCOM 2016 environment's MPs.`n`nYES = Connect directly to the old SCOM 2016 management server now and pull a live export automatically (requires network access + read rights to it from this session).`n`nNO = I already have a folder of exported MPs from the old environment -- let me browse to it.`n`nCANCEL = Skip this -- any missing dependencies will just be reported instead of auto-resolved.",
            "Step 3 of 3 - Old (Source) Environment, for Finding Dependencies",
            [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
    }
    finally {
        $promptOwner.Close()
        $promptOwner.Dispose()
    }

    if ($sourceChoice -eq "Yes") {
        $SourceManagementServer = Read-Host "Enter the OLD (source, SCOM $SourceVersion) management server to connect to"
        if ([string]::IsNullOrWhiteSpace($SourceManagementServer)) {
            Write-Host "No server name entered -- continuing without dependency auto-resolution." -ForegroundColor Yellow
            $SourceManagementServer = $null
        }
        else {
            $wantsSourceCred = Read-Host "Use different credentials to connect to '$SourceManagementServer'? Old environments are sometimes a different domain/forest. (y/N)"
            if ($wantsSourceCred -match '^(y|yes)$') {
                $SourceCredential = Get-Credential -Message "Credentials for source management server '$SourceManagementServer'"
            }
        }
    }
    elseif ($sourceChoice -eq "No") {
        Write-Host "Select Source Repository Folder (for dependency auto-resolution)..." -ForegroundColor Yellow
        $SourceRepositoryFolder = Select-Folder "STEP 3 of 3: Select the SCOM $SourceVersion (source / old environment) MP folder"
        if (-not $SourceRepositoryFolder) {
            Write-Host "No folder selected -- continuing without dependency auto-resolution." -ForegroundColor Yellow
        }
    }
    # Cancel: leave both unset, same as declining the Yes/No before.
}

if (-not $OutputFolder -and -not $NonInteractive) {
    Write-Host "Select Output Folder..." -ForegroundColor Yellow
    $OutputFolder = Select-Folder "Select Output Folder"
}

###########################################################
# VALIDATION
###########################################################

if (-not $CatalogOnly -and (-not $InputPath -or @($InputPath).Count -eq 0)) {
    throw "No input Management Pack(s) or folder was provided or selected."
}
if (-not $RepositoryFolder) {
    throw "No repository folder was provided or selected."
}
if (-not $OutputFolder) {
    throw "No output folder was provided or selected."
}

if ($CatalogOnly -and -not $CheckCatalog) {
    $CheckCatalog = $true
}
if ($DownloadFromCatalog -and -not $CheckCatalog) {
    Write-Log "-DownloadFromCatalog was supplied without -CheckCatalog or -CatalogOnly -- enabling -CheckCatalog automatically, since there's nothing to download from otherwise." -Level WARN -NoConsole
    $CheckCatalog = $true
}

# Normalize to an array throughout: -InputPath accepts one path, several
# paths (e.g. from -Multiselect in the picker, or a comma-separated -Param
# on the command line), or a single folder. Trim each entry individually.
#
# Filter out null/empty entries BEFORE trimming: an array can be non-empty
# overall (so the -not $InputPath check above passes) while still containing
# a stray $null element -- e.g. a trailing comma in a comma-separated
# -InputPath on the command line (`-InputPath "a.xml", -RepositoryFolder ...`
# is a common way to accidentally produce one). Without this filter, .Trim()
# on that null element crashes with a bare "cannot call a method on a
# null-valued expression" instead of a clear message.
#
# Every count below goes through @(...).Count rather than $var.Count:
# under Set-StrictMode -Version Latest (active in this script), reading the
# .Count PROPERTY off a value that is $null or a bare scalar throws
# "property 'Count' cannot be found on this object". @(...) forces an array
# first, whose .Count is always safe. This is the same class of strict-mode
# trap the changelog's earlier "Empty-Collection Fix" addressed.
$inputBefore = @($InputPath)
$InputPath   = @($InputPath | Where-Object { $_ } | ForEach-Object {
    # "a,b" arriving as ONE string (e.g. via powershell.exe -File) -> split it.
    if (-not (Test-Path -LiteralPath $_.Trim()) -and $_ -like '*,*') { $_.Split(',') | Where-Object { $_.Trim() } | ForEach-Object { $_.Trim() } }
    else { $_.Trim() }
})
$droppedCount = @($inputBefore).Count - @($InputPath).Count
if ($droppedCount -gt 0 -and -not $CatalogOnly) {
    Write-Log "$droppedCount empty/null -InputPath entry(ies) were ignored. If that's unexpected, check your command line for a stray trailing comma after an -InputPath value (e.g. '-InputPath `"a.xml`", -RepositoryFolder ...' silently adds an empty element)." -Level WARN
}
if (-not $CatalogOnly -and @($InputPath).Count -eq 0) {
    throw "No usable input Management Pack(s) remained after removing empty/null entries. Check your -InputPath argument for a stray trailing comma or similar."
}
$RepositoryFolder = $RepositoryFolder.Trim()
$OutputFolder     = $OutputFolder.Trim()

foreach ($p in $InputPath) {
    if (-not (Test-Path -LiteralPath $p)) {
        throw "Input path not found: $p"
    }
}

if (-not (Test-Path -LiteralPath $RepositoryFolder)) {
    throw "Repository folder not found: $RepositoryFolder"
}

if (-not (Test-Path -LiteralPath $OutputFolder)) {
    Write-Host "Creating output folder..." -ForegroundColor Yellow
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
}
# Absolute paths from here on: the generated import script and the manifest
# record file paths, and they must still be valid when run from elsewhere.
$OutputFolder = (Resolve-Path -LiteralPath $OutputFolder).ProviderPath
$RepositoryFolder = (Resolve-Path -LiteralPath $RepositoryFolder).ProviderPath
$InputPath = @($InputPath | ForEach-Object { (Resolve-Path -LiteralPath $_).ProviderPath })

# Start the transcript-style log as soon as OutputFolder exists (moved
# earlier than before) so that a live -SourceManagementServer connect/export
# below -- which can take a while and is worth having a permanent record of
# -- is captured in the log file, not just echoed to the console.
$runStamp = Get-Date -Format "yyyyMMdd-HHmmss"
$script:LogFile = Join-Path $OutputFolder "MPCompiler.$runStamp.log"
Write-Log "Log file: $script:LogFile" -Level INFO -NoConsole

# -SourceManagementServer runs BEFORE the -SourceRepositoryFolder check
# below, because it PRODUCES the folder that check validates: a live
# connect-and-export into a fresh, timestamped folder under -OutputFolder,
# then $SourceRepositoryFolder is set to point at it. Everything from here
# on (including Step 2.5's indexing later in the script) treats it exactly
# like a manually-supplied -SourceRepositoryFolder, because it now is one.
$script:SourceRepositoryIsLive = $false
$script:SourceMGName = $null
$script:SourceMSName = $null
if ($SourceManagementServer) {
    $liveExportFolder = Join-Path $OutputFolder "SourceLiveExport_$(Get-Date -Format 'yyyyMMdd-HHmmss')"
    $script:LiveSourceInventoryPath = $null
    $SourceRepositoryFolder = Connect-AndExportSourceEnvironment -ServerName $SourceManagementServer.Trim() -Credential $SourceCredential -ExportFolder $liveExportFolder
    $script:SourceRepositoryIsLive = $true
    if (-not $SourceInventory -and $script:LiveSourceInventoryPath) { $SourceInventory = $script:LiveSourceInventoryPath }
}

# Source inventory: which MPs were SEALED in the old environment.
$script:SourceSealedMap = @{}   # MPID -> $true when sealed in source
if ($SourceInventory) {
    if (-not (Test-Path -LiteralPath $SourceInventory)) { throw "Source inventory not found: $SourceInventory" }
    foreach ($row in @(Import-Csv -LiteralPath $SourceInventory)) {
        if ($row.Name -and ([string]$row.Sealed -match '^(true|1|yes)$')) { $script:SourceSealedMap[[string]$row.Name] = $true }
    }
    Write-Log "Source inventory loaded: $($script:SourceSealedMap.Count) sealed MP(s) recorded in the source environment." -NoConsole
}

if ($SourceRepositoryFolder) {
    $SourceRepositoryFolder = $SourceRepositoryFolder.Trim()
    if (-not (Test-Path -LiteralPath $SourceRepositoryFolder)) {
        throw "Source repository folder not found: $SourceRepositoryFolder"
    }
}

# Per-MP candidate XML/reports are written into a CandidateMPs subfolder so
# the OutputFolder root stays readable when migrating dozens of MPs at once.
$CandidateFolder = Join-Path $OutputFolder "CandidateMPs"
if (-not (Test-Path -LiteralPath $CandidateFolder)) {
    New-Item -ItemType Directory -Path $CandidateFolder -Force | Out-Null
}

# Normalize -ExcludeMPIDs into a hashtable for fast lookups everywhere this
# is checked (Step 2.5 auto-resolution, Step 3.5 live import, etc.) rather
# than a linear -contains scan repeated per-reference across a potentially
# large batch.
$ExcludedSet = @{}
foreach ($exId in $ExcludeMPIDs) {
    if (-not [string]::IsNullOrWhiteSpace($exId)) {
        $ExcludedSet[$exId.Trim()] = $true
    }
}

Write-Banner "SCOM MP BATCH COMPILER $($script:BuildLabel) - STEP 0"
Write-Log "If you're troubleshooting a fix that should already be in place, confirm THIS build number matches what was most recently shared before reporting an issue." -Level WARN -NoConsole
Write-Log "Source version label : $SourceVersion"
Write-Log "Target version label : $TargetVersion"
Write-Log "Strict mode          : $($Strict.IsPresent)"
Write-Log "Allow ID rewrite     : $($AllowIDRewrite.IsPresent)"
Write-Log "Skip sealed (.mp)    : $($SkipSealedExtraction.IsPresent)"
$sourceRepoLogText = if ($SourceRepositoryFolder) {
    if ($script:SourceRepositoryIsLive) { "$SourceRepositoryFolder (LIVE export just pulled from '$SourceManagementServer')" }
    else { $SourceRepositoryFolder }
}
else {
    '(none -- missing deps will only be reported, not auto-resolved)'
}
Write-Log "Source repo (deps)   : $sourceRepoLogText"
Write-Log "Audit-only mode      : $($AuditOnly.IsPresent)$(if ($AuditOnly) { ' -- every reference checked directly against -RepositoryFolder; batch-internal self-consistency does NOT count as resolved' })"
Write-Log "Live import mode     : $($LiveImport.IsPresent)$(if ($LiveImport) { " -- WILL CONNECT TO AND IMPORT INTO A REAL SCOM MANAGEMENT GROUP ($(if ($ManagementServer) { $ManagementServer } else { 'localhost (default)' }))" })" -Level $(if ($LiveImport) { "WARN" } else { "INFO" })
Write-Log "Excluded MP ID(s)    : $(if ($ExcludedSet.Count -gt 0) { $ExcludedSet.Keys -join ', ' } else { '(none)' })"
Write-Log "Force core rewrite   : $($ForceRewriteCoreLibraries.IsPresent)$(if ($ForceRewriteCoreLibraries) { ' -- CoreLibrary/SystemCenter/IISCommonLibrary references WILL be auto-rewritten to a newer repo version even when they normally would not be. This is NOT guaranteed to import successfully -- test before trusting it.' })" -Level $(if ($ForceRewriteCoreLibraries) { "WARN" } else { "INFO" })
Write-Log "Catalog check mode   : $($CheckCatalog.IsPresent)$(if ($CatalogOnly) { ' (-CatalogOnly: will stop after Step 1.5, no batch processed)' })$(if ($DownloadFromCatalog) { ' + -DownloadFromCatalog (experimental)' })$(if ($CatalogStrictMatch) { ' + -CatalogStrictMatch (generic-library soft matches reported as gaps)' })"

###########################################################
# STEP 1.5 HELPERS - MICROSOFT MP CATALOG (-CheckCatalog / -CatalogOnly)
###########################################################
# Ported from the standalone Get-SCOM2025MPCatalog.ps1 companion script,
# folded in here so the catalog check runs as part of the normal workflow
# instead of a separate tool. Same caveats apply as in that script: the
# Microsoft Learn catalog page is explicitly documented (by Microsoft) as
# reference-only and subject to change, and matching against your
# repository is by fuzzy display-name comparison, not MP ID, because the
# catalog page carries no IDs. Treat results as a strong hint, not a verdict.

function Get-MPCatalog {
    param([Parameter(Mandatory)][string]$View)

    $url = "https://learn.microsoft.com/en-us/system-center/scom/management-pack-list?view=$View"
    Write-Log "Fetching official Microsoft management pack catalog: $url"

    try {
        $resp = Invoke-WebRequest -Uri $url -UseBasicParsing -ErrorAction Stop
    }
    catch {
        throw "Could not fetch the Microsoft Learn catalog page at '$url': $($_.Exception.Message). Check network/proxy access to learn.microsoft.com from this machine, or skip -CheckCatalog/-CatalogOnly and keep using -RepositoryFolder the way you already do."
    }

    $html = $resp.Content
    $rows = New-Object System.Collections.Generic.List[object]

    # Primary parser: matches the real docfx-rendered table structure as of
    # when this was written -- <tr><td><a href=..>Name</a>...</td>
    # <td>Version</td><td>Date</td></tr>. Falls back to plain link
    # extraction rather than silently reporting zero results as "the
    # catalog is empty" if Microsoft changes the page layout.
    $rowPattern = '<tr>\s*<td>\s*<a\s+href="([^"]+)"[^>]*>([^<]+)</a>([^<]*)</td>\s*<td[^>]*>\s*([^<]*?)\s*</td>\s*<td[^>]*>\s*([^<]*?)\s*</td>\s*</tr>'
    $rowMatches = [regex]::Matches($html, $rowPattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::Singleline)

    if ($rowMatches.Count -eq 0) {
        Write-Log "Primary catalog table parser found 0 rows -- Microsoft may have changed the page layout since this script was written. Falling back to link-only extraction (Name + URL, no Version/Date column)." -Level WARN
        foreach ($link in $resp.Links) {
            if ($link.href -match 'microsoft\.com/(en-us/)?download/details\.aspx\?id=(\d+)') {
                $rows.Add([PSCustomObject]@{
                    Name       = ($link.innerText -replace '\s+', ' ').Trim()
                    Url        = $link.href
                    Version    = $null
                    Date       = $null
                    DownloadId = $Matches[2]
                })
            }
        }
        if ($rows.Count -eq 0) {
            throw "Could not extract any management pack entries from '$url' via either the table parser or plain link fallback. Open the URL in a browser to confirm it still shows the expected table."
        }
        Write-Log "Fallback link extraction found $($rows.Count) download links." -Level WARN
    }
    else {
        Write-Log "Parsed $($rowMatches.Count) rows from the official catalog table." -Level SUCCESS
        foreach ($m in $rowMatches) {
            $entryUrl = $m.Groups[1].Value
            $name = (($m.Groups[2].Value + $m.Groups[3].Value) -replace '\s+', ' ').Trim()
            $version = $m.Groups[4].Value.Trim()
            $date = $m.Groups[5].Value.Trim()
            $downloadId = $null
            if ($entryUrl -match 'id=(\d+)') { $downloadId = $Matches[1] }
            $rows.Add([PSCustomObject]@{
                Name       = $name
                Url        = $entryUrl
                Version    = $version
                Date       = $date
                DownloadId = $downloadId
            })
        }
    }

    return $rows
}

function Merge-DuplicateDownloadIds {
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Rows)

    $merged = New-Object System.Collections.Generic.List[object]
    $grouped = $Rows | Group-Object -Property DownloadId
    foreach ($g in $grouped) {
        $names = $g.Group.Name | Select-Object -Unique
        $primary = $names | Select-Object -First 1
        $aliases = $names | Where-Object { $_ -ne $primary }
        $merged.Add([PSCustomObject]@{
            Name        = $primary
            AlsoKnownAs = ($aliases -join '; ')
            Url         = $g.Group[0].Url
            Version     = $g.Group[0].Version
            Date        = $g.Group[0].Date
            DownloadId  = $g.Name
        })
    }
    return $merged
}

# Fuzzy name matching -- the catalog page has no MP IDs, only display names,
# so this compares normalized tokens rather than doing exact ID lookups the
# way the rest of this script does against $Repository.
#
# The catalog's display names are prose stuffed with version-range noise
# ("... version agnostic 2012-2022+ (Windows and Linux)"), while a local MP
# is best identified by its ID ("Microsoft.SQLServer.ReportingServices...").
# A naive symmetric Jaccard over the full token union scores real matches
# far too low, because the noise words inflate the union. So this does three
# things differently: (1) strips year/version-range noise tokens, (2)
# tokenizes the local MP ID by splitting dotted segments AND CamelCase
# within them ("InternetInformationServices" -> internet information
# services), and (3) scores by COVERAGE of the catalog's distinguishing
# tokens (how many appear locally) rather than symmetric overlap -- a short,
# specific ID that covers most of the catalog's real keywords should score
# high even though it's much shorter than the catalog's verbose name. Each
# catalog entry is scored against both the local display name and the local
# ID, and the better of the two is taken.
$script:CatalogMatchStopWords = @('microsoft', 'system', 'center', 'management', 'pack', 'packs', 'for', 'the', 'and', 'of', 'on', 'in', 'to', 'version', 'agnostic', 'plus')

function Get-CatalogNormalizedTokens {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return @() }
    $clean = ($Text -replace '[^a-zA-Z0-9]+', ' ').ToLowerInvariant()
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($t in ($clean -split '\s+')) {
        if (-not $t) { continue }
        if ($script:CatalogMatchStopWords -contains $t) { continue }
        # Drop version noise: bare 4-digit years (2012, 2019) and long
        # all-digit tokens (1709, 108512-style ids) that don't distinguish
        # one MP family from another.
        if ($t -match '^\d{4}$') { continue }
        if ($t -match '^\d{4,}$') { continue }
        $out.Add($t)
    }
    return $out.ToArray()
}

# Tokenizes a local MP ID by splitting on dots/spaces AND breaking CamelCase
# inside each segment, so "Microsoft.Windows.InternetInformationServices.2016"
# yields windows, internet, information, services (year dropped). This is
# what lets a dotted ID match the catalog's spelled-out prose name.
function Get-CatalogIdTokens {
    param([string]$MPId)
    if ([string]::IsNullOrWhiteSpace($MPId)) { return @() }
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($seg in ($MPId -split '[.\s]+')) {
        if (-not $seg) { continue }
        # CamelCase / digit-run splitter
        $words = [regex]::Matches($seg, '[A-Z]+(?=[A-Z][a-z])|[A-Z]?[a-z]+|[A-Z]+|\d+')
        foreach ($w in $words) {
            $lw = $w.Value.ToLowerInvariant()
            if (-not $lw) { continue }
            if ($script:CatalogMatchStopWords -contains $lw) { continue }
            if ($lw -match '^\d{4}$') { continue }
            if ($lw -match '^\d{4,}$') { continue }
            $out.Add($lw)
        }
    }
    return $out.ToArray()
}

# Coverage score: fraction of the CATALOG's distinguishing tokens that also
# appear in the local token set. Asymmetric on purpose -- see the block
# comment above. Returns 0.0 when either side is empty.
#
# Guard against generic-word false positives: infrastructure words like
# "windows"/"server"/"services"/"library" are shared by dozens of unrelated
# MPs, so an overlap made up ONLY of those shouldn't count as a match (e.g.
# "Windows Server Cluster" must not match "SQL Server ... (Windows)" just on
# windows+server). A match therefore requires at least one shared token that
# is NOT in this generic set -- something family-identifying like "sql",
# "iis"/"internet", "exchange", "sharepoint", etc.
#
# "linux"/"unix" are here as PLATFORM QUALIFIERS: a catalog name ending in
# "(Windows and Linux)" must not match "Microsoft.Linux.RHEL.7" on the word
# "linux" alone -- that's a platform tag, not the MP's identity. Without
# this, the version-agnostic SQL Server MP mis-matched RHEL instead of an
# actual SQL Server MP.
$script:CatalogGenericTokens = @('windows', 'server', 'service', 'services', 'core', 'library', 'common', 'discovery', 'monitoring', 'monitor', 'views', 'reports', 'linux', 'unix')

# Set once from -CatalogStrictMatch before the catalog check runs. In strict
# mode, a match requires that EVERY distinctive (non-generic) token from the
# catalog entry appears locally -- not just one. This flips soft matches to a
# generic family library (e.g. all the SQL sub-feature entries matching the
# generic SQLServer presentation MP on "sql" alone) into "No Local Match
# Found", so they land in the actionable/shopping-list bucket instead of the
# reassurance bucket. See the -CatalogStrictMatch help for the full rationale.
# Named distinctly from the -CatalogStrictMatch PARAMETER (which lives at
# script scope too) so assigning this flag can't clobber the param value.
$script:CatalogStrictMatchActive = $false

function Get-CatalogCoverageScore {
    param([string[]]$CatalogTokens, [string[]]$LocalTokens)
    if (-not $CatalogTokens -or -not $LocalTokens) { return 0.0 }
    $catSet = @($CatalogTokens | Select-Object -Unique)
    $localSet = @($LocalTokens | Select-Object -Unique)
    if ($catSet.Count -eq 0) { return 0.0 }

    $shared = @($catSet | Where-Object { $localSet -contains $_ })
    if ($shared.Count -eq 0) { return 0.0 }

    # Require at least one distinctive (non-generic) shared token.
    $catDistinctive = @($catSet | Where-Object { $script:CatalogGenericTokens -notcontains $_ })
    $sharedDistinctive = @($shared | Where-Object { $script:CatalogGenericTokens -notcontains $_ })
    if ($sharedDistinctive.Count -eq 0) { return 0.0 }

    # Strict mode: EVERY distinctive catalog token must be present locally,
    # not just one. "SQL Server Reporting Services" (distinctive: sql,
    # reporting) only matches a local MP that has BOTH sql AND reporting --
    # a generic SQLServer library that has only "sql" no longer qualifies.
    if ($script:CatalogStrictMatchActive) {
        if ($catDistinctive.Count -eq 0) { return 0.0 }
        $missingDistinctive = @($catDistinctive | Where-Object { $sharedDistinctive -notcontains $_ })
        if ($missingDistinctive.Count -gt 0) { return 0.0 }
    }

    return [math]::Round($shared.Count / $catSet.Count, 3)
}

# Gets a human-readable display name for an already-loaded $Repository
# entry, reusing the cached .XML clone from Step 1 rather than re-reading
# the file from disk. Falls back to the MP ID if no display string is found.
function Get-RepositoryDisplayName {
    param([Parameter(Mandatory)][object]$RepositoryEntry)

    try {
        $dsNodes = $RepositoryEntry.XML.ManagementPack.LanguagePacks.LanguagePack.DisplayStrings.DisplayString
        if ($dsNodes) {
            $self = $dsNodes | Where-Object { $_.ElementID -eq $RepositoryEntry.ID } | Select-Object -First 1
            if ($self -and $self.Name) { return $self.Name }
        }
    }
    catch {
        # fall through to ID fallback below
    }
    return $RepositoryEntry.ID
}

function Find-BestCatalogMatch {
    param(
        [Parameter(Mandatory)][string]$CatalogName,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$LocalEntries,
        [double]$Threshold = 0.5
    )

    if (-not $LocalEntries -or $LocalEntries.Count -eq 0) {
        return [PSCustomObject]@{ Status = 'No Local Repository Loaded'; LocalMPId = $null; LocalVersion = $null; Score = 0.0 }
    }

    $catalogTokens = Get-CatalogNormalizedTokens -Text $CatalogName
    $best = $null
    $bestScore = 0.0
    $bestSharedCount = 0

    foreach ($local in $LocalEntries) {
        # Score against both the display name and the ID-derived tokens, and
        # keep whichever is higher -- some local MPs have rich display names,
        # others only meaningful IDs.
        $nameTokens = Get-CatalogNormalizedTokens -Text $local.DisplayName
        $idTokens   = Get-CatalogIdTokens -MPId $local.MPId
        $nameScore = Get-CatalogCoverageScore -CatalogTokens $catalogTokens -LocalTokens $nameTokens
        $idScore   = Get-CatalogCoverageScore -CatalogTokens $catalogTokens -LocalTokens $idTokens
        $score = [math]::Max($nameScore, $idScore)

        # Raw count of shared tokens (from whichever side scored higher) used
        # ONLY as a tie-breaker: when two local MPs cover the same FRACTION of
        # catalog tokens, prefer the one sharing more tokens outright -- i.e.
        # the more specific match (Microsoft.SQLServer.ReportingServices...)
        # over a generic library (Microsoft.SQLServer.Generic.Presentation)
        # that happens to hit the same coverage ratio on fewer words.
        $winningLocalTokens = if ($idScore -ge $nameScore) { $idTokens } else { $nameTokens }
        $sharedCount = @($catalogTokens | Select-Object -Unique | Where-Object { $winningLocalTokens -contains $_ }).Count

        if ($score -gt $bestScore -or ($score -eq $bestScore -and $sharedCount -gt $bestSharedCount)) {
            $bestScore = $score
            $bestSharedCount = $sharedCount
            $best = $local
        }
    }

    if ($best -and $bestScore -ge $Threshold) {
        return [PSCustomObject]@{ Status = 'Likely Match'; LocalMPId = $best.MPId; LocalVersion = $best.Version; Score = $bestScore }
    }
    else {
        return [PSCustomObject]@{ Status = 'No Local Match Found'; LocalMPId = $null; LocalVersion = $null; Score = $bestScore }
    }
}

function Resolve-CatalogDownloadFileUrl {
    param([Parameter(Mandatory)][string]$DetailsUrl)

    try {
        $resp = Invoke-WebRequest -Uri $DetailsUrl -UseBasicParsing -ErrorAction Stop
    }
    catch {
        Write-Log "Could not load download page '$DetailsUrl': $($_.Exception.Message)" -Level WARN -NoConsole
        return $null
    }

    $html = $resp.Content

    # Best-effort only -- see -DownloadFromCatalog help. If all patterns
    # fail, the caller falls back to the details page URL so a human can
    # click "Download" manually; this never fabricates a link.
    $patterns = @(
        '"downloadUrl"\s*:\s*"([^"]+)"',
        'data-bi-cn="download"[^>]*href="([^"]+)"',
        'href="([^"]+\.msi)"',
        'href="([^"]+\.exe)"'
    )

    foreach ($p in $patterns) {
        $m = [regex]::Match($html, $p, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if ($m.Success) {
            return ($m.Groups[1].Value -replace '\\/', '/')
        }
    }

    return $null
}

function Save-CatalogEntry {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$DetailsUrl,
        [Parameter(Mandatory)][string]$DownloadsFolder
    )

    $fileUrl = Resolve-CatalogDownloadFileUrl -DetailsUrl $DetailsUrl
    if (-not $fileUrl) {
        Write-Log "Could not resolve a direct file URL for '$Name' -- get it manually from $DetailsUrl" -Level WARN
        return [PSCustomObject]@{ ResolvedFileUrl = $null; DownloadedTo = $null; DownloadStatus = 'Resolution Failed - use Details URL manually' }
    }

    $safeName = ($Name -replace '[^a-zA-Z0-9\.\-_ ]', '') -replace '\s+', '_'
    $ext = [System.IO.Path]::GetExtension($fileUrl)
    if ([string]::IsNullOrWhiteSpace($ext) -or $ext.Length -gt 5) { $ext = '.msi' }
    $destPath = Join-Path $DownloadsFolder "$safeName$ext"

    try {
        Invoke-WebRequest -Uri $fileUrl -OutFile $destPath -UseBasicParsing -ErrorAction Stop
        Write-Log "Downloaded '$Name' -> $destPath" -Level SUCCESS
        return [PSCustomObject]@{ ResolvedFileUrl = $fileUrl; DownloadedTo = $destPath; DownloadStatus = 'Downloaded' }
    }
    catch {
        Write-Log "Resolved a file URL for '$Name' but the download failed: $($_.Exception.Message). Try manually: $fileUrl" -Level WARN
        return [PSCustomObject]@{ ResolvedFileUrl = $fileUrl; DownloadedTo = $null; DownloadStatus = "Download Failed: $($_.Exception.Message)" }
    }
}

###########################################################
# STEP 0 - SEALED MP (.mp) EXTRACTION SUPPORT
###########################################################
# Sealed Management Packs (.mp) are NOT plain XML -- they are a compiled
# binary container (XML + resources). [xml]Get-Content cannot read them.
# To get at the XML we need the SCOM SDK assemblies that ship with the
# console / PowerShell module. We discover and load them once, then use
# the documented unseal pattern (Microsoft.EnterpriseManagement.Configuration
# .ManagementPack + ManagementPackXmlWriter) to read each .mp and write a
# plain .xml working copy -- the same approach Microsoft has published since
# SCOM 2007 (Boris Yanushpolsky's MpToXml.ps1), unchanged through SCOM 2025.
#
# If the SDK cannot be located/loaded, .mp files are logged and skipped --
# the rest of the batch (native .xml files) still processes normally.

$script:SdkLoaded = $false

function Initialize-ScomSdk {

    if ($SkipSealedExtraction) {
        Write-Log "Sealed (.mp) extraction explicitly skipped via -SkipSealedExtraction." -Level WARN
        return $false
    }

    # Common install locations for the SCOM SDK assemblies, newest-first.
    # We search broadly (Console, Server, PowerShell module dirs) because
    # the exact path varies by SCOM version and whether this runs on a
    # management server vs. a console-only workstation.
    $probeRoots = @(
        "C:\Program Files\Microsoft System Center\Operations Manager\Console\",
        "C:\Program Files\Microsoft System Center\Operations Manager\Server\",
        "C:\Program Files\Microsoft System Center 2025\Operations Manager\Console\",
        "C:\Program Files\Microsoft System Center 2025\Operations Manager\Server\",
        "C:\Program Files (x86)\Microsoft System Center\Operations Manager\Console\",
        "$env:ProgramFiles\WindowsPowerShell\Modules\OperationsManager\"
    )

    # The unseal/dump-to-XML pattern used below (ManagementPack +
    # ManagementPackXmlWriter, both in the Configuration namespace) lives in
    # Microsoft.EnterpriseManagement.OperationsManager.dll, with Core.dll as
    # its dependency.
    $neededAssemblies = @(
        "Microsoft.EnterpriseManagement.Core.dll",
        "Microsoft.EnterpriseManagement.OperationsManager.dll",
        # Optional -- only needed to read .mpb bundles. Missing is not fatal.
        "Microsoft.EnterpriseManagement.Packaging.dll"
    )

    $foundPaths = @{}

    foreach ($root in $probeRoots) {
        if (-not (Test-Path -LiteralPath $root)) { continue }

        foreach ($asmName in $neededAssemblies) {
            if ($foundPaths.ContainsKey($asmName)) { continue }

            $hit = Get-ChildItem -LiteralPath $root -Filter $asmName -Recurse -ErrorAction SilentlyContinue |
                Select-Object -First 1

            if ($hit) {
                $foundPaths[$asmName] = $hit.FullName
            }
        }
    }

    if (-not $foundPaths.ContainsKey("Microsoft.EnterpriseManagement.OperationsManager.dll")) {
        Write-Log "Could not locate Microsoft.EnterpriseManagement.OperationsManager.dll under any probed SCOM install path. Sealed (.mp) files will be skipped." -Level WARN
        return $false
    }

    try {
        foreach ($asmName in $neededAssemblies) {
            if ($foundPaths.ContainsKey($asmName)) {
                try {
                    Add-Type -Path $foundPaths[$asmName] -ErrorAction Stop
                    Write-Log "Loaded SDK assembly: $($foundPaths[$asmName])" -NoConsole
                }
                catch {
                    if ($asmName -like '*Packaging*') {
                        Write-Log "Packaging assembly could not be loaded -- .mpb bundles will be skipped: $($_.Exception.Message)" -Level WARN
                    }
                    else { throw }
                }
            }
        }
        return $true
    }
    catch {
        Write-Log "Failed to load SCOM SDK assemblies: $($_.Exception.Message). Sealed (.mp) files will be skipped." -Level WARN
        return $false
    }
}

function Convert-SealedMPToXml {
    param(
        [Parameter(Mandatory)][string]$MpFilePath,
        [Parameter(Mandatory)][string]$WorkingFolder
    )

    # Returns the path to an extracted .xml working copy, or $null on failure.
    #
    # Uses the documented unseal pattern (same one Microsoft has shipped
    # since SCOM 2007, still valid through SCOM 2025's SDK):
    #   $mp     = New-Object Microsoft.EnterpriseManagement.Configuration.ManagementPack($MpFilePath)
    #   $writer = New-Object Microsoft.EnterpriseManagement.Configuration.IO.ManagementPackXmlWriter($folder)
    #   $writer.WriteManagementPack($mp)
    # WriteManagementPack writes "<MP_ID>.xml" into $WorkingFolder and returns
    # that ID as a string -- it does not take an explicit output filename, so
    # we resolve the actual written path afterward rather than assuming the
    # source .mp file's base name matches the MP's ID (they often don't).
    try {
        $mpObject = New-Object Microsoft.EnterpriseManagement.Configuration.ManagementPack($MpFilePath)

        $writer    = New-Object Microsoft.EnterpriseManagement.Configuration.IO.ManagementPackXmlWriter($WorkingFolder)
        $writtenId = $writer.WriteManagementPack($mpObject)

        $expectedPath = Join-Path $WorkingFolder "$writtenId.xml"

        if (Test-Path -LiteralPath $expectedPath) {
            return $expectedPath
        }

        # Fallback: WriteManagementPack's return value didn't match the
        # filename it actually wrote (seen on some SDK builds). Find the
        # newest .xml file in the working folder as a best-effort recovery
        # rather than failing outright.
        $newestXml = Get-ChildItem -LiteralPath $WorkingFolder -Filter *.xml -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending |
            Select-Object -First 1

        if ($newestXml) {
            Write-Log "Note: extracted XML for '$MpFilePath' was located by timestamp ('$($newestXml.Name)') rather than the expected '$writtenId.xml'." -Level WARN -NoConsole
            return $newestXml.FullName
        }

        throw "ManagementPackXmlWriter reported writing '$writtenId' but no matching .xml file was found in '$WorkingFolder'."
    }
    catch {
        Write-Log "Failed to extract sealed MP '$MpFilePath': $($_.Exception.Message)" -Level WARN
        return $null
    }
}

function Import-MPFile {
    param(
        [Parameter(Mandatory)][System.IO.FileSystemInfo]$File,
        [Parameter(Mandatory)][string]$WorkingFolder
    )

    # Shared by Step 0 (initial batch ingestion) and the dependency
    # auto-resolution pass (Step 2.5) below -- loading a single .xml or
    # sealed .mp file into a [PSCustomObject] with its parsed XML, MP ID,
    # and version, or $null if it could not be loaded. Centralized here so
    # both call sites unseal/parse identically rather than drifting apart.

    # Returns an ARRAY of loaded MP objects (empty on failure). A .xml or .mp
    # file yields one; a .mpb bundle can yield several, all sharing the bundle
    # as their SourcePath (Import-SCOMManagementPack imports the whole bundle).

    $xmlPathToLoad = $null

    if ($File.Extension -ieq ".mpb") {
        if (-not $script:SdkLoaded) { return @() }
        $bundleResults = @()
        try {
            $reader = [Microsoft.EnterpriseManagement.Packaging.ManagementPackBundleFactory]::CreateBundleReader()
            $store  = New-Object Microsoft.EnterpriseManagement.Configuration.IO.ManagementPackFileStore
            $store.AddDirectory($File.DirectoryName)
            $bundle = $reader.Read($File.FullName, $store)
            $writer = New-Object Microsoft.EnterpriseManagement.Configuration.IO.ManagementPackXmlWriter($WorkingFolder)
            foreach ($bmp in $bundle.ManagementPacks) {
                $writtenId = [string]$writer.WriteManagementPack($bmp)
                $bxml = Join-Path $WorkingFolder "$writtenId.xml"
                if (-not (Test-Path -LiteralPath $bxml)) {
                    # Some SDK builds return the full written path instead of the ID.
                    if (Test-Path -LiteralPath $writtenId) { $bxml = $writtenId }
                    elseif (Test-Path -LiteralPath (Join-Path $WorkingFolder "$($bmp.Name).xml")) { $bxml = Join-Path $WorkingFolder "$($bmp.Name).xml" }
                    else { Write-Log "Bundle '$($File.Name)': could not locate extracted XML for '$($bmp.Name)' -- skipped." -Level WARN; continue }
                }
                $bdoc = Read-MPXml -Path $bxml
                $bundleResults += [PSCustomObject]@{
                    ID             = [string]$bdoc.ManagementPack.Manifest.Identity.ID
                    Version        = [string]$bdoc.ManagementPack.Manifest.Identity.Version
                    SourcePath     = [string]$File.FullName
                    WorkingXmlPath = [string]$bxml
                    WasSealed      = $true
                    IsBundle       = $true
                    XML            = $bdoc
                }
            }
        }
        catch {
            Write-Log "Failed to read bundle '$($File.FullName)': $($_.Exception.Message)" -Level WARN
        }
        return $bundleResults
    }

    if ($File.Extension -ieq ".xml") {
        $xmlPathToLoad = $File.FullName
    }
    elseif ($File.Extension -ieq ".mp") {
        if (-not $script:SdkLoaded) {
            return @()
        }

        $extracted = Convert-SealedMPToXml -MpFilePath $File.FullName -WorkingFolder $WorkingFolder
        if (-not $extracted) {
            return @()
        }
        $xmlPathToLoad = $extracted
    }
    else {
        return @()
    }

    try {
        $candidateXml = Read-MPXml -Path $xmlPathToLoad

        $id  = $candidateXml.ManagementPack.Manifest.Identity.ID
        $ver = $candidateXml.ManagementPack.Manifest.Identity.Version

        if (-not $id) {
            Write-Log "Skipping '$($File.FullName)': no Manifest.Identity.ID found (not a valid MP XML)." -Level WARN
            return @()
        }

        return [PSCustomObject]@{
            ID             = [string]$id
            Version        = [string]$ver
            SourcePath     = [string]$File.FullName
            WorkingXmlPath = [string]$xmlPathToLoad
            WasSealed      = ($File.Extension -ieq ".mp")
            IsBundle       = $false
            XML            = $candidateXml
        }
    }
    catch {
        Write-Log "Failed to load '$($File.FullName)' as XML: $($_.Exception.Message)" -Level WARN
        return @()
    }
}

###########################################################
# STEP 0 - BATCH INGESTION
###########################################################
# Skipped entirely in -CatalogOnly mode: there is no batch to ingest when
# all you're doing is a catalog check against -RepositoryFolder. Step 1
# (repository load) and Step 1.5 (catalog check) still run below; the script
# returns at the end of Step 1.5 before any batch-dependent step.
if (-not $CatalogOnly) {

Write-Log ""
Write-Log "Resolving batch input..." -Level WARN

$BatchWorkingFolder = Join-Path $OutputFolder "_ExtractedSource"
if (-not (Test-Path -LiteralPath $BatchWorkingFolder)) {
    New-Item -ItemType Directory -Path $BatchWorkingFolder -Force | Out-Null
}

# $InputPath is always an array at this point (one entry for a single file
# or folder, several entries when multiple files were selected/passed).
# Each entry is resolved independently -- a folder is scanned recursively,
# a file is taken as-is -- and the results are pooled into one candidate
# list, same as if everything had been dropped into one folder.
$candidateFiles = @()

foreach ($p in $InputPath) {

    $inputItem = Get-Item -LiteralPath $p

    if ($inputItem.PSIsContainer) {
        Write-Log "Input '$p' is a folder. Scanning recursively for .xml, .mp and .mpb files..."
        $candidateFiles += Get-ChildItem -LiteralPath $p -Filter *.xml -Recurse -ErrorAction SilentlyContinue
        $candidateFiles += Get-ChildItem -LiteralPath $p -Filter *.mp   -Recurse -ErrorAction SilentlyContinue
        $candidateFiles += Get-ChildItem -LiteralPath $p -Filter *.mpb  -Recurse -ErrorAction SilentlyContinue
    }
    else {
        Write-Log "Input '$p' is a single file."
        $candidateFiles += $inputItem
    }
}

# Never treat our own outputs as input if someone points -InputPath at a
# parent of -OutputFolder.
$outFull = [System.IO.Path]::GetFullPath($OutputFolder).TrimEnd('\', '/')
$candidateFiles = @($candidateFiles | Where-Object { -not ([System.IO.Path]::GetFullPath($_.FullName).StartsWith($outFull + [System.IO.Path]::DirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase)) })

if (-not $candidateFiles -or $candidateFiles.Count -eq 0) {
    throw "No .xml, .mp or .mpb Management Pack files were found under: $($InputPath -join ', ')"
}

Write-Log "Candidate source files found : $($candidateFiles.Count)"

$sealedCount = @($candidateFiles | Where-Object { $_.Extension -ieq ".mp" -or $_.Extension -ieq ".mpb" }).Count
if ($sealedCount -gt 0) {
    Write-Log "Sealed (.mp/.mpb) files detected : $sealedCount -- attempting SDK extraction" -Level WARN
    $script:SdkLoaded = Initialize-ScomSdk
}
else {
    Write-Log "No sealed (.mp/.mpb) files detected; SDK extraction not required."
}

$skippedSealed = @()

foreach ($file in $candidateFiles) {

    $isSealedFile = ($file.Extension -ieq ".mp" -or $file.Extension -ieq ".mpb")

    if ($isSealedFile -and -not $script:SdkLoaded) {
        $skippedSealed += $file.FullName
        continue
    }

    $loadedList = @(Import-MPFile -File $file -WorkingFolder $BatchWorkingFolder)

    if ($loadedList.Count -eq 0) {
        if ($isSealedFile) { $skippedSealed += $file.FullName }
        continue
    }

    foreach ($loaded in $loadedList) {
        if ($BatchSource.ContainsKey($loaded.ID)) {
            # Prefer the SEALED original over an exported .xml of the same MP:
            # only the sealed file can actually be imported for a sealed MP.
            $existing = $BatchSource[$loaded.ID]
            $replace = $false
            if ($loaded.WasSealed -and -not $existing.WasSealed) { $replace = $true }
            elseif ($loaded.WasSealed -eq $existing.WasSealed) {
                try { if ([version]$loaded.Version -gt [version]$existing.Version) { $replace = $true } } catch { }
            }
            if ($replace) {
                Write-Log "Duplicate MP ID '$($loaded.ID)': using '$($loaded.SourcePath)' (v$($loaded.Version)$(if ($loaded.WasSealed) { ', sealed original' })) instead of '$($existing.SourcePath)' (v$($existing.Version))." -Level WARN -NoConsole
                $BatchSource[$loaded.ID] = $loaded
            }
            else {
                Write-Log "Duplicate MP ID '$($loaded.ID)' (source: $($loaded.SourcePath)) ignored; keeping '$($existing.SourcePath)'." -Level WARN -NoConsole
            }
            continue
        }

        $BatchSource[$loaded.ID] = $loaded
    }
}

if ($skippedSealed.Count -gt 0) {
    Write-Log ""
    Write-Log "$($skippedSealed.Count) sealed (.mp/.mpb) file(s) were skipped (SDK unavailable or extraction failed):" -Level WARN
    $skippedSealed | ForEach-Object { Write-Log "  - $_" -Level WARN }
}

if ($BatchSource.Count -eq 0) {
    throw "No usable Management Packs were loaded from the input batch. Cannot continue."
}

###########################################################
# STEP 0.5 - MANIFEST FILTER (-Manifest)
###########################################################
# The disposition workbook decides what moves. -InputPath can then simply be
# the full 2016 export (plus a folder of sealed originals); only rows marked
# Migrate = Y are kept.

function Get-MPSelfDisplayName {
    param([System.Xml.XmlDocument]$Doc, [string]$Id)
    try {
        $n = $Doc.SelectSingleNode("//LanguagePacks/LanguagePack[@ID='ENU']/DisplayStrings/DisplayString[@ElementID='$Id']/Name")
        if ($n -and $n.InnerText) { return [string]$n.InnerText }
        $n = $Doc.SelectSingleNode("/ManagementPack/Manifest/Name")
        if ($n -and $n.InnerText) { return [string]$n.InnerText }
    }
    catch { }
    return ''
}

$script:ManifestActionMap = @{}   # MPID -> manifest Action (MIGRATE / VENDOR_SEALED / OVERRIDES)
$script:ManifestNameMap   = @{}   # MPID -> workbook ManagementPack name

if ($Manifest) {
    if (-not (Test-Path -LiteralPath $Manifest)) { throw "Manifest not found: $Manifest" }
    $manifestRows = @(Import-Csv -LiteralPath $Manifest)
    if ($manifestRows.Count -eq 0 -or -not $manifestRows[0].PSObject.Properties['Migrate'] -or -not $manifestRows[0].PSObject.Properties['ManagementPack']) {
        throw "Manifest '$Manifest' must have at least the columns 'Migrate' and 'ManagementPack' (MatchPattern and Action are optional)."
    }
    $wantedRows = @($manifestRows | Where-Object { [string]$_.Migrate -match '^(y|yes|true|1)$' })
    Write-Log ""
    Write-Log "Manifest: $($wantedRows.Count) of $($manifestRows.Count) row(s) marked Migrate = Y." -Level WARN

    $displayNames = @{}
    foreach ($bid in @($BatchSource.Keys)) { $displayNames[$bid] = Get-MPSelfDisplayName -Doc $BatchSource[$bid].XML -Id $bid }

    $keep = @{}
    $matchReport = New-Object System.Collections.Generic.List[object]
    foreach ($row in $wantedRows) {
        $pat = if ($row.PSObject.Properties['MatchPattern'] -and $row.MatchPattern) { [string]$row.MatchPattern } else { [string]$row.ManagementPack }
        $pat = $pat.Trim()
        $act = if ($row.PSObject.Properties['Action'] -and $row.Action) { [string]$row.Action } else { 'MIGRATE' }
        # A workbook name containing [ or ] is not a valid wildcard -- match it literally.
        try { [void]('x' -like $pat) } catch { $pat = [System.Management.Automation.WildcardPattern]::Escape($pat) }
        $hits = @($BatchSource.Keys | Where-Object { $_ -like $pat })
        $how = 'MPID'
        if ($hits.Count -eq 0) {
            $hits = @($BatchSource.Keys | Where-Object { $displayNames[$_] -and $displayNames[$_] -like $pat })
            $how = 'DisplayName'
        }
        foreach ($h in $hits) {
            $keep[$h] = $true
            $script:ManifestActionMap[$h] = $act
            $script:ManifestNameMap[$h] = [string]$row.ManagementPack
        }
        $matchReport.Add([PSCustomObject]@{
            ManagementPack = $row.ManagementPack
            MatchPattern   = $pat
            Action         = $act
            Status         = if ($hits.Count -eq 0) { 'NO MATCH in -InputPath' } elseif ($hits.Count -gt 1) { "MATCHED $($hits.Count) MPs" } else { 'Matched' }
            MatchedBy      = if ($hits.Count -gt 0) { $how } else { '' }
            MatchedMPIDs   = ($hits -join '; ')
            SealedFile     = (@($hits | ForEach-Object { $BatchSource[$_].WasSealed }) -join '; ')
        })
    }

    $dropped = 0
    foreach ($bid in @($BatchSource.Keys)) {
        if (-not $keep.ContainsKey($bid)) { $BatchSource.Remove($bid); $dropped++ }
    }

    $matchFile = Join-Path $OutputFolder "ManifestMatch.csv"
    $matchReport | Export-Csv -LiteralPath $matchFile -NoTypeInformation -Encoding UTF8
    $noMatch = @($matchReport | Where-Object { $_.Status -like 'NO MATCH*' })
    Write-Log "Manifest filter: kept $($BatchSource.Count) MP(s), ignored $dropped not marked Migrate = Y. Match report: $matchFile" -Level SUCCESS
    if ($noMatch.Count -gt 0) {
        Write-Log "$($noMatch.Count) manifest row(s) marked Migrate = Y matched NOTHING in -InputPath (name differs from the MP ID/display name, or the MP was not exported):" -Level WARN
        foreach ($nm in $noMatch) { Write-Log "  - $($nm.ManagementPack)  [pattern: $($nm.MatchPattern)]" -Level WARN }
    }
    if ($BatchSource.Count -eq 0) {
        throw "The manifest matched none of the MPs under -InputPath. Check $matchFile."
    }
}

# Sealed-in-source MPs that only arrived here as an exported .xml cannot be
# imported (the signature is gone). Flag them now so they are reported as
# BLOCKED with a clear reason instead of failing at import time.
foreach ($bid in @($BatchSource.Keys)) {
    $e = $BatchSource[$bid]
    $needsOriginal = $false
    if (-not $e.WasSealed) {
        if ($script:SourceSealedMap.ContainsKey($bid)) { $needsOriginal = $true }
        elseif ($script:ManifestActionMap.ContainsKey($bid) -and $script:ManifestActionMap[$bid] -eq 'VENDOR_SEALED') { $needsOriginal = $true }
    }
    $e | Add-Member -NotePropertyName NeedsSealedOriginal -NotePropertyValue $needsOriginal -Force
}
$needsOrigCount = @($BatchSource.Values | Where-Object { $_.NeedsSealedOriginal }).Count
if ($needsOrigCount -gt 0) {
    Write-Log "$needsOrigCount MP(s) were SEALED in the source but only an exported .xml was supplied. They will be analysed but marked BLOCKED until the original .mp/.mpb is added to -InputPath:" -Level WARN
    foreach ($e in @($BatchSource.Values | Where-Object { $_.NeedsSealedOriginal } | Sort-Object ID)) { Write-Log "  - $($e.ID) (v$($e.Version))" -Level WARN }
}

Write-Log ""
Write-Log "Batch Source MPs Loaded : $($BatchSource.Count)" -Level SUCCESS
foreach ($bid in $BatchSource.Keys | Sort-Object) {
    Write-Log "  - $bid (v$($BatchSource[$bid].Version))$(if ($BatchSource[$bid].WasSealed) { ' [sealed original]' })" -NoConsole
}

# Snapshot of the ORIGINAL -InputPath targets, before Step 2.5 can add any
# auto-resolved dependencies to $BatchSource. -LiveImport (Step 3.5) needs
# this distinction: dependencies get imported live as-is, but the MP(s) you
# actually named here are never auto-imported -- they always go through the
# normal rewrite/candidate pipeline, regardless of -LiveImport.
$OriginalInputTargets = @($BatchSource.Keys)

# Declared unconditionally here (not just inside Step 2.5's -SourceRepository
# Folder branch) so it's always safe to reference downstream -- e.g. by the
# missing-dependencies report -- regardless of whether a source repository
# was even provided this run.
$script:ExcludedFromSource = @{}   # refId -> $true, for IDs explicitly skipped via -ExcludeMPIDs

Write-Banner "STEP 0 COMPLETE"
Write-Log "Batch MPs to migrate : $($BatchSource.Count)"
Write-Log "Sealed extracted     : $(@($BatchSource.Values | Where-Object { $_.WasSealed }).Count)"
Write-Log "Sealed skipped       : $($skippedSealed.Count)"
Write-Log "Output folder        : $OutputFolder"
Write-Log "Candidate MP folder  : $CandidateFolder"

}  # end: if (-not $CatalogOnly) -- Step 0 batch ingestion skipped in catalog-only mode

###########################################################
# STEP 1 - LOAD TARGET REPOSITORY
###########################################################

Write-Banner "STEP 1 - LOADING TARGET REPOSITORY"

Write-Log ""
Write-Log "Loading Repository MPs..." -Level WARN

$Files = Get-ChildItem -Path $RepositoryFolder -Filter *.xml -Recurse -ErrorAction SilentlyContinue

if (-not $Files -or $Files.Count -eq 0) {
    Write-Log "No XML files found under '$RepositoryFolder'." -Level WARN
}

foreach ($file in $Files) {

    try {
        $mp = Read-MPXml -Path $file.FullName

        $id  = $mp.ManagementPack.Manifest.Identity.ID
        $ver = $mp.ManagementPack.Manifest.Identity.Version

        if (-not $id) { continue }

        # Same ID exported twice (e.g. two exports merged into one folder):
        # keep the HIGHER version, since that is what is really installed.
        if ($Repository.ContainsKey([string]$id)) {
            $keepExisting = $true
            try { $keepExisting = ([version]$Repository[[string]$id].Version -ge [version]$ver) } catch { }
            if ($keepExisting) { continue }
        }

        $Repository[[string]$id] = [PSCustomObject]@{
            ID      = [string]$id
            Version = [string]$ver
            Path    = [string]$file.FullName
            XML     = $mp
        }

    }
    catch {
        $msg = "Failed loading '$($file.FullName)': $($_.Exception.Message)"
        $Warnings += $msg
        Write-Log $msg -Level WARN -NoConsole
    }
}

Write-Log "Repository MPs Loaded : $($Repository.Count)" -Level SUCCESS

if ($Repository.Count -eq 0) {
    throw "No usable Management Packs were loaded from the repository folder. Cannot continue."
}

###########################################################
# BUILD VERSION MAP
###########################################################

Write-Log ""
Write-Log "Building Version Map..." -Level WARN

foreach ($entry in $Repository.Values) {
    if ($null -eq $entry -or -not $entry.ID) { continue }
    $VersionMap[$entry.ID] = [string]$entry.Version
}

Write-Log "Version Entries       : $($VersionMap.Count)"

Write-Banner "STEP 1 COMPLETE"
Write-Log "Repository MPs        : $($Repository.Count)"
Write-Log "Version Map Entries   : $($VersionMap.Count)"

if ($Warnings.Count -gt 0) {
    Write-Log ""
    Write-Log "Warnings so far:" -Level WARN
    $Warnings | Sort-Object -Unique | ForEach-Object { Write-Log "  - $_" -Level WARN }
}

###########################################################
# STEP 1.5 - TARGET REPOSITORY CATALOG CHECK (-CheckCatalog / -CatalogOnly)
###########################################################
if ($CheckCatalog) {

    Write-Banner "STEP 1.5 - CHECKING MICROSOFT'S OFFICIAL MP CATALOG"
    Write-Log "CatalogView            : $CatalogView"
    Write-Log "CatalogFilter          : $(if ($CatalogFilter) { $CatalogFilter -join ', ' } else { '(none -- full catalog)' })"
    Write-Log "DownloadFromCatalog    : $($DownloadFromCatalog.IsPresent) (experimental if enabled)"
    Write-Log "StrictMatch            : $($CatalogStrictMatch.IsPresent)$(if ($CatalogStrictMatch) { ' -- soft matches to generic family libraries will be reported as gaps, not matches' })"

    # Propagate the strict-match switch to the script-scoped flag the scorer
    # reads. Done here (not at param time) so it sits right next to the rest
    # of the catalog-run setup and is obvious in context.
    $script:CatalogStrictMatchActive = $CatalogStrictMatch.IsPresent

    $catalogFolder = Join-Path $OutputFolder "MPCatalog"
    if (-not (Test-Path -LiteralPath $catalogFolder)) {
        New-Item -ItemType Directory -Path $catalogFolder -Force | Out-Null
    }

    $rawCatalogRows = @(Get-MPCatalog -View $CatalogView)
    $catalogRows = @(Merge-DuplicateDownloadIds -Rows $rawCatalogRows)
    Write-Log "Catalog contains $($catalogRows.Count) unique download entries after de-duplication (from $($rawCatalogRows.Count) raw rows)."

    if ($CatalogFilter -and $CatalogFilter.Count -gt 0) {
        $catalogRows = @($catalogRows | Where-Object {
            $entry = $_
            $haystack = "$($entry.Name) $($entry.AlsoKnownAs)"
            $matched = $false
            foreach ($kw in $CatalogFilter) {
                if ($haystack -like "*$kw*") { $matched = $true; break }
            }
            $matched
        })
        Write-Log "$($catalogRows.Count) catalog entries remain after applying -CatalogFilter."
    }

    if ($catalogRows.Count -eq 0) {
        Write-Log "No catalog entries left to check -- verify your -CatalogFilter keywords." -Level WARN
    }
    else {
        # Build local display-name index from the ALREADY-LOADED $Repository
        # (Step 1) rather than re-reading files from disk -- keeps this
        # perfectly in sync with what the rest of the run considers "the
        # repository," and avoids a second, possibly-divergent file scan.
        $localCatalogEntries = @($Repository.Values | ForEach-Object {
            [PSCustomObject]@{
                MPId        = $_.ID
                DisplayName = Get-RepositoryDisplayName -RepositoryEntry $_
                Version     = $_.Version
            }
        })

        $catalogReport = New-Object System.Collections.Generic.List[object]
        $ci = 0
        foreach ($entry in $catalogRows) {
            $ci++
            Write-Progress -Activity "Checking Microsoft MP catalog against your repository" -Status $entry.Name -PercentComplete (($ci / $catalogRows.Count) * 100)
            $match = Find-BestCatalogMatch -CatalogName $entry.Name -LocalEntries $localCatalogEntries -Threshold $CatalogMatchThreshold
            $catalogReport.Add([PSCustomObject]@{
                CatalogName     = $entry.Name
                AlsoKnownAs     = $entry.AlsoKnownAs
                CatalogVersion  = $entry.Version
                CatalogDate     = $entry.Date
                DownloadPageUrl = $entry.Url
                LocalStatus     = $match.Status
                LocalMPId       = $match.LocalMPId
                LocalVersion    = $match.LocalVersion
                MatchScore      = $match.Score
                ResolvedFileUrl = $null
                DownloadedTo    = $null
                DownloadStatus  = $null
            })
        }
        Write-Progress -Activity "Checking Microsoft MP catalog against your repository" -Completed

        $catalogMatchedCount = @($catalogReport | Where-Object { $_.LocalStatus -eq 'Likely Match' }).Count
        $catalogMissingCount = @($catalogReport | Where-Object { $_.LocalStatus -eq 'No Local Match Found' }).Count
        Write-Log "Catalog gap-check complete: $catalogMatchedCount likely already in -RepositoryFolder, $catalogMissingCount with no local match found." -Level SUCCESS

        if ($DownloadFromCatalog) {
            Write-Banner "STEP 1.5 - DOWNLOADING MISSING MPs FROM CATALOG (EXPERIMENTAL)"
            $catalogDownloadsFolder = Join-Path $catalogFolder "Downloads"
            if (-not (Test-Path -LiteralPath $catalogDownloadsFolder)) {
                New-Item -ItemType Directory -Path $catalogDownloadsFolder -Force | Out-Null
            }

            $catalogToDownload = @($catalogReport | Where-Object { $_.LocalStatus -eq 'No Local Match Found' })
            Write-Log "Attempting to resolve and download $($catalogToDownload.Count) installer(s) into '$catalogDownloadsFolder'. This is best-effort -- see -DownloadFromCatalog help for why."

            $di = 0
            foreach ($row in $catalogToDownload) {
                $di++
                Write-Progress -Activity "Downloading installers" -Status $row.CatalogName -PercentComplete (($di / $catalogToDownload.Count) * 100)
                $dlResult = Save-CatalogEntry -Name $row.CatalogName -DetailsUrl $row.DownloadPageUrl -DownloadsFolder $catalogDownloadsFolder
                $row.ResolvedFileUrl = $dlResult.ResolvedFileUrl
                $row.DownloadedTo = $dlResult.DownloadedTo
                $row.DownloadStatus = $dlResult.DownloadStatus
            }
            Write-Progress -Activity "Downloading installers" -Completed

            $catalogDownloadedCount = @($catalogToDownload | Where-Object { $_.DownloadStatus -eq 'Downloaded' }).Count
            Write-Log "Downloaded $catalogDownloadedCount of $($catalogToDownload.Count) installer(s). Anything not downloaded still has its Download Center URL in MPCatalog.csv for you to grab by hand." -Level $(if ($catalogDownloadedCount -lt $catalogToDownload.Count) { 'WARN' } else { 'SUCCESS' })
            Write-Log "Reminder: downloaded files are installers only -- nothing has been extracted or imported. Run the installer (or 'msiexec /a <file> TARGETDIR=<folder>' for an admin extraction without installing), export/copy the resulting MP(s), then re-run with -RepositoryFolder pointed at wherever they land."
        }

        $catalogCsvPath = Join-Path $catalogFolder "MPCatalog.csv"
        $catalogReport | Export-Csv -LiteralPath $catalogCsvPath -NoTypeInformation -Encoding UTF8
        Write-Log "Wrote $($catalogReport.Count) row(s) to $catalogCsvPath" -Level SUCCESS
    }

    Write-Banner "STEP 1.5 COMPLETE"

    if ($CatalogOnly) {
        Write-Log "Stopping here: -CatalogOnly was specified, so no batch is being processed. Re-run without -CatalogOnly (with -InputPath) to actually migrate MP(s)." -Level SUCCESS
        return
    }
}

###########################################################
# STEP 2 - SYMBOL TABLE BUILDER (target repository)
###########################################################

Write-Banner "STEP 2 - BUILDING SYMBOL TABLE"

$Symbols = @{
    Classes     = @{}
    Rules       = @{}
    Monitors    = @{}
    Discoveries = @{}
    Tasks       = @{}
}

$IndexStats = @{
    Classes     = 0
    Rules       = 0
    Monitors    = 0
    Discoveries = 0
    Tasks       = 0
}

foreach ($repoEntry in $Repository.Values) {

    $repoMp = $repoEntry.XML

    # -------------------------
    # CLASSES
    # -------------------------
    try {
        $classNodes = $repoMp.ManagementPack.TypeDefinitions.EntityTypes.ClassTypes.ClassType

        foreach ($c in $classNodes) {
            if ($c.ID) {
                $Symbols.Classes[$c.ID] = $repoEntry.ID
                $IndexStats.Classes++
            }
        }
    } catch {
        Write-Log "Class enumeration failed for $($repoEntry.ID): $($_.Exception.Message)" -Level WARN -NoConsole
    }

    # -------------------------
    # RULES
    # -------------------------
    try {
        $ruleNodes = $repoMp.ManagementPack.Monitoring.Rules.Rule

        foreach ($r in $ruleNodes) {
            if ($r.ID) {
                $Symbols.Rules[$r.ID] = $repoEntry.ID
                $IndexStats.Rules++
            }
        }
    } catch {
        Write-Log "Rule enumeration failed for $($repoEntry.ID): $($_.Exception.Message)" -Level WARN -NoConsole
    }

    # -------------------------
    # MONITORS
    # -------------------------
    try {
        $monitorNodes = $repoMp.ManagementPack.Monitoring.Monitors.ChildNodes |
            Where-Object { $_.NodeType -eq [System.Xml.XmlNodeType]::Element }

        foreach ($m in $monitorNodes) {
            if ($m.ID) {
                $Symbols.Monitors[$m.ID] = $repoEntry.ID
                $IndexStats.Monitors++
            }
        }
    } catch {
        Write-Log "Monitor enumeration failed for $($repoEntry.ID): $($_.Exception.Message)" -Level WARN -NoConsole
    }

    # -------------------------
    # DISCOVERIES
    # -------------------------
    try {
        $discNodes = $repoMp.ManagementPack.Monitoring.Discoveries.Discovery

        foreach ($d in $discNodes) {
            if ($d.ID) {
                $Symbols.Discoveries[$d.ID] = $repoEntry.ID
                $IndexStats.Discoveries++
            }
        }
    } catch {
        Write-Log "Discovery enumeration failed for $($repoEntry.ID): $($_.Exception.Message)" -Level WARN -NoConsole
    }

    # -------------------------
    # TASKS (SAFE VERSION)
    # -------------------------
    try {
        $taskNodes = @()

        if ($repoMp.ManagementPack.Tasks) {
            $taskNodes += $repoMp.ManagementPack.Tasks.ChildNodes |
                Where-Object { $_.NodeType -eq [System.Xml.XmlNodeType]::Element }
        }

        if ($repoMp.ManagementPack.Monitoring.Tasks) {
            $taskNodes += $repoMp.ManagementPack.Monitoring.Tasks.ChildNodes |
                Where-Object { $_.NodeType -eq [System.Xml.XmlNodeType]::Element }
        }

        foreach ($t in $taskNodes) {
            if ($t.ID) {
                $Symbols.Tasks[$t.ID] = $repoEntry.ID
                $IndexStats.Tasks++
            }
        }
    } catch {
        Write-Log "Task enumeration failed for $($repoEntry.ID): $($_.Exception.Message)" -Level WARN -NoConsole
    }
}

Write-Log ""
Write-Log "Symbol Table Build Complete" -Level SUCCESS
Write-Log "------------------------------------------"
Write-Log "Classes     : $($IndexStats.Classes)"
Write-Log "Rules       : $($IndexStats.Rules)"
Write-Log "Monitors    : $($IndexStats.Monitors)"
Write-Log "Discoveries : $($IndexStats.Discoveries)"
Write-Log "Tasks       : $($IndexStats.Tasks)"

# Per-MP element index (MPID -> set of every element ID it defines). This is
# what lets the candidate check say "the Windows Server MP IS installed in
# 2025, but monitor X that this override targets no longer exists in it".
$ElementIndex = @{}
foreach ($repoEntry in $Repository.Values) {
    try { $ElementIndex[$repoEntry.ID] = Get-MPElementIdSet -Doc $repoEntry.XML }
    catch { Write-Log "Element indexing failed for $($repoEntry.ID): $($_.Exception.Message)" -Level WARN -NoConsole }
}
Write-Log "Element index built for $($ElementIndex.Count) target MP(s)." -NoConsole

###########################################################
# STEP 2.5 - AUTO-RESOLVE MISSING DEPENDENCIES FROM SOURCE REPOSITORY
###########################################################
# If -SourceRepositoryFolder was given (e.g. a SCOM 2016 share with a wider
# set of exported MPs than just your -InputPath batch), any reference that
# isn't satisfied by the batch itself or by the target Repository is now
# looked up there by MP ID. A match is pulled into $BatchSource -- NOT
# copied in blindly: it joins the batch and goes through the exact same
# pipeline as every other batch MP (override validation, dependency graph,
# reference rewriting, candidate emission) in Steps 3-8 below.
#
# Newly-added dependency MPs can themselves have unresolved references, so
# this repeats until a pass adds nothing new (a fixed-point/transitive-
# closure loop), capped at a sane iteration limit so a bad/circular source
# folder can't spin forever.

# Declared unconditionally (not just inside the -SourceRepositoryFolder
# branch below) so the final batch summary in Step 9 can always reference
# it safely under Set-StrictMode, whether or not a source repo was given.
$totalPulled = 0
$OriginalBatchSize = $BatchSource.Count
$script:SourceXmlOnlyIndex = @{}
$script:UnresolvableFromSource = @{}

if (-not $SourceRepositoryFolder) {
    Write-Log ""
    Write-Log "No -SourceRepositoryFolder provided; skipping dependency auto-resolution. Unresolved references will be reported in MissingDependencies.csv as before." -Level WARN
}
else {

    Write-Banner "STEP 2.5 - AUTO-RESOLVING DEPENDENCIES FROM SOURCE REPOSITORY"
    Write-Log "Source repository folder: $SourceRepositoryFolder"

    # IMPORTANT (3.40): every MP reference in SCOM points at a SEALED MP and
    # carries its PublicKeyToken. An .xml exported from the source (which is
    # what Export-SCOMManagementPack produces, even for sealed MPs) can never
    # satisfy such a reference -- importing it creates an unsigned copy with
    # the same ID, the reference still fails, and the target now has a bogus
    # unsealed "Microsoft.*" MP in it. So only ORIGINAL sealed files (.mp /
    # .mpb) are used to auto-resolve dependencies. An .xml-only hit is
    # recorded so the report can say "exists in 2016, but you need the
    # original sealed file".
    Write-Log "Indexing source repository (this may take a moment for a large share)..."

    $sourceFiles = @()
    $sourceFiles += Get-ChildItem -LiteralPath $SourceRepositoryFolder -Filter *.xml -Recurse -ErrorAction SilentlyContinue
    $sourceFiles += Get-ChildItem -LiteralPath $SourceRepositoryFolder -Filter *.mp   -Recurse -ErrorAction SilentlyContinue
    $sourceFiles += Get-ChildItem -LiteralPath $SourceRepositoryFolder -Filter *.mpb  -Recurse -ErrorAction SilentlyContinue

    Write-Log "Source repository files found: $($sourceFiles.Count)"

    $script:SourceXmlOnlyIndex = @{}   # MPID -> version(s) present only as exported .xml
    $SealedSourceIndex = @{}           # MPID -> array of loaded sealed MP objects

    $sealedSourceFiles = @($sourceFiles | Where-Object { $_.Extension -ieq '.mp' -or $_.Extension -ieq '.mpb' })
    if ($sealedSourceFiles.Count -gt 0 -and -not $script:SdkLoaded) {
        $script:SdkLoaded = Initialize-ScomSdk
    }

    foreach ($sf in $sourceFiles) {
        if ($sf.Extension -ieq ".xml") {
            try {
                $peekXml = Read-MPXml -Path $sf.FullName
                $peekId = [string]$peekXml.ManagementPack.Manifest.Identity.ID
                $peekVer = [string]$peekXml.ManagementPack.Manifest.Identity.Version
                if ($peekId) {
                    if (-not $script:SourceXmlOnlyIndex.ContainsKey($peekId)) { $script:SourceXmlOnlyIndex[$peekId] = @() }
                    $script:SourceXmlOnlyIndex[$peekId] += $peekVer
                }
            }
            catch {
                Write-Log "Could not peek MP ID from '$($sf.FullName)': $($_.Exception.Message)" -Level WARN -NoConsole
            }
        }
        elseif ($script:SdkLoaded) {
            foreach ($lm in @(Import-MPFile -File $sf -WorkingFolder $BatchWorkingFolder)) {
                if (-not $SealedSourceIndex.ContainsKey($lm.ID)) { $SealedSourceIndex[$lm.ID] = @() }
                $SealedSourceIndex[$lm.ID] += $lm
            }
        }
    }

    Write-Log "Source repository indexed: $($SealedSourceIndex.Count) sealed MP ID(s) usable for dependency resolution; $($script:SourceXmlOnlyIndex.Count) MP ID(s) present only as exported .xml (analysis only)."
    if ($sealedSourceFiles.Count -gt 0 -and -not $script:SdkLoaded) {
        Write-Log "$($sealedSourceFiles.Count) sealed file(s) in the source folder could not be read because the SCOM SDK assemblies were not found. Run this on a management server or a machine with the Operations console installed." -Level WARN
    }

    function Find-InSourceRepository {
        param(
            [string]$MissingID,
            [string]$RequestedVersion
        )

        if (-not $SealedSourceIndex.ContainsKey($MissingID)) { return $null }
        $cands = @($SealedSourceIndex[$MissingID])

        # SCOM references are MINIMUM-version: prefer the exact version, else
        # the lowest version that is >= the requested one, else the highest
        # available (with a loud warning -- lower than requested will fail).
        $exact = $cands | Where-Object { [string]$_.Version -eq $RequestedVersion } | Select-Object -First 1
        if ($exact) { return $exact }

        $vReq = $null
        try { $vReq = [version]$RequestedVersion } catch { }
        if ($vReq) {
            $ge = @($cands | Where-Object { try { [version]$_.Version -ge $vReq } catch { $false } } |
                    Sort-Object { [version]$_.Version })
            if ($ge.Count -gt 0) { return $ge[0] }
        }

        if (-not $vReq) {
            return ($cands | Sort-Object { try { [version]$_.Version } catch { [version]"0.0.0.0" } } -Descending | Select-Object -First 1)
        }
        Write-Log "Sealed '$MissingID' is in the source folder, but only at v$(($cands | ForEach-Object { $_.Version }) -join ', ') -- lower than the v$RequestedVersion required, so SCOM would reject it. Not used." -Level WARN
        return $null
    }

    $maxPasses   = 10
    $passCount   = 0
    $totalPulled = 0
    $script:UnresolvableFromSource = @{}   # refId -> $true, once a lookup has failed this run (avoid re-searching)

    do {
        $passCount++
        $addedThisPass = 0

        # Snapshot current batch IDs so we scan each MP's references against
        # a stable list for this pass (mutating $BatchSource mid-foreach is
        # safe in PowerShell, but reasoning about it is not -- a snapshot
        # keeps this pass's logic simple and correct).
        $idsToScan = @($BatchSource.Keys)

        foreach ($scanId in $idsToScan) {

            $scanXml = $BatchSource[$scanId].XML
            $scanRefs = $null
            try { $scanRefs = $scanXml.ManagementPack.Manifest.References.Reference } catch {}

            foreach ($r in $scanRefs) {

                # Defensive: each reference is processed in its own try/catch
                # so that if something DOES go wrong, the error names the
                # exact MP and reference being processed at the time --
                # rather than surfacing as a bare top-level ArgumentException
                # with no indication of which of potentially hundreds of
                # references was responsible.
                try {
                    if (-not $r.ID) { continue }
                    $refId = [string]$r.ID

                    if ($ExcludedSet.ContainsKey($refId)) {
                        # Explicitly excluded via -ExcludeMPIDs -- known,
                        # intentional gap. Don't search for it, don't
                        # re-check it on later passes, and record it
                        # distinctly so the missing-dependencies report can
                        # tell "you told me to skip this" apart from
                        # "genuinely couldn't find this."
                        if (-not $script:UnresolvableFromSource.ContainsKey($refId)) {
                            Write-Log "  Skipping '$refId' (required by '$scanId') -- explicitly excluded via -ExcludeMPIDs." -Level WARN
                            $script:ExcludedFromSource[$refId] = $true
                        }
                        $script:UnresolvableFromSource[$refId] = $true
                        continue
                    }

                    if ($BatchSource.ContainsKey($refId)) { continue }   # already in batch
                    if ($Repository.ContainsKey($refId)) {
                        # Satisfied by the target only if the installed version
                        # is >= the referenced (minimum) version.
                        $repoOk = $true
                        try { $repoOk = ([version]$Repository[$refId].Version -ge [version][string]$r.Version) } catch { }
                        if ($repoOk) { continue }
                    }

                    # Already tried and failed this run? Don't re-search.
                    if ($script:UnresolvableFromSource.ContainsKey($refId)) { continue }

                    $resolved = Find-InSourceRepository -MissingID $refId -RequestedVersion ([string]$r.Version)

                    if ($resolved) {
                        $resolved | Add-Member -NotePropertyName NeedsSealedOriginal -NotePropertyValue $false -Force
                        $BatchSource[$resolved.ID] = $resolved
                        $addedThisPass++
                        $totalPulled++
                        Write-Log "  Auto-resolved '$refId' from source repository (v$($resolved.Version), sealed original: $([System.IO.Path]::GetFileName($resolved.SourcePath))) -- added to batch for full processing." -Level SUCCESS
                    }
                    else {
                        if ($script:SourceXmlOnlyIndex.ContainsKey($refId)) {
                            Write-Log "  '$refId' (needed by '$scanId') exists in the source only as an exported .xml -- that cannot satisfy a sealed reference. Supply the original sealed .mp/.mpb, or install a current version in the target." -Level WARN -NoConsole
                        }
                        $script:UnresolvableFromSource[$refId] = $true
                    }
                }
                catch {
                    $failLine = $_.InvocationInfo.ScriptLineNumber
                    $failText = $_.InvocationInfo.Line.Trim()
                    Write-Log "Dependency auto-resolution failed while processing reference '$($r.ID)' required by '$scanId' (pass $($passCount)): $($_.Exception.GetType().FullName): $($_.Exception.Message) [at script line ${failLine}: $failText]" -Level ERROR
                    if ($Strict) {
                        throw
                    }
                    else {
                        Write-Log "Continuing without -Strict: this reference will be treated as unresolved." -Level WARN
                        $script:UnresolvableFromSource[[string]$r.ID] = $true
                    }
                }
            }
        }

        Write-Log "Pass ${passCount}: $addedThisPass new dependency MP(s) pulled in from source repository."

    } while ($addedThisPass -gt 0 -and $passCount -lt $maxPasses)

    if ($passCount -ge $maxPasses -and $addedThisPass -gt 0) {
        Write-Log "Reached the $maxPasses-pass safety limit for dependency auto-resolution; some references may remain unresolved. This usually means a very deep dependency chain -- check MissingDependencies.csv after this run." -Level WARN
    }

    Write-Banner "STEP 2.5 COMPLETE"
    Write-Log "Dependency MPs auto-resolved from source repository : $totalPulled"
    Write-Log "Batch size after auto-resolution                    : $($BatchSource.Count)"
    if ($totalPulled -gt 0) {
        Write-Log "NOTE: auto-resolved MPs are now full batch members -- they go through override validation, the dependency graph, and reference rewriting like every other batch MP. Check the manifest for them by ID." -Level WARN
    }
}

###########################################################
# STEP 3 - BATCH DEPENDENCY GRAPH + TOPOLOGICAL IMPORT ORDER
###########################################################
# This is the core new capability for batch migrations: MPs in the source
# batch frequently reference EACH OTHER (e.g. a custom "Company.App.Monitoring"
# MP referencing a "Company.App.Library" MP that is also in the same export).
# Those internal edges determine import order; getting the order wrong means
# Import-SCOMManagementPack fails outright on a missing reference.
#
# References that point outside the batch are resolved against the target
# Repository instead (handled per-MP in Steps 4-8) and do not participate in
# the topological sort -- only batch-internal edges do, because only those
# represent an ordering constraint WITHIN this import run.

Write-Banner "STEP 3 - BATCH DEPENDENCY GRAPH"

$BatchGraph = @{}      # MPID -> @(MPID, MPID, ...) batch-internal dependencies
$ExternalRefs = @{}    # MPID -> @(ref ID, ref ID, ...) references outside the batch

foreach ($srcEntry in $BatchSource.Values) {

    $mpId = $srcEntry.ID
    $BatchGraph[$mpId] = New-Object System.Collections.Generic.List[string]
    $ExternalRefs[$mpId] = New-Object System.Collections.Generic.List[string]

    $refs = $null
    try {
        $refs = $srcEntry.XML.ManagementPack.Manifest.References.Reference
    }
    catch {
        Write-Log "Reference enumeration failed for batch MP '$mpId': $($_.Exception.Message)" -Level WARN -NoConsole
    }

    foreach ($r in $refs) {
        if (-not $r.ID) { continue }
        $refId = [string]$r.ID

        if ($BatchSource.ContainsKey($refId)) {
            if (-not $BatchGraph[$mpId].Contains($refId)) {
                $BatchGraph[$mpId].Add($refId)
            }
        }
        else {
            if (-not $ExternalRefs[$mpId].Contains($refId)) {
                $ExternalRefs[$mpId].Add($refId)
            }
        }
    }
}

$batchEdgeCount = 0
foreach ($k in $BatchGraph.Keys) { $batchEdgeCount += $BatchGraph[$k].Count }

Write-Log "Batch nodes          : $($BatchGraph.Count)"
Write-Log "Batch-internal edges : $batchEdgeCount"

Write-Log ""
Write-Log "Batch-Internal Dependencies" -Level WARN
Write-Log "----------------------------"
foreach ($mpId in $BatchGraph.Keys | Sort-Object) {
    if ($BatchGraph[$mpId].Count -gt 0) {
        foreach ($dep in $BatchGraph[$mpId]) {
            Write-Log "$mpId  ->  $dep"
        }
    }
}

###########################################################
# TOPOLOGICAL SORT (Kahn's algorithm)
###########################################################
# Kahn's algorithm is used (over recursive DFS) because it naturally exposes
# cycle detection as "nodes left over with nonzero in-degree" -- a clean,
# explicit failure mode rather than a stack-overflow risk on deep/circular
# chains, which matters here since MP reference chains in the wild can be
# unexpectedly deep.

function Get-TopologicalOrder {
    param(
        [Parameter(Mandatory)][hashtable]$Graph
    )

    # Build in-degree counts. Edge A -> B means "A depends on B", so B must
    # be imported BEFORE A. For import ordering we want dependencies first,
    # so we sort on the REVERSED graph: treat B as having an incoming edge
    # from A, and emit nodes with in-degree 0 (nothing depends on them... )
    # -- actually for import order we need the opposite: a node can be
    # imported only once everything IT depends on has already been
    # imported. So we compute out-degree-style readiness: a node is ready
    # when all of its dependencies have already been emitted.

    $remainingDeps = @{}   # MPID -> count of not-yet-satisfied dependencies
    $dependents    = @{}   # MPID -> list of MPIDs that depend on it (reverse edges)

    foreach ($node in $Graph.Keys) {
        $remainingDeps[$node] = $Graph[$node].Count
        if (-not $dependents.ContainsKey($node)) {
            $dependents[$node] = New-Object System.Collections.Generic.List[string]
        }
    }

    foreach ($node in $Graph.Keys) {
        foreach ($dep in $Graph[$node]) {
            if (-not $dependents.ContainsKey($dep)) {
                $dependents[$dep] = New-Object System.Collections.Generic.List[string]
            }
            $dependents[$dep].Add($node)
        }
    }

    # Ready queue = nodes with zero remaining (unsatisfied) dependencies.
    # Sorted alphabetically at each step for deterministic, reviewable output
    # across repeated runs (rather than hashtable enumeration order, which
    # is not guaranteed stable). @(...) forces an array even when exactly
    # zero or one key qualifies, since Where-Object/Sort-Object unwrap a
    # single match to a scalar otherwise and break the List[string] cast.
    $readyArray = @($remainingDeps.Keys | Where-Object { $remainingDeps[$_] -eq 0 } | Sort-Object)
    $ready = New-Object System.Collections.Generic.List[string]
    foreach ($n in $readyArray) { $ready.Add($n) }

    $order = New-Object System.Collections.Generic.List[string]
    $visited = @{}

    while ($ready.Count -gt 0) {

        # Always re-sort before taking the next item so newly-added nodes
        # (appended out of order below) never violate the alphabetical
        # tiebreak guarantee documented above. Simpler than maintaining a
        # sorted-insert and the list sizes here (dozens, not thousands of
        # MPs) make the extra sort negligible.
        $sortedArray = @($ready | Sort-Object)
        $ready.Clear()
        foreach ($n in $sortedArray) { $ready.Add($n) }

        $current = $ready[0]
        $ready.RemoveAt(0)

        if ($visited.ContainsKey($current)) { continue }
        $visited[$current] = $true
        $order.Add($current)

        foreach ($dependent in $dependents[$current]) {
            $remainingDeps[$dependent]--
            if ($remainingDeps[$dependent] -eq 0 -and -not $visited.ContainsKey($dependent)) {
                $ready.Add($dependent)
            }
        }
    }

    $cyclic = @($Graph.Keys | Where-Object { -not $visited.ContainsKey($_) })

    return [PSCustomObject]@{
        Order  = $order
        Cyclic = $cyclic
    }
}

$TopoResult   = Get-TopologicalOrder -Graph $BatchGraph
$ImportOrder  = $TopoResult.Order
$CyclicNodes  = $TopoResult.Cyclic

Write-Log ""
Write-Log "Topological Import Order Computed" -Level SUCCESS
Write-Log "-----------------------------------"
$position = 1
foreach ($mpId in $ImportOrder) {
    Write-Log ("{0,3}. {1}" -f $position, $mpId)
    $position++
}

if ($CyclicNodes.Count -gt 0) {
    Write-Log ""
    Write-Log "CIRCULAR DEPENDENCY DETECTED among $($CyclicNodes.Count) batch MP(s):" -Level ERROR
    $CyclicNodes | Sort-Object | ForEach-Object { Write-Log "  - $_" -Level ERROR }
    Write-Log "These MPs reference each other in a cycle and CANNOT be ordered for import as-is." -Level ERROR
    Write-Log "SCOM does not support circular MP references -- this must be fixed in the source MPs before import (break the cycle by removing/restructuring one of the references)." -Level ERROR

    if ($Strict) {
        throw "Strict validation failed: circular dependency detected among $($CyclicNodes.Count) batch MP(s). See log: $script:LogFile"
    }
    else {
        Write-Log "Continuing without -Strict: cyclic MPs are appended to the END of the import order (unordered, best-effort) and flagged in the manifest." -Level WARN
        foreach ($c in ($CyclicNodes | Sort-Object)) {
            $ImportOrder.Add($c)
        }
    }
}

Write-Banner "STEP 3 COMPLETE"
Write-Log "Batch MPs ordered    : $($ImportOrder.Count)"
Write-Log "Cyclic / unorderable : $($CyclicNodes.Count)"

###########################################################
# STEP 3.5 - OPTIONAL LIVE IMPORT OF DEPENDENCIES (-LiveImport only)
###########################################################
# Off by default. When -LiveImport is set, this connects to a real SCOM
# management group and ACTUALLY IMPORTS every auto-resolved DEPENDENCY
# (anything in $ImportOrder that is NOT one of the original -InputPath
# targets), at its genuine old version, exactly as exported from the
# source environment. The MP(s) you actually named in -InputPath are
# skipped here unconditionally -- they always go through the normal
# rewrite/candidate pipeline in Steps 4-8 below, live import or not.
#
# DEFAULT BEHAVIOR (changed from earlier versions): a decline or an actual
# import failure no longer stops the whole run. It's recorded, that one
# dependency (and anything that depends on it) is skipped, and the run
# continues with everything else -- so a single bad legacy dependency
# doesn't block you from seeing the full picture across a whole batch in
# one pass. Decisions persist across separate runs too (see
# LiveImportDecisions.json in -OutputFolder), so a dependency you've
# already declined or that already failed is never re-prompted for again.

if ($LiveImport) {

    Write-Banner "STEP 3.5 - LIVE DEPENDENCY IMPORT"

    if (-not $ManagementServer) {
        if ($NonInteractive) { $ManagementServer = "localhost" }
        else {
            $ManagementServer = Read-Host "Enter the SCOM management server to connect to (blank = localhost)"
            if ([string]::IsNullOrWhiteSpace($ManagementServer)) { $ManagementServer = "localhost" }
        }
    }

    Write-Log "Connecting to SCOM management group on '$ManagementServer'..." -Level WARN

    try {
        Import-Module OperationsManager -ErrorAction Stop
        # Capture the exact connection object this call creates, and pin to
        # IT (not to a server-name match, which fails for 'localhost').
        $connBefore = @(Get-SCOMManagementGroupConnection -ErrorAction SilentlyContinue)
        New-SCOMManagementGroupConnection -ComputerName $ManagementServer -ErrorAction Stop
        $connAfter = @(Get-SCOMManagementGroupConnection -ErrorAction Stop)
        $keyOf = { param($c) "$($c.ManagementGroupName)|$($c.ManagementServerName)" }
        $beforeKeys = @($connBefore | ForEach-Object { & $keyOf $_ })
        $newConns = @($connAfter | Where-Object { $beforeKeys -notcontains (& $keyOf $_) })
        $script:TargetConnection = if ($newConns.Count -eq 1) { $newConns[0] } else { @($connAfter | Where-Object { $_.IsActive }) | Select-Object -First 1 }
        if (-not $script:TargetConnection) { throw "Connected, but could not identify the new connection object." }
        $script:TargetServer = [string]$script:TargetConnection.ManagementServerName
        Use-TargetConnection
        $mgConn = $script:TargetConnection
        if ($SourceManagementServer -and $script:SourceMSName -and ([string]$mgConn.ManagementServerName -eq $script:SourceMSName)) {
            throw "The -ManagementServer connection resolved to '$($mgConn.ManagementServerName)', which is the SOURCE management server. Refusing to import."
        }
        Write-Log "Connected (active): Management Group '$($mgConn.ManagementGroupName)' via '$($mgConn.ManagementServerName)'" -Level SUCCESS
        if ($SourceManagementServer) {
            Write-Log "A SOURCE connection is also open in this session. All import/lookup calls are pinned to the TARGET connection above -- confirm that management group name is the SCOM $TargetVersion one before answering any prompt." -Level WARN
        }
        Update-InstalledMPCache
    }
    catch {
        throw "LiveImport requested but could not connect to SCOM on '$ManagementServer': $($_.Exception.Message). Nothing has been imported. Fix the connection and re-run, or omit -LiveImport to use the safe, file-only mode instead."
    }

    # ---- Persistent decision file (across separate runs) ----
    # Records every MP ID that was previously DECLINED or FAILED during a
    # live import, plus the version and reason, so a future run never
    # re-prompts for or re-attempts something already known to be a dead
    # end. Keyed by "ID|Version" since a different version of the same ID
    # might genuinely work even if an earlier one didn't.
    $decisionsFile = Join-Path $OutputFolder "LiveImportDecisions.json"
    $priorDecisions = @{}
    if (Test-Path -LiteralPath $decisionsFile) {
        try {
            $loaded = Get-Content -LiteralPath $decisionsFile -Raw | ConvertFrom-Json
            foreach ($prop in $loaded.PSObject.Properties) {
                $priorDecisions[$prop.Name] = $prop.Value
            }
            Write-Log "Loaded $($priorDecisions.Count) prior live-import decision(s) from $decisionsFile -- these will be skipped without re-prompting." -Level WARN
        }
        catch {
            Write-Log "Could not read $decisionsFile ($($_.Exception.Message)) -- starting with no prior decisions." -Level WARN
        }
    }

    function Save-LiveImportDecisions {
        try {
            $priorDecisions | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $decisionsFile -Encoding UTF8
        }
        catch {
            Write-Log "Could not save live-import decisions to $decisionsFile : $($_.Exception.Message)" -Level WARN
        }
    }

    $dependenciesToImport = @($ImportOrder | Where-Object { $OriginalInputTargets -notcontains $_ -and -not $ExcludedSet.ContainsKey($_) })

    $excludedFromLiveImport = @($ImportOrder | Where-Object { $OriginalInputTargets -notcontains $_ -and $ExcludedSet.ContainsKey($_) })
    if ($excludedFromLiveImport.Count -gt 0) {
        Write-Log "Excluded from live import (per -ExcludeMPIDs, will not be prompted for): $($excludedFromLiveImport -join ', ')" -Level WARN
    }

    Write-Log "Dependencies eligible for live import : $($dependenciesToImport.Count)"
    Write-Log "Migration target(s) (never auto-imported, always go through the normal candidate pipeline below): $($OriginalInputTargets -join ', ')"

    # Detect headless/non-interactive sessions explicitly (e.g. a remote
    # PowerShell session or scheduled task on the management server, which
    # -LiveImport is specifically likely to run under) rather than relying
    # on an exception from MessageBox.Show -- passing a null owner to it is
    # valid and does NOT reliably throw, so detecting up front is safer
    # than trying to catch our way out of it after the fact.
    $isInteractive = [System.Environment]::UserInteractive -and $script:UIAvailable -and -not $NonInteractive

    $liveImportOwner = $null
    if ($isInteractive) {
        try { $liveImportOwner = Get-TopMostOwner } catch { $isInteractive = $false }
    }
    if (-not $isInteractive) {
        Write-Log "Running headless (or no owner window available) -- confirmation prompts will use the console instead of a popup." -Level WARN
    }

    # Per-this-run results, keyed by MP ID, so cascading-skip detection and
    # the end-of-step summary both have a single source of truth.
    # Status is one of: Imported, AlreadyPresent, Declined, Failed, SkippedCascade
    $script:LiveImportResults = @{}

    foreach ($depId in $dependenciesToImport) {

        $depEntry = $BatchSource[$depId]

        if (-not $depEntry -or -not $depEntry.SourcePath) {
            Write-Log "Skipping live import of '$depId': no source file path on record (unexpected -- this MP should have come from -SourceRepositoryFolder)." -Level WARN
            $script:LiveImportResults[$depId] = [PSCustomObject]@{ Status = "Failed"; Reason = "No source file path on record" }
            continue
        }

        $decisionKey = "$depId|$($depEntry.Version)"

        # Cascading skip: if anything THIS dependency itself needs already
        # failed/was declined/excluded earlier in this same run, there is
        # no point prompting for or attempting this one -- it will fail
        # the same way SCOM-side. Record why and move on without asking.
        $myDeps = if ($BatchGraph.ContainsKey($depId)) { @($BatchGraph[$depId]) } else { @() }
        $failedPrereqs = @($myDeps | Where-Object {
            ($script:LiveImportResults.ContainsKey($_) -and $script:LiveImportResults[$_].Status -in @("Declined", "Failed", "SkippedCascade")) -or
            $ExcludedSet.ContainsKey($_)
        })

        if ($failedPrereqs.Count -gt 0) {
            Write-Log "Skipping '$depId': depends on $($failedPrereqs -join ', '), which $(if ($failedPrereqs.Count -eq 1) { 'was' } else { 'were' }) declined, failed, or excluded earlier in this run." -Level WARN
            $script:LiveImportResults[$depId] = [PSCustomObject]@{ Status = "SkippedCascade"; Reason = "Depends on: $($failedPrereqs -join ', ')" }
            continue
        }

        # Previously declined or failed in an EARLIER run (persisted)? Skip
        # without re-prompting -- that's the whole point of remembering.
        if ($priorDecisions.ContainsKey($decisionKey)) {
            $prior = $priorDecisions[$decisionKey]
            Write-Log "Skipping '$depId' (v$($depEntry.Version)): previously $($prior.status) on $($prior.date) -- $($prior.reason). Delete the relevant entry from $decisionsFile if you want to try again." -Level WARN
            $script:LiveImportResults[$depId] = [PSCustomObject]@{ Status = "SkippedCascade"; Reason = "Previously $($prior.status): $($prior.reason)" }
            continue
        }

        # Already installed at a matching version? Skip silently, no prompt
        # -- re-running after a partial success should not re-ask about
        # things already settled.
        $existing = Get-InstalledMP -Name $depId

        $existingOk = $false
        if ($existing) {
            try { $existingOk = ([version][string]$existing.Version -ge [version][string]$depEntry.Version) }
            catch { $existingOk = ([string]$existing.Version -eq [string]$depEntry.Version) }
        }
        if ($existingOk) {
            Write-Log "Already installed at v$($existing.Version) (>= v$($depEntry.Version)): $depId -- skipping." -Level SUCCESS
            $script:LiveImportResults[$depId] = [PSCustomObject]@{ Status = "AlreadyPresent"; Reason = "" }
            continue
        }

        if (-not $depEntry.WasSealed) {
            Write-Log "Not importing '$depId': only an exported .xml is available, which cannot stand in for a sealed MP." -Level WARN
            $script:LiveImportResults[$depId] = [PSCustomObject]@{ Status = "Failed"; Reason = "Only an exported .xml is available -- supply the original sealed .mp/.mpb" }
            continue
        }

        if ($existing) {
            Write-Log "NOTE: '$depId' is already installed, but at a DIFFERENT version (installed: v$($existing.Version), about to import: v$($depEntry.Version)). SCOM will reject this if it's not a valid upgrade path -- you'll be prompted either way." -Level WARN
        }

        if ($NonInteractive -or $AutoApproveDependencies) {
            $proceed = "Yes"
            Write-Log "Auto-approved (-NonInteractive/-AutoApproveDependencies): importing dependency '$depId'." -Level WARN -NoConsole
        }
        elseif ($isInteractive) {
            $proceed = [System.Windows.Forms.MessageBox]::Show(
                $liveImportOwner,
                "About to LIVE IMPORT a dependency into SCOM on '$ManagementServer':`n`nMP ID      : $depId`nVersion    : $($depEntry.Version)`nFile       : $($depEntry.SourcePath)`n`nThis is a real import into the connected management group, exactly as exported from the source environment (no rewrite applied -- dependencies are imported as-is).`n`nYes = import now.  No = skip this one (and anything that depends on it) and keep going with the rest of the batch.",
                "Confirm Live Import ($depId)",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
        }
        else {
            $consoleAnswer = Read-Host "About to LIVE IMPORT '$depId' (v$($depEntry.Version)) from '$($depEntry.SourcePath)' into '$ManagementServer'. Proceed? (Y/N -- N skips this one and continues with the rest)"
            $proceed = if ($consoleAnswer -match '^[Yy]') { "Yes" } else { "No" }
        }

        if ($proceed -ne "Yes") {
            Write-Log "Declined: '$depId' (v$($depEntry.Version)) -- skipping this dependency and anything that depends on it. Continuing with the rest of the batch." -Level WARN
            $script:LiveImportResults[$depId] = [PSCustomObject]@{ Status = "Declined"; Reason = "Declined by user" }
            $priorDecisions[$decisionKey] = @{ status = "declined"; reason = "Declined by user"; date = (Get-Date -Format "yyyy-MM-dd HH:mm:ss") }
            Save-LiveImportDecisions
            continue
        }

        Write-Log "Importing '$depId' (v$($depEntry.Version)) from '$($depEntry.SourcePath)'..." -Level WARN

        try {
            Use-TargetConnection
            Import-SCOMManagementPack -Fullname $depEntry.SourcePath -ErrorAction Stop
            Write-Log "Successfully imported: $depId (v$($depEntry.Version))" -Level SUCCESS
            $script:LiveImportResults[$depId] = [PSCustomObject]@{ Status = "Imported"; Reason = "" }
            $script:InstalledMPCache[$depId] = [PSCustomObject]@{ Name = $depId; Version = $depEntry.Version; KeyToken = 'sealed'; Sealed = $true }
        }
        catch {
            # Walk the full exception chain and surface every level's
            # message -- SCOM's own errors are routinely wrapped 2-3 levels
            # deep (outer "not valid, see inner exception" wrapper -> a
            # FailedVerification-type exception -> the actual specific
            # reason, e.g. "MP not found in the store"), and the outer
            # message alone is rarely useful on its own.
            $exceptionChain = New-Object System.Collections.Generic.List[string]
            $currentEx = $_.Exception
            $depth = 0
            while ($currentEx -and $depth -lt 6) {
                $msgText = if ($null -ne $currentEx.Message) { $currentEx.Message.Trim() } else { "(no message)" }
                $exceptionChain.Add("  [$depth] $($currentEx.GetType().Name): $msgText")
                $currentEx = $currentEx.InnerException
                $depth++
            }

            Write-Log "Live import FAILED for '$depId'. Full exception chain:" -Level ERROR
            foreach ($line in $exceptionChain) { Write-Log $line -Level ERROR }
            Write-Log "Continuing with the rest of the batch -- '$depId' and anything depending on it will be skipped." -Level WARN

            $innermostMsg = $exceptionChain[$exceptionChain.Count - 1]
            $script:LiveImportResults[$depId] = [PSCustomObject]@{ Status = "Failed"; Reason = $innermostMsg }
            $priorDecisions[$decisionKey] = @{ status = "failed"; reason = $innermostMsg; date = (Get-Date -Format "yyyy-MM-dd HH:mm:ss") }
            Save-LiveImportDecisions
        }
    }

    if ($liveImportOwner) { $liveImportOwner.Close(); $liveImportOwner.Dispose() }

    $importedCount       = @($script:LiveImportResults.Values | Where-Object { $_.Status -eq "Imported" }).Count
    $alreadyPresentCount = @($script:LiveImportResults.Values | Where-Object { $_.Status -eq "AlreadyPresent" }).Count
    $declinedCount       = @($script:LiveImportResults.Values | Where-Object { $_.Status -eq "Declined" }).Count
    $failedCount         = @($script:LiveImportResults.Values | Where-Object { $_.Status -eq "Failed" }).Count
    $cascadeCount        = @($script:LiveImportResults.Values | Where-Object { $_.Status -eq "SkippedCascade" }).Count

    Write-Banner "STEP 3.5 COMPLETE"
    Write-Log "Imported just now      : $importedCount"
    Write-Log "Already present        : $alreadyPresentCount"
    Write-Log "Declined               : $declinedCount"
    Write-Log "Failed                 : $failedCount"
    Write-Log "Skipped (cascade/prior): $cascadeCount"

    if ($failedCount -gt 0 -or $declinedCount -gt 0 -or $cascadeCount -gt 0) {
        Write-Log ""
        Write-Log "NOT fully resolved this run (review before relying on $($OriginalInputTargets -join ', ')):" -Level WARN
        foreach ($depId in $script:LiveImportResults.Keys | Sort-Object) {
            $r = $script:LiveImportResults[$depId]
            if ($r.Status -in @("Declined", "Failed", "SkippedCascade")) {
                Write-Log "  [$($r.Status)] $depId -- $($r.Reason)" -Level WARN
            }
        }
        Write-Log ""
        Write-Log "Decisions saved to $decisionsFile -- these will not be re-prompted on future runs. Delete entries there to retry them." -Level WARN
    }

    Write-Log ""
    Write-Log "Proceeding to candidate generation for your actual migration target(s). Note: if any dependency above did not fully resolve, the target's compatibility score/report will reflect that honestly -- it does not assume success." -Level SUCCESS
}

###########################################################
# SHARED HELPER FUNCTIONS (used inside the per-MP loop below)
###########################################################

function Resolve-MPReference {
    param(
        [string]$Reference,
        [hashtable]$AliasMapForMP
    )

    if (-not $Reference) { return $null }

    if ($Reference -like "*!*") {
        $parts = $Reference.Split("!")

        $alias = $parts[0]
        $id    = $parts[1]

        if ($AliasMapForMP.ContainsKey($alias)) {
            return $AliasMapForMP[$alias].ID
        }

        return $id
    }

    return $Reference
}

function Get-MPFamily {

    param([string]$MPID)

    switch -Regex ($MPID) {
        # CommonLibrary must be checked BEFORE the general IIS pattern below
        # -- it is NOT safe to treat the same as the version-specific IIS
        # MPs. Old IIS MPs (2003/2008/2012) require the EXACT OLD version of
        # CommonLibrary; the modern unified CommonLibrary (10.x) uses a
        # different schema/reference chain that those old MPs were never
        # updated to use. Auto-rewriting this reference to "the newest IIS-
        # family MP in the repo" silently breaks old IIS MPs by pointing
        # them at an incompatible CommonLibrary version they can't actually
        # use -- confirmed in practice: SCOM rejects the resulting MP.
        '^Microsoft\.Windows\.InternetInformationServices\.CommonLibrary$' { return 'IISCommonLibrary' }
        '^Microsoft\.Windows\.InternetInformationServices'                { return 'IIS' }
        '^Microsoft\.SQLServer'                            { return 'SQL' }
        '^Microsoft\.SystemCenter'                          { return 'SystemCenter' }
        '^Microsoft\.Windows\.Server\.AD'                   { return 'ActiveDirectory' }
        '^Microsoft\.Windows\.Server'                       { return 'Windows' }
        '^Microsoft\.Windows\.Client'                       { return 'WindowsClient' }
        '^Microsoft\.ConfigurationManager'                  { return 'SCCM' }
        '^Microsoft\.Exchange'                              { return 'Exchange' }
        '^Microsoft\.AD'                                    { return 'ActiveDirectory' }
        '^Microsoft\.SharePoint'                            { return 'SharePoint' }
        '^Microsoft\.Unix'                                  { return 'Unix' }
        '^Microsoft\.Linux'                                 { return 'Linux' }
        # Foundational MPs that ship with every SCOM install (confirmed
        # against Microsoft's own "Management Packs Installed with
        # Operations Manager" documentation) -- these are never something
        # to download separately; if missing from the target repo scan,
        # it almost always means the wrong folder was scanned, not that
        # SCOM itself is missing them.
        '^System\.Library$'                                 { return 'CoreLibrary' }
        '^System\.Health\.Library$'                         { return 'CoreLibrary' }
        '^System\.Performance\.Library$'                    { return 'CoreLibrary' }
        '^System\.Snmp\.Library$'                            { return 'CoreLibrary' }
        '^Microsoft\.Windows\.Library$'                      { return 'CoreLibrary' }
        '^Microsoft\.Windows\.Cluster\.Library$'             { return 'CoreLibrary' }
        '^Microsoft\.SystemCenter\.InstanceGroup\.Library$'  { return 'CoreLibrary' }
        default                                              { return 'Unknown' }
    }
}

# --------------------------------------------------------
# MP FAMILY KNOWLEDGE TABLE
# --------------------------------------------------------
# Encodes what's actually true, as of SCOM 2025, about how each Microsoft MP
# family evolved -- verified against Microsoft's own MP download pages and
# the SCOM management-pack changes-history docs (not guessed):
#
#   - IIS, SQL Server, Windows Server (Base OS), and Active Directory (ADDS)
#     all UNIFIED: Microsoft replaced per-OS-version MPs (e.g. one MP per
#     Windows Server release) with a single current MP that covers a wide
#     OS/product range. For these families, "the newest version already in
#     your target repo" is a safe recommendation.
#
#   - Exchange unified similarly: one "Exchange Server 2013 and above" MP
#     replaced the old per-version (2007/2010/2013) MPs.
#
#   - SharePoint did NOT unify -- Microsoft still ships separate MPs per
#     major SharePoint version (2013, 2016, 2019, Subscription Edition).
#     Swapping in "whatever SharePoint MP is newest in the repo" is WRONG
#     here if your source MP monitors a different SharePoint version --
#     this needs a human to confirm which on-prem SharePoint version is
#     actually running before picking a target MP.
#
#   - SystemCenter (library/core) MPs are foundational and version-locked
#     by design -- never auto-swapped (existing behavior, unchanged).
#
#   - A small number of MP families have an announced END-OF-SUPPORT date
#     even though they still technically import and run. These are flagged
#     distinctly so they don't get silently treated as "safe to upgrade."
#
# This table is reference material for humans AND the advisor logic below;
# extend FamilyBehavior / DeprecationNotices as you verify more families.

$FamilyBehavior = @{
    "IIS"              = "Unified"        # one current MP covers all supported OS/IIS versions
    "SQL"              = "Unified"        # version-agnostic MP, currently covers SQL 2014-2025+
    "Windows"          = "Unified"        # Base OS MP covers Server 2016-2025 in one package
    "ActiveDirectory"  = "Unified"        # current "ADDS" MP covers 2016/2019/2022 DCs in one package
    "Exchange"         = "Unified"        # "Exchange Server 2013 and above" covers 2013-2019/2022
    "SharePoint"       = "PerVersion"     # NOT unified -- 2013/2016/2019/SE remain separate MPs
    "SystemCenter"     = "ExactMatchOnly" # foundational/library MPs -- never auto-swapped
    "CoreLibrary"      = "ExactMatchOnly" # ships with every SCOM install -- never auto-swapped, never "missing" from a real SCOM 2025 environment, only from an incomplete repo scan
    "IISCommonLibrary" = "ExactMatchOnly" # old IIS MPs (2003/2008/2012) need the EXACT original version -- the modern unified version uses an incompatible schema/reference chain for them
    "SCCM"             = "PerVersion"     # ConfigMgr client MPs track ConfigMgr release lines
    "WindowsClient"    = "PerVersion"     # Windows 10/11 client MPs track OS release lines
    "Unix"             = "Unified"
    "Linux"            = "Unified"
}

# Known end-of-support notices for MP families that still IMPORT and RUN on
# SCOM 2025 today but have an announced sunset -- flagged, not blocked.
# Source: Microsoft Tech Community deprecation announcements.
$DeprecationNotices = @{
    "Microsoft.SQLServer.Reporting"        = "Microsoft has announced SCOM support for SSRS/PBIRS management packs ends January 2027. They continue to function on SCOM 2019/2022 but compatibility with SQL Server 2025 / SCOM 2025 / future releases is not guaranteed. Plan to migrate this monitoring to Azure-based alternatives."
    "Microsoft.SQLServer.Analysis"         = "Microsoft has announced SCOM support for SSAS management packs ends January 2027. They continue to function on SCOM 2019/2022 but compatibility with SQL Server 2025 / SCOM 2025 / future releases is not guaranteed. Plan to migrate this monitoring to Azure-based alternatives."
}

# Where to point a human to find a specific missing MP. systemcenter.wiki is
# a long-running, community-maintained mirror of nearly every Microsoft MP
# ever released, with direct per-version download links keyed off MP ID --
# more useful here than a generic search because it's MP-ID-addressable.
function Get-MPLookupHint {
    param([string]$MPID, [string]$Family)

    $wikiUrl = "https://systemcenter.wiki/?Get-ManagementPack=$MPID"

    $sourceNote = switch ($Family) {
        "IIS"             { "Microsoft Download Center (search 'Management Pack for Internet Information Services') or the link below." }
        "SQL"             { "Microsoft Download Center (search 'Management Pack for SQL Server') or the link below." }
        "Windows"         { "Microsoft Download Center (search 'Management Pack for Windows Server Operating System') or the link below." }
        "ActiveDirectory" { "Microsoft Download Center (search 'Management Pack for ADDS' / 'Active Directory') or the link below." }
        "Exchange"        { "Microsoft Download Center (search 'Exchange Server Management Pack') or the link below." }
        "SharePoint"      { "Microsoft Download Center (search 'Management Pack for SharePoint Server', matching your exact SharePoint version) or the link below." }
        "SystemCenter"    { "This is a foundational/library MP -- it should already be present on any SCOM install; verify your SCOM $TargetVersion install/setup media or the management server's MP folder first." }
        "CoreLibrary"     { "This ships with EVERY SCOM installation (confirmed: it's one of the core MPs installed automatically with Operations Manager). It is almost certainly already present on your SCOM 2025 management server -- this is very likely a sign that -RepositoryFolder did not scan a complete repository. Run 'Get-SCOMManagementPack | Export-SCOMManagementPack -Path <folder>' from a live SCOM 2025 connection to get a guaranteed-complete export, then point -RepositoryFolder at that instead." }
        "IISCommonLibrary" { "This needs the EXACT SAME VERSION already used by the old IIS MP it's paired with -- not the newest available, and not the modern unified CommonLibrary (10.x), which uses an incompatible schema for old IIS MPs. Find the matching old version in your SCOM $SourceVersion source export." }
        default           { "Microsoft Download Center, the MP vendor's site (for third-party/custom MPs), or the link below." }
    }

    return [PSCustomObject]@{
        SuggestedSource = $sourceNote
        LookupUrl        = $wikiUrl
    }
}

# --------------------------------------------------------
# FAMILY TABLE (built once from the target Repository, shared across all
# per-MP advisory passes below)
# --------------------------------------------------------

$FamilyTable = @{}

foreach ($repoMP in $Repository.Keys) {

    $family = Get-MPFamily $repoMP

    if (-not $FamilyTable.ContainsKey($family)) {
        $FamilyTable[$family] = @()
    }

    $FamilyTable[$family] += $repoMP
}

function Get-BestCandidate {

    param(
        [string]$Family,
        [string]$ReferenceID
    )

    if (-not $FamilyTable.ContainsKey($Family)) {
        return $null
    }

    $candidates = $FamilyTable[$Family]
    $behavior   = if ($FamilyBehavior.ContainsKey($Family)) { $FamilyBehavior[$Family] } else { "Unified" }

    # ----------------------------------------
    # EXACT-MATCH-ONLY FAMILIES (SystemCenter library/core)
    # ----------------------------------------
    # Foundational MPs where swapping in "any newer one alphabetically" is
    # unsafe -- they must match exactly, or be left alone for manual review.
    if ($behavior -eq "ExactMatchOnly") {
        if ($candidates -contains $ReferenceID) {
            return $ReferenceID
        }
        return $null
    }

    # ----------------------------------------
    # PER-VERSION FAMILIES (SharePoint, SCCM, WindowsClient)
    # ----------------------------------------
    # These did NOT unify across versions in real life -- picking "the
    # highest version present" could silently swap SharePoint 2013 for
    # SharePoint Subscription Edition, which monitors a DIFFERENT product
    # version and will not work correctly. Only auto-match if the OLD
    # reference ID and a repo candidate share their version-identifying
    # token (a crude but safe heuristic); otherwise defer to manual review.
    if ($behavior -eq "PerVersion") {
        $exactToken = $candidates | Where-Object {
            # crude shared-substring check: same family, look for a shared
            # 4-digit year-like token between old ID and candidate ID
            $oldTokens = [regex]::Matches($ReferenceID, '\d{4}') | ForEach-Object { $_.Value }
            $newTokens = [regex]::Matches($_, '\d{4}') | ForEach-Object { $_.Value }
            $oldTokens -and $newTokens -and (Compare-Object $oldTokens $newTokens -IncludeEqual -ExcludeDifferent | Measure-Object).Count -gt 0
        } | Select-Object -First 1

        if ($exactToken) { return $exactToken }
        return $null   # ambiguous -- needs a human, not a guess
    }

    # ----------------------------------------
    # UNIFIED FAMILIES (IIS / SQL / Windows / ActiveDirectory / Exchange / etc.)
    # ----------------------------------------
    # Prefer the highest-versioned candidate in the repository -- the real,
    # verified pattern for these families (one current MP supersedes the
    # old per-version line).
    return $candidates |
        ForEach-Object {
            [PSCustomObject]@{ ID = $_; Version = $Repository[$_].Version }
        } |
        Sort-Object { try { [version]$_.Version } catch { [version]"0.0.0.0" } } -Descending |
        Select-Object -First 1 -ExpandProperty ID
}

# --------------------------------------------------------
# REFERENCE PATTERN MAP (old/versioned ID -> family-prefix to match)
# --------------------------------------------------------
# Maps a known OLD reference ID directly to the family prefix it should be
# matched against in the target Repository. Extend this table as you
# encounter additional renamed/restructured MPs during migration testing.
# This is a fast-path lookup for well-known renames; Get-BestCandidate
# (above) handles the general case via family + behavior.
#
# Primary tested path for this tool: SCOM 2016 -> SCOM 2025.
# Kept generic (not hardcoded to one source version) so the same script
# can be reused for other source versions later -- see -SourceVersion.

$ReferencePatterns = @{

    # --- IIS family (unified) ---
    "Microsoft.Windows.InternetInformationServices.2003" = "Microsoft.Windows.InternetInformationServices"
    "Microsoft.Windows.InternetInformationServices.2008" = "Microsoft.Windows.InternetInformationServices"
    "Microsoft.Windows.InternetInformationServices.2012" = "Microsoft.Windows.InternetInformationServices"
    "Microsoft.Windows.InternetInformationServices.2016" = "Microsoft.Windows.InternetInformationServices"

    # --- System Center family (exact-match only) ---
    "Microsoft.SystemCenter.2007" = "Microsoft.SystemCenter"

    # --- Windows Server family (unified; common 2016-era MP names) ---
    "Microsoft.Windows.Server.2012"            = "Microsoft.Windows.Server"
    "Microsoft.Windows.Server.2012.Discovery"  = "Microsoft.Windows.Server.Discovery"
    "Microsoft.Windows.Server.2012.Monitoring" = "Microsoft.Windows.Server.Monitoring"
    "Microsoft.Windows.Server.2016"            = "Microsoft.Windows.Server"
    "Microsoft.Windows.Server.2016.Discovery"  = "Microsoft.Windows.Server.Discovery"
    "Microsoft.Windows.Server.2016.Monitoring" = "Microsoft.Windows.Server.Monitoring"

    # --- Active Directory family (unified; old per-OS AD MPs -> ADDS line) ---
    "Microsoft.Windows.Server.AD.2008"            = "Microsoft.Windows.Server.ADDS"
    "Microsoft.Windows.Server.AD.2008.Monitoring" = "Microsoft.Windows.Server.ADDS"
    "Microsoft.Windows.Server.AD.2012"            = "Microsoft.Windows.Server.ADDS"
    "Microsoft.Windows.Server.AD.2012.Monitoring" = "Microsoft.Windows.Server.ADDS"
    "Microsoft.Windows.Server.AD.2016"            = "Microsoft.Windows.Server.ADDS"
    "Microsoft.Windows.Server.AD.2016.Monitoring" = "Microsoft.Windows.Server.ADDS"

    # --- SQL Server family (unified) ---
    "Microsoft.SQLServer.2012.Discovery" = "Microsoft.SQLServer.Discovery"
    "Microsoft.SQLServer.2014.Discovery" = "Microsoft.SQLServer.Discovery"
    "Microsoft.SQLServer.2016.Discovery" = "Microsoft.SQLServer.Discovery"

    # --- Exchange family (unified; old per-version MPs -> "2013 and above") ---
    "Microsoft.Exchange.2007" = "Microsoft.Exchange.15"
    "Microsoft.Exchange.2010" = "Microsoft.Exchange.15"
}

function Find-BestRepositoryMatch {

    param(
        [string]$OldID
    )

    if ($ReferencePatterns.ContainsKey($OldID)) {
        $pattern = $ReferencePatterns[$OldID]
    }
    else {
        return $null
    }

    $repoMatches = $Repository.Keys |
        Where-Object { $_ -like "$pattern*" }

    if (-not $repoMatches -or @($repoMatches).Count -eq 0) {
        return $null
    }

    $best = $repoMatches |
        Sort-Object {
            try { [version]$Repository[$_].Version } catch { [version]"0.0.0.0" }
        } -Descending |
        Select-Object -First 1

    return $Repository[$best]
}

###########################################################
# STEPS 4-8 - PER-MP PROCESSING (run in topological order)
###########################################################
# Each MP in the batch goes through: alias map -> override validation ->
# dependency listing (vs. target repo) -> rewrite table -> candidate XML
# emission -> migration advisor / semantic recommendations. Results are
# collected per-MP into $BatchResults and rolled into one combined manifest
# in Step 9.

Write-Banner "STEPS 4-8 - PER-MP PROCESSING ($($ImportOrder.Count) MPs)"

# Batch-wide collector for every reference that resolved to neither the
# batch itself nor the target repository. Populated inside the Migration
# Advisor section of the per-MP loop below; rolled into its own CSV report
# in Step 9 so "what do I need to go find" is one list, not buried per-MP.
$script:MissingDependencies = @()

# Batch MPs go into the element index too: when a batch MP is imported, ITS
# version of the elements is what the target will have, so it takes
# precedence over an older/newer copy already in the repository export.
foreach ($bid in @($BatchSource.Keys)) {
    if ($AuditOnly -and $ElementIndex.ContainsKey($bid)) { continue }
    if ($Repository.ContainsKey($bid)) {
        # Target already has this MP at a HIGHER version: the batch copy will be
        # skipped at import time, so the target's element list is the truth.
        $repoNewer = $false
        try { $repoNewer = ([version][string]$Repository[$bid].Version -gt [version][string]$BatchSource[$bid].Version) } catch { }
        if ($repoNewer) { continue }
    }
    try { $ElementIndex[$bid] = Get-MPElementIdSet -Doc $BatchSource[$bid].XML }
    catch { Write-Log "Element indexing failed for batch MP $($bid): $($_.Exception.Message)" -Level WARN -NoConsole }
}

# Sealed targets are imported from their original file; the extracted XML is
# written here for reference only, so nobody imports it by mistake.
$AnalysisFolder = Join-Path $CandidateFolder "_analysis_only_do_not_import"
if (-not (Test-Path -LiteralPath $AnalysisFolder)) { New-Item -ItemType Directory -Path $AnalysisFolder -Force | Out-Null }

# Checks one "Alias!ElementID" (or bare local "ElementID") reference.
# Returns: OK | NO_ALIAS (alias not declared) | NO_MP (MP not known in target
# repo or batch) | NO_ELEMENT (MP known, element not in it) | EMPTY
function Test-ElementRef {
    param(
        [string]$Ref,
        [hashtable]$AliasToId,
        [string]$OwnId
    )
    if ([string]::IsNullOrWhiteSpace($Ref)) { return [PSCustomObject]@{ Status = 'EMPTY'; MPID = $null; Element = $null } }
    $mpId = $OwnId
    $el = $Ref
    if ($Ref -like '*!*') {
        $bang = $Ref.IndexOf('!')
        $refAlias = $Ref.Substring(0, $bang)
        $el = $Ref.Substring($bang + 1)
        if (-not $AliasToId.ContainsKey($refAlias)) { return [PSCustomObject]@{ Status = 'NO_ALIAS'; MPID = $null; Element = $el } }
        $mpId = [string]$AliasToId[$refAlias]
    }
    if (-not $ElementIndex.ContainsKey($mpId)) { return [PSCustomObject]@{ Status = 'NO_MP'; MPID = $mpId; Element = $el } }
    if ($ElementIndex[$mpId].Contains($el)) { return [PSCustomObject]@{ Status = 'OK'; MPID = $mpId; Element = $el } }
    return [PSCustomObject]@{ Status = 'NO_ELEMENT'; MPID = $mpId; Element = $el }
}

# Attributes that carry element references on override / category / folder
# item elements.
$script:OverrideRefAttributes = @('Monitor', 'Rule', 'Discovery', 'Diagnostic', 'Recovery', 'Context', 'SecureReference')

# Any batch MP that another MP REFERENCES must be sealed (SCOM does not allow
# references to unsealed MPs). If we only hold an exported .xml of it, the
# original was sealed -- even without -SourceInventory we can tell.
foreach ($bid in @($BatchSource.Keys)) {
    $bRefs = @()
    try { $bRefs = @($BatchSource[$bid].XML.SelectNodes('/ManagementPack/Manifest/References/Reference/ID') | ForEach-Object { [string]$_.InnerText }) } catch { }
    foreach ($rid in $bRefs) {
        if ($BatchSource.ContainsKey($rid) -and -not $BatchSource[$rid].WasSealed -and -not $BatchSource[$rid].NeedsSealedOriginal) {
            $BatchSource[$rid].NeedsSealedOriginal = $true
            Write-Log "'$rid' is referenced by '$bid', so it was a SEALED MP in the source -- the exported .xml supplied cannot be imported in its place. Marked BLOCKED until the original .mp/.mpb is supplied." -Level WARN
        }
    }
}

$script:AllStripLog = New-Object System.Collections.Generic.List[object]

$script:InstanceMap = $null
if ($InstanceMapFile) {
    if (-not (Test-Path -LiteralPath $InstanceMapFile)) { throw "Instance map not found: $InstanceMapFile" }
    $script:InstanceMap = @{}
    foreach ($im in @(Import-Csv -LiteralPath $InstanceMapFile)) {
        if ($im.NewGuid) { $script:InstanceMap[([string]$im.OldGuid).Trim().ToLowerInvariant()] = [string]$im.NewGuid }
    }
    Write-Log "Instance map loaded: $($script:InstanceMap.Count) source object(s) mapped to their target equivalent."
}

###########################################################
# STATIC -> DYNAMIC GROUP CONVERSION SETUP (-GroupConversionFile)
###########################################################
$script:GroupConvRows = @{}   # "MP|DiscoveryID|RuleIndex" -> row
if ($GroupConversionFile) {
    if (-not (Test-Path -LiteralPath $GroupConversionFile)) { throw "Group conversion file not found: $GroupConversionFile" }
    foreach ($gr in @(Import-Csv -LiteralPath $GroupConversionFile)) {
        $script:GroupConvRows["$($gr.MP)|$($gr.DiscoveryID)|$($gr.RuleIndex)"] = $gr
    }
    Write-Log "Group conversion file loaded: $($script:GroupConvRows.Count) static membership rule(s), $(@($script:GroupConvRows.Values | Where-Object { $_.Convert -match '^(y|yes|true|1)$' }).Count) marked Convert = Y."
}
# Singleton classes (groups) anywhere in the target or the batch: a static
# member of one of these is a nested subgroup, whose dynamic equivalent is
# simply "all instances of that class".
$script:SingletonClasses = New-Object 'System.Collections.Generic.HashSet[string]'
foreach ($docSrc in @(@($Repository.Values | ForEach-Object { $_.XML }) + @($BatchSource.Values | ForEach-Object { $_.XML }))) {
    foreach ($ct in @($docSrc.SelectNodes("//ClassType[@Singleton='true']"))) { [void]$script:SingletonClasses.Add([string]$ct.GetAttribute('ID')) }
}
$script:GroupConversionResults = New-Object System.Collections.Generic.List[object]

function Convert-StaticGroupRules {
    param([System.Xml.XmlDocument]$Doc, [string]$MPID, [string]$TargetLabel)

    $refsNode = $Doc.SelectSingleNode('/ManagementPack/Manifest/References')
    $winAlias = $null
    if ($refsNode) {
        foreach ($rn in @($refsNode.SelectNodes('Reference'))) {
            if ([string]$rn.SelectSingleNode('ID').InnerText -eq 'Microsoft.Windows.Library') { $winAlias = [string]$rn.GetAttribute('Alias') }
        }
    }
    $propPath = $null
    $ensureWin = {
        if (-not $script:cvWinAlias) {
            $script:cvWinAlias = 'MigWin'
            $ver = if ($Repository.ContainsKey('Microsoft.Windows.Library')) { [string]$Repository['Microsoft.Windows.Library'].Version } else { '7.5.8501.0' }
            if (-not $refsNode) {
                $refsNode = $Doc.CreateElement('References')
                [void]$Doc.SelectSingleNode('/ManagementPack/Manifest').AppendChild($refsNode)
            }
            $frag = $Doc.CreateDocumentFragment()
            $frag.InnerXml = "<Reference Alias=`"MigWin`"><ID>Microsoft.Windows.Library</ID><Version>$ver</Version><PublicKeyToken>31bf3856ad364e35</PublicKeyToken></Reference>"
            [void]$refsNode.AppendChild($frag)
        }
    }
    $script:cvWinAlias = $winAlias

    foreach ($disc in @($Doc.SelectNodes('/ManagementPack/Monitoring/Discoveries/Discovery'))) {
        $ds = $disc.SelectSingleNode('DataSource')
        if (-not $ds -or [string]$ds.GetAttribute('TypeID') -notlike '*GroupPopulator*') { continue }
        $discId = [string]$disc.GetAttribute('ID')
        $gi = $ds.SelectSingleNode('GroupInstanceId')
        $groupClass = if ($gi) { ([string]$gi.InnerText) -replace '^\$MPElement\[Name="?', '' -replace '"?\]\$$', '' } else { '' }
        $ri = 0
        foreach ($mr in @($ds.SelectNodes('MembershipRules/MembershipRule'))) {
            $ri++
            $incNode = $mr.SelectSingleNode('IncludeList')
            $excNode = $mr.SelectSingleNode('ExcludeList')
            if (-not $incNode -and -not $excNode) { continue }
            $mcText = [string]$mr.SelectSingleNode('MonitoringClass').InnerText
            $mcRef = $mcText -replace '^\$MPElement\[Name="?', '' -replace '"?\]\$$', ''
            $mcId = if ($mcRef -like '*!*') { $mcRef.Substring($mcRef.IndexOf('!') + 1) } else { $mcRef }
            $res = [PSCustomObject]@{ MP = $MPID; GroupClass = $groupClass; DiscoveryID = $discId; RuleIndex = $ri; Action = ''; Operator = ''; Pattern = ''; ExcludePattern = ''; Note = '' }

            if ($script:SingletonClasses.Contains($mcId)) {
                if ($incNode) { [void]$mr.RemoveChild($incNode) }
                if ($excNode) { [void]$mr.RemoveChild($excNode) }
                $res.Action = 'NestedGroup'; $res.Note = "Static subgroup reference replaced by 'all instances of $mcId' (a singleton group)"
                $script:GroupConversionResults.Add($res); continue
            }

            $row = $script:GroupConvRows["$MPID|$discId|$ri"]
            if (-not $row -or [string]$row.Convert -notmatch '^(y|yes|true|1)$' -or -not $row.Pattern) {
                $res.Action = 'LeftStatic'; $res.Note = if ($row) { "Convert = N or no pattern ($($row.Recommendation))" } else { 'No row in GroupConversion.csv -- will import EMPTY' }
                $script:GroupConversionResults.Add($res); continue
            }

            & $ensureWin
            $prop = "`$MPElement[Name=`"$($script:cvWinAlias)!Microsoft.Windows.Computer`"]/NetbiosComputerName`$"
            $op = if ($row.Operator) { [string]$row.Operator } else { 'MatchesRegularExpression' }
            $pats = @(if ($op -eq 'MatchesWildcard') { ([string]$row.Pattern) -split ';' | Where-Object { $_ } } else { [string]$row.Pattern })
            # Computers: <Property>. Hosted classes (OS, disks, services...):
            # <HostProperty> -- the host computer's NetBIOS name. Both are
            # exactly what the console's group wizard writes, so the group
            # stays editable with Create/Edit rules (a <Contained> expression
            # is valid but greys the editor out).
            $isComputer = ([string]$row.MemberKind -eq 'Computer') -or ($mcId -match '^Microsoft\.Windows\.(Server\.|Client\.)?Computer$')
                        # HostProperty must name the host class AND the property:
            #   <HostProperty><MonitoringClass>..Computer..</MonitoringClass><Property>..NetbiosComputerName..</Property></HostProperty>
            $propEsc = [System.Security.SecurityElement]::Escape($prop)
            $hostCls = [System.Security.SecurityElement]::Escape("`$MPElement[Name=`"$($script:cvWinAlias)!Microsoft.Windows.Computer`"]`$")
            $valueXml = if ($isComputer) { "<Property>$propEsc</Property>" } else { "<HostProperty><MonitoringClass>$hostCls</MonitoringClass><Property>$propEsc</Property></HostProperty>" }
            if ([string]$row.MemberKind -eq 'NameOnly') {
                # Objects not hosted on the server they represent (Health Service
                # Watchers): match on Display Name = the agent FQDN or NetBIOS name.
                $sysAlias = $null
                foreach ($rn in @($Doc.SelectNodes('/ManagementPack/Manifest/References/Reference'))) {
                    if ([string]$rn.SelectSingleNode('ID').InnerText -eq 'System.Library') { $sysAlias = [string]$rn.GetAttribute('Alias') }
                }
                if (-not $sysAlias) {
                    $sysAlias = 'MigSys'
                    $sver = if ($Repository.ContainsKey('System.Library')) { [string]$Repository['System.Library'].Version } else { '7.5.8501.0' }
                    $rf = $Doc.SelectSingleNode('/ManagementPack/Manifest/References')
                    $fr = $Doc.CreateDocumentFragment()
                    $fr.InnerXml = "<Reference Alias=`"MigSys`"><ID>System.Library</ID><Version>$sver</Version><PublicKeyToken>31bf3856ad364e35</PublicKeyToken></Reference>"
                    [void]$rf.AppendChild($fr)
                }
                $dnProp = [System.Security.SecurityElement]::Escape("`$MPElement[Name=`"$sysAlias!System.Entity`"]/DisplayName`$")
                $valueXml = "<Property>$dnProp</Property>"
                # NAME or NAME.domain, any case
                # Display names can be upper or lower case; list both forms rather
                # than rely on an inline (?i) flag.
                $pats = @($pats | ForEach-Object {
                    if ($_ -match '^\^\((.*)\)\$$') {
                        $alts = @($Matches[1] -split '\|' | ForEach-Object { $_; $_.ToLowerInvariant() } | Select-Object -Unique)
                        '^(' + ($alts -join '|') + ')(\..*)?$'
                    } else { $_ } })
                $res.Note = "Health Service Watcher members -- matched on Display Name"
            }
            $regexXml = { param($o, $p) "<Expression><RegExExpression><ValueExpression>$valueXml</ValueExpression><Operator>$o</Operator><Pattern>$([System.Security.SecurityElement]::Escape($p))</Pattern></RegExExpression></Expression>" }
            $incXml = if ($pats.Count -eq 1) { & $regexXml $op $pats[0] } else { "<Expression><Or>$((@($pats | ForEach-Object { & $regexXml $op $_ })) -join '')</Or></Expression>" }
            $exPat = if ($row.PSObject.Properties['ExcludePattern']) { [string]$row.ExcludePattern } else { '' }
            if ($exPat) { $incXml = "<Expression><And>$incXml$(& $regexXml 'DoesNotMatchRegularExpression' $exPat)</And></Expression>" }

            if (-not $isComputer -and [string]$row.MemberKind -ne 'NameOnly') {
                $res.Note = "Hosted class $mcId -- matched on its host computer's NetBIOS name (HostProperty)"
            }

            $oldExpr = $mr.SelectSingleNode('Expression')
            $frag = $Doc.CreateDocumentFragment()
            if ($oldExpr) {
                $frag.InnerXml = "<Expression><Or><Expression>$($oldExpr.InnerXml)</Expression>$incXml</Or></Expression>"
                [void]$mr.ReplaceChild($frag, $oldExpr)
                $res.Note = ($res.Note + ' Existing dynamic expression kept (OR).').Trim()
            }
            else {
                $frag.InnerXml = $incXml
                $anchor = $mr.SelectSingleNode('RelationshipClass')
                if (-not $anchor) { $anchor = $mr.SelectSingleNode('MonitoringClass') }
                [void]$mr.InsertAfter($frag, $anchor)
            }
            if ($incNode) { [void]$mr.RemoveChild($incNode) }
            if ($excNode) { [void]$mr.RemoveChild($excNode) }
            $res.Action = 'ConvertedToDynamic'; $res.Operator = $op; $res.Pattern = [string]$row.Pattern; $res.ExcludePattern = $exPat
            $script:GroupConversionResults.Add($res)
        }
    }
}

$mpIndex = 0

foreach ($RootMP in $ImportOrder) {

    $mpIndex++

    # If -LiveImport already imported this MP live as a dependency (i.e.
    # it's not one of the originally-named -InputPath targets), skip the
    # candidate-generation pipeline for it entirely -- it's already in
    # SCOM as-is, a rewritten candidate file for it would be redundant and
    # was never the point. Only the MP(s) you actually named get the
    # rewrite/candidate treatment, live import or not.
    if ($LiveImport -and $OriginalInputTargets -notcontains $RootMP) {
        Write-Log "Skipping candidate generation for '$RootMP' -- already handled by live import in Step 3.5." -NoConsole
        continue
    }

    $srcEntry = $BatchSource[$RootMP]
    $SourceMP = $srcEntry.XML
    $RootVersion = $srcEntry.Version

    Write-Banner "[$mpIndex / $($ImportOrder.Count)] PROCESSING: $RootMP (v$RootVersion)"

    $safeRootMP = $RootMP -replace '[\\/:*?"<>|]', '_'

    ###########################################################
    # STATIC GROUP MEMBERSHIP CHECK (always, unless suppressed)
    ###########################################################
    # A group MP with an explicit <IncludeList> of object GUIDs will import
    # cleanly and populate EMPTY in the target, because those GUIDs are
    # source-environment instance IDs that don't exist in the new management
    # group. This surfaces that silent failure -- and, with a source
    # connection + -ResolveStaticGroupMembers, turns the dead GUIDs into a
    # named server list.
    if (-not $SkipStaticMembershipCheck) {
        # @() wrapping: Get-StaticGroupMembership returns a List that PowerShell
        # unrolls to $null (empty) or a scalar (single item) on return, which
        # makes .Count throw PropertyNotFoundStrict under Set-StrictMode Latest.
        # Forcing an array keeps .Count safe -- same class of fix as elsewhere.
        $staticGroups = @(Get-StaticGroupMembership -MPXml $SourceMP)
        if ($staticGroups.Count -gt 0) {
            $totalStatic = ($staticGroups | Measure-Object -Property StaticCount -Sum).Sum
            Write-Log "STATIC GROUP MEMBERSHIP DETECTED in '$RootMP': $($staticGroups.Count) group(s) with $totalStatic explicitly-listed member object GUID(s)." -Level WARN
            Write-Log "  These GUIDs are OLD-environment object IDs. Imported as-is, the group(s) will populate EMPTY in SCOM $TargetVersion (same servers get new GUIDs there), and anything targeting the group silently monitors nothing." -Level WARN

            # Try to attach a friendly group name from the MP's display strings.
            $dsMap = @{}
            try {
                foreach ($ds in @($SourceMP.ManagementPack.LanguagePacks.LanguagePack.DisplayStrings.DisplayString)) {
                    if ($ds.ElementID) { $dsMap[[string]$ds.ElementID] = [string]$ds.Name }
                }
            } catch { }

            # Optional GUID -> name resolution against the source environment.
            $nameLookup = $null
            if ($ResolveStaticGroupMembers) {
                if ($SourceManagementServer) {
                    $allGuids = @($staticGroups | ForEach-Object { $_.MemberGuids } | Select-Object -Unique)
                    Write-Log "  Resolving $($allGuids.Count) distinct member GUID(s) to server names against source '$SourceManagementServer' (read-only)..." -Level WARN
                    $nameLookup = Resolve-SourceObjectNames -Guids $allGuids -SourceServer $SourceManagementServer -Credential $SourceCredential
                }
                else {
                    Write-Log "  -ResolveStaticGroupMembers was set but there's no -SourceManagementServer connection to resolve names against. Reporting GUIDs only. Re-run with -SourceManagementServer to get server names." -Level WARN
                }
            }

            foreach ($sg in $staticGroups) {
                $friendly = if ($dsMap.ContainsKey($sg.GroupTargetID)) { $dsMap[$sg.GroupTargetID] } else { $sg.GroupTargetID }
                $dynNote = if ($sg.HasDynamicToo) { ' (also has a dynamic rule)' } else { '' }
                Write-Log "  - Group '$friendly': $($sg.StaticCount) static member(s)$dynNote" -Level WARN
                foreach ($g in $sg.MemberGuids) {
                    $resolvedName = if ($nameLookup -and $nameLookup.ContainsKey($g)) { $nameLookup[$g] } else { '' }
                    $script:StaticGroupReport.Add([PSCustomObject]@{
                        MP            = $RootMP
                        GroupID       = $sg.GroupTargetID
                        GroupName     = $friendly
                        MemberGuid    = $g
                        ResolvedName  = $resolvedName
                        HasDynamicToo = $sg.HasDynamicToo
                    })
                }
            }
            Write-Log "  Recommended handling: do NOT rely on the migrated static list. Dual-home these servers into SCOM $TargetVersion, then rebuild membership as a DYNAMIC rule (or re-add), using the resolved name list as your validation set." -Level WARN
        }
    }

    ###########################################################
    # ALIAS MAP (per-MP; each MP has its own References/Alias scope)
    ###########################################################

    $AliasMap = @{}

    # Wrapped in try/catch: at small-batch scale every test MP happened to
    # have a References block, but a genuinely minimal or unusual MP (a
    # real possibility across a 745-MP real-world export) can have NO
    # <References> element under <Manifest> at all -- not even an empty
    # one. In that case .Manifest.References itself is $null, and accessing
    # .Reference on $null throws PropertyNotFoundStrict under StrictMode.
    # This MP just has zero references; treat it that way instead of
    # crashing the whole batch run over it.
    $refs = $null
    try {
        $refs = $SourceMP.ManagementPack.Manifest.References.Reference
    }
    catch {
        Write-Log "Could not read References block for '$RootMP': $($_.Exception.Message)" -Level WARN -NoConsole
    }

    if (-not $refs) {
        Write-Log "Source MP has no References block under Manifest." -Level WARN
    }

    foreach ($r in $refs) {
        $AliasMap[$r.Alias] = [PSCustomObject]@{
            Alias   = $r.Alias
            ID      = $r.ID
            Version = $r.Version
        }
        Write-Log "$($r.Alias) -> $($r.ID)" -NoConsole
    }

    Write-Log "Alias Count : $($AliasMap.Count)"

    ###########################################################
    # OVERRIDE NODES + VALIDATION
    ###########################################################
    # 3.40: overrides live under <Monitoring><Overrides>. Earlier builds read
    # ManagementPack.Rules / .Monitors / .Discoveries, which do not exist in
    # the MP schema -- under StrictMode that threw, the catch swallowed it,
    # and NO override was ever validated. Every override element type is
    # read here (Monitor/Rule/Discovery/Diagnostic/Recovery property and
    # configuration overrides, category and secure-reference overrides).

    $sourceAliasToId = @{}
    foreach ($ak in $AliasMap.Keys) { $sourceAliasToId[$ak] = [string]$AliasMap[$ak].ID }

    $overrideNodes = @()
    $ovParentSrc = $SourceMP.SelectSingleNode('/ManagementPack/Monitoring/Overrides')
    if ($ovParentSrc) {
        $overrideNodes = @($ovParentSrc.ChildNodes | Where-Object { $_.NodeType -eq [System.Xml.XmlNodeType]::Element })
    }

    $Diagnostics = @()

    if ($overrideNodes.Count -eq 0) {
        Write-Log "No overrides found in MP. Skipping override validation." -NoConsole
    }
    else {

        Write-Log "Overrides Found: $($overrideNodes.Count)" -Level SUCCESS

        foreach ($o in $overrideNodes) {

            $ovId = [string]$o.GetAttribute('ID')
            $targetAttr = $null
            foreach ($an in @('Monitor', 'Rule', 'Discovery', 'Diagnostic', 'Recovery', 'SecureReference')) {
                if ($o.HasAttribute($an)) { $targetAttr = $an; break }
            }
            $target  = if ($targetAttr) { [string]$o.GetAttribute($targetAttr) } else { $null }
            $context = if ($o.HasAttribute('Context')) { [string]$o.GetAttribute('Context') } else { $null }

            if (-not $target) {
                $Diagnostics += [PSCustomObject]@{
                    Severity = "WARNING"; Type = "NoTargetAttribute"; OverrideID = $ovId; Target = $null; Context = $context; Resolved = $null
                    Message  = "Override has no Monitor/Rule/Discovery/Diagnostic/Recovery attribute ($($o.LocalName)) -- not validated"
                }
                continue
            }

            $tr = Test-ElementRef -Ref $target -AliasToId $sourceAliasToId -OwnId $RootMP
            if ($tr.Status -ne 'OK') {
                $Diagnostics += [PSCustomObject]@{
                    Severity = "ERROR"; Type = "MissingTarget:$($tr.Status)"; OverrideID = $ovId; Target = $target; Context = $context; Resolved = "$($tr.MPID)!$($tr.Element)"
                    Message  = switch ($tr.Status) {
                        'NO_MP'      { "Target MP '$($tr.MPID)' is not in the target repository or this batch" }
                        'NO_ELEMENT' { "MP '$($tr.MPID)' is present, but its current version has no element '$($tr.Element)' (renamed/removed by the vendor)" }
                        default      { "Override target could not be resolved ($($tr.Status))" }
                    }
                }
                continue
            }

            if ($context) {
                $cr = Test-ElementRef -Ref $context -AliasToId $sourceAliasToId -OwnId $RootMP
                if ($cr.Status -ne 'OK') {
                    $Diagnostics += [PSCustomObject]@{
                        Severity = "ERROR"; Type = "MissingContext:$($cr.Status)"; OverrideID = $ovId; Target = $target; Context = $context; Resolved = "$($cr.MPID)!$($cr.Element)"
                        Message  = "Override context class '$context' does not exist in the target"
                    }
                    continue
                }
            }

            $Diagnostics += [PSCustomObject]@{
                Severity = "INFO"; Type = "ValidOverride"; OverrideID = $ovId; Target = $target; Context = $context; Resolved = "$($tr.MPID)!$($tr.Element)"
                Message  = "Target and context exist in the target environment"
            }
        }
    }

    $ErrorDiagnostics = @($Diagnostics | Where-Object { $_.Severity -eq "ERROR" })

    Write-Log "Diagnostics -> ERROR: $($ErrorDiagnostics.Count), WARNING: $(@($Diagnostics | Where-Object { $_.Severity -eq 'WARNING' }).Count), INFO: $(@($Diagnostics | Where-Object { $_.Severity -eq 'INFO' }).Count)"

    if ($Strict -and $ErrorDiagnostics.Count -gt 0) {
        Write-Log "STRICT MODE: $($ErrorDiagnostics.Count) ERROR-severity diagnostic(s) found for '$RootMP'. Skipping candidate emission for this MP (continuing with the rest of the batch)." -Level ERROR
        $BatchResults[$RootMP] = [PSCustomObject]@{
            MPID                = $RootMP
            Version             = $RootVersion
            Status              = "FAILED_STRICT"
            Diagnostics         = $Diagnostics
            RewriteTable        = @()
            CandidatePath       = $null
            ImportPath          = $null
            WasSealed           = [bool]$srcEntry.WasSealed
            ReadyToImport       = $false
            BlockingReasons     = @("-Strict: $($ErrorDiagnostics.Count) override(s) point at elements that do not exist in the target. Re-run without -Strict to strip them automatically.")
            StrippedOverrides   = 0
            StrippedOther       = 0
            StrippedReferences  = 0
            StaticGroups        = 0
            OverridesKept       = 0
            InstanceRemapped    = 0
            InstanceDropped     = 0
            InstanceUnmapped    = 0
            IsDependency        = ($OriginalInputTargets -notcontains $RootMP)
            CompatibilityScore  = 0
            MigrationAdvice     = @()
            SemanticRecs        = @()
            ExternalRefs        = @($ExternalRefs[$RootMP])
            BatchDependsOn      = @($BatchGraph[$RootMP])
        }
        continue
    }
    elseif ($ErrorDiagnostics.Count -gt 0) {
        Write-Log "$($ErrorDiagnostics.Count) ERROR-severity diagnostic(s) found, but -Strict was not specified. Continuing; output will be written but may be incomplete." -Level WARN
    }

    ###########################################################
    # REWRITE TABLE (vs. target repository)
    ###########################################################

    $RewriteTable = @()

    foreach ($ref in $refs) {

        $oldID = $ref.ID

        # 3.40: an exact ID that already exists in the target (or is being
        # imported in this batch) is ALWAYS handled by the exact-ID branch
        # below. Previously the family pattern table ran first, so e.g.
        # Microsoft.Windows.InternetInformationServices.2016 (which still
        # exists in 2025) was "matched" to whichever IIS-family MP had the
        # highest version and -- with ID rewrite off -- got THAT MP's version
        # number stamped onto the old ID: a reference to a version that does
        # not exist, and a guaranteed import failure.
        $newEntry = $null
        if (-not $Repository.ContainsKey([string]$oldID) -and -not $BatchSource.ContainsKey([string]$oldID)) {
            $newEntry = Find-BestRepositoryMatch $oldID
        }

        if ($null -eq $newEntry) {

            if ($Repository.ContainsKey($oldID)) {
                $repoEntry = $Repository[$oldID]
                $refFamily = Get-MPFamily $oldID
                $refBehavior = if ($FamilyBehavior.ContainsKey($refFamily)) { $FamilyBehavior[$refFamily] } else { "Unified" }

                # CRITICAL: this exact-ID-match branch previously bumped to
                # ANY newer version found in the repo unconditionally, with
                # NO awareness of family behavior. That bypassed every
                # safeguard built elsewhere (ExactMatchOnly/CoreLibrary/
                # IISCommonLibrary) and was the actual cause of old IIS MPs
                # getting rewritten to reference the incompatible modern
                # unified CommonLibrary -- confirmed in practice: SCOM
                # rejected the resulting MP outright.
                #
                # IMPORTANT DISTINCTION (fixed after real-world testing):
                # that original failure was a DIFFERENT-MP substitution (old
                # IIS MP -> modern unified IIS CommonLibrary, a genuinely
                # different MP), which is handled in the family/pattern path,
                # NOT here. THIS branch is the exact-ID-match branch: the repo
                # MP has the SAME ID as the reference. For the same library ID,
                # SCOM's reference model is MINIMUM-version -- a reference to
                # v7.0.8438.6 is satisfied by an installed v10.25.10132.0 of
                # the same ID. (Confirmed in practice: manually editing the
                # old reference version up to match the new environment is the
                # standard, working fix.) So for ExactMatchOnly families here,
                # accept the repo version when it is EQUAL OR HIGHER, and only
                # flag it as unresolvable when the repo version is genuinely
                # LOWER than required (the one case a higher-min reference
                # truly can't bind to).
                if ($refBehavior -eq "ExactMatchOnly" -and -not $ForceRewriteCoreLibraries) {

                    # Compare as real versions; fall back to string equality
                    # only if either value isn't a parseable version.
                    $repoGE = $false          # repo >= required?
                    $repoEQ = $false          # repo == required?
                    $parsed = $false
                    try {
                        $vRepo = [version]$repoEntry.Version
                        $vRef  = [version]$ref.Version
                        $repoGE = ($vRepo -ge $vRef)
                        $repoEQ = ($vRepo -eq $vRef)
                        $parsed = $true
                    }
                    catch {
                        $repoEQ = ([string]$repoEntry.Version -eq [string]$ref.Version)
                        $repoGE = $repoEQ
                    }

                    if ($repoEQ) {
                        # Identical version -- reference is already satisfied,
                        # nothing to change.
                        $RewriteTable += [PSCustomObject]@{
                            Alias      = $ref.Alias
                            OldID      = $oldID
                            OldVersion = $ref.Version
                            NewID      = $repoEntry.ID
                            NewVersion = $repoEntry.Version
                            Action     = "UNCHANGED"
                            Source     = "TargetRepository"
                        }
                    }
                    elseif ($repoGE) {
                        # Repo has an EQUAL-OR-HIGHER version of the SAME core
                        # library ID -- bump the reference up to it. This is
                        # exactly the manual fix (edit old reference version to
                        # match the new environment) that is known to work, now
                        # done automatically. Distinct source tag so it's
                        # visible in the rewrite report as a core-library
                        # min-version bump, not a risky different-MP swap.
                        Write-Log "Reference '$oldID' (v$($ref.Version)) resolves to the SAME core library at a higher version (v$($repoEntry.Version)) in the target repository. SCOM binds minimum-version references to equal-or-higher builds of the same ID, so the reference is bumped up to v$($repoEntry.Version)." -Level INFO
                        $RewriteTable += [PSCustomObject]@{
                            Alias      = $ref.Alias
                            OldID      = $oldID
                            OldVersion = $ref.Version
                            NewID      = $repoEntry.ID
                            NewVersion = $repoEntry.Version
                            Action     = "VERSION_ONLY"
                            Source     = "CoreLibraryMinVersionBump"
                        }
                    }
                    else {
                        # Repo version is genuinely LOWER than the reference
                        # requires -- a higher-minimum reference cannot bind to
                        # a lower installed version. This is a real block.
                        Write-Log "Reference '$oldID' (v$($ref.Version)) is an exact-match-only dependency ($refFamily family) -- the target repository only has an OLDER version (v$($repoEntry.Version)), which cannot satisfy a reference requiring v$($ref.Version). This MP needs v$($ref.Version) or higher of $oldID present in SCOM $TargetVersion before it can import successfully." -Level WARN
                        $RewriteTable += [PSCustomObject]@{
                            Alias      = $ref.Alias
                            OldID      = $oldID
                            OldVersion = $ref.Version
                            NewID      = $oldID
                            NewVersion = $ref.Version
                            Action     = "UNCHANGED"
                            Source     = "ExactVersionMismatch"
                        }
                    }
                    continue
                }

                if ($refBehavior -eq "ExactMatchOnly" -and $ForceRewriteCoreLibraries -and [string]$repoEntry.Version -ne [string]$ref.Version) {
                    # The risky path: -ForceRewriteCoreLibraries explicitly
                    # asked for this safeguard to be bypassed. Still logged
                    # loudly and tagged distinctly in the rewrite report --
                    # never silent -- since this is the exact class of
                    # rewrite that has previously produced a candidate that
                    # looked clean but that SCOM rejected on actual import.
                    Write-Log "FORCED REWRITE: '$oldID' (v$($ref.Version) -> v$($repoEntry.Version)) is normally an exact-match-only dependency ($refFamily family), but -ForceRewriteCoreLibraries is set. This MAY NOT actually work -- SCOM has rejected this exact pattern before. Test the resulting candidate before trusting it." -Level WARN

                    $RewriteTable += [PSCustomObject]@{
                        Alias      = $ref.Alias
                        OldID      = $oldID
                        OldVersion = $ref.Version
                        NewID      = $repoEntry.ID
                        NewVersion = $repoEntry.Version
                        Action     = "VERSION_ONLY"
                        Source     = "ForcedCoreLibraryRewrite"
                    }
                    continue
                }

                $action = "UNCHANGED"
                try {
                    if ([version]$repoEntry.Version -gt [version]$ref.Version) {
                        $action = "VERSION_ONLY"
                    }
                } catch {}

                $RewriteTable += [PSCustomObject]@{
                    Alias      = $ref.Alias
                    OldID      = $oldID
                    OldVersion = $ref.Version
                    NewID      = $repoEntry.ID
                    NewVersion = $repoEntry.Version
                    Action     = $action
                    Source     = "TargetRepository"
                }
                continue
            }

            # Not found via pattern, not found by exact ID in target repo --
            # check whether it resolves within the BATCH itself (an internal
            # dependency that we are migrating side-by-side in this same run).
            if ($BatchSource.ContainsKey($oldID)) {
                $RewriteTable += [PSCustomObject]@{
                    Alias      = $ref.Alias
                    OldID      = $oldID
                    OldVersion = $ref.Version
                    NewID      = $oldID
                    NewVersion = $ref.Version
                    Action     = "UNCHANGED"
                    Source     = "BatchInternal"
                }
                continue
            }

            $RewriteTable += [PSCustomObject]@{
                Alias      = $ref.Alias
                OldID      = $oldID
                OldVersion = $ref.Version
                NewID      = ""
                NewVersion = ""
                Action     = "UNCHANGED"
                Source     = "NotFound"
            }

            continue
        }

        $action = "UPDATE"

        # NOTE: unlike the exact-ID-match branch above, this path (driven by
        # an explicit $ReferencePatterns entry) does not currently check
        # family behavior before rewriting. This is lower-risk today because
        # $ReferencePatterns is a small, manually-curated table and nothing
        # in it currently maps to an ExactMatchOnly family -- but if an
        # ExactMatchOnly-family rename is ever added there, apply the same
        # version-match guard used above before trusting this to rewrite it.
        if ($oldID -eq $newEntry.ID) {
            $action = "VERSION_ONLY"
        }

        $RewriteTable += [PSCustomObject]@{
            Alias      = $ref.Alias
            OldID      = $oldID
            OldVersion = $ref.Version
            NewID      = $newEntry.ID
            NewVersion = $newEntry.Version
            Action     = $action
            Source     = "TargetRepository"
        }
    }

    Write-Log "Rewrite candidates computed: $($RewriteTable.Count)"

    ###########################################################
    # EMIT CANDIDATE MP
    ###########################################################

    # IMPORTANT: building this from $SourceMP.OuterXml (a plain string) was
    # the cause of every candidate MP failing import with a generic "not
    # valid" error -- .OuterXml does NOT include the XML declaration node,
    # so the resulting document's XmlDeclaration was $null, and
    # XmlDocument.Save() omits the <?xml ... ?> line entirely when there is
    # no declaration to write. SCOM's MP parser rejects that. .Clone() is a
    # true deep-copy of the DOM, including the XmlDeclaration (and its
    # encoding attribute), so the saved file is byte-for-byte structurally
    # equivalent to the source except for the deliberate reference edits
    # made below.
    [xml]$OutputMP = $SourceMP.Clone()

    # Defensive belt-and-suspenders: if the source document somehow had no
    # XmlDeclaration of its own (unusual, but seen on a few hand-edited or
    # very old MPs), explicitly add a standard one rather than silently
    # saving without one again.
    if (-not $OutputMP.FirstChild -or $OutputMP.FirstChild.NodeType -ne [System.Xml.XmlNodeType]::XmlDeclaration) {
        $decl = $OutputMP.CreateXmlDeclaration("1.0", "utf-8", $null)
        $OutputMP.InsertBefore($decl, $OutputMP.DocumentElement) | Out-Null
    }

    # Static -> dynamic group conversion (unsealed MPs only; sealed MPs are
    # imported untouched from their original file).
    if (-not $srcEntry.WasSealed) {
        $gcBefore = $script:GroupConversionResults.Count
        Convert-StaticGroupRules -Doc $OutputMP -MPID $RootMP
        $gcNew = @($script:GroupConversionResults | Select-Object -Skip $gcBefore)
        $gcConv = @($gcNew | Where-Object { $_.Action -ne 'LeftStatic' }).Count
        if ($gcNew.Count -gt 0) {
            Write-Log "Group membership: $gcConv of $($gcNew.Count) static rule(s) converted to dynamic$(if ($gcNew.Count -gt $gcConv) { "; $($gcNew.Count - $gcConv) left static (will import EMPTY)" })." -Level $(if ($gcConv -eq $gcNew.Count) { 'SUCCESS' } else { 'WARN' })
        }
    }

    $RewriteResults = @()

    $outputRefs = $null
    try {
        $outputRefs = $OutputMP.ManagementPack.Manifest.References.Reference
    }
    catch {
        Write-Log "Could not read References block from candidate document for '$RootMP': $($_.Exception.Message)" -Level WARN -NoConsole
    }

    foreach ($ref in $outputRefs) {

        $row = $RewriteTable |
            Where-Object { $_.Alias -eq $ref.Alias } |
            Select-Object -First 1

        if ($null -eq $row) {
            continue
        }

        $oldID      = $ref.ID
        $oldVersion = $ref.Version

        switch ($row.Action) {
            "VERSION_ONLY" {
                if (-not [string]::IsNullOrWhiteSpace($row.NewVersion)) {
                    $ref.Version = $row.NewVersion
                }
            }

            "UPDATE" {
                # ID and version move TOGETHER or not at all -- the new
                # version number belongs to the new ID.
                if ($AllowIDRewrite -and -not [string]::IsNullOrWhiteSpace($row.NewID)) {
                    $ref.ID = $row.NewID
                    if (-not [string]::IsNullOrWhiteSpace($row.NewVersion)) { $ref.Version = $row.NewVersion }
                }
                else {
                    Write-Log "Superseded reference left unchanged for alias '$($row.Alias)' ($oldID -> $($row.NewID)); pass -AllowIDRewrite to repoint it. Anything using this alias is checked (and stripped if it is only overrides) below." -Level WARN -NoConsole
                }
            }

            default {
                # UNCHANGED - no action (covers BatchInternal / NotFound too)
            }
        }

        $RewriteResults += [PSCustomObject]@{
            Alias          = $ref.Alias
            OldID          = $oldID
            NewID          = $ref.ID
            OldVersion     = $oldVersion
            NewVersion     = $ref.Version
            RewriteID      = ($oldID -ne $ref.ID)
            RewriteVersion = ($oldVersion -ne $ref.Version)
            Source         = $row.Source
        }
    }

    ###########################################################
    # DEAD REFERENCE / DEAD ELEMENT STRIPPING  (rewritten in 3.40)
    ###########################################################
    # Two different things make a candidate fail import:
    #   (a) a <Reference> to an MP that will not be in the target at an
    #       equal-or-higher version (or is an unsigned copy of a sealed MP);
    #   (b) an override / category / folder item that points at an element
    #       which no longer exists in the version of the MP that IS there
    #       (vendor renamed/removed it -- very common across the Windows,
    #       SQL and IIS MP rebuilds).
    # For UNSEALED targets both are fixed by removing the dead override /
    # category / folder item (plus its display strings and knowledge), and a
    # dead <Reference> is removed ONLY if nothing else in the MP still uses its
    # alias. If a class, group, monitor, rule, view, etc. still uses it, the
    # reference is kept and the MP is reported BLOCKED with the exact elements
    # responsible -- removing the reference there would just trade one import
    # error for another. SEALED targets cannot be modified at all, so any
    # unresolvable reference blocks them.

    $isSealedTarget = [bool]$srcEntry.WasSealed
    $needsOriginal  = [bool]($srcEntry.PSObject.Properties['NeedsSealedOriginal'] -and $srcEntry.NeedsSealedOriginal)

    $candAlias = @{}
    # Sealed targets are imported from the original file, so judge them on
    # the ORIGINAL references, not on anything the rewrite step changed.
    $aliasDoc = if ($isSealedTarget) { $SourceMP } else { $OutputMP }
    $candRefParent = $aliasDoc.SelectSingleNode('/ManagementPack/Manifest/References')
    if ($candRefParent) {
        foreach ($rn in @($candRefParent.ChildNodes | Where-Object { $_.NodeType -eq [System.Xml.XmlNodeType]::Element })) {
            $idN = $rn.SelectSingleNode('ID'); $vN = $rn.SelectSingleNode('Version'); $kN = $rn.SelectSingleNode('PublicKeyToken')
            $candAlias[[string]$rn.GetAttribute('Alias')] = [PSCustomObject]@{
                ID       = if ($idN) { [string]$idN.InnerText } else { '' }
                Version  = if ($vN)  { [string]$vN.InnerText }  else { '' }
                KeyToken = if ($kN)  { [string]$kN.InnerText }  else { '' }
                Node     = $rn
            }
        }
    }
    $candAliasToId = @{}
    foreach ($ak in $candAlias.Keys) { $candAliasToId[$ak] = $candAlias[$ak].ID }

    # ---- (a) which referenced MPs will not be there at import time ----
    $unresolvableIDs = @{}

    if ($LiveImport -and (Test-Path variable:script:LiveImportResults) -and $script:LiveImportResults) {
        foreach ($depId in $script:LiveImportResults.Keys) {
            $depStatus = $script:LiveImportResults[$depId].Status
            if ($depStatus -in @("Failed", "Declined", "SkippedCascade")) { $unresolvableIDs[$depId] = "LiveImport:$depStatus" }
        }
    }
    foreach ($exId in $ExcludedSet.Keys) { $unresolvableIDs[$exId] = "ExcludedByRequest" }

    foreach ($al in $candAlias.Keys) {
        $ce = $candAlias[$al]
        $cid = $ce.ID
        $cver = $ce.Version
        if (-not $cid -or $unresolvableIDs.ContainsKey($cid)) { continue }

        if (-not $AuditOnly -and $BatchSource.ContainsKey($cid)) {
            # Imported earlier in this same batch (topological order) -- fine,
            # unless that batch member can't itself be imported.
            $bm = $BatchSource[$cid]
            $bmLower = $false
            try { $bmLower = ([version][string]$bm.Version -lt [version]$cver) } catch { }
            if ($bm.PSObject.Properties['NeedsSealedOriginal'] -and $bm.NeedsSealedOriginal) {
                $unresolvableIDs[$cid] = "BatchMemberNeedsSealedOriginal"
            }
            elseif ($bmLower) {
                $unresolvableIDs[$cid] = "BatchVersionLower:Have=v$($bm.Version),Need=v$cver"
            }
            elseif (-not $bm.WasSealed -and $ce.KeyToken) {
                # Unsealed MPs cannot be referenced at all.
                $unresolvableIDs[$cid] = "ReferencedMPIsUnsealed"
            }
            continue
        }

        $haveVer = $null
        $haveKey = $null
        if ($LiveImport) {
            $live = Get-InstalledMP -Name $cid
            if (-not $live) { $unresolvableIDs[$cid] = "NotInstalledInTarget"; continue }
            $haveVer = [string]$live.Version
            $haveKey = if ($live.PSObject.Properties['KeyToken'] -and $live.KeyToken) { [string]$live.KeyToken } else { '' }
        }
        else {
            if (-not $Repository.ContainsKey($cid)) { $unresolvableIDs[$cid] = "NotInTargetRepo"; continue }
            $haveVer = [string]$Repository[$cid].Version
        }

        $isLower = $false
        try { $isLower = ([version]$haveVer -lt [version]$cver) } catch { $isLower = ($haveVer -ne $cver) }
        if ($isLower) { $unresolvableIDs[$cid] = "TargetVersionLower:Have=v$haveVer,Need=v$cver"; continue }

        if ($LiveImport -and $ce.KeyToken -and $haveKey -ne $null) {
            if (-not $haveKey) { $unresolvableIDs[$cid] = "InstalledCopyIsUnsigned" }
            elseif ($haveKey -ne 'sealed' -and $haveKey -ne $ce.KeyToken) { $unresolvableIDs[$cid] = "KeyTokenMismatch:Installed=$haveKey,Need=$($ce.KeyToken)" }
        }
    }

    $unresolvableAliases = @{}
    foreach ($al in $candAlias.Keys) {
        if ($unresolvableIDs.ContainsKey($candAlias[$al].ID)) { $unresolvableAliases[$al] = $unresolvableIDs[$candAlias[$al].ID] }
    }

    # ---- (b) strip dead overrides / categories / folder items ----
    $strippedOverrideCount  = 0
    $strippedOtherCount     = 0
    $strippedReferenceCount = 0
    $strippedIds = New-Object 'System.Collections.Generic.HashSet[string]'
    $StripLog = New-Object System.Collections.Generic.List[object]

    function Get-DeadReason {
        param([System.Xml.XmlElement]$Node, [string[]]$Attrs)
        foreach ($an in $Attrs) {
            if (-not $Node.HasAttribute($an)) { continue }
            $val = [string]$Node.GetAttribute($an)
            if (-not $val) { continue }
            if ($val -like '*!*') {
                $al = $val.Substring(0, $val.IndexOf('!'))
                if ($unresolvableAliases.ContainsKey($al)) {
                    return "$an=$val -> MP '$($candAlias[$al].ID)' unavailable ($($unresolvableAliases[$al]))"
                }
            }
            elseif ($strippedIds.Contains($val)) {
                return "$an=$val -> local element was itself stripped"
            }
            $t = Test-ElementRef -Ref $val -AliasToId $candAliasToId -OwnId $RootMP
            if ($t.Status -eq 'NO_ELEMENT') { return "$an=$val -> '$($t.Element)' no longer exists in '$($t.MPID)' (target version)" }
            if ($t.Status -eq 'NO_ALIAS')   { return "$an=$val -> alias is not declared in References" }
        }
        return $null
    }

    $instRemapped = 0
    $instDropped = 0
    $instUnmapped = 0
    if (-not $isSealedTarget) {

        # Overrides aimed at ONE object (ContextInstance = 2016 object GUID).
        foreach ($ciNode in @($OutputMP.SelectNodes('/ManagementPack/Monitoring/Overrides/*[@ContextInstance]'))) {
            $oldG = ([string]$ciNode.GetAttribute('ContextInstance')).Trim().ToLowerInvariant()
            if ($script:InstanceMap) {
                if ($script:InstanceMap.ContainsKey($oldG)) {
                    $ciNode.SetAttribute('ContextInstance', $script:InstanceMap[$oldG])
                    $instRemapped++
                }
                else {
                    $oid = [string]$ciNode.GetAttribute('ID')
                    [void]$ciNode.ParentNode.RemoveChild($ciNode)
                    [void]$strippedIds.Add($oid)
                    $strippedOverrideCount++; $instDropped++
                    $StripLog.Add([PSCustomObject]@{ MP = $RootMP; Kind = $ciNode.LocalName; ElementID = $oid; Reason = "Targets one specific source object ($oldG) with no match in SCOM $TargetVersion (server retired, not yet an agent in 2025, or object type gone)" })
                }
            }
            else { $instUnmapped++ }
        }
        if ($instRemapped -gt 0 -or $instDropped -gt 0) {
            Write-Log "Per-object overrides: $instRemapped re-pointed to the target object, $instDropped removed (no matching object in the target)." -Level $(if ($instDropped) { 'WARN' } else { 'SUCCESS' })
        }
        if ($instUnmapped -gt 0) {
            Write-Log "$instUnmapped override(s) target one specific source object by ID. Without -InstanceMapFile they import but apply to NOTHING in SCOM $TargetVersion. Run 'Invoke-ScomMigrationStep.ps1 MapInstances' first." -Level WARN
        }

        $ovParent = $OutputMP.SelectSingleNode('/ManagementPack/Monitoring/Overrides')
        if ($ovParent) {
            foreach ($ov in @($ovParent.ChildNodes | Where-Object { $_.NodeType -eq [System.Xml.XmlNodeType]::Element })) {
                $why = Get-DeadReason -Node $ov -Attrs $script:OverrideRefAttributes
                if ($why) {
                    $oid = [string]$ov.GetAttribute('ID')
                    [void]$ovParent.RemoveChild($ov)
                    [void]$strippedIds.Add($oid)
                    $strippedOverrideCount++
                    $StripLog.Add([PSCustomObject]@{ MP = $RootMP; Kind = $ov.LocalName; ElementID = $oid; Reason = $why })
                }
            }
        }

        foreach ($cat in @($OutputMP.SelectNodes('/ManagementPack/Categories/Category'))) {
            $why = Get-DeadReason -Node $cat -Attrs @('Target')
            if ($why) {
                $cid2 = [string]$cat.GetAttribute('ID')
                [void]$cat.ParentNode.RemoveChild($cat)
                if ($cid2) { [void]$strippedIds.Add($cid2) }
                $strippedOtherCount++
                $StripLog.Add([PSCustomObject]@{ MP = $RootMP; Kind = 'Category'; ElementID = $cid2; Reason = $why })
            }
        }

        foreach ($fi in @($OutputMP.SelectNodes('/ManagementPack/Presentation/FolderItems/FolderItem'))) {
            $why = Get-DeadReason -Node $fi -Attrs @('ElementID', 'Folder')
            if ($why) {
                $fid = [string]$fi.GetAttribute('ID')
                [void]$fi.ParentNode.RemoveChild($fi)
                if ($fid) { [void]$strippedIds.Add($fid) }
                $strippedOtherCount++
                $StripLog.Add([PSCustomObject]@{ MP = $RootMP; Kind = 'FolderItem'; ElementID = $fid; Reason = $why })
            }
        }

        # ---- groups that cannot exist in the target ----
        # A membership rule that depends on an unavailable MP/element is
        # removed; a group left with no rules is removed entirely, and then
        # anything in this MP that points at a removed element (overrides,
        # folder items, views, monitors, other groups that nest it...) is
        # removed too, repeated until nothing else depends on what was removed.
        if (-not $KeepBlockedGroups) {
            $rxRef = New-Object System.Text.RegularExpressions.Regex '(?<![A-Za-z0-9_.])([A-Za-z_][A-Za-z0-9_]*)!([A-Za-z_][A-Za-z0-9_.]*)'
            $deadIn = {
                param($node)
                $texts = New-Object System.Collections.Generic.List[string]
                foreach ($d in @($node.SelectNodes('descendant-or-self::*'))) {
                    foreach ($at in @($d.Attributes)) { $texts.Add([string]$at.Value) }
                    foreach ($c in @($d.ChildNodes)) { if ($c.NodeType -eq [System.Xml.XmlNodeType]::Text -or $c.NodeType -eq [System.Xml.XmlNodeType]::CDATA) { $texts.Add([string]$c.Value) } }
                }
                foreach ($tx in $texts) {
                    if (-not $tx) { continue }
                    foreach ($m in $rxRef.Matches($tx)) {
                        $al = $m.Groups[1].Value
                        if (-not $candAlias.ContainsKey($al)) { continue }
                        if ($unresolvableAliases.ContainsKey($al)) { return "uses $($candAlias[$al].ID) ($($unresolvableAliases[$al]))" }
                        $t = Test-ElementRef -Ref "$al!$($m.Groups[2].Value.TrimEnd('.'))" -AliasToId $candAliasToId -OwnId $RootMP
                        if ($t.Status -eq 'NO_ELEMENT') { return "uses $($t.MPID)!$($t.Element), which no longer exists in the target version" }
                    }
                    foreach ($rid in $strippedIds) {
                        if ($tx -eq $rid -or $tx.Contains('Name="' + $rid + '"') -or $tx.Contains("Name='$rid'")) { return "points at '$rid', which was removed" }
                    }
                }
                return $null
            }

            $removedGroups = 0
            for ($pass = 1; $pass -le 10; $pass++) {
                $changed = $false

                foreach ($disc in @($OutputMP.SelectNodes('/ManagementPack/Monitoring/Discoveries/Discovery'))) {
                    $ds = $disc.SelectSingleNode('DataSource')
                    if (-not $ds -or [string]$ds.GetAttribute('TypeID') -notlike '*GroupPopulator*') { continue }
                    $rulesNode = $ds.SelectSingleNode('MembershipRules')
                    if (-not $rulesNode) { continue }
                    foreach ($mr in @($rulesNode.SelectNodes('MembershipRule'))) {
                        $why = & $deadIn $mr
                        if ($why) {
                            [void]$rulesNode.RemoveChild($mr); $changed = $true; $strippedOtherCount++
                            $StripLog.Add([PSCustomObject]@{ MP = $RootMP; Kind = 'GroupMembershipRule'; ElementID = [string]$disc.GetAttribute('ID'); Reason = "Membership rule $why" })
                        }
                    }
                    if (@($rulesNode.SelectNodes('MembershipRule')).Count -eq 0) {
                        $gi = $ds.SelectSingleNode('GroupInstanceId')
                        $gcls = if ($gi) { ([string]$gi.InnerText) -replace '^\$MPElement\[Name="?', '' -replace '"?\]\$$', '' } else { '' }
                        $did = [string]$disc.GetAttribute('ID')
                        [void]$disc.ParentNode.RemoveChild($disc); [void]$strippedIds.Add($did)
                        $gNode = if ($gcls -and $gcls -notlike '*!*') { $OutputMP.SelectSingleNode("/ManagementPack/TypeDefinitions/EntityTypes/ClassTypes/ClassType[@ID='$gcls']") } else { $null }
                        if ($gNode) { [void]$gNode.ParentNode.RemoveChild($gNode); [void]$strippedIds.Add($gcls) }
                        $gName = $gcls
                        try { $dn = $OutputMP.SelectSingleNode("//LanguagePack[@ID='ENU']/DisplayStrings/DisplayString[@ElementID='$gcls']/Name"); if ($dn) { $gName = "$($dn.InnerText) ($gcls)" } } catch { }
                        $removedGroups++; $changed = $true
                        $StripLog.Add([PSCustomObject]@{ MP = $RootMP; Kind = 'Group (DROPPED)'; ElementID = $gName; Reason = 'No membership rule left that can work in the target -- group removed so the rest of the MP can import' })
                    }
                }

                # Anything else in this MP pointing at a removed element.
                if ($strippedIds.Count -gt 0) {
                    foreach ($xp in @('/ManagementPack/Monitoring/*/*', '/ManagementPack/Presentation/*/*', '/ManagementPack/Categories/Category',
                                      '/ManagementPack/TypeDefinitions/EntityTypes/ClassTypes/ClassType', '/ManagementPack/TypeDefinitions/EntityTypes/RelationshipTypes/RelationshipType',
                                      '/ManagementPack/Reporting/*/*')) {
                        foreach ($el in @($OutputMP.SelectNodes($xp))) {
                            if (-not $el.ParentNode) { continue }
                            $eid = [string]$el.GetAttribute('ID')
                            if ($eid -and $strippedIds.Contains($eid)) { continue }
                            # Group discoveries are handled rule-by-rule above, so a
                            # parent group only loses the rule that nested the dropped one.
                            $gds = $el.SelectSingleNode('DataSource')
                            if ($el.LocalName -eq 'Discovery' -and $gds -and [string]$gds.GetAttribute('TypeID') -like '*GroupPopulator*') { continue }
                            $why = & $deadIn $el
                            if ($why -and $why -like 'points at*') {
                                [void]$el.ParentNode.RemoveChild($el); $changed = $true; $strippedOtherCount++
                                if ($eid) { [void]$strippedIds.Add($eid) }
                                $StripLog.Add([PSCustomObject]@{ MP = $RootMP; Kind = $el.LocalName; ElementID = $eid; Reason = "Removed with dropped group: $why" })
                            }
                        }
                    }
                }
                if (-not $changed) { break }
            }
            if ($removedGroups -gt 0) {
                Write-Log "Dropped $removedGroups group(s) whose members live in MPs/classes not available in the target (see $safeRootMP.Stripped.csv). The rest of the MP is kept." -Level WARN
            }
        }

        # Knowledge articles / display strings / image references written
        # for elements in OTHER MPs (company knowledge on Microsoft monitors,
        # e.g. ElementID="SQL2012!...") die with that MP or element -- remove
        # them the same way as dead overrides.
        foreach ($kn in @($OutputMP.SelectNodes('//LanguagePacks//*[@ElementID] | /ManagementPack/Presentation/ImageReferences/ImageReference[@ElementID]'))) {
            $eid = [string]$kn.GetAttribute('ElementID')
            if ($eid -notlike '*!*') { continue }
            $why = Get-DeadReason -Node $kn -Attrs @('ElementID')
            if ($why -and $kn.ParentNode) {
                [void]$kn.ParentNode.RemoveChild($kn)
                $strippedOtherCount++
                $StripLog.Add([PSCustomObject]@{ MP = $RootMP; Kind = $kn.LocalName; ElementID = $eid; Reason = $why })
            }
        }

        # Display strings, knowledge articles, image references etc. that
        # describe a stripped element must go too, or SCOM rejects them as
        # "refers to an invalid element".
        if ($strippedIds.Count -gt 0) {
            foreach ($n in @($OutputMP.SelectNodes('//*[@ElementID]'))) {
                if ($strippedIds.Contains([string]$n.GetAttribute('ElementID')) -and $n.ParentNode) {
                    [void]$n.ParentNode.RemoveChild($n)
                }
            }
        }

        # Remove containers left empty.
        foreach ($xp in @('/ManagementPack/Monitoring/Overrides', '/ManagementPack/Categories', '/ManagementPack/Presentation/FolderItems',
                          '/ManagementPack/Presentation/ImageReferences', '//LanguagePack/DisplayStrings', '//LanguagePack/KnowledgeArticles')) {
            foreach ($c in @($OutputMP.SelectNodes($xp))) {
                if (@($c.ChildNodes | Where-Object { $_.NodeType -eq [System.Xml.XmlNodeType]::Element }).Count -eq 0 -and $c.ParentNode) {
                    [void]$c.ParentNode.RemoveChild($c)
                }
            }
        }
    }

    # ---- remaining usage of each unresolvable alias ----
    function Find-AliasUsage {
        param([System.Xml.XmlDocument]$Doc, [string]$Alias)
        $hits = New-Object System.Collections.Generic.List[string]
        $rx = New-Object System.Text.RegularExpressions.Regex ('(?<![A-Za-z0-9_.])' + [regex]::Escape($Alias) + '!')
        foreach ($top in @($Doc.DocumentElement.ChildNodes)) {
            if ($top.NodeType -ne [System.Xml.XmlNodeType]::Element -or $top.LocalName -eq 'Manifest') { continue }
            foreach ($n in @($top.SelectNodes('descendant-or-self::*'))) {
                $hit = $false
                foreach ($at in @($n.Attributes)) { if ($rx.IsMatch([string]$at.Value)) { $hit = $true; break } }
                if (-not $hit) {
                    foreach ($c in @($n.ChildNodes)) {
                        if (($c.NodeType -eq [System.Xml.XmlNodeType]::Text -or $c.NodeType -eq [System.Xml.XmlNodeType]::CDATA) -and $rx.IsMatch([string]$c.Value)) { $hit = $true; break }
                    }
                }
                if ($hit) {
                    $owner = $n
                    while ($owner -and $owner.NodeType -eq [System.Xml.XmlNodeType]::Element -and -not $owner.HasAttribute('ID')) { $owner = $owner.ParentNode }
                    $label = if ($owner -and $owner.NodeType -eq [System.Xml.XmlNodeType]::Element) { "$($owner.LocalName) '$($owner.GetAttribute('ID'))'" } else { $n.LocalName }
                    if (-not $hits.Contains($label)) { $hits.Add($label) }
                }
            }
        }
        return , $hits
    }

    $BlockingReasons = New-Object System.Collections.Generic.List[string]
    if ($needsOriginal) {
        $BlockingReasons.Add("SEALED in the source environment but only an exported .xml was supplied -- add the original .mp/.mpb to -InputPath (or get the vendor's current build).")
    }

    foreach ($al in @($unresolvableAliases.Keys)) {
        $ce = $candAlias[$al]
        $usage = Find-AliasUsage -Doc $OutputMP -Alias $al   # returns a List; do NOT wrap in @() (that nests it)
        if ($usage.Count -eq 0 -and -not $isSealedTarget) {
            [void]$ce.Node.ParentNode.RemoveChild($ce.Node)
            $strippedReferenceCount++
            $StripLog.Add([PSCustomObject]@{ MP = $RootMP; Kind = 'Reference'; ElementID = "$al -> $($ce.ID) v$($ce.Version)"; Reason = $unresolvableAliases[$al] })
            $RewriteResults += [PSCustomObject]@{
                Alias = $al; OldID = $ce.ID; NewID = "(removed)"; OldVersion = $ce.Version; NewVersion = "(removed)"
                RewriteID = $false; RewriteVersion = $false; Source = "DeadReferenceStripped:$($unresolvableAliases[$al])"
            }
        }
        else {
            $usedBy = if ($usage.Count -gt 0) { " -- still used by $($usage.Count) element(s): $((@($usage) | Select-Object -First 6) -join ', ')$(if ($usage.Count -gt 6) { ', ...' })" } elseif ($isSealedTarget) { " -- sealed MP, cannot be edited" } else { "" }
            $BlockingReasons.Add("Needs $($ce.ID) v$($ce.Version) [$($unresolvableAliases[$al])]$usedBy")
        }
    }

    # ---- every remaining Alias!Element must exist in the target version ----
    # Overrides/categories/folder items were fixed above. Anything else (a
    # rule's Target, a monitor's TypeID, a discovery's data source, a group
    # formula's $MPElement[Name="..."]$, a view's target...) that points at an
    # element the target version of that MP no longer has cannot be fixed by
    # stripping -- it is reported so the MP is BLOCKED instead of "READY".
    $deadElementUse = @{}   # "MPID!Element" -> list of owning elements
    $liveAliases = @($candAlias.Keys | Where-Object { -not $unresolvableAliases.ContainsKey($_) -and $candAlias[$_].ID -and $ElementIndex.ContainsKey($candAlias[$_].ID) })
    if ($liveAliases.Count -gt 0) {
        $aliasAlt = (@($liveAliases | Sort-Object Length -Descending | ForEach-Object { [regex]::Escape($_) })) -join '|'
        $rxEl = New-Object System.Text.RegularExpressions.Regex ('(?<![A-Za-z0-9_.])(' + $aliasAlt + ')!([A-Za-z_][A-Za-z0-9_.]*)')
        foreach ($top in @($OutputMP.DocumentElement.ChildNodes)) {
            if ($top.NodeType -ne [System.Xml.XmlNodeType]::Element -or $top.LocalName -eq 'Manifest') { continue }
            foreach ($n in @($top.SelectNodes('descendant-or-self::*'))) {
                $texts = New-Object System.Collections.Generic.List[string]
                foreach ($at in @($n.Attributes)) { $texts.Add([string]$at.Value) }
                foreach ($c in @($n.ChildNodes)) {
                    if ($c.NodeType -eq [System.Xml.XmlNodeType]::Text -or $c.NodeType -eq [System.Xml.XmlNodeType]::CDATA) { $texts.Add([string]$c.Value) }
                }
                foreach ($tx in $texts) {
                    if (-not $tx -or $tx.IndexOf('!') -lt 0) { continue }
                    foreach ($m in $rxEl.Matches($tx)) {
                        $mAlias = $m.Groups[1].Value
                        $mEl = $m.Groups[2].Value.TrimEnd('.')
                        $mMp = $candAlias[$mAlias].ID
                        if ($ElementIndex[$mMp].Contains($mEl)) { continue }
                        $owner = $n
                        while ($owner -and $owner.NodeType -eq [System.Xml.XmlNodeType]::Element -and -not $owner.HasAttribute('ID')) { $owner = $owner.ParentNode }
                        $label = if ($owner -and $owner.NodeType -eq [System.Xml.XmlNodeType]::Element) { "$($owner.LocalName) '$($owner.GetAttribute('ID'))'" } else { $n.LocalName }
                        $k = "$mMp!$mEl"
                        if (-not $deadElementUse.ContainsKey($k)) { $deadElementUse[$k] = New-Object System.Collections.Generic.List[string] }
                        if (-not $deadElementUse[$k].Contains($label)) { $deadElementUse[$k].Add($label) }
                    }
                }
            }
        }
    }
    foreach ($k in @($deadElementUse.Keys | Sort-Object)) {
        $users = $deadElementUse[$k]
        $BlockingReasons.Add("Element '$k' does not exist in the $TargetVersion version of that MP -- used by $($users.Count) element(s): $((@($users) | Select-Object -First 6) -join ', ')$(if ($users.Count -gt 6) { ', ...' })")
        $StripLog.Add([PSCustomObject]@{ MP = $RootMP; Kind = 'NOT STRIPPED - BLOCKING'; ElementID = $k; Reason = "Used by: $(@($users) -join ', ')" })
    }

    if ($strippedOverrideCount + $strippedOtherCount + $strippedReferenceCount -gt 0) {
        Write-Log "Stripped from candidate: $strippedOverrideCount override(s), $strippedOtherCount category/folder item/knowledge/display string(s), $strippedReferenceCount reference(s) -- details in $safeRootMP.Stripped.csv" -Level WARN
        foreach ($sl in ($StripLog | Select-Object -First 15)) { Write-Log "  - [$($sl.Kind)] $($sl.ElementID): $($sl.Reason)" -Level WARN -NoConsole }
    }

    $ReadyToImport = ($BlockingReasons.Count -eq 0)
    if ($ReadyToImport) {
        Write-Log "Verdict: READY$(if ($isSealedTarget) { ' (sealed -- import the original file as-is)' })" -Level SUCCESS
    }
    else {
        Write-Log "Verdict: BLOCKED" -Level ERROR
        foreach ($br in $BlockingReasons) { Write-Log "  - $br" -Level ERROR }
    }

    # IMPORTANT: SCOM's Import-SCOMManagementPack enforces that the filename
    # (minus extension) must exactly match the MP's own Identity.ID inside
    # the XML -- it is NOT just a label. Earlier versions of this script
    # appended ".$TargetVersion" here for readability, which breaks import
    # with an "Identity mismatch" error. The candidate file MUST be named
    # exactly "$safeRootMP.xml"; version/target context goes in the report
    # filenames instead (those have no such constraint).
    # Sealed targets (and sealed-in-source MPs we only hold as .xml) are never
    # imported from a re-emitted XML, so their XML goes to the analysis-only
    # folder and the import path is the original sealed file (or nothing).
    $xmlIsImportable = (-not $isSealedTarget -and -not $needsOriginal)
    $outputFile = if ($xmlIsImportable) { Join-Path $CandidateFolder "$safeRootMP.xml" } else { Join-Path $AnalysisFolder "$safeRootMP.xml" }
    $reportFile = Join-Path $CandidateFolder "$safeRootMP.RewriteReport.csv"
    $stripFile  = Join-Path $CandidateFolder "$safeRootMP.Stripped.csv"
    if ($StripLog.Count -gt 0) {
        try { $StripLog | Export-Csv -LiteralPath $stripFile -NoTypeInformation -Encoding UTF8 } catch { }
        foreach ($sl in $StripLog) { $script:AllStripLog.Add($sl) }
    }
    $diagFile   = Join-Path $CandidateFolder "$safeRootMP.Diagnostics.csv"

    try {
        # IMPORTANT: XmlDocument.Save(string) writes a UTF-8 BOM whenever the
        # document's XmlDeclaration explicitly specifies encoding="UTF-8" --
        # which real Microsoft-shipped MPs do (confirmed: the source files
        # here declare encoding="UTF-8"). The original files are BOM-less,
        # so .Save() silently introduces a BOM that wasn't there before.
        # This is a known, documented .NET behavior (not a bug in our
        # control), and the standard workaround is to save through an
        # explicit XmlWriter with a BOM-less UTF8Encoding instead of letting
        # .Save(path) pick the encoding implicitly from the declaration.
        $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
        $writerSettings = New-Object System.Xml.XmlWriterSettings
        $writerSettings.Encoding = $utf8NoBom
        $writerSettings.Indent = $true

        $xmlWriter = [System.Xml.XmlWriter]::Create($outputFile, $writerSettings)
        try {
            $OutputMP.Save($xmlWriter)
        }
        finally {
            $xmlWriter.Close()
            $xmlWriter.Dispose()
        }

        Write-Log "Candidate MP saved: $outputFile" -Level SUCCESS
    }
    catch {
        Write-Log "Failed to save candidate MP to '$outputFile': $($_.Exception.Message)" -Level ERROR
        $outputFile = $null
    }

    try {
        $RewriteResults | Export-Csv -Path $reportFile -NoTypeInformation -Encoding UTF8
    }
    catch {
        Write-Log "Failed to save rewrite report to '$reportFile': $($_.Exception.Message)" -Level WARN
    }

    try {
        if ($Diagnostics.Count -gt 0) {
            $Diagnostics | Export-Csv -Path $diagFile -NoTypeInformation -Encoding UTF8
        }
    }
    catch {
        Write-Log "Failed to save diagnostics report to '$diagFile': $($_.Exception.Message)" -Level WARN
    }

    $idRewrites      = @($RewriteResults | Where-Object { $_.RewriteID }).Count
    $versionRewrites = @($RewriteResults | Where-Object { $_.RewriteVersion }).Count

    Write-Log "Version updates: $versionRewrites | ID updates: $idRewrites | Overrides stripped: $strippedOverrideCount | Other elements stripped: $strippedOtherCount | References stripped: $strippedReferenceCount"

    $ImportPath = if ($isSealedTarget) { [string]$srcEntry.SourcePath } elseif ($needsOriginal) { $null } else { $outputFile }

    ###########################################################
    # CLIENT-FACING EXPLAINER (only when -ForceRewriteCoreLibraries changed
    # something for THIS MP) -- teaches the underlying concept, not just
    # what changed, so this can be handed off and repeated without the
    # consultant on the call next time.
    ###########################################################

    $forcedRewritesForThisMP = @($RewriteTable | Where-Object { $_.Source -eq "ForcedCoreLibraryRewrite" })

    if ($forcedRewritesForThisMP.Count -gt 0) {

        $explainerFile = Join-Path $CandidateFolder "$safeRootMP.WhatChanged.md"

        $explainerLines = New-Object System.Collections.Generic.List[string]

        $explainerLines.Add('# What Changed in `' + $RootMP + '`, and Why')
        $explainerLines.Add("")
        $explainerLines.Add("This file explains, in plain language, what was edited in this Management Pack to prepare it for SCOM $TargetVersion -- and how to make the same kind of edit yourself on a different MP, without needing this script or a consultant on the call.")
        $explainerLines.Add("")
        $explainerLines.Add("## The short version")
        $explainerLines.Add("")
        $explainerLines.Add("Every Management Pack lists the other MPs it depends on, including an exact version number for each one. That version number is just a few digits sitting inside the MP's own file -- it does not update itself. When this MP was originally written (years ago, for an older SCOM), it recorded whatever version of a few foundational library MPs existed *at that time*. SCOM $TargetVersion now has newer versions of those same libraries, so the old number written inside this MP no longer matches what is actually installed.")
        $explainerLines.Add("")
        $explainerLines.Add("The fix is: update that recorded version number to match what is actually installed now. This script just did that for you, automatically, for the items listed below.")
        $explainerLines.Add("")
        $explainerLines.Add("## What was actually changed")
        $explainerLines.Add("")
        $explainerLines.Add("| Library | Old version (what the MP used to ask for) | New version (what SCOM $TargetVersion actually has) |")
        $explainerLines.Add("|---|---|---|")
        foreach ($fr in $forcedRewritesForThisMP) {
            $explainerLines.Add("| $($fr.OldID) | $($fr.OldVersion) | $($fr.NewVersion) |")
        }
        $explainerLines.Add("")
        $explainerLines.Add("## Why this is usually safe")
        $explainerLines.Add("")
        $explainerLines.Add("For most MPs -- especially ones that are mainly Overrides (turning monitors/rules on or off, changing thresholds) rather than introducing brand-new custom monitoring logic -- the library version number is just a reference declaration. The MP isn't relying on anything specific that changed between the old and new library version; it just needs *some* valid, currently-installed version on file. Bumping the number to match is usually enough.")
        $explainerLines.Add("")
        $explainerLines.Add("## When this is NOT safe to assume")
        $explainerLines.Add("")
        $explainerLines.Add("If an MP defines its own custom Classes, Discoveries, Monitors, or Rules that build directly on specific things defined inside that library (not just an Override), a version bump alone may not be enough -- the underlying library could have renamed or removed something this MP's custom logic depends on. In that case, the import may still fail even after the version number is corrected, and that failure needs to be looked at on its own merits, not assumed away.")
        $explainerLines.Add("")
        $explainerLines.Add("This script flagged this rewrite because it could not verify which of these two situations applies -- it bumped the number because you explicitly asked it to (`-ForceRewriteCoreLibraries`), but the result still needs to be tested by actually attempting the import.")
        $explainerLines.Add("")
        $explainerLines.Add("## How to do this yourself, by hand, on a different MP")
        $explainerLines.Add("")
        $explainerLines.Add("1. Open the MP's `.xml` file in any plain text editor (Notepad is fine).")
        $explainerLines.Add("2. Find the `<Manifest><References>` section near the top of the file.")
        $explainerLines.Add("3. Each `<Reference>` block looks like this:")
        $explainerLines.Add('   ```xml')
        $explainerLines.Add('   <Reference Alias="SC">')
        $explainerLines.Add('     <ID>Microsoft.SystemCenter.Library</ID>')
        $explainerLines.Add('     <Version>7.0.8438.6</Version>')
        $explainerLines.Add('     <PublicKeyToken>31bf3856ad364e35</PublicKeyToken>')
        $explainerLines.Add('   </Reference>')
        $explainerLines.Add('   ```')
        $explainerLines.Add("4. Find out what version is *actually* installed in your SCOM $TargetVersion environment for that same `<ID>` -- from the SCOM console (Administration > Management Packs, check the Version column), or with PowerShell:")
        $explainerLines.Add('   ```powershell')
        $explainerLines.Add('   Get-SCOMManagementPack -Name "Microsoft.SystemCenter.Library" | Select-Object Name, Version')
        $explainerLines.Add('   ```')
        $explainerLines.Add("5. Replace the number inside `<Version>...</Version>` with that installed version.")
        $explainerLines.Add("6. Save the file, then try importing it.")
        $explainerLines.Add('7. If SCOM rejects it, the error message will usually name the EXACT library and version it still expects (if running the import yourself in PowerShell, check $Error[0].Exception.InnerException for the full detail) -- repeat the same steps for that one too.')
        $explainerLines.Add("")
        $explainerLines.Add("---")
        $explainerLines.Add("*Generated automatically by the SCOM MP Batch Compiler (`-ForceRewriteCoreLibraries`). This explains what changed in this one file; it is not a guarantee the import will succeed -- test it.*")

        try {
            Set-Content -LiteralPath $explainerFile -Value $explainerLines -Encoding UTF8
            Write-Log "Client-facing explainer saved: $explainerFile" -Level SUCCESS
        }
        catch {
            Write-Log "Failed to save client-facing explainer to '$explainerFile': $($_.Exception.Message)" -Level WARN
        }
    }

    ###########################################################
    # MIGRATION ADVISOR (per-MP)
    ###########################################################
    # Beyond "found / missing", this now flags:
    #   - PerVersion families (e.g. SharePoint) where blindly picking the
    #     newest repo candidate would be WRONG -- these need a human to
    #     confirm the matching product version.
    #   - Known deprecation notices (e.g. SSRS/SSAS) for MPs that still
    #     work today but have an announced end-of-support date.
    # Missing references are also appended to the batch-wide
    # $MissingDependencies collector (see Step 9) with a concrete lookup
    # hint, instead of just being reported as "MISSING" in isolation.

    $MigrationAdvice = @()

    foreach ($ref in $refs) {

        $oldID  = $ref.ID
        $family = Get-MPFamily $oldID
        $behavior = if ($FamilyBehavior.ContainsKey($family)) { $FamilyBehavior[$family] } else { "Unified" }

        $deprecationNote = $null
        foreach ($depKey in $DeprecationNotices.Keys) {
            if ($oldID -like "$depKey*") { $deprecationNote = $DeprecationNotices[$depKey] }
        }

        if (-not $AuditOnly -and $BatchSource.ContainsKey($oldID)) {
            # Skipped entirely under -AuditOnly: when -InputPath is an
            # entire source environment rather than a real migration batch,
            # "this reference is satisfied by another MP in the same batch"
            # is true but meaningless -- it just means two old MPs from the
            # same era are mutually consistent, not that either is actually
            # compatible with the target. Falling through lets the real
            # checks below (exact match / family behavior / missing) decide.
            $MigrationAdvice += [PSCustomObject]@{
                Reference      = $oldID
                Family         = $family
                Status         = "BATCH_INTERNAL"
                Candidates     = 1
                Recommendation = "Resolved within this batch -- ensure it is imported first (see import order)"
                Deprecation    = $deprecationNote
            }
        }
        elseif ($behavior -eq "PerVersion" -and $FamilyTable.ContainsKey($family)) {
            # Family exists in the repo, but versions don't unify -- flag
            # for manual confirmation rather than silently picking one.
            $MigrationAdvice += [PSCustomObject]@{
                Reference      = $oldID
                Family         = $family
                Status         = "NEEDS_REVIEW"
                Candidates     = $FamilyTable[$family].Count
                Recommendation = "This MP family ($family) does not unify across versions -- confirm which target-version MP actually matches before importing. Candidates present in repo: $($FamilyTable[$family] -join ', ')"
                Deprecation    = $deprecationNote
            }
        }
        elseif ($Repository.ContainsKey($oldID)) {
            # Exact ID match in target repo. For ExactMatchOnly families the
            # VERSION must be compatible -- but "compatible" means SCOM's
            # minimum-version rule: an EQUAL-OR-HIGHER repo version of the
            # SAME library ID satisfies the reference (the manual fix that is
            # known to work -- bump the old reference version up to the new
            # environment's). Only a genuinely LOWER repo version is missing.
            # Mirror the rewrite engine's check exactly so scoring and the
            # actual rewrite always agree.
            $repoVersionHere = $Repository[$oldID].Version

            # Determine relationship as real versions (fall back to string
            # (in)equality if either can't be parsed).
            $repoIsLower = $false
            $repoIsHigher = $false
            try {
                $repoIsLower  = ([version]$repoVersionHere -lt [version]$ref.Version)
                $repoIsHigher = ([version]$repoVersionHere -gt [version]$ref.Version)
            }
            catch {
                $neq = ([string]$repoVersionHere -ne [string]$ref.Version)
                $repoIsLower = $neq   # can't tell direction; treat as needing attention
            }

            if ($behavior -eq "ExactMatchOnly" -and -not $ForceRewriteCoreLibraries -and $repoIsLower) {
                # Repo has an OLDER version than required -- genuinely can't
                # satisfy a higher-minimum reference. Real gap.
                $MigrationAdvice += [PSCustomObject]@{
                    Reference      = $oldID
                    Family         = $family
                    Status         = "MISSING"
                    Candidates     = 0
                    Recommendation = "Target repository has an OLDER version (v$repoVersionHere) of this dependency than required (v$($ref.Version)). A lower installed version cannot satisfy a higher-minimum reference -- v$($ref.Version) or higher must be present in SCOM $TargetVersion."
                    Deprecation    = $deprecationNote
                }

                $script:MissingDependencies += [PSCustomObject]@{
                    RequiredBy      = $RootMP
                    MissingID       = $oldID
                    RequestedVersion= $ref.Version
                    Family          = $family
                    SuggestedSource = "Need v$($ref.Version) or higher -- only v$repoVersionHere found in target repo. Import a newer build into SCOM $TargetVersion."
                    LookupUrl       = "https://systemcenter.wiki/?Get-ManagementPack=$oldID"
                }
            }
            elseif ($behavior -eq "ExactMatchOnly" -and -not $ForceRewriteCoreLibraries -and $repoIsHigher) {
                # Repo has an EQUAL-OR-HIGHER version of the SAME core library
                # ID -- the reference is bumped up to it (min-version binding).
                # This is a genuine, safe resolution, not a risky different-MP
                # swap, so it counts as FOUND.
                $MigrationAdvice += [PSCustomObject]@{
                    Reference      = $oldID
                    Family         = $family
                    Status         = "FOUND"
                    Candidates     = 1
                    Recommendation = "Reference bumped to the same core library at a higher version (v$($ref.Version) -> v$repoVersionHere). SCOM binds minimum-version references to equal-or-higher builds of the same ID."
                    Deprecation    = $deprecationNote
                }
            }
            elseif ($behavior -eq "ExactMatchOnly" -and $ForceRewriteCoreLibraries -and ($repoIsLower -or $repoIsHigher)) {
                # -ForceRewriteCoreLibraries is set, so this is being
                # rewritten rather than flagged -- still reported distinctly
                # from a genuine, safe match, since the forced path carries
                # real risk.
                $MigrationAdvice += [PSCustomObject]@{
                    Reference      = $oldID
                    Family         = $family
                    Status         = "FOUND"
                    Candidates     = 1
                    Recommendation = "FORCED REWRITE (-ForceRewriteCoreLibraries): rewritten v$($ref.Version) -> v$repoVersionHere despite being an exact-match-only dependency. This is not guaranteed to actually import successfully -- test before trusting it."
                    Deprecation    = $deprecationNote
                }
            }
            else {
                $MigrationAdvice += [PSCustomObject]@{
                    Reference      = $oldID
                    Family         = $family
                    Status         = "FOUND"
                    Candidates     = 1
                    Recommendation = "Safe to upgrade"
                    Deprecation    = $deprecationNote
                }
            }
        }
        elseif ($behavior -eq "Unified" -and $FamilyTable.ContainsKey($family) -and $family -ne 'Unknown' -and (Get-BestCandidate -Family $family -ReferenceID $oldID)) {
            # NOT an exact-ID match, but this family is known to unify
            # across versions (IIS/SQL/Windows/etc.) AND there's an actual,
            # specific newer MP in the repo this reference can be rewritten
            # to point at -- checked via the same Get-BestCandidate logic
            # the rewrite engine itself uses, so "FOUND" here means a real
            # rewrite target exists, not just "some MP in this family
            # happens to be present somewhere in the repo."
            # 3.40: this used to report FOUND / "Safe to upgrade", but nothing
            # actually repoints the reference unless -AllowIDRewrite is used,
            # so the MP scored 100% and then failed import. The old MP ID does
            # NOT exist in the target; say so.
            $candidateCount = $FamilyTable[$family].Count
            $successor = Get-BestCandidate -Family $family -ReferenceID $oldID

            $MigrationAdvice += [PSCustomObject]@{
                Reference      = $oldID
                Family         = $family
                Status         = "SUPERSEDED"
                Candidates     = $candidateCount
                Recommendation = "This MP ID does not exist in SCOM $TargetVersion; its family was rebuilt (current: $successor). Overrides against it are stripped automatically; anything else using it must be re-authored against the new MP."
                Deprecation    = $deprecationNote
            }
            $script:MissingDependencies += [PSCustomObject]@{
                RequiredBy      = $RootMP
                MissingID       = $oldID
                RequestedVersion= $ref.Version
                Family          = $family
                SuggestedSource = "Superseded in $TargetVersion by $successor (different MP ID and element IDs) -- not installable; re-author against the new MP if still needed."
                LookupUrl       = "https://systemcenter.wiki/?Get-ManagementPack=$oldID"
            }
        }
        elseif ($ExcludedSet.ContainsKey($oldID)) {
            # Explicitly excluded via -ExcludeMPIDs -- a known, intentional
            # gap (e.g. confirmed in earlier testing that the required
            # version doesn't exist anywhere accessible). Still reported as
            # MISSING (whatever needs it genuinely doesn't have it), but
            # tagged distinctly so the report can tell "you told me to skip
            # this" apart from "couldn't find this."
            $MigrationAdvice += [PSCustomObject]@{
                Reference      = $oldID
                Family         = $family
                Status         = "MISSING"
                Candidates     = 0
                Recommendation = "Explicitly excluded via -ExcludeMPIDs -- known, intentional gap, not searched for."
                Deprecation    = $deprecationNote
            }

            $script:MissingDependencies += [PSCustomObject]@{
                RequiredBy      = $RootMP
                MissingID       = $oldID
                RequestedVersion= $ref.Version
                Family          = $family
                SuggestedSource = "EXCLUDED BY REQUEST (-ExcludeMPIDs) -- not searched for. Remove from -ExcludeMPIDs to have this auto-resolved again."
                LookupUrl       = "https://systemcenter.wiki/?Get-ManagementPack=$oldID"
            }
        }
        else {
            $lookupHint = Get-MPLookupHint -MPID $oldID -Family $family

            $MigrationAdvice += [PSCustomObject]@{
                Reference      = $oldID
                Family         = $family
                Status         = "MISSING"
                Candidates     = 0
                Recommendation = "Not found in target ($TargetVersion) repository or this batch. Suggested source: $($lookupHint.SuggestedSource)"
                Deprecation    = $deprecationNote
            }

            # Accumulate into the batch-wide missing-dependencies report
            # (see Step 9) so it's visible as one list across the whole
            # batch, not buried per-MP.
            $script:MissingDependencies += [PSCustomObject]@{
                RequiredBy      = $RootMP
                MissingID       = $oldID
                RequestedVersion= $ref.Version
                Family          = $family
                SuggestedSource = $lookupHint.SuggestedSource
                LookupUrl       = $lookupHint.LookupUrl
            }
        }
    }

    $missingCount = @($MigrationAdvice | Where-Object { $_.Status -eq "MISSING" -or $_.Status -eq "SUPERSEDED" }).Count
    $needsReviewCount = @($MigrationAdvice | Where-Object { $_.Status -eq "NEEDS_REVIEW" }).Count
    $totalRefs    = [Math]::Max(1, $MigrationAdvice.Count)
    # NEEDS_REVIEW counts as a half-penalty -- it's not missing, but it's
    # not a safe auto-upgrade either, so it shouldn't score identically to
    # a clean FOUND.
    $CompatibilityScore = [Math]::Round(100 - ((($missingCount + ($needsReviewCount * 0.5)) / $totalRefs) * 100))
    if ($CompatibilityScore -lt 0) { $CompatibilityScore = 0 }

    Write-Log "Compatibility Score: $CompatibilityScore% (Missing: $missingCount, Needs Review: $needsReviewCount)" -Level SUCCESS

    $deprecatedRefs = @($MigrationAdvice | Where-Object { $_.Deprecation })
    if ($deprecatedRefs.Count -gt 0) {
        Write-Log ""
        Write-Log "DEPRECATION NOTICES for this MP's references:" -Level WARN
        foreach ($d in $deprecatedRefs) {
            Write-Log "  - $($d.Reference): $($d.Deprecation)" -Level WARN
        }
    }

    ###########################################################
    # SEMANTIC REWRITE RECOMMENDATIONS (per-MP)
    ###########################################################

    $SemanticRecommendations = @()

    foreach ($ref in $refs) {

        $family    = Get-MPFamily $ref.ID
        $candidate = Get-BestCandidate -Family $family -ReferenceID $ref.ID

        $status     = "UNKNOWN"
        $confidence = "LOW"

        if (-not $AuditOnly -and $BatchSource.ContainsKey($ref.ID)) {
            $status     = "BATCH_INTERNAL"
            $confidence = "HIGH"
            $candidate  = $ref.ID
        }
        elseif ($ExcludedSet.ContainsKey($ref.ID)) {
            $status     = "EXCLUDED"
            $confidence = "HIGH"
        }
        elseif ($Repository.ContainsKey($ref.ID)) {
            # Exact ID in target repo -- but for ExactMatchOnly families,
            # the VERSION must also match (same reasoning as the advisor
            # above): a different version of a CoreLibrary/SystemCenter/
            # IISCommonLibrary-family MP is not a safe substitute, even
            # though the ID is technically present.
            $behaviorHere = if ($FamilyBehavior.ContainsKey($family)) { $FamilyBehavior[$family] } else { "Unified" }

            if ($behaviorHere -eq "ExactMatchOnly" -and -not $ForceRewriteCoreLibraries -and [string]$Repository[$ref.ID].Version -ne [string]$ref.Version) {
                $status     = "MISSING"
                $confidence = "LOW"
            }
            elseif ($behaviorHere -eq "ExactMatchOnly" -and $ForceRewriteCoreLibraries -and [string]$Repository[$ref.ID].Version -ne [string]$ref.Version) {
                $status     = "FOUND"
                $confidence = "LOW"
                $candidate  = $ref.ID
            }
            else {
                $status     = "FOUND"
                $confidence = "HIGH"
                $candidate  = $ref.ID
            }
        }
        elseif ($candidate) {
            # Get-BestCandidate found an actual, specific rewrite target --
            # not just "the family exists somewhere in the repo."
            $status     = "FOUND"
            $confidence = "HIGH"
        }
        else {
            $status     = "MISSING"
            $confidence = "LOW"
        }

        $SemanticRecommendations += [PSCustomObject]@{
            Reference  = $ref.ID
            Family     = $family
            Candidate  = $candidate
            Confidence = $confidence
            Status     = $status
        }
    }

    try {
        $semanticFile = Join-Path $CandidateFolder "$safeRootMP.SemanticRecommendations.csv"
        $SemanticRecommendations | Export-Csv -Path $semanticFile -NoTypeInformation -Encoding UTF8
    }
    catch {
        Write-Log "Failed to save semantic recommendations report: $($_.Exception.Message)" -Level WARN
    }

    ###########################################################
    # RECORD RESULT FOR THIS MP
    ###########################################################

    # Dependencies for import ordering / cascade-skip must come from the FINAL
    # candidate: a reference removed by stripping (e.g. a third-party MP after its
    # group was dropped) is no longer a dependency.
    $finalRefDoc = if ($isSealedTarget) { $SourceMP } else { $OutputMP }
    $finalRefIds = @($finalRefDoc.SelectNodes('/ManagementPack/Manifest/References/Reference/ID') | ForEach-Object { [string]$_.InnerText })
    $finalBatchDeps = @($BatchGraph[$RootMP] | Where-Object { $finalRefIds -contains $_ })

    # Static groups still left in the CANDIDATE (after any conversion).
    $staticGroupCountForMP = @(Get-StaticGroupMembership -MPXml $OutputMP).Count

    $BatchResults[$RootMP] = [PSCustomObject]@{
        MPID                = $RootMP
        Version             = $RootVersion
        Status              = "PROCESSED"
        Diagnostics         = $Diagnostics
        RewriteTable        = $RewriteTable
        CandidatePath       = $outputFile
        ImportPath          = $ImportPath
        WasSealed           = $isSealedTarget
        ReadyToImport       = $ReadyToImport
        BlockingReasons     = @($BlockingReasons)
        StrippedOverrides   = $strippedOverrideCount
        StrippedOther       = $strippedOtherCount
        StrippedReferences  = $strippedReferenceCount
        StaticGroups        = $staticGroupCountForMP
        OverridesKept       = @($OutputMP.SelectNodes('/ManagementPack/Monitoring/Overrides/*')).Count
        InstanceRemapped    = $instRemapped
        InstanceDropped     = $instDropped
        InstanceUnmapped    = $instUnmapped
        IsDependency        = ($OriginalInputTargets -notcontains $RootMP)
        CompatibilityScore  = $CompatibilityScore
        MigrationAdvice     = $MigrationAdvice
        SemanticRecs        = $SemanticRecommendations
        ExternalRefs        = @($ExternalRefs[$RootMP])
        BatchDependsOn      = $finalBatchDeps
    }
}

Write-Banner "STEPS 4-8 COMPLETE"
Write-Log "MPs processed : $(@($BatchResults.Values | Where-Object { $_.Status -eq 'PROCESSED' }).Count)"
Write-Log "MPs failed    : $(@($BatchResults.Values | Where-Object { $_.Status -ne 'PROCESSED' }).Count)"

###########################################################
# STEP 9 - BATCH MANIFEST + IMPORT SEQUENCE
###########################################################
# Rolls every per-MP result up into one ordered manifest CSV and one
# generated PowerShell snippet that calls Import-SCOMManagementPack in the
# correct order. Per your migration plan, this script does NOT connect to
# a live SCOM server or import anything -- it only prepares this for you to
# review and run by hand against your SCOM 2025 management server.

Write-Banner "STEP 9 - BATCH MANIFEST"

$ManifestRows = @()
$position = 0

foreach ($mpId in $ImportOrder) {

    $position++
    $result = $BatchResults[$mpId]

    if ($null -eq $result) {
        # Should not happen, but guards against an MP that was in the order
        # but somehow never got a results entry recorded.
        continue
    }

    $isCyclic = $CyclicNodes -contains $mpId

    $ManifestRows += [PSCustomObject]@{
        Order              = $position
        MPID               = $result.MPID
        WorkbookName       = if ($script:ManifestNameMap.ContainsKey($result.MPID)) { $script:ManifestNameMap[$result.MPID] } else { '' }
        ManifestAction     = if ($script:ManifestActionMap.ContainsKey($result.MPID)) { $script:ManifestActionMap[$result.MPID] } elseif ($result.IsDependency) { 'DEPENDENCY (auto-resolved)' } else { '' }
        SourceVersion      = $result.Version
        Status             = $result.Status
        Verdict            = if ($result.ReadyToImport) { 'READY' } else { 'BLOCKED' }
        BlockingReasons    = ($result.BlockingReasons -join ' | ')
        Sealed             = $result.WasSealed
        ImportPath         = $result.ImportPath
        StrippedOverrides  = $result.StrippedOverrides
        StrippedOther      = $result.StrippedOther
        StrippedReferences = $result.StrippedReferences
        StaticGroups       = $result.StaticGroups
        OverridesKept      = $result.OverridesKept
        InstanceRemapped   = $result.InstanceRemapped
        InstanceDropped    = $result.InstanceDropped
        InstanceUnmapped   = $result.InstanceUnmapped
        CompatibilityScore = $result.CompatibilityScore
        DeadOverrideTargets= @($result.Diagnostics | Where-Object { $_.Severity -eq 'ERROR' }).Count
        WarningCount       = @($result.Diagnostics | Where-Object { $_.Severity -eq 'WARNING' }).Count
        BatchDependsOn     = ($result.BatchDependsOn -join '; ')
        ExternalRefs       = ($result.ExternalRefs -join '; ')
        CyclicFlag         = $isCyclic
        CandidatePath      = $result.CandidatePath
    }
}

$manifestFile = Join-Path $OutputFolder "BatchImportManifest.csv"

try {
    $ManifestRows | Export-Csv -Path $manifestFile -NoTypeInformation -Encoding UTF8
    Write-Log "Batch import manifest saved: $manifestFile" -Level SUCCESS
}
catch {
    Write-Log "Failed to save batch import manifest: $($_.Exception.Message)" -Level WARN
}

if ($script:GroupConversionResults.Count -gt 0) {
    $gcFile = Join-Path $OutputFolder "GroupConversionResults.csv"
    $script:GroupConversionResults | Export-Csv -LiteralPath $gcFile -NoTypeInformation -Encoding UTF8
    $gcDyn = @($script:GroupConversionResults | Where-Object { $_.Action -eq 'ConvertedToDynamic' }).Count
    $gcNest = @($script:GroupConversionResults | Where-Object { $_.Action -eq 'NestedGroup' }).Count
    $gcLeft = @($script:GroupConversionResults | Where-Object { $_.Action -eq 'LeftStatic' }).Count
    Write-Log "Static group rules: $gcDyn converted to dynamic NetBIOS patterns, $gcNest nested subgroups fixed, $gcLeft left static. Detail: $gcFile" -Level $(if ($gcLeft -gt 0) { 'WARN' } else { 'SUCCESS' })
}

if ($script:AllStripLog.Count -gt 0) {
    $allStripFile = Join-Path $OutputFolder "StrippedElements.csv"
    try {
        $script:AllStripLog | Export-Csv -LiteralPath $allStripFile -NoTypeInformation -Encoding UTF8
        Write-Log "Every stripped override/category/folder item/reference, with the reason: $allStripFile" -Level WARN
    }
    catch { Write-Log "Failed to save StrippedElements.csv: $($_.Exception.Message)" -Level WARN }
}

# Static group membership report (only if any static groups were found).
if ($script:StaticGroupReport.Count -gt 0) {
    $staticFile = Join-Path $OutputFolder "StaticGroupMembership.csv"
    try {
        $script:StaticGroupReport | Export-Csv -Path $staticFile -NoTypeInformation -Encoding UTF8
        $distinctGroups = @($script:StaticGroupReport | Select-Object -ExpandProperty GroupID -Unique).Count
        $resolvedNames = @($script:StaticGroupReport | Where-Object { $_.ResolvedName }).Count
        Write-Log "Static group membership report saved: $staticFile -- $($script:StaticGroupReport.Count) member rows across $distinctGroups groups, $resolvedNames resolved to names." -Level WARN
        if ($resolvedNames -eq 0 -and -not $ResolveStaticGroupMembers) {
            Write-Log "  Tip: re-run with -ResolveStaticGroupMembers and -SourceManagementServer to turn those member GUIDs into actual server names." -Level WARN
        }
    }
    catch {
        Write-Log "Failed to save static group membership report: $($_.Exception.Message)" -Level WARN
    }
}

Write-Log ""
Write-Log "Batch Import Order Summary"
Write-Log "----------------------------"
($ManifestRows | Format-Table Order, MPID, Verdict, Sealed, StrippedOverrides, StrippedReferences, CyclicFlag -AutoSize | Out-String -Width 250).Trim() |
    ForEach-Object { Write-Log $_ -NoConsole; Write-Host $_ }

###########################################################
# MISSING DEPENDENCIES REPORT (consolidated across the whole batch)
###########################################################
# Every reference across every MP in the batch that resolved to neither
# another batch MP nor anything already in the target repository, with a
# concrete suggestion of where it normally comes from and a direct
# systemcenter.wiki lookup link keyed off the MP ID. This does NOT download
# anything automatically -- there is no single reliable, scriptable source
# for arbitrary Microsoft/third-party MPs -- but it replaces "go figure out
# what's missing" with a ready list of exact IDs and where to look.

$MissingRollup = @($script:MissingDependencies |
    Group-Object MissingID |
    ForEach-Object {
        $first = $_.Group[0]
        [PSCustomObject]@{
            MissingID       = $first.MissingID
            Family          = $first.Family
            RequestedVersion= $first.RequestedVersion
            RequiredByCount = $_.Count
            RequiredBy      = (($_.Group | Select-Object -ExpandProperty RequiredBy) -join '; ')
            SuggestedSource = $first.SuggestedSource
            LookupUrl       = $first.LookupUrl
        }
    } |
    Sort-Object Family, MissingID)

$missingDepsFile = Join-Path $OutputFolder "MissingDependencies.csv"

if ($MissingRollup.Count -gt 0) {
    try {
        $MissingRollup | Export-Csv -Path $missingDepsFile -NoTypeInformation -Encoding UTF8
        Write-Log "Missing dependencies report saved: $missingDepsFile" -Level WARN
    }
    catch {
        Write-Log "Failed to save missing dependencies report: $($_.Exception.Message)" -Level WARN
    }

    $genuineMissing = @($MissingRollup | Where-Object { $_.SuggestedSource -notlike "EXCLUDED BY REQUEST*" })
    $excludedRollup = @($MissingRollup | Where-Object { $_.SuggestedSource -like "EXCLUDED BY REQUEST*" })

    if ($genuineMissing.Count -gt 0) {
        Write-Log ""
        Write-Log "MISSING DEPENDENCIES (not found in this batch or the target repository)" -Level WARN
        Write-Log "---------------------------------------------------------------------------"
        foreach ($m in $genuineMissing) {
            Write-Log "  [$($m.Family)] $($m.MissingID)  (wanted by $($m.RequiredByCount) MP(s): $($m.RequiredBy))" -Level WARN
            Write-Log "      Source : $($m.SuggestedSource)" -Level WARN -NoConsole
            Write-Log "      Lookup : $($m.LookupUrl)" -Level WARN -NoConsole
        }
        Write-Log ""
        Write-Log "Full source/lookup-link detail for each is in: $missingDepsFile" -Level WARN
    }

    if ($excludedRollup.Count -gt 0) {
        Write-Log ""
        Write-Log "EXCLUDED BY REQUEST (-ExcludeMPIDs -- known, intentional gaps, not searched for)" -Level WARN
        Write-Log "---------------------------------------------------------------------------"
        foreach ($m in $excludedRollup) {
            Write-Log "  [$($m.Family)] $($m.MissingID)  (wanted by $($m.RequiredByCount) MP(s): $($m.RequiredBy))" -Level WARN
        }
    }

    $missingCoreLibs = @($genuineMissing | Where-Object { $_.Family -eq "CoreLibrary" -or $_.Family -eq "IISCommonLibrary" -or $_.Family -eq "SystemCenter" })
    if ($missingCoreLibs.Count -gt 0) {

        # Distinguish two genuinely different situations that both end up
        # in MissingDependencies.csv, so the guidance given actually fits
        # what's wrong:
        #   1) GENUINELY ABSENT -- the family/ID isn't in $Repository at
        #      all. A re-export of -RepositoryFolder can fix this if the
        #      scan was incomplete.
        #   2) VERSION MISMATCH -- the ID IS in $Repository, just at a
        #      different (usually much newer) version than this old MP
        #      needs. Re-exporting -RepositoryFolder changes NOTHING here --
        #      the requested old version simply doesn't exist in any modern
        #      SCOM install, because the MP requesting it was abandoned
        #      rather than updated to use current core libraries.
        $genuinelyAbsent  = @($missingCoreLibs | Where-Object { -not $Repository.ContainsKey($_.MissingID) })
        $versionMismatch  = @($missingCoreLibs | Where-Object { $Repository.ContainsKey($_.MissingID) })

        if ($genuinelyAbsent.Count -gt 0) {
            Write-Log ""
            Write-Log "*** LIKELY ROOT CAUSE: INCOMPLETE REPOSITORY SCAN ***" -Level ERROR
            Write-Log "$($genuinelyAbsent.Count) item(s) ($($genuinelyAbsent.MissingID -join ', ')) are core/foundational MPs that ship with every SCOM installation, and are NOT present anywhere in -RepositoryFolder at all." -Level ERROR
            Write-Log "This usually means -RepositoryFolder did not scan a complete repository. Fix: from a live connection to SCOM $TargetVersion, run:" -Level ERROR
            Write-Log "    Get-SCOMManagementPack | ForEach-Object { `$_ | Export-SCOMManagementPack -Path 'C:\Temp\Target_AllMPs' }" -Level ERROR
            Write-Log "Then re-run this script with -RepositoryFolder 'C:\Temp\Target_AllMPs' instead." -Level ERROR
        }

        if ($versionMismatch.Count -gt 0) {
            Write-Log ""
            Write-Log "*** DIFFERENT ISSUE: OLD MP REQUIRES A VERSION THAT NO LONGER EXISTS ***" -Level ERROR
            Write-Log "$($versionMismatch.Count) item(s) ($($versionMismatch.MissingID -join ', ')) ARE present in -RepositoryFolder -- just at a different version than these old MP(s) require." -Level ERROR
            Write-Log "Re-exporting/re-scanning -RepositoryFolder will NOT fix this: the old version these MPs need genuinely does not exist in any modern SCOM install, because the requesting MP(s) were superseded rather than updated to reference current core libraries (this is a documented, known SCOM pattern, not specific to your environment)." -Level ERROR
            Write-Log "This usually means the source MP(s) requesting it are too old to be carried forward as-is. Options: (1) find the genuinely matching old version in your SCOM $SourceVersion source export and attempt to import it alongside the modern version (SCOM's strict version model may reject this), (2) keep the systems that need this monitoring on the old SCOM environment, or (3) find a current/different MP or tool to monitor that workload going forward." -Level ERROR
        }
    }
}
else {
    Write-Log ""
    Write-Log "No missing dependencies detected -- every reference resolved within the batch or the target repository." -Level SUCCESS
}

###########################################################
# DEPRECATION NOTICES ROLLUP (consolidated across the whole batch)
###########################################################

$DeprecationRollup = @()
foreach ($result in $BatchResults.Values) {
    foreach ($advice in $result.MigrationAdvice) {
        if ($advice.Deprecation) {
            $DeprecationRollup += [PSCustomObject]@{
                MPID        = $result.MPID
                Reference   = $advice.Reference
                Notice      = $advice.Deprecation
            }
        }
    }
}

if ($DeprecationRollup.Count -gt 0) {
    Write-Log ""
    Write-Log "DEPRECATION NOTICES (still functional today, but Microsoft has announced an end date)" -Level WARN
    Write-Log "---------------------------------------------------------------------------------------"
    foreach ($d in ($DeprecationRollup | Sort-Object Reference -Unique)) {
        Write-Log "  - $($d.Reference): $($d.Notice)" -Level WARN
    }
}

###########################################################
# GENERATE IMPORT SCRIPT (3.40: resilient, ordered, reports every error)
###########################################################
# The generated script reads BatchImportManifest.csv (so you can edit a
# Verdict or delete a row before running it), imports READY rows in
# dependency order, keeps going past failures, skips anything whose batch
# dependency failed, and writes ImportResults_<timestamp>.csv plus
# ImportErrors_<timestamp>.txt with SCOM's full exception chain for every
# failure. Those two files are what to send back for triage.

$importScriptFile = Join-Path $OutputFolder "Import-Batch.$TargetVersion.ps1"

$importTemplate = @'
<#
    Generated by SCOM MP Batch Compiler __BUILD__ on __DATE__
    Source batch  : __SOURCE__
    Target repo   : __REPO__

    Imports the rows of BatchImportManifest.csv (same folder) whose Verdict is
    READY, in the Order column (dependencies first). Keeps going after a
    failure; anything that depends on a failed MP is skipped, not attempted.

    Examples:
      .\Import-Batch.__TARGET__.ps1 -ManagementServer SCOMMS01 -WhatIfOnly   # dry run, no changes
      .\Import-Batch.__TARGET__.ps1 -ManagementServer SCOMMS01               # import READY rows
      .\Import-Batch.__TARGET__.ps1 -ManagementServer SCOMMS01 -Only "Contoso*" # subset
      .\Import-Batch.__TARGET__.ps1 -ManagementServer SCOMMS01 -IncludeBlocked  # also TRY blocked rows, to collect SCOM's exact errors

    Send back: ImportResults_<timestamp>.csv and ImportErrors_<timestamp>.txt
#>
[CmdletBinding()]
param(
    [string]$ManagementServer = "localhost",
    [switch]$WhatIfOnly,
    [switch]$IncludeBlocked,
    [switch]$Reimport,
    [string]$Only,
    # Import only if the connected management group has this name. If not
    # given, you are shown the name and asked to confirm (unless -Force).
    [string]$ExpectedManagementGroup = '__EXPECTED_MG__',
    [switch]$Force
)

Set-StrictMode -Off
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$manifestPath = Join-Path $here 'BatchImportManifest.csv'
if (-not (Test-Path -LiteralPath $manifestPath)) { throw "BatchImportManifest.csv not found next to this script ($here)." }
$entries = @(Import-Csv -LiteralPath $manifestPath | Sort-Object { [int]$_.Order })

Import-Module OperationsManager
New-SCOMManagementGroupConnection -ComputerName $ManagementServer | Out-Null
$target = if ($ManagementServer -eq 'localhost') { $env:COMPUTERNAME } else { $ManagementServer }
$short = ($target -split '\.')[0]
$conn = @(Get-SCOMManagementGroupConnection) | Where-Object { (([string]$_.ManagementServerName) -split '\.')[0] -eq $short } | Select-Object -First 1
if ($conn) { $conn | Set-SCOMManagementGroupConnection }
$active = @(Get-SCOMManagementGroupConnection | Where-Object { $_.IsActive }) | Select-Object -First 1
if (-not $active) { throw "No active SCOM connection after connecting to '$ManagementServer'." }
Write-Host "Target management group: $($active.ManagementGroupName) via $($active.ManagementServerName)" -ForegroundColor Cyan
if ($ExpectedManagementGroup -and $ExpectedManagementGroup -ne ('__EXPECTED' + '_MG__')) {
    if ($active.ManagementGroupName -ne $ExpectedManagementGroup) { throw "Connected to management group '$($active.ManagementGroupName)', expected '$ExpectedManagementGroup'. Nothing imported." }
}
elseif (-not $WhatIfOnly -and -not $Force) {
    $ans = Read-Host "Import into management group '$($active.ManagementGroupName)' on '$($active.ManagementServerName)'? This must be the NEW SCOM environment. Type YES to continue"
    if ($ans -ne 'YES') { throw "Not confirmed. Nothing imported." }
}
if ($WhatIfOnly) { Write-Host "DRY RUN -- nothing will be imported." -ForegroundColor Yellow }

$installed = @{}
foreach ($m in @(Get-SCOMManagementPack)) { $installed[[string]$m.Name] = $m }
Write-Host "Installed MPs in target: $($installed.Count)"

$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$resultsPath = Join-Path $here "ImportResults_$stamp.csv"
$errorsPath  = Join-Path $here "ImportErrors_$stamp.txt"
$results = New-Object System.Collections.Generic.List[object]
$status = @{}
$importedFiles = @{}   # a .mpb bundle holds several MPs but is imported once
$okStates = @('Imported', 'AlreadyPresent', 'WouldImport')

function Resolve-ImportFile([string]$p) {
    if ($p -and (Test-Path -LiteralPath $p)) { return $p }
    if ($p) {
        $leaf = Split-Path -Leaf $p
        foreach ($c in @((Join-Path $here "CandidateMPs\$leaf"), (Join-Path $here $leaf))) { if (Test-Path -LiteralPath $c) { return $c } }
    }
    return $null
}

$i = 0
foreach ($e in $entries) {
    $i++
    if ($Only -and $e.MPID -notlike $Only) { continue }
    $t0 = Get-Date
    $result = $null; $detail = ''
    $deps = @(([string]$e.BatchDependsOn) -split ';\s*' | Where-Object { $_ })
    $badDeps = @($deps | Where-Object { $status.ContainsKey($_) -and $okStates -notcontains $status[$_] })
    $file = Resolve-ImportFile $e.ImportPath

    if ($e.Verdict -ne 'READY' -and -not $IncludeBlocked) {
        $result = 'SkippedBlocked'; $detail = [string]$e.BlockingReasons
    }
    elseif ($badDeps.Count -gt 0) {
        $result = 'SkippedDependencyFailed'; $detail = "Depends on: $($badDeps -join ', ')"
    }
    elseif (-not $file) {
        $result = 'Failed'; $detail = if ($e.ImportPath) { "Import file not found: $($e.ImportPath)" } else { "No importable file (sealed MP without its original .mp/.mpb)" }
    }
    else {
        $inst = $installed[[string]$e.MPID]
        $skipPresent = $false
        if ($inst) {
            $cmp = 0
            try { $cmp = ([version][string]$inst.Version).CompareTo([version][string]$e.SourceVersion) } catch { $cmp = if ([string]$inst.Version -eq [string]$e.SourceVersion) { 0 } else { -1 } }
            if ($cmp -gt 0) { $skipPresent = $true; $detail = "Newer version already installed (v$($inst.Version))" }
            elseif ($cmp -eq 0 -and ($inst.Sealed -or -not $Reimport)) { $skipPresent = $true; $detail = "Already installed at v$($inst.Version)" }
        }
        if (-not $skipPresent -and $importedFiles.ContainsKey($file)) { $skipPresent = $true; $detail = "Imported with bundle $(Split-Path -Leaf $file) earlier in this run" }
        if ($skipPresent) { $result = 'AlreadyPresent' }
        elseif ($WhatIfOnly) { $result = 'WouldImport'; $detail = $file; $importedFiles[$file] = $true }
        else {
            try {
                Import-SCOMManagementPack -Fullname $file -ErrorAction Stop
                $result = 'Imported'
                $importedFiles[$file] = $true
                try { $installed[[string]$e.MPID] = Get-SCOMManagementPack -Name $e.MPID -ErrorAction Stop } catch { }
            }
            catch {
                $result = 'Failed'
                $chain = New-Object System.Collections.Generic.List[string]
                $ex = $_.Exception; $d = 0
                while ($ex -and $d -lt 8) { $chain.Add("[$d] $($ex.GetType().Name): $(([string]$ex.Message).Trim())"); $ex = $ex.InnerException; $d++ }
                $detail = ($chain -join ' || ')
                Add-Content -LiteralPath $errorsPath -Encoding UTF8 -Value (@("==== $($e.MPID) (v$($e.SourceVersion))  $file") + @($chain) + @(""))
            }
        }
    }

    $status[[string]$e.MPID] = $result
    $secs = [math]::Round(((Get-Date) - $t0).TotalSeconds, 1)
    $results.Add([PSCustomObject]@{
        Order = $e.Order; MPID = $e.MPID; WorkbookName = $e.WorkbookName; Verdict = $e.Verdict; Result = $result
        Detail = $detail; Sealed = $e.Sealed; SourceVersion = $e.SourceVersion; ImportFile = $file; Seconds = $secs; Time = (Get-Date -Format 's')
    })
    $color = switch ($result) { 'Imported' { 'Green' } 'WouldImport' { 'Green' } 'AlreadyPresent' { 'DarkGreen' } 'Failed' { 'Red' } default { 'Yellow' } }
    Write-Host ("[{0}/{1}] {2,-24} {3}" -f $i, $entries.Count, $result, $e.MPID) -ForegroundColor $color
    if ($result -eq 'Failed') { Write-Host "      $detail" -ForegroundColor Red }
    $results | Export-Csv -LiteralPath $resultsPath -NoTypeInformation -Encoding UTF8   # rewritten each time so a crash still leaves a record
}

Write-Host ""
$results | Group-Object Result | Sort-Object Name | ForEach-Object { Write-Host ("{0,-26} {1}" -f $_.Name, $_.Count) }
Write-Host ""
Write-Host "Results: $resultsPath" -ForegroundColor Cyan
if (Test-Path -LiteralPath $errorsPath) { Write-Host "Errors : $errorsPath" -ForegroundColor Red }
'@

$importText = $importTemplate.Replace('__BUILD__', $script:BuildLabel)
$importText = $importText.Replace('__DATE__', (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
$importText = $importText.Replace('__SOURCE__', ($InputPath -join ', '))
$importText = $importText.Replace('__REPO__', $RepositoryFolder)
$importText = $importText.Replace('__TARGET__', $TargetVersion)
$expectedMg = if ($LiveImport -and $script:TargetConnection) { [string]$script:TargetConnection.ManagementGroupName } else { '' }
$importText = $importText.Replace('__EXPECTED_MG__', $expectedMg)

try {
    Set-Content -LiteralPath $importScriptFile -Value $importText -Encoding UTF8
    Write-Log "Generated import script: $importScriptFile" -Level SUCCESS
}
catch {
    Write-Log "Failed to write generated import script: $($_.Exception.Message)" -Level WARN
}

###########################################################
# STEP 10 - LIVE IMPORT OF TARGETS (-LiveImport only)
###########################################################
# Runs the exact script generated above, in-process, so a -LiveImport run and
# a later manual run behave identically: dependency order, cascade-skip on
# failure, and ImportResults / ImportErrors files for triage.
$script:LiveTargetResultsPath = $null
if ($LiveImport) {
    Write-Banner "STEP 10 - LIVE IMPORT OF READY TARGETS (dependency order)"
    try {
        & $importScriptFile -ManagementServer $ManagementServer -Force
        $latest = Get-ChildItem -LiteralPath $OutputFolder -Filter 'ImportResults_*.csv' -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($latest) { $script:LiveTargetResultsPath = $latest.FullName }
        Use-TargetConnection
    }
    catch {
        Write-Log "Live import of targets stopped: $($_.Exception.Message)" -Level ERROR
    }
}

###########################################################
# FINAL SUMMARY
###########################################################

$scored = @($BatchResults.Values | Where-Object { $_.Status -eq 'PROCESSED' })
$readyRows   = @($ManifestRows | Where-Object { $_.Verdict -eq 'READY' })
$blockedRows = @($ManifestRows | Where-Object { $_.Verdict -ne 'READY' })
$sealedReady = @($readyRows | Where-Object { $_.Sealed })
$totalStrippedOv  = ($ManifestRows | Measure-Object -Property StrippedOverrides -Sum).Sum
$totalStrippedRef = ($ManifestRows | Measure-Object -Property StrippedReferences -Sum).Sum
$staticRows = @($ManifestRows | Where-Object { [int]$_.StaticGroups -gt 0 })

Write-Banner "MIGRATION BATCH RUN COMPLETE"
Write-Log "Source batch          : $($InputPath -join ', ')"
Write-Log "Target Repository     : $RepositoryFolder ($TargetVersion)"
$sourceRepoFinalText = if ($SourceRepositoryFolder) { $SourceRepositoryFolder } else { '(not provided)' }
if ($script:SourceRepositoryIsLive) { $sourceRepoFinalText += " [LIVE export from '$SourceManagementServer']" }
Write-Log "Source Repository     : $sourceRepoFinalText"
Write-Log "Manifest              : $(if ($Manifest) { $Manifest } else { '(none -- every MP under -InputPath)' })"
Write-Log "MPs in original batch : $OriginalBatchSize"
Write-Log "MPs auto-resolved     : $totalPulled"
Write-Log "MPs processed         : $($scored.Count)"
Write-Log "Cyclic / unorderable  : $($CyclicNodes.Count)"
Write-Log "Distinct Missing Deps : $($MissingRollup.Count)"
Write-Log ""
Write-Log "Batch Manifest          : $manifestFile"
if ($MissingRollup.Count -gt 0) { Write-Log "Missing Dependencies    : $missingDepsFile" -Level WARN }
Write-Log "Generated Import Script : $importScriptFile"
Write-Log "Candidate MP Folder     : $CandidateFolder"
Write-Log "Log File                : $script:LogFile"

Write-Log ""
Write-Log "==========================================================" -Level WARN
Write-Log "BOTTOM LINE -- READ THIS FIRST" -Level WARN
Write-Log "==========================================================" -Level WARN
Write-Log "READY to import   : $($readyRows.Count)  ($($sealedReady.Count) of them sealed originals, imported as-is)" -Level SUCCESS
Write-Log "BLOCKED           : $($blockedRows.Count)" -Level $(if ($blockedRows.Count -gt 0) { 'ERROR' } else { 'SUCCESS' })
Write-Log "Dead overrides stripped across the batch : $totalStrippedOv (targets no longer exist in $TargetVersion -- see StrippedElements.csv)"
Write-Log "Dead references stripped                 : $totalStrippedRef"
if ($staticRows.Count -gt 0) {
    Write-Log "Static-membership groups still left in $($staticRows.Count) MP(s) will import EMPTY (old-environment object GUIDs). $(if ($GroupConversionFile) { 'See GroupConversionResults.csv (LeftStatic rows).' } else { 'Run Export-ScomEnvironment.ps1 -Role Source and pass -GroupConversionFile to convert them to dynamic groups.' })" -Level WARN
}

if ($blockedRows.Count -gt 0) {
    Write-Log ""
    Write-Log "BLOCKED MPs and why:" -Level ERROR
    foreach ($br in $blockedRows) {
        Write-Log "  $($br.MPID)" -Level ERROR
        foreach ($reason in (([string]$br.BlockingReasons) -split ' \| ')) { if ($reason) { Write-Log "      - $reason" -Level ERROR } }
    }
}

if ($CyclicNodes.Count -gt 0) {
    Write-Log ""
    Write-Log "Circular references among: $($CyclicNodes -join ', ') -- these were appended at the end of the order; import them by hand after reviewing." -Level WARN
}

Write-Log ""
Write-Log "NEXT ACTION:" -Level WARN
if ($LiveImport) {
    Write-Log "  Live import ran. Results: $(if ($script:LiveTargetResultsPath) { $script:LiveTargetResultsPath } else { '(see ImportResults_*.csv in the output folder)' })" -Level WARN
}
else {
    Write-Log "  1) Dry run on the SCOM $TargetVersion management server:  .\Import-Batch.$TargetVersion.ps1 -ManagementServer <MS> -WhatIfOnly" -Level WARN
    Write-Log "  2) Real run:                                             .\Import-Batch.$TargetVersion.ps1 -ManagementServer <MS>" -Level WARN
    Write-Log "  3) Send back ImportResults_*.csv, ImportErrors_*.txt, BatchImportManifest.csv, MissingDependencies.csv and this log." -Level WARN
}
Write-Log "==========================================================" -Level WARN
Write-Log ""
