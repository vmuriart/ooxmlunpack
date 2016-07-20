# THIS SCRIPT UPDATES THE FOLLOWING:
# Reads the major / minor version from AssemblyInfo.cs
# Reads the branch name from Mercurial
# Reads the revision number from Mercurial
# Reads the build number from Teamcity
# Sets the Teamcity build number to major.minor.revision.build
# Sets the version for any published NuGet packages to major.minor.revision-branch

$original = Get-Content 'Properties/AssemblyInfo.cs' | Out-String

# Get version from template.
[void] ($original -match '(?<=\r\n\[assembly: AssemblyVersion\(")(?<version>\d+\.\d+)(?=(\.\d+){0,2}"\)\]\r\n)')
$version = $matches['version']

# Get branch name.
$branch = ((hg id -b) | Out-String) -replace "[ \t\r\n]", ""
# if ($branch -match '^(?<version>\d+\.\d+)$') {
  # # Branch looks like a version, use it as such.
  # $version = $branch
# }

# Get revision number.
if (((hg id -n) | Out-String) -replace "[ \t\r\n\+]", "" -match '^(?<revision>\d+)$') {
  $revision = $matches['revision']
  Write-Host "Using Hg revision number: $revision"
} else {
  $revision = 0
  Write-Host "No revision number found so using default: 0"
}

# Get build number.
if (test-path env:BUILD_NUMBER) {
  $build = $env:BUILD_NUMBER
  Write-Host "Using TeamCity build number: $build"
} else {
  $build = 0
  Write-Host "No build number found so using default: 0"
}

# This line is a command to TeamCity to update the build number
# (see https://confluence.jetbrains.com/display/TCD8/Build+Script+Interaction+with+TeamCity )
Write-Host "##teamcity[buildNumber '$version.$revision.$build']"

# This line is a command to TeamCity to update the named build parameter
if($branch -eq 'default') {
  Write-Host "##teamcity[setParameter name='NuGetPackageVersion' value='$version.$revision']"
} else {
  Write-Host "##teamcity[setParameter name='NuGetPackageVersion' value='$version.$revision-$branch']"
}
