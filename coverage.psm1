function Get-CodeCoverage{
    param(
        # Path to a test project file (*.csproj)
        [Parameter(Mandatory=$true)]
        [string] $TestProjectPath,
        # Path to a Test settings file (*.runsettings)
        [Parameter(Mandatory=$false)]
        [string] $TestSettingsPath,
        # Open coverage report when it is generated
        [Parameter(Mandatory=$false)]
        [bool] $OpenReport
    )

    try {

        if (-not (Test-Path $TestProjectPath))
        {
            throw [System.IO.FileNotFoundException] "$TestProjectPath not found."
        }

        $testResultDirectoryName = 'TestResults'
        $coverageXmlFilename = 'output.coveragexml'
        $coverageReportDirectory = 'coveragereport'

        # ensuring absolute path to the project is always used
        Write-Debug -Message "Test project path $testProjectPath"
        $absoluteTestProjectPath = Resolve-Path $TestProjectPath
        Write-Debug -Message "Absolute path to test project: $absoluteTestProjectPath"

        $testProjectDirectory = Split-Path $absoluteTestProjectPath
        Write-Debug -Message "Test project directory: $testProjectDirectory"

        if ($null -eq $TestSettingsPath -or $TestSettingsPath -eq "")
        {
            # if test settings file path is not provided, try to auto-guess it
            $testSettingsFileName = Get-ChildItem -File -Filter *.runsettings -Path $testProjectDirectory -Name | Select-Object -First 1;
            Write-Debug -Message "Test settings file $testSettingsFileName"
            $testSettingsPath = Join-Path -Path $testProjectDirectory -ChildPath $testSettingsFileName
            Write-Debug -Message "Test settings full path $testSettingsPath"
        }

        $testSettingsPresent = $true
        if (-not (Test-Path $TestSettingsPath))
        {
            $testSettingsPresent = $false
            Write-Error -Message "$TestSettingsPath not found."
            Write-Error -Message "Coverage report will be generated, but it might include thrid-party dependencies"
        }
        if ($testSettingsPresent)
        {
            # ensure absolute path is used
            Write-Debug -Message "Test settings path $TestSettingsPath"
            $absoluteTestSettingsPath = Resolve-Path $TestSettingsPath
            Write-Debug -Message "Absolute path to test settings file: $absoluteTestSettingsPath"
        }
        # test results directory
        $testResultDirectory = Join-Path -Path $testProjectDirectory -ChildPath $testResultDirectoryName
        Write-Debug -Message "TestResult directory: $testResultDirectory"

        # check if required packages are installed and added as project references
        $package = 'microsoft.codecoverage'
        $coverageInstalled = Assert-Package -PackageName $package
        $coverageDirectory = Get-LatestPackageDirectory -PackageName $package

        $package = 'reportgenerator'
        $reportGeneratorInstalled = Assert-Package -PackageName $package
        $reportGeneratorDirectory = Get-LatestPackageDirectory -PackageName $package

        # run tests with rebuild if packages were not installed before running this script
        # permit running tests without settings file
        $testRunnerCommand = "dotnet test " + $absoluteTestProjectPath + " -v q"
        if ($testSettingsPresent)
        {
            $testRunnerCommand += " --settings:" + $absoluteTestSettingsPath
        }
        if (-not $coverageInstalled -and -not $reportGeneratorInstalled) {
            $testRunnerCommand += " --no-build"
        }
        Write-Debug -Message "Running tests via $testRunnerCommand"
        Invoke-Expression $testRunnerCommand

        # get the most recent coverage file
        $recentCoverageFile = Get-ChildItem -File -Filter *.coverage -Path $testResultDirectory -Name -Recurse | Select-Object -First 1;
        $recentCoverage = Join-Path -Path $testResultDirectory -ChildPath $recentCoverageFile
        Write-Host "Test Completed, coverage file : $recentCoverage"  -ForegroundColor Green

        # convert coverage file from binary to xml
        $coverageFile = Join-Path -Path $testResultDirectory -ChildPath $coverageXmlFilename
        $coverageCall = $coverageDirectory + '\build\netstandard1.0\CodeCoverage\CodeCoverage.exe analyze  /output:' + $coverageFile + ' ' + $recentCoverage
        Write-Debug -Message "Converting coverage results via: $coverageCall"
        Invoke-Expression $coverageCall
        Write-Host 'CoverageXML Generated in $coverageFile'  -ForegroundColor Green

        if (Test-Path $coverageFile)
        {
            # if binary-to-xml conversion succeeded, prepare an html report
            $outputDirectory = Join-Path -Path $testResultDirectory -ChildPath $coverageReportDirectory
            $reportCall = 'dotnet ' + $reportGeneratorDirectory + '\tools\netcoreapp3.0\ReportGenerator.dll "-reports:' + $coverageFile + '" "-targetdir:' + $outputDirectory + '"'
            Write-Debug -Message "Generating coverage report via: $reportCall"
            Invoke-Expression $reportCall
            Write-Host 'CoverageReport Published'  -ForegroundColor Green

            if ($OpenReport)
            {
                # open browser with report if required
                $reportPath = Join-Path -Path $testResultDirectory -ChildPath $coverageReportDirectory | Join-Path -ChildPath 'index.htm'
                if (Test-Path $reportPath)
                {
                    Invoke-Expression $reportPath
                }
                else {
                    Write-Error "$reportPath was not generated"
                }
            }
        }
        else {
            Write-Error "$coverageFile was not generated"
        }
    }
    catch {

        Write-Host "Caught an exception:" -ForegroundColor Red
        Write-Host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
        Write-Host "Exception Message: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Exception Line: $($_.ScriptStackTrace)" -ForegroundColor Red
    }
    finally{
        [Console]::ResetColor()
    }
}

function Assert-Package{
    param (
        [Parameter(Mandatory=$true)]
        [string] $PackageName
    )

    Write-Debug -Message "Assert-Package $PackageName"
    $rebuildRequired = $false
    $packageInstalled = dotnet list $testProjectPath package | Select-String $PackageName -Quiet
    if (-not $packageInstalled)
    {
        $rebuildRequired = $true
        Write-Host $package 'is not installed, installing' -ForegroundColor Yellow
        dotnet add $testProjectPath package $package
    }
    Write-Debug -Message "Rebuild required: $rebuildRequired"
    return $rebuildRequired
}

function Get-LatestPackageDirectory{
    param(
        [Parameter(Mandatory=$true)]
        [string] $PackageName
    )

    Write-Debug -Message "Get-LatestPackageDirectory $PackageName"
    $packageDirectory = Join-Path -Path $env:USERPROFILE -ChildPath '.nuget' | Join-Path -ChildPath 'packages' | Join-Path -ChildPath $PackageName
    Write-Debug -Message "Package Directory: $packageDirectory"
    $latestPackageVersion = Get-ChildItem -Directory -Path $packageDirectory -Name | Select-Object -Last 1;
    Write-Debug -Message "Latest package version: $latestPackageVersion"
    return Join-Path -Path $packageDirectory -ChildPath $latestPackageVersion
}
