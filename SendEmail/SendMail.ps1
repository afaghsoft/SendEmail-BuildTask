param
(
    [bool] $isDebug = $false
)

$global:PATTERN_ZIP_FILE_OUTPUT = "(\{\{\s*ZIP_FILE_OUTPUT\s*\}\})" # e.g. {{ Git:info }}

# Splits the given string value with the given separators and
# removes the empty entries.
function Split-StringValue {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline)]
        [string]$InputObject,
        [string[]]$Separators = " "
    )
    
    
    process {
        return $InputObject.Split($Separators, 
            [System.StringSplitOptions]::RemoveEmptyEntries) |
        ForEach-Object {
            if ([string]::IsNullOrWhiteSpace($_)) {
                return $null
            }

            return $_.Trim()
        } | Where-Object { $null -ne $_ }
    }
}

# fetches and returns the current git branch that we are on.
# If the current directory that we are in at the moment, doesn't belong to
# any git repository (AKA: does not have .git dir and its content), this
# function will write the error using Write-Error cmdlet and returns $null.
function Get-CurrentGitBranch {
    [CmdletBinding()]
    param (
    )
    
    process {
        # provided by TFS / Azure DevOps server itself, so we prioritize it.
        if ($Env:BUILD_SOURCEBRANCHNAME -and $Env:BUILD_SOURCEBRANCHNAME.Trim()) {
            return $Env:BUILD_SOURCEBRANCHNAME
        }

        $theBranch = (git branch 2>&1) -split "\n" | Where-Object { $_.StartsWith("*") }
        if ($null -eq $theBranch) {
            # putting this here just in case of exceptional situations (like not being in a dir with .git dir, etc).
            # we might want to change the behavior of it in future.
            return $null
        }
        elseif ($gitOutput -is [System.Management.Automation.ErrorRecord]) {
            $gitOutput | Write-Error
            return $null
        }

        $theBranch = ($theBranch -as [string]).Substring(2, $theBranch.Length - 2)

        # are we in detached state?
        # if yes, then simply invoking "git branch" won't make us reach the
        # correct answer. 
        if ($theBranch.StartsWith("(HEAD detached at")) {
            $theBranch = (git show -s --pretty=%d HEAD 2>&1)
            # The expected output for this command is something like this:
            #  (HEAD, tag: v1.2.10, origin/Release1.2)
            # we better not rely on the "tag: ..." part (since the branch might have no tag)
            # and just find the branch ourselves.
            $myStrs = $theBranch | Split-StringValue -Separators @(", ") 
            for ($i = 1; $i -lt $myStrs.Count; $i++) {
                if (($myStrs[$i] -as [string]).Contains("/")) {
                    # the last index will be the correct branch name.
                    $theBranch = ($myStrs[$i]).Split('/')[-1].Trim(")")
                    break
                }
            }
        }

        return $theBranch
    }
}

# returns file path(s) for the output zip file. If it exists.
function Get-ZipFileOutputPath {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [string]$OutputPathContainer = $null
    )
    
    process {
        if (!$OutputPathContainer -or !(Test-Path $OutputPathContainer)) {
            return "Unknown Path"
        }

        $fileContent = Get-Content -Path $OutputPathContainer
        return $fileContent.Split(";")[0].Trim()
    }
}

function Send-MailFromPipeline {
    $To = Get-VstsInput -Name 'To' -Require
    $Subject = Get-VstsInput -Name 'Subject' -Require
    $Body = Get-VstsInput -Name 'Body' -Require
    $From = Get-VstsInput -Name 'From' -Require
    $BodyAsHtml = Get-VstsInput -Name 'BodyAsHtml'
    $SmtpServer = Get-VstsInput -Name 'SmtpServer' -Require
    $SmtpPort = Get-VstsInput -Name 'SmtpPort'
    $SmtpUsername = Get-VstsInput -Name 'SmtpUsername'
    $SmtpPassword = Get-VstsInput -Name 'SmtpPassword'
    $UseSSL = Get-VstsInput -Name 'UseSSL'
    $AddAttachment = Get-VstsInput -Name 'AddAttachment'
    $Attachment = Get-VstsInput -Name 'Attachment'
    $CC = Get-VstsInput -Name 'CC'
    $BCC = Get-VstsInput -Name 'BCC'


    if (!$SmtpServer -or $SmtpServer -eq "none" -or !$From -or $From -eq "none") {
        Write-Output "Server-address/From is set to 'none'. Doing nothing!"
        return
    }

    # Additional parameters
    $BranchFilter = Get-VstsInput -Name 'BranchFilter'
    $ZipFilePathContainer = Get-VstsInput -Name 'ZipFilePathContainer'

    $Body = [Regex]::Replace($Body, $global:PATTERN_ZIP_FILE_OUTPUT, { 
        return Get-ZipFileOutputPath -OutputPathContainer $ZipFilePathContainer
    })

    $MailParams = @{}

    Write-Output "Input Vars"
    Write-Output "Branch Filter: $BranchFilter"
    Write-Output "Send Email To: $To"
    Write-Output "Send Email CC: $CC"
    Write-Output "Send Email BCC: $BCC"
    Write-Output "Subject: $Subject"
    Write-Output "Send Email From: $From"
    Write-Output "Body as Html?: $BodyAsHtml"
    Write-Output "SMTP Server: $SmtpServer"
    Write-Output "SMTP Username: $SmtpUsername"
    Write-Output "SMTP Port: $SmtpPort"
    Write-Output "Use SSL?: $UseSSL"
    Write-Output "Add Attachment?: $AddAttachment"
    Write-Output "Attachment: $Attachment"

    if ($BranchFilter -and !($BranchFilter -eq "*" -or $BranchFilter -eq "**")) {
        $currentGitBranch = Get-CurrentGitBranch
        if (!$currentGitBranch) {
            Write-Output "Couldn't determine current git branch!"
            return
        }

        $BranchFilters = $BranchFilter.Split(";").Trim()
        $matchedOnce = $false
        foreach ($currentFilter in $BranchFilters) {
            if ($currentGitBranch -like $currentFilter) {
                $matchedOnce = $true
                break
            }
        }

        if (!$matchedOnce) {
            Write-Output "Couldn't match current git branch '$currentGitBranch' with any filters applied!"
            return
        }
    }

    [string[]]$toMailAddresses = $To.Split(';');
    [string[]]$ccMailAddresses = $CC.Split(';');
    [string[]]$bccMailAddresses = $BCC.Split(';');

    [bool]$BodyAsHtmlBool = [System.Convert]::ToBoolean($BodyAsHtml)
    [bool]$UseSSLBool = [System.Convert]::ToBoolean($UseSSL)
    [bool]$AddAttachmentBool = [System.Convert]::ToBoolean($AddAttachment)

    $MailParams.Add("To", $toMailAddresses)
    if ($null -ne $ccMailAddresses -and $ccMailAddresses[0] -ne "") { 
        $MailParams.Add("Cc", $ccMailAddresses)
    }

    if ($null -ne $bccMailAddresses -and $bccMailAddresses[0] -ne "") { 
        $MailParams.Add("Bcc", $bccMailAddresses)
    }
    $MailParams.Add("From", $From)

    $SubjectExpanded = $ExecutionContext.InvokeCommand.ExpandString($Subject) 
    $MailParams.Add("Subject", $SubjectExpanded)

    $BodyExpanded = $ExecutionContext.InvokeCommand.ExpandString($Body) 
    $MailParams.Add("Body", $BodyExpanded)

    $MailParams.Add("SmtpServer", $SmtpServer)
    $MailParams.Add("Port", $SmtpPort)
    $MailParams.Add("Encoding", "UTF8")

    if (!([string]::IsNullOrEmpty($SmtpUsername))) {
        $securePassword = ConvertTo-SecureString $SmtpPassword -AsPlainText -Force
        $credential = New-Object System.Management.Automation.PSCredential ($SmtpUsername, $securePassword)
        $MailParams.Add("Credential", $credential)
    }

    if ($BodyAsHtmlBool) {
        $MailParams.Add("BodyAsHtml", $true)
    }

    if ($UseSSLBool) {
        $MailParams.Add("UseSSL", $true)
    }

    if ($AddAttachmentBool) {
        $MailParams.Add("Attachments", $Attachment)
    }

    Send-MailMessage @MailParams
}


if ($isDebug -eq $false) {
    try {
        Send-MailFromPipeline
    }
    catch {
        throw
    }
}

