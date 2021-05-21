################################################
##  Copy Planner Plan  #########################
##  ############################################
##  Created by:                               ##
##  Heather Wilson  ############################
##  ############################################
##  Version 1.03                              ##
##  ############################################
##  Last Update Date                          ##
##  May 20, 2021                              ##
################################################
##  This script is based on the               ##
##  PlannerMigration.ps1 script created by    ##
##  Github user smbm.                         ##
##  I made significant changes to smbm's      ##
##  script, mainly because it was meant for   ##
##  copying Plans to Groups in different M365 ##
##  tenants. In my use case, Plans need to be ##
##  copied to Groups within the same M365     ##
##  tenant.                                   ##
##  Many of the original script's functions   ##
##  relied on array item comparisons in order ##
##  for the script to proceed, so those had   ##
##  to be modified or removed.                ##
##  This script requires the user to input    ##
##  the source Group and destination Group    ##
##  for copying all Plans. Beyond that, the   ##
##  script runs alone without any other input.##
##  Login method was changed from passing     ##
##  plaintext credentials to the device code  ##
##  login method. It's a little inconvenient  ##
##  but much more secure.                     ##
################################################
##  Other inspiration for creating this       ##
##  script was disappointment in Microsoft's  ##
##  own Plan copying function that does not   ##
##  return task comments, task completion     ##
##  status, checklist item completion status, ##
##  task assignments, and who completed the   ##
##  task. I think that it is unreasonable     ##
##  that those features were not added when   ##
##  the features are clearly accessible       ##
##  through the Microsoft Graph API.          ##
##                                            ##
##  Lastly, I believe in freely sharing this  ##
##  kind of time-saving information because   ##
##  you should not have to pay a vendor       ##
##  thousands of dollars for this. Or pay     ##
##  them a lot and not even get this Plan     ##
##  copying feature.                          ##
################################################
##  The original script was missing a lot of  ##
##  Plan data that could be copied to the new ##
##  destination plan. My script copies:       ##
##  1. Task start and due dates               ##
##  2. Task progress: not started,            ##
##     in progress, completed                 ##
##  3. Task notes                             ##
##  4. Task comments, who completed task, and ##
##     when task was completed was also added ##
##     to task notes section                  ##
##  5. If task is completed, the task preview ##
##     will be set as "description" so the    ##
##     task completion information is visible ##
##     on the task card. Task completion data ##
##     is always added to the top of the      ##
##     notes section, even if other data is   ##
##     present in the notes section           ##
##  6. Task assigned users                    ##
##  7. Task checklist items and completion    ##
##     status of checklist items              ##
##  8. Copy task label                        ##
################################################
##  FUNCTIONS STILL MISSING  ###################
##  1. Automatically add file attachments to  ##
##     tasks. Probably the best way to do     ##
##     this is to get the file names of tasks ##
##     with attached files, automate moving   ##
##     the files to one folder in the         ##
##     destination Group's SharePoint site,   ##
##     and then re-add the file to its        ##
##     corresponding task and Plan.           ##
##  2. Task urgency. I don't think there is a ##
##     graph API query for this.              ##
################################################
##  POWERSHELL MODULE REQUIREMENTS #############
##  TEAMS  #####################################
##  EXCHANGEONLINEV2  ##########################
################################################
$Header = @{
    Authorization = $onlyToken
}
filter displayname-filter{
    param ([string]$filterString)
    if ($_.displayname -ceq "$filterString"){$_}
}
##  This function returns each Plan in one group. It retrieves the Plan's title
##  and the Plan ID.
##  It writes the current Group name to the console.
function GetPlans($groups, $token){
    $Header = @{
        Authorization = $token
    }
    $plans = @()
    foreach ($group in $groups){
        Write-Host "`r`nQuerying plans in $($group.displayname) ..." -ForegroundColor Yellow
        try{$checkforplans = Invoke-RestMethod -Headers $Header -Uri $('https://graph.microsoft.com/v1.0/groups/' + $group.id + '/planner/plans') -Method Get}
        catch {}
        if ($checkforplans.value.count -gt 0){
            ForEach($plan in $checkforplans.value){
                $output = @{
                    DisplayName = $group.displayname
                    GroupID = $group.id
                    Title = $plan.title
                    PlanID = $plan.ID
                }
                $plans += new-object psobject -Property $output
            }
        }
    }
    return $plans
}
##  This function returns the $groupID in the form of an array. This
#   is an intermediary function that produces output that will be
#   passed to the GetPlans function.
Function GetUnifiedGroups ($token, $GroupID) {
    $Header = @{
        Authorization = $token
    }
    $content = @()
    $unifiedgroupsout = @()
    $groupsUrl = "https://graph.microsoft.com/v1.0/Groups/$GroupID"
    Write-Host "Getting unified groups from tenant"  -ForegroundColor Yellow
    while (-not [string]::IsNullOrEmpty($groupsUrl)) {
        $Data = Invoke-RestMethod -Headers $Header -Uri $groupsUrl -Method Get
        $Data      
        if ($data.'@odata.nextLink'){$groupsUrl = $data.'@odata.nextLink'}
        else {$groupsUrl = $null}
        $content += ($Data | select-object Value).Value
    }
    Foreach ($group in $content){
        if ($group.grouptypes -eq 'Unified'){
        $unifiedgroupsout += $group
        }
    }
    return $unifiedgroupsout
}
##  This function retrieves all buckets and their properties for one Plan.
Function GetBucket ($bucketid, $token){
    $Header = @{
        Authorization = $token
    }
    return Invoke-RestMethod -Headers $Header -Uri $('https://graph.microsoft.com/v1.0/planner/buckets/' + $bucketid ) -Method Get
}
##  This function uses the Device Code authentication
##  method to pass the admin's user credentials to
##  get the authorization token.
##  This is a more secure method than the previous
##  method of storing the username, password, and
##  client secret all in plain text.
Function GetDeviceCode {
    $deviceCodeResponse = @()
    $deviceCodeRequest = @()
    do {
        Clear-Host
        Write-Host "You will need to go to https://microsoft.com/devicecode."
        Write-Host "Enter the code shown below."
        $deviceCodeRequestParams = @{
            Method = 'POST'
            Uri = "https://login.microsoftonline.com/$tenantID/oauth2/devicecode"
            Body = @{
                client_id = $clientID
                resource = $resource
                client_secret = $clientSecret
            }
        }
        $deviceCodeRequest = Invoke-RestMethod @DeviceCodeRequestParams
        Write-Host $deviceCodeRequest.message -ForegroundColor Yellow
        Write-Host "After you enter the code, respond with Y to proceed."
        $deviceCodeResponse = Read-Host "Enter Y after you complete the device code login."
    }
    until ($deviceCodeResponse)
    return $deviceCodeRequest
}
Function GetPasswordAuthToken ($deviceCodeRequest){
    $TokenRequestParams = @{
        Method = 'POST'
        Uri = "https://login.microsoftonline.com/$tenantID/oauth2/token"
        Body = @{
            grant_type = "urn:ietf:params:oauth:grant-type:device_code"
            code = $deviceCodeRequest.device_code
            client_secret = $clientSecret
            client_id = $clientID
        }
    }
    $tokenRequest = Invoke-RestMethod @TokenRequestParams
    Start-Sleep -Milliseconds 4000
    $codeToken = @()
    $codeToken = $tokenRequest.access_token
    return $codeToken
}
#   This function retrieves all users currently in the source Group. This
#   is a function that came directly from the original script, and I have
#   not touched this yet, so I can tell you that it works but can't really
#   explain how it works.
Function GetAllUsers($token) {
    $Header = @{
        Authorization = $token
    }
    Write-Host "Getting all Users from tenant" -ForegroundColor Yellow
    $content = @()
    $usersUrl = "https://graph.microsoft.com/v1.0/Users"
    while (-not [string]::IsNullOrEmpty($usersUrl)) {
        $Data = Invoke-RestMethod -Headers $Header -Uri $UsersUrl -Method Get
        if ($data.'@odata.nextLink'){$usersUrl = $data.'@odata.nextLink'}
        else {$usersUrl = $null}
        $content += ($Data | select-object Value).Value
    }
    return $content
}
#   This function gets the names of each bucket per Plan per Group.
#   The returned data includes Plan name and ID, bucket name and
#   ID, and the bucket order. The bucket order places the new
#   buckets into the same left to right order as in the source
#   Plan.
function GetBuckets($plan, $Token){
    $Header = @{
        Authorization = $Token
    }
    $buckets = @()
    $plan.SourcePlanID
    Write-Host "`r`nQuerying Buckets in $($plan.title) ..." -ForegroundColor Yellow
    $checkForBuckets = Invoke-RestMethod -Headers $Header -Uri $('https://graph.microsoft.com/v1.0/planner/plans/' + $plan.SourcePlanID + '/buckets') -Method Get
    if ($checkForBuckets.value.count -gt 0){
        ForEach($bucket in $checkForBuckets.value){
            if ($null -ne $bucket.name){
            $output = @{
                Title = $plan.title
                PlanID = $bucket.planid
                BucketID = $bucket.id
                BucketName = $bucket.name
                BucketOrderHint = $bucket.orderhint
            }
            $buckets += new-object psobject -Property $output
            }
        }
    }
    return $buckets
}
#  This function is practically the same as the previous GetBuckets
#  function except that it is intended to get the bucket names
#  and other properties present in the newly created Plan.
#  Function name means get target buckets.
function GetTGTBuckets($plan, $Token){
    $Header = @{
        Authorization = $Token
    }
    $buckets = @()
    Write-Host "`r`nQuerying TGTBuckets in plan id $plan ..." -ForegroundColor Yellow
    $checkForBuckets = Invoke-RestMethod -Headers $Header -Uri $('https://graph.microsoft.com/v1.0/planner/plans/' + $plan + '/buckets') -Method Get
    if ($checkForBuckets.value.count -gt 0){
        ForEach($bucket in $checkForBuckets.value){
            $output = @{
                Title = $plan.title
                PlanID = $bucket.planid
                BucketID = $bucket.id
                BucketName = $bucket.name
                BucketOrderHint = $bucket.orderhint
            }
            $buckets += new-object psobject -Property $output
        }
    }
    return $buckets
}
#  The following function gets source Plan buckets.
function GetBucket ($bucketid, $token){
    $Header = @{
        Authorization = $token
    }
    return Invoke-RestMethod -Headers $Header -Uri $('https://graph.microsoft.com/v1.0/planner/buckets/' + $bucketid ) -Method Get
} 
#  This function gets all tasks within one Plan. This function
#  was present in the original script, so I don't know much
#  about how this was written.
function GetTasksByPlan($token, $planid){
    $Header = @{
        Authorization = $onlyToken
    }
    $content = @()
    $tasksUrl = $('https://graph.microsoft.com/v1.0/planner/plans/' + $planid + '/tasks')
    while (-not [string]::IsNullOrEmpty($tasksUrl)) {
        Write-Host "`r`nQuerying $tasksUrl..." -ForegroundColor Yellow
        $Data = Invoke-RestMethod -Headers $Header -Uri $tasksUrl -Method Get
        if ($data.'@odata.nextLink'){$tasksUrl = $data.'@odata.nextLink'}
        else {$tasksUrl = $null}
        $content += ($Data | select-object Value).Value
    }
    return $content
}
#  This function gets the task comments threads, if present, for each task
#  within one Plan. This function provides the comment ID but does not
#  return the full comment value.
function GetTaskComments ($token, $GroupID) {
    $Header = @{
        Authorization = $token
    }
    $threadsURL = $('https://graph.microsoft.com/v1.0/groups/' + $GroupID + '/threads')
    while (-not [string]::IsNullOrEmpty($threadsUrl)) {
        Write-Host "`r`nQuerying $threadsUrl..." -ForegroundColor Yellow
        $Data = Invoke-RestMethod -Headers $Header -Uri $threadsUrl -Method Get
        if ($data.'@odata.nextLink'){$threadsUrl = $data.'@odata.nextLink'}
        else {$threadsUrl = $null}
        $content += ($Data | select-object Value).Value
    }
    return $content
}
#  This function gets the full task comments for each task with
#  comments within one Plan.
function GetTaskCommentsDetails ($token, $GroupID, $threadID) {
    $Header = @{
        Authorization = $token
    }
    $threadsURL = $('https://graph.microsoft.com/v1.0/groups/' + $GroupID + '/threads/' + $threadID + '/posts')
    while (-not [string]::IsNullOrEmpty($threadsUrl)) {
        Write-Host "`r`nQuerying $threadsUrl..." -ForegroundColor Yellow
        $Data = Invoke-RestMethod -Headers $Header -Uri $threadsUrl -Method Get
        if ($data.'@odata.nextLink'){$threadsUrl = $data.'@odata.nextLink'}
        else {$threadsUrl = $null}
        $content += ($Data | select-object Value).Value
    }
    return $content
}
function mainfunction {
#   Initialize variables to be used in the GetDeviceCode
#   and the GetPasswordAuthToken functions.
$clientID = "INSERT CLIENT ID AKA APPLICATION ID"
$tenantID = "INSERT TENANT ID"
$resource = "https://graph.microsoft.com"
#   $clientSecret was encrypted using:
#   ConvertTo-SecureString "clientsecret" -AsPlainText -Force | ConvertFrom-SecureString | Out-File C:\FilePath\clientsecret.txt
$clientSecret = Get-Content "C:\Path\To\secureclientsecret.txt" | ConvertTo-SecureString
#   This calls the GetDeviceCode and GetPasswordAuthToken 
#   functions and stores the returned data in the
#   variables $deviceCodeRequest and $onlyToken.
$deviceCodeRequest = GetDeviceCode
$onlyToken = GetPasswordAuthToken $deviceCodeRequest
#   Sometimes the $onlyToken array variable can contain more
#   than one item in the array so this checks if there is more
#   than one item in the array and if that is true, it then sets
#   the $onlyToken variable as the value of the first (position 0) item.
$count = $onlyToken.count
if ($count -gt 1) {
    $onlyToken = $onlyToken[-1]
}
#   This is the first interactive part of the script. The script user
#   enters the source Group name, which can be entered in the form of
#   the current DisplayName or the original Group name or email address.
#   It requires connecting to the ExchangeOnlineManagement PowerShell module.
#   
#   Side note: you probably can retrieve this info through a Graph API
#   call but I did what I knew would work.
$getSourceGroup = Read-Host -Prompt "Enter the source group name"
$getSourceGroupID = Get-UnifiedGroup -Identity $getSourceGroup | Select-Object 'ExternalDirectoryObjectID'
#   The following for loop retrieves the value of the ExternalDirectoryObjectID (GroupId)
#   from the array $getSourceGroupID and sets it as the string variable for $groupSourceID.
foreach ($p in $getSourceGroupID){
    $groupSourceID = $p.ExternalDirectoryObjectID
}
#   This is the second interactive part of the script. The script user
#   enters the destination (target) Group name, which can be entered
#   in the form of the current DisplayName or the original Group name
#   or email address. This also requires a connection to the 
#   ExchangeOnlineManagement module.
$getTargetGroup = Read-Host -Prompt "Enter the destination group name"
$getTargetGroupID = Get-UnifiedGroup -Identity $getTargetGroup | Select-Object 'ExternalDirectoryObjectID'
#   The following for loop retrieves the value of the ExternalDirectoryObjectID (GroupId)
#   from the array $getSourceGroupID and sets it as the string variable for $groupSourceID.
foreach ($q in $getTargetGroupID){
    $groupTargetID = $q.ExternalDirectoryObjectID
}
#   Pass the $onlyToken and $groupSourceID parameters to the
#   GetUnifiedGroups function. Return data in the variable $sourceGroup.
$sourceGroup = GetUnifiedGroups $onlyToken $groupSourceID
#   Pass the $onlyToken and $groupTargetID parameters to the
#   GetUnifiedGroups function. Return data in the variable $targetGroup.
#
#   Pass the $onlyToken and the previously created $sourceGroup parameters
#   to the GetPlans function. This returns data of each Plan present
#   within the $sourceGroup.
$sourcePlans = GetPlans $sourceGroup $onlyToken
#
#   The following 2 commands require connecting to the MicrosofTeams
#   PowerShell module. This function could be accomplished with a Graph
#   API call, but again I know that this works which is why I put it
#   in this script.
#   This returns the Team's (associated with the Group) user list
#   and excludes guest accounts from being added to the lists of 
#   $sourceTeamUsers and $targetTeamUsers.
#   This is another function that was present in the original script, and
#   this function is useful because it does not add the same users
#   from the source Group to the destination Group when they already
#   exist in the destination Group.
$sourceTeamUsers = Get-Team -GroupId $groupSourceID | Get-TeamUser | Select-Object 'User' | Where-Object {($_.User -notlike '*EXT#*')}
$targetTeamUsers = Get-Team -GroupId $groupTargetID | Get-TeamUser | Select-Object 'User' | Where-Object {($_.User -notlike '*EXT#*')}
Write-Output "Comparing users and adding new users to target team."
$userCompare = Compare-Object -ReferenceObject $sourceTeamUsers -DifferenceObject $targetTeamUsers -Property User -PassThru -IncludeEqual
foreach ($diff in $userCompare) {
    if ($diff.SideIndicator -eq "<="){
        Add-TeamUser -GroupID $groupTargetID -User $diff.User
    }
}
#   The following commands use the GetAllUsers function to populate
#   the $sourceUsers and $targetUsers variables. These variables
#   will be used later for assigning tasks in the destination
#   Group's Plan(s) to the correct user.
#   This function was present in the original script and uses the
#   filters that are at the top of this script to get the user
#   name data in the correct format.
$sourceUsers = GetAllUsers $onlyToken
$targetUsers = GetAllUsers $onlyToken
$sourceAndTargetUserIDs = @()
foreach($i in $sourceUsers){
    $targetUser = $targetUsers | displayname-filter -filterString $i.displayname
    
    $output = @{
        displayname = $sourceuser.displayname
        sourceuserid = $sourceuser.id
        targetuserid = $targetUser.id
    }
    $sourceAndTargetUserIDs += New-Object psobject -Property $output
}
#   Filter the $sourcePlans variable to create two separate variables:
#   $sourcePlanList contains DisplayName and Title
#   $sourcePlanIDs contains PlanID
$sourcePlanList = $sourcePlans | Select-Object 'DisplayName','Title'
$sourcePlanIDs = $sourcePlans | Select-Object 'PlanID'
#   Initialize the $newPlanIDs & $newPlanNames arrays.
$newPlanIDs = @()
$newPlanNames = @()
#   The following for loop creates a new Plan in the destination Group
#   that has the same DisplayName as the Plan in $sourcePlanList.
#   Success/failure at this step is written to the console for each
#   Plan in $sourcePlanList.
foreach ($plan in $sourcePlanList) {
    $Header = @{
        Authorization = $onlyToken
    }
    $payloadNewPlan = @{ owner = $groupTargetID; title = $plan.title }
    $jsonpayload = $payloadNewPlan | ConvertTo-Json
    $success = Invoke-RestMethod -Method POST -Uri 'https://graph.microsoft.com/v1.0/planner/plans' -Headers $Header -Body $jsonpayload -ContentType 'application/json'
    if ($null -eq $success){
        Write-Host "Failed to add plan " $plan.title " to group " $getTargetGroup "." -ForegroundColor Red
    }
    elseif ($success){
        Write-Host "The Plan " $plan.title " has been added to the group " $getTargetGroup -ForegroundColor Green
        $newPlanIDs += $success.id
        $newPlanNames += $success.Title
    }
}
#   Initialize counter variable $g that will be used in a for
#   loop to get the correct indexed item in the $sourcePlanIDs
#   and $newPlanIDs arrays. Increment $g by plus 1 until
#   the value of $g is equal to $sourcePlanIDs' array length.
$g = 0;
$groupsPlansList = @()
foreach ($group in $newPlanIDs){
    $output = @{
        DisplayName = $getTargetGroup
        SourceGroupID = $groupSourceID
        Title = $success.title
        SourcePlanID = $($sourcePlanIDs[$g]).PlanID
        TargetGroupID = $groupTargetID
        TargetPlanID = $newPlanIDs[$g]
    }
    $groupsPlansList += new-object psobject -Property $output
    if ($g -lt $sourcePlanIDs.count){
        $g++
    }   
}
#  Initialize arrays that will be plus-oned throughout the
#  script so that if the script is run again the values in
#  the arrays are cleared.
$sourceTasks = @()
$targetTasks = @()
$taskcompare = @()
$data = @()
#  The following retrieves the buckets from the source
#  Plan.
foreach ($plan in $groupsPlansList) {
    $planBuckets = @()
    $Header = @{
        Authorization = $onlyToken
    }
    Write-Host "Getting buckets from plan" $plan.title -ForegroundColor Yellow
    $planBuckets += GetBuckets $plan $onlyToken
   [array]::Reverse($planBuckets)
   #  Based on previous notes, it seems like the rest of this for
   #  loop doesn't work and could probably be removed. Will need
   #  to test this.
    $targetPlanID = $plan.TargetPlanID
    $amountplanBuckets = $planBuckets.count
    $amountplanBucketsminusone = $amountplanBuckets - 1
    $h = 0
    foreach ($bucket in $planBuckets) {
        if ($h -lt $amountplanBucketsminusone){
        $payload = @{ name = $bucket.bucketname; planId = $targetPlanID; orderHint = " !" }
        $jsonpayload = $payload | ConvertTo-Json
        Invoke-RestMethod -Method POST -Uri 'https://graph.microsoft.com/v1.0/planner/buckets' -Headers $Header -Body $jsonpayload -ContentType 'application/json'
        $h++
    }
}
}

#  The following retrieves the newly created buckets
#  in the target Plan.
Start-Sleep -Milliseconds 4000
foreach ($plan in $groupsPlansList) {
    $Header = @{
        Authorization = $onlyToken
    }
    Write-Host "Getting target buckets from plan" $plan.title -ForegroundColor Yellow
    $planTGTBuckets += GetTGTBuckets $plan.TargetPlanID $onlyToken
}
#  The following function is the biggest function within
#  this script. It is going to query the task
#  attributes and put them in the same order and
#  in the same bucket.
#  The first part of the script compares source tasks
#  and target tasks. There should be no tasks in the new
#  Plan.
#  Next it creates each task and adds certain details
#  before moving on to the next task.
foreach ($plan in $groupsPlansList) {
    $sourceTasks = GetTasksByPlan $onlyToken $plan.SourcePlanID | Sort-Object orderHint -Descending
    $targetTasks = GetTasksByPlan $onlyToken $plan.TargetPlanID
    $Header2 = @{
        Authorization = $onlyToken;
    }
    $sourcePlanDetailsURI = 'https://graph.microsoft.com/v1.0/planner/plans/' + $plan.SourcePlanID + '/details'
    $sourcePlanDetails = Invoke-RestMethod -Method GET -Uri $sourcePlanDetailsURI -Header $Header2
    $targetPlanDetailsURI = 'https://graph.microsoft.com/v1.0/planner/plans/' + $plan.TargetPlanID + '/details'
    $targetPlanDetails = Invoke-RestMethod -Method GET -Uri $targetPlanDetailsURI -Header $Header2
    $payload = @{categoryDescriptions = $sourcePlanDetails.categoryDescriptions}
    $jsonTaskPayload = $payload | ConvertTo-Json
    $targetPlanETag = $targetPlanDetails.'@odata.etag'
    $Header3 = @{
        Authorization = $onlyToken;
        'If-Match' = $targetPlanETag
    }
    Invoke-RestMethod -Method PATCH -Uri $targetPlanDetailsURI -Header $Header3 -Body $jsonTaskPayload -ContentType 'application/json'
    $success = $null
#  If there are tasks within the Plan, do the
#  following:
    if($sourceTasks){
        $Header = @{
            Authorization = $onlyToken
        }
            foreach($task in $sourceTasks){
                #  Initialize arrays
                $taskUpdateComment = @()
                $taskAssignees = @()
                $taskAssignments = @{}
                $taskPayload = @()
                #  Get the source Plan's buckets' names,
                #  target Plan's ID, source Plan's ID, target buckets'
                #  names, and filter the buckets in the new Plan
                #  in order to select the correct bucket for a task.
                $sourcePlanBucketName = $(GetBucket $task.bucketid $onlyToken).name
                $targetPlanID = $plan.TargetPlanID
                $sourcePlanID = $(GetBucket $task.bucketid $onlyToken).PlanID
                $targetBucket = ($planTGTBuckets).BucketName
                $targetBucketForTask = $targetBucket | Where-Object {$_ -match $sourcePlanBucketName}
                $filterTargetBucket = $planTGTBuckets | Where-Object {($_.BucketName) -match $targetBucketForTask}
                $filterTargetBucketID = $filterTargetBucket.BucketID
                #  If the task is assigned to users
                if (-not [string]::IsNullOrEmpty($task.assignments)){
                    Write-Host "Checking for task assignees and building JSON payload" -ForegroundColor Yellow
                    $sourceTaskAssignees = Get-Member -InputObject $task.assignments -MemberType NoteProperty
                    $sourceTaskAssigneesNames = $sourceTaskAssignees.Name
                    foreach ($name in $sourceTaskAssigneesNames) {
                        $taskAssignees += $sourceAndTargetUserIDs.targetuserid | Where-Object {$_ -match $name}
                    }
                    foreach($assignee in $taskAssignees){
                        $payload = @{$assignee = @{ '@odata.type' = "microsoft.graph.plannerAssignment"; orderHint = " !" }}
                        $taskAssignments += $payload
                    }
                }             
                if($taskAssignments.keys.count -gt 0){
                    $taskPayload = @{ planId = $plan.TargetPlanID ; bucketId = $filterTargetBucketID ; title = $task.title; assignments = $taskAssignments}
                }
                else {$taskPayload = @{ planId = $plan.TargetPlanID ; bucketId = $filterTargetBucketID ; title = $task.title}} 
                $jsonTaskPayload = $taskPayload | ConvertTo-Json -Depth 4             
                #  Create the target Task in the target Plan.
                #  Store the success of creating the target task
                #  in order to continue adding more attributes
                #  to it, so long as the task was successfully
                #  created.
                $createNewTaskURI = $('https://graph.microsoft.com/v1.0/planner/tasks')
                $success = Invoke-RestMethod -Method POST -Uri $createNewTaskURI -Headers $Header -Body $jsonTaskPayload -ContentType 'application/json'
                #  Brief pause before adding more attributes
                #  to the new target Task.
                Start-Sleep -Milliseconds 2000
                #  Initialize array that will store task completion,
                #  task completed by, notes, and comments
                #  depending on if the task contains these
                #  attributes.
                $notesHeaderPayload = @()
                #  If the task is completed, then get the name of
                #  the user who completed the task.
                if(($task.percentComplete) -eq "100") {
                    $userCompletedTaskID = $task.completedBy.user.id
                    $getURI = 'https://graph.microsoft.com/v1.0/users/' + $userCompletedTaskID
                    $getResponse = Invoke-RestMethod -Method GET -Uri $getURI -Header $Header
                    $taskCompletedBy = $getResponse.displayName
                    if ($taskCompletedBy){
                        $notesHeaderPayload = "Task completed on $($task.completedDateTime) by $taskCompletedBy."
                    }
                }
                #  If the target task is successfully created
                if($success){
                    Write-host -ForegroundColor Green "Task successfully created."
                    #  Initialize arrays
                    $taskPayload = @()
                    $taskUpdateComment = @()
                    #  Get the target task's ID and title.
                    $newTaskID = $success.id
                    $newTaskTitle = $success.Title
                    #  Get source and target task's details.
                    $sourceTaskDetailsURI = $('https://graph.microsoft.com/v1.0/planner/tasks/' + $task.id + '/details')
                    $sourceTaskDetails = Invoke-RestMethod -Headers $Header -Method 'GET' -Uri $sourceTaskDetailsURI -ContentType 'application/json'
                    $targetTaskDetailsURI = $('https://graph.microsoft.com/v1.0/planner/tasks/' + $newTaskID + '/details')
                    $targetTaskDetails = Invoke-RestMethod -Headers $Header -Method 'GET' -Uri $targetTaskDetailsURI -ContentType 'application/json'
                    #  Create new variable for the new task's etag.
                    $newTaskETag = $targetTaskDetails.'@odata.etag'
                    ###################################
                    #  APPLY TASK LABELS ##############
                    ###################################
                    $Header = @{
                        Authorization = $onlyToken;
                        'If-Match' = $success.'@odata.etag'
                    }
                    $labelCategories = @{appliedCategories = $task.appliedCategories}
                    $targetLabelNames = $labelCategories | ConvertTo-Json
                    $targetTaskURI = $('https://graph.microsoft.com/v1.0/planner/tasks/' + $newTaskID)
                    Invoke-RestMethod -Headers $Header -Method 'PATCH' -Uri $targetTaskURI -Body $targetLabelNames -ContentType 'application/json'
                    ###################################
                    # END APPLY TASK LABELS ###########
                    ###################################                 
                    #  Create new $header variable to use with
                    #  Graph for the new task details. 
                    #  Token is the newTaskETag.
                    $Header = @{
                        Authorization = $onlyToken;
                        'If-Match' = $newTaskETag
                    }
                    ###################################
                    #  GET TASK COMMENTS ##############
                    ###################################
                    #  The following section gets the task comments
                    #  if any exist, and strips the actual comment
                    #  from the html-coded comment data that is
                    #  returned through Graph API.
                    ###################################
                    #  Initialize array to store comment
                    #  details from the source Group.
                    $sourceGroupComments = @()
                    $sourceGroupComments = GetTaskComments $onlyToken $groupSourceID
                    #  Initialize arrays to store which comment
                    #  is associated with which task, and get
                    #  each comments' IDs.
                    $taskCommentTaskTitle = @()
                    $taskCommentIDArray = @()
                    foreach ($sourceComment in $sourceGroupComments) {
                        #  Filter $sourceGroupComments' data for actual
                        #  task comments since it includes other data.
                        $comment = $sourceComment.topic
                        #  If the topic has the word Comments in it
                        if ($comment -like "*Comments*"){
                            #  Split the comment and add the task Comment ID
                            #  if it exists.
                            $newTaskTitle2 = $newTaskTitle.Substring(0,$newTaskTitle.Length-1)
                            if (($comment -like "*$($newTaskTitle2)*") -or ($comment -like "*$($newTaskTitle)*")){
                                $taskTitleName = $sourceComment.topic -split "\`""
                                $taskTitleName2 = $taskTitleName[1]
                                $taskCommentIDArray += @($sourceComment.id)
                                $taskCommentTaskTitle += @($taskTitleName2)
                            }
                        }
                    }
                    #  If there are task comments
                        if($taskCommentIDArray){
                            foreach ($id in $taskCommentIDArray){
                                #  Initialize variables.
                                $commentFirstSplit = ""
                                $commentFirstSplitFirstItem = ""
                                $commentSecondSplit = ""
                                $commentSecondSplitLastItem = ""
                                $taskCommentDetails = GetTaskCommentsDetails $onlyToken $groupSourceID $id
                                if ($taskCommentDetails.count -gt 1) {[array]::Reverse($taskCommentDetails)}
                                foreach ($q in $taskCommentDetails){
                                    #### need to figure out what to do when someone has a <br> after the end of
                                    #### the comment before the <table id
                                $commentBody = $q.body.content
                                $commentFirstSplit = $commentBody -split "<table id="
                                $commentFirstSplitFirstItem = $commentFirstSplit[0]
                                $commentSecondSplit = $commentFirstSplitFirstItem -split "<div>"
                                $commentSecondSplitLastItem = $commentSecondSplit[-1]
                                #  If the last letter?? in the second split is $null
                                if ($null -eq $commentSecondSplitLastItem) {
                                    #  Find the <br and then create another split
                                    foreach ($j in $commentSecondSplitLastItem) {
                                        $breakTrailer = "<br"
                                        if ($j -like "*$($breakTrailer)*") {
                                            $commentSecondSplitAlone = $j -split $breakTrailer
                                            $comment = $commentSecondSplitAlone[0]
                                        }
                                }
                                $taskUpdateComment += @($comment,"`r`n",$q.lastModifiedDateTime,"`r`n",$q.from.emailAddress.name,"`r`n")
                                }
                                #  If the last letter in the second split is NOT
                                #  a null value
                                else {
                                    $comment = $("`r`n" + $commentSecondSplitLastItem)
                                    $taskUpdateComment += @($comment,"`r`n",$q.lastModifiedDateTime,"`r`n",$q.from.emailAddress.name,"`r`n")
                                }
                            }
                        }
                        #  Reverse the order of the $taskUpdateComment
                        #  so that the comment order is in the same
                        #  order from newest to oldest.
                        [array]::Reverse($taskUpdateComment)
                        }
                        ###################################
                        # END GET TASK COMMENTS ###########
                        ###################################
                        ###################################
                        ###################################
                        #  TASK NOTES PAYLOAD #############
                        ###################################
                        #  This section adds any combo of task completion,
                        #  task comments, and actual notes together and
                        #  then publishes that info to the target task.
                        $breakHTML = "`r`n"
                        if ($task.hasDescription) {
                            if ($taskUpdateComment -and $notesHeaderPayload) {
                                write-output "description and comments and completedby"
                            $taskUpdateDescription = @{
                                "description" = "$($notesHeaderPayload + "`r`n"+ $sourceTaskDetails.description + $taskUpdateComment)";
                                "previewType" = "description"
                            }
                            $taskPayload = $taskUpdateDescription
                            }
                            elseif ($taskUpdateComment -and !($notesHeaderPayload)){
                                Write-Output "description and comments, no completedby"
                                $taskUpdateDescription = @{
                                    "description" = "$($sourceTaskDetails.description + $taskUpdateComment)"
                                    "previewType" = "$($sourceTaskDetails.previewType)"
                                }
                                $taskPayload = $taskUpdateDescription
                            }
                            elseif (!($taskUpdateComment) -and $notesHeaderPayload) {
                                write-output "description, no comments, completedby"
                            $taskUpdateDescription = @{
                                "description" = "$($notesHeaderPayload + "`r`n" + $sourceTaskDetails.description)";
                                "previewType" = "description"
                            }
                            $taskPayload = $taskUpdateDescription
                            }
                            elseif (!($taskUpdateComment) -and !($notesHeaderPayload)){
                                Write-Output "description, no comments, no completedby"
                            $taskUpdateDescription = @{
                                "description" = "$($sourceTaskDetails.description)"
                                "previewType" = "$($sourceTaskDetails.previewType)"
                            }
                            $taskPayload = $taskUpdateDescription
                            }
                        }
                        elseif (!($task.hasDescription)) {
                            if ($taskUpdateComment -and $notesHeaderPayload) {
                                write-output "comments and completedby"
                            $taskUpdateDescription = @{
                                "description" = "$($notesHeaderPayload + "`r`n" + $taskUpdateComment)";
                                "previewType" = "description"
                            }
                            $taskPayload = $taskUpdateDescription
                            }
                            elseif ($taskUpdateComment -and !($notesHeaderPayload)){
                                Write-Output "comments, no completedby"
                                $taskUpdateDescription = @{
                                    "description" = "$($taskUpdateComment)"
                                    "previewType" = "$($sourceTaskDetails.previewType)"
                                }
                                $taskPayload = $taskUpdateDescription
                            }
                            elseif (!($taskUpdateComment) -and $notesHeaderPayload) {
                                write-output "description, no comments, completedby"
                            $taskUpdateDescription = @{
                                "description" = "$($notesHeaderPayload)";
                                "previewType" = "description"
                            }
                            $taskPayload = $taskUpdateDescription
                            }
                            elseif (!($taskUpdateComment) -and !($notesHeaderPayload)){
                                Write-Output "nodescription, no comments, no completedby"
                            $taskUpdateDescription = @()
                            $taskPayload = $taskUpdateDescription
                            }
                        }
                        if ($taskPayload){           
                            $jsonpayload = $taskPayload | ConvertTo-Json
                            $updateURI = 'https://graph.microsoft.com/v1.0/planner/tasks/' + $newTaskID + '/details'
                            Invoke-RestMethod -Method PATCH -Uri $updateURI -Header $Header -Body $jsonpayload -ContentType 'application/json'
                        }
                        ###################################
                        #  END TASK NOTES PAYLOAD #########
                        ###################################
                        ###################################
                        #  GET TASK CHECKLISTS ############
                        ###################################              
                        #  This section gets the checklist items
                        #  and checks off any completed items.
                        #  If there are checklist items:
                        if (![string]::IsNullOrEmpty($sourceTaskDetails.checklist)) {
                                foreach ($item in $sourceTaskDetails.checklist.psobject.Properties) {
                                    $guid = [guid]::newGuid().guid
                                    $checklistValue = $item.value
                                    $checklistValue.orderhint = " !"
                                    $checklistValue.PSObject.Properties.Remove('LastModifiedDateTime')
                                    $checklistValue.PSObject.Properties.Remove('LastModifiedBy')
                                    $checklistPayload = @()
                                    $checklistPayload = @{$guid = $checklistValue}
                                    $taskUpdateChecklist = @{"checklist" =    $($checklistPayload)}
                                    $jsonTaskPayload = $taskUpdateChecklist | ConvertTo-Json
                                    $updateURI = 'https://graph.microsoft.com/v1.0/planner/tasks/' + $newTaskID + '/details'
                                    Invoke-RestMethod -Method PATCH -Uri $updateURI -Header $Header -Body $jsonTaskPayload -ContentType 'application/json'
                                    if ($counterVar -lt $sourceTaskDetails.checklist.psobject.Properies.length){
                                        $counterVar++
                                    }
                                    }
                                }
                        ###################################
                        #  END GET CHECKLISTS #############
                        ###################################
                        ###################################
                        #  GET TASK COMPLETION AND DUE ####
                        #  DATES AND START DATES ##########
                        ###################################
                        #  The following section gets the task completion
                        #  or progress status, and the due/start dates.
                        #  If the task is in progress or completed:
                        if ($task.percentComplete -ne 0){
                            $newTaskETag2 = $success.'@odata.etag'
                            $Header = @{
                                Authorization = $onlyToken;
                                'If-Match' = $newTaskETag2
                            }
                            $taskCompletion = $task.percentComplete
                            $completion = @{
                                "percentComplete" = $taskCompletion
                                }
                                $jsonTaskPayload = $completion | ConvertTo-Json
                                $updateURI = 'https://graph.microsoft.com/v1.0/planner/tasks/' + $newTaskID
                                Invoke-RestMethod -Method PATCH -Uri $updateURI -Header $Header -Body $jsonTaskPayload -ContentType 'application/json'
                                }
                        #  If the task has a due date
                        try {if ($task.dueDateTime){
                            $newTaskETag2 = $success.'@odata.etag'
                            $Header = @{
                                Authorization = $onlyToken
                                'If-Match' = $newTaskETag2
                            }
                            $taskDueDate = $task.dueDateTime                    
                            $payload = @{"dueDateTime" = $taskDueDate}
                            if ($task.startDateTime){
                                $taskStartDate = $task.startDateTime
                                $taskStartDatePayload = @{"startDateTime" = $taskStartDate}
                                $payload += $taskStartDatePayload
                            }
                            $jsonTaskPayload = $payload | ConvertTo-Json
                            $updateURI = 'https://graph.microsoft.com/v1.0/planner/tasks/' + $newTaskID
                            Invoke-RestMethod -Method PATCH -Uri $updateURI -Header $Header -Body $jsonTaskPayload -ContentType 'application/json'
                        }
                        #  If the task does not have a due date
                        elseif (!($task.dueDateTime)) {
                            $payload = @()
                            if ($task.startDateTime){
                                $taskStartDate = $task.startDateTime
                                $taskStartDatePayload = @{"startDateTime" = $taskStartDate}
                                $payload += $taskStartDatePayload
                            }
                            $jsonTaskPayload = $payload | ConvertTo-Json
                            $updateURI = 'https://graph.microsoft.com/v1.0/planner/tasks/' + $newTaskID
                            Invoke-RestMethod -Method PATCH -Uri $updateURI -Header $Header -Body $jsonTaskPayload -ContentType 'application/json'
                        }
                    }
                    catch {#prevent error from being shown in the console
                    }                                       
                }
                else{Write-Host "Errors occurred." -ForegroundColor Red}
            }
        }
        else{
            Write-Host "Tasks were not found in this Plan." -ForegroundColor Yellow
    }
}
}

do {
    Clear-Host
    Write-Host "You will need to connect to:"
    Write-Host -ForegroundColor Red "MicrosoftTeams PowerShell Module"
    Write-Host "&"
    Write-Host -ForegroundColor Green "ExchangeOnline V2 PowerShell Module"
    Write-Host "********************************************"
    Write-Host "Have you run Connect-MicrosoftTeams & Connect-ExchangeOnline?"
    $moduleConnection = Read-Host "Enter Y or N."
    if ($moduleConnection -eq "Y"){
        mainfunction
    }
    else {
        Write-Host "Connect to the Teams PowerShell Module"
        Connect-MicrosoftTeams
        Write-Host "Connect to Exchange Online PowerShell Module"
        Connect-ExchangeOnline
    }
}
until ($moduleConnection -eq "Y")


