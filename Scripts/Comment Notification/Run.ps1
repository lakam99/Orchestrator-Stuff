##Better Notify Analyst {Comments}
#Written entirely by Arkam Mazrui
#arkam.mazrui@nserc-crsng.gc.ca
#arkam.mazrui@gmail.com
#The primary purpose of this script is to
#1) Check the assigned user didn't make the comment
#2) Retrieve assigned user email and send notification to them 

$sesh = New-PSSession ottansm1;
$r = Invoke-Command -Session $sesh -ScriptBlock {
    $evId = 5;
    Write-EventLog -LogName 'Orchestrator Script Messages' -Source 'Orchestrator' -EventId $evId -Message "A new comment has been created." -EntryType Information;

    Import-Module "D:\Program Files\Microsoft System Center\Service Manager\Powershell\System.Center.Service.Manager.psd1";

    $userProjectionType = Get-SCSMTypeProjection -Name System.User.Preferences.Projection;
    $srqProjection = Get-SCSMTypeProjection -Name System.WorkItem.ServiceRequestProjection;
    
    $ticket_id = '';
    $entered_by = '';
    
    $ticket = Get-SCSMObjectProjection -Projection $srqProjection -Filter "Id -eq $ticket_id";
    
    if (!($ticket)) {
        #failed to get ticket
        Write-EventLog -LogName 'Orchestrator Script Messages' -Source 'Orchestrator' -EventId $evId -Message "Failed to retrieve $ticket_id for comment entered by $entered_by." -EntryType Error;
    } else {
        #retrieved ticket
        $ticket_au = $ticket.AssignedTo; #retrieve assigned user

        if (!($ticket_au)) {
            #no assigned user
            Write-EventLog -LogName 'Orchestrator Script Messages' -Source 'Orchestrator' -EventId $evId -Message "Ticket $ticket_id has no assigned user. Cancelling notification." -EntryType Information;
        } else {
            #there is an assigned user
            if ($ticket_au.DisplayName -eq $entered_by) {
                Write-EventLog -LogName 'Orchestrator Script Messages' -Source 'Orchestrator' -EventId $evId -Message "New comment on ticket $ticket_id was entered by assigned user. Cancelling notification." -EntryType Information;
            } else {
                $user_projection = Get-SCSMObjectProjection -Projection $userProjectionType -Filter "Id -eq '$($ticket_au.Id.Guid)'";

                if (!($user_projection)) {
                    #couldn't get user projection
                     Write-EventLog -LogName 'Orchestrator Script Messages' -Source 'Orchestrator' -EventId $evId -Message "Failed to retrieve projection of $entered_by." -EntryType Error;
                } else {
                    #got user projection
                    Write-EventLog -LogName 'Orchestrator Script Messages' -Source 'Orchestrator' -EventId $evId -Message "Successfully retrieved projection of $entered_by." -EntryType Information;
                    $user_email = $null;
                    $user_smtp = $user_projection.Notification | ?{$_.DisplayName -like '*SMTP*'} | Select TargetAddress -ExpandProperty TargetAddress;
                    if (!($user_smtp)) {
                        #SMTP TargetAddress not found
                        Write-EventLog -LogName 'Orchestrator Script Messages' -Source 'Orchestrator' -EventId $evId -Message "Failed to retrieve SMTP TargetAddress of $entered_by." -EntryType Warning;
                        $user_email = $user_projection.UPN;
                    } else {
                        $user_email = $user_smtp;
                    }

                    if (!($user_email)) {
                        #User has no email assosciated with their account..?
                        Write-EventLog -LogName 'Orchestrator Script Messages' -Source 'Orchestrator' -EventId $evId -Message "Failed to retrieve email of $entered_by." -EntryType Error;
                    } else {
                        Write-EventLog -LogName 'Orchestrator Script Messages' -Source 'Orchestrator' -EventId $evId -Message "Retrieved email $user_email for $entered_by." -EntryType Information;
                        return @{email=$user_email;failed=$false};
                    }
                }
            }  
        }
    }
    return @{email=$null;$failed=true};
}
$failed = $r.failed;
$email = $r.email;

Disconnect-PSSession $sesh;
$sesh | Remove-PSSession;