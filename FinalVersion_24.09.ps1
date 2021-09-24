
$url_base = Read-Host 'Please write your base address'
$PAT = Read-Host 'Please write your organization PAT'
$user = Read-Host 'Please write your organization user name'
$computer_user_name = Read-Host 'Please write your computer user name for file path'
$file_name = Read-Host 'Please write your file name for output(please add .csv to the end of file name) '

Set-Content -Path C:\Users\$computer_user_name\Desktop\$file_name -Value   "WorkID, Title, IterationPath, Add, Delete, Edit, EpicID, FeatureID, ReqID" 

#Generate PAT
$token = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $user, $PAT)))
$header = @{authorization = "Basic $token"}

#List all repositories for the project
$url_endpoint_repositories = "_apis/git/repositories?api-version=6.0"
$url_repositories = $url_base + $url_endpoint_repositories

$repositories = Invoke-RestMethod -Uri "$url_repositories" -Method GET -ContentType "application/json" -headers $header 

   $repo_arr = @()
    foreach ($counter in $repositories.value) {

        $repo_arr += $counter
    }


    Write-Host "================ Repositories ================"
    $index= 0
    $repo = @()
    $repo_name = @()
    foreach ($item in $repo_arr){
        $ID = $item.id
        $name = $item.name
        
        $repo +=  $ID
        $repo_name += $name

        echo "$index)  Repo_name:  $name"
        $index += 1
   }   
    $repository_index = Read-Host 'Please enter your choice'
    $repository_id = $repo[$repository_index]
    $url_endpoint_pullrequests = "_apis/git/repositories/$repository_id/pullrequests?searchCriteria.status=completed&api-version=6.0"
    $url_pullrequests = $url_base + $url_endpoint_pullrequests

    $pullrequests = Invoke-RestMethod -Uri "$url_pullrequests" -Method GET -ContentType "application/json" -headers $header 
    
 
    $pull_request_arr = @()
    foreach ($pullrequest in $pullrequests.value) {

        $ref_name_path = $pullrequest.targetRefName
        $ref_name = $ref_name_path.Split("/")#split ayırmayı sağlıyor
        
        If($ref_name[2] -ne 'master')  
        {   
            $pull_request_arr += $pullrequest
            
           
            
        }
    }
    
    $pull = @{}
    $s = 0
    $updates = @()
    $final_info2 = [System.Collections.ArrayList]::new()
    foreach ($p in $pull_request_arr) {
        $addCount = 0
        $removeCount = 0
        $editCount = 0
        $pullId = $p.pullRequestId
        
        $url_endpoint_pullCommit = "_apis/git/repositories/$repository_id/pullRequests/$pullId/commits?api-version=6.0"
        $url_pullCommit = $url_base + $url_endpoint_pullCommit
        $pullCommit = Invoke-RestMethod -Uri "$url_pullCommit" -Method GET -ContentType "application/json" -headers $header 
        foreach ($commit in $pullCommit.value) {
           
            $commit_edits = @()
            $commit_adds = @()
            $commitId =$commit.commitId

            $url_endpoint_commit_info = "_apis/git/repositories/$repository_id/commits/$commitId ?api-version=5.0"
            $url_commit_info  = $url_base + $url_endpoint_commit_info 
            $commit_info  = Invoke-RestMethod -Uri "$url_commit_info" -Method GET -ContentType "application/json" -headers $header
            
            $parent_arr = $commit_info.'parents'

            $url_endpoint_forCommit = "_apis/git/repositories/$repository_id/commits/$commitId/changes?api-version=6.0"
            $url_forCommit = $url_base + $url_endpoint_forCommit
            $forCommit = Invoke-RestMethod -Uri "$url_forCommit" -Method GET -ContentType "application/json" -headers $header 
            if($null -ne  $forCommit.changes){
                foreach($changes in $forCommit.changes){
                    if($changes.item.'isFolder'.count -eq 0){
                        if($changes.changeType -eq 'edit'){
                            $commit_edits += $changes.item
                        }
                        elseif($changes.changeType -eq 'add'){
                            $commit_adds += $changes.item
                        }
                     }
                }
                
                foreach($edit in $commit_edits){
                    $path = $edit.path
                    foreach ($parent in $parent_arr) {
                          $parentID =  $parent
                          
                          $value = @{"originalPath"=$path;"originalVersion"=$parentID;"modifiedPath"=$path;"modifiedVersion"=$commitId} | ConvertTo-Json -Compress
                          $url_endpoint_post = "_api/_versioncontrol/fileDiff?__v=5&diffParameters=$value&repositoryId=$repository_id"
                          $url_post = $url_base + $url_endpoint_post
                          $post = Invoke-RestMethod -Uri "$url_post" -Method POST -ContentType "application/json" -headers $header #| convertto-json -depth 100
                          foreach($blocks in $post."blocks")
                        { 
                            $changeType = $blocks."changeType"

                            if($changeType -eq 1)
                            {
                                $addCount += $blocks."mLinesCount"
                            }elseif($changeType -eq 2)
                            {
                                $removeCount += $blocks.oLinesCount
                            }elseif($changeType -eq 3)
                            {
                                if($blocks.mLinesCount -eq $blocks.oLinesCount){
                                    $editCount += $blocks.mLinesCount
                                }
                                elseif($blocks.mLinesCount -gt $blocks.oLinesCount){
                                    $addCount += ($blocks.mLinesCount - $blocks.oLinesCount)
                                    $editCount += $blocks.oLinesCount
                                }
                                elseif($blocks.mLinesCount -lt $blocks.oLinesCount){
                                    $removeCount += ($blocks.oLinesCount - $blocks.mLinesCount)
                                    $editCount += $blocks.mLinesCount
                
                                }  
                            }
                        }
                    }    
                }
                foreach ($add in $commit_adds) {
                    $line_number = 0
                    $oId = $add."objectId"
                    $url_endpoint_object = "_apis/git/repositories/$repository_id/blobs/$oId ?api-version=6.0"
                    $url_object  = $url_base + $url_endpoint_object
                    $object  = Invoke-RestMethod -Uri "$url_object" -Method GET -ContentType "application/json" -headers $header 
                    $number = $object | Measure-Object -Line
                    $line_number = $number.Lines
                    $paths = $add.path
                    $splitting = $paths.Split('.')
                    if($line_number -eq 0){
                        if($splitting -contains 'json'){
                            $f = $object | convertto-json -depth 100
                            $number = $f | Measure-Object -Line
                            $line_number = $number.Lines    
                        }
                    }
                    $addCount += $line_number  
                }
            }
    }
    echo "pullID : $pullId "
    echo "addCount , removeCount , editCount : $addCount , $removeCount , $editCount"
    $itemChanges = @($editCount, $addCount, $removeCount)
    $pull.add($pullId,$itemChanges) 
    
    $pull_req_id = $p.pullRequestId
    $url_endpoint_workitems = "_apis/git/repositories/$repository_id/pullRequests/$pull_req_id/workitems?api-version=6.0"
    $url_workitems = $url_base + $url_endpoint_workitems

    $workitems_ids = Invoke-RestMethod -Uri "$url_workitems" -Method GET -ContentType "application/json" -headers $header  
        
    $ids = @()
    foreach($workitem in $workitems_ids.value){
        $workitem_id = $workitem.id 
        echo "Work Items: , $workitem_id"
        $ids += $workitem_id
            
    }
    $updates = @()
        
    foreach($i in $ids)
    {
        $url_endpoint_workitem = "_apis/wit/workitems?ids=$i&api-version=6.0"
        $url_workitem = $url_base + $url_endpoint_workitem
        $workitem_info = Invoke-RestMethod -Uri "$url_workitem" -Method GET -ContentType "application/json" -headers $header 
     
        If($workitem_info.value.fields.'System.WorkItemType' -eq 'Product Backlog Item')
        {   
            echo "...........................A PBI"
            $id = $workitem_info.value.id
            $url_endpoint_relations = "_apis/wit/workItems/$id/updates?api-version=6.0"
            $url_relations = $url_base + $url_endpoint_relations
            $relations = Invoke-RestMethod -Uri "$url_relations" -Method GET -ContentType "application/json" -headers $header  #| convertto-json -depth 100
            $updates += $relations
               
        }
        If($workitem_info.value.fields.'System.WorkItemType' -eq 'Task')
        {   
            echo "..........................A TASK"
            $id = $workitem_info.value.id
            $url_endpoint_relations = "_apis/wit/workItems/$id/updates?api-version=6.0"
            $url_relations = $url_base + $url_endpoint_relations
            $relations = Invoke-RestMethod -Uri "$url_relations" -Method GET -ContentType "application/json" -headers $header 
            If($relations.value.fields.'System.Parent'.'newValue'.count -ne 0){
                $updates += $relations
            } 
        }
    }
        
    foreach ($update in $updates) 
    {   
        if($update.value.fields.'System.Parent'.'newValue'.count -ne 0){
          
            $parent_id = $update.value.fields.'System.Parent'.'newValue'
            echo "Parent Id: , $parent_id "
            $url_endpoint_parent = "_apis/wit/workItems/$parent_id/updates?api-version=6.0"
            $url_parent = $url_base + $url_endpoint_parent
               
            $parent_info = Invoke-RestMethod -Uri "$url_parent" -Method GET -ContentType "application/json" -headers $header
               
        } 
            
            
            If($update.value.fields."System.WorkItemType".'newValue' -eq 'Task'){ 
                
                If($parent_info.value.fields.'System.WorkItemType'.'newValue' -eq 'Product Backlog Item') 
                { 
                    
                    $final_info_1 = $parent_info.value.workItemId[0]
                    $final_info_2 = $parent_info.value.fields.'System.Title'.'newValue'
                    $final_info_3 = $parent_info.value.fields.'System.IterationPath'.'newValue'
                    $final_info_4 = $pull.$pull_req_id[1]
                    $final_info_5 = $pull.$pull_req_id[2]
                    $final_info_6 = $pull.$pull_req_id[0]
                    $final_info_7 = @()
                    $final_info_8 = @()
                    $final_info_9 = @()
                   
                    $iter = 0
                    while($iter -le 10){
                       
                   
                        If($parent_info.value.fields.'System.WorkItemType'.'newValue' -eq 'Epic') 
                        { 
                            $final_info_7 += $parent_info.value.fields.'System.Id'.'newValue'
                            echo "....................A EPIC"
                            
                         }
                        elseif ($parent_info.value.fields.'System.WorkItemType'.'newValue' -eq 'Feature') {
                             
                            $final_info_8  += $parent_info.value.fields.'System.Id'.'newValue'
                            echo "....................A Feature"

                        }
                        elseif ($parent_info.value.fields.'System.WorkItemType'.'newValue' -eq 'Requirement') {
                              
                            $final_info_9  += $parent_info.value.fields.'System.Id'.'newValue'
                            echo "....................A Requirement"

                        }
                        if($parent_info.value.fields.'System.Parent'.'newValue'.count -ne 0){
                          $new_parent_id = $parent_info.value.fields.'System.Parent'.'newValue'
                          $url_endpoint_new_parent = "_apis/wit/workItems/$new_parent_id/updates?api-version=6.0"
                          $url_new_child = $url_base + $url_endpoint_new_parent
                            
                          try{
                            $parent_info = Invoke-RestMethod -Uri "$url_new_child" -Method GET -ContentType "application/json" -headers $header
                            $iter += 1

                          }catch{
                              break
                          }
                        }
                        else{
                            break
                        }
                        
                        
                    }
                    if(0 -eq $final_info_7.count)
                                {
                                    $final_info_7 = "-"
                                }
                    if(0 -eq $final_info_8.count)
                                {
                                    $final_info_8 = "-"
                                }
                    if(0 -eq $final_info_9.count)
                                {
                                    $final_info_9 = "-"
                                }

                 $final_info = @(
                      "$final_info_1,$final_info_2,$final_info_3,$final_info_4,$final_info_5,$final_info_6,$final_info_7,$final_info_8,$final_info_9"
                      ) 
                 echo $final_info
                 $final_info2.Add($final_info)  
                  
                    
                }
            }
            elseif($update.value.fields."System.WorkItemType".'newValue' -eq 'Product Backlog Item'){

                $iter = 0
                $final_info_4 = @()
                $final_info_5 = @()
                $final_info_6 = @()
                $final_info_1 = $update.value.workItemId[0]
                $final_info_2 = $update.value.fields.'System.Title'.'newValue'
                $final_info_3 = $update.value.fields.'System.IterationPath'.'newValue'
                $final_info_4 = $pull.$pull_req_id[1]
                $final_info_5 = $pull.$pull_req_id[2]
                $final_info_6 = $pull.$pull_req_id[0]
                $final_info_7 = @()
                $final_info_8 = @()
                $final_info_9 = @()
                
                if($null -ne $update.value.fields.'System.Parent'.'newValue'){
                    

                    while($iter -le 10){
                        
                   
                       If($parent_info.value.fields.'System.WorkItemType'.'newValue' -eq 'Epic') 
                       { 
                          $final_info_7 += $parent_info.value.fields.'System.Id'.'newValue'

                        }
                        elseif ($parent_info.value.fields.'System.WorkItemType'.'newValue' -eq 'Feature') {
                        
                           $final_info_8 += $parent_info.value.fields.'System.Id'.'newValue'

                        }
                        elseif ($parent_info.value.fields.'System.WorkItemType'.'newValue' -eq 'Requirement') {
                         
                           $final_info_9 = $parent_info.value.fields.'System.Id'.'newValue'
                        }
                         
                        if($parent_info.value.fields.'System.Parent'.'newValue'.count -ne 0)
                        {
                            $new_parent_id = $parent_info.value.fields.'System.Parent'.'newValue'
                            $url_endpoint_new_parent = "_apis/wit/workItems/$new_parent_id/updates?api-version=6.0"
                            $url_new_child = $url_base + $url_endpoint_new_parent
                            
                            try{
                                $parent_info = Invoke-RestMethod -Uri "$url_new_child" -Method GET -ContentType "application/json" -headers $header
                                $iter += 1
                            }catch
                            { 
                                break
                            }    
                        }
                        else{
                            break
                        }
                    }
                }
                    if($final_info_7.count -eq 0) 
                    {
                        $final_info_7 = "-"
                    }
                    if($final_info_8.count -eq 0)
                    {
                        $final_info_8 = "-"
                    }
                    if($final_info_9.count -eq 0)
                    {
                        $final_info_9 = "-"
                    }

                 $final_info = @(
                     "$final_info_1,$final_info_2,$final_info_3,$final_info_4,$final_info_5,$final_info_6,$final_info_7,$final_info_8,$final_info_9"
                     ) 
                 $final_info2.Add($final_info)   
            }       
 }
 if($s -eq 3){
     break
    }
$s+=1
}
$final_info2| foreach { Add-Content -Path  C:\Users\$computer_user_name\Desktop\final.csv -Value $PSItem } 
