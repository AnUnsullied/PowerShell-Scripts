function GetAnalyzedFollowing($followers, $following){
    Write-Host "Compiling users you follow but they don't follow you back..."
    foreach($userFollowing in $following){

        if ($userFollowing -in $followers){
            Write-Host ("You are following " + $userFollowing + " and they are following you back" ) -ForegroundColor Green
        }
        else{
            Write-Host ("You are following " + $userFollowing + " but they are not following you back" ) -ForegroundColor Red
        }

    }
}

function GetAnalyzedFollowers($followers, $following){
    Write-Host "Compiling users following you but you don't follow them back..."
    foreach($userFollower in $followers){

        if ($userFollower -in $following){
            Write-Host ($userFollower + " is following you and you are following them back" ) -ForegroundColor Green
        }
        else{
            Write-Host ($userFollower + " us following you but you are not following back" ) -ForegroundColor Red
        }

    }
}

$path = Split-Path -Parent $PSCommandPath

$followersFile = "$path\followers.json"
$jsonFollowerItem = Get-Content $followersFile | ConvertFrom-Json

$followingFile = "$path\following.json"
$jsonFollowingItem = Get-Content $followingFile | ConvertFrom-Json

$followers = @()
$following = @()


foreach($user in $jsonFollowerItem.relationships_followers){
    $followers += $user.string_list_data.value
}

foreach($user in $jsonFollowingItem.relationships_following){
    $following += $user.string_list_data.value
}



$userInput = ""

while($userInput -ne 'q'){
    Write-Host "If you want to find accounts you are following that does not follow you back, enter '1'." -BackgroundColor White -ForegroundColor Black
    Write-Host "If you want to find accounts following you but you do not follow back, enter '2'." -BackgroundColor White -ForegroundColor Black
    Write-Host "If you want to end this program, enter 'q'." -BackgroundColor White -ForegroundColor Black
    $userInput = Read-Host "Enter your input here"

    switch($userInput){
        "1"{
            GetAnalyzedFollowing -followers $followers -following $following
        }
        "2"{
            GetAnalyzedFollowers -followers $followers -following $following
        }
        "q"{
            break
        }
        default{
            Write-Host "User input unknown. Please try again. `n"
        }
    }

}
