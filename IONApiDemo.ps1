# determine demo API to call
$APIProgram = "MNS150MI"
$APITransaction = "SelUsers"

# read credentials from ionapi file
$Creds = Get-Content -Raw -Path "CredFile.ionapi" | ConvertFrom-Json

# get bearer token
$TokenUri = $Creds.pu + $Creds.ot
$Body = @{
	grant_type = 'password'
	username = $Creds.saak
	password = $Creds.sask
	client_id = $Creds.ci
	client_secret = $Creds.cs
	scope = ''
	redirect_uri = 'https://localhost/'
}
$AuthResult = Invoke-RestMethod -Method Post -Uri $TokenUri -Body $Body
$AccessToken = $AuthResult.access_token

# call MI
$Headers = @{
    Authorization="Bearer $AccessToken"
}
$MIUri = $Creds.iu + "/" + $Creds.ti + "/M3/m3api-rest/v2/execute/" + $APIProgram + "/" + $APITransaction
$Output = Invoke-RestMethod -Uri $MIUri -Headers $Headers

# output results
Write-Host "CalledUri: " + $MIUri
Write-Host "Result: " + $Output