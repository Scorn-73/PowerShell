Import-Module Microsoft.Graph.Identity.SignIns
Connect-MgGraph -Scopes "IdentityRiskyUser.ReadWrite.All"
$userID = "INPUTID"
$params = @{
	userIds = @(
	"$userID")
}

Invoke-MgDismissRiskyUser -BodyParameter $params

Remove-Module Microsoft.Graph.Identity.SignIns
