Clear-Host

Import-Module ActiveDirectory

$i = 2

function Get-User($naa)
{
	try
	{
		$user = Get-ADUser $naa -Properties "office", "department", "displayname"
		return $user
	}
	catch
	{
		$user  = "Fail"
	}
}

function Get-Location($user)
{
	try
	{
		[string]$location = (($user.office).subString(1, 3))
		$proxyAdd = Get-ADGroup $location -properties "proxyaddresses"
		foreach ($obj in $proxyAdd.proxyaddresses)
		{
			if ($obj -like "smtp:C*")
			{
				$location = ($obj.substring(6, 3))
			}
		}
	}
	catch
	{
		$location = "No Location"
	}
	return $location
}

function Get-Information($naa)
{
	$user = Get-User($naa)
	if ($user -ne "Fail")
	{
		$location = Get-Location($user)
		if ($location -ne "No Location")
		{
			$cstCenter = ($location + "-" + (($user.department).substring(0, 5)))
			#write-Host($cstCenter)
			$excelWkSheet.Cells.Item($i, 5) = $cstCenter
			$excelWkSheet.Cells.Item($i, 3) = $user.displayname
		}
	}
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true

$excelWkBk = $excel.Workbooks.Open("C:\Users\naajuq\OneDrive - Coca-Cola Bottling Consolidated\Verizon Cleanup\June Invoice\June Invoice.xlsx")
$excelWkSheet = $excel.WorkSheets.Item("JuneInvoice")
$excelWkSheet.activate()


while (($excelWkSheet.Cells.Item($i, 1).text -ne "") -or ($excelWkSheet.Cells.Items($i, 1).text -ne $null))
{
	Write-Host($excelWkSheet.Cells.Item($i, 1).text)
	[string]$user1 = ($excelWkSheet.Cells.Item($i, 4).text).tostring().trim()
	[string]$user2 = ($excelWkSheet.Cells.Item($i, 7).text).tostring().trim()
	
	if (($user1 -ne "") -and ($user1 -ne $null))
	{
		Get-Information($user1)
	}
	elseif(($user2 -ne "") -and ($user2 -ne $null))
	{
		Get-Information($user2)
	}

	$i++
}

$excel.Quit

