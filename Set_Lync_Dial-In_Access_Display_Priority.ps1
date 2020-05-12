<##########################################################################  

Script Name:		
Change Dial-In Access Number Order

Contact Info: 
Name	  		C. Anthony Caragol 
LinkedIn  		http://www.linkedin.com/pub/anthony-caragol/6/48/488

Description:		
By request of a user, change the ordering of dial-in access numbers in a 
meeting request.

Notes:			
This must be done per-region.  Users may need to wait a few minutes and reload
Outlook before the change is seen.

Please excuse the sloppy coding for now, I don't use a
development environment, IDE or ISE.  I use notepad, not
even Notepad++, just notepad.  I am not a developer, just
an enthusiast.

Version:
1.000	First Draft


##########################################################################>  

#Run this when the form is resized
$CAC_FormSizeChanged = { 
		$dataGridView.Columns[0].Width = ($ObjForm.Width - 83) / 7.42
		$dataGridView.Columns[1].Width = ($ObjForm.Width - 83) / 2.79
		$dataGridView.Columns[2].Width = ($ObjForm.Width - 83) / 5.57
		$dataGridView.Columns[3].Width = ($ObjForm.Width - 83) / 4.45
		$dataGridView.Columns[4].Width = ($ObjForm.Width - 83) / 9.95
} 

#Run this when the combo box selection changes
$OnSelect_RegionDropDown = { 
	Refresh_DataGridView
}

#Run this to refresh data in the datagridview
function Refresh_DataGridView {
	$selectedregion = $RegionDropDown.SelectedItem.tostring()
	$DataGridView.Rows.Clear()
	$getdata= Get-CsDialInConferencingAccessNumber -region "$selectedregion" -WarningAction SilentlyContinue| select *

	foreach ($Row in $getdata) {
		$dataGridView.Rows.Add($selectedregion, $Row.DisplayName, $Row.DisplayNumber, $Row.PrimaryURI, $Row.Priority)
	}
}

#Run when Move To Top is clicked
function OnClick_MoveToTop  { 
	$DataGridView.SelectedRows[0].Cells[3].Value
	Set-CsDialInConferencingAccessNumber -Identity $DataGridView.SelectedRows[0].Cells[3].Value -Priority 0 -ReorderedRegion $DataGridView.SelectedRows[0].Cells[0].Value
	Refresh_DataGridView
	$datagridview.Rows[0].Selected = $true
}

#Run when Move to Bottom is clicked
function OnClick_MoveToBottom  {
	if (($DataGridView.SelectedRows[0].Index + 2) -lt $DataGridView.RowCount) {
		$newrow= $DataGridView.RowCount - 2
		Set-CsDialInConferencingAccessNumber -Identity $DataGridView.SelectedRows[0].Cells[3].Value -Priority 999 -ReorderedRegion $DataGridView.SelectedRows[0].Cells[0].Value
		Refresh_DataGridView
		$datagridview.Rows[$newrow].Selected = $true
	}
}

#Run when Move Up is clicked
function OnClick_MoveUp  { 
	if ([int]$DataGridView.SelectedRows[0].Cells[4].Value -gt 0) {
		$newrow= $DataGridView.SelectedRows[0].Index - 1
		$NewPriority = [int]$DataGridView.SelectedRows[0].Cells[4].Value - 1
		Set-CsDialInConferencingAccessNumber -Identity $DataGridView.SelectedRows[0].Cells[3].Value -Priority $NewPriority -ReorderedRegion $DataGridView.SelectedRows[0].Cells[0].Value
		Refresh_DataGridView
		$datagridview.Rows[$newrow].Selected = $true
	}
}

#Run when Move Down is clicked
function OnClick_MoveDown  {
	if (($DataGridView.SelectedRows[0].Index + 2) -lt $DataGridView.RowCount) {
		$newrow= $DataGridView.SelectedRows[0].Index + 1
		$NewPriority = [int]$DataGridView.SelectedRows[0].Cells[4].Value + 1
		Set-CsDialInConferencingAccessNumber -Identity $DataGridView.SelectedRows[0].Cells[3].Value -Priority $NewPriority -ReorderedRegion $DataGridView.SelectedRows[0].Cells[0].Value
		Refresh_DataGridView
		$datagridview.Rows[$newrow].Selected = $true
	}
}

#Start script and load form

#I know we don't specifically need this, but in my head it might load the Lync module faster than if we make it guess.
write-host Please be patient while we import the Lync command set.
write-host Users may need to wait a few minutes and reload Outlook before the change is seen.
import-module lync

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

$objForm = New-Object System.Windows.Forms.Form 
$objForm.Text = "Change Dial-In Access Number Order"
$objForm.Size = New-Object System.Drawing.Size(640,600) 
$objForm.StartPosition = "CenterScreen"
$ObjForm.Add_SizeChanged($CAC_FormSizeChanged) 
$objForm.KeyPreview = $True

$MoveUpButton = New-Object System.Windows.Forms.Button
$MoveUpButton.Location = New-Object System.Drawing.Size(110,525)
$MoveUpButton.Size = New-Object System.Drawing.Size(100,25)
$MoveUpButton.Text = "Move Up"
$MoveUpButton.Add_Click({ OnClick_MoveUp })
$MoveUpButton.Anchor = 'Bottom, Right'
$objForm.Controls.Add($MoveUpButton)

$MoveDownButton = New-Object System.Windows.Forms.Button
$MoveDownButton.Location = New-Object System.Drawing.Size(210,525)
$MoveDownButton.Size = New-Object System.Drawing.Size(100,25)
$MoveDownButton.Text = "Move Down"
$MoveDownButton.Add_Click({ OnClick_MoveDown })
$MoveDownButton.Anchor = 'Bottom, Right'
$objForm.Controls.Add($MoveDownButton)

$MoveToTopButton = New-Object System.Windows.Forms.Button
$MoveToTopButton.Location = New-Object System.Drawing.Size(310,525)
$MoveToTopButton.Size = New-Object System.Drawing.Size(100,25)
$MoveToTopButton.Text = "Move To Top"
$MoveToTopButton.Add_Click({ OnClick_MoveToTop })
$MoveToTopButton.Anchor = 'Bottom, Right'
$objForm.Controls.Add($MoveToTopButton)

$MoveToBottomButton = New-Object System.Windows.Forms.Button
$MoveToBottomButton.Location = New-Object System.Drawing.Size(410,525)
$MoveToBottomButton.Size = New-Object System.Drawing.Size(100,25)
$MoveToBottomButton.Text = "Move To Bottom"
$MoveToBottomButton.Add_Click({ OnClick_MoveToBottom })
$MoveToBottomButton.Anchor = 'Bottom, Right'
$objForm.Controls.Add($MoveToBottomButton)

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Size(510,525)
$CancelButton.Size = New-Object System.Drawing.Size(100,25)
$CancelButton.Text = "Quit"
$CancelButton.Add_Click({$objForm.Close()})
$CancelButton.Anchor = 'Bottom, Right'
$objForm.Controls.Add($CancelButton)

$dataGridView = New-Object System.Windows.Forms.DataGridView
$dataGridView.Location = New-Object System.Drawing.Size(10,60) 
$dataGridView.Size = New-Object System.Drawing.Size(600,400) 
$dataGridView.SelectionMode = 'FullRowSelect'
$dataGridView.Multiselect=$false
$dataGridView.Anchor = 'Top, Bottom, Left, Right'
$dataGridView.ColumnCount = 5
$dataGridView.Columns[0].Width = 75
$dataGridView.Columns[1].Width = 200
$dataGridView.Columns[2].Width = 100
$dataGridView.Columns[3].Width = 125
$dataGridView.Columns[4].Width = 57
$dataGridView.Columns[0].Name = "Region"
$dataGridView.Columns[1].Name = "Display Name"	
$dataGridView.Columns[2].Name = "Display Number"
$dataGridView.Columns[3].Name = "URI"
$dataGridView.Columns[4].Name = "Priority"
$objForm.Controls.Add($dataGridView) 

$WorkFlowLabel = New-Object System.Windows.Forms.Label
$WorkFlowLabel.Location = New-Object System.Drawing.Size(10,20) 
$WorkFlowLabel.Size = New-Object System.Drawing.Size(230,20) 
$WorkFlowLabel.Text = "Please select an existing dialin region:"
$objForm.Controls.Add($WorkFlowLabel) 

$RegionDropDown = new-object System.Windows.Forms.ComboBox
$RegionDropDown.Location = new-object System.Drawing.Size(240,20) 
$RegionDropDown.Size = new-object System.Drawing.Size(370,20) 
$RegionDropDown.add_SelectedIndexChanged($OnSelect_RegionDropDown) 
$RegionDropDown.Anchor = 'Top, Left, Right'

$getdata= Get-CsDialInConferencingAccessNumber
$RegionArray = @()
foreach ($Row in $getdata) {
	foreach ($regionentry in $row.regions) {
		if ($RegionArray -notcontains $regionentry) {  $RegionArray += $regionentry  }
	}
}

for ($i=0; $i -lt $RegionArray.length; $i++) {
	[void]$RegionDropDown.Items.Add($RegionArray[$i])
}
$objForm.Controls.Add($RegionDropDown)
$RegionDropDown.SelectedIndex = 0

$objForm.Topmost = $True

$objForm.Add_Shown({$objForm.Activate()})
[void] $objForm.ShowDialog()


