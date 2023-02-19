[Reflection.Assembly]::LoadWithPartialName("System.Drawing")

function Screenshot([Drawing.Rectangle]$bounds, $path) {
   $bmp = New-Object Drawing.Bitmap $bounds.width, $bounds.height
   $graphics = [Drawing.Graphics]::FromImage($bmp)

   $graphics.CopyFromScreen($bounds.Location, [Drawing.Point]::Empty, $bounds.size)

   $bmp.Save($path)

   $graphics.Dispose()
   $bmp.Dispose()
}

$data = Import-Excel 'D:\temp\PROVA AUTOMATION.xlsx' -Sheet websites

$data | ForEach-Object {
   $website = $_
   
   $url=$website.Hostnames.Split(',')[0]
   
   Write-Host $website.Name $url
   
   $IE=new-object -com internetexplorer.application
   $IE.visible=$true
   $IE.FullScreen=$false
   $IE.ToolBar = $false
   $IE.StatusBar = $false
   $IE.MenuBar = $false
   $IE.AddressBar = $true
   $IE.Resizable = $true
   $IE.Top = 0
   $IE.Left = 577
   $IE.Width = 1024
   $IE.Height = 747
   
   $IE.navigate2( $url )
   
   $i=0
   While ( $IE.busy -eq $true ) { 
      Start-Sleep -s 1
      $i = $i + 1
      if ( $i -ge 20 ) { break }
   }

   $bounds = [Drawing.Rectangle]::FromLTRB($IE.Left, $IE.Top, $IE.Left + $IE.Width, $IE.Top + $IE.Height)
   
   $filename = "D:\temp\urlshots\"+ ($website.Name) +".png"
   
   Screenshot $bounds $filename
   
   $IE.Quit()
}
