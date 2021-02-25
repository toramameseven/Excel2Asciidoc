# fparam([string]$fullFileName)  
  
# $filePath = Join-Path $pwd $fullFileName  

# 最初に1度ロードすればよい
#[void][Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.Excel")

# 型名が長いのでいったん変数に入れる
#Set-Alias [xlDirection] [Microsoft.Office.Interop.Excel.XlDirection]

#[Microsoft.Office.Interop.Excel.XlDirection]::
# attach to excel file



$fullFileName = 'C:\Users\maru\Desktop\desk_temp\excel_powershell\vbaDeveloperx.xlam'
$bookName = [System.IO.Path]::GetFileName($fullFileName)
$srcPath = [System.IO.Path]::GetDirectoryName($fullFileName) + '\src\' + $bookName
New-Item $srcPath -ItemType Directory -Force | Out-Null


# エクセルアプリケーションを起動
$objExcel = New-Object -ComObject Excel.Application
 
# バックグランドで実行（$false）表示する場合は（$true）
$objExcel.Visible = $true
 
# ブックを新規作成して、そのオブジェクトを取得
$objBook  = $objExcel.Workbooks.Add()

#export $fullFileName
import $fullFileName

# ブックを保存
#$objBook.SaveAs($path_file_output)
 
# エクセルのアプリケーションを終了する
#$objExcel.Quit()





function export([string]$fullFileName){
  $Excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
  if ($null -eq $Excel) {
    Write-Output "i can not find excel objects"
    exit 
  }

  $VBProjects = $Excel.VBE.VBProjects
  if ($null -eq $VBProjects) {
    Write-Output "i can not find excel objects"
    exit 
  }

  $VBProjects | ForEach-Object {
    if ($_.FileName -eq $fullFileName){
        $_.VBComponents | ForEach-Object {
          [string]$extension = (ResolveModuleExtension $_.type)
          if ($extension -ne ""){
            [string]$modName =  $srcPath + "\" + $_.Name + $extension
            $modName
            $_.export($modName)
          }
        }
    }
  }
}

function import([string]$fullFileName){
  $Excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
  if ($null -eq $Excel) {
    Write-Output "i can not find excel objects"
    exit 
  }

  $VBProjects = $Excel.VBE.VBProjects
  if ($null -eq $VBProjects) {
    Write-Output "i can not find excel objects"
    exit 
  }

  $VBProject = ($VBProjects | Where-Object {$_.FileName -eq $fullFileName})

  $bookName = [System.IO.Path]::GetFileName($fullFileName)
  $srcPath = [System.IO.Path]::GetDirectoryName($fullFileName) + '\src\' + $bookName

  $modules = Get-ChildItem $srcPath | Where-Object {@(".cls",".frm",".bas").Contains([System.IO.Path]::GetExtension($_))}

  $modules

  $VBProject.VBComponents | ForEach-Object {
    [string]$extension = (ResolveModuleExtension $_.type)
    if ($extension -ne ""){
      $VBProject.VBComponents.remove($_)
    }
  }

  $modules | ForEach-Object {
    $_.FullName
    $VBProject.VBComponents.import($_.FullName)
  } 

}



function ResolveModuleExtension([int]$moduleType) {
  switch ($moduleType) {
    1 { ".bas" }
    2 { ".cls" }
    3 { ".frm" }
    Default {""}
  }
}

  

