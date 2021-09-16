Update-TypeData -AppendPath ($PSScriptRoot + "\comObject.types.ps1xml")
enum openMode
{
    msiOpenDatabaseModeReadOnly = 0
    msiOpenDatabaseModeTransact = 1
    msiOpenDatabaseModeDirect = 2
    msiOpenDatabaseModeCreate = 3
    msiOpenDatabaseModeCreateDirect = 4
}

enum ViewModify 
{
    msiViewModifyInsert         = 1
    msiViewModifyUpdate         = 2
    msiViewModifyAssign         = 3
    msiViewModifyReplace        = 4
    msiViewModifyDelete         = 6
}

enum InstallerUILevel
{
    msiUILevelNone = 2
}

enum SessionMode
{
    msiRunModeSourceShortNames = 9
}

enum FileAttributes
{
    msidbFileAttributesNoncompressed = 8192
    msidbFileAttributesCompressed = 16384
}
function Release-Ref ($ref) {

    [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$ref) | out-null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
function Get-InstallProperties {
    [CmdletBinding()]
    param (
        # windowsinstaller object
        [Parameter(Mandatory=$true,Position=1)]
        [System.__ComObject]$WindowsInstaller
    )
    
    begin {
        $Products = $wi.GetParamProperty('ProductsEx','','',7)
        $AllOuts =@()
    }
    
    process {
        foreach($prod in $Products){
            $ut = New-Object PSObject
#            $prod = Add-ComMembers $prod 
            Add-Member -InputObject $ut -Name 'ProductCode' -Value $prod.GetProperty('ProductCode') -MemberType NoteProperty
            $InstallProperties = @('PackageName','Transforms','Language','ProductName','AssignmentType','PackageCode','Version','ProductIcon','InstanceType','AuthorizedLUAApp','InstalledProductName','VersionString','InstallDate','InstallLocation','InstallSource',
            'Publisher','LocalPackage','VersionMinor','VersionMajor','ProductID','RegCompany','RegOwner','MediaPackagePath')
            foreach ($ip in $InstallProperties){
                try{
                    $installPropertyValue = $prod.GetParamProperty('InstallProperty', $ip)
                    if (!$installPropertyValue -eq ""){
                       Add-Member -InputObject $ut -Name $ip -Value $installPropertyValue -MemberType NoteProperty
                        
                    }
                }
                catch{
                }
            }
            $AllOuts += $ut
        }
    }
    end {
        $AllOuts
    }
}

function IsTable{
    param( 
        [System.__ComObject]$SourceMSIDatabase, 
        [string]$sTableName 
        )
	$TableView = $SourceMSIDatabase.InvokeMethod('OpenView',( "SELECT * FROM ``_Tables`` WHERE ``Name`` = '" + $sTableName + "'"))
    $TableView.InvokeMethod('Execute') | Out-Null
    
	return ($null -ne $TableView.InvokeMethod('Fetch'))
}
function NewKeyName(){
    param (
        [String]$oldName,
        [String]$msiName
    )
    $rndChar = -join ((65..90) + (97..122) | Get-Random -Count 2 | ForEach-Object {[char]$_})
    if(($oldName.Length + $msiName.Length) -gt 70){
        $start = $oldName.Length - 70+ $msiName.Length
        $len = $oldName.Length -$start
        $oldName.Substring(($start ),($len )) + $rndChar+ $msiName
    }
    else{
        $oldName + $rndChar + $msiName
    }

}
function Get-ComponentColumnNo {
    param (
        [string]$sTableName,
        [System.__ComObject]$db
    )
    [int]$colNo = 0
    #fix the componentname
    if($sTableName -eq 'Component'){
        $sSQL = "SELECT ``Number`` FROM ``_Columns`` WHERE ``Table``='" +$sTableName +"' AND ``Name``='Component'"
    }
    else {
        $sSQL = "SELECT ``Number`` FROM ``_Columns`` WHERE ``Table``='" +$sTableName +"' AND ``Name``='Component_'"
    }
    $ColumnView = $db.InvokeMethod('OpenView',($sSQL))
    $ColumnView.InvokeMethod('Execute') |Out-Null
    $ColumnRecord = $ColumnView.InvokeMethod('Fetch')
    if ($null -ne $ColumnRecord){
        $colNo = $ColumnRecord.GetParamProperty('StringData',1)
    }
    $colNo
}
function Get-FileColumnNo {
    param (
        [string]$sTableName,
        [System.__ComObject]$db
    )
    [int]$colNo = 0
    #fix the componentname
    if($sTableName -eq 'Component'){
        $sSQL = "SELECT ``Number`` FROM ``_Columns`` WHERE ``Table``='" +$sTableName +"' AND ``Name``='KeyPath'"
    }
    else {
        $sSQL = "SELECT ``Number`` FROM ``_Columns`` WHERE ``Table``='" +$sTableName +"' AND ``Name``='File_'"
    }
    $ColumnView = $db.InvokeMethod('OpenView',($sSQL))
    $ColumnView.InvokeMethod('Execute') |Out-Null
    $ColumnRecord = $ColumnView.InvokeMethod('Fetch')
    if ($null -ne $ColumnRecord){
        $colNo = $ColumnRecord.GetParamProperty('StringData',1)
    }
    $colNo
}
function MSICopyData {
    param (
        [System.__ComObject]$SourceMSIDatabase,
        [System.__ComObject]$DestDatabase,
        [string]$sTableName,
        [string]$msiName,
        [hashtable]$fileRenameList,
        [hashtable]$componentRenameList,
        [hashtable]$regRenameList
    )
    
    
	If ((IsTable $SourceMSIDatabase $sTableName) -ne $true){return}

    $SourceView = $SourceMSIDatabase.InvokeMethod('OpenView',("SELECT * FROM ``" + $sTableName  + "``" ))
	$SourceView.InvokeMethod('Execute')
	$RecordD = $SourceView.InvokeMethod('Fetch')
	If (!$RecordD -is [__ComObject]){return} # no data to copy

    #'create a table if necc
	If ((IsTable $DestDatabase $sTableName) -eq $false){
        MSICopyTableDefinitions $SourceMSIDatabase $DestDatabase $sTableName
    }
	#' get a view for putting data into
	$DestView = $DestDatabase.InvokeMethod('OpenView', ("SELECT * FROM ``" + $sTableName + "``"))
	$DestView.InvokeMethod('Execute')

    $compColNo = Get-ComponentColumnNo $sTableName $SourceMSIDatabase
    $fileColNo = Get-FileColumnNo $sTableName $SourceMSIDatabase


    while ($RecordD -is [__ComObject]) {#	'get remaining data
        #uniquify
        if($compColNo -ne 0){
            $oldName = $RecordD.GetParamProperty('StringData',$compColNo)
            if (($null -ne $componentRenameList) -and $componentRenameList.ContainsKey($oldName)){
                $RecordD.SetParamProperty('StringData',$compColNo,($componentRenameList[$oldName]))
            }
            else{
                $newname = NewKeyName $RecordD.GetParamProperty('StringData',$compColNo) $msiName
                $componentRenameList[$oldName] = $newname
                $RecordD.SetParamProperty('StringData',$compColNo,($newname))
            }
        }
        if($fileColNo -ne 0){   #check to see if the filename has been renamed
            $oldName = $RecordD.GetParamProperty('StringData',$fileColNo)
            if (($null -ne $fileRenameList) -and $fileRenameList.ContainsKey($oldName)){
                $RecordD.SetParamProperty('StringData',$fileColNo,($fileRenameList[$oldName]))
            }
        }
        if($sTableName -eq 'Component'){ #check to see if the keypath has been renamed
            if($RecordD.GetParamProperty('IntegerData',4) -bor 4){ #test for registry keypath
                $oldname = $RecordD.GetParamProperty('StringData',6)
                if ($regRenameList.ContainsKey($oldName)){
                    $RecordD.SetParamProperty('StringData',6,($regRenameList[$oldName]))
                }
            }
        }
    
        try{
            $DestView.InvokeMethod('Modify', [ViewModify]::msiViewModifyInsert, $RecordD)
        }
        catch{<#
            $expectedTablesWithIssues = @('Registry','File','Upgrade','ModuleComponents','ModuleSignature','Signature','CheckBox','Icon','Binary','CustomAction',
            'ControlCondition','Control','Dialog','ControlEvent','RadioButton','TextStyle','CompLocator','Error','DrLocator','PatchPackage',
            'UIText','ActionText','EventMapping','RegLocator','AppSearch','LaunchCondition','_Validation','AdminUISequence','AdminExecuteSequence','AdvtExecuteSequence',
            'Directory','InstallExecuteSequence','InstallUISequence','Property','Media','Feature')
            if(!$expectedTablesWithIssues.contains($sTableName)){
                "ERROR Table: " + $sTableName + " 1st Field: " + $RecordD.GetParamProperty('StringData',1)
            }#>

            if($sTableName -eq 'File'){
                $oldName = $RecordD.GetParamProperty('StringData',1)
                $newname = NewKeyName $oldName $msiName
                $fileRenameList[$oldName]=$newname
                $RecordD.SetParamProperty('StringData',1,$newname)
                $DestView.InvokeMethod('Modify', [ViewModify]::msiViewModifyInsert, $RecordD)

            }
            if($sTableName -eq 'Registry'){
                $oldName = $RecordD.GetParamProperty('StringData',1)
                $newname = NewKeyName $oldName $msiName
                $regRenameList[$oldName]=$newname
                $RecordD.SetParamProperty('StringData',1,$newname)
                try{$DestView.InvokeMethod('Modify', [ViewModify]::msiViewModifyInsert, $RecordD)}
                catch{write-host "wtf oldname"+ $oldname +" newname "+ $newname }

            }
        }
		$RecordD = $SourceView.InvokeMethod('Fetch')
    }
} 

function MSICopyTableDefinitions {
    param (
        [System.__ComObject]$SourceMSIDatabase,
        [System.__ComObject]$DestDatabase,
        [string]$sTableName
    )    
    If (((IsTable $SourceMSIDatabase $sTableName) -eq $false) -or ((IsTable $DestDatabase $sTableName) -eq $true)){return}
    

    $sSQL = "CREATE TABLE ``" + $sTableName + "`` ( "

	$ColumnView = $SourceMSIDatabase.InvokeMethod('OpenView', ( 
		"SELECT ``Name``  " +
		"FROM ``_Columns`` " +
		"WHERE ``Table`` = '" + $sTableName + "' " +
        "ORDER BY ``Number``"))
        
	$ColumnView.InvokeMethod('Execute')
	$ColumnRecord = $ColumnView.InvokeMethod('Fetch')

	while ( $ColumnRecord -is [__ComObject]){
		$sColumnName = $ColumnRecord.GetParamProperty('StringData', 1)

#	'//get the columns
		$TableColumnNameView = $SourceMSIDatabase.InvokeMethod('OpenView',( 
			"SELECT ``" + $sColumnName + "`` " +
			"FROM ``" + $sTableName + "``"))
		$TableColumnNameView.InvokeMethod('Execute')

#	'//get their definitions
		$ColumnInfoRecord = $TableColumnNameView.GetParamProperty('ColumnInfo',(1))

#	'//write out the SQL to make the column
		$sSQL = $sSQL + "``" + $sColumnName + "`` " + (TranslateColumnDescriptor $ColumnInfoRecord.GetParamProperty('StringData',(1))) + ", "

		$ColumnRecord = $ColumnView.InvokeMethod('Fetch')
    }

#	'remove last comma and space
	$sSQL = $sSQL.TrimEnd(', ')

#	'get the list of primary keys
#	Dim sPrimaryKey
	$PrimaryKeysRecord = $SourceMSIDatabase.InvokeMethod('PrimaryKeys',$sTableName)
	for ($X = 1; $X -le $PrimaryKeysRecord.GetParamProperty('FieldCount'); $X++){
		$sPrimaryKey = $sPrimaryKey + "``" + $PrimaryKeysRecord.GetParamProperty('StringData',$X) + "``, "
    }

#	'remove last comma and space
	$sPrimaryKey = $sPrimaryKey.TrimEnd(', ')

	$sSQL = $sSQL + " PRIMARY KEY " + $sPrimaryKey + ")"

	$ColumnView = $DestDatabase.InvokeMethod('OpenView',$sSQL)	#'run the generated sql

	$ColumnView.InvokeMethod('Execute')
}

function TranslateColumnDescriptor {
    [OutputType([string])]
    param (
        [string]$sColumnInfo
    )
    
	$sSize = $sColumnInfo.Substring(1)  #Mid(sColumnInfo, 2, Len(sColumnInfo) -1)
	$cDataType = $sColumnInfo.Substring(0,1) #Left(sColumnInfo,1)  '//grab the column info, re "column definition format" in msi.chm

	switch ($cDataType){

        {($_ -in "i", "I")}{
            if ($sSize -eq "2"){$outString = "SHORT"}
            else{$outString = "LONG"}
        }
        {($_ -in "s", "S","l", "L")}{
        	if ($sSize -eq "0"){$outString = "LONGCHAR"}                
			else {$outString = "CHAR(" + $sSize + ")"}}
        {($_ -in "v", "V")}{$outString = "OBJECT"}
    }	

	if ($cDatatype.ToLower() -ceq $cDataType){
		$outString = $outString + " NOT NULL"	#'//not nullable
    }

	if ($cDataType.ToLower() -eq "l"){
		$outString = $outString + " LOCALIZABLE"	#'// must come at end
    }
    return $outString
}

class CabPacker {
    CabPacker($destDatabase,$windowsInstaller){
        $this.destDatabase = $destDatabase
        $this.windowsInstaller = $windowsInstaller
        $this.installedComponents =$this.windowsinstaller.GetProperty('Components')
    }
    [int]$previousDiskID = 0
    [int]$LastSequence=1
    [System.__ComObject]$destDatabase
    [System.__ComObject]$windowsInstaller
    [string]$cabName = "setup"
    [string]$DDFFile = $PSScriptRoot + "\setup.ddf"
    [psobject]$installedComponents 
    
    [void]InitDDFFile() {
    #' Create DDF file and write header properties
        ";This temporary intermediate file is safe to delete" | Out-File -FilePath $this.DDFFILE -Encoding ascii
        ".Set Compress=ON" | Out-File -FilePath $this.DDFFILE -append -Encoding ascii
        ".Set Cabinet=ON" | Out-File -FilePath $this.DDFFILE -append -Encoding ascii
        ".Set MaxErrors=1" | Out-File -FilePath $this.DDFFILE -append -Encoding ascii
        ".Set DiskDirectoryTemplate=."  | Out-File -FilePath $this.DDFFILE -append -Encoding ascii
        ".Set DiskDirectory1=$($PSScriptRoot)"  | Out-File -FilePath $this.DDFFILE -append -Encoding ascii
        ".Set ReservePerCabinetSize=8"	| Out-File -FilePath $this.DDFFILE -append -Encoding ascii
        ".Set CabinetName1=" + $this.cabName + ".cab" | Out-File -FilePath $this.DDFFILE -append -Encoding ascii
        ".Set CabinetNameTemplate=" + $this.cabName + "*.CAB" | Out-File -FilePath $this.DDFFILE -append -Encoding ascii
        ".Set MaxDiskSize=CDROM" 	| Out-File -FilePath $this.DDFFILE -append -Encoding ascii
        ".Set CabinetFileCountThreshold=16000" 	| Out-File -FilePath $this.DDFFILE -append -Encoding ascii
        ".Set CompressionType=LZX"	| Out-File -FilePath $this.DDFFILE -append -Encoding ascii
        ".Set RptFileName=" + $this.cabName + ".RPT"	| Out-File -FilePath $this.DDFFILE -append -Encoding ascii
        ".Set InfFileName=" + $this.cabName + ".cab.Log"	| Out-File -FilePath $this.DDFFILE -append -Encoding ascii
        ".Set InfSectionOrder=CF" | Out-File -FilePath $this.DDFFILE -append -Encoding ascii
        ".Set InfFileLineFormat=(*cab#*) *file#*: *file*, *ver*, *Date*, *lang*" | Out-File -FilePath $this.DDFFILE -append -Encoding ascii
        ".Set InfCabinetLineFormat=(*cab#*) *cabfile* " | Out-File -FilePath $this.DDFFILE -append -Encoding ascii
        ".Set InfHeader=" | Out-File -FilePath $this.DDFFILE -append -Encoding ascii
        ".Set InfFooter=" | Out-File -FilePath $this.DDFFILE -append -Encoding ascii
        ".InfWrite (*cab#*) *file#*: *file*, *ver*, *Date*, *lang* " | Out-File -FilePath $this.DDFFILE -append -Encoding ascii
        ".InfWrite The file# goes into the lastdisk column of the media table" | Out-File -FilePath $this.DDFFILE -append -Encoding ascii
        ".Set UniqueFiles=OFF"	| Out-File -FilePath $this.DDFFILE -append -Encoding ascii
        ".Set GenerateInf=ON" | Out-File -FilePath $this.DDFFILE -append -Encoding ascii
    }

    [void]AddFiles([System.__ComObject]$session,$fileRenameList){
        $session.InvokeMethod('DoAction','CostInitialize') |Out-Null
        $session.InvokeMethod('DoAction','CostFinalize')    | Out-Null
        $db = $session.GetProperty('Database')
        $view = $db.InvokeMethod('OpenView',"SELECT File,FileName,Directory_,File.Attributes,ComponentId,Component,Version,FileSize,Condition,KeyPath" +
        " FROM File,Component WHERE Component_=Component ORDER BY File.Sequence")
        $View.InvokeMethod('Execute')
    # Fetch each file and request the source path, then verify the source path
        $productCode = $session.GetParamProperty('Property','ProductCode')
        $previousSearchedFolders=@()
        
        Do{
            $record = $view.InvokeMethod('Fetch')
            if ($null -eq $record){break}

            $ComponentID = $record.GetParamProperty('StringData',5)
            $componentName = $record.GetParamProperty('StringData',6)
            if (!$this.installedComponents.contains($ComponentID)){
                Write-Host "Component $ComponentID is not installed anywhere"
                continue
            }
            $fileKey    = $record.GetParamProperty('StringData',1)
            if($fileRenameList.ContainsKey($fileKey)){$fileKey=$fileRenameList[$fileKey]}
            $fileName   = $record.GetParamProperty('StringData',2)
            $folder     = $record.GetParamProperty('StringData',3)
            [int]$attributes = $record.GetParamProperty('IntegerData',4)
            $version    = $record.GetParamProperty('StringData',7)
            [int]$size = $record.GetParamProperty('IntegerData',8)
            $condition    = $record.GetParamProperty('StringData',9)
            $keyPath    = $record.GetParamProperty('StringData',10)

            if("" -ne $condition){
                if($session.InvokeMethod('EvaluateCondition',$condition) -eq 0){
                    Write-Host "Component:$ComponentID File:$fileName is not installed by this package"
                    continue
                }
            }

            $destView = $this.destDatabase.InvokeMethod('OpenView', 
                "UPDATE File SET ``Sequence``='" + $this.LastSequence +
                    "',``Attributes``='" + ($attributes -bor [FileAttributes]::msidbFileAttributesCompressed) + 
                    "' WHERE ``File``='"+$fileKey+"'")
            $destView.InvokeMethod('Execute')

            if($fileName.contains('|')){$fileName=$fileName.split('|')[1]}

            $sourcePath = $Session.GetParamProperty('TargetPath',$folder) + $fileName   #try the session's TargetPath
            if(!([System.IO.File]::Exists($sourcePath))){    #there are reasons why this may happen such as installing to GAC
                #ask the installer object for the component path                
                                
                $sourcePath = $this.windowsInstaller.GetParamProperty('ComponentPath', $productCode, $ComponentID)  #try the installer object's Component Path
                
                if ($null -ne $sourcePath){ 
                    $sourcePath =  ( $sourcePath | Split-Path -Parent) +"\" +$fileName
                }
                if(!([System.IO.File]::Exists($sourcePath))){
                    $AssemblyView = $db.InvokeMethod('OpenView', "SELECT Attributes FROM ``MsiAssembly`` WHERE ``Component_``='" + $componentName + "'")
                    $AssemblyView.InvokeMethod('Execute')
                    $AssemblyRecord = $AssemblyView.InvokeMethod('Fetch')
                    if ($null -ne $AssemblyRecord){
                        $sourcePath = $this.ProvideAssembly($db,$componentName, $AssemblyRecord.GetParamProperty('IntegerData',1))
                        if(!([System.IO.File]::Exists($sourcePath))){
                            ";damn the assembly is not here: " + $sourcePath | Out-File -FilePath $this.DDFFILE -append -Encoding ascii
                            continue        
                        }
                        $ext = [system.io.path]::GetExtension($fileName)
                        if(([system.io.path]::GetExtension($sourcePath) -eq '.manifest') -and ($ext -eq '.cat')){
                            $sourcePath = "$($sourcePath.Substring(0,$sourcePath.Length -8))cat"

                            if((!([System.IO.File]::Exists($sourcePath))) -or ((Get-item $sourcePath).Length -ne $size)){
                                ";nice try: " + $sourcePath | Out-File -FilePath $this.DDFFILE -append -Encoding ascii
                                continue
                            }
                        }
                        elseif(($ext -ne [system.io.path]::GetExtension($sourcePath)) -and ($fileKey -ne $keyPath) ){
                            $sourcePath = (Join-Path -path (get-item env:\windir).value -ChildPath "Winsxs\Manifests\") + (split-path(split-path $sourcePath -Parent) -leaf) + $ext
                            
                            if((!([System.IO.File]::Exists($sourcePath))) -or ((Get-item $sourcePath).Length -ne $size)){
                                ";damn the manifest or cat is not here: " + $sourcePath | Out-File -FilePath $this.DDFFILE -append -Encoding ascii
                                continue
                            }
                        }
                    }
                    else{
                        ## ok not an assembly or anything else ... bruteforce time
                        $hashView = $db.InvokeMethod('OpenView',"SELECT HashPart1,HashPart2,HashPart3,HashPart4 FROM MsiFileHash WHERE File_='$fileKey'")
                        $hashView.InvokeMethod('Execute')
                        $hashRecord = $hashView.InvokeMethod('Fetch')
                        $hashData = @{}
                        if($null -ne $hashRecord){
                            for($x=1; $x -le 4; $x++){
                                $hashData[$x] = $hashRecord.GetParamProperty('IntegerData',$x)
                            }
                        }
            
                        #add any folder found to a list for later use
                        $searchCmd = "(Get-Item {0} | Where-Object{{ (`$_.Length -eq $size)"
                        if("" -ne $version){ $searchCmd+=" -and (`$_.VersionInfo.FileVersionRaw -eq '$version')"}
                        $searchCmd+= "}}).FullName"
                        
                        $sourcePath = $null
                        foreach($searchFolder in $previousSearchedFolders){
                            if(!([System.IO.File]::Exists("$searchFolder\$fileName"))){continue}
                            $result = Invoke-Expression ($searchCmd -f """$searchFolder\$fileName""")
                            if($null -ne $result){
                                if($null -ne $hashRecord){
                                    foreach ($r in $result){
                                        $hashRecord = $this.windowsInstaller.InvokeMethod('FileHash', $r ,0)
                                        $match = $true
                                        for($x=1; $x -le 4; $x++){
                                            if ($hashData[$x] -ne $hashRecord.GetParamProperty('IntegerData',$x)){
                                                $match = $false
                                            }
                                        }
                                        if($match){
                                            $sourcePath = $r
                                            break
                                        }
                                    }
                                }
                                else{
                                    if($result -is [array]){
                                        $sourcePath = $result[1]
                                    }else{
                                        $sourcePath = $result
                                    }
                                }
                            }
                        }
                        if($null -eq $sourcePath){
                            $searchCmd = "(Get-ChildItem ""C:\$fileName"" -recurse -exclude 'C:\Windows\CSC\' | Where-Object{{ (`$_.Length -eq $size)"
                            if("" -ne $version){ $searchCmd+=" -and (`$_.VersionInfo.FileVersionRaw -eq '$version')"}
                            $searchCmd+= "}}).FullName"
                            $result = Invoke-Expression ($searchCmd)
                            if($null -ne $hashRecord){
                                foreach ($r in $result){
                                    $hashRecord = $this.windowsInstaller.InvokeMethod('FileHash', $r ,0)
                                    $match = $true
                                    for($x=1; $x -le 4; $x++){
                                        if ($hashData[$x] -ne $hashRecord.GetParamProperty('IntegerData',$x)){
                                            $match = $false
                                        }
                                    }
                                    if($match){
                                        $sourcePath = $r
                                        break
                                    }
                                }
                                $result = $null
                            }
                            else{
                                if($result -is [array]){
                                    $sourcePath=$result[0]
                                }
                                else {
                                    $sourcePath=$result                                
                                }
                            }
                        }

                        if(!($previousSearchedFolders -contains (Split-Path $sourcePath -Parent))){
                            $previousSearchedFolders += (Split-Path $sourcePath -Parent)
                        }

                        if($null -eq $sourcePath){
                            "shit"
                        }
                    }
                }
            }
            $this.LastSequence++
            if(($this.LastSequence % 9000) -eq 0){
                ".new Cabinet" | Out-File -FilePath $this.DDFFILE -append -Encoding ascii

            }
            $SourcePathWithoutUnicode = ($sourcePath -replace "[^\u0000-\u007F]+")
            if($sourcePath -ne $SourcePathWithoutUnicode){
                Rename-Item $sourcePath -NewName $SourcePathWithoutUnicode
                $sourcePath = $SourcePathWithoutUnicode
            }
            """" + $sourcePath + """ """ + $fileKey +"""" | Out-File -FilePath $this.DDFFILE -append -Encoding ascii
        }while($null -ne $record) 
    }
    [void]FinaliseDDFFile(){
        [int]$diskID = 0
        [int]$MedialastSequence = 0
        Set-Location $PSScriptRoot
        makecab.exe  /f $this.DDFFILE

        Get-Content -Path ($PSScriptRoot + "\" + $this.cabName + ".cab.Log") | ForEach-Object {[regex]::Matches($_,'^\(([0-9]+)\) (?![0-9]+:)(.+\.(?i)cab)')} | ForEach-Object{
            [int]$diskID =  $_.groups[1].Value 
            $cabinet = $_.groups[2].Value
            [int]$MedialastSequence = ((Select-String ($PSScriptRoot + "\" +$this.cabName + ".cab.log") -pattern ('^\(' + $diskID + '\) ([0-9]+):') )[-1].Matches[0].Groups[1].Value)

            $destView = $this.destDatabase.InvokeMethod('OpenView', 
                "INSERT INTO `Media` (`DiskId`, `LastSequence`, `Cabinet`) VALUES (" + 
                $diskID + ", " + ($MedialastSequence) + ", '" + $cabinet + "')")
            $destView.InvokeMethod('Execute')
        }
 
    }
    [string]ProvideAssembly([System.__ComObject]$db,[string]$componentID,[int]$attribute){
        [string]$strongName = ""
        $AssemblyView = $db.InvokeMethod('Openview',"SELECT Name, Value FROM MsiAssemblyName WHERE Component_='"+ $componentID+"'")
        $AssemblyView.InvokeMethod('Execute') |Out-Null
        $AssemblyRecord = $AssemblyView.InvokeMethod('Fetch')
        while ($null -ne $AssemblyRecord){
            if ($AssemblyRecord.GetParamProperty('StringData',1) -eq 'Name'){
                $strongName = $AssemblyRecord.GetParamProperty('StringData',2) + "," + $strongName
            }
            else{
                $strongName = $strongName + $AssemblyRecord.GetParamProperty('StringData',1) + "=" + """""" + $AssemblyRecord.GetParamProperty('StringData',2) + """""" + ","
            }
            $AssemblyRecord = $AssemblyView.InvokeMethod('Fetch')
        }
        "wscript.echo CreateObject(""WindowsInstaller.Installer"").ProvideAssembly("""+ $strongName +""",vbnullstring,-2," + $attribute + ")"| Out-File -FilePath ($PSScriptRoot + "\" + $componentID + ".vbs") -Encoding ascii

        $pinfo = New-Object System.Diagnostics.ProcessStartInfo
        $pinfo.FileName = (Join-Path -path (get-item env:\windir).value -ChildPath "system32\cscript.exe")
        $pinfo.RedirectStandardError = $true
        $pinfo.RedirectStandardOutput = $true
        $pinfo.UseShellExecute = $false
        $pinfo.Arguments = "//Nologo " + $PSScriptRoot + "\" + $componentID + ".vbs"
        $pinfo.CreateNoWindow = $true
        $p = New-Object System.Diagnostics.Process
        $p.StartInfo = $pinfo
        $p.Start() | Out-Null
        $stdout = $p.StandardOutput.ReadLine()
        $p.WaitForExit()
        return $stdout
    }

    [void]SetFeatureStates([System.__ComObject]$session) {
        [string]$productCode = $session.GetParamProperty('ProductProperty','ProductCode') 
        #  grab a db object from the session
        $db = $session.GetProperty('Database')
        
        $view = $db.InvokeMethod('OpenView',"Select Feature FROM Feature")
        $view.InvokeMethod('Execute')
        $record = $view.InvokeMethod('Fetch')
        while ($null -ne $record) {
            [string]$featureName = $record.GetParamProperty('StringData',1)
            $featureCurrentState = $this.windowsInstaller.GetParamProperty('FeatureState',$productCode,$featureName)
            if ($featureCurrentState -ne 3){
                $destView = $this.destDatabase.InvokeMethod('OpenView',"UPDATE Feature SET Level=0 WHERE Feature='"+ $featureName +"'")
                $destView.InvokeMethod('Execute')
            }
            $record = $view.InvokeMethod('Fetch')
        }   
    }
}
function Set-SIS {
    param (
        [System.__ComObject]$destDatabase
    )
    $SIS = $destDatabase.GetParamProperty('SummaryInformation',4)
    $SIS.SetParamProperty('Property',7,'Intel')
    $SIS.SetParamProperty('Property',14,400)
    $SIS.SetParamProperty('Property',15,2)
    $SIS.SetParamProperty('Property',9,('{' + ([guid]::NewGuid()).toString() + '}'))
    $SIS.InvokeMethod('Persist')    
}
$wi = New-Object -ComObject WindowsInstaller.Installer

$destDatabase = $wi.InvokeMethod('opendatabase', ($PSScriptRoot +'\Merged.msi'),[openMode]::msiOpenDatabaseModeCreate)
Set-SIS $destDatabase
$myCabPacker = [CabPacker]::new($destDatabase, $wi)
$myCabPacker.InitDDFFile()

$view = $destDatabase.InvokeMethod('OpenView', "CREATE TABLE `Media` ( `DiskId` SHORT NOT NULL, `LastSequence` LONG NOT NULL, `DiskPrompt` CHAR(64) LOCALIZABLE, `Cabinet` CHAR(255), `VolumeLabel` CHAR(32), `Source` CHAR(72) PRIMARY KEY `DiskId`)")
$view.InvokeMethod('Execute')
$view = $destDatabase.InvokeMethod('OpenView', "CREATE TABLE `_ARPackageInfo` ( `Id` CHAR(72) NOT NULL, `ProductName` LONGCHAR NOT NULL, `ProductCode` LONGCHAR PRIMARY KEY `Id`)")
$view.InvokeMethod('Execute')

 Get-InstallProperties $wi| 
    Where-Object -Property InstallDate -gt (get-date).AddDays(-5).ToString("yyyyMMdd") |
    Select-Object -Property 'LocalPackage' |
    ForEach-Object{
    $basename = [io.path]::GetFileNameWithoutExtension($_.LocalPackage)
    $session = $wi.InvokeMethod('OpenPackage', $_.LocalPackage ,[openMode]::msiOpenDatabaseModeReadOnly)
    $sourceDatabase = $session.GetProperty('Database')

    #populate the _ARPackageInfo Table
    $view = $destDatabase.InvokeMethod('OpenView',"INSERT INTO _ARPackageInfo (Id, ProductName,ProductCode)  VALUES ('"+ $basename +
        "','"+ $session.GetParamProperty('Property','ProductName') +
        "','"+ $session.GetParamProperty('Property','ProductCode') +
        "' )")
    $view.InvokeMethod('Execute')

    #these tables get special treatment because of keypaths
    $fileRenameList=@{}
    $componentRenameList=@{}
    $regRenameList=@{}

#do something with the sequence tables?
#do something with the properties to pull existing values?
    MSICopyData $sourceDatabase $destDatabase 'File' $basename $fileRenameList $componentRenameList $regRenameList
    MSICopyData $sourceDatabase $destDatabase 'Registry' $basename $fileRenameList $componentRenameList $regRenameList

#copy all tables
    $view = $sourceDatabase.InvokeMethod('OpenView',	( "SELECT * FROM _Tables WHERE Name<>'File' AND Name<>'Registry' AND Name<>'_Property' AND Name<>'Media' AND Name<>'#_BaselineData' AND Name<>'#_FolderCache' AND Name<>'#_BaselineCost' AND Name<>'#_BaselineFile' AND Name <> '#_PatchCache' " ))
    $view.InvokeMethod('Execute')
    $RecordD = $view.InvokeMethod('Fetch')
    while  ($null -ne $RecordD) {
        MSICopyData $sourceDatabase $destDatabase ($RecordD.GetParamProperty('StringData',1)) $basename $fileRenameList $componentRenameList $regRenameList
        $RecordD = $view.InvokeMethod('Fetch')
    }
       
    $myCabPacker.AddFiles($session,$fileRenameList)
    $mycabPacker.SetFeatureStates($session)
    
    Release-Ref $session
    Release-Ref $sourceDatabase
}
$myCabPacker.FinaliseDDFFile()
#set the productname to the UPN if passed on the commandline
if($null -ne $args[0]){
    $view = $destDatabase.InvokeMethod('OpenView', "UPDATE `Property` SET `Value`='$($args[0])' WHERE `Property`='ProductName'")
    $view.InvokeMethod('Execute')
}

$destDatabase.InvokeMethod('commit')
Release-Ref $destDatabase
@('*.cab.log','*.rpt','*.ddf','*.vbs') | ForEach-Object {Remove-Item ($PSScriptRoot + "\" + $_) }
