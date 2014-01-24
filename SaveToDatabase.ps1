Function Export-ToVCheckDB { 
    [CmdletBinding()] 
    param 
       ( 
       [Parameter(ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $InputObject, # The data that you would like written to your SQL table
       [Parameter(Mandatory=$true)]  [string] $ProviderType, # This is the type of provider e.g. VirtualCenter/vCheck
       [Parameter(Mandatory=$true)]  [string] $Description, # Description field to be populated from each plugin
       [Parameter(Mandatory=$true)]  [string] $InstanceName, # This the instance of $ProviderType (i.e. This would be the name of the VC against which the report was run)
       [Parameter(Mandatory=$false)] [string] $Version="1",  # This is provided to allow for creation of new tables in the SQL Database - most people would not use this
       [Parameter(Mandatory=$false)] [switch] $CreateTableIfDoesNotExist, # Tells the script to generate tables if they do not exist. If not specified, the script bombs out if the table does not exist
       [Parameter(Mandatory=$false)] [DateTime] $DateTime, # Allows for overwriting of current DateTime (You may want to date all entried for the log run at exactly the same time?)
       [Parameter(Mandatory=$true)] [string] $ConnectionString # Standard SQL connection string to be used for accessing / writing to the DB.
       )
       # Use current Datetime if this has not been provided
       if (!$DateTime) { $DateTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss" }     
       
       # Convert $inputObject to a Datatable
           If (($InputObject.GetType().Name) -notmatch "DataTable")
       {
			  $Inputobject = Out-DataTable -inputobject $inputobject
       }
        
       # Strip spaces from $description as this is used to name the SQL table
       $Description = Remove-SpecialCharacters ((Get-Culture).TextInfo.ToTitleCase($Description.ToLower()))
       
       # construct the SQL Table Name
       $TableName = $ProviderType+"_"+$Description+"_v"+$Version

       # Check if the Table already exists in the SQL DB
       # If table does not exist, check if we should create it
       If (!(Submit-SQLQuickQuery -connectionstring $connectionstring -query "Select * from sys.Tables" | Where-Object{$_.name.tolower() -eq  ($tablename.tolower())}))
           {
                If ($CreateTableIfDoesNotExist)
                {
                        Add-SQLTable -TableName $TableName -InputObject $InputObject -ConnectionString $ConnectionString
                }
                Else
                {      
                        return "Table not found - please create this manually or specify -CreateTableIfDoesNotExist to attempt automatic table generation"
                }
           }
       
       # Add $InputObjectDT to the SQL table
        Write-DataTable -TableName $tablename -InputObject $InputObject -InstanceName $InstanceName -datetime $datetime -ConnectionString $ConnectionString
 
} #end function
###############################################################################

Function Add-SqlTable { 
    [CmdletBinding()] 
    param( 
    [Parameter(Mandatory=$true)]  $InputObject,
    [Parameter(Mandatory=$true)] [String]$TableName, 
    [Parameter(Mandatory=$false)] [Int32]$MaxLength=256,
    [Parameter(Mandatory=$true)] [string]$connectionstring	
    )
    # Convert $inputObject to a Datatable
           If (($InputObject.GetType().Name) -notmatch "DataTable")
       {
			  $Inputobject = Out-DataTable -inputobject $inputobject
       }
      
#	$coldatetime = New-Object system.Data.DataColumn Datetime,([System.DateTime])
#	$inputobject.Columns.Add($coldatetime)
	
 try {
# Check if Table exists
    If ((Submit-SQLQuickQuery -connectionstring $connectionstring -query "Select * from sys.Tables" | Where-Object{$_.name.tolower() -eq  ($tablename.tolower())}))
	 {
	 return "Table $tablename already exists"
	 }
      
	# Create new table containing P_Id and a datetimeStamp
	Submit-SQLQuickQuery  -connectionstring $connectionstring -query "Create Table $TableName (P_Id int PRIMARY KEY IDENTITY, InstanceName nvarchar(50))" 
	
	# Now trawl through each Column in our DataTable and create an SQL Column for this.
	
   Foreach ($column in $inputobject.Columns) 
    { 
        $sqlDbType = "$(Get-SqlType $column.DataType.Name)" 
	
        if ($sqlDbType -eq 'VarBinary' -or $sqlDbType -eq 'VarChar') 
        { 
		
            if ($MaxLength -gt 0) 
            {$dataType = "$sqlDbType`($MaxLength`)"}
            else
            { $sqlDbType  = "$(Get-SqlType $column.DataType.Name)Max"
              $dataType =  $sqlDbType
		    }
        } 
        else 
        {
	$dataType =  $sqlDbType 
	} 
    Submit-SQLQuickQuery -query "ALTER TABLE $TableName ADD `"$column`" $dataType" -connectionstring $connectionstring
   }
    Submit-SQLQuickQuery -query "ALTER TABLE $TableName ADD `"SourceProxy`" varchar(50)" -connectionstring $connectionstring 
    Submit-SQLQuickQuery -query "ALTER TABLE $TableName ADD `"DateTime`" DateTime" -connectionstring $connectionstring 
 }
catch {
    $message = $_.Exception.GetBaseException().Message
    Write-Error $message
	}
  



} #end function
###############################################################################

Function Out-DataTable { 
  
  
    [CmdletBinding()] 
    param([Parameter(Position=0, Mandatory=$true, ValueFromPipeline = $true)] [PSObject[]]$InputObject)
begin {
 
        $dt = new-object Data.datatable   
        $First = $true  
    
}
process {
 
        foreach ($object in $InputObject) 
        { 
            $DR = $DT.NewRow()   
            foreach($property in $object.PsObject.get_properties()) 
            {   
                if ($first) 
                {   
                    $Col =  new-object Data.DataColumn   
                    $Col.ColumnName = $property.Name.ToString()   
                    if ($property.value) 
                    { 
                        If ($property.value -eq $null)
                            {$Col.DataType =  [System.DBNull]}
                       elseif ($property.value -isnot [System.DBNull]) { 
                            $Col.DataType = [System.Type]::GetType("$(Get-Type $property.TypeNameOfValue)") 
                         } 
                    } 
                    $DT.Columns.Add($Col) 
                }   
                if ($property.Gettype().IsArray) { 
                    $DR.Item($property.Name) =$property.value | ConvertTo-XML -AS String -NoTypeInformation -Depth 1 
                }   
               else { 
                    If (!($property.value))
                        {$DR.Item($property.Name) = ""}
                    Else
                        {$DR.Item($property.Name) = $property.value}
                } 
            }   
            $DT.Rows.Add($DR)   
            $First = $false 
        } 
    
}
end {
 
       return @(,($dt)) 
    
}


} #end function
###############################################################################

Function Write-DataTable { 
    [CmdletBinding()] 
    param( 
    [Parameter(Position=0, Mandatory=$false)] [string]$ServerInstance, 
    [Parameter(Position=1, Mandatory=$false)] [string]$Database, 
    [Parameter(Position=2, Mandatory=$false)] $InputObject,
    [Parameter(Position=3, Mandatory=$false)] [String]$TableName, 
    [Parameter(Position=4, Mandatory=$false)] [DateTime]$DateTime,
    [Parameter(Position=5, Mandatory=$false)] [string]$InstanceName,
	 [Parameter(Position=6, Mandatory=$false)] [string]$ConnectionString

   )
    $Properties = $InputObject | Get-Member | Where-Object{$_.Membertype -eq "Property"} | ForEach-Object{$_.Name}
    $SourceProxy = $ENV:Computername
    forEach ($row in $InputObject)
	{
	#Prepare SQL insert statement and execute it
	$StartInsertstatement =   "INSERT Into $TableName (" 
       $InsertStatement =  "VALUES("
       $col = 0
	   foreach($Property in $Properties)
       {
              $col++;
              if ($col -gt 1) { $InsertStatement += "," ; $StartInsertstatement += ","}
		     $StartInsertstatement += "[$Property]"
		
		#In the INSERT statement, do special tratment for Nulls, Dates and XML. Other special cases can be added as needed.
              if (-not $row.$Property)
              {
                     $InsertStatement += "null `n"
              }
              elseif ($row.$Property -is [datetime])
              {
                     $InsertStatement += "'" + $row.$Property.ToString("yyyy-MM-dd HH:mm:ss.fff") + "'`n"
              }
              elseif ($row.$Property -is [System.Xml.XmlNode] -or $row.$Property -is [System.Xml.XmlElement])
              {
                     $InsertStatement += "'" + ([string]$($row.$Property.Get_OuterXml())).Replace("'","''") + "'`n"
              }
              else
              {
                     $InsertStatement += "'" + ([string]$($row.$Property)).Replace("'","''") + "'`n"
              }
       }
       $InsertStatement +=",'$Datetime','$SourceProxy','$InstanceName')"
	   $StartInsertStatement +=",[Datetime],[SourceProxy],[InstanceName])"
	   $InsertStatement = $StartInsertstatement + $InsertStatement
	 Submit-SQLQuickQuery -ConnectionString $ConnectionString  -query "$InsertStatement" 

	}

} #end function
###############################################################################

Function Get-SqlType { 
  
  
    param([string]$TypeName)
    switch ($TypeName)  
        { 
            'Boolean' {[Data.SqlDbType]::Bit} 
            'Byte[]' {[Data.SqlDbType]::VarBinary} 
            'Byte'  {[Data.SQLDbType]::VarBinary} 
            'Datetime'  {[Data.SQLDbType]::DateTime} 
            'Decimal' {[Data.SqlDbType]::Decimal} 
            'Double' {[Data.SqlDbType]::Float} 
            'Guid' {[Data.SqlDbType]::UniqueIdentifier} 
            'Int16'  {[Data.SQLDbType]::SmallInt} 
            'Int32'  {[Data.SQLDbType]::Int} 
            'Int64' {[Data.SqlDbType]::BigInt} 
            'UInt16'  {[Data.SQLDbType]::SmallInt} 
            'UInt32'  {[Data.SQLDbType]::Int} 
            'UInt64' {[Data.SqlDbType]::BigInt} 
            'Single' {[Data.SqlDbType]::Decimal}
            default {[Data.SqlDbType]::VarChar} 
        } 
     


} #end function
###############################################################################

Function Submit-SQLQuickQuery { 
            Param (
	    [Parameter(Mandatory=$true)]$query,
        [Parameter(Mandatory=$true)] [string]$connectionstring	
			)
            $conn = new-object ('System.Data.SqlClient.SqlConnection')
		    $conn.ConnectionString = $ConnectionString
	        if (test-path variable:\conn) {
	            $conn.close()
	        } else {
	            $conn = new-object ('System.Data.SqlClient.SqlConnection')
	        }
	        $conn.Open()
	        $sqlCmd = New-Object System.Data.SqlClient.SqlCommand
	        $sqlCmd.CommandTimeout = $CommandTimeout
	        $sqlCmd.CommandText = $query
            $sqlCmd.Connection = $conn
	        $data = $sqlCmd.ExecuteReader()
	        while ($data.read() -eq $true) {
	            $max = $data.FieldCount -1
	            $obj = New-Object Object
	            For ($i = 0; $i -le $max; $i++) {
	                $name = $data.GetName($i)
	                if ($name.length -eq 0) {
	                    $name = "field$i"
	                }
	                $obj | Add-Member Noteproperty $name -value $data.GetValue($i) -Force
	            }
            $obj
	        }
	    $conn.close()
	    $conn = $null
} #end function
###############################################################################

Function Get-Type { 
  
    param($type) 
 
$types = @( 
'System.Boolean', 
'System.Byte[]', 
'System.Byte', 
'System.Char', 
'System.Datetime', 
'System.Decimal', 
'System.Double', 
'System.Guid', 
'System.Int16', 
'System.Int32', 
'System.Int64', 
'System.Single', 
'System.UInt16', 
'System.UInt32', 
'System.UInt64') 
 
    if ( $types -contains $type ) { 
        Write-Output "$type" 
    } 
    else { 
        Write-Output 'System.String' 
         
    } 

} #end function
###############################################################################

Function Object-ContainsProperty { 
 
       param 
       (
              [Parameter(Mandatory=$true)] $Object,
              [Parameter(Mandatory=$true)] [string] $PropertyName
       )
       foreach ($ObjectMember in ($Object | Get-Member)) 
       {
              if ($ObjectMember.Name -eq $PropertyName) { return $true }
       }
       return $false

} #end function
###############################################################################

function Encrypt-String($string, $path="$ENV:Temp\connstring.csv") {
  $secure = ConvertTo-SecureString $string -asPlainText -force
  $export = $secure | ConvertFrom-SecureString
  Set-Content $path $export
  "Script has been encrypted as '$path'"
}#end function
###############################################################################


function Get-EncryptedString($path="$ENV:Temp\connstring.csv") {
  trap { "Decryption failed"; break }
  $raw = Get-Content $path
  $secure = ConvertTo-SecureString $raw
  $helper = New-Object system.Management.Automation.PSCredential("test", $secure)
  $plain = $helper.GetNetworkCredential().Password
  return $plain
}#end function
###############################################################################


# Copied directly from : http://poshcode.org/2720
function New-SqlConnectionString {
#.Synopsis
#  Create a new SQL ConnectionString
#.Description
#  Builds a SQL ConnectionString using SQLConnectionStringBuilder with the supplied parameters
#.Example
#  New-SqlConnectionString -Server DBServer12 -Database NorthWind -IntegratedSecurity -MaxPoolSize 4 -Pooling
#.Example
#  New-SqlConnectionString -Server DBServer4 -Database NorthWind -Login SA -Password ""
[CmdletBinding(DefaultParameterSetName='Default')]
PARAM(
   # A full-blown connection string to start from
   [String]${ConnectionString},
   # The name of the application associated with the connection string.
   [String]${ApplicationName},
   # Whether asynchronous processing is allowed by the connection created by using this connection string.
   [Switch]${AsynchronousProcessing},
   # The name of the primary data file. This includes the full path name of an attachable database.
   [String]${AttachDBFilename},
   # The length of time (in seconds) to wait for a connection to the server before terminating the attempt and generating an error.
   [String]${ConnectTimeout},
   # Whether a client/server or in-process connection to SQL Server should be made.
   [Switch]${ContextConnection},
   # The SQL Server Language record name.
   [String]${CurrentLanguage},
   # The name and/or network address of the instance of SQL Server to connect to.
   [Parameter(Position=0)]
   [Alias("Server","Address")]
   [String]${DataSource},
   # Whether SQL Server uses SSL encryption for all data sent between the client and server if the server has a certificate installed.
   [Switch]${Encrypt},
   # Whether the SQL Server connection pooler automatically enlists the connection in the creation thread's current transaction context.
   [Switch]${Enlist},
   # The name or address of the partner server to connect to if the primary server is down.
   [String]${FailoverPartner},
   # The name of the database associated with the connection.
   [Parameter(Position=1)]
   [Alias("Database")]
   [String]${InitialCatalog},
   # Whether User ID and Password are specified in the connection (when false) or whether the current Windows account credentials are used for authentication (when true).
   [Switch]${IntegratedSecurity},
   # The minimum time, in seconds, for the connection to live in the connection pool before being destroyed.
   [Int]${LoadBalanceTimeout},
   # The maximum number of connections allowed in the connection pool for this specific connection string.
   [Int]${MaxPoolSize},
   # The minimum number of connections allowed in the connection pool for this specific connection string.
   [Int]${MinPoolSize},
   # Whether multiple active result sets can be associated with the associated connection.
   [Switch]${MultipleActiveResultSets},
   # The name of the network library used to establish a connection to the SQL Server.
   [String]${NetworkLibrary},
   # The size in bytes of the network packets used to communicate with an instance of SQL Server.
   [Int]${PacketSize},
   # The password for the SQL Server account.
   [AllowEmptyString()]
   [String]${Password},
   # Whether security-sensitive information, such as the password, is returned as part of the connection if the connection is open or has ever been in an open state.
   [Switch]${PersistSecurityInfo},
   # Whether the connection will be pooled or explicitly opened every time that the connection is requested.
   [Switch]${Pooling},
   # Whether replication is supported using the connection.
   [Switch]${Replication},
   # How the connection maintains its association with an enlisted System.Transactions transaction.
   [String]${TransactionBinding},
   # Whether the channel will be encrypted while bypassing walking the certificate chain to validate trust.
   [Switch]${TrustServerCertificate},
   # The type system the application expects.
   [String]${TypeSystemVersion},
   # The user ID to be used when connecting to SQL Server.
   [Alias("UserName","Login")]
   [String]${UserID},
   # Whether to redirect the connection from the default SQL Server Express instance to a runtime-initiated instance running under the account of the caller.
   [Switch]${UserInstance},
   # The name of the workstation connecting to SQL Server.
   [String]${WorkstationID},
   # Whether to return the SqlConnectionStringBuilder for further modification instead of just a connection string.
   [Switch]${AsBuilder}
)
BEGIN {
   if(!( 'System.Data.SqlClient.SqlConnectionStringBuilder' -as [Type] )) {
     $null = [Reflection.Assembly]::Load( 'System.Data, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089' ) 
   }
}
PROCESS {
   $Builder = New-Object System.Data.SqlClient.SqlConnectionStringBuilder -Property $PSBoundParameters
   if($AsBuilder) {
      Write-Output $Builder
   } else {
      Write-Output $Builder.ToString()
   }
}
}


# Provided by : http://tomasdeceuninck.wordpress.com/2013/03/11/powershell-replace-special-characters/
Function Remove-SpecialCharacters {
    <#
    .SYNOPSIS
        Removes special characters from a string.
 
    .DESCRIPTION
        Any character appart from alphanumerical characters and underscores will be replaced by an empty character.
     
    .EXAMPLE
        Remove-SpecialCharacters "Test-String's"
        This command will return "TestStrings"
 
    .OUTPUTS
        System.String
    #>
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $true,
            Position = 1,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias("Input")]
        [string] $InputString
    )
    Begin {}
    Process {
        $InputString = $InputString.Replace('á','a')
        $InputString = $InputString.Replace('Á','A')
        $InputString = $InputString.Replace('à','a')
        $InputString = $InputString.Replace('À','A')
        $InputString = $InputString.Replace('â','a')
        $InputString = $InputString.Replace('Â','A')
        $InputString = $InputString.Replace('ä','a')
        $InputString = $InputString.Replace('Ä','A')
        $InputString = $InputString.Replace('ã','a')
        $InputString = $InputString.Replace('Ã','A')
         
        $InputString = $InputString.Replace('ç','c')
         
        $InputString = $InputString.Replace('é','e')
        $InputString = $InputString.Replace('É','E')
        $InputString = $InputString.Replace('è','e')
        $InputString = $InputString.Replace('È','E')
        $InputString = $InputString.Replace('ê','e')
        $InputString = $InputString.Replace('Ê','E')
        $InputString = $InputString.Replace('ë','e')
        $InputString = $InputString.Replace('Ë','E')
         
        $InputString = $InputString.Replace('í','i')
        $InputString = $InputString.Replace('Í','I')
        $InputString = $InputString.Replace('ì','i')
        $InputString = $InputString.Replace('Ì','I')
        $InputString = $InputString.Replace('î','i')
        $InputString = $InputString.Replace('Î','I')
        $InputString = $InputString.Replace('ï','i')
        $InputString = $InputString.Replace('Ï','I')
         
        $InputString = $InputString.Replace('ñ','n')
        $InputString = $InputString.Replace('Ñ','N')
         
        $InputString = $InputString.Replace('ó','o')
        $InputString = $InputString.Replace('Ó','O')
        $InputString = $InputString.Replace('ò','o')
        $InputString = $InputString.Replace('Ò','O')
        $InputString = $InputString.Replace('ô','o')
        $InputString = $InputString.Replace('Ô','O')
        $InputString = $InputString.Replace('ö','o')
        $InputString = $InputString.Replace('Ö','O')
        $InputString = $InputString.Replace('õ','o')
        $InputString = $InputString.Replace('Õ','O')
         
        $InputString = $InputString.Replace('ú','u')
        $InputString = $InputString.Replace('Ú','U')
        $InputString = $InputString.Replace('ù','u')
        $InputString = $InputString.Replace('Ù','U')
        $InputString = $InputString.Replace('û','u')
        $InputString = $InputString.Replace('Û','U')
        $InputString = $InputString.Replace('ü','u')
        $InputString = $InputString.Replace('Ü','U')
         
        # Remove rest
        $strOutput = [System.Text.RegularExpressions.Regex]::Replace($InputString,"[^0-9a-zA-Z_]","")
        Write-Output $strOutput
    }
    End {}
}
