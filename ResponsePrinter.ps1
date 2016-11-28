<#
V1.1.0
This script is used to parse the call log from Microsip and match the phone numbers with a list of fire fighters and their skillset.
This information is then displayed on screen so that other fire fighters may see who is on their way to the station. 

Windows Management Framework 5 is required to run this script. 
#>

###################################
##### Application Configuration####
###################################

$FireFightersRegister = 'C:\Repositories\ResponsePrinter\Sample Data\FireFighters.csv'
$IniFile = "C:\Repositories\ResponsePrinter\Sample Data\microsip.ini"

###################################
###End Application Configuration###
###################################

Function Get-IniContent {  
    <#  
    .Synopsis  
        Gets the content of an INI file  
          
    .Description  
        Gets the content of an INI file and returns it as a hashtable  
          
    .Notes  
        Author        : Oliver Lipkau <oliver@lipkau.net>  
        Blog        : http://oliver.lipkau.net/blog/  
        Source        : https://github.com/lipkau/PsIni 
                      http://gallery.technet.microsoft.com/scriptcenter/ea40c1ef-c856-434b-b8fb-ebd7a76e8d91 
        Version        : 1.0 - 2010/03/12 - Initial release  
                      1.1 - 2014/12/11 - Typo (Thx SLDR) 
                                         Typo (Thx Dave Stiff) 
          
        #Requires -Version 2.0  
          
    .Inputs  
        System.String  
          
    .Outputs  
        System.Collections.Hashtable  
          
    .Parameter FilePath  
        Specifies the path to the input file.  
          
    .Example  
        $FileContent = Get-IniContent "C:\myinifile.ini"  
        -----------  
        Description  
        Saves the content of the c:\myinifile.ini in a hashtable called $FileContent  
      
    .Example  
        $inifilepath | $FileContent = Get-IniContent  
        -----------  
        Description  
        Gets the content of the ini file passed through the pipe into a hashtable called $FileContent  
      
    .Example  
        C:\PS>$FileContent = Get-IniContent "c:\settings.ini"  
        C:\PS>$FileContent["Section"]["Key"]  
        -----------  
        Description  
        Returns the key "Key" of the section "Section" from the C:\settings.ini file  
          
    .Link  
        Out-IniFile  
    #>  
      
    [CmdletBinding()]  
    Param(  
        [ValidateNotNullOrEmpty()]  
        [ValidateScript({(Test-Path $_) -and ((Get-Item $_).Extension -eq ".ini")})]  
        [Parameter(ValueFromPipeline=$True,Mandatory=$True)]  
        [string]$FilePath  
    )  
      
    Begin  
        {Write-Verbose "$($MyInvocation.MyCommand.Name):: Function started"}  
          
    Process  
    {  
        Write-Verbose "$($MyInvocation.MyCommand.Name):: Processing file: $Filepath"  
              
        $ini = @{}  
        switch -regex -file $FilePath  
        {  
            "^\[(.+)\]$" # Section  
            {  
                $section = $matches[1]  
                $ini[$section] = @{}  
                $CommentCount = 0  
            }  
            "^(;.*)$" # Comment  
            {  
                if (!($section))  
                {  
                    $section = "No-Section"  
                    $ini[$section] = @{}  
                }  
                $value = $matches[1]  
                $CommentCount = $CommentCount + 1  
                $name = "Comment" + $CommentCount  
                $ini[$section][$name] = $value  
            }   
            "(.+?)\s*=\s*(.*)" # Key  
            {  
                if (!($section))  
                {  
                    $section = "No-Section"  
                    $ini[$section] = @{}  
                }  
                $name,$value = $matches[1..2]  
                $ini[$section][$name] = $value  
            }  
        }  
        Write-Verbose "$($MyInvocation.MyCommand.Name):: Finished Processing file: $FilePath"  
        Return $ini  
    }  
          
    End  
        {Write-Verbose "$($MyInvocation.MyCommand.Name):: Function ended"}  
} 

# Call Details Class
class CallDetails
{
   # Property
   hidden [int]$iniId
   [string]$FirstName
   [string]$LastName
   [string] $FullName
   [string[]]$Skills
   [string]$PhoneNo
   [DateTime]$DateTime
   [string]$TimeCalled
   [string]$Brigade
   # Constructor
   CallDetails ([int]$iniId,[string]$PhoneNo,[int]$UnixTime,[string]$FirstName,[string]$LastName, [string]$Skills, [string]$Brigade)
   {
        $startDate = [DateTime]::new(1970, 1, 1);
        $this.DateTime = $startDate.AddSeconds($UnixTime);
        $this.PhoneNo = $PhoneNo;
        $this.iniId = $iniId;
        $this.Skills = $Skills -split ";" | sort;
        $this.FirstName = $FirstName;
        $this.LastName = $LastName;
        $this.FullName = "$FirstName $LastName";
        $this.TimeCalled = $this.DateTime.ToLocalTime().ToString('HH:MM')
        $this.Brigade = $Brigade
        
   }
   
}
$FireFighters = import-csv $FireFightersRegister

do
{
    
    $iniContents = Get-IniContent -FilePath $IniFile
    $CallList = [System.Collections.ArrayList]::new()
    foreach($key in $iniContents.Calls.Keys)
        {
            $PhoneNo = ((($iniContents.Calls[$key]) -split "@")[0])
            if($PhoneNo -match ";")
            {
                $PhoneNo = $PhoneNo.Split(";")[0]
            }
            $ff = $FireFighters.Where({$_.PhoneNumber -eq $PhoneNo})
            $CallList.Add([CallDetails]::New("$key",$PhoneNo,$((($iniContents.Calls[$key]) -split ";")[3]),$ff.FirstName,$ff.LastName,$ff.Skills,$ff.Brigade)) | Out-Null
        }
    $FilteredCalls = $CallList | where DateTime -Ge $((get-date).ToUniversalTime().AddMinutes(-30))
    $FilteredCalls.ForEach({$_.DateTime = $_.DateTime.ToLocalTime()})
    cls
    $FilteredCalls | sort DateTime| FT FullName, Brigade, Skills, TimeCalled -AutoSize
    if(!$FilteredCalls)
    {
        Write-Host -ForegroundColor Yellow "******* Waiting For Calls *******"
    }
    Start-Sleep -Seconds 2
}
while($true)

#([datetime]::UtcNow.Subtract(([DateTime]::new(1970, 1, 1)).AddMinutes(24))).TotalSeconds #use to help with the INI File Testing.
