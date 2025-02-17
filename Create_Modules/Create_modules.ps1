<#
    $psISE.CurrentFile.Editor.ToggleOutliningExpansion()
#>
$ErrorActionPreference = "SilentlyContinue"
$ErrorActionPreference = "Stop"
$Global:Scriptpath = Split-Path $MyInvocation.MyCommand.Definition -Parent 
$scriptname = $([system.io.path]::GetFileNameWithoutExtension($(Split-path $MyInvocation.MyCommand.Definition -LEAF)))
$logdir = "$Scriptpath\Logs"
$Configdir = "$Scriptpath\Config"
$Inputdir = "$Scriptpath\Input"
$Outputdir = "$Scriptpath\Output"
$SharedPath = "E:\Data\Scripts\Shared\"
$FunctionsPath = join-path -Path $SharedPath -childpath 'New Functions'
$FunctionsFiles = Get-ChildItem -Path $FunctionsPath -Filter '*.ps1' -recurse
Foreach ($Functions in $FunctionsFiles){
    Write-Verbose "Including '$($Functionsze.fullname)'"
    . $Functions.fullname
}

cls
Try{
    #Import-Module "$env:USERPROFILE\documents\modules\Standaard\Standaard.psm1"
}catch{
    Write-Error "Import-module not found"
}
If (!(test-path $logdir)){
    New-Item $logdir -ItemType directory | out-null
}
If (!(test-path $Configdir)){
    New-Item $Configdir -ItemType directory | out-null
}
If (!(test-path $Inputdir)){
    New-Item $Inputdir -ItemType directory | out-null
}
If (!(test-path $Outputdir)){
    New-Item $Outputdir -ItemType directory | out-null
}
$logdir = "$Logdir\"
$Outputdir = "$Outputdir\"
$Inputdir = "$Inputdir\"
$GLOBAL:Configfile = "$($Configdir)\$($scriptname)_Config.xml"
$Global:ERRFILE = "$($logdir)$($scriptname)_Error.txt"
$Global:VerboseFile = "$($logdir)$($scriptname)_Verbose.txt"
$Global:LogFile = "$($logdir)$($scriptname)_Log.txt"
$Global:VoortgangFile = "$($logdir)$($scriptname)_Voortgang.txt"
#$Global:DomainController = (Get-ADDomainController).name 
#$Global:DomainName = (Get-ADDomain).dnsroot
#region Record-definitions
#endregion Record-definitions

#region Functies
Function Main{
    $Functionname = "Main"
    Try{
        Write-log -Category Verbose -message "[$(get-date -format "HH:mm")] [$Functionname] : Execute : "
        #$Global:ExecNew = $True#
        If ($Global:ExecProces1){            
            $StatusTekst  = "Aanmaken van modules"
            Write-log -Category Host -message "[$Functionname] : Execute : $StatusTekst" -color Magenta
            Write-log -Category Log -message "[$(Get-date -Format "dd-MM-yyyy HH:mm")] $StatusTekst" -logfile $Global:VoortgangFile
            Write-log -Category Verbose -message "[$(get-date -format "HH:mm")] [$Functionname] : Execute : Init_Module"
            $Result  = Init_Module            
            If ($Result){
                Write-log -Category Verbose -message "[$(get-date -format "HH:mm")] [$Functionname] : Execute : Create_Module"
                Create_module 
                Write-log -Category Verbose -message "[$(get-date -format "HH:mm")] [$Functionname] : Execute : Create_Module_file"
                Create_Module_file
                IF (Test-Path $Global:ModuleFullname){
                    notepad.exe $Global:ModuleFullname
                }
            }                        
            #Execute_New
            #Versturen_Backup_Mail
        }
    }catch{
        Write-log -Category Error -message "[$Functionname] : Unknown error. "   
        Write-log -Category Error -message "[$Functionname] : Targetname   : $($_.CategoryInfo.targetname)"
        Write-log -Category Error -message "[$Functionname] : Fullname     : $($_.exception.gettype().Fullname)"
        Write-log -Category Error -message "[$Functionname] : Type fout    : $($_.CategoryInfo.category)"
        Write-log -Category Error -message "[$Functionname] : Position     : $($_.invocationinfo.positionmessage)"
        Write-log -Category Error -message "[$Functionname] : Errormessage : $($_.Exception.message)"   
    }
}
#region Basisfuncties
Function Create_XML{
    $FunctionName = "Create_XML"
    Write-log -Category Debug -message "[$Functionname] : `$Configfile = `'$Configfile`'"
    Try{
	    $XmlWriter = New-Object System.XMl.XmlTextWriter($GLOBAL:Configfile,$Null)
	    $xmlWriter.Formatting = "Indented"
	    $xmlWriter.Indentation = "4"
	    $xmlWriter.WriteStartDocument()
	    $xmlWriter.WriteStartElement("Config")# Write Config Element
	    $xmlWriter.WriteStartElement("Instellingen")# Write Instellingen
        $xmlWriter.WriteElementString("DisplayLogfiles","1")
        $xmlWriter.WriteElementString("RemoveLogfiles","1")
	    $xmlWriter.WriteElementString("RollBack","0")
        $xmlWriter.WriteElementString("Testrun","1")
	    $xmlWriter.WriteEndElement() # <-- Closing Instellingen
        $xmlWriter.WriteStartElement("InputFiles")# Write Element
        $xmlWriter.WriteElementString("InputFile","")
	    $xmlWriter.WriteEndElement() # <-- Closing 
        $xmlWriter.WriteStartElement("OutputFiles")# Write Element
        $xmlWriter.WriteElementString("Outputfile","Output.json")
	    $xmlWriter.WriteEndElement() # <-- Closing 
        $xmlWriter.WriteStartElement("Execute")# Write Element
        $xmlWriter.WriteElementString("ExecProces1","True")
	    $xmlWriter.WriteEndElement()	    
        $xmlWriter.WriteStartElement("Modules")# Write Element
        $xmlWriter.WriteElementString("ModuleProjectFile","Create_Word.xml")
        $xmlWriter.WriteElementString("ModuleProjectdir","E:\Data\Scripts\Repository\Create_Modules\Config\Module_Projects")
        $xmlWriter.WriteElementString("Modulesdir","E:\Data\Scripts\Modules")
        $xmlWriter.WriteEndElement()	    
        $xmlWriter.WriteStartElement("ModuleInfo")# Write Element
        $xmlWriter.WriteElementString("Author","Marten Wiltens")
        $xmlWriter.WriteElementString("Companyname","Bloemendaal Consultancy")
        $xmlWriter.WriteEndElement()	    
        $xmlWriter.WriteStartElement("Email")# Write Element
        $xmlWriter.WriteElementString("From","info@rdw.nl")
        $xmlWriter.WriteElementString("TestTo","mwiltens@rdw.nl")
        $xmlWriter.WriteElementString("To","ddp@rdw.nl")
        $xmlWriter.WriteElementString("Priority","Normal")
        $xmlWriter.WriteElementString("SmtpServer","smtp.beheer.tld")
        $xmlWriter.WriteEndElement() # <-- Closing 
	    $xmlWriter.WriteEndElement() # <-- Closing Config
	    $xmlWriter.WriteEndDocument()
	    $xmlWriter.Flush() | out-null
	    $xmlWriter.Close()
    }Catch{
        Write-log -Category Error -message "[$Functionname] : Unknown error. "   
        Write-log -Category Error -message "[$Functionname] : Targetname   : $($_.CategoryInfo.targetname)"
        Write-log -Category Error -message "[$Functionname] : Fullname     : $($_.exception.gettype().Fullname)"
        Write-log -Category Error -message "[$Functionname] : Type fout    : $($_.CategoryInfo.category)"
        Write-log -Category Error -message "[$Functionname] : Position     : $($_.invocationinfo.positionmessage)"
        Write-log -Category Error -message "[$Functionname] : Errormessage : $($_.Exception.message)"   
    }
}
Function Create_ProjectXML{
    $FunctionName = "Create_ProjectXML"
    Write-log -Category Debug -message "[$Functionname] : `$Configfile = `'$Configfile`'"
    Try{
	    $XmlWriter = New-Object System.XMl.XmlTextWriter($GLOBAL:Configfile,$Null)
	    $xmlWriter.Formatting = "Indented"
	    $xmlWriter.Indentation = "4"
	    $xmlWriter.WriteStartDocument()
        $xmlWriter.WriteStartElement("Modules")# Write Element
        $xmlWriter.WriteElementString("ModuleName","Create_xxx")
        $xmlWriter.WriteElementString("Description","Omschrijving")
	    $xmlWriter.WriteElementString("Functionsdir","E:\Data\Scripts\FunctionsDir\Create xxx Functions")
	    $xmlWriter.WriteElementString("Moduleversion","0.1")
	    $xmlWriter.WriteElementString("FunctionsToExport","*")
	    $xmlWriter.WriteEndElement()	    
	    $xmlWriter.WriteEndDocument()
	    $xmlWriter.Flush() | out-null
	    $xmlWriter.Close()
    }Catch{
        Write-log -Category Error -message "[$Functionname] : Unknown error. "   
        Write-log -Category Error -message "[$Functionname] : Targetname   : $($_.CategoryInfo.targetname)"
        Write-log -Category Error -message "[$Functionname] : Fullname     : $($_.exception.gettype().Fullname)"
        Write-log -Category Error -message "[$Functionname] : Type fout    : $($_.CategoryInfo.category)"
        Write-log -Category Error -message "[$Functionname] : Position     : $($_.invocationinfo.positionmessage)"
        Write-log -Category Error -message "[$Functionname] : Errormessage : $($_.Exception.message)"   
    }
}
Function Read_XML{
	Param ($Default = $True)
    $FunctionName = "Read_XML"
    Write-log -Category Debug -message "[$Functionname] : `$Configfile = `'$Configfile`'"
    Try{
	#return
        [xml]$XML= get-content $Configfile
    	#$InputFile = "$($XML.Config.InputFiles.InputFile)"
        #$OutputFile = "$($XML.Config.OutputFiles.OutputFile)"
        #$DisplayLogfiles = $XML.Config.Instellingen.DisplayLogfiles
	    #$RollBack = $XML.Config.Instellingen.RollBack
        #$RemoveLogfiles = $XML.Config.Instellingen.RemoveLogfiles
	    #$TestRun = $XML.Config.Instellingen.TestRun
        #region Uitlezen-Instellingen

    	$InputFile = "$($XML.Config.InputFiles.InputFile)"
        $OutputFile = "$($XML.Config.OutputFiles.OutputFile)"
        $DisplayLogfiles = $XML.Config.Instellingen.DisplayLogfiles
	    $RollBack = $XML.Config.Instellingen.RollBack
        $RemoveLogfiles = $XML.Config.Instellingen.RemoveLogfiles
	    $TestRun = $XML.Config.Instellingen.TestRun
        #endregion Uitlezen-Instellingen
        #region Uitlezen-Execute
        Try{            
            $Global:ExecProces1 = $XML.Config.Execute.ExecProces1
        }Catch{
            $Global:ExecProces1 = $False
        }#ExecProces1
        #endregion Uitlezen-Execute
        #region Uitlezen-Module
        Try{            
            $Global:ModuleProjectFile = $XML.Config.Modules.ModuleProjectFile
        }Catch{
            $Global:ModuleProjectFile = $False
        }#ModuleProjectFile
        Try{            
            $Global:ModuleProjectdir = $XML.Config.Modules.ModuleProjectdir
        }Catch{
            $Global:ModuleProjectdir = $False
        }#ModuleProjectdir
        Try{            
            $Global:Modulesdir = $XML.Config.Modules.Modulesdir
        }Catch{
            $Global:Modulesdir = $False
        }#Functionsdir
        #endregion Uitlezen-Module
        #region Uitlezen-ModuleProject        
        $ProjectFile = "$($Global:ModuleProjectdir)\$($Global:ModuleProjectFile)"
        [xml]$ProjectXML= get-content $ProjectFile
        Try{            
            $Global:ModuleName = $ProjectXML.Modules.ModuleName
        }Catch{
            $Global:ModuleName = $False
        }#ModuleName
        Try{            
            $Global:Description = $ProjectXML.Modules.Description
        }Catch{
            $Global:Description = $False
        }#Description
        Try{            
            $Global:Functionsdir = $ProjectXML.Modules.Functionsdir
        }Catch{
            $Global:Functionsdir = $False
        }#Functionsdir
        Try{            
            $Global:Moduleversion = $ProjectXML.Modules.Moduleversion
        }Catch{
            $Global:Moduleversion = $False
        }#Moduleversion
        Try{            
            $Global:FunctionsToExport = $ProjectXML.Modules.FunctionsToExport
        }Catch{
            $Global:FunctionsToExport = $False
        }#FunctionsToExport
        #endregion Uitlezen-ModuleProject
        #region Uitlezen-ModuleInfo
        Try{            
            $Global:Author = $XML.Config.ModuleInfo.Author
        }Catch{
            $Global:Author = $False
        }#Author
        Try{            
            $Global:Companyname = $XML.Config.ModuleInfo.Companyname
        }Catch{
            $Global:Companyname = $False
        }#Companyname
        #endregion Uitlezen-ModuleInfo
        #region Uitlezen-Email
        $Global:From = $XML.Config.Email.From
        $Global:TestTo = $XML.Config.Email.TestTo
        $Global:To = $XML.Config.Email.To
        $Global:Priority = $XML.Config.Email.Priority
        $Global:SmtpServer = $XML.Config.Email.SmtpServer
        #endregion Uitlezen-Email


<#

	    <NewXML></NewXML>
	    $xmlWriter.WriteElementString("NewXML","")
	    $NewXML = $XML.Config.Instellingen.NewXML
        Display_info -Omschrijving "NewXML" -Waarde $Global:NewXML -Color Yellow
	        If ($NewXML -eq 1){
	            $Global:NewXML = $True
	        }Else{
	            $Global:NewXML = $False
	        }
#>
    #region global Instellingen
	$Global:InputFile = "$($Inputdir)$InputFile"
	$Global:OutputFile = "$($Outputdir)$OutputFile"
    $DisplayLogfiles = $XML.Config.Instellingen.DisplayLogfiles
	$RemoveLogfiles = $XML.Config.Instellingen.RemoveLogfiles
	If ($DisplayLogfiles -eq 1){
	    $Global:DisplayLogfiles = $True
	}Else{
	    $Global:DisplayLogfiles = $False
	}
	If ($RemoveLogfiles -eq 1){
	    $Global:RemoveLogfiles = $True
	}Else{
	    $Global:RemoveLogfiles = $False
	}
	If ($RollBack -eq 1){
	    $Global:RollBack = $True
	}Else{
	    $Global:RollBack = $False
	}
	If ($TestRun -eq 1){
	    $Global:TestRun = $True
	}Else{
	    $Global:TestRun = $False
	}
    #endregion global Instellingen
    #region global Execute
    <#
    Try{
	    If ($ExecProces1 -eq 1){
	        $Global:ExecProces1 = $True
	    }Else{
	        $Global:ExecProces1 = $False
	    }
    }Catch{
        $Global:NewDCs = $False
    }#ExecProces1
    #>
    #endregion global Execute

    }Catch{
        Write-log -Category Error -message "[$Functionname] : Unknown error. "   
        Write-log -Category Error -message "[$Functionname] : Targetname   : $($_.CategoryInfo.targetname)"
        Write-log -Category Error -message "[$Functionname] : Fullname     : $($_.exception.gettype().Fullname)"
        Write-log -Category Error -message "[$Functionname] : Type fout    : $($_.CategoryInfo.category)"
        Write-log -Category Error -message "[$Functionname] : Position     : $($_.invocationinfo.positionmessage)"
        Write-log -Category Error -message "[$Functionname] : Errormessage : $($_.Exception.message)"   
    }
}
#endregion Basisfuncties
#region Template functies
Function New{
    $Functionname = "New"
    #Write-log -Category Debug -message "[$Functionname] : `$Param1 = `'$Param1`' "
    Try{
        $KopVoettekst = "Nieuwe functie"
        Write-log -Category Host -message "Execute : $KopVoettekst" -color Magenta
        Write-log -Category Log -message "[$(Get-date -Format "dd-MM-yyyy HH:mm")] $KopVoettekst" -logfile $Global:VoortgangFile
       <#
        $KopVoettekst = "Nieuwe functie"
        Write-log -Category Verbose -message "[$Functionname] : Execute :  "
        KopVoettekst -Tekst $Koptekst -Functie Kop
        Write-log -Category Verbose -message "[$Functionname] : Execute :  Write-log -Category Log -message `"[$(Get-date -Format "dd-MM-yyyy HH:mm")] $Koptekst is uitgevoerd`""
        Write-log -Category Log -message "[$(Get-date -Format "dd-MM-yyyy HH:mm")] $Koptekst is uitgevoerd" 
        KopVoettekst -Tekst $Koptekst -Functie Voet
        $Aantalrecords  = ($Variabelenaam | Measure-object).count
        $i = 0
        Foreach ($loop in $Loops){
            $Pct = [math]::Ceiling((($I/$AantalRecords)) * 100)
            Write-progress -Activity $ActivityOmschrijving -Status "Complete: $Pct" -PercentComplete $Pct
            $I++
        }

        #>

    }catch{
        Write-log -Category Error -message "[$Functionname] : Unknown error. "   
        Write-log -Category Error -message "[$Functionname] : Targetname   : $($_.CategoryInfo.targetname)"
        Write-log -Category Error -message "[$Functionname] : Fullname     : $($_.exception.gettype().Fullname)"
        Write-log -Category Error -message "[$Functionname] : Type fout    : $($_.CategoryInfo.category)"
        Write-log -Category Error -message "[$Functionname] : Position     : $($_.invocationinfo.positionmessage)"
        Write-log -Category Error -message "[$Functionname] : Errormessage : $($_.Exception.message)"   
    }
}
Function New_Invoke{
    $Functionname = "New"
    #Write-log -Category Debug -message "[$Functionname] : `$Param1 = `'$Param1`' "
    Try{
        $KopVoettekst = "Nieuwe functie"
        Write-log -Category Host -message "Execute : $KopVoettekst" -color Magenta
        Write-log -Category Log -message "[$(Get-date -Format "dd-MM-yyyy HH:mm")] $KopVoettekst" -logfile $Global:VoortgangFile
        Write-log -Category Verbose -message "[$Functionname] : Execute :  "
        $Foodef = "Function Write-log { ${Function:Write-log} }"
        #$Foodef = "Function Write-log { ${Function:Write-log} }; Function Test-NLASetting { ${Function:Test-NLASetting} }"
        $RemoteVerboseFile  = "D:\$($scriptname)_Verbose.txt"
        $RemoteLogFile = "D:\$($scriptname)_Log.txt"
        $RemoteErrFile = "D:\$($scriptname)_ERR.txt"
         foreach ($Server in $Servers){
            Write-log -Category Host -message "[$Functionname] : Process $(($Server).servername)" -color Gray
            Invoke-Command -ComputerName $Server -ScriptBlock {
                Param ($Foodef, $VerboseFile, $LogFile, $ErrFile, $Functionname,$TestRun)
                . ([scriptblock]::Create($Foodef))
                $Functionname = "$($Functionname)_Remote"
                If (test-path $Verbosefile){Remove-item $Verbosefile -Confirm:$False}
                If (test-path $LogFile){Remove-item $LogFile -Confirm:$False}
                If (test-path $ErrFile){Remove-item $ErrFile -Confirm:$False}
            }-ArgumentList $Foodef, $RemoteVerboseFile, $RemoteLogFile, $RemoteErrFile, $Functionname,$Global:TestRun
            Read_RemoteLogfile -RemoteLogfile "\\$($Server)\D$\$($scriptname)_Verbose.txt" -Logfile $Global:VerboseFile
            Read_RemoteLogfile -RemoteLogfile "\\$($Server)\D$\$($scriptname)_Log.txt" -Logfile $Global:LogFile
        }
   }catch{
        Write-log -Category Error -message "[$Functionname] : Unknown error. "   
        Write-log -Category Error -message "[$Functionname] : Targetname   : $($_.CategoryInfo.targetname)"
        Write-log -Category Error -message "[$Functionname] : Fullname     : $($_.exception.gettype().Fullname)"
        Write-log -Category Error -message "[$Functionname] : Type fout    : $($_.CategoryInfo.category)"
        Write-log -Category Error -message "[$Functionname] : Position     : $($_.invocationinfo.positionmessage)"
        Write-log -Category Error -message "[$Functionname] : Errormessage : $($_.Exception.message)"   
    }
}
#endregion Template functies
#region Script functies
Function Versturen_Mail{
    $Functionname = "Versturen_Mail"
    Try{
        $MailSubject = ""
        Write-log -Category Verbose -message "[$Functionname] : Execute :  Execute_Sendmail -Subject `"$MailSubject`" -Body `$HTMLmeldingen -Priority High -bodyashtml True -From $Global:From"
        Execute_Sendmail -Subject $MailSubject -Body $HTMLmeldingen -bodyashtml True -From $Global:From -Attachment $VoortgangFile

    }catch{
        Write-log -Category Error -message "[$Functionname] : Unknown error. "   
        Write-log -Category Error -message "[$Functionname] : Targetname   : $($_.CategoryInfo.targetname)"
        Write-log -Category Error -message "[$Functionname] : Fullname     : $($_.exception.gettype().Fullname)"
        Write-log -Category Error -message "[$Functionname] : Type fout    : $($_.CategoryInfo.category)"
        Write-log -Category Error -message "[$Functionname] : Position     : $($_.invocationinfo.positionmessage)"
        Write-log -Category Error -message "[$Functionname] : Errormessage : $($_.Exception.message)"   
    }

}
Function Init_Module{
    $Functionname = "Init_Module"
    #Write-log -Category Debug -message "[$Functionname] : `$Param1 = `'$Param1`' "
    Try{
        $InitTab = 40
        
        $Global:ModuleDir = "$($Global:Modulesdir)\$Global:Modulename\$(Get-Childitem "$($Global:Modulesdir)\$Global:Modulename" -Recurse  -Directory | sort name | Select -Last 1)"
        $Global:Manifestname = "$Global:ModuleDir\$Global:Modulename.psd1"
        $NewVersion = Read_Manifest
        $Global:ModuleDir = "$($Global:Modulesdir)\$Global:Modulename\$NewVersion"
        $Global:ModuleFullname = "$($Global:Modulesdir)\$Global:Modulename\$Global:Modulename.psm1"
        $Global:Manifestname = "$($Global:Modulesdir)\$Global:Modulename\$Global:Modulename.psd1"
        $Global:ModuleFullname = "$Global:ModuleDir\$Global:Modulename.psm1"
        $Global:Manifestname = "$Global:ModuleDir\$Global:Modulename.psd1"

        Write-log -Category Host -message "$($("Functionsdir").Padright($InitTab)) : $Global:FunctionsDir" -color Magenta
        Write-log -Category Host -message "$($("Modulename").Padright($InitTab)) : $Global:Modulename" -color Magenta
        Write-log -Category Host -message "$($("ModulesDir").Padright($InitTab)) : $Global:Modulesdir" -color Magenta
        Write-log -Category Host -message "$($("ModuleDir").Padright($InitTab)) : $ModuleDir" -color Magenta
        Write-log -Category Host -message "$($("ModuleFullname").Padright($InitTab)) : $Global:ModuleFullname" -color Magenta
        Write-log -Category Host -message "$($("Manifestname").Padright($InitTab)) : $Global:Manifestname" -color Magenta
        IF (!(test-path $Global:Functionsdir)){
            Write-log -Category Host -message "[$Functionname] : $Global:Functionsdir is niet gevonden" -color Red
            return $False
        }else{
            IF (!([bool](Get-ChildItem -Path $Global:FunctionsDir))){
                Write-log -Category Host -message "[$Functionname] : Er zijn geen bestanden gevonden in $Global:Functionsdir " -color Red
                return $False
            }
        }
        return $True
    }catch{
        Write-log -Category Error -message "[$Functionname] : Unknown error. "   
        Write-log -Category Error -message "[$Functionname] : Targetname   : $($_.CategoryInfo.targetname)"
        Write-log -Category Error -message "[$Functionname] : Fullname     : $($_.exception.gettype().Fullname)"
        Write-log -Category Error -message "[$Functionname] : Type fout    : $($_.CategoryInfo.category)"
        Write-log -Category Error -message "[$Functionname] : Position     : $($_.invocationinfo.positionmessage)"
        Write-log -Category Error -message "[$Functionname] : Errormessage : $($_.Exception.message)"   
    }
}
Function Create_Module{
    $Functionname = "Create_Module"
    #Write-log -Category Debug -message "[$Functionname] : `$Param1 = `'$Param1`' "
    Try{
        $KopVoettekst = "Aanmaken Module $Global:Modulename"
        Write-log -Category Host -message "Execute : $KopVoettekst" -color Magenta
        Write-log -Category Log -message "[$(Get-date -Format "dd-MM-yyyy HH:mm")] $KopVoettekst" -logfile $Global:VoortgangFile
        IF (!(test-path $Global:Moduledir)){
            If ($Global:TestRun){
                Write-log -Category Verbose -message "[$Functionname] : Execute (Testrun) : New-Item -Path $Global:ModuleDir -ItemType Directory -WhatIf | Out-Null" 
                New-Item -Path $Global:ModuleDir -ItemType Directory -WhatIf | Out-Null
            }else{
                Write-log -Category Verbose -message "[$Functionname] : Execute : New-Item -Path $Global:ModuleDir -ItemType Directory | Out-Null" 
                New-Item -Path $Global:ModuleDir -ItemType Directory | Out-Null
            }
        }`
        IF (!(test-path $Global:Functionsdir)){
            If ($Global:TestRun){
                Write-log -Category Verbose -message "[$Functionname] : Execute (Testrun) : New-Item -Path $Global:Functionsdir -ItemType Directory -WhatIf | Out-Null" 
                New-Item -Path $Global:Functionsdir -ItemType Directory -WhatIf | Out-Null
                $Global:FunctionNames = ""
            }else{
                Write-log -Category Verbose -message "[$Functionname] : Execute :  New-Item -Path $Global:Functionsdir -ItemType Directory | Out-Null" 
                New-Item -Path $Global:Functionsdir -ItemType Directory | Out-Null
            }
        }Else{
            Write-log -Category Verbose -message "[$Functionname] : Execute: `$Global:FunctionNames = Get-ChildItem -Path $Global:FunctionsDir"
            $Global:FunctionNames = Get-ChildItem -Path $Global:FunctionsDir
        }
        IF (!(test-path $Global:Manifestname)){
            If ($Global:TestRun){
                Write-log -Category Verbose -message "[$Functionname] : Execute (Testrun) : Create_Manifest -Version $Global:Moduleversion -Desciption $Global:Description"
                Create_Manifest -Version $Global:Moduleversion -Desciption $Global:Description
            }else{
                Write-log -Category Verbose -message "[$Functionname] : Execute : Create_Manifest -Version $Global:Moduleversion -Desciption $Global:Description"
                Create_Manifest -Version $Global:Moduleversion -Desciption $Global:Description
            }
        }Else{
            If ($Global:TestRun){
                Write-log -Category Verbose -message "[$Functionname] : Execute (Testrun) : $NewVersion = Read_Manifest"
                Write-log -Category Verbose -message "[$Functionname] : Execute (Testrun) : Create_Manifest -Version $Global:Moduleversion -Desciption $Global:Description"
            }else{
                Write-log -Category Verbose -message "[$Functionname] : Execute (Testrun) : $NewVersion = Read_Manifest"
                $NewVersion = Read_Manifest
                Write-log -Category Verbose -message "[$Functionname] : Execute (Testrun) : Create_Manifest -Version $NewVersion -Desciption $Global:Description"
                Create_Manifest -Version $NewVersion -Desciption $Global:Description

            }
        }
    }catch{
        Write-log -Category Error -message "[$Functionname] : Unknown error. "   
        Write-log -Category Error -message "[$Functionname] : Targetname   : $($_.CategoryInfo.targetname)"
        Write-log -Category Error -message "[$Functionname] : Fullname     : $($_.exception.gettype().Fullname)"
        Write-log -Category Error -message "[$Functionname] : Type fout    : $($_.CategoryInfo.category)"
        Write-log -Category Error -message "[$Functionname] : Position     : $($_.invocationinfo.positionmessage)"
        Write-log -Category Error -message "[$Functionname] : Errormessage : $($_.Exception.message)"   
    }
}
Function New{
    $Functionname = "New"
    #Write-log -Category Debug -message "[$Functionname] : `$Param1 = `'$Param1`' "
    Try{
        $KopVoettekst = "Nieuwe functie"
        Write-log -Category Host -message "Execute : $KopVoettekst" -color Magenta
        Write-log -Category Log -message "[$(Get-date -Format "dd-MM-yyyy HH:mm")] $KopVoettekst" -logfile $Global:VoortgangFile
       <#
        $KopVoettekst = "Nieuwe functie"
        Write-log -Category Verbose -message "[$Functionname] : Execute :  "
        KopVoettekst -Tekst $Koptekst -Functie Kop
        Write-log -Category Verbose -message "[$Functionname] : Execute :  Write-log -Category Log -message `"[$(Get-date -Format "dd-MM-yyyy HH:mm")] $Koptekst is uitgevoerd`""
        Write-log -Category Log -message "[$(Get-date -Format "dd-MM-yyyy HH:mm")] $Koptekst is uitgevoerd" 
        KopVoettekst -Tekst $Koptekst -Functie Voet
        #>

    }catch{
        Write-log -Category Error -message "[$Functionname] : Unknown error. "   
        Write-log -Category Error -message "[$Functionname] : Targetname   : $($_.CategoryInfo.targetname)"
        Write-log -Category Error -message "[$Functionname] : Fullname     : $($_.exception.gettype().Fullname)"
        Write-log -Category Error -message "[$Functionname] : Type fout    : $($_.CategoryInfo.category)"
        Write-log -Category Error -message "[$Functionname] : Position     : $($_.invocationinfo.positionmessage)"
        Write-log -Category Error -message "[$Functionname] : Errormessage : $($_.Exception.message)"   
    }
}
Function Create_Manifest{
    Param ($Version)
    $Functionname = "Create_Manifest"
    Write-log -Category Debug -message "[$Functionname] : `$Version = `'$Version`' "
    Try{
        $KopVoettekst = "Create new Manifest : $Global:Manifestname"
        Write-log -Category Host -message "Execute : $KopVoettekst" -color Magenta
        Write-log -Category Log -message "[$(Get-date -Format "dd-MM-yyyy HH:mm")] $KopVoettekst" -logfile $Global:VoortgangFile
        If ($Global:TestRun){
            Write-log -Category Verbose -message "[$Functionname] : Execute (Testrun) : New-ModuleManifest -Path $Global:Manifestname -Author $Global:Author -CompanyName $Global:Companyname -ModuleVersion $Version -Description $Global:Description -FunctionsToExport $Global:FunctionsToExport -WhatIf"
            New-ModuleManifest -Path $Global:Manifestname -Author $Global:Author -CompanyName $Global:Companyname -ModuleVersion $Version -Description $Global:Description -FunctionsToExport $Global:FunctionsToExport -WhatIf
        }else{
            Write-log -Category Verbose -message "[$Functionname] : Execute : New-ModuleManifest -Path $Global:Manifestname -Author $Global:Author -CompanyName $Global:Companyname -ModuleVersion $Version -Description $Global:Description -FunctionsToExport $Global:FunctionsToExport"
            New-ModuleManifest -Path $Global:Manifestname -Author $Global:Author -CompanyName $Global:Companyname -ModuleVersion $Version -Description $Global:Description -FunctionsToExport $Global:FunctionsToExport
        }
    }catch{
        Write-log -Category Error -message "[$Functionname] : Unknown error. "   
        Write-log -Category Error -message "[$Functionname] : Targetname   : $($_.CategoryInfo.targetname)"
        Write-log -Category Error -message "[$Functionname] : Fullname     : $($_.exception.gettype().Fullname)"
        Write-log -Category Error -message "[$Functionname] : Type fout    : $($_.CategoryInfo.category)"
        Write-log -Category Error -message "[$Functionname] : Position     : $($_.invocationinfo.positionmessage)"
        Write-log -Category Error -message "[$Functionname] : Errormessage : $($_.Exception.message)"   
    }
}
Function Read_Manifest{
    $Functionname = "Read_Manifest"
    #Write-log -Category Debug -message "[$Functionname] : `$Param1 = `'$Param1`' "
    Try{
        $KopVoettekst = "Read Manifest-file $Global:Manifestname"
        Write-log -Category Host -message "Execute : $KopVoettekst" -color Magenta
        Write-log -Category Log -message "[$(Get-date -Format "dd-MM-yyyy HH:mm")] $KopVoettekst" -logfile $Global:VoortgangFile
        If (!(Test-path $Global:Manifestname)){
            return "0.1"
        }else{
            $SearchText = "ModuleVersion = "
            $Line = Get-Content $Global:Manifestname | Where {$_ -like $("$SearchText*")}
            $Line = $Line.Substring($Line.IndexOf("=") + 3)
            $Version = $Line.Substring(0,$Line.Length-1)
            $Major = [int]$Version.Substring(0,$Line.IndexOf("."))
            $Minor = [int]$Version.Substring($Line.IndexOf(".")+1)
            Switch ($Minor){
                0 {$Minor += 1}
                1 {$Minor += 1}
                2 {$Minor += 1}
                3 {$Minor += 1}
                4 {$Minor += 1}
                5 {$Minor += 1}
                6 {$Minor += 1}
                7 {$Minor += 1}
                8 {$Minor += 1}
                9 {$Major += 1;$Minor = 0}
            }
            return "$Major.$Minor"
        }
    }catch{
        Write-log -Category Error -message "[$Functionname] : Unknown error. "   
        Write-log -Category Error -message "[$Functionname] : Targetname   : $($_.CategoryInfo.targetname)"
        Write-log -Category Error -message "[$Functionname] : Fullname     : $($_.exception.gettype().Fullname)"
        Write-log -Category Error -message "[$Functionname] : Type fout    : $($_.CategoryInfo.category)"
        Write-log -Category Error -message "[$Functionname] : Position     : $($_.invocationinfo.positionmessage)"
        Write-log -Category Error -message "[$Functionname] : Errormessage : $($_.Exception.message)"   
    }
}
Function Create_Module_File{
    $Functionname = "Create_Module_File"
    #Write-log -Category Debug -message "[$Functionname] : `$Param1 = `'$Param1`' "
    Try{
        $KopVoettekst = "Create Module File : $Global:ModuleFullname"
        Write-log -Category Host -message "Execute : $KopVoettekst" -color Magenta
        Write-log -Category Log -message "[$(Get-date -Format "dd-MM-yyyy HH:mm")] $KopVoettekst" -logfile $Global:VoortgangFile
        If ([string]::IsNullOrEmpty($Global:Functionnames)){
            Write-log -Category Host -message "Er zijn geen functies gevonden om toe te voegen aan de module $Global:ModuleFullname " -color Red
        }else{
            $NewModule = @()
            foreach ($NewFunction in $Global:Functionnames){
                If ($TestRun){
                        Write-log -Category Host -message "Functionname $($NewFunction.fullname) is toegevoegd" -color Yellow
                        #$NewModule += Get-Content $NewFunction.fullname
                        #$NewModule += ""
                    }else{
                        Write-log -Category Host -message "Functionname $($NewFunction.fullname) is toegevoegd" -color Green
                        $NewModule += Get-Content $NewFunction.fullname
                        $NewModule += ""
    
                }
                

            }
            $NewModule | Out-File $Global:ModuleFullname
        }
    }catch{
        Write-log -Category Error -message "[$Functionname] : Unknown error. "   
        Write-log -Category Error -message "[$Functionname] : Targetname   : $($_.CategoryInfo.targetname)"
        Write-log -Category Error -message "[$Functionname] : Fullname     : $($_.exception.gettype().Fullname)"
        Write-log -Category Error -message "[$Functionname] : Type fout    : $($_.CategoryInfo.category)"
        Write-log -Category Error -message "[$Functionname] : Position     : $($_.invocationinfo.positionmessage)"
        Write-log -Category Error -message "[$Functionname] : Errormessage : $($_.Exception.message)"   
    }
}


#endregion Script functies
#endregion Functies
$StopWatch = [Diagnostics.Stopwatch]::StartNew()
$tab = 40
<#
    If (test-path $Configfile){Remove-Item $Global:Configfile -Confirm:$false}
    If (test-path $Configfile){&Notepad $Configfile}
#>
If (!(test-path $Configfile)){
    Create_XML
}

read_xml
$time = get-date -Format "HH:mm:ss dd-MM-yyyy"
If ($Global:RemoveLogfiles){
    If (test-path $VerboseFile){Remove-Item $Global:VerboseFile -Confirm:$false}
    If (test-path $ErrFile){Remove-Item $Global:ERRFILE -Confirm:$false}
}
Clear-Host
Display_info -Omschrijving "Time start" -Waarde $time -Color Green
Display_info -Omschrijving "Config file" -Waarde $Configfile -Color Cyan
Display_info -Omschrijving "Error-file" -Waarde $ERRFILE -Color Cyan
Display_info -Omschrijving "Verbose-File" -Waarde $VerboseFile -Color Cyan
Display_info -Omschrijving "Voortgang-file" -Waarde $Global:VoortgangFile -Color Cyan
Display_info -Omschrijving "Log-file" -Waarde $LogFile -Color Cyan
Display_info -Omschrijving "Inputdir" -Waarde $Inputdir -Color Cyan
Display_info -Omschrijving "Input file" -Waarde $Global:InputFile -Color Cyan
Display_info -Omschrijving "Outputdir" -Waarde $Outputdir -Color Cyan
Display_info -Omschrijving "Output file" -Waarde $Global:OutputFile -Color Cyan
Display_info -Omschrijving "Rollback" -Waarde $Global:Rollback -Color Red
Display_info -Omschrijving "Testrun" -Waarde $Global:Testrun -Color Red
Display_info -Omschrijving "Module ProjectFile : " -Waarde $Global:ModuleProjectFile -Color Magenta
Display_info -Omschrijving "Module Projectdirectory : " -Waarde $Global:ModuleProjectdir -Color Magenta
Display_info -Omschrijving "Author : " -Waarde $Global:Author -Color Magenta
Display_info -Omschrijving "Companyname : " -Waarde $Global:Companyname -Color Magenta
Display_info -Omschrijving "Module Naam : " -Waarde $Global:ModuleName -Color Yellow
Display_info -Omschrijving "Description : " -Waarde $Global:Description -Color Yellow
Display_info -Omschrijving "Functions directory : " -Waarde $Global:Functionsdir -Color Yellow
Display_info -Omschrijving "Modules directory : " -Waarde $Global:Modulesdir -Color Yellow
Display_info -Omschrijving "Module Version : " -Waarde $Global:Moduleversion -Color Yellow
Display_info -Omschrijving "Functions To Export : " -Waarde $Global:FunctionsToExport -Color Yellow

<#
ModuleProjectFile
ModuleProjectdir
$Global:ModuleNaam
$Global:Description
$Global:Functionsdir
$Global:Modulesdir 
$Global:Moduleversion 
$Global:FunctionsToExport
$Global:Author
$Global:Companyname
#>
#Display_info -Omschrijving "Process PO" -Waarde $ProcessPO -Color Yellow
Main

#Sleep -Seconds 30
$time = get-date -Format "HH:mm:ss dd-MM-yyyy"
$Message = "Time end".padright($tab," ") + ": $time"
Write-log -Category host -message $Message -color Green
$StopWatch.stop()
$message  = "Elapsed time ".padright($tab, " ") + ": {0:hh}:{0:mm}:{0:ss}" -f $($StopWatch.Elapsed)
Write-log -Category host -message $Message -color Yellow
If ($Global:DisplayLogfiles){
    If (test-path $VerboseFile){&notepad $Global:VerboseFile}
    If (test-path $LogFile){&notepad $LogFile}
    If (test-path $VoortgangFile){&notepad $VoortgangFile}
}

