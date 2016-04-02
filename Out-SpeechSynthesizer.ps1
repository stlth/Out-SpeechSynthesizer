<#PSScriptInfo
.VERSION 1.0.0
.GUID 8ea83384-9a96-48b3-b29e-ae98acf8c9f9
.AUTHOR Cory Calahan
.COMPANYNAME
.COPYRIGHT (C) Cory Calahan. All rights reserved.
.TAGS Speech
.LICENSEURI 
.PROJECTURI
   https://github.com/stlth/Out-SpeechSynthesizer
.ICONURI 
.EXTERNALMODULEDEPENDENCIES 
.REQUIREDSCRIPTS 
.EXTERNALSCRIPTDEPENDENCIES 
.RELEASENOTES
   Inspired by: https://gallery.technet.microsoft.com/scriptcenter/Out-Voice-1be16d5e

.Synopsis
   Allow PowerShell to speak back to you.
.DESCRIPTION
   Loads the System.Speech assemblies to allow you to pipeline information to an audible output.
.PARAMETER InputObject
   Data that will be spoken.
.PARAMETER Rate
   Sets the speaking rate of speech. Defaults to (0).
.PARAMETER Volume
   Sets the output volume. Defaults to (100).
.PARAMETER Mode
   Sets the speech to speak synchronously or asynchronously called. Defaults to (Synchronous).
.EXAMPLE
   "Hello world." | Out-SpeechSynthesizer
#>
function Out-SpeechSynthesizer
{
    [CmdletBinding(DefaultParameterSetName='Default', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  ConfirmImpact='Medium')]
    [Alias("Out-Voice")]
    Param
    (
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   Position=0,
                   ParameterSetName='Default')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Alias("Statement")] 
        [string[]]
        $InputObject,
        [Parameter(Mandatory=$false,
                   Position=1,
                   ParameterSetName='Default')]
        [ValidateRange(-10,10)]
        [int]
        $Rate=0,
        [Parameter(Mandatory=$false,
                   Position=2,
                   ParameterSetName='Default')]
        [ValidateRange(1,100)]
        [int]
        $Volume=100,
        [Parameter(Mandatory=$false,
                   Position=3,
                   ParameterSetName='Default')]
        [ValidateSet('Synchronous','Asynchronous')]
        [string]
        $Mode='Synchronous'
    )
    DynamicParam
    {
        # Set the dynamic parameters' name:
        $ParameterName = 'Voice'
        $RuntimeParameterDictionary = New-Object -TypeName System.Management.Automation.RuntimeDefinedParameterDictionary
        $AttributeCollection = New-Object -TypeName System.Collections.ObjectModel.Collection[System.Attribute]
        # Create and set the parameters' attributes:
        $ParameterAttribute = New-Object -TypeName System.Management.Automation.ParameterAttribute
        $ParameterAttribute.HelpMessage = 'Select a voice:'
        $ParameterAttribute.Mandatory = $false
        $ParameterAttribute.Position = 4
        $ParameterAttribute.ParameterSetName = 'Default'
        $AttributeCollection.Add($ParameterAttribute)
        # Generate the list of available voices and set the ValidateSet
        Add-Type -AssemblyName System.Speech
        $temp = New-Object -TypeName System.Speech.Synthesis.SpeechSynthesizer
        $voiceSet = $temp.GetInstalledVoices().VoiceInfo | Select-Object -ExpandProperty Name
        $temp.Dispose()
        Remove-Variable -Name temp
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($voiceSet)
        $AttributeCollection.Add($ValidateSetAttribute)
        # Create and return the dynamic parameter
        $RuntimeParameter = New-Object -TypeName System.Management.Automation.RuntimeDefinedParameter($ParameterName,[string],$AttributeCollection)
        $RuntimeParameterDictionary.Add($ParameterName, $RuntimeParameter)
        return $RuntimeParameterDictionary
    }
    Begin
    {
        # Bind the parameter to a new instance variable created in the Dynamic Parameter
        $VoiceSelected = $PsBoundParameters[$ParameterName]
        
        Write-Verbose -Message 'Listing Parameters utilized:'
        $PSBoundParameters.GetEnumerator() | ForEach-Object { Write-Verbose -Message "$($PSItem)" }
        
        Write-Verbose -Message 'Adding System.Speech assembly.'
        Add-Type -AssemblyName System.Speech
        $synth = New-Object -TypeName System.Speech.Synthesis.SpeechSynthesizer

        if ($PSBoundParameters['Rate'])
        {
            Write-Verbose -Message "Setting 'Rate' to: $Rate"
            $synth.Rate = $PSBoundParameters['Rate']
        }
        if ($PSBoundParameters['Volume'])
        {
            Write-Verbose -Message "Setting 'Volume' to: $Volume"
            $synth.Volume = $PSBoundParameters['Volume']
        }
        if ($PSBoundParameters['Voice'])
        {
            Write-Verbose -Message "Setting 'Voice' to: $($VoiceSelected)"
            $synth.SelectVoice($VoiceSelected)
        }
    } # END: Begin
    Process
    {
        if ($pscmdlet.ShouldProcess("Target", "Operation"))
        {
            foreach ($statement in $InputObject)
            {                
                switch($Mode)
                {
                    'Synchronous'
                    {
                        Write-Verbose -Message "'Synchronous' mode selected. Speaking: `'$statement`'"
                        $synth.Speak( $($statement | Out-String) ) | Out-Null
                    }
                    'Asynchronous'
                    {
                        Write-Verbose -Message "'Asynchronous' mode selected. Speaking: `'$statement`'"
                        $synth.SpeakAsync( $($statement | Out-String) ) | Out-Null
                    }
                default{}
                }
            }
        }
    } # END: Process
    End
    {
        $synth.Dispose()
    } # END: End
}