#Depth
#BottomGas
#DecoGas
#BottomSetpoint
#DecoSetpoint
#metabolic
#VO2


Function OpenExcel ($excelmasterfile){
    $excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $true
    $Excel.DisplayAlerts = $true # Disable comfirmation prompts
    $workbook  = $excel.Workbooks.Open($excelmasterfile)
    $worksheet = $workbook.Worksheets.Item("DivePlanning")
    return $worksheet
}

Function UpdateExcelRow($inputArray){
    #First need to get row without text (based on A column starting in row 3)
    $row=2
    $Cell=("A" + $row)
    While ($WorkSheet.Range($Cell).Text -ne ""){
        $row++
        $Cell=("A" + $row)
    }
    
    #Total DiveTime Calculation
    If ($row -ne 3) {
        $PreviousDiveTimeCell=("C" + ($row -1))
        [int]$PreviousDiveTime=$WorkSheet.Range($PreviousDiveTimeCell).Text
        $TotalDiveTime=$PreviousDiveTime + $inputArray.minutes
    }else{
        $TotalDiveTime=$inputArray.minutes
    }


    $worksheet.cells.item($row,1) = $inputArray.Depth
    $worksheet.cells.item($row,2) = $inputArray.minutes
    $worksheet.cells.item($row,3) = $TotalDiveTime
    $worksheet.cells.item($row,4) = $inputArray.bar 
    $worksheet.cells.item($row,5) = [int]($inputArray.gas * 100)
    $worksheet.cells.item($row,6) = $inputArray.Fo2
    $worksheet.cells.item($row,7) = $inputArray.SCRFlowRate
    $worksheet.cells.item($row,8) = $inputArray.SCRLiter
    $worksheet.cells.item($row,9) = $inputArray.PP02
    $worksheet.cells.item($row,10) = $inputArray.NOAADlimit
    $worksheet.cells.item($row,11) = $inputArray.NOAADPct
    $worksheet.cells.item($row,12) = $inputArray.NOAAD24limit
    $worksheet.cells.item($row,13) = $inputArray.NOAA24Pct
    $worksheet.cells.item($row,14) = $inputArray.OTUMinute
    $worksheet.cells.item($row,15) = $inputArray.OTUTotal
    If ($inputArray.PP02 -gt 1.6) {
        $Target=$worksheet.cells.item($row,9)
        $Target.Interior.ColorIndex =3
    }
    if (($inputArray.Fo2 + 0.03) -gt $inputArray.gas) {
        $Target=$worksheet.cells.item($row,6)
        $Target.Interior.ColorIndex =3
    }

    #BAILOUT PLAN
    #bailout plan starts at row 20 so always $row + 17 (as default starts at 3)
    If ($row -eq 3) {
        $brow=$row+17
        $worksheet.cells.item($brow,1) = $inputArray.Depth
        $worksheet.cells.item($brow,2) = "3" #3 minutes to sort shit out and then go up
        $worksheet.cells.item($brow,4) = $inputArray.bar 
        $worksheet.cells.item($brow,5) = [int]($inputArray.gas * 100)
        $worksheet.cells.item($brow,6) = $inputArray.Fo2
        $worksheet.cells.item($brow,7) = "30" #At failure - the SCR Flow rate is fixed to 30
        $worksheet.cells.item($brow,9) = "20" #SAC rate set to 20 which is a bit high
        $worksheet.cells.item($brow,12) = $inputArray.PP02
        $worksheet.cells.item($brow,13) = $inputArray.NOAADlimit
        $worksheet.cells.item($brow,14) = $inputArray.NOAADPct
        $worksheet.cells.item($brow,15) = $inputArray.NOAAD24limit
        $worksheet.cells.item($brow,16) = $inputArray.NOAA24Pct
        $worksheet.cells.item($brow,17) = $inputArray.OTUMinute
        $worksheet.cells.item($brow,18) = $inputArray.OTUTotal


        If ($inputArray.PP02 -gt 1.6) {
            $Target=$worksheet.cells.item($brow,11)
            $Target.Interior.ColorIndex =3
        }

        #BAILOUT CALCULATIONS
        $depth=$inputArray.Depth
        $AscentTime=[math]::Round(($depth/10),0)
        $averageDepth=$depth/2
        $averagebar= ($averageDepth /10) + 1
        $SCRLiterbailout=$AscentTime * 30
        $BailOutLitres=20*$AscentTime*$averagebar*1.5 #SAC RATE of 20

        $brow++
        $worksheet.cells.item($brow,1) = $averageDepth #average depth for recovering
        $worksheet.cells.item($brow,2) = $AscentTime #Average time to go to next depth
        $worksheet.cells.item($brow,4) = $averagebar
        $worksheet.cells.item($brow,5) = [int]($inputArray.gas * 100)
        $worksheet.cells.item($brow,6) = $inputArray.Fo2
        $worksheet.cells.item($brow,7) = "30" #At failure - the SCR Flow rate is fixed to 30
        $worksheet.cells.item($brow,9) = "20" #SAC rate set to 20 which is a bit high
        
        $ppo2=pp02 -gas ($inputArray.gas * 100) -depth $averageDepth
        
        $worksheet.cells.item($brow,12) = $ppo2

                $brow++
        $worksheet.cells.item($brow,1) = "ADD YOUR MANDATORY SAFETY STOP / DECO HERE"
        $worksheet.cells.item($brow,1) = "THIS SHEET USES NO DECO /SAFETY STOP AT ALL"


    }



}

Function CloseExcel(){
    $Gasmixxes= $WorkSheet.Range("E3:E11")
    [array]$Gasses=$Gasmixxes | where {$_.Text -ne "GAS MIX" -and $_.text -ne ""}
    $Gasses=$Gasses.text | Get-Unique
    If ($Gasses.count -gt 2) {
        Write-Host "too many gasses selected - manual calculation required"
    }else{
    #BOTTOM GAS CALCULATION
        [int]$gas1=$gasses[0]
        $worksheet.cells.item(14,1) = $gas1
        #Used Rows
        $RowsWithGas1=$Gasmixxes | where {$_.Text -eq [string]$gas1} | select Row
        $GasRequired=0
        ForEach ($row in $RowsWithGas1.row) {
            $Cell=("H" + $row)
            $GasRequired=$GasRequired+[int]$WorkSheet.Range($Cell).Text
        }
        $worksheet.cells.item(14,2) = $GasRequired
        $7LDeco= [math]::Round(($GasRequired /7),0)
        $11LDeco= [math]::Round(($GasRequired /11.1),0)
        $worksheet.cells.item(14,7) = $7LDeco
        $worksheet.cells.item(14,8) = $11LDeco

        If ($gasses.count -eq 2) {
    #DECO GAS CALCULATION
        [int]$gas2=$gasses[1]
        $worksheet.cells.item(15,1) = $gas2
        #Used Rows
        $RowsWithGas2=$Gasmixxes | where {$_.Text -eq [string]$gas2} | select Row
        $Gas2Required=0
        ForEach ($row in $RowsWithGas2.row) {
            $Cell=("H" + $row)
            $Gas2Required=$Gas2Required+[int]$WorkSheet.Range($Cell).Text
        }
        $worksheet.cells.item(15,2) = $Gas2Required
        $7LDeco= [math]::Round(($Gas2Required /7),0)
        $11LDeco= [math]::Round(($Gas2Required /11.1),0)
        $worksheet.cells.item(15,7) = $7LDeco
        $worksheet.cells.item(15,8) = $11LDeco

        }



    }
    #$excel.Workbooks.Close


}

Function CalculateFlowRate($vo2,$LoopPerc,$GasPerc, $Depth) {
    if (($LoopPerc +0.03) -gt $GasPerc) {
        write-host  "ERROR - LOOP MUST BE -3%" -ForegroundColor Red
        
    }
    if ($LoopPerc -gt 1) {$LoopPerc=$LoopPerc/100}
    if ($GasPerc -gt 1) {$GasPerc=$GasPerc/100}
    $FlowRate= ($vo2 *(1-$LoopPerc)) / ($GasPerc - $LoopPerc)
    #minimum of 5
    If ($SCRFlowRate -lt 5) {$SCRFlowRate=5}
    
    
    #Adjusted Flowrate at depth - as optional
    #AdjustedFlowWare = (IP at Depth / IP at the surface * flowrate)
    #IP at surface = 11  (depth bar + 10)
    #IP at depth = depth + 10
    if ($Depth) {
        $bar=($depth/10) + 1
        $IPatDepth=$bar + 10
        $IPatSurface=11
        $AdjustedFlowrate=($IPatDepth / $IPatSurface) * $FlowRate    
        $FlowRate=[math]::Round($AdjustedFlowrate,1)
    }else{
        $FlowRate=[math]::Round($FlowRate,1)
    }
    Return $FlowRate
}

Function CalculateO2($flowrate,$GasPerc,$vo2){
    #Calculating the O2 in the Loop based on 
    # Loop FO2 = [FlowRate * Cylinder FO2 - Vo2] / [FlowRate - vo2]
    if ($GasPerc -gt 1) {
        $GasPerc=$GasPerc/100
    }
    $FirstPartOfCalc = ($FlowRate * $GasPerc)   
    $SecondPartOfCalc = ($flowrate - $vo2)        
    $NetFO2InLoop = ($FirstPartOfCalc - $vo2) / $SecondPartOfCalc
    return $NetFO2InLoop
}

Function MaxO2($ppo2,$depth) {
    $bar=($depth/10) + 1
    $maxo2 = $ppo2 / $bar
    $maxO2Perc = $maxo2 * 100
    return $maxo2
}

Function MaxDepth($ppo2,$o2) {
    if ($o2 -gt 1) {
        $o2 = $o2 / 100
    }
    $maxbar = $ppo2 / $o2
    $MaxDepth = ($maxbar -1)*10
    return $MaxDepth
}

function pp02($gas,$depth){
    if ($depth -gt 9) {
        #we assume Metres converting to actual BAR
        $depth = ($depth /10) + 1
    }
    if ($gas -gt 1) {
        $gas = $gas / 100
    }
    $ppo2=$gas * $depth
    return $ppo2
}

Function EquivalentAirDepth ($FO2,$Depth) {
    if ($o2 -gt 1) {
        $o2 = $o2 / 100
    }
    $EDepth = ($depth + 10) / .79
    $EAD = (1- $o2) * $EDepth 
    $EAD = $EAD - 10
    $EAD=[math]::Round($EAD,1)
    return $EAD
}

Function SACRate($PressureStart,$PressureEnd,$AverageDepth,$Minutes,$CylinderCapacity) {
    if ($AverageDepth -gt 10) {
        #we assume Metres converting to actual BAR
        $AverageDepth = ($AverageDepth /10) + 1
    }
    $PressureDifference= $PressureStart-$PressureEnd
    $SACatTime =  ($PressureDifference * $CylinderCapacity) / $AverageDepth
    $SAC=$SACatTime / $Minutes
    return $SAC
}

#Depth 
#Time
#GASMix
#FO2
#SCR FlowRate
#SCRLiter = time * flow rate
#BAILOUTSAC
#OC LITERS REquired=time*bar*balout
#PP02 limit
#NOAA Single Dive limit (per PP02)
#% of dive limit
#NOAAL DAILY dive limit
#% of daily dive limit
#OTU Minute
#OTU total

Function DivePlanRow($depth,$minutes,$gas,$Fo2Setpoint,$vo2) {
    if (!($vo2)) {$vo2=0.8}
    if ($gas -gt 1) {
        $gas = $gas / 100
    }
    if ($Fo2Setpoint -gt 1) {
        $Fo2Setpoint = $Fo2Setpoint / 100
    }
    #we assume Metres converting to actual BAR
    $bar = ($depth /10) + 1
    
    $SCRFlowRate=CalculateFlowRate -vo2 $vo2 -LoopPerc $Fo2Setpoint -GasPerc $gas -Depth $depth



    #Now that we know the new flowrate - we can calculate FO2 again - based on that
    $Fo2Setpoint=CalculateO2 -flowrate $SCRFlowRate -GasPerc $gas -vo2 $vo2
    $Fo2Setpoint=[math]::Round($Fo2Setpoint,2)

    #write-host "FlowRate = " $SCRFlowRate
    $SCRLiter = $SCRFlowRate * $minutes
    $SCRLiter=[math]::Round($SCRLiter,0)
    #write-host "SCR Litres: " $SCRLiter
    #write-host "Bar:" $bar
    #write-host "Gas:" $Fo2Setpoint is being used as this is now the actual FO2 in the loop
    $ppo2 = ($bar * $Fo2Setpoint)
    #write-host "PPO2:" $ppo2
    $ppo2ForNOAA = [math]::Round($ppo2,1)
    switch ($ppo2ForNOAA)
        {
            0.6 {
                $Dlimit=720
                $24Limit=720
                }
            0.7 {
                $Dlimit=570
                $24Limit=570
                }
            0.8 {
                $Dlimit=450
                $24Limit=450
                }
            0.9 {
                $Dlimit=360
                $24Limit=360
            }
            1.0 {
                $Dlimit=300
                $24Limit=300
                }
            1.1 {
                $Dlimit=240
                $24Limit=270
                }
            1.2 {
                $Dlimit=219
                $24Limit=240
                }
            1.3 {
                $Dlimit=180
                $24Limit=210
                }
            1.4 {
                $Dlimit=150
                $24Limit=180
                }
            1.5 {
                $Dlimit=120
                $24Limit=180
                }
            1.6 {
                $Dlimit=45
                $24Limit=150
                }
        }
    $NOAASdivePercent=($minutes/$Dlimit)*100
    $NOAASdivePercent = [math]::Round($NOAASdivePercent,0)
    $NOAADayPercent=($minutes/$24Limit)*100
    $NOAADayPercent = [math]::Round($NOAADayPercent,0)
    $OTUMinute=OTU -ppo2 $ppo2ForNOAA
    $OTUTotal=$OTUMinute * $minutes


    #BAILOUT CALCULATIONS
        $AscentTime=[math]::Round(($depth/10),0)
        $averageDepth=$depth/2
        $averagebar= ($averageDepth /10) + 1
        $SCRLiterbailout=$AscentTime * 30
        $BailOutLitres=20*$AscentTime*$averagebar*1.5
        $totalNeededforBailout=$BailOutLitres+$SCRLiterbailout + 30 #the 30 is for the 1 minute of getting ready to go up


$OutputArray = @{}
$OutputArray.Depth = $depth
$OutputArray.minutes = $minutes
$OutputArray.bar = $bar
$OutputArray.gas = $gas
$OutputArray.Fo2 = $Fo2Setpoint
$OutputArray.SCRFlowRate = $SCRFlowRate
$OutputArray.SCRLiter = $SCRLiter
$OutputArray.PP02 = $ppo2ForNOAA
$OutputArray.NOAADlimit = $Dlimit
$OutputArray.NOAADPct = $NOAASdivePercent
$OutputArray.NOAAD24limit = $24Limit
$OutputArray.NOAA24Pct = $NOAADayPercent
$OutputArray.OTUMinute = $OTUMinute
$OutputArray.OTUTotal = $OTUTotal
$OutputArray.totalNeededforBailout=$totalNeededforBailout
$OutputArray.AscentTime=$AscentTime
$OutputArray.averageDepth=$averageDepth
$OutputArray.SCRLiterbailout=$SCRLiterbailout
$OutputArray.BailOutLitres=$BailOutLitres


$Objectname = New-Object PSobject -Property $OutputArray

$objectname #| select Depth,minutes,bar,gas,fo2,SCRFlowRate,SCRLiter,PP02,NOAADlimit,NOAADPct,NOAAD24limit,NOAA24Pct,OTUMinute,OTUTotal,totalNeededforBailout | fl
UpdateExcelRow $OutputArray
}

Function OTU($ppo2){
        switch ($ppo2)
        {
            0.5 {
                $OTU=0
                }
            0.6 {
                $OTU=0.27
                }
            0.7 {
                $OTU=0.47
                }
            0.8 {
                $OTU=0.65
                }
            0.9 {
                $OTU=0.83
                }
            1.0 {
                $OTU=1.0
                }
            1.1 {
                $OTU=1.16
                }
            1.2 {
                $OTU=1.32
                }
            1.3 {
                $OTU=1.48
                }
            1.4 {
                $OTU=1.63
                }
            1.5 {
                $OTU=1.78
                }
            1.6 {
                $OTU=1.92
                }
            1.7 {
                $OTU=2.07
                }
            1.8 {
                $OTU=2.21
                }
            1.9 {
                $OTU=2.35
                }
            2.0 {
                $OTU=2.49
                }
        }
        return $OTU
}