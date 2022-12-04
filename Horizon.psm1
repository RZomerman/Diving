#Depth
#BottomGas
#DecoGas
#BottomSetpoint
#DecoSetpoint
#metabolic
#VO2



Function CalculateFlowRate($vo2,$LoopPerc,$GasPerc) {
    $FlowRate= ($vo2 *(1-$LoopPerc)) / ($GasPerc - $LoopPerc)
    Return $FlowRate
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
    
    $SCRFlowRate=CalculateFlowRate -vo2 $vo2 -LoopPerc $Fo2Setpoint -GasPerc $gas

    #write-host "FlowRate = " $SCRFlowRate
    $SCRLiter = $SCRFlowRate * $minutes

    #write-host "SCR Litres: " $SCRLiter
    #write-host "Bar:" $bar
    #write-host "Gas:" $gas
    $ppo2 = ($bar * $gas)
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