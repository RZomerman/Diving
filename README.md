# Diving
Everything is without warranty - always doublecheck the calculations. You are responsible what you dive including your dive planning. Always double analyze and act on that.


This is for the Mares Horizon only - to determine / calculate the required info.

All calculations use meters and the digits for o2 % *(such as 40 and 32   - for 40 meters and 32 percent nitrox)


# V2 implements the Excel file automatic addition - note that this is still on a per line item, but it also allows you to "close" the excel file and perform all calculations for the Excel file as well.. I use iDecoPro for example to determine every line. 

Cells will be highlited for "ppo2" violations (above 1.6) 

# v2 working: 

copy the excel file into a new one (to keep the original clean)

copy .\HorizonPlanning.xlsx test10.xlsx

>open the excel file
>
$Sheet=OpenExcel c:\diving\test10.xlsx

DivePlanRow -depth 30 -minutes 20 -gas 32 -Fo2Setpoint 28 -vo2 0.8 -WorkSheet $Sheet

(add as required)

CloseExcel -WorkSheet $Excel

You will then see required litres etc.. 

# Bailout
Note that the system can only provide max depth + average bailout calculations (meaning no decompression = straight up) .. if you need to deco -- you need to manually calculate this all - based on mandatory deco!!

Add the lines for diving: 

the following modules are available

# DivePlanRow
Calculates the required output for a single row in the Horizon Diveplanner, including NOAA and OTU limits

DivePlanRow -depth -minutes -gas- Fo2Setpoint -vo2

* Depth  - depth for the row to be filled (in meters)
* Minutes  - minutes at that depth
* Gas - bottomgas (cylinder) or deco gas (cylinder)
* Fo2Setpoint - setpoint set in the system (bailout or bottom)
* Vo2 - vo2 set - usually 0.8 or 1.0



# Flow Rate Calculation
Calculate the flow rate in your system

CalculateFlowRate -vo2 -looppercentage -gaspercentage

* vo2 - is your VO2 (usually 0.8 or 1.0)
* looppercentage - is your setpoint 
* gaspercentage - is what you have in your cylinder

# Max O2 Calculation


calculate the maximum amount of O2 you can have for your depth (at usually 1.4 or 1.6 - bailout)

MaxO2 -ppo2 -depth


* ppo2 - ppo2 you want to use
* depth - depth in metres

# MaxDepth Calculation
Determines the maximum depth for a given gas and ppo2 value (usually 1.4 or 1.6 for bailout)
Use 1.6 with your cylinder bottom gas and use with 1.4 with your setpoint

* ppo2 - ppo2 to be used
* o2 - O2 to calculate with

# pp02 calculation

determines the pp02 at a given depth for a gax mixture

ppo2 -gas -depth

* gas - gas to be used
* depth - depth in meters

# EquivalentAirDepth 
Calculates the equivalent air depth with nitrox for your setpoint

EquivalentAirDepth -fo2 -depth

* fo2 - setpoint in loop or actual gas (bailout)
* depth - depth in meters

#SACRate
Calculate your SAC rate

SACRate -PressureStart -PressureEnd -AverageDepth -Minutes -CylinderCapacity

* PressureStart - pressure at start of measurement 
* PressureEnd - pressure at end of measurement
* AverageDepth - average depth from start to finish of measurement
* Minutes - minutes used for measurement
* CylinderCapacity - capacity of cilinder in litres



