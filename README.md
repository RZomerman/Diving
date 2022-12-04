# Diving

This is for the Mares Horizon only - to determine / calculate the required info, the following modules are available

#DivePlanRow
Calculates the required output for a single row in the Horizon Diveplanner, including NOAA and OTU limits

DivePlanRow -depth -minutes -gas- Fo2Setpoint -vo2

* Depth  - depth for the row to be filled
* Minutes  - minutes at that depth
* Gas - bottomgas (cylinder) or deco gas (cylinder)
* Fo2Setpoint - setpoint set in the system (bailout or bottom)
* Vo2 - vo2 set - usually 0.8 or 1.0



# Flow Rate Calculation
Calculate the flow rate in your system


CalculateFlowRate -vo2 -looppercentage -gaspercentage


* vo2 is your VO2 (usually 0.8 or 1.0)
* looppercentage is your setpoint 
* gaspercentage is what you have in your cylinder

# Max O2 Calculation


calculate the maximum amount of O2 you can have for your depth (at usually 1.4 or 1.6 - bailout)

MaxO2 -ppo2 -depth


* ppo2
* depth

# MaxDepth Calculation
Determines the maximum depth for a given gas and ppo2 value (usually 1.4 or 1.6 for bailout)
Use 1.6 with your cylinder bottom gas and use with 1.4 with your setpoint

* ppo2
* o2

# pp02 calculation

determines the pp02 at a given depth for a gax mixture

ppo2 -gas -depth

* gas
* depth

#EquivalentAirDepth 
Calculates the equivalent air depth with nitrox for your setpoint


EquivalentAirDepth -fo2 -depth

* fo2
* depth

#SACRate
Calculate your SAC rate

SACRate -PressureStart -PressureEnd -AverageDepth -Minutes -CylinderCapacity

*PressureStart
*PressureEnd
*AverageDepth
*Minutes
*CylinderCapacity



