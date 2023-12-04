;---------------------------------------------------------------------------
;                           Autoexecute Section
;---------------------------------------------------------------------------
#Include AE.V2.ahk
#Requires AutoHotkey v2

;---------------------------------------------------------------------------
;                              CHANGELOG
;---------------------------------------------------------------------------

; Version 6, Released 11/3/2023
; Separated the FMDS titles into a separate document from the common abbrevations.
; Removed Block Inputs and SendLevel(1).

; Version 5.0, Released 9/17/2023
; Converted script to V2. Moved Autocorrect stuff from this to the Autocorrect script
; as the Autocorrect script now detects duplicate items with the new Win+H hotkey script.
; This will prevent users from accidentally adding duplicate items. Also, moved OS/DS
; quick launch script from the main script to this script to make management for the 
; stand alone app easier as it will be able to share this script without modification.

; Version 4.0, Released 7/28/2023
; Reworked script based on new Horizon italics functionality so that standard control keys
; will work. Some items have been moved to Quick Access Pop-up. Updated name for OS 3-36.

; Version 3.4 - Released 5/30/2023
; Changed "sgsv" to "sgsvf" based on user feedback. Added "degf" for addeding a degree symbol.

; Version 3.3 - Released 5/3/2023
; Added case specific hot string areas to reduce rework on strange cases.

; Version 3.2 - Released 4/27/2023
; Updated OS 7-11 title.

; Version 3.1 - Released 4/11/2023
; Updated OS 10-1 title. Improved fast launch of OS GUI.

; Version 3 - 7 Mar 2023
; OS Titles now input in italics except for Horizon.

; Version 2.0 - Released 3/6/2023
; Added the Shift+Ctrl+WIN+O hot key for fast launch of a specific OS.;

; Version 1.0 - Released 2/24/2023
; Modified to correct a few standard titles and add the "densityf" hotstring with input for the varibles.
; Also modified to break this script out into a sub-script to make maintenance easier for users, so they
; are not forced to update the starter script with all the personal information, which likely won't change
; as much as the sub-scripts.

;---------------------------------------------------------------------------
;                   FM Global Standard Abbreviations
;---------------------------------------------------------------------------

:*:1-0f::
{ 
Send("1-0, ")
Send("^i")
Sleep(100)
Send("Safeguards During Construction, Alteration and Demolition")
Sleep(100)
Send("^i")
}
return

:*:1-1f::
{ 
Send("1-1, ")
Send("^i")
Sleep(100)
Send("Firesafe Building Construction and Materials")
Sleep(100)
Send("^i")
}
return

:*:1-2f::
{ 
Send("1-2, ")
Send("^i")
Sleep(100)
Send("Earthquakes")
Sleep(100)
Send("^i")
}
return

:*:1-3f::
{ 
Send("1-3, ")
Send("^i")
Sleep(100)
Send("High-Rise Buildings")
Sleep(100)
Send("^i")
}
return

:*:1-4f::
{ 
Send("1-4, ")
Send("^i")
Sleep(100)
Send("Fire Tests")
Sleep(100)
Send("^i")
}
return

:*:1-6f::
{ 
Send("1-6, ")
Send("^i")
Sleep(100)
Send("Cooling Towers")
Sleep(100)
Send("^i")
}
return

:*:1-8f::
{ 
Send("1-8, ")
Send("^i")
Sleep(100)
Send("Antenna Towers and Signs")
Sleep(100)
Send("^i")
}
return

:*:1-9f::
{ 
Send("1-9, ")
Send("^i")
Sleep(100)
Send("Roof Anchorage for Older, Wood-Roofed Buildings")
Sleep(100)
Send("^i")
}
return

:*:1-10f::
{ 
Send("1-10, ")
Send("^i")
Sleep(100)
Send("Smoke and Heat Venting in Sprinklered Buildings")
Sleep(100)
Send("^i")
}
return

:*:1-11f::
{ 
Send("1-11, ")
Send("^i")
Sleep(100)
Send("Fire Following Earthquakes")
Sleep(100)
Send("^i")
}
return

:*:1-12f::
{ 
Send("1-12, ")
Send("^i")
Sleep(100)
Send("Ceilings and Concealed Spaces")
Sleep(100)
Send("^i")
}
return

:*:1-13f::
{ 
Send("1-13, ")
Send("^i")
Sleep(100)
Send("Chimneys")
Sleep(100)
Send("^i")
}
return

:*:1-15f::
{ 
Send("1-15, ")
Send("^i")
Sleep(100)
Send("Roof Mounted Solar Photovoltaic Panels")
Sleep(100)
Send("^i")
}
return

:*:1-17f::
{ 
Send("1-17, ")
Send("^i")
Sleep(100)
Send("Reflective Ceiling Insulation")
Sleep(100)
Send("^i")
}
return

:*:1-20f::
{ 
Send("1-20, ")
Send("^i")
Sleep(100)
Send("Protection Against Exterior Fire Exposure")
Sleep(100)
Send("^i")
}
return

:*:1-21f::
{ 
Send("1-21, ")
Send("^i")
Sleep(100)
Send("Fire Resistance of Building Assemblies")
Sleep(100)
Send("^i")
}
return

:*:1-22f::
{ 
Send("1-22, ")
Send("^i")
Sleep(100)
Send("Maximum Foreseeable Loss")
Sleep(100)
Send("^i")
}
return

:*:1-24f::
{ 
Send("1-24, ")
Send("^i")
Sleep(100)
Send("Protection Against Liquid Damage")
Sleep(100)
Send("^i")
}
return

:*:1-26f::
{ 
Send("1-26, ")
Send("^i")
Sleep(100)
Send("Steep-Slope Roof Systems")
Sleep(100)
Send("^i")
}
return

:*:1-27f::
{ 
Send("1-27, ")
Send("^i")
Sleep(100)
Send("Windstorm  Retrofit and Loss Expectancy (LE) Guidelines")
Sleep(100)
Send("^i")
}
return

:*:1-28f::
{ 
Send("1-28, ")
Send("^i")
Sleep(100)
Send("Wind Design")
Sleep(100)
Send("^i")
}
return

:*:1-29f::
{ 
Send("1-29, ")
Send("^i")
Sleep(100)
Send("Roof Deck Securement and Above-Deck Roofing  Components")
Sleep(100)
Send("^i")
}
return

:*:1-30f::
{ 
Send("1-30, ")
Send("^i")
Sleep(100)
Send("Repair of Wind Damaged Roof Systems")
Sleep(100)
Send("^i")
}
return

:*:1-31f::
{ 
Send("1-31, ")
Send("^i")
Sleep(100)
Send("Panel Roof Systems")
Sleep(100)
Send("^i")
}
return

:*:1-32f::
{ 
Send("1-32, ")
Send("^i")
Sleep(100)
Send("Inspection and Maintenance of Roof Assemblies")
Sleep(100)
Send("^i")
}
return

:*:1-33f::
{ 
Send("1-33, ")
Send("^i")
Sleep(100)
Send("Safeguarding Torch - Applied Roof Installations")
Sleep(100)
Send("^i")
}
return

:*:1-34f::
{ 
Send("1-34, ")
Send("^i")
Sleep(100)
Send("Hail Damage")
Sleep(100)
Send("^i")
}
return

:*:1-35f::
{ 
Send("1-35, ")
Send("^i")
Sleep(100)
Send("Vegetative Roof Systems Vegetative Roof Systems, Occupied Roof Areas & Decks")
Sleep(100)
Send("^i")
}
return

:*:1-36f::
{ 
Send("1-36, ")
Send("^i")
Sleep(100)
Send("Mass Engineered Timber")
Sleep(100)
Send("^i")
}
return

:*:1-37f::
{ 
Send("1-37, ")
Send("^i")
Sleep(100)
Send("Hospitals")
Sleep(100)
Send("^i")
}
return

:*:1-40f::
{ 
Send("1-40, ")
Send("^i")
Sleep(100)
Send("Flood")
Sleep(100)
Send("^i")
}
return

:*:1-41f::
{ 
Send("1-41, ")
Send("^i")
Sleep(100)
Send("Dam Evaluations")
Sleep(100)
Send("^i")
}
return

:*:1-42f::
{ 
Send("1-42, ")
Send("^i")
Sleep(100)
Send("MFL Limiting Factors")
Sleep(100)
Send("^i")
}
return

:*:1-44f::
{ 
Send("1-44, ")
Send("^i")
Sleep(100)
Send("Damage-Limiting Construction")
Sleep(100)
Send("^i")
}
return

:*:1-45f::
{ 
Send("1-45, ")
Send("^i")
Sleep(100)
Send("Air Conditioning and Ventilating Systems")
Sleep(100)
Send("^i")
}
return

:*:1-49f::
{ 
Send("1-49, ")
Send("^i")
Sleep(100)
Send("Perimeter Flashing")
Sleep(100)
Send("^i")
}
return

:*:1-52f::
{ 
Send("1-52, ")
Send("^i")
Sleep(100)
Send("Field  Verification of Roof Wind Uplift Resistance")
Sleep(100)
Send("^i")
}
return

:*:1-53f::
{ 
Send("1-53, ")
Send("^i")
Sleep(100)
Send("Anechoic Chambers")
Sleep(100)
Send("^i")
}
return

:*:1-54f::
{ 
Send("1-54, ")
Send("^i")
Sleep(100)
Send("Roof Loads and Drainage")
Sleep(100)
Send("^i")
}
return

:*:1-55f::
{ 
Send("1-55, ")
Send("^i")
Sleep(100)
Send("Weak Construction and Design")
Sleep(100)
Send("^i")
}
return

:*:1-56f::
{ 
Send("1-56, ")
Send("^i")
Sleep(100)
Send("Cleanrooms")
Sleep(100)
Send("^i")
}
return

:*:1-57f::
{ 
Send("1-57, ")
Send("^i")
Sleep(100)
Send("Plastics in Construction")
Sleep(100)
Send("^i")
}
return

:*:1-59f::
{ 
Send("1-59, ")
Send("^i")
Sleep(100)
Send("Fabric and Membrane Covered Structures")
Sleep(100)
Send("^i")
}
return

:*:1-60f::
{ 
Send("1-60, ")
Send("^i")
Sleep(100)
Send("Asphalt-Coated/Protected Metal Buildings")
Sleep(100)
Send("^i")
}
return

:*:1-61f::
{ 
Send("1-61, ")
Send("^i")
Sleep(100)
Send("Impregnated Fire-Retardant Lumber")
Sleep(100)
Send("^i")
}
return

:*:1-62f::
{ 
Send("1-62, ")
Send("^i")
Sleep(100)
Send("Cranes")
Sleep(100)
Send("^i")
}
return

:*:1-63f::
{ 
Send("1-63, ")
Send("^i")
Sleep(100)
Send("Exterior Insulation and Finish Systems")
Sleep(100)
Send("^i")
}
return

:*:1-64f::
{ 
Send("1-64, ")
Send("^i")
Sleep(100)
Send("Exterior Walls and Facades")
Sleep(100)
Send("^i")
}
return

:*:2-0f::
{ 
Send("2-0, ")
Send("^i")
Sleep(100)
Send("Installation Guidelines for Automatic Sprinklers")
Sleep(100)
Send("^i")
}
return

:*:2-1f::
{ 
Send("2-1, ")
Send("^i")
Sleep(100)
Send("Corrosion in Automatic Sprinkler Systems")
Sleep(100)
Send("^i")
}
return

:*:2-8f::
{ 
Send("2-8, ")
Send("^i")
Sleep(100)
Send("Earthquake Protection for Water-Based Fire Protection Systems")
Sleep(100)
Send("^i")
}
return

:*:2-81f::
{ 
Send("2-81, ")
Send("^i")
Sleep(100)
Send("Fire Protection System Inspection, Testing and Maintenance")
Sleep(100)
Send("^i")
}
return

:*:2-89f::
{ 
Send("2-89, ")
Send("^i")
Sleep(100)
Send("Pipe Friction Loss Tables")
Sleep(100)
Send("^i")
}
return

:*:3-0f::
{ 
Send("3-0, ")
Send("^i")
Sleep(100)
Send("Hydraulics of Fire Protection Systems")
Sleep(100)
Send("^i")
}
return

:*:3-1f::
{ 
Send("3-1, ")
Send("^i")
Sleep(100)
Send("Tanks and Reservoirs for Interconnected Fire Service and Public Mains")
Sleep(100)
Send("^i")
}
return

:*:3-2f::
{ 
Send("3-2, ")
Send("^i")
Sleep(100)
Send("Water Tanks for Fire Protection")
Sleep(100)
Send("^i")
}
return

:*:3-3f::
{ 
Send("3-3, ")
Send("^i")
Sleep(100)
Send("Cross Connections")
Sleep(100)
Send("^i")
}
return

:*:3-4f::
{ 
Send("3-4, ")
Send("^i")
Sleep(100)
Send("Embankment-Supported Fabric Tanks")
Sleep(100)
Send("^i")
}
return

:*:3-6f::
{ 
Send("3-6, ")
Send("^i")
Sleep(100)
Send("Lined Earth Reservoirs for Fire Protection")
Sleep(100)
Send("^i")
}
return

:*:3-7f::
{ 
Send("3-7, ")
Send("^i")
Sleep(100)
Send("Fire Protection Pump")
Sleep(100)
Send("^i")
}
return

:*:3-10f::
{ 
Send("3-10, ")
Send("^i")
Sleep(100)
Send("Installation/Maintenance of  Fire Service Mains")
Sleep(100)
Send("^i")
}
return

:*:3-11f::
{ 
Send("3-11, ")
Send("^i")
Sleep(100)
Send("Flow and Pressure Regulating Devices for Fire Protection Service")
Sleep(100)
Send("^i")
}
return

:*:3-26f::
{ 
Send("3-26, ")
Send("^i")
Sleep(100)
Send("Fire Protection for Nonstorage Occupancies")
Sleep(100)
Send("^i")
}
return

:*:3-29f::
{ 
Send("3-29, ")
Send("^i")
Sleep(100)
Send("Reliability of Fire Protection Water Supplies")
Sleep(100)
Send("^i")
}
return

:*:4-0f::
{ 
Send("4-0, ")
Send("^i")
Sleep(100)
Send("Special Protection Systems")
Sleep(100)
Send("^i")
}
return

:*:4-1Nf::
{ 
Send("4-1N, ")
Send("^i")
Sleep(100)
Send("Fixed Water Spray Systems for Fire Protection")
Sleep(100)
Send("^i")
}
return

:*:4-2f::
{ 
Send("4-2, ")
Send("^i")
Sleep(100)
Send("Water Mist Systems")
Sleep(100)
Send("^i")
}
return

:*:4-3Nf::
{ 
Send("4-3N, ")
Send("^i")
Sleep(100)
Send("Medium and High Expansion Foam Systems")
Sleep(100)
Send("^i")
}
return

:*:4-4Nf::
{ 
Send("4-4N, ")
Send("^i")
Sleep(100)
Send("Standpipe and Hose Systems")
Sleep(100)
Send("^i")
}
return

:*:4-5f::
{ 
Send("4-5, ")
Send("^i")
Sleep(100)
Send("Portable Extinguishers")
Sleep(100)
Send("^i")
}
return

:*:4-6f::
{ 
Send("4-6, ")
Send("^i")
Sleep(100)
Send("Hybrid Fire Extinguishing Systems")
Sleep(100)
Send("^i")
}
return

:*:4-8Nf::
{ 
Send("4-8N, ")
Send("^i")
Sleep(100)
Send("Halon 1301 Extinguishing Systems")
Sleep(100)
Send("^i")
}
return

:*:4-9f::
{ 
Send("4-9, ")
Send("^i")
Sleep(100)
Send("Halocarbon and Inert Gas (Clean Agent) Fire Extinguishing Systems")
Sleep(100)
Send("^i")
}
return

:*:4-10f::
{ 
Send("4-10, ")
Send("^i")
Sleep(100)
Send("Dry Chemical Systems")
Sleep(100)
Send("^i")
}
return

:*:4-11Nf::
{ 
Send("4-11N, ")
Send("^i")
Sleep(100)
Send("Carbon Dioxide Extinguishing Systems (NFPA)")
Sleep(100)
Send("^i")
}
return

:*:4-12f::
{ 
Send("4-12, ")
Send("^i")
Sleep(100)
Send("Foam Water Extinguishing  Systems")
Sleep(100)
Send("^i")
}
return

:*:4-13f::
{ 
Send("4-13, ")
Send("^i")
Sleep(100)
Send("Oxygen Reduction Systems")
Sleep(100)
Send("^i")
}
return

:*:5-1f::
{ 
Send("5-1, ")
Send("^i")
Sleep(100)
Send("Electrical Equipment in Hazardous (Classified) Locations")
Sleep(100)
Send("^i")
}
return

:*:5-3f::
{ 
Send("5-3, ")
Send("^i")
Sleep(100)
Send("Hydroelectric Power Plants")
Sleep(100)
Send("^i")
}
return

:*:5-4f::
{ 
Send("5-4, ")
Send("^i")
Sleep(100)
Send("Transformers")
Sleep(100)
Send("^i")
}
return

:*:5-8f::
{ 
Send("5-8, ")
Send("^i")
Sleep(100)
Send("Static Electricity")
Sleep(100)
Send("^i")
}
return

:*:5-11f::
{ 
Send("5-11, ")
Send("^i")
Sleep(100)
Send("Lightning and Surge Protection for Electrical Systems")
Sleep(100)
Send("^i")
}
return

:*:5-12f::
{ 
Send("5-12, ")
Send("^i")
Sleep(100)
Send("Electric AC Generators")
Sleep(100)
Send("^i")
}
return

:*:5-14f::
{ 
Send("5-14, ")
Send("^i")
Sleep(100)
Send("Telecommunications")
Sleep(100)
Send("^i")
}
return

:*:5-17f::
{ 
Send("5-17, ")
Send("^i")
Sleep(100)
Send("Motors & Adjustable Speed Drives")
Sleep(100)
Send("^i")
}
return

:*:5-18f::
{ 
Send("5-18, ")
Send("^i")
Sleep(100)
Send("Protection of Electrical Equipment")
Sleep(100)
Send("^i")
}
return

:*:5-19f::
{ 
Send("5-19, ")
Send("^i")
Sleep(100)
Send("Switchgear and Circuit Breakers")
Sleep(100)
Send("^i")
}
return

:*:5-20f::
{ 
Send("5-20, ")
Send("^i")
Sleep(100)
Send("Electrical Testing")
Sleep(100)
Send("^i")
}
return

:*:5-21f::
{ 
Send("5-21, ")
Send("^i")
Sleep(100)
Send("Metal Halide High-Intensity Discharge Lighting")
Sleep(100)
Send("^i")
}
return

:*:5-23f::
{ 
Send("5-23, ")
Send("^i")
Sleep(100)
Send("Design and Fire Protection for Emergency and Standby Power Systems")
Sleep(100)
Send("^i")
}
return

:*:5-24f::
{ 
Send("5-24, ")
Send("^i")
Sleep(100)
Send("Miscellaneous Electrical Equipment")
Sleep(100)
Send("^i")
}
return

:*:5-25f::
{ 
Send("5-25, ")
Send("^i")
Sleep(100)
Send("High Voltage Direct Current Converter  Stations")
Sleep(100)
Send("^i")
}
return

:*:5-28f::
{ 
Send("5-28, ")
Send("^i")
Sleep(100)
Send("DC Battery Systems")
Sleep(100)
Send("^i")
}
return

:*:5-30f::
{ 
Send("5-30, ")
Send("^i")
Sleep(100)
Send("Power Factor Correction and Static Reactive Compensator Systems")
Sleep(100)
Send("^i")
}
return

:*:5-31f::
{ 
Send("5-31, ")
Send("^i")
Sleep(100)
Send("Cables and Bus Bars")
Sleep(100)
Send("^i")
}
return

:*:5-32f::
{ 
Send("5-32, ")
Send("^i")
Sleep(100)
Send("Data Centers and Related Facilities")
Sleep(100)
Send("^i")
}
return

:*:5-33f::
{ 
Send("5-33, ")
Send("^i")
Sleep(100)
Send("Electrical Energy Storage Systems")
Sleep(100)
Send("^i")
}
return

:*:5-40f::
{ 
Send("5-40, ")
Send("^i")
Sleep(100)
Send("Fire Alarm Systems")
Sleep(100)
Send("^i")
}
return

:*:5-48f::
{ 
Send("5-48, ")
Send("^i")
Sleep(100)
Send("Automatic Fire Detection")
Sleep(100)
Send("^i")
}
return

:*:5-49f::
{ 
Send("5-49, ")
Send("^i")
Sleep(100)
Send("Gas and Vapor Detectors and Analysis Systems")
Sleep(100)
Send("^i")
}
return

:*:6-0f::
{ 
Send("6-0, ")
Send("^i")
Sleep(100)
Send("Elements of Industrial Heating Equipment")
Sleep(100)
Send("^i")
}
return

:*:6-2f::
{ 
Send("6-2, ")
Send("^i")
Sleep(100)
Send("Pulverized Coal Fired  Boilers")
Sleep(100)
Send("^i")
}
return

:*:6-3f::
{ 
Send("6-3, ")
Send("^i")
Sleep(100)
Send("Induction and Dielectric Heating Equipment")
Sleep(100)
Send("^i")
}
return

:*:6-4f::
{ 
Send("6-4, ")
Send("^i")
Sleep(100)
Send("Oil- and Gas-Fired Single-Burner Boilers")
Sleep(100)
Send("^i")
}
return

:*:6-5f::
{ 
Send("6-5, ")
Send("^i")
Sleep(100)
Send("Oil- or Gas-Fired Multiple Burner Boilers")
Sleep(100)
Send("^i")
}
return

:*:6-6f::
{ 
Send("6-6, ")
Send("^i")
Sleep(100)
Send("Boiler-Furnaces Implosions")
Sleep(100)
Send("^i")
}
return

:*:6-7f::
{ 
Send("6-7, ")
Send("^i")
Sleep(100)
Send("Fluidized Bed Combustors")
Sleep(100)
Send("^i")
}
return

:*:6-8f::
{ 
Send("6-8, ")
Send("^i")
Sleep(100)
Send("Combustion Air Heaters")
Sleep(100)
Send("^i")
}
return

:*:6-9f::
{ 
Send("6-9, ")
Send("^i")
Sleep(100)
Send("Industrial Ovens and Dryers")
Sleep(100)
Send("^i")
}
return

:*:6-10f::
{ 
Send("6-10, ")
Send("^i")
Sleep(100)
Send("Process Furnaces")
Sleep(100)
Send("^i")
}
return

:*:6-11f::
{ 
Send("6-11, ")
Send("^i")
Sleep(100)
Send("Thermal and Regenerative Catalytic Oxidizers")
Sleep(100)
Send("^i")
}
return

:*:6-12f::
{ 
Send("6-12, ")
Send("^i")
Sleep(100)
Send("Low-Water Protection")
Sleep(100)
Send("^i")
}
return

:*:6-13f::
{ 
Send("6-13, ")
Send("^i")
Sleep(100)
Send("Waste Fuel Fired Facilities")
Sleep(100)
Send("^i")
}
return

:*:6-14f::
{ 
Send("6-14, ")
Send("^i")
Sleep(100)
Send("Waste Heat Boilers")
Sleep(100)
Send("^i")
}
return

:*:6-17f::
{ 
Send("6-17, ")
Send("^i")
Sleep(100)
Send("Rotary Kilns and Dryers")
Sleep(100)
Send("^i")
}
return

:*:6-21f::
{ 
Send("6-21, ")
Send("^i")
Sleep(100)
Send("Chemical Recovery Boilers")
Sleep(100)
Send("^i")
}
return

:*:6-22f::
{ 
Send("6-22, ")
Send("^i")
Sleep(100)
Send("Firetube Boilers")
Sleep(100)
Send("^i")
}
return

:*:6-23f::
{ 
Send("6-23, ")
Send("^i")
Sleep(100)
Send("Watertube Boilers")
Sleep(100)
Send("^i")
}
return

:*:7-0f::
{ 
Send("7-0, ")
Send("^i")
Sleep(100)
Send("Causes and Effects of Fires and Explosions")
Sleep(100)
Send("^i")
}
return

:*:7-1f::
{ 
Send("7-1, ")
Send("^i")
Sleep(100)
Send("Fire Protection for Textile Mills")
Sleep(100)
Send("^i")
}
return

:*:7-2f::
{ 
Send("7-2, ")
Send("^i")
Sleep(100)
Send("Waste Solvent Recovery")
Sleep(100)
Send("^i")
}
return

:*:7-3f::
{ 
Send("7-3, ")
Send("^i")
Sleep(100)
Send("Flight Simulator System Protection")
Sleep(100)
Send("^i")
}
return

:*:7-4f::
{ 
Send("7-4, ")
Send("^i")
Sleep(100)
Send("Paper Machines and Pulp Dryers")
Sleep(100)
Send("^i")
}
return

:*:7-6f::
{ 
Send("7-6, ")
Send("^i")
Sleep(100)
Send("Heated Plastic and Plastic-Lined Tanks")
Sleep(100)
Send("^i")
}
return

:*:7-7f::
{ 
Send("7-7, ")
Send("^i")
Sleep(100)
Send("Semiconductor Fabrication Facilities")
Sleep(100)
Send("^i")
}
return

:*:7-9f::
{ 
Send("7-9, ")
Send("^i")
Sleep(100)
Send("Dip Tanks, Flow Coaters and Roll Coaters")
Sleep(100)
Send("^i")
}
return

:*:7-10f::
{ 
Send("7-10, ")
Send("^i")
Sleep(100)
Send("Wood Processing and Woodworking Facilities")
Sleep(100)
Send("^i")
}
return

:*:7-11f::
{ 
Send("7-11, ")
Send("^i")
Sleep(100)
Send("Conveyors")
Sleep(100)
Send("^i")
}
return

:*:7-12f::
{ 
Send("7-12, ")
Send("^i")
Sleep(100)
Send("Mining and Mineral Processing")
Sleep(100)
Send("^i")
}
return

:*:7-13f::
{ 
Send("7-13, ")
Send("^i")
Sleep(100)
Send("Mechanical Refrigeration")
Sleep(100)
Send("^i")
}
return

:*:7-14f::
{ 
Send("7-14, ")
Send("^i")
Sleep(100)
Send("Fire Protection for Chemical Plants")
Sleep(100)
Send("^i")
}
return

:*:7-15f::
{ 
Send("7-15, ")
Send("^i")
Sleep(100)
Send("Garages")
Sleep(100)
Send("^i")
}
return

:*:7-17f::
{ 
Send("7-17, ")
Send("^i")
Sleep(100)
Send("Explosion Suppression Systems")
Sleep(100)
Send("^i")
}
return

:*:7-19Nf::
{ 
Send("7-19N, ")
Send("^i")
Sleep(100)
Send("Fire Hazard Properties of Flammable Liquids, Gases, and Volatile Solids")
Sleep(100)
Send("^i")
}
return

:*:7-20f::
{ 
Send("7-20, ")
Send("^i")
Sleep(100)
Send("Oil Cookers")
Sleep(100)
Send("^i")
}
return

:*:7-21f::
{ 
Send("7-21, ")
Send("^i")
Sleep(100)
Send("Rolling Mills")
Sleep(100)
Send("^i")
}
return

:*:7-22f::
{ 
Send("7-22, ")
Send("^i")
Sleep(100)
Send("Hydrazine and its Derivatives")
Sleep(100)
Send("^i")
}
return

:*:7-23f::
{ 
Send("7-23, ")
Send("^i")
Sleep(100)
Send("Data on General Class of Chemicals")
Sleep(100)
Send("^i")
}
return

:*:7-23Nf::
{ 
Send("7-23N, ")
Send("^i")
Sleep(100)
Send("Hazardous Chemicals Data")
Sleep(100)
Send("^i")
}
return

:*:7-24f::
{ 
Send("7-24, ")
Send("^i")
Sleep(100)
Send("Blowing Agents")
Sleep(100)
Send("^i")
}
return

:*:7-25f::
{ 
Send("7-25, ")
Send("^i")
Sleep(100)
Send("Molten Steel Production")
Sleep(100)
Send("^i")
}
return

:*:7-26f::
{ 
Send("7-26, ")
Send("^i")
Sleep(100)
Send("Glass Manufacturing")
Sleep(100)
Send("^i")
}
return

:*:7-27f::
{ 
Send("7-27, ")
Send("^i")
Sleep(100)
Send("Spray Application of Ignitable and Combustible Materials")
Sleep(100)
Send("^i")
}
return

:*:7-28f::
{ 
Send("7-28, ")
Send("^i")
Sleep(100)
Send("Energetic Materials")
Sleep(100)
Send("^i")
}
return

:*:7-29f::
{ 
Send("7-29, ")
Send("^i")
Sleep(100)
Send("Ignitable Liquid Storage in Portable Containers")
Sleep(100)
Send("^i")
}
return

:*:7-31f::
{ 
Send("7-31, ")
Send("^i")
Sleep(100)
Send("Storage of Aerosol Products")
Sleep(100)
Send("^i")
}
return

:*:7-32f::
{ 
Send("7-32, ")
Send("^i")
Sleep(100)
Send("Ignitable Liquid Operations")
Sleep(100)
Send("^i")
}
return

:*:7-33f::
{ 
Send("7-33, ")
Send("^i")
Sleep(100)
Send("High-Temperature Molten Materials")
Sleep(100)
Send("^i")
}
return

:*:7-35f::
{ 
Send("7-35, ")
Send("^i")
Sleep(100)
Send("Air Separation Processes")
Sleep(100)
Send("^i")
}
return

:*:7-36f::
{ 
Send("7-36, ")
Send("^i")
Sleep(100)
Send("Pharmaceutical Operations")
Sleep(100)
Send("^i")
}
return

:*:7-37f::
{ 
Send("7-37, ")
Send("^i")
Sleep(100)
Send("Cutting Fluids")
Sleep(100)
Send("^i")
}
return

:*:7-39f::
{ 
Send("7-39, ")
Send("^i")
Sleep(100)
Send("Industrial Trucks")
Sleep(100)
Send("^i")
}
return

:*:7-40f::
{ 
Send("7-40, ")
Send("^i")
Sleep(100)
Send("Heavy Duty Mobile Equipment")
Sleep(100)
Send("^i")
}
return

:*:7-41f::
{ 
Send("7-41, ")
Send("^i")
Sleep(100)
Send("Heat Treating of Materials Using Oil Quenching and Molten Salt Baths")
Sleep(100)
Send("^i")
}
return

:*:7-42f::
{ 
Send("7-42, ")
Send("^i")
Sleep(100)
Send("Vapor Cloud Explosions")
Sleep(100)
Send("^i")
}
return

:*:7-43f::
{ 
Send("7-43, ")
Send("^i")
Sleep(100)
Send("Process Safety")
Sleep(100)
Send("^i")
}
return

:*:7-45f::
{ 
Send("7-45, ")
Send("^i")
Sleep(100)
Send("Instrumentation and Control in Safety Applications")
Sleep(100)
Send("^i")
}
return

:*:7-46f::
{ 
Send("7-46, ")
Send("^i")
Sleep(100)
Send("Chemical Reactors and Reactions")
Sleep(100)
Send("^i")
}
return

:*:7-49f::
{ 
Send("7-49, ")
Send("^i")
Sleep(100)
Send("Emergency Venting of Vessels")
Sleep(100)
Send("^i")
}
return

:*:7-50f::
{ 
Send("7-50, ")
Send("^i")
Sleep(100)
Send("Compressed Gases in Cylinders")
Sleep(100)
Send("^i")
}
return

:*:7-51f::
{ 
Send("7-51, ")
Send("^i")
Sleep(100)
Send("Acetylene")
Sleep(100)
Send("^i")
}
return

:*:7-53f::
{ 
Send("7-53, ")
Send("^i")
Sleep(100)
Send("Liquefied Natural Gas (LNG)")
Sleep(100)
Send("^i")
}
return

:*:7-54f::
{ 
Send("7-54, ")
Send("^i")
Sleep(100)
Send("Natural Gas and Gas Piping")
Sleep(100)
Send("^i")
}
return

:*:7-55f::
{ 
Send("7-55, ")
Send("^i")
Sleep(100)
Send("Liquefied Petroleum Gas")
Sleep(100)
Send("^i")
}
return

:*:7-58f::
{ 
Send("7-58, ")
Send("^i")
Sleep(100)
Send("Chlorine Dioxide")
Sleep(100)
Send("^i")
}
return

:*:7-59f::
{ 
Send("7-59, ")
Send("^i")
Sleep(100)
Send("Inerting and Purging of Vessels and Equipment")
Sleep(100)
Send("^i")
}
return

:*:7-61f::
{ 
Send("7-61, ")
Send("^i")
Sleep(100)
Send("Facilities Processing Radioactive Materials")
Sleep(100)
Send("^i")
}
return

:*:7-64f::
{ 
Send("7-64, ")
Send("^i")
Sleep(100)
Send("Aluminum Industries")
Sleep(100)
Send("^i")
}
return

:*:7-72f::
{ 
Send("7-72, ")
Send("^i")
Sleep(100)
Send("Reformer and Cracking Furnace")
Sleep(100)
Send("^i")
}
return

:*:7-73f::
{ 
Send("7-73, ")
Send("^i")
Sleep(100)
Send("Dust Collectors and Collection Systems")
Sleep(100)
Send("^i")
}
return

:*:7-74f::
{ 
Send("7-74, ")
Send("^i")
Sleep(100)
Send("Distilleries")
Sleep(100)
Send("^i")
}
return

:*:7-75f::
{ 
Send("7-75, ")
Send("^i")
Sleep(100)
Send("Grain Storage and Milling")
Sleep(100)
Send("^i")
}
return

:*:7-76f::
{ 
Send("7-76, ")
Send("^i")
Sleep(100)
Send("Prevention and Mitigation of Combustible Dust Explosion and Fire")
Sleep(100)
Send("^i")
}
return

:*:7-77f::
{ 
Send("7-77, ")
Send("^i")
Sleep(100)
Send("Testing Internal Combustion Engines and Accessories")
Sleep(100)
Send("^i")
}
return

:*:7-78f::
{ 
Send("7-78, ")
Send("^i")
Sleep(100)
Send("Industrial Exhaust Systems")
Sleep(100)
Send("^i")
}
return

:*:7-79f::
{ 
Send("7-79, ")
Send("^i")
Sleep(100)
Send("Fire Protection for Gas Turbine and Electric Generators")
Sleep(100)
Send("^i")
}
return

:*:7-80f::
{ 
Send("7-80, ")
Send("^i")
Sleep(100)
Send("Organic Peroxides")
Sleep(100)
Send("^i")
}
return

:*:7-83f::
{ 
Send("7-83, ")
Send("^i")
Sleep(100)
Send("Drainage Systems for Ignitable Liquids")
Sleep(100)
Send("^i")
}
return

:*:7-85f::
{ 
Send("7-85, ")
Send("^i")
Sleep(100)
Send("Combustible and Reactive Metals")
Sleep(100)
Send("^i")
}
return

:*:7-86f::
{ 
Send("7-86, ")
Send("^i")
Sleep(100)
Send("Cellulose Nitrate")
Sleep(100)
Send("^i")
}
return

:*:7-88f::
{ 
Send("7-88, ")
Send("^i")
Sleep(100)
Send("Outdoor Ignitable Storage Tanks")
Sleep(100)
Send("^i")
}
return

:*:7-89f::
{ 
Send("7-89, ")
Send("^i")
Sleep(100)
Send("Ammonium Nitrate and Mixed Fertilizers Containing Ammonium Nitrate")
Sleep(100)
Send("^i")
}
return

:*:7-91f::
{ 
Send("7-91, ")
Send("^i")
Sleep(100)
Send("Hydrogen")
Sleep(100)
Send("^i")
}
return

:*:7-92f::
{ 
Send("7-92, ")
Send("^i")
Sleep(100)
Send("Ethylene Oxide")
Sleep(100)
Send("^i")
}
return

:*:7-93f::
{ 
Send("7-93, ")
Send("^i")
Sleep(100)
Send("Aircraft Hangars, Aircraft Manufacturing and Assembly Facilities")
Sleep(100)
Send("^i")
}
return

:*:7-95f::
{ 
Send("7-95, ")
Send("^i")
Sleep(100)
Send("Compressors")
Sleep(100)
Send("^i")
}
return

:*:7-96f::
{ 
Send("7-96, ")
Send("^i")
Sleep(100)
Send("Printing Plants")
Sleep(100)
Send("^i")
}
return

:*:7-97f::
{ 
Send("7-97, ")
Send("^i")
Sleep(100)
Send("Metal Cleaning")
Sleep(100)
Send("^i")
}
return

:*:7-98f::
{ 
Send("7-98, ")
Send("^i")
Sleep(100)
Send("Hydraulic Fluids")
Sleep(100)
Send("^i")
}
return

:*:7-99f::
{ 
Send("7-99, ")
Send("^i")
Sleep(100)
Send("Heat Transfer Fluid Systems")
Sleep(100)
Send("^i")
}
return

:*:7-101f::
{ 
Send("7-101, ")
Send("^i")
Sleep(100)
Send("Fire Protection for Steam Turbines and Electric Generators")
Sleep(100)
Send("^i")
}
return

:*:7-103f::
{ 
Send("7-103, ")
Send("^i")
Sleep(100)
Send("Turpentine Recovery in Pulp and Paper Mills")
Sleep(100)
Send("^i")
}
return

:*:7-104f::
{ 
Send("7-104, ")
Send("^i")
Sleep(100)
Send("Metal Treatment Processes for Steel Mills")
Sleep(100)
Send("^i")
}
return

:*:7-105f::
{ 
Send("7-105, ")
Send("^i")
Sleep(100)
Send("Concentrating Solar Power")
Sleep(100)
Send("^i")
}
return

:*:7-106f::
{ 
Send("7-106, ")
Send("^i")
Sleep(100)
Send("Ground Mounted Photovoltaic Solar Power")
Sleep(100)
Send("^i")
}
return

:*:7-107f::
{ 
Send("7-107, ")
Send("^i")
Sleep(100)
Send("Natural Gas Transmission and Storage")
Sleep(100)
Send("^i")
}
return

:*:7-108f::
{ 
Send("7-108, ")
Send("^i")
Sleep(100)
Send("Silane")
Sleep(100)
Send("^i")
}
return

:*:7-109f::
{ 
Send("7-109, ")
Send("^i")
Sleep(100)
Send("Fuel Fired Thermal Electric Power Generation Facilities")
Sleep(100)
Send("^i")
}
return

:*:7-110f::
{ 
Send("7-110, ")
Send("^i")
Sleep(100)
Send("Industrial Control Systems")
Sleep(100)
Send("^i")
}
return

:*:7-111f::
{ 
Send("7-111, ")
Send("^i")
Sleep(100)
Send("Chemical Process Industries")
Sleep(100)
Send("^i")
}
return

:*:7-111Af::
{ 
Send("7-111A, ")
Send("^i")
Sleep(100)
Send("Fuel-Grade Ethanol")
Sleep(100)
Send("^i")
}
return

:*:7-111Bf::
{ 
Send("7-111B, ")
Send("^i")
Sleep(100)
Send("Carbon Black")
Sleep(100)
Send("^i")
}
return

:*:7-111Cf::
{ 
Send("7-111C, ")
Send("^i")
Sleep(100)
Send("Titanium Dioxide")
Sleep(100)
Send("^i")
}
return

:*:7-111Df::
{ 
Send("7-111D, ")
Send("^i")
Sleep(100)
Send("Oilseed Processing")
Sleep(100)
Send("^i")
}
return

:*:7-111Ef::
{ 
Send("7-111E, ")
Send("^i")
Sleep(100)
Send("Chloro-Alkali")
Sleep(100)
Send("^i")
}
return

:*:7-111Ff::
{ 
Send("7-111F, ")
Send("^i")
Sleep(100)
Send("Sulfuric Acid")
Sleep(100)
Send("^i")
}
return

:*:7-111Gf::
{ 
Send("7-111G, ")
Send("^i")
Sleep(100)
Send("Ammonia and Ammonia Derivatives")
Sleep(100)
Send("^i")
}
return

:*:7-111Hf::
{ 
Send("7-111H, ")
Send("^i")
Sleep(100)
Send("Olefins")
Sleep(100)
Send("^i")
}
return

:*:7-111If::
{ 
Send("7-111I, ")
Send("^i")
Sleep(100)
Send("Ink, Paint and Coating Formulations")
Sleep(100)
Send("^i")
}
return

:*:8-1f::
{ 
Send("8-1, ")
Send("^i")
Sleep(100)
Send("Commodity Classification")
Sleep(100)
Send("^i")
}
return

:*:8-3f::
{ 
Send("8-3, ")
Send("^i")
Sleep(100)
Send("Rubber Tire Storage")
Sleep(100)
Send("^i")
}
return

:*:8-7f::
{ 
Send("8-7, ")
Send("^i")
Sleep(100)
Send("Baled Fiber Storage")
Sleep(100)
Send("^i")
}
return

:*:8-9f::
{ 
Send("8-9, ")
Send("^i")
Sleep(100)
Send("Storage of Class 1, 2, 3, 4 and Plastic Commodities")
Sleep(100)
Send("^i")
}
return

:*:8-10f::
{ 
Send("8-10, ")
Send("^i")
Sleep(100)
Send("Coal and Charcoal Storage")
Sleep(100)
Send("^i")
}
return

:*:8-18f::
{ 
Send("8-18, ")
Send("^i")
Sleep(100)
Send("Storage of Hanging Garments")
Sleep(100)
Send("^i")
}
return

:*:8-21f::
{ 
Send("8-21, ")
Send("^i")
Sleep(100)
Send("Roll Paper Storage")
Sleep(100)
Send("^i")
}
return

:*:8-22f::
{ 
Send("8-22, ")
Send("^i")
Sleep(100)
Send("Storage of Baled Waste Paper")
Sleep(100)
Send("^i")
}
return

:*:8-23f::
{ 
Send("8-23, ")
Send("^i")
Sleep(100)
Send("Rolled Nonwoven Fabric Storage")
Sleep(100)
Send("^i")
}
return

:*:8-24f::
{ 
Send("8-24, ")
Send("^i")
Sleep(100)
Send("Idle Pallet Storage")
Sleep(100)
Send("^i")
}
return

:*:8-27f::
{ 
Send("8-27, ")
Send("^i")
Sleep(100)
Send("Storage of Wood Chips")
Sleep(100)
Send("^i")
}
return

:*:8-28f::
{ 
Send("8-28, ")
Send("^i")
Sleep(100)
Send("Pulpwood and Outdoor Log Storage")
Sleep(100)
Send("^i")
}
return

:*:8-30f::
{ 
Send("8-30, ")
Send("^i")
Sleep(100)
Send("Storage of Carpets")
Sleep(100)
Send("^i")
}
return

:*:8-33f::
{ 
Send("8-33, ")
Send("^i")
Sleep(100)
Send("Carousel Storage and Retrieval Systems")
Sleep(100)
Send("^i")
}
return

:*:8-34f::
{ 
Send("8-34, ")
Send("^i")
Sleep(100)
Send("Protection for Automatic Storage and Retrieval Systems")
Sleep(100)
Send("^i")
}
return

:*:9-0f::
{ 
Send("9-0, ")
Send("^i")
Sleep(100)
Send("Asset Integrity")
Sleep(100)
Send("^i")
}
return

:*:9-1f::
{ 
Send("9-1, ")
Send("^i")
Sleep(100)
Send("Supervision of Property")
Sleep(100)
Send("^i")
}
return

:*:9-16f::
{ 
Send("9-16, ")
Send("^i")
Sleep(100)
Send("Burglary and Theft")
Sleep(100)
Send("^i")
}
return

:*:9-18f::
{ 
Send("9-18, ")
Send("^i")
Sleep(100)
Send("Prevention of Freeze-ups")
Sleep(100)
Send("^i")
}
return

:*:9-19f::
{ 
Send("9-19, ")
Send("^i")
Sleep(100)
Send("Wildland Fire")
Sleep(100)
Send("^i")
}
return

:*:10-0f::
{ 
Send("10-0, ")
Send("^i")
Sleep(100)
Send("The Human Factor of Property Conservation")
Sleep(100)
Send("^i")
}
return

:*:10-1f::
{ 
Send("10-1, ")
Send("^i")
Sleep(100)
Send("Pre-Incident and Emergency Response Planning")
Sleep(100)
Send("^i")
}
return

:*:10-3f::
{ 
Send("10-3, ")
Send("^i")
Sleep(100)
Send("Hot Work Management")
Sleep(100)
Send("^i")
}
return

:*:10-4f::
{ 
Send("10-4, ")
Send("^i")
Sleep(100)
Send("Contractor Management")
Sleep(100)
Send("^i")
}
return

:*:10-5f::
{ 
Send("10-5, ")
Send("^i")
Sleep(100)
Send("Disaster Recovery and Contingency Plan")
Sleep(100)
Send("^i")
}
return

:*:10-6f::
{ 
Send("10-6, ")
Send("^i")
Sleep(100)
Send("Protection Against Arson and Other Incendiary Fires")
Sleep(100)
Send("^i")
}
return

:*:10-7f::
{ 
Send("10-7, ")
Send("^i")
Sleep(100)
Send("Fire Protection Impairment Management")
Sleep(100)
Send("^i")
}
return

:*:10-8f::
{ 
Send("10-8, ")
Send("^i")
Sleep(100)
Send("Operators")
Sleep(100)
Send("^i")
}
return

:*:12-2f::
{ 
Send("12-2, ")
Send("^i")
Sleep(100)
Send("Vessels & Piping")
Sleep(100)
Send("^i")
}
return

:*:12-3f::
{ 
Send("12-3, ")
Send("^i")
Sleep(100)
Send("Continuous Digesters & Related Vessels")
Sleep(100)
Send("^i")
}
return

:*:12-6f::
{ 
Send("12-6, ")
Send("^i")
Sleep(100)
Send("Batch Digesters & Related Vessels")
Sleep(100)
Send("^i")
}
return

:*:12-43f::
{ 
Send("12-43, ")
Send("^i")
Sleep(100)
Send("Pressure Relief Devices")
Sleep(100)
Send("^i")
}
return

:*:12-53f::
{ 
Send("12-53, ")
Send("^i")
Sleep(100)
Send("Absorption Refrigeration Systems")
Sleep(100)
Send("^i")
}
return

:*:13-1f::
{ 
Send("13-1, ")
Send("^i")
Sleep(100)
Send("Cold Mechanical Repairs")
Sleep(100)
Send("^i")
}
return

:*:13-2f::
{ 
Send("13-2, ")
Send("^i")
Sleep(100)
Send("Hydroelectric Power Plants")
Sleep(100)
Send("^i")
}
return

:*:13-3f::
{ 
Send("13-3, ")
Send("^i")
Sleep(100)
Send("Steam Turbines")
Sleep(100)
Send("^i")
}
return

:*:13-6f::
{ 
Send("13-6, ")
Send("^i")
Sleep(100)
Send("Flywheels and Pulleys")
Sleep(100)
Send("^i")
}
return

:*:13-7f::
{ 
Send("13-7, ")
Send("^i")
Sleep(100)
Send("Gears")
Sleep(100)
Send("^i")
}
return

:*:13-8f::
{ 
Send("13-8, ")
Send("^i")
Sleep(100)
Send("Power Presses")
Sleep(100)
Send("^i")
}
return

:*:13-10f::
{ 
Send("13-10, ")
Send("^i")
Sleep(100)
Send("Wind Turbines and Farms")
Sleep(100)
Send("^i")
}
return

:*:13-17f::
{ 
Send("13-17, ")
Send("^i")
Sleep(100)
Send("Gas Turbines")
Sleep(100)
Send("^i")
}
return

:*:13-18f::
{ 
Send("13-18, ")
Send("^i")
Sleep(100)
Send("Industrial Clutches and Clutch Couplings")
Sleep(100)
Send("^i")
}
return

:*:13-24f::
{ 
Send("13-24, ")
Send("^i")
Sleep(100)
Send("Fans and Blowers")
Sleep(100)
Send("^i")
}
return

:*:13-26f::
{ 
Send("13-26, ")
Send("^i")
Sleep(100)
Send("Internal Combustion Engines")
Sleep(100)
Send("^i")
}
return

:*:13-27f::
{ 
Send("13-27, ")
Send("^i")
Sleep(100)
Send("Heavy Duty Mobile Equipment")
Sleep(100)
Send("^i")
}
return

:*:13-28f::
{ 
Send("13-28, ")
Send("^i")
Sleep(100)
Send("Aluminum Industries")
Sleep(100)
Send("^i")
}
return

:*:17-0f::
{ 
Send("17-0, ")
Send("^i")
Sleep(100)
Send("Asset Integrity")
Sleep(100)
Send("^i")
}
return

:*:17-2f::
{ 
Send("17-2, ")
Send("^i")
Sleep(100)
Send("Process Safety")
Sleep(100)
Send("^i")
}
return

:*:17-4f::
{ 
Send("17-4, ")
Send("^i")
Sleep(100)
Send("Monitoring and Diagnosis of Vibration in Rotating Machinery")
Sleep(100)
Send("^i")
}
return

:*:17-11f::
{ 
Send("17-11, ")
Send("^i")
Sleep(100)
Send("Chemical Reactors and Reactions")
Sleep(100)
Send("^i")
}
return

:*:17-12f::
{ 
Send("17-12, ")
Send("^i")
Sleep(100)
Send("Semiconductor Fabrication Facilities")
Sleep(100)
Send("^i")
}
return

:*:17-16f::
{ 
Send("17-16, ")
Send("^i")
Sleep(100)
Send("Cranes")
Sleep(100)
Send("^i")
}
return