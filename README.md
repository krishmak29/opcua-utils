# PLC / SCADA Object Verification Utility

This version uses the existing engineering Excel structure and groups all rows belonging to the same PLC + Area Code + Equipment Number into one object.

Features:
- PLC tabs generated from the PLC column
- Object cards with no animations
- Main value defaults to OPC Tag Suffix RD.PV
- Tested status: Not Tested / Correct / Incorrect / Recheck
- Color indication for verification status
- Next / Previous pages
- All parameters shown in popup
- Template name shown in popup
- Live OPC UA values
- Correct / Incorrect / Recheck
- Configurable cause list and tester/comment
- Persistent SQLite verification database
- Separate utility configuration Excel

Engineering Excel expected columns include:
PLC, Area Code*, Equipment Number*, Template, OPC Tag Suffix, GE PLC Tag Name, Tag Name, Address*, Data Type, Client Access, Eng Units, Description*, plus the remaining project columns.

Configuration workbook:
PLC_Config: PLC, Vendor, Protocol, IP, Port, EndpointPath, Username, Password
Display_Config: Parameter, Value
Cause_Master: Cause Code, Cause Description

Run:
py -m pip install -r requirements.txt
py main.py

EXE:
pyinstaller --noconfirm --onefile --windowed --name PLC_SCADA_Verification_Utility main.py

Important: Address* is currently passed directly as an OPC UA NodeId. If Address* is a vendor PLC address/item name rather than an OPC UA NodeId, the communication adapter must resolve it through the relevant OPC server/gateway.
