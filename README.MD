# IONApiDemo
## What is this
Demos of how to connect to Infor's ION API to call MI Programs. Duh.
One done in Powershell, one done in VBA to be used inside Excel.

## How can I use these demoes?
Do not clone one Onedrive based directory as VBA will act wonky and will not find the cred file. Yes, it's a demo.
### General
Clone repo and place your .ionapi credential file inside the folder as CredFile.ionapi (included in .gitignore).
### Powershell
Will execute MNS150MI/SelUsers
- Start IONApiDemo.ps1
### Excel
Will execute MNS150MI/SelUsers
- Open IONApiDemo.xlsm in Excel
- Open VBS IDE
- Execute TestMICall()

If you do not trust this repo (I would not) assemble your own:
- create xslm file in Excel
- add IONApiDemo.bas to the file
- add JsonConverter.bas either from here or from VBA-JSON (https://github.com/VBA-tools/VBA-JSON)
- enable reference "Microsoft Scripting Runtime"