# Prime KPI Scripts

This set of scripts provide you with examples around how to get wireless KPI data from Cisco Prime Infrastructure

Contacts:

* Santiago Flores (sfloresk@cisco.com)

## Instruction

Download the code to your computer
```
git clone https://github.com/sfloresk/prime-kpi-scripts.git
```

Add to the primecredentials file your Prime Infrastructure credentials. For example

```
#Format: primeUrl,username,password
https://primeinfrastructure.cisco.com/,primeuser,supersecretpassword
```

Go to the project directory and install python requirements with pip:
```
pip install -r requirements.txt
```

You are ready to go!

## Available Scripts

### createReport.py
This script gets APs along with RF Stats, counters and load information, and creates an excel file.


#### To run this script:
```
python createReport.py
```
