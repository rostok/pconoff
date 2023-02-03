# pconoff
PCONOFF analyzes Windows events and shows when computer was turned on and off. 
Program mimics behaviour of PcOnOffTime propertiary app but is rather simple as I wanted to see what ChatGPT can produce.

# instalation
Get the repo and compile it running `csc pconoff.cs`

# options
```
syntax: pconoff [days-back] [-t] [-h]
 days    how deep in time should the events be analyzed (90 by default)
 -c      will print report to console
 -t      will generate periods.js and table.html & will NOT show gird
```