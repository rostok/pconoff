# pconoff
PCONOFF analyzes Windows events and shows when computer was turned on and unlocked. 
Program mimics behaviour of PcOnOffTime propertiary app. But this is rather simple software 
as I wanted to see what ChatGPT can produce.

# instalation
Get the repo and compile it running `csc pconoff.cs`

# options
```
PCONOFF will show when this computer was turned on and unlocked.
syntax: pconoff [days] [-l] [-k] [-c] [-t]
 days    how deep in time should the events be analyzed (90 by default)
 -l      will investigate log on/off events
 -k      will investigate unlocked/locked events
 -c      will print report to console
 -v      verbose output of events
 -t      will generate periods.js and table.html & will NOT show gird
```

# license
MIT