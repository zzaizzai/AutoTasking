# auto tasking system with data in complex files

I use it in my work. becuz I handle with many data values which I have to evaluate for new or current materials as my job.   
Without this code, I have to spend much time finding data files in share folder and considering which data should be nessasery.   
however, with this code, I can **save about 90% of my time** handling with data in excels. Also I dont have to confirm that certain data were completed and exist on share folder. It restains excessive effort and make me concentrate on considering the mehcanism of phenomenon.   


## Process
Collect Files having name of target from share folder to my desktop so that i can see data files more quickly and easily     
-> Read Data Files with pandas       
-> normalize Data values from excel files with pandas     
-> Make Data Sheet we can use it as a report data      

```
python 3.7.8
```

```
pip install -r requirements.txt
python Main.py
```

### .env
```
HOSE_DIR = "データが入っている場所"
AUTO_TENSION_ALL_XLSX_1 = "AutoTension用のエクセルファイル"
AUTO_TENSION_ALL_XLSX_2 = "AutoTension用のエクセルファイル"
```

**There are some confidential information but I manage that in a environment file**