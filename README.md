# Convert_GAMS_to_Python_Optimization_Model

Introdutcion
This repository contains the steps of translating the deterministic Global Gas Model (GGM) from GAMS to Python. The comments in the code script will determine how each line of code works. Some 'Notes' indicate the differences between GAMS and Python regarding preparing data and modeling GGM. There are two main code scripts on this project. The first one reads data from excel files, cleans, and structures for the Pyomo modeling package. The second file constructs the Pyomo modeling, which includes the construction of the GGM model with its components. 

Data
Since some data are proprietary it is not possible to make available all input data that has been used. The deterministic GGM reads input from three MS Excel workbooks3, which is processed further in the model. 
There are three types of actors along the natural gas value chain that are represented in the Global Gas Model. For each actor, model uses specific input data that is documented as following.

Gas market data categories:
- Production
- Consumption
- Piplines
- Liquefaction terminals
- Regasification terminals
- LNG ship
- Storages (is not included in this project)


MS Excel workbooks with input data:
- data
- data projection
- data calibration


GGM
