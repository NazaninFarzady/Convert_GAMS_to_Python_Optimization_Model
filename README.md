# Convert_GAMS_to_Python_Optimization_Model

Introdutcion
This repository contains the steps of translating the deterministic Global Gas Model (GGM) from GAMS to Python. The comments in the code script will determine how each line of code works. Some 'Notes' indicate the differences between GAMS and Python regarding preparing data and modeling GGM. There are two main code scripts on this project. The first one reads data from excel files, cleans, and structures for the Pyomo modeling package. The second file constructs the Pyomo modeling, which includes the construction of the GGM model with its components. 

Data

Since some data are proprietary it is not possible to make available all input data that has been used. The deterministic GGM reads input from three MS Excel workbooks3, which is processed further in the model. 
The relevant data for the GGM are collected, categorized, and organized in three MS Excel workbooks. The primary documentation  of the model and data provides a comprehensive explanation of data structure and data processing. 
The essential part of the data reading process is the name manager in MS Excel files which shows defined ranges of data used by GAMS when reading the data. Similarly, the python code script reads these defined ranges to get data and compute the parameters' value

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
- data: Containes different worksheet including: Parameters values, Nodes (Geographical,liquefaction and regasification) , Terminals (pipelines, liquefaction and regasification), Distance matrix for shipping of LNG, Gas storages, Market power, Production resources, Years used in the data sets, Sectors residential and commercial sector (building heating), industry, electric power generation, transport, Arcs (Piplines, liquefaction, Shipping, regasification), LNG shipping distances matrix, Storage, and Market power.
- data projection: Production and consumption data (based on different scenarios)
- data calibration: Calibration for each scenario (Calibration is the process to reconcile model outputs with reference values by means of input data
adjustments.)

Note: In MS Excel, the name manager shows definitions that are used by GAMS when reading the Workbook. It has been used in code to read relevant range of data and as the name of output file for storing the defined range.

Note: Data from three MS Excel has been used to calculate parameters as model inputs, and defining model sets.


GGM

A multi-period model for analyzing the world natural gas market.
Country level; with large countries disaggregated (USA, CAN, RUS, CHN, IND).
Focus on infrastructure investment and trade, taking into account market power.
Production, pipelines, liquefaction, regasification, shipping, storage.
Implemented in GAMS.
Input data files are MS Excel workbooks.


Model input:
Reference values for production, consumption, prices, market power for base year and future years’ projections
Capacities, investment and operational costs, depreciation and loss rates of production, transportation, and storage infrastructure.
Demand seasonality, sector shares and elasticities, production costs.


Model output:
Pipeline, liquefaction, regasification, storage expansions and utilization.
Seasonal production, consumption, trade and prices.
Sector profits, costs, consumer surplus and social welfare impacts.


Model execution:
- Gathering, cleaning, processing data from three MS Excel files.
- Defining Sets, Parameters, and their values in proper format to enter the Pyomo model
- Build Pyomo model, enter input and run the model


The code script will read the data from an external source by following steps:
1.	Read the value of the global parameter from 'Setting.txt.'
2.	Set the working directory to the data folder
3.	Read the excel files and store defined_range values as a data frame in 'CSV format
4.	Save the first batch of 'CSV files in sub_folder (Step 3 and 4 is necessary to construct data frame with the specific name of defined_ranges for the further process)
5.	Read the 'CSV files and store them as new data frames with the name of defined_range individually
6.	Clean each data frame by considering the structure
    a.	Set the header by the column names. In some data frames, it is required to combine two rows and set header, then
    b.	Convert data type, integer, string, and float
    c.	Replace the value of 'EPS' to 0
    d.	Replace (0) values with specific values for some parameters (as determined in GAMS and data document)
7.	Merge data frames where it is needed to compute parameter values easier
8.	Compute parameter values and add a new column with specific labels to a related data frame
9.	Build separate data frames for each parameter value along with its indices
10.	Create a list of parameters names from GAMS and data document. (This list will be used to name data frames and save in 'CSV' format for Pyomo package modeling)
11.	Store data frame for parameters and indexes with the name from the list of parameters and save as 'CSV' format in new sub_folder


Constructing GGM using Pyomo package
By installing the Pyomo package, we can construct the GGM model with its components, as Sets, parameters, variables, constraints, and objective functions. To make modeling more flexible, we consider an abstract model which will read values for sets and parameters from the external data source; here are prepared 'CSV files from the previous code script.
After declaring the model structure, the code script will read the data values from the second sub_folder, 'CSV files, for parameters, indices, and sets values.


Running the model and solve
To solve GGM using Pyomo package of Python, we consider the 'Gurobi' academic license. To install Guorbi solver, you need to register on the website (https://www.gurobi.com/), download, and install the license. By following the documentation on the official website of Gurobi, you can install and run it on your machine. 
After installing the Guorbi solver, the code will run the model and solver, then return the result with a model summary.


Extracting the result of the model
Finally, the code script will save the value of variables and objective function in a separate excel file, in the root of the project folder, to present the result of the model.
NOTE: there are comments in the code script to explain codes in detail for later implementation.
