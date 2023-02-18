'''
This file will construct and run the pyomo abstract model on the data sub_folder (Test1)
The steps are as follow:
1. Import libraries
2. Create solver
3. Build the model
4. Declare the Sets, Parameters and Variables
5. Declare the constraints and objective function
6. Load related data for sets and parameters
7. Create an instance of the model
8. Run the solver on the instance
9. Get the data folder and sub-folder from 'Setting.txt' file
10. Check and change directory, create folder for storing the result
11. Get the value from Variables and Objective function
12. Create a Dataframe for results
13. Write the result in xlsx format
14. Save the Excel file in desired path

Note: This code script is imported to the main script and will be run by calling main.py
Note: This file only runs on test1 data and save the result in this path(data/test1/Results)

'''
# load libraries
import sys
import os
from pyomo.core import *
from pyomo.environ import *
import numpy as np
import pandas as pd
from pyomo.opt import SolverFactory
import json
import xlsxwriter
from copy import deepcopy


# Create a solver
opt = SolverFactory('gurobi')

model = AbstractModel(name="(GGM_Test1)")

#### Sets
model.n = Set(ordered=True) # the set of all nodes
model.a = Set(ordered=True) # the set of all arcs, including LNG
model.y = Set(ordered=True) # set of year
model.r = Set(ordered=True) # set of resources


# subsets
#model.aprl = Set(ordered=True) # Set of Arcs except for LNG
model.n_p = Set(ordered=True) # the production nodes
model.n_c = Set(ordered=True) # the consumption nodes
model.a_s = Set(ordered=True) # set of start node of arcs
model.a_e = Set(ordered=True) # set of end node of arcs
# We do not consider Countries (CN) as an index for variables

### Parameters

#int(n_c,y): Intercept of Demand Curve, indexed with consumption node and year
model.int = Param(model.n_c,model.y,mutable=True,initialize = 0)

#slp(n_c,y): slope of demand curve, indexed with consumption node and year
model.slp = Param(model.n_c,model.y,mutable=True,initialize = 0)

#model.cour, indexed with country from (data excel file, M tab, market power table), consumption node and year
model.cour = Param(model.n_p,model.n_c,model.y,mutable=True,initialize = 0)

#model.cost_pl, indexed with supplier (prod_node), resource and year
model.cost_pl = Param(model.n_p,model.r,model.y,mutable=True,initialize = 0)

#model.cost_pq, indexed with supplier (prod_node), resource and year
model.cost_pq = Param(model.n_p,model.r,model.y,mutable=True,initialize = 0)

#model.cap_p, indexed with supplier (prod_node), resource and year
model.cap_p = Param(model.n_p,model.r,model.y,mutable=True,initialize = 0)

#model.cap_a, indexed with arcs_main and year. It does not include LNG.
model.cap_a = Param(model.a,model.y,mutable=True,initialize = 0)

#model.cost_a, indexed with arc and year
model.cost_a = Param(model.a,model.y,mutable=True,initialize = 0)

#model.d_max_a, indexed with arcs_main and year.It does not include LNG.
model.d_a_max1 = Param(model.a,model.y,mutable=True,initialize = 0) #bcm

#model.inv_a, indexed with arcs_main and year.It does not include LNG.
model.inv_a = Param(model.a,model.y,mutable=True,initialize = 0)

#model.l_a, indexed with arc and year
model.l_a = Param(model.a,model.y,mutable=True,initialize = 0)

#model.disc, indexed with year
model.disc = Param(model.y,mutable=True,initialize = 0)

### Variables
#model.Q_P is the quantity produced by Resource, then it is indexed over (r, n (as prod_node), y). The supplier is an auxiliary
model.Q_P = Var(model.n_p, model.r,model.y, within=NonNegativeReals,initialize = 0)

#model.Q_S, is the quantity sold, indexed over (n (as cons_node),y). Which it appears to be the quantity consumed by consumption nodes
model.Q_S = Var(model.n_c,model.y, within=NonNegativeReals,initialize = 0) #how much is consumed

#model.F_A. indexed with (supplier (as country with prod_node), a, y)
model.F_A = Var(model.n_p,model.a,model.y, within=NonNegativeReals,initialize = 0)

#model.D_A, indexed with (a,y). It is indexed with arcs_main which includes all arcs except LNG because vessels has no capacity nor expansion for capacity
model.D_A = Var(model.a,model.y, within=NonNegativeReals,initialize = 0)


# NOTE: the indices for Cap_p are different in Word document and GAMs. I followed the GAMS to avoid errors
def prod_cap_lim_rule(model,n_p,r,y):
    return(
    model.Q_P[n_p, r, y] <= model.cap_p[n_p,r,y]
    )
model.prod_cap_lim = Constraint(model.n_p,model.r,model.y,rule=prod_cap_lim_rule)

# Arc flow capacity  fa <= cap_a + da
def arc_flow_cap_rule(model,a,y):
    return sum(model.F_A[n_p,a,y] for n_p in model.n_p) <= model.cap_a[a,y] + sum(model.D_A[a,y] for y in model.y if y<y+1)
model.arc_flow_cap = Constraint(model.a, model.y, rule=arc_flow_cap_rule)

# mass balance  qp + (1-la)fa == qs + fa
def mass_balance_rule(model,n_p,n_c,a,y):
    return sum(model.Q_P[n_p,r,y] for r in model.r) + sum(model.F_A[n_p,a[0],y] * (1-model.l_a[a[0],y]) for a in model.a_e if a[1]==n_c) == model.Q_S[n_c,y] + sum(model.F_A[n_p,a[0],y] for a in model.a_s if a[1]==n_p)


#return sum(model.Q_P[n_p,r,y] for r in model.r) if n_p in model.n_p else 0 + sum(model.F_A[a[0],y] * (1-model.l_a[a[0]]) for a in model.a_e if a[1]==n) == model.Q_S[n_c,y] if n_c in model.n_c else 0 + sum(model.F_A[a[0],y] for a in model.a_s if a[1]==n)
model.mass_balance = Constraint(model.n_p,model.n_c,model.a ,model.y, rule=mass_balance_rule)

# expansion da <= d_max_a
def arc_expan_rule(model,a,y):
    return model.D_A[a,y] <= model.d_a_max1[a,y]
model.arc_expan = Constraint(model.a,model.y, rule=arc_expan_rule)

### Objective function
# obj = REV + CS - TC - MPA
def obj_rule(model):
    #return sum((model.cost_pl[n_p, r, y] + 0.5 * model.cost_pq[n_p, r, y] * model.qp[n_p, r, y] * model.qp[n_p, r, y]) * model.disc[y] for n_p[1] in model.n_p if n_p[0]="NOR"  for r in model.r  for y in model.y) \
    return sum((model.cost_pl[n_p, r, y] + 0.5 * model.cost_pq[n_p, r, y] * model.Q_P[n_p, r, y]) * model.Q_P[n_p, r, y] * model.disc[y] for n_p in model.n_p for r in model.r  for y in model.y) \
           + sum(model.cost_a[a, y] * model.F_A[n_p,a, y] * model.disc[y] for a in model.a for n_p in model.n_p for y in model.y) \
           + sum(model.inv_a[a, y] * model.D_A[a, y] * model.disc[y] for y in model.y for a in model.a) \
           + sum(model.cour[n_p, n_c, y] * model.slp[n_c, y] *  (model.Q_S[n_c, y] ** 2) * model.disc[y] for y in model.y for n_c in model.n_c for n_p in model.n_p) / 2 \
           - sum((model.int[n_c, y] - (model.slp[n_c, y] * model.Q_S[n_c, y])) * model.Q_S[n_c, y]  * model.disc[y] for y in model.y for n_c in model.n_c) \
           - sum(model.slp[n_c, y] * model.Q_S[n_c, y] ** 2 * model.disc[y] for y in model.y for n_c in model.n_c)  / 2


model.obj = Objective(sense=minimize, rule=obj_rule)

######################## Load data values
# Load Sets value
data = DataPortal()
data.load(filename = 'data/test1/ggm/Nodes.csv', format = 'set', set=model.n)
#data.load(filename = 'data/test1/ggm/Arcs_prl.csv', format = 'set', set=model.aprl)
data.load(filename = 'data/test1/ggm/Arcs.csv', format = 'set', set=model.a)
data.load(filename = 'data/test1/ggm/Years.csv', format = 'set', set=model.y)
data.load(filename = 'data/test1/ggm/Resources.csv', format = 'set', set=model.r)
data.load(filename = 'data/test1/ggm/np.csv', format = 'set', set=model.n_p)
data.load(filename = 'data/test1/ggm/nc.csv', format = 'set', set=model.n_c)
data.load(filename = 'data/test1/ggm/a_s.csv', format = 'set', set=model.a_s)
data.load(filename = 'data/test1/ggm/a_e.csv', format = 'set', set=model.a_e)

data.load(filename = 'data/test1/ggm/int.csv', index=[model.n_c,model.y], param=model.int)
data.load(filename = 'data/test1/ggm/slp.csv', index=[model.n_c,model.y], param=model.slp)
data.load(filename = 'data/test1/ggm/cour.csv', index=[model.n_p,model.n_c,model.y], param=model.cour)
data.load(filename = 'data/test1/ggm/cost_pl.csv', index=[model.n_p,model.r,model.y], param=model.cost_pl)
data.load(filename = 'data/test1/ggm/cost_pq.csv', index=[model.n_p,model.r,model.y], param=model.cost_pq)
data.load(filename = 'data/test1/ggm/cap_p.csv', index=[model.n_p,model.r,model.y], param=model.cap_p)
data.load(filename = 'data/test1/ggm/cap_a.csv', index=[model.a,model.y], param=model.cap_a)
data.load(filename = 'data/test1/ggm/cost_a.csv', index=[model.a,model.y], param=model.cost_a)
data.load(filename = 'data/test1/ggm/d_a_max1.csv', index=[model.a,model.y], param=model.d_a_max1)
data.load(filename = 'data/test1/ggm/inv_a.csv', index=[model.a,model.y], param=model.inv_a)
data.load(filename = 'data/test1/ggm/l_a.csv', index=[model.a,model.y], param=model.l_a)
data.load(filename = 'data/test1/ggm/disc.csv', param=model.disc)
#################### Solve the model

# Create a model instance and optimize
instance = model.create_instance(data)
results = SolverFactory('gurobi', Verbose=True).solve(instance, tee=True)
instance.solutions.load_from(results)
################# Save the result of the model in the Excel file ####################3

# Variables
Q_P_data = {(n_p, r,y): value(v) for (n_p, r,y), v in instance.Q_P.items()}
df_Q_P = pd.DataFrame.from_dict(Q_P_data, orient="index", columns=["variable value"])
Q_S_data = {(n_c, y): value(v) for (n_c,y), v in instance.Q_S.items()}
df_Q_S = pd.DataFrame.from_dict(Q_S_data, orient="index", columns=["variable value"])
F_A_data = {(n_p,a,y): value(v) for (n_p,a,y) , v in instance.F_A.items()}
df_F_A = pd.DataFrame.from_dict(F_A_data,orient='index', columns=['variable value'])
D_A_data = {(a,y): value(v) for (a,y), v in instance.D_A.items()}
df_D_A = pd.DataFrame.from_dict(D_A_data, orient='index', columns=['variable value'])

header0 = []
header1 = []

for o in instance.component_data_objects(Objective, active=True):
        header0.append("Objective")
        header1.append(str(o.name))
MultiHeaders = [header0, header1]

AllData = []
pov_data = []

for o in instance.component_data_objects(Objective, active=True):
        pov_data.append(value(o))

AllData.append(deepcopy(pov_data))
pov_data.clear()


obj_results = pd.DataFrame(data = np.array(AllData), columns =MultiHeaders)
print(obj_results)

############## ُSave the result as excel file #####################
# store the 'Setting.txt' file as a variable
file_setting = 'Setting.txt'
param_list = {}
try:
    cur_path = os.getcwd()  # get the current path
    with open(file_setting) as f:  # open text file as 'f'
        for line in f:  # read each line in file
            splitLine = line.split(':')  # in each line, parameter's name and value are separated by (:)
            param_list[splitLine[0]] = ",".join(splitLine[1:]).strip(
                '\n')  # Get the first term as key and the second as value, then it stores the value in a variable and remove the extra character (\n)
    for key, val in param_list.items():  # get the parameters' value and save in variables for further reference
        if key == 'Main':  # The main data folder
            main_data = val
        if key == 'Data':  # The dataset sub_folder
            test_path = val
except FileNotFoundError:
    print("Directory: {0} does not exist".format(file_setting))  # Print this message if file is nto found
####################### Read the excel files from data test folder directory #########################
try:
    data_file = main_data  # get the main folder of all data ('data'), it reads from setting text file
    os.chdir(data_file)  # change directory the 'data' main folder
    main_path = os.getcwd()  # get directory after changed to 'data' folder
    test_dir = test_path  # get the sub_folder directory (test folders), reads from setting text file
    join_dir = os.path.join(main_path, test_dir)  # join test sub_folder to the changed directory
    res_dir = os.path.join(join_dir,'Results')  # create new sub_folder to store defined_ranges as dataframe in 'CSV' format
    if not os.path.exists(res_dir):  # If the name of sub_folder is not in desired directory, create it
        os.makedirs(os.path.join(join_dir, res_dir))

    # Create a Pandas Excel w
    writer = pd.ExcelWriter(f'{res_dir}/Test1Model.xlsx', engine='xlsxwriter')  # header=0)
    # Write each dataframe to a different worksheet.
    df_Q_S.to_excel(writer, sheet_name='Quantity Sold')
    df_Q_P.to_excel(writer, sheet_name='Quantity Produced')
    df_F_A.to_excel(writer, sheet_name='Flow of Arc')
    df_D_A.to_excel(writer, sheet_name='Expansion Arc')
    obj_results.to_excel(writer, sheet_name='Objective Result')

    # Close the Pandas Excel writer and output the Excel file.
    writer.save()
except FileNotFoundError:
    print("Directory: {0} does not exist".format(main_path))  # Print this message if the file or directory could not be be found
