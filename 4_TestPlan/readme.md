# TEST PLAN:

## Table 1: Unit Testing

| **Test ID** | **Description**                                              | **Exp I/P** | **Exp O/P** | **Actual Out** | 
|-------------|--------------------------------------------------------------|------------|-------------|----------------|
|  U_01     |Checking cgpa calculation|Score obtained|Score in cgpa|As expected|
|  U_02       |Analysis of every batch code| Batch code, Emp Name &PS, Total number of batches|min, max, avg, SD, Kurt, Top, Mid, Bottom blockers of batches. |Not exactly as expected|
|  U_03       |Analysis of every module|  Module Score , Hours, credits, Module name, |Min, max, avg, SD, Kurt of modules.Top, Mid, Bottom blockers of modules.|Not exactly as expected|
|U_04|Bar chart representation |Data of Top, Mid, Bottom blockers. |Visualisation of input data on graph|As expected|



## Table 2: Integration testing

| **Test ID** | **Description**                                              | **Exp I/P** | **Exp O/P** | **Actual Out** | 
|-------------|--------------------------------------------------------------|------------|-------------|----------------|
|I_01|Calculation of cgpa for each module|Projected points of modules|Finding score in cgpa|As expected|
|I_02|Batch wise Analysis by considering the scores of all the modules|Batch code, Emp name &PS,Total number of batches, score|Finding the batch wise min, max, avg, SD, Kurt score. 
Top, Mid, Bottom blockers of batches.
|Not exactly as expected|
|I_03|By using the given data, calculation for every modules - Module wise  Analysis |Number of modules, each module score and , credits of each module|Finding the module wise min, max, avg, SD, Kurt score.
Top, Mid, Bottom blockers of modules.
|Not exactly as expected|
|I_04|Module wise and batch wise analysis result will form the Bar chart.|Data of Top, Mid, Bottom blockers. |Visualisation of input data on graph|As expected|


## Table 3: System testing

| **Test ID** | **Description**                                              | **Exp I/P** | **Exp O/P** | **Actual Out** | 
|-------------|--------------------------------------------------------------|------------|-------------|----------------|
|S_01|Functionality and  Calculation of cgpa for each module|Projected points of modules|Output should be in the form of cgpa|As expected|
|S_02|Checking the functionality of a code in batch wise by considering the scores of all the modules.|Batch code, Emp name &PS, Total number of batches, score|Finding the batch wise min, max, avg, SD, Kurt score.
Top, Mid, Bottom blockers of batches
|Not exactly as expected|
|S_03 |Checking the functionality of a code for Module wise analysis. |Number of modules, each module score and  credits of each module|Finding the module wise min, max, avg, SD, Kurt for given data.
Top, Mid, Bottom blockers of modules.
|Not exactly as expected|
|S_04|Module wise and batch wise analysis result will form the Bar chart. Functionality checking of the graphs by using obtained result.|Data of Top, Mid, Bottom blockers. |Visualisation of input data on graph|As expected|
