# Checking-the-quality-of-Revit-and-IFC-projects

![](https://opendatabim.io/wp-content/uploads/2023/02/projects-data-2.gif)

> In the construction industry, having accurate and high-quality building information models (BIM) is essential to the successful completion of a project. 

## Solution logic
With the help of a table describing the parameters in the category to be checked and with the help of Pipelines from OpenDataBIM – we check the projects in Revit and IFC formats. We automatically generate a PDF document, which contains the basic information about the parameters and categories to be checked.

![](https://opendatabim.io/wp-content/uploads/2023/02/Solution-Check-Data-Revit-IFC.gif)

## Prerequisites

Before you dive into running the application, you need to prepare all the components for the start-up:
1.  A recent version of Python: You can download the latest version of Python from the official website (https://www.python.org/)
2. You need to get the data from the project as a CSV or XLSX Excel table. To export Revit or IFC format into tables, you can use any solution you work with:

> 👨‍💻 **Manual extraction of tables:**
Revit & Dynamo, Revit & Schedule, Revit & ODBC, pyRevit, Forge, SimpleBIM, Desite, IfcOpenShell, IFCjs and others
> ⚙️ **Automatic and batch table retrieval: using noBIM Lite converters** (Revit 2018, Revit 2019, Revit 2020, Revit 2021, IFC2X3, IFC4X1, IFC4X, IFC4 – IFC4.3), noBIM Full converters (Revit 2015, Revit 2016, Revit 2017, Revit 2018, Revit 2019, Revit 2020, Revit 2021, Revit 2022, Revit 2023, IFC2X3, IFC4X1, IFC4X, IFC4 – IFC4.3)

![](https://opendatabim.io/wp-content/uploads/2023/02/github.com-OpenDataBIM-5.gif)

## Excel with checking rules
Any project elements in Revit and IFC format belong to some group (OST_ in Revit, IfcEntity in IFC) or type group (TypeName in Revit, ObjectType in IFC). Also, each element has its own unique set of properties and parmeters (which are columns in the overall project table when the noBIM concept is applied).

Using the check table, we specify in the first column of the table, for the elements to be checked and tested, first the category (OST_ in Revit, IfcEntity in IFC) or type (TypeNamein Revit, ObjectType in IFC) and in the second column we specify the parameters that we want to check for this group of elements.

![](https://opendatabim.io/wp-content/uploads/2023/02/Excel-Check-Revit-IFC-Project.gif)

For each parameter from the check table in the project will be checked:
- presence of the parameter for the group of elements
- percentage of the content of the values in the parameter
- the unique values of this parameter
Once you have the prerequisites installed, you can obtain the source code for the application. You can do this by either cloning the repository using Git or by downloading a zip file of the code from Github.

## Clone or Download



    git clone  https://github.com/OpenDataBIM/Checking-the-quality-of-Revit-and-IFC-projects.git

Install required libraries: The libraries required by your code are specified in the requirements.txt file. To install these libraries, run the following command:




    pip install -r requirements.txt

In the code ODB_Pipeline_Batch_reporting_Revit_IFC_projects.py change the parameters – the path to the folders where your projects are located: pathfold, path, namemap, path_conv




    ########################   Parameters    ########################
    # Folder where the converter are located
    pathfold = r'C:\ODB_Check\ODBCheck\\'
    
    # Folder where the conversion files are located
    path = r'C:\ODB_Check\Sample\\'
    
    # Path to the Excel file with test parameters
    namemap = pathfold + '\ODB Table_of_Parameters_Revit _IFC_Check.xlsx'
    
    # path to noBIM converter (for streaming file conversion)
    path_conv = r'C:\ODB_Check\noBIM_Lite_v1_23-v2jfja\\'
    ################################################################


## Run the code


Now that you have installed Python and the required libraries, you can run your Python code. Simply navigate to the folder where your code is located and run the following command:


    python ODB_Pipeline_Batch_reporting_Revit_IFC_projects.py
    
With these steps, you should be able to run Python code on a new computer without any issues.

![](https://opendatabim.io/wp-content/uploads/2023/02/Excel-Check-Revit-IFC-Project.gif)

## Troubleshooting

 If you encounter any problems, make sure to check the requirements.txt file and ensure that all the required libraries are installed in your virtual environment.

## OpenDataBIM

 At OpenDataBIM, we strive to provide the best possible support to our users. We understand that errors and issues can arise, and we’re committed to finding solutions to these problems as quickly and efficiently as possible.

To that end, we offer consultations with our team of experts to help users address any issues they’re facing. Our team has extensive experience working with BIM software and data, and we’re always happy to share our knowledge and expertise with our users.

Contact Us: info@opendatabim.io
