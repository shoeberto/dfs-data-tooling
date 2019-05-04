# dfs-data-tooling

A tool for batch conversion and validation of field-collected vegetation data.

### Installation

This tool requires a minimum of Python 3.7, which can be downloaded from here:
[Download Python](https://www.python.org/downloads/)

If you are using Windows, it is recommended that you select the option in the installer to add Python to your path:
[Add Python to PATH](https://docs.python.org/3.7/using/windows.html#finding-the-python-executable)

You will need to set up a Python library called `pipenv` to manage this tool. Open the command line and type:
`pip install --user pipenv`

Now, clone or download the repository. If downloading the zip file, extract it somewhere on your computer.

On the command line, go to the directory where you extracted the project files, such as:
```
user> cd /home/[your_name]/[project_path]/
```

Run this command to invoke a pipenv shell:
```
user> pipenv shell
```

Install the project dependencies:
```
(dfs-data-tooling-dGjV9NyZ) user> pipenv install
```

Now run the `run.py` script. You should see the following output:
```
(dfs-data-tooling-dGjV9NyZ) user> python run.py
usage: run.py [-h] -i INPUT_DIRECTORY -o OUTPUT_DIRECTORY -y YEAR
              [-t {plot,treatment,superplot}]
```


### Usage

This tool ingests a directory full of Excel files (in the .xlsx format - **.xls is not supported!**), validates that they match the expected criteria for data collection, and then re-writes them to a unified output format.

The arguments at the command line are used to control the following processing parameters:

* -i: Input directory. Must be a path on your computer containing raw input files in .xlsx format.
* -o: Output directory. Must be a path on your computer where the converted files will be written to.
* -y: Year. The year that the input data was collected. This is used to determine how the input data will be parsed.
* -t: Data type. May be `plot`, `treatment`, or `superplot`. Affects both how data is parsed and written. Default is 'plot'.

All files in the input directory are processed when the script is invoked. There is currently no way to process only a single file at a time.

Once the process starts running, all validation errors are printed directly as they are encountered. Typically I recommend piping the output to a file to assist in debugging errors. You can accomplish this on Linux and Mac OSX using output redirection:
```
(dfs-data-tooling-dGjV9NyZ) user> python run.py -i /home/user/input_data/ -o /home/user/output_data/ -y 2015 > /home/user/input_data/validation_output.txt
```

The validation output will include specific details about errors in the source files. All errors should be investigated, but unless explicitly flagged as "fatal", all errors can potentially be ignored. It is up to the user to investigate and determine the severity of all errors that are emitted. Ideally, by the time the source data has been reviewed and fixed, there will be no error output produced by the script.

Errors should be corrected in the original input data, and after the errors have been fixed, the script should be re-run. Sometimes fixing an issue, particularly fatal ones, will lead to new errors being discovered in the file. As such, this process should be considered iterative, and will need to be repeated until the source files have achieved an acceptable level of data cleanliness.

### Methodology

This tool uses a lightweight ETL pattern to extract all of the data from the sources using specialized parser. The parsers read the Excel spreadsheets based on how humans would interpret them; for example, knowing that cell B1 on the "General" sheet was used to store the data collection date in 2013.

This parsed data is then stored into a generalized data model for all of the collected data. This data model normalizes the aspects of all data that were collected across all years, and contains validation rules for how the data is expected to look. For example, the data model knows what fields can be empty or required, and what values are considered valid for any given field, such as specific species names.

After validation, the normalized data is re-written into a standardized Excel format that can be used for import into Microsoft Access.
