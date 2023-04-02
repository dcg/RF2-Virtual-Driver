# RF2 Virtual Driver

RF2 Virtual Driver is a tool that allows the creation of virtual drivers for rFactor2. It is developed using Python and requires the openpyxl package to be installed.

## Installation

To install RF2 Virtual Driver, follow the steps below:

1. Install openpyxl using pip: `pip install openpyxl`
2. Modify the `mod_dir` and `vehicles_dir` parameters in the `init.py` file if needed
3. Modify the `output_dir` parameter in the `create_league.py` file if needed

## Usage

### Prepare the content

1. Run `python init.py` to create a configuration file `config.xlsx`
2. In the `config.xlsx` file, select the cars you want to extract liveries from by changing the value from `False` to `True`
3. Run `python init.py` again to extract the liveries to the `Data` folder

### Setup a Roster

1. Run `python create_league.py` to create a dummy `data.xlsx` file
2. Add liveries (with the folder names and vehicle files from the initial created data folder)
3. Add drivers
4. Run `python create_league.py` again to generate the data for rFactor2
5. Repeat step 4 as needed, as files are simply overwritten.
