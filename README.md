# Excel To Insert Statements Tool

A small Flet application for making insert statements from Excel sheets

by Zach Sarette in 2024

## Installation

### Python

install python 3.9 if you don't already have it.
https://www.python.org/downloads/release/python-390/

### Virtual Environment

Make sure you are in the correct directory

```shell
cd Excel-To-Insert-Statements-Tool
```

Set up your virtual environment

```shell
python -m venv venv
```

#### Activate the virtual environment

Powershell
```Powershell
./venv/Scripts/Activate.ps1
```

Command Prompt
```command prompt
.\venv\Scripts\activate.bat
```

Bash (mac/linux)
```Bash
source venv/bin/activate
```

### Install the Python Packages

```shell
pip install -r requirements.txt

```
### Run
```
python main.py
```
### Excel File Format
- Sheets must be .xlsx files. Files using .ods extention will not work currently.
- Excel sheet is the name of the table to insert into
- Column headers (the first row) are the name of the fields
- The other rows are the records for each insert statement. They contain the statement's data.

<img width="502" alt="image" src="https://github.com/zacharysarette/Excel-To-Insert-Statements-Tool/assets/650130/f339f606-bffb-4775-821b-b2cb41487bfa">

[Example data source: 'Comparing the Sizes of Dinosaurs in the Lost World' by Giulia De Amicis ](https://www.visualcapitalist.com/cp/comparing-the-sizes-of-dinosaurs-in-the-lost-world/)

### Usage

1. Select File
2. Select Sheet
3. Get Insert Statements (copied to clipboard)

<img width="490" alt="image" src="https://github.com/zacharysarette/Excel-To-Insert-Statements-Tool/assets/650130/fbaf148d-a358-4cfe-890f-8b8af5767ca2">


### Output Example
```sql
INSERT INTO `dinosaurs` (`name`, `length_meters`, `length_feet`) VALUES ('Argentinosaurus', '39', '128.0');
INSERT INTO `dinosaurs` (`name`, `length_meters`, `length_feet`) VALUES ('Brachiosaurus', '26', '85.0');
INSERT INTO `dinosaurs` (`name`, `length_meters`, `length_feet`) VALUES ('Diplodocus', '26', '85.0');
INSERT INTO `dinosaurs` (`name`, `length_meters`, `length_feet`) VALUES ('Barosaurus', '24', '79.0');
INSERT INTO `dinosaurs` (`name`, `length_meters`, `length_feet`) VALUES ('Spinosaurus', '15', '49.0');
INSERT INTO `dinosaurs` (`name`, `length_meters`, `length_feet`) VALUES ('Tyrannosaurus rex', '12', '30.0');
INSERT INTO `dinosaurs` (`name`, `length_meters`, `length_feet`) VALUES ('Iguanodon', '10', '33.0');
INSERT INTO `dinosaurs` (`name`, `length_meters`, `length_feet`) VALUES ('Baryonyx', '10', '33.0');
INSERT INTO `dinosaurs` (`name`, `length_meters`, `length_feet`) VALUES ('Triceratops', '9', '30.0');
INSERT INTO `dinosaurs` (`name`, `length_meters`, `length_feet`) VALUES ('Epidextipteryx', '44', '1.4');
INSERT INTO `dinosaurs` (`name`, `length_meters`, `length_feet`) VALUES ('Parvicursor', '39', '1.3');
```



