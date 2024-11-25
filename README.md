# VBA-Excel

VBA Macros for Excel

This repository contains useful VBA macros for Microsoft Excel. The macros are organized into _packages_. 

A package encapsulates a single functionality, or a set of related functionalities. It may contain VBA modules (.bas files), class modules (also .bas files), user forms (.frm and .frx files), and Office Ribbon customizations (a file called customUI14.xml in a directory called CustomUI). It must have at least one VBA module.

Each subdirectory under the `packages` directory in this repository contains an independent package. Each package directory also has a `Tests` subdirectory, which contains a .xlsm file which has test cases and usage instructions for the package, as well as the package code.

## Using these macros

### To use as an Excel Add-in (Recommended)

1. Download the `MaxOffice-Excel-Macros-Collection.xlam` file from the [latest release](https://github.com/MaxOffice/VBA-Excel/releases/latest).
2. Enable it from the Excel user interface as an Add-in:
    1. Click the **File** tab
    2. Select **Options** (Windows) or **Preferences** (Mac)
    3. Click **Add-ins**
    4. At the bottom of the window, in the "Manage:" dropdown:
        - Select "Excel Add-ins"
        - Click **Go**
    5. In the Add-ins dialog, click **Browse** and navigate to where you saved the .xlam file
    6. Check the box next to the Add-in name to enable it
    7. Click **OK**

3. Access all macros through the Macros button on the View ribbon tab, or the Efficiency 365 tab installed by the Add-in.

### To use all packages without installing the add-in

1. Download the MaxOffice-Excel-Macros-Collection.xlsm file from the [latest release](https://github.com/MaxOffice/VBA-Excel/releases/latest).
2. Open it, remembering to enable macro content. Keep it open.
3. Invoke any macro from the Macros button on the View ribbon tab, or the Efficiency 365 tab installed by the file.

### To use an individual package

1. Download the .xlsm file in the package's `Tests` subdirectory.
2. Open it, remembering to enable macro content.
3. Invoke the macro from the Macros button on the View ribbon tab. 


## To use packages of your choice

1. Create or open a .xlsm file
2. Use `File/Import File...` to import all files found directly under all your chosen package subdirectories. These may include `.bas`, `.cls`, `.frm` and `.frx` files. Do  not import anything in the `Tests` subdirectories. 
3. Save the file, and keep it open.
4. Invoke any macro from the Macros button on the View ribbon tab.
