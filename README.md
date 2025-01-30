C# JSON to Excel Converter
C# console application that reads a JSON file, deserializes its data into a model class, stores the values in a DataTable, and exports the DataTable to an Excel file (.xlsx) using the OpenXML library

Overview
This application performs the following tasks:
1) Reads a JSON file located in the JsonFile folder within the solution directory.
2) Deserializes the JSON into a strongly-typed model class, handling nested arrays and objects.
3) Converts the deserialized object into a DataTable format.
4) Exports the DataTable to an Excel file using the EPPlus library.

Prerequisites
To run the application, make sure you have the following:
1) .NET Core SDK or .NET Framework installed on your system.
2) EPPlus library for working with Excel files.



