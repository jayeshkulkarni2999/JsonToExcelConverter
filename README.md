# C# JSON to Excel Converter

## Overview
**C# JSON to Excel Converter** is a C# Console Application that reads a JSON file, deserializes its data into a model class, stores the values in a **DataTable**, and exports the DataTable to an **Excel file (.xlsx)** using the **OpenXML** library.

## Features
- Reads JSON data from a file located in the `JsonFile` folder.
- Deserializes JSON into a strongly-typed **model class**, handling nested objects and arrays.
- Converts the deserialized data into a **DataTable** format.
- Exports the DataTable to an **Excel file (.xlsx)** using the **OpenXML** library.
- Supports structured and hierarchical JSON data.

## Prerequisites
Ensure you have the following installed:
- [.NET Core SDK or .NET Framework](https://dotnet.microsoft.com/download)
- [OpenXML SDK](https://www.nuget.org/packages/DocumentFormat.OpenXml/) for working with Excel files
- Git (if using version control)

## Setup Instructions
1. **Clone the Repository:**
   ```sh
   git clone <repository-url>
   cd CSharpJsonToExcelConverter
   ```
2. **Restore Dependencies:**
   ```sh
   dotnet restore
   ```
3. **Build the Project:**
   ```sh
   dotnet build
   ```
4. **Run the Application:**
   ```sh
   dotnet run
   ```

## Git Setup
Make sure `App.config` is ignored in `.gitignore`:
```
App.config
```

### Git Commands to Push Code
```sh
git init
git remote add origin <repo-url>
echo App.config > .gitignore
git add .
git commit -m "Initial commit"
git push -u origin main
```

## How It Works
1. The application reads the **JSON file** from the `JsonFile` folder.
2. It deserializes the JSON data into a **C# model class**.
3. The deserialized data is converted into a **DataTable**.
4. The **OpenXML** library is used to export the DataTable to an **Excel file (.xlsx)**.
5. The generated Excel file is saved in the output directory.

## Configuration
- Ensure the JSON file follows a structured format to be correctly parsed.
- Modify the model class as needed to match the structure of your JSON data.

## Testing
- Place sample JSON files in the `JsonFile` folder.
- Run the application and verify that the generated Excel file contains the expected data.
- Check logs for any errors related to JSON parsing or Excel file creation.

## License
This project is licensed under the MIT License.

