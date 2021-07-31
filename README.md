# Clean up Content Performance report spreadsheet

The data from the [Topic-level Content Performance report](https://aka.ms/contentperformancedashboard) can be downloaded to Excel, but the spreadsheet is not easy to use. To make it more convenient to use, this VBA module makes changes such as:

* Deletes the two extra rows at the top of the spreadsheet.
* Pins the top row so the headings are visible while you scroll down.
* Hides columns that aren't usually useful.
* Removes redundant words from column headings.
* Sizes column widths for readability.

To use the Cleanup VBA module, follow these steps:

* CLone or download this repo to your local machine.

* Run the [Content Performance report](https://aka.ms/contentperformancedashboard) and select the **Documentation** tab.

  ![image](https://user-images.githubusercontent.com/3605364/120519165-0acda400-c36e-11eb-805f-bc431c8cf9b4.png)

* Find the ellipsis (`...`) and select **Export Data**. Hover around the area in the red circle until `...` appears.

  ![image](https://user-images.githubusercontent.com/3605364/120519553-74e64900-c36e-11eb-9016-9ac35e4b6380.png)

* Select the default **Summarized Data**.

* Open the Excel file, and select **Enable Editing**.

  ![image](https://user-images.githubusercontent.com/3605364/120518419-44ea7600-c36d-11eb-9c46-9e3a72799d0a.png)

* If you want to be able to save the VBA code with the spreadsheet, select **File** > **Save As Excel Macro-Enabled Workbook (\*.xlsm)**. If you skip this step you can still save the cleaned-up spreadsheet as an *.xlsx* file without the VBA code.

* [Show the **Developer** tab](https://support.microsoft.com/topic/show-the-developer-tab-e1192344-5e56-4d45-931b-e5fd9bea2d45)
* Select **Developer** > **Visual Basic**.
* Right click **VBAProject** > **Import File**.
* Select the *Cleanup.bas* file  and then select **Open**.
* Expand **Modules** and then double-click **Cleanup**.
* Click somewhere in the `Sub EveryThing_ASPNET()` or the `Sub EveryThing_DOTNET()` line, and select the **Run** button.

  The repo currently has code customized for ASP.NET Core and .NET; to customize for your docset, create subroutines equivalent to the ones with _ASPNET or _DOTNET suffixes on the names. For example, create a set of _EFCORE subroutines.
