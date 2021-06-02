# Clean up Content Performance report spreadsheet

* Run the [Content Performance report](https://aka.ms/contentperformancedashboard) and select the **Documentation** tab.

  ![image](https://user-images.githubusercontent.com/3605364/120519165-0acda400-c36e-11eb-805f-bc431c8cf9b4.png)

* Find the `...` and select **Export Data***. Hover around the area in the red circle until `...` appears.

  ![image](https://user-images.githubusercontent.com/3605364/120519553-74e64900-c36e-11eb-9016-9ac35e4b6380.png)

* Select the default **Summarized Data**.

* Open the Excel file, select **Enable Editing**.

  ![image](https://user-images.githubusercontent.com/3605364/120518419-44ea7600-c36d-11eb-9c46-9e3a72799d0a.png)

* **File > Save As Excel Macro-Enabled Workbook (\*.xlsm)**

* [Show the Developer tab](https://support.microsoft.com/en-us/topic/show-the-developer-tab-e1192344-5e56-4d45-931b-e5fd9bea2d45)

* Select **Developer > Visual Basic**

* Right click VBAProject > **Import File** and select *cleanTrafficReport.bas*  > **Open** - You'll need to select VB Files to do this.

* In the VBA window, select **Modules/Module1**.

* Select `Sub EveryThing()` and press the run icon
