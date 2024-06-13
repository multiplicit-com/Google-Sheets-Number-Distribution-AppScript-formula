# Distribution Models formula and Appscript for Google Sheets

This Google Sheets AppScript macro script provides various distribution models for forecasting purposes. Simply provide a target number and specify the desired number of steps in the formula. The script allows you to apply different mathematical models to change the distribution.

All the numbers at each step will add up to the the 'goal' you provided. When expressed as a percentage the percentage of each 'step' will add up to 100%.

Note - there is also a version for Excel, which you can find here: https://github.com/multiplicit-com/Excel-Number-Distribution-VBA . Once you have it installed the instructions for using both versions are almost identical, using the same formula name and parameters in the same order.

## Distribution Models

The script currently supports the following distribution models:

1. **Linear Distribution**: Equally distributes the target value across all months.
2. **Logarithmic Distribution**: Distributes the target value such that the values start low and increase gradually.
3. **Exponential Distribution**: Distributes the target value such that the values start low and increase rapidly.
4. **Normal Distribution**: Distributes the target value according to a normal distribution curve, peaking in the middle.
5. **Quadratic Distribution**: Distributes the target value such that the values start low and increase towards the end, creating a "ski jump" incline which increases more gradually than the steeper and more severe 'exponential' model.

![distribution_models](https://github.com/multiplicit-com/Excel-Number-Distribution-VBA/assets/127529943/ba33b90a-df10-4d72-a0cb-845f72149f7b)


## Installation
To use the AppScript, follow these steps:

1. Open your Google Sheets.
2. Go to Extensions > Apps Script.
3. Delete any code in the script editor and replace it with the VBA code from this repository.
4. Save the script by clicking on the floppy disk icon or pressing Ctrl + S.

## Usage
Once installed, the **DISTRIBUTEGOAL** function can be called like any formula in your Sheet.

### Parameters

* **DistributionType**: the type of distribution model to apply to the target number.
  The accepted distribution types are:
  * linear
  * logarithmic
  * exponential
  * normal
  * quadratic
    
* **TotalMonths**: The total number of months to distribute the goal across.
* **CurrentPosition**: The current position in the month range.
* **Target**: The goal value to be distributed.


### Formula examples

Logarithmic distribution over 7 steps, show position 1:

 **_=DISTRIBUTEGOAL("logarithmic", 7, 1, $A$8)_**


quadratic distribution over 9 steps, show position 8:

 **_=DISTRIBUTEGOAL("quadratic", 9, 8, $A10)_**

