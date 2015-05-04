NormDistS-Excel-Function-VBA
===========================

If you want to model a normal distributed bell curve in a MS Excel worksheet, you would need to calculate a lot of values to use `=NORMDIST(x; mu; sigma; cumulative)` and any method of implementing this standard worksheet-function won't give you an interactive worksheet for scenarios.

The `=NORMDISTS(x; UVal)`-function is easier to use and gives you more control over an interactive model of bell-curve-shaping and normal distribution over a defined (or even partially unknown) period of time.

#### Install
The installation proces is easy: you just need to import NormDistS.bas into your desired worksheet from the VBA-editor and you are ready to use `=NORMDISTS(x; UVal)` inside your worksheet.
Tip: If you import NormDistS.bas into an empty workbook you can save this workbook as an Excel-AddIn (.xla / .xlam) for use in all your workbooks.

#### Variables
* `x` is the respective position on the x-axis (e.g. the day/month/year on your timeline)
* `UVal` is the highest point of your data (or the exact middle of your time period)

#### How to use
* `=NORMDISTS(...)` only retrieves a factor to be used with the a sum of things you want to distribute
* E.g.: You want to model the onboarding of 100 customers over 24 months and already know the twelfth month to be the highest point. The resulting formula woul like like this: `=100*NORMDISTS(x; 12)` with `x` being the month for which you want to retrieve the number of customers.
* Consider the month #0, when choosing UVal.
* This function can be used in conjunction with known values. If you have periodized data, you can start from any following month using `=NORMDISTS(...)`. Tip: `UVal` needs to be equal to all months and not just the remaining.
