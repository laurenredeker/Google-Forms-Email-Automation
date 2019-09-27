# Google-Forms-Email-Automation

During an internship, I created a portal where users could go to fill out a Google Form for various system troubleshooting. Using Google Sheets, my team could:<br/>
* Recieve alerts when new responses appeared on the Google sheet and easily access, track and organize them
* Send automated emails to users with a reference number to update them on the status of their request rather than having to manually email them</br>

Attached are samples of the solution I used with Google Apps Script. 

**Sample 1**</br>
View each form entry through Google Sheets and send automated email responses to users to update them on the status of their request<br /> 
  * Google Form: https://docs.google.com/forms/d/e/1FAIpQLSce38Bq0yD-qopugheO8GpJTTvb4WCtgt5YYLYLp9MQ8Pbodg/viewform?usp=sf_link
  * Sample1_Example.csv is a sample Google Sheet in excel

**Sample 2**</br>
Sometimes, a user may have to fill out the same request more than once for different instances. This form allows users to fill out a request more than once without having to exit the Google Form (which saves the user time / is more convenient).</br>
* Problem: When a Google Form iterates the questions multiple times through the same form, Google Sheets displays all responses in the same row, rather than each response having its own row
* Solution: This script organizes a form response that has multiple responses row by row on another sheet for greater visibility of each unique request (in this sample you can fill out the form up to 5X).
* Google Form: https://docs.google.com/forms/d/e/1FAIpQLSfGlYfZjMTings01E1-Nxu0OhmBfHgakHnR9R_oekBQ-4h15A/viewform
* Sample2_Example.xlsb is a sample Google Sheet in excel
