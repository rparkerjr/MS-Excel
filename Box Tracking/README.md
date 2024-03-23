# Box Tracking
Custom workbook to facilitate the tracking of boxes through a production process. This repository only includes the VBA and obfuscates identifying information. The workbook itself is branded with company logos in various spots and thus not publicly available.

# Description
Users create individual project tabs via a form (accessed via a button on the Master Tracking sheet). Data entered into the form is passed into the new project sheet and a range of box numbers is automatically generated and added to the table. Summary data from each project sheet is pulled into a Master Tracking sheet via cell formulas. 

## MainModule.bas
Helper functions for initializing user forms and updating pivot tables throughout the workbook.

## NewProject.frm and .frx
This is the code behind the New Project form, including data validation and the creation of a new project tab based on a template worksheet.

## TabListing.frm and .frx
Code for a simple listing of all tabs in a workbook and their visibility status (Visible/Hidden). Allows the user to selectively hide/unhide multiple worksheets from a single window.
