# How to Use the HMC BBQ Add-On

**What is it?**  
The HMC BBQ Add-On is a script (installed with Google Sheets) that helps each Harvey Mudd dorm manage their BBQs automatically. Using the add-on, dorms can set up automatic BBQ request submission to the dining hall and automatic reminder emails sent to their BBQ participants. This way, BBQ requests no longer have to be submitted via the usual [BBQ Request Form](https://hmc.formstack.com/forms/bbqrequest); the script will do that for you. The hope is that this will make it significantly easier to manage BBQs and decrease the likelihood of the occasional mistake.

**How to Install | (Quick and Easy\!)**  
Installation is easy\! Just make sure that you’re signed into your HMC account.

1. If you already have a Google Sheets spreadsheet for managing BBQs at your dorm, open it. If you don’t, open a new, blank spreadsheet.  
2. Click on the menu item for “Add-ons” and then click on “Get add-ons…”. A pop-up should appear (shown below).  
   ![image](https://github.com/user-attachments/assets/dbe9ceea-fd40-4af8-a3bd-6fac235d04a1)
3. In the top right drop-down menu choose “For g.hmc.edu” and, if necessary, type the name of our add-on “HMC BBQ” in the search box in the upper left corner.  
4. Choose our add-on “HMC BBQ” and install it.  
5. Confirm any permissions and then let the program add the necessary sheets to your spreadsheet.  
6. Now open the protected “Admin” sheet (at the bottom of the page) and enter the necessary info. The sheet will help you enter the right data. **Don’t forget to enable the app and enable email reminders by changing both boolean values to 0\!**  
7. *If you already had a sheet for managing BBQs, you must copy the data from that sheet to the one the app created for you.* Afterward, we recommend deleting all other sheets besides the ones that the app just created for you.

You’re done\! You don’t have to keep reading, but you can certainly continue if you want some more advanced info.

\------------------------------------------------------------------------------

**Notes and Recommendations**

* If you ever need to disable the app for a week or a month, we recommend simply changing the boolean value in the “Admin” sheet.  
* At the beginning of each school year when you need to transfer ownership of the BBQ Script, simply [uninstall the script](https://dottech.org/175256/how-to-install-and-uninstall-add-ons-in-google-docs-guide/) under your account and let the new BBQ Frosh reinstall the script under their account (by having them follow [the same installation instructions](#bookmark=id.notz09y4zoj2) you did). The script is configured to help you transition easily between re-installs.  
* The add-on will send you an email after each time it submits the BBQ form. If you don’t get an email an hour after the time you specified on the Admin sheet, an error likely occurred, in which case you should just submit the BBQ form manually.  
* Upon install, the script will change the sharing settings of your Google Sheets file to “editable by link only”. *After that, don’t change the sharing settings of your Google Sheets file.* Doing so will mess up the script.  
* *Be careful editing any of the cells in The List from A1:D5* (essentially, anything from Row 5 and up), especially B2, B4, and C4. These are automatically updated and managed by the script. If you change them, it might mess up the script.  
* *Don’t change the format or arrangement of anything in the Admin sheet*. The script depends on the location of each cell, the data validation, and the auto-formatting to function properly.  
* Both the Admin sheet and everything above Row 5 in The List sheet have been “protected”, which means only you can edit them. This way, you don’t have to worry about anybody else messing up the script when your sharing setting is “Editable by anyone with the link.” If you’d like, you can give permissions to specific people, so that they can edit the Admin sheet, but it is recommended that you don’t give anybody permission to change anything at or above Row 5 in The List.  
* Try not to delete the trigger the script creates for you. The only way to recover from this is to uninstall and reinstall the app. If you don’t know what triggers are or how you could delete them, then don’t worry about it :)

**Ok. That’s great. But what does the script *actually* do?**  
We get it. You want the details. Here’s a bit of an overview.

Upon install, the add-on will generate two preformatted sheets (so that it knows how to read them) and add them to your spreadsheet. It will also change the sharing settings of the spreadsheet to “Editable by link” so that it (and your BBQ participants\!) can access the spreadsheet later. Lastly, it will add “protections” to the Admin sheet and to a specific range A1:D5 from The List, so that nobody but you can edit those cells.

Note that the script will also store your email in a list accessible by me, the author of the script. This is so that I can contact you if a problem with the script ever arises.

The script will also create a “trigger” for you. This trigger will run every hour (at approximately the same time each time) and check whether the script needs to either

1) submit the BBQ request form to the dining hall

	or

2) send an email reminder to the BBQ participants, reminding them to either subscribe or unsubscribe from the BBQ list.

If you’ve set the app to “disabled,” none of these functions will run. If you’ve set email reminders to “disabled,” the BBQ request will be submitted, but an email reminder will never have been sent to the BBQ participants reminding them of it.

Every time either of these functions are run, the script will also clean up the list to get rid of any empty spaces between entries. It will also automatically populate the “Date of BBQ Event:” entry with the correct date.

When the script detects that it’s time to submit the BBQ request form, it will automatically export The List as an excel doc and submit it along with the data you provided. It will also count the number of vegetarians on your list by looking for “vegetarian” or “Vegetarian” next to each student ID in the “Notes” section. In essence, it fills out the form for you.
