# The problem 

A lot off legal teams need to create contract 
However those contracts depend on information coming from the business.

This implies a lot off black and forth between legal representative and business leaders in order to make the contract right

So the usual disadvantages:
- Loss of time 
- Loss of informations
- Loss of data quality

In order to avoid being contacted through many channels, legal department often asks to submit a Google form where business can give all precises answers for the contract 
Then all answers are sent to a Google sheets and placed on a Google docs 

We still have a manual step between the forms and the Google Docs (without talking about the Google sheets  storage count )

So how to send form answers directly to Google docs ?

With an Addon that links your Google forms and a Google docs template to be copied and filled with forms answers 

# The constraints

Since it will be a Google forms Addon , we will need to create an editor Addon 
Therefore , some HTML,CSS and JavaScript will be there 

After that, a drive picker will be used to create an user-friendly way to select thé Google docs template to link with the form. 
We will need the Google picker api and an api Key in pur Google cloud project (the one for the Addon) 
All is described here : The Google Picker API | Google Drive

# Architecture solution 
We need to create a link between the google form and a Google Docs that will be used as a template. To do so we are using the document properties. We will add the Google Docs Id in the Google Form properties for every Google Forms where the addon is used 
For the user to choose the Google Docs to link, we’ll use the Drive picker in a modal dialog. But in order to work as the person using the drive picker, we need a Drive picker Api key as described here : The Google Picker API | Google Drive 
Once the link is made, we need to provide to the user a matching table to show a table with the label for each questions or question/choice duet to use in the Google Docs template. 
- For example : 
  - Question ⇒ What’s your name ? 
  - Label ⇒ A 
  - Label to place in the Google Docs template ⇒ {A} 

The label will then be replaced by the question’s answer.


After that, the script can get answers, copy the Google Docs template and edit the new Google Docs copied from the template.

Here is a diagram to better understand : 

Don’t forget to configure your Google Workspace marketplace SDK API with the script ID and version & as a Google Form Addon(i.e. editor addon)

# Code
The is divided in 3 parts : 
- Code.gs contains every functions to : 
  - Display the Drive picker 
  - Set document properties
  - Copy the Google Docs template
  - Edit the new Google Sheets with text manipulation functions
- instruction.html contains instructions to display
- drivePicker.html contains the drive picker code with javascript functions calling code.gs functions
- authEmail.html contains the email to send if the form trigger requires authorization that has not been granted yet; if so, warn the user via this email. This check is required when using triggers with add-ons to maintain functional triggers.
# What’s next ? 

Apart from the UX that could still be improved, we thought of 3 mains improvements that could be done : 

- Managing more question type in forms ⇒ We only manage Text , Paragraph and List type
- Automatic sharing with the form respondent option ⇒ Once the document is created and edited, having the possibility to share it to the respondent automatically

The code could be improved with you, that's why it's open source !
Fare the well, fellow Apps Scripters














