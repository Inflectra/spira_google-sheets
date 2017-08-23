
# SpiraTeam Google Sheets Integration Add-on
The web-based interface of SpiraTeam® is ideal for creating and managing requirements, test cases and incidents for a new project. However when migrating requirements, test cases, test steps and incidents for an existing project from another system, it is useful to be able to load in a batch of artifacts, rather than having to manually enter them one at a time.

To simplify this task  we’ve created a Google sheets add-on for SpiraTeam® that can export requirements, test cases, test steps and incidents from a generated spreadsheet into SpiraTeam®.

![SpiraTeam Google Sheets Integration Screenshot](https://github.com/inflectra/spira_google-sheets/blob/master/src/assets/screenshots/STGSIGithubScreenshot.png)


## Installation
The add-on runs using the Google App Script Engine so all of the files must be downloaded into a single file and run using Google App Script. [Google App Script Docs](https://developers.google.com/apps-script).

The easiest path is to simply create a new Google App Script project and import all of the downloaded files, then you may test the code in an IDE-like test environment.


## Usage
You must have a SpiraTeam® account with the proper permissions to utilize this app. For usage instructions reference the SpiraTeam® documentation located at [SpiraTeam® Product Add-ons and Downloads](https://www.inflectra.com/SpiraTest/Downloads.aspx#ImportTools).



---

Hello,

My name is Toni. I was previously an intern at Inflectra and I was tasked with adding functionality to spira_googlesheets.  I was able to get Inflectra’s Rest API to work in conjunction with google sheets in order to create a task template.  The next goal for this project would be to use Inflectra’s REST API to post this template onto SpiraTeam.  After this is completed the next goal would to be to add more artifact functionality to spira_googlesheets.  One of the difficulties you may face when trying to accomplish this in the current code setup. The program was originally built with requirement artifact functionality in mind. This has led to a number of hard codings throughout the software.  In order to aid in the lifecycle of this software it would be useful to refactor those hard codings out of the software. 

The text editor you will be most likely using is google scripts. This google creation comes with a number of built in functionalities that are called throughout the software. Here is a reference https://developers.google.com/apps-script/overview

