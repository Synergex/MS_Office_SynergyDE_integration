SYNERGY & MICROSOFT WORD - INTEGRATION DEMO
===========================================
AUTHOR:         
Steve Ives
Synergex
<steve.ives@synergex.com>

DISCLAIMER:
This software is provided "as is" and without warranty. Synergex 
accepts no responsibility from any loss or damage which may result 
from the use of this software.

INTRODUCTION:
The files provided present a demonstration of the integration of 
Synergy/DE with Microsoft Office 97.  Integration is achieved using 
VBA (Visual Basic for Applications), the scripting language available 
in the Microsoft Office 97 applications.  Essentially, the files 
provided enable Microsoft Word to communicate with a Synergy ISAM 
database via the Synergy/DE xfODBC Driver.

The demo provides two Microsoft word templates as follows:
* Back Order Letter.dot
* SynTools.dot

REQUIREMENTS:
* Microsoft Office 97
* Synergy/DE. V6.1 (or higher), with a standalone xfODBC installation 
  and the demo database configured.

INSTALLATION:
* Install Microsoft Office 97 (if it is not already installed).  When 
  installing Office 97, you must select the optional "Data access 
  objects for Visual Basic" component.

* Install Synergy/DE if it is not already installed.  You must install 
  the "Connectivity" option and during the connectivity installation 
  select a "Standalone" ODBC installation.

* Configure the Synergy ODBC demo database (if it is not already 
  configured).  To do this, select the "Install Sample Synergy 
  Database" option in the "SynergyDE" start menu folder.

* Move the "Back Order Letter.dot" file into your Microsoft Office 
  templates directory (usually C:\Program Files\Microsoft Office\Templates

* Move the "SynTools.dot" file into the Microsoft Office startup 
  directory (usually C:\Program Files\Microsoft Office\Office\Startup

* Run Microsoft Word

* Select the "Tools" menu, then the "Templates and Add-Ins" menu entry.

* If the "SynTools.dot" global template is not checked then check it, 
  and press OK.

* Right-click an area on the toolbar, and select "Customize"

* In the "Toolbars" tab you should now see a toolbar called "Synergy 
  Tools".  If this toolbar is not checked then check it.  Press the 
  "Close" button.

* If the new toolbar has appeared as a "floating" toolbar, then you 
  may prefer to dock it on one of the main toolbar areas by dragging 
  it with the mouse.

NOTES:
When using these templates remember that this simply a demonstration 
of what is possible, not a comprehensive demonstration.  These 
templates are programs, like any other, and can be coded to behave in 
any way you require.

USING BACK ORDER LETTER.DOT:
Back Order Letter.dot is a document template.  Document templates 
facilitate the repeated creation of "standard" documents (in this case 
a letter to a customer).

Run Microsoft Word and select "File/New" from the menu (do not click 
the "New" button on the toolbar, as this always creates a new document 
based on the standard template (Normal.dot).  The "New" dialog will be 
displayed, and in the "General" tab you should see the "Back Order 
Letter" template. Click on the template and press "OK".  A new standard 
customer letter will be created.

On using this template for the first time you will see a setup dialog, 
which prompts you to enter your name and job title.  In addition you 
can set various options which affect the behaviour of the document 
template.  The options available are:

* Automatic print
* Confirm before printing
* Automatic save
* Close document after save

These options, along with your user information, are normally saved in 
the Windows Registry for future use.

Next you will see another dialog which prompts you to enter a customer 
code.  Having entered a customer code Word will retrieve the required 
data for the selected customer and insert the customer name, address, 
contact name etc. at appropriate points in the letter.

The template is hard-coded to communicate with a demo database, which 
is provided with the standalone version of the Synergy/DE xfODBC 
Driver.  Valid customer codes are 1 to 37. 

USING SYNTOOLS.DOT:
SynTools.dot is a global template.  Global templates allow Microsoft 
Word to be customized, or have additional functionality added.  This 
functionality will be generally available to the user when they are 
editing any document.

The template adds a custom menu column called "Synergy Tools" and a 
custom toolbar called "Synergy Tools".  Both contain four entries, 
which will insert a specific piece of information about a selected 
customer into the current document at the current position:

* Customer Address
* Customer Credit Limit
* Customer Contact Name
* Customer Phone Number

When you select any of these functions you will see a dialog, which 
prompts you to enter a customer code. When you enter a customer code 
the database is referenced and the resulting information is inserted.

In this example the customer code field also has been given a drill 
button.  Pressing the button causes a list of available customers to 
be displayed.

When a customer has been selected, that customer number will be the 
default the next time you select any of the available functions.  If 
you press the "Cancel" button, the default is cleared.

VIEWING THE SOURCE CODE:
Both demonstrations are delivered in Microsoft Word template files.  
To view the source code, use Explorer to navigate the directory 
containing the template file, then right-click the file and select 
"Open".  When prompted with the "Warning" dialog, press the "Enable 
Macros" button.  Next, select "Tools / Macro / Visual Basic Editor" 
from the menu to display the source code contained within the template 
document.

Note that if you have a copy of SynTools.dot in your "Startup" 
directory when working on the SynTools.dot document, you will see two 
"Synergy Tools" menu columns and two "Synergy Tools" toolbars.  One is 
from the global template in the Startup directory, the other from the 
current document.  This can become confusing, you may prefer to remove 
SynTools.dot from the "Startup" directory before working on the 
template files.
