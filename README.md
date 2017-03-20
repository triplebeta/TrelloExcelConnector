## TrelloExcelConnector

This project is intended to help users leverage the power of Trello from within their Excel file.
Using this Excel file you can import a Trello board into Excel, publish rows of data as cards and sync back changes from existing cards to your Trello board.

### Features

You can use it for many purposes:
* Make a backup of your board for your archive
* Clone a board with all its cards by importing and re-publishing
* Create an Excel-dashboard based on the status of the cards on your board
* Use Excel to link other types of data to card on your board
* Quickly setup a board using data exported from another source

...and anything else you can think of...

Although I'm aware putting online an Excelfile containing VBA macros might not be considered safe, I do think it's the best
way to publish it so non-developers can use it as well.

Please feel free to download, review and improve the code.

### Download the Excel file
Follow these steps to connect your Excel connection to Trello:
1. Download the [Excel file](https://github.com/triplebeta/TrelloExcelConnector/blob/master/Trello%20Excel%20Connector.xlsm)
2. Find the downloaded file in Windows Explorer
3. Right-click the file, go to properties and Unblock the file
4. Open the Excel file, click the button in the yellow bar to enable the content
5. This will bring up a dialog:
<p><img src="https://github.com/triplebeta/TrelloExcelConnector/img/CannotLoadTrelloConfig.png"/></p>

6. Click Yes and you will see a dialog:
<p><img src="https://github.com/triplebeta/TrelloExcelConnector/img/ConfigureConnection.png"/></p>


### Get your API key
You need to setup the connection details just once:
1. Login to Trello using your account
2. Then change your url to https://trello.com/app-key
3. At the top of the page you find the Developer key  (copy this to the Application Key field)
4. Then scroll to the bottom of the page to get the Secret (copy this to the Secret field)
5. Scroll back up to where you go the Developer key, and click the Token link.
Click `Allow` to grant the application permission to read/write your Trello board:
<p><img src="https://github.com/triplebeta/TrelloExcelConnector/img/Token.png"/></p>

6. Paste this value in the field Read/Write token.
7. Click OK to store the settings and you're good to go!

Your settings are saved in your Documents folder in a file named "Trello Excel Connector.ini"

_Warning: keep these values secure!__

### Using the Trello tab
The Excelfile will show you an extra tab named TRELLO

![](https://github.com/triplebeta/TrelloExcelConnector/img/TrelloTab.png)

Play around with them, use one worksheet for each Trello board.
