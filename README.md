fresh-sheets is a simple python script that can automate updating a Google Docs spreadsheet from CSV data.  In order to use this script you will first need to set up some things in your Google account.  Follow the steps below to get started:

## Google Account Setup

You must enable the Google Sheets API in your Google account and create a service account for the fresh-sheets script to use when updating your spreadsheet.

- Visit the [Developer Console](https://console.developers.google.com)
- In the upper-left corner (next to the Google APIs logo) click the project selection button and create a new project or select an existing one.
- Click the "Enable APIs and Services" button.
- Search for Google Sheets API and enable the API for your project.
- Return to the [Developer Console Dashboard](https://console.developers.google.com).
- On the left, click on "Credentials".
- Click the "Create credentials" button and choose "Service Account key"
- Select "New service account" and give it a sensible name such as "sheet-updater".
- For Role, choose Project->Viewer.
- Leave JSON selected for "Key type" and click "Create".
- Collect the file that just downloaded.  You will need to use it with the fresh-sheets script.

## Spreadsheet setup

You must grant the service account edit access to the desired spreadsheet in order for updates to be allowed.

- Return to the [Developer Console](https://console.developers.google.com)
- Click "Credentials" and choose "Manage service accounts" in the content pane.
- Copy the email address associated with the service account you created in the previous section.
- Visit your target spreadsheet in Google Docs.
- Make note of the URL which should look something like: https://docs.google.com/spreadsheets/d/1bOBnVx8RgJjzKV9sa6M5F91oQc9wOn39iXmc8wZsG18/edit#gid=1392089204
- The URL contains the document ID and the sheet ID.  In this example the document ID is the component beginning with 1b and the sheet ID is the number 1392089204.  You will need these two values to configure your script run.
- Click Share in the upper-right corner of the document view.
- Grant edit access to the service account email address.

## Running the script

Congratulations!  Setup is complete and you are ready to run the script.  You will need the following:

- CSV Data: It's outside the scope this project how you generate this data.  You will pipe it to the fresh-sheets command.
- Google service account key: Generated in the first section of this guide.
- Document ID: Identifies the document you want to update (discovered from the URL of your spreadsheet)
- Sheet ID: Identifies the sheet/tab within the document (discovered from the URL of your spreadsheet)

Optional parameters:
- Starting row/column: By default the entire sheet is replaced by fresh-sheets.  If you have a header or other static content at the top or left of the sheet you can use these values to preserve those areas.
- Delimiter: By default the separator is \t (Tab) but you may change this if your CSV data uses a different delimiter.

To run the script, pipe the CSV data to the script an supply required and optional parameters, for example:

generate-csv-data | fresh-sheets -k sa.key -d 1bOBnVx8RgJjzKV9sa6M5F91oQc9wOn39iXmc8wZsG18 -s 1392089204

In some cases it may be more convenient to supply parameters as environment variables (ie. for containerized deployments).
See the usage help for the environment variables that can be set.