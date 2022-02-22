# OfficeJS Addin for Unsplash
Sample code of a Word, OneNote & Powerpoint Add-In to insert PHOTOS from UNSPLASH.com.
It leverages Unsplash API.

## Register an Unsplash.com ApplicationID
* Register an application at [UNSPLASH Developer program](https://unsplash.com/oauth/applications)
* Create an .env file in the root folder and put in...

    `REACT_APP_UNSPLASH_API_KEY=<your-unsplash-client-id>`
* Make sure you honor the rules of the Unsplash Developer Program when making any changes.

## Create Azure ApplicationInsights
* [Create a new AppInsights resource in Azure](https://docs.microsoft.com/en-us/azure/azure-monitor/app/create-new-resource)
* Add this to your .env file

    `REACT_APP_APPINSIGHTS_API_KEY=<your-appinsights-key>`
* IF you decide to NOT use AppInsights, you need to at least have to put an empty value to the .env file

    `REACT_APP_APPINSIGHTS_API_KEY=`

## Run the code & debug in Office
* Run `npm start`
