
# spHelper.js

**spHelper.js** is a lightweight JavaScript class created to simplify communication to SharePoint services using JSOM (JavaScript Object Model). **spHelper.js** will shorten the amount of code you need to write, as well as reduce repetition. This class can manage same and cross domain requests.

## Builds

Due to the various ways you can develop on the SharePoint platform there are multiple builds available to use depending on your needs.

-  **build/spHelper.js**
	- Node module build (NPM).
- **build/spHelper-stand-alone-min.js**
	- Independent JS library build (minified).
- **build/spHelper-stand-alone-poly.min.js**
	- Independent JS library build including babel polyfill (minified). 

## Features

- Cross Domain support
- Read SharePoint site properties
- Read SharePoint list properties
- CRUD SharePoint list items [Create/Read/Update/Delete]
- Read SharePoint list default content type
- Read SharePoint users
- Read SharePoint user properties
- Read SharePoint user profile
- Get the current user profile
- Get the current users manager profile

## Installation

### NPM (Webpack Project)

1. Install the module.
```shell
npm i --save-dev sphelper
```
2. Include in your project.
```javascript
import spHelper from 'spHelper';
```
### SharePoint WebPart/Script Editor/Content Editor/Master Page

1. Upload **spHelper-stand-alone-min.js** or  **spHelper-stand-alone-poly.min.js** from the build directory.

2. Include in your project.
```html
<!-- Required MicrosoftAjax.js Script for JSOM functionality -->
<script  type="text/javascript"  src="/mySite/_layouts/15/MicrosoftAjax.js"></script>

<!-- Required SP.Runtime.js & SP.UserProfiles.js Scripts for User Profile functionality -->
<script  type="text/javascript"  src="/mySite/_layouts/15/SP.Runtime.js"></script>
<script  type="text/javascript"  src="/mySite/_layouts/15/SP.UserProfiles.js"></script>

<!-- spHelper -->
<script type="text/javascript" src="spHelper-stand-alone-min.js"></script>
```
> It is necessary to include the Microsoft JS class files prior to spHelper.

### SharePoint Custom App/Add-In

1. Upload **spHelper-stand-alone-min.js** or  **spHelper-stand-alone-poly.min.js** from the build directory.

2. Include in your project.

##### Non-Cross Domain
```html
<!-- spHelper -->
<script type="text/javascript" src="spHelper-stand-alone-min.js"></script>
```
##### Cross Domain

```html
<!--  Required  Init.js  &  MicrosoftAjax.js  Script  for JSOM functionality -->
<script type="text/javascript" src="http://my.other.sharepoint.site.com/mySite/_layouts/15/init.js"></script>
<script type="text/javascript" src="http://my.other.sharepoint.site.com/mySite/_layouts/15/MicrosoftAjax.js"></script>

<!-- Required SP.Runtime.js / SP.core.js / SP.UserProfiles.js / SP.js Script for JSOM functionality -->
<script type="text/javascript" src="http://my.other.sharepoint.site.com/mySite/_layouts/15/SP.Runtime.js"></script>
<script type="text/javascript" src="http://my.other.sharepoint.site.com/mySite/_layouts/15/SP.core.js"></script>
<script type="text/javascript" src="http://my.other.sharepoint.site.com/mySite/_layouts/15/SP.UserProfiles.js"></script>
<script type="text/javascript" src="http://my.other.sharepoint.site.com/mySite/_layouts/15/SP.js"></script>

<!-- Required SP.RequestExecutor.js Script for cross domain execution -->
<script type="text/javascript" src="http://my.other.sharepoint.site.com/mySite/_layouts/15/SP.RequestExecutor.js"></script>

<!-- spHelper -->
<script type="text/javascript" src="spHelper-stand-alone-min.js"></script>
```

## Usage
  
### Initialize spHelper Method: *WebPart/Editor/App/Etc.*

```javascript
// Initialize spHelper.js.
var  dataConnection  =  new  spHelper
({
	crossDomain : false,
	targetSite  : 'http://mysharepoint.site.com/training/'
});

// .... Use the library functions here ....
```

>  **crossDomain**: Set this to false if you do not need cross domain support.
>  **targetSite**: This should be set to the site URL of the SharePoint site you are performing requests on.

### Initialize spHelper Method: *Cross Domain*

>  **Assuming Your Application Site**: http://add-in-7f4164e5346b2f.sharepoint.site.com/mySite/
>  **Assuming SharePoint Request Site**: http://my.other.sharepoint.site.com/mySite/
>  
```javascript
// Initialize spHelper.js.
var  dataConnection  =  new  spHelper
({
	crossDomain  :  true,
	targetSite  :  'http://my.other.sharepoint.site.com/mySite/'
});

// ... Perform spHelper requests here ...
```

>  **crossDomain**: Set this to true for cross domain support.
>  **targetSite**: This should be set to the site URL of the SharePoint site you are performing requests on.  

## spHelper Requests

  

After you have initialized **spHelper.js** you can start to perform requests to your SharePoint site/library.

  

**ASYNC**: It is important to understand that all SharePoint spHelper/JSOM requests are done ASYNC. This means that your request is sent to the server and the rest of your JavaScript code continues to run without receiving the response from the server. When your request is processed by SharePoint, a callback will be run.

  

**CALLBACK**: A JavaScript callback is a function that is ran once a set of commands are completed. Callback functions are typically used when performing ASYNC calls. When you perform a ASYNC request to a server, along with your request itself, you also send a function that will be run after the request is completed.

  
  

## Read a SharePoint Site Property

  

**Function**: getSiteProperty( siteProperty, onSuccess, onFailure )

  

**Description**: This function will ask SharePoint for information (properties) of the initialized target site. There are various types of information you can receive such as Title, Url and more.

  

**Parameters**:

  

-  ***siteProperty [ARRAY]***: An array of SharePoint site properties as [STRING]. These are case sensitive and typically start with a capital letter such as Title or Url.

- Get a full list of available properties here: https://msdn.microsoft.com/en-us/library/office/jj245288.aspx#properties

  

-  **onSuccess [FUNCTION]**: A JavaScript function that will be executed once the request is completed successfully.

  

-  **onFailure [FUNCTION]**: A JavaScript function that will be executed if the request fails.

  

#### EXAMPLE

  

This script will run as soon as the page loads and will read the site Title and URL then print them in the console.

  

```javascript

function  myApplication()

{

// Initialize spHelper.js.

var  dataConnection  =  new  spHelper

({

crossDomain  :  false,

targetSite  :  'http://mysharepoint.site.com/training/'

});

  

// Set the SharePoint site property you want to receive.

var  siteProperty  = ['Title', 'Url'];

  

// Create the callback function that will be executed after the request is completed successfully.

var  onSuccess  =  function (results)

{

console.log('The SharePoint site title is:  '  + results.Title);

console.log('The SharePoint site title is:  '  + results.Url);

}

  

// Create the callback function that will be executed if the request fails.

var  onFailure  =  function (errorMessage)

{

console.log('An error has occured!  '  +  errorMessage);

}

  

// Perform the spHelper request and send your callback functions.

dataConnection.getSiteProperty(siteProperty, onSuccess, onFailure);

}

  

// Run the myApplication function immediately after the client is fully loaded.

SP.SOD.executeFunc("clienttemplates.js", "SPClientTemplates", myApplication);

```

  

## Add a SharePoint List Item

  

**Function**: addListItem( itemDetails, onSuccess, onFailure )

  

**Description**: This function will add a new item in a list with new column data.

  

**Parameters**:

  

-  ***itemDetails [OBJECT]***: An object containing the specific details of the new item. This must be formatted in a specific manner. The structure of this object differs depending on the column field type. Review the code below for the stucture options.

  

-  **listName**: The title of the list.

  

-  **columnData**: A Javascript object with the same name as the column internal name with key/values representing the column data.

-  **Type**: The SharePoint field type of the column.

-  **Value**: The value for the column. See 'Column Data Types' note below.

-  **URL**: The URL value. Used when Type is 'URL'.

-  **Description**: The description of the URL. Used when the Type is 'URL'.

-  **LookupID**: The ID of the item in the lookup table. Used when Type is 'Lookup'.

  

**Column Data Types**:

  

>**IMPORTANT**: The "columnData" object keys are made up of the internalNames of the SharePoint list fields.

>

>**Text**: Use the key "Value" which should be a [STRING].

>

>**Note**: Use the key "Value" which should be a [STRING].

>

>**Integer/Number/Currency**: Use the key "Value" which should be a [INTEGER].

>

>**Choice**: Use the key "Value" which should be a [STRING].

>

>**MultiChoice**: Use the key "Value" which should be an [ARRAY] of [STRING].

>

>**DateTime**: Use the key "Value" which should be a JavaScript [DATE] object.

>

>**Lookup**: Use the key "LookupID" which should be a [INTEGER] representing the lookup ID.

>

>**Boolean**: Use the key "Value" which should be a Javascript [BOOL].

>

>**URL**: Use the key "URL" & "Description" which should both be [STRING].

>

>**User**: Use the key "Value" which should be a [STRING] when the field only accepts one user or [ARRAY] when the field accepts more than one.

  

```javascript

var  itemDetails  =

{

listName  :  'mySharePointList',

columnData  :

{

columnName1  :

{

Type  :  'Text',

Value  :  'This is a string.'

},

columnName2  :

{

Type  :  'Integer',

Value  :  391

},

columnName3  :

{

Type  :  'URL',

URL  :  'http://www.google.com',

Description  :  'This is a link to google.'

},

columnName4  :

{

Type  :  'Lookup',

LookupID  :  10

},

columnName5  :

{

Type  :  'MultiChoice',

Value  : ['Choice #1', 'Choice #2']

},

columnName6  :

{

Type  :  'User',

Value  : ['John Doe', 'Jane Doe']

}

}

};

```

  

-  **onSuccess [FUNCTION]**: A Javascript function that will be executed once the request is completed successfully.

  

-  **onFailure [FUNCTION]**: A Javascript function that will be executed if the request fails.

  

#### EXAMPLE

  

This script will run as soon as the page loads and it will add a new user to the Customers list then print them in the console.

  

```javascript

function  myApplication()

{

// Initialize spHelper.js.

var  dataConnection  =  new  spHelper

({

crossDomain  :  false,

targetSite  :  'http://mysharepoint.site.com/training/'

});

  

// Set the new item details.

var  itemDetails  =

{

listName  :  'Customers',

columnData  :

{

Name  :

{

Type  :  'Text',

Value  :  'John'

},

Last_x0020_Name  :

{

Type  :  'Text',

Value  :  'Doe'

},

Phone  :

{

Type  :  'Text',

Value  :  '555-123-4567'

}

}

};

  

// Create the callback function that will be executed after the request is completed successfully.

var  onSuccess  =  function (results)

{

console.log('New customer added!');

}

  

// Create the callback function that will be executed if the request fails.

var  onFailure  =  function (errorMessage)

{

console.log('An error has occured!  '  +  errorMessage);

}

  

// Perform the spHelper request and send your callback functions.

dataConnection.addListItem(itemDetails, onSuccess, onFailure);

}

  

// Run the myApplication function immediately after the client is fully loaded.

SP.SOD.executeFunc("clienttemplates.js", "SPClientTemplates", myApplication);

```

  

## Update a SharePoint List Item

  

**Function**: updateListItem( itemDetails, onSuccess, onFailure )

  

**Description**: This function will update an existing item in a list with new column data. This uses the same structure as addListItem() in addition to the itemID property.

  

**Parameters**:

  

-  ***itemDetails [OBJECT]***: An object containing the specific details of the item to update.

  

-  **listName**: The title of the list.

  

-  **itemID**: The SharePoint ID of the item you wish to update.

  

-  **columnData**: A JavaScript object with the same name as the column internal name with key/values representing the column data.

  

**Column Data Types**:

  

>See **Add a SharePoint List Item** for more details on columnData.

  

```javascript

var  itemDetails  =

{

listName  :  'mySharePointList',

itemID  :  10

columnData  :

{

columnName1  :

{

Type  :  'Text',

Value  :  'This is a string.'

},

columnName2  :

{

Type  :  'Integer',

Value  :  391

},

columnName3  :

{

Type  :  'URL',

URL  :  'http://www.google.com',

Description  :  'This is a link to google.'

},

columnName4  :

{

Type  :  'Lookup',

LookupID  :  10

},

columnName5  :

{

Type  :  'MultiChoice',

Value  : ['Choice #1', 'Choice #2']

}

columnName6  :

{

Type  :  'User',

Value  : ['John Doe']

}

}

};

```

  

-  **onSuccess [FUNCTION]**: A JavaScript function that will be executed once the request is completed successfully.

  

-  **onFailure [FUNCTION]**: A JavaScript function that will be executed if the request fails.

  
  

#### EXAMPLE

  

This script will run as soon as the page loads and it will add a new user to the Customers list then print them in the console.

  

```javascript

function  myApplication()

{

// Initialize spHelper.js.

var  dataConnection  =  new  spHelper

({

crossDomain  :  false,

targetSite  :  'http://mysharepoint.site.com/training/'

});

  

// Set the new item details.

var  itemDetails  =

{

listName  :  'Customers',

columnData  :

{

Name  :

{

Type  :  'Text',

Value  :  'John'

},

Last_x0020_Name  :

{

Type  :  'Text',

Value  :  'Doe'

},

Phone  :

{

Type  :  'Text',

Value  :  '555-123-4567'

}

}

};

  

// Create the callback function that will be executed after the request is completed successfully.

var  onSuccess  =  function (results)

{

console.log('New customer added!');

}

  

// Create the callback function that will be executed if the request fails.

var  onFailure  =  function (errorMessage)

{

console.log('An error has occured!  '  +  errorMessage);

}

  

// Perform the spHelper request and send your callback functions.

dataConnection.updateListItem(itemDetails, onSuccess, onFailure);

}

  

// Run the myApplication function immediately after the client is fully loaded.

SP.SOD.executeFunc("clienttemplates.js", "SPClientTemplates", myApplication);

```

  

## Read SharePoint List Data

  

Reading data from a SharePoint list using JSOM can be difficult as Microsoft uses CAML Query language to determine what data is received. spHelper.js uses a simplified approach and instead converts a simple JavaScript object into a CAML Query for you.

  

However, if you still want to use CAML statements to read data you may do so.

  

**Function**: getListData( queryDetails, onSuccess, onFailure )

  

**Description**: This function will retrieve items from a SharePoint list that match the query details.

  

**Parameters**:

  

-  ***queryDetails [OBJECT]***: An object containing the specific details of the items you are looking to get from the SharePoint list.

  

-  **listName**: The list title of the SharePoint list.

  

-  **listColumns**: An array of internal column names from the SharePoint list you want to receive back. This key is only optional when a full query is provided.

  

-  **query**: A full CAML string representing the list query. When this is provided, you do not need a 'listColumns' or 'where' key. For more information about CAML visiting the following:

- https://docs.microsoft.com/en-us/previous-versions/office/developer/sharepoint-team-services/dd588092(v=office.11)

- https://msdn.microsoft.com/en-us/library/office/ms426449.aspx

  

-  **where**: A object containing key/values representing the WHERE clause. This is optional.

  

-  **column**: The internal name [STRING] of the column field.

-  **operation**: The CAML operation for the WHERE clause.

-  **value**: The value being compared in the column field.

-  **values**: An array of WHERE objects for multiple column comparisons.

-  **type**: The field type of the column.

  

-  **join**: !!EXPERIMENTAL!! Microsoft limits the types of joins you can make in SharePoint because of the various types of field types. This key is an object that makes up the join details.

  

-  **direction**: The join direction (ie. 'LEFT').

-  **list**: The list you want to join.

-  **joinColumn**: The column which is common between the two tables to join with.

-  **getColumns**: An array of columns in the joined list you want to receive. These columns must also be in the listColumns key of the query object.

  

>  **Note About Joins**: A SharePoint JOIN is limited to the following columns: Calculated, ContentTypeId, Counter, Currency, DateTime, Guid, Integer, Note (One-line only), Text.

  

```javascript

var  queryDetails  =

{

listName  :  'mySharePointList', //REQUIRED

listColumns  : ['columnInternalName1', 'columnInternalName2'], //OPTIONAL

query  :  '<View Scope="RecursiveAll"><Query> ..... </Query><RowLimit>5000</RowLimit></View>'  //OPTIONAL

where  :  //OPTIONAL

{

column  :  'columnInternalName1',

operation  :  'Eq',

value  :  'Test 123',

type  :  'Text',

}

};

```

  

-  **onSuccess [FUNCTION]**: A JavaScript function that will be executed once the request is completed successfully.

  

-  **onFailure [FUNCTION]**: A JavaScript function that will be executed if the request fails.

  
  

#### EXAMPLES

  

Get all items in a list.

  

```javascript

function  myApplication()

{

// Initialize spHelper.js.

var  dataConnection  =  new  spHelper

({

crossDomain  :  false,

targetSite  :  'http://mysharepoint.site.com/training/'

});

  

// Set the query details.

// Look at the list 'Customers' and get the fields 'Name', 'Last_x0020_Name', 'Phone' and 'Created' for ALL items.

let  queryDetails  =

{

listName  :  'Customers',

listColumns  : ['Name', 'Last_x0020_Name', 'Phone', 'Created']

};

  

// Create the callback function that will be executed after the request is completed successfully.

var  onSuccess  =  function (results)

{

results.forEach( function (customer)

{

console.log('First Name:  '  + customer.Name  +  ' Last Name:  '  + customer.Last_x0020_Name  +  ' Phone:  '  + customer.Phone);

console.log('Record created on:  '  + customer.Created);

});

}

  

// Create the callback function that will be executed if the request fails.

var  onFailure  =  function (errorMessage)

{

console.log('An error has occured!  '  +  errorMessage);

}

  

// Perform the spHelper request and send your callback functions.

dataConnection.getListData(queryDetails, onSuccess, onFailure);

}

  

// Run the myApplication function immediately after the client is fully loaded.

SP.SOD.executeFunc("clienttemplates.js", "SPClientTemplates", myApplication);

```

  

Get all items from a list where one field equals a value.

  

```javascript

function  myApplication()

{

// Initialize spHelper.js.

var  dataConnection  =  new  spHelper

({

crossDomain  :  false,

targetSite  :  'http://mysharepoint.site.com/training/'

});

  

// Set the query details.

// Look at the list 'Customers' and get the 'Name', 'Last_x0020_Name', 'Phone' and 'Created' fields WHERE the 'Name' column = 'James'.

var  queryDetails  =

{

listName  :  'Customers',

listColumns  : ['Name', 'Last_x0020_Name', 'Phone', 'Created'],

where  :

{

column  :  'Name',

type  :  'Text',

operation  :  'Eq',

value  :  'James'

}

};

  

// Create the callback function that will be executed after the request is completed successfully.

var  onSuccess  =  function (results)

{

results.forEach( function (customer)

{

console.log('First Name:  '  + customer.Name  +  ' Last Name:  '  + customer.Last_x0020_Name  +  ' Phone:  '  + customer.Phone);

console.log('Record created on:  '  + customer.Created);

});

}

  

// Create the callback function that will be executed if the request fails.

var  onFailure  =  function (errorMessage)

{

console.log('An error has occured!  '  +  errorMessage);

}

  

// Perform the spHelper request and send your callback functions.

dataConnection.getListData(queryDetails, onSuccess, onFailure);

}

  

// Run the myApplication function immediately after the client is fully loaded.

SP.SOD.executeFunc("clienttemplates.js", "SPClientTemplates", myApplication);

```

  

Get all items in a list where multiple columns equal multiple values.

  

```javascript

function  myApplication()

{

// Initialize spHelper.js.

var  dataConnection  =  new  spHelper

({

crossDomain  :  false,

targetSite  :  'http://mysharepoint.site.com/training/'

});

  

// Set the query details.

// Look at the list 'Customers' and get the 'Name', 'Last_x0020_Name', 'Phone' and 'Created' fields

// WHERE the column 'Name' = 'James' AND the column 'Last_x0020_Name' = 'Druhan'.

var  queryDetails  =

{

listName  :  'Customers',

listColumns  : ['Name', 'Last_x0020_Name', 'Phone', 'Created'],

where  :

{

operation  :  'And',

values  :

[

{

column  :  'Name',

type  :  'Text',

operation  :  'Eq',

value  :  'James'

},

{

column  :  'Last_x0020_Name',

type  :  'Text',

operation  :  'Eq',

value  :  'Druhan'

}

]

}

};

  

// Create the callback function that will be executed after the request is completed successfully.

var  onSuccess  =  function (results)

{

results.forEach( function (customer)

{

console.log('First Name:  '  + customer.Name  +  ' Last Name:  '  + customer.Last_x0020_Name  +  ' Phone:  '  + customer.Phone);

console.log('Record created on:  '  + customer.Created);

});

}

  

// Create the callback function that will be executed if the request fails.

var  onFailure  =  function (errorMessage)

{

console.log('An error has occured!  '  +  errorMessage);

}

  

// Perform the spHelper request and send your callback functions.

dataConnection.getListData(queryDetails, onSuccess, onFailure);

}

  

// Run the myApplication function immediately after the client is fully loaded.

SP.SOD.executeFunc("clienttemplates.js", "SPClientTemplates", myApplication);

```

  

Get all items in a list where single column equals multiple values.

  

> The below example may seem strange. This is because the SharePoint CAML capabilities are limited. If you want to receive a list of items where a column matches many different values it is best to use the operation "In". This enables a work around and allows spHelper.js to generate a unique CAML query that will allow 1000's of column comparisons.

  
  

```javascript

function  myApplication()

{

// Initialize spHelper.js.

var  dataConnection  =  new  spHelper

({

crossDomain  :  false,

targetSite  :  'http://mysharepoint.site.com/training/'

});

  

// Create an array of strings that represent first names you want to find.

var  listOfNames  = ['James', 'John', 'Sarah', 'Fred', 'Ginny', 'David', 'Omar']

  

// Set the query details.

// Look at the list 'Customers' and get the 'Name', 'Last_x0020_Name', 'Phone' and 'Created' fields

// WHERE the column 'Name' = each of the names provided in the listOfNames array.

var  queryDetails  =

{

listName  :  'Customers',

listColumns  : ['Name', 'Last_x0020_Name', 'Phone', 'Created'],

where  :

{

column  :  'Name',

operation  :  'In',

value  :  listOfNames,

type  :  'Text'

}

};

  

// Create the callback function that will be executed after the request is completed successfully.

var  onSuccess  =  function (results)

{

results.forEach( function (customer)

{

console.log('First Name:  '  + customer.Name  +  ' Last Name:  '  + customer.Last_x0020_Name  +  ' Phone:  '  + customer.Phone);

console.log('Record created on:  '  + customer.Created);

});

}

  

// Create the callback function that will be executed if the request fails.

var  onFailure  =  function (errorMessage)

{

console.log('An error has occured!  '  +  errorMessage);

}

  

// Perform the spHelper request and send your callback functions.

dataConnection.getListData(queryDetails, onSuccess, onFailure);

}

  

// Run the myApplication function immediately after the client is fully loaded.

SP.SOD.executeFunc("clienttemplates.js", "SPClientTemplates", myApplication);

```

  

## Delete a SharePoint List Item

  

**Function**: deleteListItem( itemDetails, onSuccess, onFailure )

  

**Description**: This function will delete an existing item in a SharePoint list library.

  

**Parameters**:

  

-  ***deleteDetails [OBJECT]***: An object containing the specific details of the item to be deleted.

-  **listName**: The title of the list.

  

-  **itemID**: The SharePoint ID of the item you wish to delete.

  

```javascript

var  deleteDetails  =

{

listName  :  'mySharePointList',

itemID  :  10

};

```

  

-  **onSuccess [FUNCTION]**: A JavaScript function that will be executed once the request is completed successfully.

  

-  **onFailure [FUNCTION]**: A JavaScript function that will be executed if the request fails.

  
  

#### EXAMPLE

  

This script will run as soon as the page loads and it will delete the user.

  

```javascript

function  myApplication()

{

// Initialize spHelper.js.

var  dataConnection  =  new  spHelper

({

crossDomain  :  false,

targetSite  :  'http://mysharepoint.site.com/training/'

});

  

// Set the details of the item you wish to delete.

// Delete the item with an ID of 10 from the SharePoint list 'Customers'.

var  deleteDetails  =

{

listName  :  'Customers',

itemID  :  10

};

  

// Create the callback function that will be executed after the request is completed successfully.

var  onSuccess  =  function (results)

{

console.log('Customer deleted!');

}

  

// Create the callback function that will be executed if the request fails.

var  onFailure  =  function (errorMessage)

{

console.log('An error has occured!  '  +  errorMessage);

}

  

// Perform the spHelper request and send your callback functions.

dataConnection.deleteListItem(deleteDetails, onSuccess, onFailure);

}

  

// Run the myApplication function immediately after the client is fully loaded.

SP.SOD.executeFunc("clienttemplates.js", "SPClientTemplates", myApplication);

```

  

## Get SharePoint List Settings

  

**Function**: getListDetails( libraryName, onSuccess, onFailure )

  

**Description**: This will return a object containing SharePoint list settings of a specific list.

  

**Parameters**:

  

-  **libraryName [STRING]**: The title of the SharePoint library.

  

-  **onSuccess [FUNCTION]**: A JavaScript function that will be executed once the request is completed successfully.

  

-  **onFailure [FUNCTION]**: A JavaScript function that will be executed if the request fails.

  

## Search for SharePoint Users

  

**Function**: searchUsers( searchTerm, onSuccess, onFailure )

  

**Description**: This function will search the SharePoint user database using the preferred name field (first/last).

  

**Parameters**:

  

-  **searchTerm [STRING]**: The users preferred name (first / last) to search for.

  

-  **onSuccess [FUNCTION]**: A JavaScript function that will be executed once the request is completed successfully.

  

-  **onFailure [FUNCTION]**: A JavaScript function that will be executed if the request fails.

  

## Get User Profile Object for a SharePoint User

  

**Function**: getUserProfile( userID, onSuccess, onFailure )

  

**Description**: This function will return the user profile properties object for a specific SharePoint user.

  

**Parameters**:

  

-  **userID [STRING]**: The full qualified user ID ( DOMAIN / ID ).

-

-  **onSuccess [FUNCTION]**: A JavaScript function that will be executed once the request is completed successfully.

  

-  **onFailure [FUNCTION]**: A JavaScript function that will be executed if the request fails.

  

## Get Specific Property of a SharePoint User

  

**Function**: getUserProperty( userProperty, onSuccess, onFailure )

  

**Description**: This function will return a specific user property for a given SharePoint user.

  

**Parameters**:

  

-  **userProperty [ARRAY]**: An array of strings representing the users profile property requested.

  

-  **onSuccess [FUNCTION]**: A JavaScript function that will be executed once the request is completed successfully.

  

-  **onFailure [FUNCTION]**: A JavaScript function that will be executed if the request fails.

  

## Get SharePoint List Default Content Type Name

  

**Function**: getListContentTypeDefault( libraryName, onSuccess, onFailure )

  

**Description**: This function will return the default content type name as a [STRING].

  

**Parameters**: TBD

  

## Get the Current User's User Profile

  

**Function**: getCurrentUser( onSuccess, onFailure )

  

**Description**: This function will return an object of the current users SharePoint user profile.

  

**Parameters**:

  

-  **onSuccess [FUNCTION]**: A JavaScript function that will be executed once the request is completed successfully.

  

-  **onFailure [FUNCTION]**: A JavaScript function that will be executed if the request fails.

  

## Get the Current Users's Manager User Profile

  

**Function**: getCurrentUserManager( onSuccess, onFailure )

  

**Description**: This function will return an object of the current users managar user profile.

  

**Parameters**:

  

-  **onSuccess [FUNCTION]**: A JavaScript function that will be executed once the request is completed successfully.

  

-  **onFailure [FUNCTION]**: A JavaScript function that will be executed if the request fails.