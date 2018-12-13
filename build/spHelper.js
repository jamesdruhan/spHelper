'use strict';

Object.defineProperty(exports, "__esModule", {
    value: true
});

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var spHelper = function () {
    /**
     * Sets default class parameters and initializes the SP Context objects.
     *
     * PARAMETERS
     *    'options' - [OBJECT] : List of class options and option settings.
     *
     * OPTIONS
     *    'crossDomain' - [BOOL]   : Configures the class for cross domain communication.
     *    'targetSite'  - [STRING] : Full URL path to the SharePoint Site to communicate with.
     */
    function spHelper(options) {
        _classCallCheck(this, spHelper);

        // Class member defaults.
        this.crossDomain = false;
        this.targetSite = '';
        this.appContext = {};
        this.spWeb = {};

        // Initialize the class with options.
        this.initializeOptions(options);
    }

    /**
     * Processes the class options.
     *
     * PARAMETERS
     *    'options' - [OBJECT] : See constructor.
     */


    _createClass(spHelper, [{
        key: 'initializeOptions',
        value: function initializeOptions(options) {
            if (typeof options.targetSite !== 'undefined') {
                this.targetSite = options.targetSite;
            }

            // Sets the cross domain flag if the cross domain option is enabled.
            if (typeof options.crossDomain !== 'undefined' && options.crossDomain === true) {
                if (this.targetSite !== '') {
                    this.setCrossDomain();
                } else {
                    throw 'spHelper Error: Configuration error. Cannot set cross domain communication without a target site.';
                }
            }

            // Create the necessary context's and objects.
            this.setClientContext();
            this.setWeb();
        }

        /**
         * Configures the cross domain parameters for the class.
         */

    }, {
        key: 'setCrossDomain',
        value: function setCrossDomain() {
            this.crossDomain = true;
        }

        /**
         * Creates the necessary client context(s) to support the classes various methods.
         */

    }, {
        key: 'setClientContext',
        value: function setClientContext() {
            if (this.targetSite !== '' && this.crossDomain === false) {
                // Create a client context based off of the target URL.
                this.appContext = new SP.ClientContext(this.targetSite);
            } else {
                // Create a client context based off of the current site.
                this.appContext = new SP.ClientContext.get_current();

                if (this.crossDomain === true) {
                    // Creates a client context for the cross domain, using the current domain context.
                    this.crossContext = new SP.AppContextSite(this.appContext, this.targetSite);
                }
            }
        }

        /**
         * Creates the spWeb JSOM object used to support the classes various methods.
         */

    }, {
        key: 'setWeb',
        value: function setWeb() {
            if (this.crossDomain === true) {
                this.spWeb = this.crossContext.get_web();
            } else {
                this.spWeb = this.appContext.get_web();
            }
        }

        /**
         * Refreshes the client context and spWeb objects.
         */

    }, {
        key: 'refreshConnection',
        value: function refreshConnection() {
            this.setClientContext();
            this.setWeb();
        }

        /**
         * Retrieves one or more SPWeb properties of a SharePoint site. If the property is successfully received from the server
         * it will be passed back to the users 'onSuccess' callback function. Similarly, if an error occures, the error message
         * will be passed back to the users 'onFailure' callback function.
         *
         * PARAMETERS
         *      siteProperty  - [ARRAY]                         : An array of [STRING] that represent a spWeb property. Case sensitive.
         *      onSuccessUser - [FUNC ([OBJECT] result)]        : A callback function that is executed when the property is successfully received.
         *      onFailureUser - [FUNC ([STRING] errorMessage)]  : A callback function that is executed when the property cannot be received.
         *
         * OPTIONS
         *      siteProperty : 'Title', 'Url', 'ServerRelativeUrl', 'MasterUrl'
         */

    }, {
        key: 'getSiteProperty',
        value: function getSiteProperty(siteProperty, onSuccessUser, onFailureUser) {
            if (typeof siteProperty == 'undefined') {
                throw 'spHelper Error: Unable to get the site property. Ensure the siteProperty [ARRAY] is correctly configured.';
            }

            try {
                // Refresh connections.
                this.refreshConnection();

                // Setup the request. Limits the returned data to only the requested properties.
                this.appContext.load(this.spWeb, siteProperty);

                // Create a local 'this' for the callback functions.
                var classThis = this;

                // Callback function when the request (promise) has succeeded. This callback is used to collect the requested spWeb properties
                // in a object then return to them to the user with their call back.
                var resolve = function resolve() {
                    var results = {};

                    if (siteProperty.includes('Title')) {
                        results.Title = classThis.spWeb.get_title();
                    }

                    if (siteProperty.includes('ServerRelativeUrl')) {
                        results.ServerRelativeUrl = classThis.spWeb.get_serverRelativeUrl();
                    }

                    if (siteProperty.includes('MasterUrl')) {
                        results.MasterUrl = classThis.spWeb.get_masterUrl();
                    }

                    onSuccessUser(results);
                };

                // Callback function when the request (promise) has rejected.
                var _reject = function _reject(sender, args) {
                    onFailureUser(args.get_message());
                };

                // Execute the request.
                this.appContext.executeQueryAsync(resolve, _reject);
            } catch (error) {
                throw 'spHelper Error: Unable to get the site property. Validate the request details.' + error;
            }
        }

        /**
         * Retreive all data from a SharePoint list library. This method will use the loadListData method which has a item limit of 5000.
         * If the quantity of items returned by the server equals 5000 the loadListData method will be executed again to see if there are
         * more items in the list. This process will repeat until all items have been received. Only then will the data be returned to the
         * users call back.
         *
         * PARAMETERS
         *      queryDetails  - [OBJECT]                        : A key/value pair object with details of the query request.
         *      onSuccessUser - [FUNC ([ARRAY][OBJECT] result)] : A callback function that is executed when the data is successfully received.
         *      onFailureUser - [FUNC ([STRING] errorMessage)]  : A callback function that is executed when the data cannot be received.
         *
         * OPTIONS
         *      queryDetails
         *          'listName'     - [STRING]  : The name/title of the list to query.
         *          'listColumns'  - [ARRAY]   : Array of columns you want to retrieve.
         *          'query'        - [STRING]  : A full CAML query to define which items to retrieve. Leave empty for all items.
         *                                     : "<View Scope='RecursiveAll'><Query> ..... </Query><RowLimit>5000</RowLimit></View>"
         *          'where'        - [OBJECT]  : A object containing key/values detailing the WHERE clause.
         *                                     : "where : { column : 'tableName', operation : 'Eq', value : 'myValue', type : 'Text'}"
         *          'pagePosition' - [INTEGER] : Starting item position to retrieve. Set to 0 for all items.
         */

    }, {
        key: 'getListData',
        value: function getListData(queryDetails, onSuccessUser, onFailureUser) {
            // Set the default page position to 0 (first item) if undefined.
            if (typeof queryDetails.pagePosition == 'undefined') {
                queryDetails.pagePosition = 0;
            }

            // Refresh connections.
            this.refreshConnection();

            // This array will be the final array returned to the users callback once all data is collected.
            var listData = [];

            // Create a local 'this' for the callback functions.
            var classThis = this;

            // This function will run more than once if the 5000 item cap is hit.
            var resolve = function resolve(results) {
                listData = listData.concat(results);

                // There is a default 5000 limit in SharePoint. If reached, run the request again to collect additional items.
                if (results.length == 5000) {
                    // Since the page position starts at 0 it needs to be incremented by the value of 5001 for the first page.
                    if (listData.length == 5000) {
                        queryDetails.pagePosition = queryDetails.pagePosition + 1;
                    }

                    // Increase the page position by 5000.
                    queryDetails.pagePosition = queryDetails.pagePosition + 5000;

                    // Request for more data.
                    classThis.loadListData(queryDetails, resolve, reject);
                } else {
                    // All the data is collected. Run the users callback.
                    onSuccessUser(listData);
                }
            };

            var reject = function reject(errorMessage) {
                onFailureUser(errorMessage);
            };

            // Run the request for data. This initial attempt will cap at 5000 items.
            this.loadListData(queryDetails, resolve, reject);
        }

        /**
         * Retrieve a maximum of 5000 items from a SharePoint list library. Use method 'getListData' to get all list items.
         *
         * PARAMETERS
         *      queryDetails  - [OBJECT]                        : A key/value pair object with details of the query request.
         *      onSuccess     - [FUNC ([ARRAY][OBJECT] result)] : A callback function that is executed when the data is successfully received.
         *      onFailure     - [FUNC ([STRING] errorMessage)]  : A callback function that is executed when the data cannot be received.
         *
         * OPTIONS
         *      queryDetails : See getListData ().
         */

    }, {
        key: 'loadListData',
        value: function loadListData(queryDetails, onSuccess, onFailure) {
            if (typeof queryDetails.listName == 'undefined' && typeof queryDetails.listGuid == 'undefined' || typeof queryDetails.listColumns == 'undefined' || typeof queryDetails.pagePosition == 'undefined') {
                throw 'spHelper Error: Invalid query details. Minimum query details must include list title, columns and page position.';
            }

            try {
                var spList = '';

                if (typeof queryDetails.listName !== 'undefined') {
                    // Will store the spList object when request is complete.
                    spList = this.spWeb.get_lists().getByTitle(queryDetails.listName);
                } else {
                    // Will store the spList object when request is complete.
                    spList = this.spWeb.get_lists().getById(queryDetails.listGuid);
                }

                // Generate the Caml query for the request.
                var camlQuery = new SP.CamlQuery();

                if (typeof queryDetails.query !== 'undefined') {
                    camlQuery.set_viewXml(queryDetails.query);
                }
                // Create Custom Query
                else {
                        var customQuery = "<View Scope='RecursiveAll'>";

                        customQuery += '<ViewFields>';

                        //FieldRef defines which columns are returned.
                        for (var column in queryDetails.listColumns) {
                            customQuery += '<FieldRef Name=\'' + queryDetails.listColumns[column] + '\' />';
                        }

                        customQuery += '</ViewFields>';

                        // Check if a WHERE clause has been provided.
                        if (typeof queryDetails.where !== 'undefined') {
                            var camlWhereInFix = false;
                            var multiWhereValue = false;
                            var batches = 0;

                            // If there is no Where.Value then set it to null (vs. undefined) so it can still be used later.
                            if (typeof queryDetails.where.value === 'undefined') {
                                queryDetails.where.value = '';
                            }

                            customQuery += '<Query> <Where>';

                            // Check if the keyword "values" was used. This is intended for WHERE clauses that use operations like AND/OR
                            // that are comparing columns with their own values/operations. The 'values' property should be an array of
                            // objects [ {column: 'myColumn', operation: 'Eq', value: 'something', type: 'lookup' }, ..... ]
                            if (typeof queryDetails.where.values !== 'undefined') {
                                customQuery += '<' + queryDetails.where.operation + '>';

                                queryDetails.where.values.forEach(function (value) {
                                    customQuery += '<' + value.operation + '>';

                                    // This checks for lookup. If the value is a number then adds the LookupID = 'TRUE' tag. If not,
                                    // any lookups will be compared to the lookup value itself.
                                    if (value.type.toLowerCase() === 'lookup' && !isNaN(value.value)) {
                                        customQuery += '<FieldRef Name=\'' + value.column + '\' LookupId=\'TRUE\'/> <Value Type=\'' + value.type + '\'>' + value.value + '</Value>';
                                    } else {
                                        customQuery += '<FieldRef Name=\'' + value.column + '\'/> <Value Type=\'' + value.type + '\'>' + value.value + '</Value>';
                                    }

                                    customQuery += '</' + value.operation + '>';
                                });

                                customQuery += '</' + queryDetails.where.operation + '>';
                            } else {
                                // Check if more than one value was sent. If true then it is necessary to use <values><value></value></values> later.
                                if (queryDetails.where.value.constructor.name === 'Array') {
                                    multiWhereValue = true;
                                }

                                // Check if the operation is a IN operation.
                                if (queryDetails.where.operation.toLowerCase() === 'in') {
                                    // Check the count of the items. This is necessary because a multi-value WHERE clause will need many <value></value> tags.
                                    // SharePoint only allows a maximum of 60 of these to be used in a single statement. The workaround for over 60 is to
                                    // batch the <value></value> tags in multiple <Or></Or> tags.
                                    if (multiWhereValue && queryDetails.where.value.length > 60) {
                                        // SharePoint CAML <WHERE/IN> Fix Required.
                                        camlWhereInFix = true;

                                        // Calculate number of batches.
                                        batches = Math.ceil(queryDetails.where.value.length / 60);
                                    }
                                }

                                // Create a more complex query by creating batches of <values></values> within a <Or></Or> tag.
                                if (camlWhereInFix) {
                                    customQuery += '<Or>';

                                    for (var i = 0; i < batches; i++) {
                                        customQuery += '<In> <FieldRef Name=\'' + queryDetails.where.column + '\' /> <Values>';

                                        queryDetails.where.value.forEach(function (value) {
                                            customQuery += '<Value Type=\'' + queryDetails.where.type + '\'>' + value + '</Value>';
                                        });

                                        customQuery += '</Values></In>';
                                    }
                                } else {
                                    customQuery += '<' + queryDetails.where.operation + '>';

                                    if (multiWhereValue) {
                                        // This checks for lookup. If the value is a number then adds the LookupID = 'TRUE' tag. If not,
                                        // any lookups will be compared to the lookup value itself.
                                        if (queryDetails.where.type.toLowerCase() === 'lookup' && !isNaN(queryDetails.where.value[0])) {
                                            customQuery += ' <FieldRef Name=\'' + queryDetails.where.column + '\' LookupId=\'TRUE\'/>';
                                        } else {
                                            customQuery += ' <FieldRef Name=\'' + queryDetails.where.column + '\' />';
                                        }

                                        customQuery += '<Values>';

                                        queryDetails.where.value.forEach(function (value) {
                                            customQuery += '<Value Type=\'' + queryDetails.where.type + '\'>' + value + '</Value>';
                                        });

                                        customQuery += '</Values>';
                                    } else {
                                        // This checks for lookup. If the value is a number then adds the LookupID = 'TRUE' tag. If not,
                                        // any lookups will be compared to the lookup value itself.
                                        if (typeof queryDetails.where.type !== 'undefined' && queryDetails.where.type.toLowerCase() === 'lookup' && !isNaN(queryDetails.where.value)) {
                                            customQuery += ' <FieldRef Name=\'' + queryDetails.where.column + '\' LookupId=\'TRUE\'/>';
                                        } else {
                                            customQuery += ' <FieldRef Name=\'' + queryDetails.where.column + '\' />';
                                        }

                                        if (typeof queryDetails.where.type !== 'undefined' && typeof queryDetails.where.value !== 'undefined') {
                                            customQuery += '<Value Type=\'' + queryDetails.where.type + '\'>' + queryDetails.where.value + '</Value>';
                                        }
                                    }

                                    customQuery += '</' + queryDetails.where.operation + '>';
                                }

                                if (camlWhereInFix) {
                                    customQuery += '</Or>';
                                }
                            }

                            customQuery += '</Where></Query>';
                        }

                        if (typeof queryDetails.join !== 'undefined') {
                            customQuery += '\n                    <Joins>\n                        <Join Type="' + queryDetails.join.direction + '" ListAlias="' + queryDetails.join.list + '">\n                            <Eq>\n                                <FieldRef Name="' + queryDetails.join.joinColumn + '" RefType="Id" />\n                                <FieldRef Name="ID" List="' + queryDetails.join.list + '" />\n                            </Eq>\n                        </Join>\n                    </Joins>';

                            customQuery += '<ProjectedFields>';

                            for (var _column in queryDetails.join.getColumns) {
                                customQuery += '<Field ShowField="' + queryDetails.join.getColumns[_column] + '" Type="Lookup" Name="' + queryDetails.join.getColumns[_column] + '" List="' + queryDetails.join.list + '" />';
                            }

                            customQuery += '</ProjectedFields>';
                        }

                        customQuery += '</View>';

                        camlQuery.set_viewXml(customQuery);
                    }

                // Configure the CAML query paging position.
                var position = new SP.ListItemCollectionPosition();

                // Set to initial pull of data to start at the first item (position 0);
                position.set_pagingInfo('Paged=TRUE&p_ID=' + queryDetails.pagePosition);

                camlQuery.set_listItemCollectionPosition(position);

                // Will store the spListItemCollection object when request is complete.
                var spListItemCollection = spList.getItems(camlQuery);

                // Load the request into the client context with only the specific columns.
                this.appContext.load(spListItemCollection, 'Include(' + queryDetails.listColumns.toString() + ')');

                // Create a local 'this' for the callback functions.
                var classThis = this;

                // Callback function when the request (promise) is resolved. Turns the returned item objects into an [ARRAY] of JS [OBJECTS].
                var resolve = function resolve() {
                    var itemsEnumeration = spListItemCollection.getEnumerator();
                    var returnedItems = [];

                    // Cycle through each item returned in the item collection.
                    while (itemsEnumeration.moveNext()) {
                        var currentListItem = itemsEnumeration.get_current();
                        var itemRow = {};

                        // Foreach column requested from the user, create a property in the itemRow object.
                        // This creates a complete object for each row of data.
                        var _iteratorNormalCompletion = true;
                        var _didIteratorError = false;
                        var _iteratorError = undefined;

                        try {
                            for (var _iterator = queryDetails.listColumns[Symbol.iterator](), _step; !(_iteratorNormalCompletion = (_step = _iterator.next()).done); _iteratorNormalCompletion = true) {
                                var columnName = _step.value;

                                // Try & Catch is required because a request for list items will still be successful if the user
                                // requests a column name that doesn't exist. The returned items will simply not contain the unknown
                                // column data. However SPHelper will throw notifying the user.
                                try {
                                    itemRow[columnName] = currentListItem.get_item(columnName);
                                } catch (error) {
                                    if (typeof queryDetails.listName !== 'undefined') {
                                        throw 'spHelper Error: The column \'' + columnName + '\' requested from \'' + queryDetails.listName + '\' does not exist! Error: ' + error;
                                    } else {
                                        throw 'spHelper Error: The column \'' + columnName + '\' requested from \'' + queryDetails.listGuid + '\' does not exist! Error: ' + error;
                                    }
                                }
                            }
                        } catch (err) {
                            _didIteratorError = true;
                            _iteratorError = err;
                        } finally {
                            try {
                                if (!_iteratorNormalCompletion && _iterator.return) {
                                    _iterator.return();
                                }
                            } finally {
                                if (_didIteratorError) {
                                    throw _iteratorError;
                                }
                            }
                        }

                        returnedItems.push(itemRow);
                    }

                    onSuccess(returnedItems);
                };

                // Callback function when the request (promise) has rejected.
                var _reject2 = function _reject2(sender, args) {
                    onFailure(args.get_message());
                };

                this.appContext.executeQueryAsync(resolve, _reject2);
            } catch (error) {
                throw 'spHelper Error: Unable to read list data. Validate query details ... ' + error;
            }
        }

        /**
         * Updates a list item in a SharePoint list library. The update details object must include a number of details including the list name,
         * item ID to update as well as an additional [OBJECT] with the column details to be updated (Column name, Type, Value, etc.). The column
         * details vary depending on the type of SharePoint column. This is because special SharePoint column types like URL or Lookup require
         * additional information.
         *
         * PARAMETERS
         *      updateDetails - [OBJECT]               : A key/value pair object with details of the update request.
         *      onSuccessUser - [FUNC ()]              : A callback function that is executed when the data is successfully updated.
         *      onFailureUser - [FUNC (sender, args)]  : A callback function that is executed when the data cannot be updated.
         *
         * EXAMPLES
         *      var updateDetails =
         *          {
         *              listName      : 'mySharePointList',
         *              itemID        : 52,
         *              columnData :
         *              {
         *                  columnName1 :
         *                  {
         *                      Type  : 'Text',
         *                      Value : 'This is a string.'
         *                  },
         *                  columnName2 :
         *                  {
         *                      Type  : 'Integer',
         *                      Value : 391
         *                  },
         *                  columnName3 :
         *                  {
         *                      Type        : 'URL',
         *                      URL         : 'http://www.google.com',
         *                      Description : 'This is a link to google.'
         *                  },
         *                  columnName4 :
         *                  {
         *                      Type     : 'Lookup',
         *                      LookupID : 10
         *                  },
         *                  columnName5 :
         *                  {
         *                      Type  : 'MultiChoice',
         *                      Value : ['Choice #1', 'Choice #2']
         *                  }
         *              }
         *          };
         */

    }, {
        key: 'updateListItem',
        value: function updateListItem(updateDetails, onSuccessUser, onFailureUser) {
            if (typeof updateDetails.listName == 'undefined' || typeof updateDetails.itemID == 'undefined' || typeof updateDetails.columnData == 'undefined') {
                throw 'spHelper Error: Invalid update details. To update an item, the list name, item ID, and column update data is required.';
            }

            try {
                // Refresh connections.
                this.refreshConnection();

                // Will store the spList object when request is complete.
                var spList = this.spWeb.get_lists().getByTitle(updateDetails.listName);

                // Will store the spItem object when the request is complete.
                var spItem = spList.getItemById(updateDetails.itemID);

                for (var key in updateDetails.columnData) {
                    // Update process for URL fields.
                    if (updateDetails.columnData[key].Type.toLowerCase() == 'url') {
                        try {
                            var spURLField = new SP.FieldUrlValue();

                            spURLField.set_url(updateDetails.columnData[key].URL);
                            spURLField.set_description(updateDetails.columnData[key].Description);

                            spItem.set_item(key, spURLField);
                        } catch (error) {
                            throw 'spHelper Error: Invalid URL field details. Unable to update item ... ' + error;
                        }
                    }
                    // Update process for Lookup fields.
                    else if (updateDetails.columnData[key].Type.toLowerCase() == 'lookup') {
                            try {
                                if (updateDetails.columnData[key].LookupID !== '') {
                                    // If array --> MultiChoice Lookup.
                                    if (updateDetails.columnData[key].LookupID.constructor.name === 'Array') {
                                        (function () {
                                            var lookupIDFields = [];

                                            updateDetails.columnData[key].LookupID.forEach(function (item) {
                                                var spLookupField = new SP.FieldLookupValue();

                                                spLookupField.set_lookupId(item);

                                                lookupIDFields.push(spLookupField);
                                            });

                                            spItem.set_item(key, lookupIDFields);
                                        })();
                                    } else {
                                        var spLookupField = new SP.FieldLookupValue();

                                        spLookupField.set_lookupId(updateDetails.columnData[key].LookupID);

                                        spItem.set_item(key, spLookupField);
                                    }
                                } else {
                                    spItem.set_item(key, null);
                                }
                            } catch (error) {
                                throw 'spHelper Error: Invalid Lookup field details. Unable to update item ... ' + error;
                            }
                        } else if (updateDetails.columnData[key].Type.toLowerCase() === 'user') {
                            (function () {
                                var userList = [];

                                if (updateDetails.columnData[key].Value.constructor.name === 'String') {
                                    var user = new SP.FieldUserValue.fromUser(updateDetails.columnData[key].Value);

                                    userList.push(user);
                                } else {
                                    updateDetails.columnData[key].Value.forEach(function (itemUser) {
                                        var user = new SP.FieldUserValue.fromUser(itemUser);

                                        userList.push(user);
                                    });
                                }

                                spItem.set_item(key, userList);
                            })();
                        }
                        // Update process for normal SharePoint fields (like text, choice, integer).
                        else {
                                spItem.set_item(key, updateDetails.columnData[key].Value);
                            }
                }

                // Apply the changes to the row.
                spItem.update();

                // Update the item.
                this.appContext.executeQueryAsync(onSuccessUser, onFailureUser);
            } catch (error) {
                throw 'spHelper Error: Unable to update item. Validate update details ... ' + error;
            }
        }

        /**
         * Adds a new item to a SharePoint list library. The itemDetails property must contain the list name as well as the column data details.
         *
         * PARAMETERS
         *      itemDetails   - [OBJECT]               : A key/value pair object with details of the create request.
         *      onSuccessUser - [FUNC ()]              : A callback function that is executed when the data is successfully created.
         *      onFailureUser - [FUNC (sender, args)]  : A callback function that is executed when the data cannot be created.
         *
         * EXAMPLES
         *      See updateListItem () method for itemDetails example.
         */

    }, {
        key: 'addListItem',
        value: function addListItem(itemDetails, onSuccessUser, onFailureUser) {
            if (typeof itemDetails.listName == 'undefined' || typeof itemDetails.columnData == 'undefined') {
                throw 'spHelper Error: Invalid create item details. To create an item, listName and columnData must be defined.';
            }

            try {
                // Refresh connections.
                this.refreshConnection();

                // Get the SharePoint list and create a blank list item.
                var spList = this.spWeb.get_lists().getByTitle(itemDetails.listName);
                var itemCreateInfo = new SP.ListItemCreationInformation();
                var spItem = spList.addItem(itemCreateInfo);

                for (var key in itemDetails.columnData) {
                    // Update process for URL fields.
                    if (itemDetails.columnData[key].Type.toLowerCase() === 'url') {
                        try {
                            // Create a blank URL field and set the details.
                            var spURLField = new SP.FieldUrlValue();

                            spURLField.set_url(itemDetails.columnData[key].URL);
                            spURLField.set_description(itemDetails.columnData[key].Description);

                            spItem.set_item(key, spURLField);
                        } catch (error) {
                            throw 'spHelper Error: Invalid URL field details. Unable to add item ... ' + error;
                        }
                    }
                    // Update process for Lookup fields.
                    else if (itemDetails.columnData[key].Type.toLowerCase() === 'lookup') {
                            try {
                                // For single lookup fields.
                                if (itemDetails.columnData[key].LookupID.constructor.name === 'Number' || itemDetails.columnData[key].LookupID.constructor.name === 'String') {
                                    // Create a blank Lookup field and set the details.
                                    var spLookupField = new SP.FieldLookupValue();

                                    spLookupField.set_lookupId(itemDetails.columnData[key].LookupID);

                                    spItem.set_item(key, spLookupField);
                                }
                                // For multiple choice lookup fields.
                                else {
                                        (function () {
                                            var spLookupFields = [];

                                            itemDetails.columnData[key].LookupID.forEach(function (iD) {
                                                // Create a blank Lookup field and set the details.
                                                var spLookupField = new SP.FieldLookupValue();

                                                spLookupField.set_lookupId(iD);

                                                spLookupFields.push(spLookupField);
                                            });

                                            spItem.set_item(key, spLookupFields);
                                        })();
                                    }
                            } catch (error) {
                                throw 'spHelper Error: Invalid Lookup field details. Unable to add item ... ' + error;
                            }
                        } else if (itemDetails.columnData[key].Type.toLowerCase() === 'user') {
                            (function () {
                                var userList = [];

                                if (itemDetails.columnData[key].Value.constructor.name === 'String') {
                                    var user = new SP.FieldUserValue.fromUser(itemDetails.columnData[key].Value);

                                    userList.push(user);
                                } else {
                                    itemDetails.columnData[key].Value.forEach(function (itemUser) {
                                        var user = new SP.FieldUserValue.fromUser(itemUser);

                                        userList.push(user);
                                    });
                                }

                                spItem.set_item(key, userList);
                            })();
                        }
                        // Update process for normal SharePoint fields (like text, choice, integer).
                        else {
                                spItem.set_item(key, itemDetails.columnData[key].Value);
                            }
                }

                var onLocalSuccess = function onLocalSuccess() {
                    onSuccessUser(spItem.get_id());
                };

                // Apply the changes to the blank item.
                spItem.update();

                // Add the new item.
                this.appContext.executeQueryAsync(onLocalSuccess, onFailureUser);
            } catch (error) {
                throw 'SPHelper Error: Unable to add new item. Validate create details ... ' + error;
            }
        }

        /**
         * NOT WORKING: Adds attachment to a list item.
         *
         * PARAMETERS
         *      itemDetails   - [OBJECT]               : A key/value pair object with details of the create request.
         *      onSuccessUser - [FUNC ()]              : A callback function that is executed when the data is successfully created.
         *      onFailureUser - [FUNC (sender, args)]  : A callback function that is executed when the data cannot be created.
         */

    }, {
        key: 'addListItemAttachment',
        value: function addListItemAttachment(fileDetails, onSuccessUser, onFailureUser) {
            // Refresh connections.
            this.refreshConnection();

            // Will store the spList object when request is complete.
            var spList = this.spWeb.get_lists().getByTitle(fileDetails.listName);

            this.appContext.load(spList, 'RootFolder');

            // Will store the spListItem object when request is complete.
            var spItem = spList.getItemById(fileDetails.itemID);

            this.appContext.load(spItem);

            var vueThis = this;

            var onLoadDetailsSuccess = function onLoadDetailsSuccess() {
                // Check if attachments already exist for this item.
                if (!spItem.get_fieldValues()['Attachments']) {
                    var attachmentRootFolderURL = spList.get_rootFolder().get_serverRelativeUrl() + '/Attachments';

                    var attachmentsRootFolder = vueThis.spWeb.getFolderByServerRelativeUrl(attachmentRootFolderURL);

                    // This gets access denied.
                    var attachmentsFolder = attachmentsRootFolder.get_folders().add('_' + fileDetails.itemID);

                    // This gets moveTo function not found.
                    attachmentsFolder.moveTo(attachmentRootFolderURL + '/' + fileDetails.itemID);

                    vueThis.appContext.load(attachmentsFolder);

                    vueThis.appContext.executeQueryAsync(function (result) {
                        console.log(result);
                    }, function (sender, args) {
                        console.log(args.get_message());
                    });
                }
                // If so, we don't need to create the attachment folder.
                else {}
            };

            var onLoadDetailsFailure = function onLoadDetailsFailure(sender, args) {
                onFailureUser(args.get_message());
            };

            this.appContext.executeQueryAsync(onLoadDetailsSuccess, onLoadDetailsFailure);
        }

        /**
         * Delete an item from a SharePoint list library.
         *
         * PARAMETERS
         *      itemDetails   - [OBJECT]               : A key/value pair object with details of the delete request.
         *      onSuccessUser - [FUNC ()]              : A callback function that is executed when the data is successfully deleted.
         *      onFailureUser - [FUNC (sender, args)]  : A callback function that is executed when the data cannot be deleted.
         *
         * OPTIONS
         *      deleteDetails
         *          'listName'- [STRING]  : SharePoint list to delete from.
         *          'itemID'  - [INTEGER] : Item ID to delete.
         */

    }, {
        key: 'deleteListItem',
        value: function deleteListItem(deleteDetails, onSuccessUser, onFailureUser) {
            if (typeof deleteDetails.listName == 'undefined' || typeof deleteDetails.itemID == 'undefined') {
                throw 'spHelper Error: Invalid delete details. To update an item, the list name and item ID is required.';
            }

            try {
                // Refresh connections.
                this.refreshConnection();

                // Will store the spList object when request is complete.
                var spList = this.spWeb.get_lists().getByTitle(deleteDetails.listName);

                // Will store the spItem object when the request is complete.
                var spItem = spList.getItemById(deleteDetails.itemID);

                spItem.deleteObject();

                this.appContext.executeQueryAsync(onSuccessUser, onFailureUser);
            } catch (error) {
                throw 'spHelper Error: Unable to delete item. Validate update details ... ' + error;
            }
        }

        /**
        * Gets the default content type for a specific library/list.
        *
        * PARAMETERS
        *      libraryName   - [STRING]                       : The name of the library.
        *      onSuccessUser - [FUNC ([STRING] result)]       : A callback function that is executed when the property is successfully received.
        *      onFailureUser - [FUNC ([STRING] errorMessage)] : A callback function that is executed when the property cannot be received.
        */

    }, {
        key: 'getListContentTypeDefault',
        value: function getListContentTypeDefault(libraryName, onSuccessUser, onFailureUser) {
            try {
                // Refresh connections.
                this.refreshConnection();

                // Will store the spList object when request is complete.
                var spList = this.spWeb.get_lists().getByTitle(libraryName);

                // Will store the spContentTypeCollection object when request is complete.
                var spContentTypeCollection = spList.get_contentTypes();

                // Load the request into the client context.
                this.appContext.load(spContentTypeCollection);

                // Create a local 'this' for the callback functions.
                var classThis = this;

                // Callback function when the request (promise) is resolved.
                var resolve = function resolve() {
                    // Get the enumberator for the content types returned.
                    var contentTypeEnumerator = spContentTypeCollection.getEnumerator();

                    // Move to the first item (this is the default content type).
                    contentTypeEnumerator.moveNext();

                    // Return the default content type ID.
                    onSuccessUser(contentTypeEnumerator.get_current().get_id().toString());
                };

                // Callback function when the request (promise) has rejected.
                var _reject3 = function _reject3(sender, args) {
                    onFailureUser(args.get_message());
                };

                this.appContext.executeQueryAsync(resolve, _reject3);
            } catch (error) {
                throw 'spHelper Error: Unable to get list default content type ... ' + error;
            }
        }

        /**
         * Gets a full breakdown of a list library including column details and settings. The 'readOnlyFields' option allows the user
         * to control the quantity of columns returned. Setting this option to 'true' means only fields that can be modified in a form
         * are turned. Setting to 'false' will ensure any and every field is returned (such as 'Created', 'Author', etc.).
         *
         * PARAMETERS
         *      libraryName    - [STRING]                       : The name of the library.
         *      onSuccessUser  - [FUNC ([STRING] result)]       : A callback function that is executed when the property is successfully received.
         *      onFailureUser  - [FUNC ([STRING] errorMessage)] : A callback function that is executed when the property cannot be received.
         *      readOnlyFields - [BOOL]                         : Inidicates if only the read only fields of a list should be returned.
         */

    }, {
        key: 'getListDetails',
        value: function getListDetails(libraryName, onSuccessUser, onFailureUser) {
            var _this = this;

            var readOnlyFields = arguments.length > 3 && arguments[3] !== undefined ? arguments[3] : true;

            try {
                // Refresh connections.
                this.refreshConnection();

                var getLibraryDetails = function getLibraryDetails(contentTypeID) {
                    // This will hold all the list details.
                    var listDetails = {};

                    // Will store the spList object when request is complete.
                    var spList = _this.spWeb.get_lists().getByTitle(libraryName);

                    // Will store the spContentTypeCollection object when request is complete.
                    var spContentTypeCollection = spList.get_contentTypes();

                    // Will store the spContentType object when request is complete.
                    var spContentType = spContentTypeCollection.getById(contentTypeID);

                    // Will store the spFields object when request is complete.
                    var spFields = spContentType.get_fields();

                    // Get the root folder of the list. This is used to generate the list URL.
                    var spFolder = spList.get_rootFolder();

                    // Create a local spWeb object -- This resolves scoping issues.
                    var spWeb = _this.spWeb;

                    // Load the request into the client context.
                    _this.appContext.load(spFields);
                    _this.appContext.load(spList);
                    _this.appContext.load(spFolder);
                    _this.appContext.load(spWeb);

                    // Callback function when the request (promise) is resolved.
                    var resolve = function resolve() {
                        listDetails['settings'] = {};

                        // Get some library settings.
                        listDetails['settings']['id'] = spList.get_id().toString();
                        listDetails['settings']['title'] = spList.get_title();
                        listDetails['settings']['enableAttachments'] = spList.get_enableAttachments();
                        listDetails['settings']['contentTypesEnabled'] = spList.get_contentTypesEnabled();
                        listDetails['settings']['description'] = spList.get_description();
                        listDetails['settings']['enableFolderCreation'] = spList.get_enableFolderCreation();
                        listDetails['settings']['enableMinorVersions'] = spList.get_enableMinorVersions();
                        listDetails['settings']['enableModeration'] = spList.get_enableModeration();
                        listDetails['settings']['enableVersioning'] = spList.get_enableVersioning();
                        listDetails['settings']['forceCheckout'] = spList.get_forceCheckout();
                        listDetails['settings']['parentWebUrl'] = spList.get_parentWebUrl();
                        listDetails['settings']['template'] = spList.get_baseTemplate();
                        listDetails['settings']['rootFolder'] = spFolder.get_serverRelativeUrl();
                        listDetails['settings']['internalName'] = spFolder.get_name();

                        // Set the library relative server URL. This depends on the type of library (list/document).
                        if (listDetails.settings.template == 100) {
                            listDetails['settings']['serverRelativeURL'] = spWeb.get_url() + '/Lists/' + listDetails.settings.internalName + '/';
                        } else {
                            listDetails['settings']['serverRelativeURL'] = spWeb.get_url() + '/' + listDetails.settings.internalName + '/';
                        }

                        // Get all the library columns and details.
                        var spFieldsEnumerator = spFields.getEnumerator();

                        listDetails['columns'] = [];

                        while (spFieldsEnumerator.moveNext()) {
                            var tempColumn = {};
                            var currentField = spFieldsEnumerator.get_current();

                            if (currentField.get_internalName() !== 'ContentType') {
                                // Check if read only fields should be provided. If so, validate the field and skip or not.
                                if (readOnlyFields && currentField.get_readOnlyField() === true) {
                                    continue;
                                }

                                // Get general column details.
                                tempColumn = {
                                    id: currentField.get_id().toString(),
                                    title: currentField.get_title(),
                                    internalName: currentField.get_internalName(),
                                    default: currentField.get_defaultValue(),
                                    unique: currentField.get_enforceUniqueValues(),
                                    required: currentField.get_required(),
                                    hidden: currentField.get_hidden(),
                                    description: currentField.get_description(),
                                    fieldType: currentField.get_fieldTypeKind()
                                };

                                // Dig deeper for more details (some properties are not exposed to JSOM so we need to extract them from the schemaXML).
                                var parser = new DOMParser();
                                var fieldXML = currentField.get_schemaXml();
                                var parsedXML = parser.parseFromString(fieldXML, 'text/xml');
                                var fieldAttributes = parsedXML.getElementsByTagName("Field")[0].attributes;

                                // Get additional column detail: String
                                if (currentField.get_fieldTypeKind() === 2) {
                                    tempColumn['maxLength'] = currentField.get_maxLength();
                                }

                                // Get additional column detail: Choice
                                if (currentField.get_fieldTypeKind() === 6 || currentField.get_fieldTypeKind() === 15) {
                                    tempColumn['choices'] = currentField.get_choices();
                                    tempColumn['fillInChoice'] = currentField.get_fillInChoice();

                                    if (currentField.get_fieldTypeKind() === 6) {
                                        tempColumn['editFormat'] = currentField.get_editFormat();
                                    }
                                }

                                // Get additional column detail: Multiple Lines
                                if (currentField.get_fieldTypeKind() === 3) {
                                    tempColumn['numberOfLines'] = currentField.get_numberOfLines();
                                    tempColumn['richText'] = currentField.get_richText();
                                    tempColumn['appendOnly'] = currentField.get_appendOnly();
                                }

                                // Get additional column detail: Number & Currency
                                if (currentField.get_fieldTypeKind() === 9 || currentField.get_fieldTypeKind() === 10) {
                                    tempColumn['minimumValue'] = currentField.get_minimumValue();
                                    tempColumn['maximumValue'] = currentField.get_maximumValue();

                                    // Specific to numbers.
                                    if (currentField.get_fieldTypeKind() === 9) {
                                        if (typeof fieldAttributes.Percentage !== 'undefined') {
                                            tempColumn['showAsPercentage'] = fieldAttributes.Percentage.value === "FALSE" ? false : true;
                                        } else {
                                            tempColumn['showAsPercentage'] = false;
                                        }

                                        // Check if a decimal setting was set.
                                        if (typeof fieldAttributes.Decimals !== 'undefined') {
                                            tempColumn['displayFormat'] = parseInt(fieldAttributes.Decimals.value);
                                        }
                                    }

                                    // Specific to currency.
                                    if (currentField.get_fieldTypeKind() === 10) {
                                        tempColumn['currencyLocaleId'] = currentField.get_currencyLocaleId();
                                    }
                                }

                                // Get additional column detail: Date Time
                                if (currentField.get_fieldTypeKind() === 4) {
                                    tempColumn['displayFormat'] = currentField.get_displayFormat();
                                    tempColumn['friendlyDisplayFormat'] = currentField.get_friendlyDisplayFormat();
                                }

                                // Get additional column detail: Lookup
                                if (currentField.get_fieldTypeKind() === 7) {
                                    tempColumn['allowMultipleValues'] = currentField.get_allowMultipleValues();
                                    tempColumn['lookupList'] = currentField.get_lookupList();
                                    tempColumn['lookupField'] = currentField.get_lookupField();
                                }

                                // Get additional column detail: User
                                if (currentField.get_fieldTypeKind() === 20) {
                                    tempColumn['allowMultipleValues'] = currentField.get_allowMultipleValues();
                                }

                                // Get additional column detail: URL/Picture
                                if (currentField.get_fieldTypeKind() === 11) {
                                    tempColumn['displayFormat'] = currentField.get_displayFormat();
                                }

                                // Get additional column detail: Calculated
                                if (currentField.get_fieldTypeKind() === 17) {
                                    tempColumn['resultType'] = fieldAttributes.ResultType.value;

                                    if (fieldAttributes.ResultType.value === 'Number') {
                                        tempColumn['displayFormat'] = fieldAttributes.Decimals.value;
                                        tempColumn['showAsPercentage'] = fieldAttributes.Percentage.value;
                                    }

                                    if (fieldAttributes.ResultType.value === 'Currency') {
                                        tempColumn['displayFormat'] = fieldAttributes.Decimals.value;
                                        tempColumn['currencyLocaleId'] = fieldAttributes.LCID.value;
                                    }
                                }

                                // Create a list of fields.
                                listDetails['columns'].push(tempColumn);
                            }
                        }

                        onSuccessUser(listDetails);
                    };

                    // Callback function when the request (promise) has rejected.
                    var reject = function reject(sender, args) {
                        _onFailure(args.get_message());
                    };

                    _this.appContext.executeQueryAsync(resolve, reject);
                };

                var _onFailure = function _onFailure(errorMessage) {
                    onFailureUser(errorMessage);
                };

                // Get the content type ID (default) then get the columns.
                this.getListContentTypeDefault(libraryName, getLibraryDetails, _onFailure);
            } catch (error) {
                throw 'spHelper Error: Unable to get list details ... ' + error;
            }
        }

        /**
         * Gets a specific property for the current user.
         *
         * PARAMETERS
         *      siteProperty  - [ARRAY]                         : An array of [STRING] that represent a spWeb property. Case sensitive.
         *      onSuccessUser - [FUNC ([OBJECT] result)]        : A callback function that is executed when the property is successfully received.
         *      onFailureUser - [FUNC ([STRING] errorMessage)]  : A callback function that is executed when the property cannot be received.
         *
         * OPTIONS
         *      siteProperty : 'Title', 'Url', 'ServerRelativeUrl', 'MasterUrl'
         */

    }, {
        key: 'getUserProperty',
        value: function getUserProperty(userProperty, onSuccessUser, onFailureUser) {
            if (typeof userProperty == 'undefined') {
                throw 'spHelper Error: Unable to get the user property. Ensure the userProperty [ARRAY] is correctly configured.';
            }

            try {
                // Refresh connections.
                this.refreshConnection();

                // This variable will hold the SPUser class after the query is executed.
                var currentUser = this.spWeb.get_currentUser();

                // Setup the request. Limits the returned data to only the requested properties.
                this.appContext.load(currentUser, userProperty);

                // Callback function when the request (promise) has succeeded. This callback is used to collect the requested spWeb properties
                // in a object then return to them to the user with their call back.
                var resolve = function resolve() {
                    var results = {};

                    if (userProperty.includes('Title')) {
                        results.Title = currentUser.get_title();
                    }

                    if (userProperty.includes('id')) {
                        results.ID = currentUser.get_id();
                    }

                    if (userProperty.includes('email')) {
                        results.Email = currentUser.get_email();
                    }

                    onSuccessUser(results);
                };

                // Callback function when the request (promise) has rejected.
                var _reject4 = function _reject4(sender, args) {
                    onFailureUser(args.get_message());
                };

                // Execute the request.
                this.appContext.executeQueryAsync(resolve, _reject4);
            } catch (error) {
                throw 'spHelper Error: Unable to get the user property. Validate the request details.' + error;
            }
        }

        /**
         * Search for a SharePoint user via their preferred name (first/last).
         *
         * PARAMETERS
         *      searchTerm    - [STRING]                        : A [STRING] of the users preferred name.
         *      onSuccessUser - [FUNC ([OBJECT] result)]        : A callback function that is executed when the property is successfully received.
         *      onFailureUser - [FUNC ([STRING] errorMessage)]  : A callback function that is executed when the property cannot be received.
         *
         * OPTIONS
         *      siteProperty : 'Title', 'Url', 'ServerRelativeUrl', 'MasterUrl'
         */

    }, {
        key: 'searchUsers',
        value: function searchUsers(searchTerm, onSuccessUser, onFailureUser) {
            // Refresh connections.
            this.refreshConnection();

            // Create query object for searching people.
            // Options here: https://msdn.microsoft.com/en-us/library/microsoft.sharepoint.applicationpages.clientpickerquery.clientpeoplepickerqueryparameters_members.aspx
            var query = new SP.UI.ApplicationPages.ClientPeoplePickerQueryParameters();

            // Configure the query.
            query.set_allowMultipleEntities(false);
            query.set_maximumEntitySuggestions(50);
            query.set_principalType(1);
            query.set_principalSource(15);
            query.set_queryString(searchTerm);

            // Load the search query in the context.
            var searchResult = SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser(this.appContext, query);

            // Create a local this for the resolve function.
            var localThis = this;

            // Function that runs on success.
            var resolve = function resolve() {
                // Parse the results to a JSON object.
                var results = JSON.parse(searchResult.get_value());

                if (results) {
                    var userList = [];
                    var profilePromises = [];

                    results.forEach(function (item) {
                        // Create a list of promises to get the users profile.
                        profilePromises.push(new Promise(function (resolve, reject) {
                            var buildUserList = function buildUserList(userProfile) {
                                if (userProfile !== null) {
                                    userList.push(userProfile);
                                }

                                resolve();
                            };

                            // Error getting profile.
                            var displayError = function displayError(error) {
                                onFailureUser(error);
                            };

                            // Make the SharePoint request.
                            localThis.getUserProfile(item.Description, buildUserList, displayError);
                        }));
                    });

                    // Complete all the queued up promises then return the results.
                    Promise.all(profilePromises).then(function () {
                        onSuccessUser(userList);
                    });
                } else {
                    onSuccessUser(null);
                }
            };

            var reject = function reject(sender, args) {
                onFailureUser(args.get_message());
            };

            // Execute the query.
            this.appContext.executeQueryAsync(resolve, reject);
        }

        /**
         * Gets the user profile for a given user.
         *
         * PARAMETERS
         *      userID        - [STRING]                        : A string representing the users ID (DOMAIN\USERNAME)
         *      onSuccessUser - [FUNC ([OBJECT] result)]        : A callback function that is executed when the property is successfully received.
         *      onFailureUser - [FUNC ([STRING] errorMessage)]  : A callback function that is executed when the property cannot be received.
         */

    }, {
        key: 'getUserProfile',
        value: function getUserProfile(userID, onSuccessUser, onFailureUser) {
            // Refresh connections.
            this.refreshConnection();

            // Create a SP.UserProfile.PeopleManager object.
            var peopleManager = new SP.UserProfiles.PeopleManager(this.appContext);

            // This will be the variable loaded with the properties after execute.
            var personProperties = peopleManager.getPropertiesFor(userID);

            // Load the query to execute in the context.
            this.appContext.load(personProperties);

            // Execute the query and fill the variable (personProperties).
            this.appContext.executeQueryAsync(onRequestSuccess, onRequestFail);

            // Function run on execute success.
            function onRequestSuccess() {
                // Try and get the userProfileProperties. If this fails, the user isn't a valid SP user with properties.
                try {
                    // Send the properties back.
                    onSuccessUser(personProperties.get_userProfileProperties());
                } catch (error) {
                    onSuccessUser(null);
                }
            }

            // This function runs if the executeQueryAsync call fails.
            function onRequestFail(sender, args) {
                onFailureUser(args.get_message());
            }
        }

        /**
         * Gets the user profile for the current user.
         *
         * PARAMETERS
         *      onSuccessUser - [FUNC ([OBJECT] result)]        : A callback function that is executed when the property is successfully received.
         *      onFailureUser - [FUNC ([STRING] errorMessage)]  : A callback function that is executed when the property cannot be received.
         */

    }, {
        key: 'getCurrentUser',
        value: function getCurrentUser(onSuccessUser, onFailureUser) {
            // Refresh connections.
            this.refreshConnection();

            // Create a SP.UserProfile.PeopleManager object.
            var peopleManager = new SP.UserProfiles.PeopleManager(this.appContext);

            // This will be the variable loaded with the properties after execute.
            var personProperties = peopleManager.getMyProperties();

            // Load the query to execute in the context.
            this.appContext.load(personProperties);

            // Execute the query and fill the variable (personProperties).
            this.appContext.executeQueryAsync(onRequestSuccess, onRequestFail);

            // Function run on execute success.
            function onRequestSuccess() {
                // Try and get the userProfileProperties. If this fails, the user isn't a valid SP user with properties.
                try {
                    // Send the properties back.
                    onSuccessUser(personProperties.get_userProfileProperties());
                } catch (error) {
                    onSuccessUser(null);
                }
            }

            // This function runs if the executeQueryAsync call fails.
            function onRequestFail(sender, args) {
                onFailureUser(args.get_message());
            }
        }

        /**
         * Gets the user profile for the current users manager.
         *
         * PARAMETERS
         *      onSuccessUser - [FUNC ([OBJECT] result)]        : A callback function that is executed when the property is successfully received.
         *      onFailureUser - [FUNC ([STRING] errorMessage)]  : A callback function that is executed when the property cannot be received.
         */

    }, {
        key: 'getCurrentUserManager',
        value: function getCurrentUserManager(onSuccessUser, onFailureUser) {
            // Refresh connections.
            this.refreshConnection();

            var vueThis = this;

            var onSuccessCurrent = function onSuccessCurrent(result) {
                vueThis.getUserProfile(result['Manager'], onSuccessUser, onFailureUser);
            };

            var onFailureCurrent = function onFailureCurrent(error) {
                onFailureUser(error);
            };

            this.getCurrentUser(onSuccessCurrent, onFailureCurrent);
        }

        /**
         * Gets the number (constant) equivalent to the text based SP.FieldType if the passed variable is a STRING.
         * Gets the text (string) equivalent of the SP.FieldType if the passed variable is a NUMBER.
         *
         * PARAMETERS
         *      type - [STRING/NUMBER] : A string representing a SP.FieldType or a number (SP.FieldType).
         */

    }, {
        key: 'fieldType',
        value: function fieldType(type) {
            if (type.constructor.name === 'String') {
                return SP.FieldType[type];
            } else {
                if (type === 1) {
                    return 'integer';
                } else if (type === 2) {
                    return 'text';
                } else if (type === 3) {
                    return 'note';
                } else if (type === 4) {
                    return 'dateTime';
                } else if (type === 5) {
                    return 'counter';
                } else if (type === 6) {
                    return 'choice';
                } else if (type === 7) {
                    return 'lookup';
                } else if (type === 8) {
                    return 'boolean';
                } else if (type === 9) {
                    return 'number';
                } else if (type === 10) {
                    return 'currency';
                } else if (type === 11) {
                    return 'URL';
                } else if (type === 12) {
                    return 'computed';
                } else if (type === 13) {
                    return 'threading';
                } else if (type === 14) {
                    return 'guid';
                } else if (type === 15) {
                    return 'multiChoice';
                } else if (type === 16) {
                    return 'gridChoice';
                } else if (type === 17) {
                    return 'calculated';
                } else if (type === 18) {
                    return 'file';
                } else if (type === 19) {
                    return 'attachments';
                } else if (type === 20) {
                    return 'user';
                } else if (type === 21) {
                    return 'recurrence';
                } else if (type === 22) {
                    return 'crossProjectLink';
                } else if (type === 23) {
                    return 'modStat';
                } else if (type === 24) {
                    return 'error';
                } else if (type === 25) {
                    return 'contentTypeId';
                } else if (type === 26) {
                    return 'pageSeparator';
                } else if (type === 27) {
                    return 'threadIndex';
                } else if (type === 28) {
                    return 'workflowStatus';
                } else if (type === 29) {
                    return 'allDayEvent';
                } else if (type === 30) {
                    return 'workflowEventType';
                } else if (type === 31) {
                    return 'maxItems';
                }
            }
        }

        /**
         * Gets the number (constant) equivalent to the text based SP.DateTimeFieldFormatType.
         *
         * PARAMETERS
         *      type - [STRING] : A string representing a SP.DateTimeFieldFormatType.
         */

    }, {
        key: 'dateType',
        value: function dateType(type) {
            return SP.DateTimeFieldFormatType[type];
        }

        /**
         * Gets the number (constant) equivalent to the text based SP.Nummber.FormatType.
         *
         * PARAMETERS
         *      type - [STRING] : A string representing a SP.Nummber.FormatType.
         */

    }, {
        key: 'numberFormatType',
        value: function numberFormatType(type) {
            if (type.toLowerCase() === 'automatic') {
                return 'undefined';
            } else if (type.toLowerCase() === 'nodecimal') {
                return 0;
            } else if (type.toLowerCase() === 'onedecimal') {
                return 1;
            } else if (type.toLowerCase() === 'twodecimals') {
                return 2;
            } else if (type.toLowerCase() === 'threedecimals') {
                return 3;
            } else if (type.toLowerCase() === 'fourdecimals') {
                return 4;
            } else if (type.toLowerCase() === 'fivedecimals') {
                return 5;
            }
        }

        /**
        * Verifies if user has read access to the specified list.
        *
        * PARAMETERS
        *      libraryName - [STRING] : A string representing the list name.
        */

    }, {
        key: 'canReadList',
        value: function canReadList(libraryName, onHasAccess, onNoAccess) {
            if (typeof libraryName == 'undefined') {
                throw 'spHelper Error: Invalid can read list details. To check access a libraryName must be defined.';
            }

            try {
                // Refresh connections.
                this.refreshConnection();

                // Will store the spList object when request is complete.
                var spList = this.spWeb.get_lists().getByTitle(libraryName);

                // Load the SPList object.
                this.appContext.load(spList);

                // Function that runs on success (has access).
                var resolve = function resolve() {
                    onHasAccess();
                };

                // Function to run when no access.
                var _reject5 = function _reject5(sender, args) {
                    onNoAccess();
                };

                this.appContext.executeQueryAsync(resolve, _reject5);
            } catch (error) {
                throw 'spHelper Error: Unable to validate user permissions for list ... ' + error;
            }
        }

        /**
        * Verifies if user has write/edit access to the specified list.
        *
        * PARAMETERS
        *      libraryName - [STRING] : A string representing the list name.
        */

    }, {
        key: 'canWriteList',
        value: function canWriteList(libraryName, onHasAccess, onNoAccess) {
            if (typeof libraryName == 'undefined') {
                throw 'spHelper Error: Invalid can write list details. To check access a libraryName must be defined.';
            }

            try {
                // Refresh connections.
                this.refreshConnection();

                // Will store the spList object when request is complete.
                var spList = this.spWeb.get_lists().getByTitle(libraryName);

                // Load the SPList object with its EffectiveBasePermissions property.
                this.appContext.load(spList, 'EffectiveBasePermissions');

                // Function that runs on success (has access).
                var resolve = function resolve() {
                    if (spList.get_effectiveBasePermissions().has(SP.PermissionKind.editListItems)) {
                        onHasAccess();
                    } else {
                        onNoAccess();
                    }
                };

                // Function to run when no access.
                var _reject6 = function _reject6(sender, args) {
                    onNoAccess();
                };

                this.appContext.executeQueryAsync(resolve, _reject6);
            } catch (error) {
                throw 'spHelper Error: Unable to validate user permissions for list ... ' + error;
            }
        }

        /**
        * Verifies if user has delete access to the specified list.
        *
        * PARAMETERS
        *      libraryName - [STRING] : A string representing the list name.
        */

    }, {
        key: 'canDeleteList',
        value: function canDeleteList(libraryName, onHasAccess, onNoAccess) {
            if (typeof libraryName == 'undefined') {
                throw 'spHelper Error: Invalid can delete list details. To check access a libraryName must be defined.';
            }

            try {
                // Refresh connections.
                this.refreshConnection();

                // Will store the spList object when request is complete.
                var spList = this.spWeb.get_lists().getByTitle(libraryName);

                // Load the SPList object with its EffectiveBasePermissions property.
                this.appContext.load(spList, 'EffectiveBasePermissions');

                // Function that runs on success (has access).
                var resolve = function resolve() {
                    if (spList.get_effectiveBasePermissions().has(SP.PermissionKind.deleteListItems)) {
                        onHasAccess();
                    } else {
                        onNoAccess();
                    }
                };

                // Function to run when no access.
                var _reject7 = function _reject7(sender, args) {
                    onNoAccess();
                };

                this.appContext.executeQueryAsync(resolve, _reject7);
            } catch (error) {
                throw 'spHelper Error: Unable to validate user permissions for list ... ' + error;
            }
        }

        /**
         * Verifies if the user executing the command has specific access to a specific list.
         *
         * PARAMETERS
         *      libraryName - [STRING] : A string representing the list name.
         *      permissionType - [STRING] : A string representing the SPBasePermission (https://docs.microsoft.com/en-us/previous-versions/office/developer/sharepoint-2010/ee556747(v%3Doffice.14))
         */

    }, {
        key: 'hasListAccess',
        value: function hasListAccess(libraryName, permissionType, onHasAccess, onNoAccess) {
            if (typeof libraryName == 'undefined' || typeof permissionType == 'undefined') {
                throw 'spHelper Error: Invalid has list access details. To check access, libraryName and permissionType must be defined.';
            }

            try {
                // Refresh connections.
                this.refreshConnection();

                // Will store the spList object when request is complete.
                var spList = this.spWeb.get_lists().getByTitle(libraryName);

                // Load the SPList object with its EffectiveBasePermissions property.
                this.appContext.load(spList, 'EffectiveBasePermissions');

                // Callback function when the request (promise) is resolved.
                var resolve = function resolve() {
                    if (spList.get_effectiveBasePermissions().has(SP.PermissionKind[permissionType])) {
                        onHasAccess();
                    } else {
                        onNoAccess();
                    }
                };

                // Callback function when the request (promise) has rejected.
                var _reject8 = function _reject8(sender, args) {
                    onNoAccess();
                };

                this.appContext.executeQueryAsync(resolve, _reject8);
            } catch (error) {
                throw 'spHelper Error: Unable to validate user permissions ... ' + error;
            }
        }

        /**
        * A collection of SharePoint ribbon utilities.
        *
        * UTILITIES [OBJECTS]
        *      removeEdit - [FUNCTION] : Selects the browse tab and removes the edit tab.
        */

    }, {
        key: 'spRibbon',
        value: function spRibbon() {
            return {
                // Remove the EDIT tab in the SharePoint ribbon. Makes 3 attempts due to ribbon loading timing.
                removeEdit: function removeEdit() {
                    var removeEditPromise = new Promise(function (resolve, reject) {
                        ExecuteOrDelayUntilScriptLoaded(function () {
                            var attempts = 3;
                            var editRemoved = false;

                            var startRemoveEdit = function startRemoveEdit() {
                                try {
                                    // Get the SharePoint Ribbon object.
                                    var spRibbon = SP.Ribbon.PageManager.get_instance().get_ribbon();

                                    // If the ribbon is active.
                                    if (Boolean(spRibbon)) {
                                        // Check if the edit tab is present.
                                        if (Boolean(spRibbon.getChildByTitle('Edit'))) {
                                            // Remove the edit tab.
                                            spRibbon.removeChild('Ribbon.ListForm.Edit');
                                        }

                                        // Select the read (browse) tab.
                                        spRibbon.selectTabById('Ribbon.Read');

                                        // Indicate that the title was set.
                                        editRemoved = true;

                                        resolve();
                                    }
                                    // Ribbon is not found / ready.
                                    else {
                                            if (!editRemoved && attempts > 0) {
                                                // Wait 0.2 seconds and try again.
                                                setTimeout(startRemove, 200);

                                                attempts--;
                                            } else {
                                                reject('Unable to remove the edit tab (Ribbon took to long to load). Try refreshing this page.');
                                            }
                                        }
                                } catch (errorMessage) {
                                    reject(errorMessage);
                                }
                            };

                            startRemoveEdit();
                        }, 'sp.ui.dialog.js');
                    });

                    return removeEditPromise;
                }
            };
        }

        /**
         * A collection of SharePoint modal dialog utilities.
         *
         * UTILITIES [OBJECTS]
         *      isActive - [FUNCTION] : Checks if a SharePoint modal dialog is active. Accepts success/fail callbacks.
         *      resize   - [FUNCTION] : Resizes the modal dialog.
         *      close    - [FUNCTION] : Closes the top most dialog.
         *      setTitle - [FUNCTION] : Sets the title of the modal dialog window.
         */

    }, {
        key: 'spModalDialog',
        value: function spModalDialog() {
            return {
                // Checks if there is a modal dialog window active.
                isActive: function isActive() {
                    var isActivePromise = new Promise(function (resolve, reject) {
                        ExecuteOrDelayUntilScriptLoaded(function () {
                            try {
                                // Get the SharePoint Ribbon object.
                                var modalDialog = SP.UI.ModalDialog.get_childDialog();

                                // If the dialog is active.
                                if (Boolean(modalDialog)) {
                                    resolve(true);
                                }
                                // Dialog is not found / ready.
                                else {
                                        resolve(false);
                                    }
                            } catch (errorMessage) {
                                reject(errorMessage);
                            }
                        }, 'sp.ui.dialog.js');
                    });

                    return isActivePromise;
                },

                // Automatically resize the modal window according to the content within it.
                resize: function resize() {
                    var resizePromise = new Promise(function (resolve, reject) {
                        ExecuteOrDelayUntilScriptLoaded(function () {
                            var attempts = 3;
                            var resizeComplete = false;

                            var startResize = function startResize() {
                                try {
                                    // Get the SharePoint Ribbon object.
                                    var modalDialog = SP.UI.ModalDialog.get_childDialog();

                                    // If the dialog is active & if the content inside has loaded.
                                    if (Boolean(modalDialog) && modalDialog.$e_0 !== null) {
                                        // Ensure the dialog is not maximized.
                                        if (!modalDialog.$Q_0) {
                                            // Auto resize the dialog.
                                            modalDialog.autoSize();
                                        }

                                        // Indicate that the title was set.
                                        resizeComplete = true;

                                        resolve();
                                    }
                                    // Dialog is not found / ready.
                                    else {
                                            // If the title hasn't been set and there are attempts left, try again.
                                            if (!resizeComplete && attempts > 0) {
                                                // Wait 0.2 seconds and try again.
                                                setTimeout(startResize, 200);

                                                attempts--;
                                            } else {
                                                reject('Unable to resize the modal dialog (Took to long to load). Try refreshing this window.');
                                            }
                                        }
                                } catch (errorMessage) {
                                    reject(errorMessage);
                                }
                            };

                            startResize();
                        }, 'sp.ui.dialog.js');
                    });

                    return resizePromise;
                },

                // Closes the top most modal dialog.
                close: function close() {
                    ExecuteOrDelayUntilScriptLoaded(function () {
                        var attempts = 3;

                        var startClose = function startClose() {
                            try {
                                // Get the SharePoint Ribbon object.
                                var modalDialog = SP.UI.ModalDialog.get_childDialog();

                                // If the dialog is active & if the content inside has loaded.
                                if (Boolean(modalDialog)) {
                                    // Closes the dialog.
                                    modalDialog.close();
                                }
                                // Dialog is not found / ready.
                                else {
                                        if (attempts > 0) {
                                            // Wait 0.2 seconds and try again.
                                            setTimeout(startClose, 200);

                                            attempts--;
                                        }
                                    }
                            } catch (errorMessage) {
                                console.log(errorMessage);
                            }
                        };

                        startClose();
                    }, 'sp.ui.dialog.js');
                },

                // Closes the top most modal dialog.
                setTitle: function setTitle(newTitle) {
                    var setTitlePromise = new Promise(function (resolve, reject) {
                        ExecuteOrDelayUntilScriptLoaded(function () {
                            var attempts = 3;
                            var titleSet = false;

                            var startSetTitle = function startSetTitle() {
                                try {
                                    // Get the SharePoint Ribbon object.
                                    var modalDialog = SP.UI.ModalDialog.get_childDialog();

                                    // If the dialog is active & if the content inside has loaded.
                                    if (Boolean(modalDialog) && modalDialog.$e_0 !== null) {
                                        // Closes the dialog.
                                        SP.UI.UIUtility.setInnerText(parent.document.getElementById('dialogTitleSpan'), newTitle);

                                        // Indicate that the title was set.
                                        titleSet = true;

                                        resolve();
                                    }
                                    // Dialog is not found / ready.
                                    else {
                                            // If the title hasn't been set and there are attempts left, try again.
                                            if (!titleSet && attempts > 0) {
                                                // Wait 0.2 seconds and try again.
                                                setTimeout(startSetTitle, 200);

                                                attempts--;
                                            } else {
                                                reject('Unable to set SharePoint Modal Title (Took to long to load). Try refreshing this window.');
                                            }
                                        }
                                } catch (errorMessage) {
                                    reject(errorMessage);
                                }
                            };

                            startSetTitle();
                        }, 'sp.ui.dialog.js');
                    });

                    return setTitlePromise;
                }
            };
        }
    }]);

    return spHelper;
}();

exports.default = spHelper;
