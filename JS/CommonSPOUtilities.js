//Register Scripts which won't load By default
RegisterSod('sp.taxonomy.js', '/_layouts/15/SP.Taxonomy.js');
RegisterSod('sp.requestexecutor.js', '/_layouts/15/SP.RequestExecutor.js');

var LoadAndExecuteSodfunction = window.LoadAndExecuteSodfunction || function (scriptKey, callback) {
    if (!ExecuteOrDelayUntilScriptLoaded(callback, scriptKey)) {
        LoadSodByKey(NormalizeSodKey(scriptKey));
    }
};


window.SPOUtilities = window.SPOUtilities || {};
window.NotificationUtilities = window.NotificationUtilities || {};
window.NotificationUtilities.Desktop = window.NotificationUtilities.Desktop || function () {
    var getPermissionsAndShow = function (myNotification) {
        if (!(window.NotificationUtilities.Desktop.isNotificationsAllowed)) {
            return;
        }
        if (Notify.isSupported && Notify.needsPermission && window.NotificationUtilities.Desktop.AskPermissionIfNeeded) {
            Notify.requestPermission(onPermissionGranted, onPermissionDenied);
        } else {
            myNotification.show();
        }

        function onPermissionGranted() {
            console.log('Permission has been granted by the user');
            myNotification.show();
        }

        function onPermissionDenied() {
            console.warn('Permission has been denied by the user');
        }
    }
    return { GetPermissionsAndShow: getPermissionsAndShow };
}();

//used http://adodson.com/notification.js for browser native notifications
window.NotificationUtilities.Desktop.isNotificationsAllowed = false;
window.NotificationUtilities.Desktop.AskPermissionIfNeeded = true;

//below function contains the constants which can be utilized just by calling the appropriate one instead of hard coding it.
window.SPOUtilities.Constants = window.SPOUtilities.Constants || function () {
    var FieldTypes = {
        Text: 'Text',
        Number: 'Number',
        Choice: 'Choice',
        TaxonomyFieldType: 'TaxonomyFieldType',
        Boolean: '',
        Computed: '',
        DateTime: ''
    };
}();

//defined namespaces seperately for CSOM calls and REST calls
window.SPOUtilities.CSOM = window.SPOUtilities.CSOM || {};
window.SPOUtilities.REST = window.SPOUtilities.REST || {};

var _arrayBufferToBase64 = function (buffer) {
    var binary = '';
    var bytes = new window.Uint8Array(buffer);
    var len = bytes.byteLength;
    for (var i = 0; i < len; i++) {
        binary += String.fromCharCode(bytes[i]);
    }
    return binary;
};

/**************************************CSOM Utilities *************************************/
//Taxonomy '
window.SPOUtilities.CSOM.Taxonomy = window.SPOUtilities.CSOM.Taxonomy || function () {

    //below functions gets the term names and values based on the parameters send to this function using properties.
    //Paramaters:
    // SiteUrl - current site url,
    // TermSetName - Parent TermSet to load the terms from,
    // Successcallback  - success callback that takes response as parameter.
    //failurecallback - failure call that takes parameter as error message
    var loadTerms = function (props) {
        LoadAndExecuteSodfunction('sp.js', function () {
            LoadAndExecuteSodfunction('sp.taxonomy.js', function () {
                var context = new SP.ClientContext(props.SiteUrl);
                var site = context.get_site();
                var terms = SP.Taxonomy.TaxonomySession.getTaxonomySession(context).getDefaultSiteCollectionTermStore().getSiteCollectionGroup(site, false).get_termSets().getByName(props.TermSetName, 1033).get_terms();
                context.load(terms);
                context.executeQueryAsync(function () {
                    var enumer = terms.getEnumerator();

                    var TermValues = [];
                    while (enumer.moveNext()) {
                        var term = enumer.get_current();
                        var TermValue = {
                            Label: term.get_name(),
                            Desc: term.get_description(),
                            ID: term.get_id().ToSerialized()
                        };
                        for (var index = 0; index < props.properties.length; index++) {
                            TermValue[props.properties[index]] = term.get_customProperties()[props.properties[index]];
                        }
                        TermValues.push(TermValue);
                    }

                    if (props.successcallback && typeof (props.successcallback) === "function")
                        props.successcallback(TermValues)
                }, function (s, a) {
                    props.failurecallback(a.get_message());
                });
            });
        })

    };

    return {
        loadTerms: loadTerms
    };
}();



window.SPOUtilities.CSOM.List = window.SPOUtilities.CSOM.List || function () {
    //below functions gets the list items based on the parameters send to this function using properties.
    //Paramaters:
    // SiteUrl - current site url,
    // camlViewXML - caml Query to load the items from list,
    // ListTitle - List Name for Query
    // Successcallback  - success callback that takes response as parameter.
    //failurecallback - failure call that takes parameter as error message
    var loadItems = function (props) {
        LoadAndExecuteSodfunction('sp.js', function () {
            var context = new SP.ClientContext(props.SiteUrl);
            var camlQuery = new SP.CamlQuery();
            camlQuery.set_viewXml(props.camlViewXML);
            var items = context.get_web().get_lists().getByTitle(props.ListTitle).getItems(camlQuery);
            context.load(items);
            context.executeQueryAsync(function () {
                var enumer = items.getEnumerator();
                var ListItems = [];
                while (enumer.moveNext()) {
                    var currentItem = enumer.get_current();
                    ListItems.push(currentItem.get_fieldValues());
                }
                if (props.successcallback && typeof (props.successcallback) === "function") {
                    if (window.Notify) {
                    var notification = new Notify('Items Loaded', {
                        body: props.ListTitle + ' Items Loaded',
                        icon: _spPageContextInfo.siteAbsoluteUrl + '/SiteAssets/Value-eXcellence/images/vx.png',
                        timeout: 4,
                        notifyClick: function () { window.open(); }
                    });
                    window.NotificationUtilities.Desktop.GetPermissionsAndShow(notification);
                    }
                    props.successcallback(ListItems);
                }

            }, function (s, a) {
                if (props.failurecallback && typeof (props.failurecallback) === "function") {
                    props.failurecallback(a.get_message());
                }
            });
        });
    };

    var loadContentTypes = function (props) {
        LoadAndExecuteSodfunction('sp.js', function () {
            var context = new SP.ClientContext(props.webUrl);
            var web = context.get_web();
            var list = web.get_lists().getByTitle(props.listTitle);
            var ctypes = list.get_contentTypes();
            context.load(ctypes);
            context.executeQueryAsync(function () {
                var ctypesArray = [];
                var enumer = ctypes.getEnumerator();
                while (enumer.moveNext()) {
                    var currentItem = enumer.get_current();
                    ctypesArray.push(currentItem);
                }
                props.successcallback(ctypesArray);
            }, function (s, a) {
                props.failurecallback(a.get_message());
            });
        });
    }
    //below functions loads the list on the parameters send to this function using properties.
    //Paramaters:
    // SiteUrl - current site url,
    // ListTitle - List Name for Query
    // Successcallback  - success callback that takes response as parameter.
    //failurecallback - failure call that takes parameter as error message
    var loadList = function (props) {
        LoadAndExecuteSodfunction('sp.js', function () {
            var context = new SP.ClientContext(props.SiteUrl);
            var list = context.get_web().get_lists().getByTitle(props.ListTitle);
            props.SelectedFields[0] = list;
            context.load(list);
            context.executeQueryAsync(function () {
                props.successcallback(list);
            }, function (s, a) {
                props.failurecallback(a.get_message());
            });
        });
    };

    //below functions loads the list of items on the ItemID, the parameters send to this function using properties.
    //Paramaters:
    // SiteUrl - current site url,
    // ListTitle - List Name for Query
    //ItemID - Item ID for Query
    // Successcallback  - success callback that takes response as parameter.
    //failurecallback - failure call that takes parameter as error message
    var loadItem = function (props) {
        LoadAndExecuteSodfunction('sp.js', function () {
            var context = new SP.ClientContext(props.SiteUrl);
            var list = context.get_web().get_lists().getByTitle(props.ListTitle);
            var item = list.getItemById(props.ItemID);
           // props.selectfields[0] = item;
            context.load(item);
            context.executeQueryAsync(function () {
                props.successcallback(item.get_fieldValues());
            }, function (s, a) {
                props.failurecallback(a.get_message());
            });
        });
    };

    //below functions gets the URL of the Icon File, parameters.
    //Paramaters:
    // filename - name of the file in library,
    // iconSize - Size of the Image/Icon
    // Successcallback  - success callback that takes response as parameter, failure call back.
    //failurecallback - failure call that takes parameter as error message
    var getIconUrl = function (filename, iconSize, successcallback, failurecallback) {
        LoadAndExecuteSodfunction('sp.js', function () {
            var ctx = SP.ClientContext.get_current();
            var iconname = ctx.get_web().mapToIcon(filename, '', iconSize);
            ctx.executeQueryAsync(function () {
                successcallback("/_layouts/15/images/" + iconname.m_value);
            }, function (s, a) {
                failurecallback(a.get_messsage());
            });
        });
    };

    //below functions sets the Taxonomy field in the list based on the parameters.
    //Paramaters:
    // context - Current context,
    // list - List Name in which the field has to update
    //item - List Item Object
    //columnName - columnName which you are updating
    //TaxonomyValue - new taxonomy field value for updation
    // Successcallback  - success callback that takes response as parameter, failure call back.
    //failurecallback - failure call that takes parameter as error message
    var setTaxonomyFieldValue = function (context, list, item, columnName, TaxonomyValue) {
        var taxColumn = list.get_fields().getByInternalNameOrTitle(columnName);
        var TypedTaxColumn = context.castTo(taxColumn, SP.Taxonomy.TaxonomyField);

        var newTaxValue = new SP.Taxonomy.TaxonomyFieldValue();
        newTaxValue.set_label(TaxonomyValue.label);
        newTaxValue.set_termGuid(TaxonomyValue.ID);
        newTaxValue.set_wssId(-1);
        TypedTaxColumn.setFieldValueByValue(item, newTaxValue);
    };

    //below functions adds the list item in a list on the parameters send to this function using properties.
    //Paramaters:
    // webUrl - current site url,
    // ListTitle - List Name for Query
    //FieldValues - fields values to add in the SharePoint list
    // Successcallback  - success callback that takes response as parameter, failure call back.
    //failurecallback - failure call that takes parameter as error message
    var addListItems = function (props) {
        var ctx = {};
        if (props.webUrl) {
            ctx = new SP.ClientContext(props.webUrl);
        }
        else {
            ctx = SP.ClientContext.get_current();
        }
        var list = ctx.get_web().get_lists().getByTitle(props.ListTitle);
        var columnsTypes = {}, columnNames = [], columns = [];
        if (props.FieldValues.length && props.FieldValues.length > 0) {
            var FieldValue = props.FieldValues[0];
            for (var ColumnName in FieldValue) {
                if (FieldValue.hasOwnProperty(ColumnName)) {
                    columnNames.push(ColumnName);
                }
            }
        }
        else {
            props.successcallback('No values to save');
            return;
        }

        $.each(columnNames, function (index, columnName) {
            columns[index] = list.get_fields().getByInternalNameOrTitle(columnName);
            ctx.load(columns[index]);
        });

        ctx.executeQueryAsync(function () {
            $.each(columnNames, function (index, columnName) {
                columnsTypes[columnName] = columns[index].get_typeAsString();
            });
            // save all items to list at a time
            var items = [];
            $.each(props.FieldValues, function (index, item) {
                var newItemCreateInfo = new SP.ListItemCreationInformation();
                var newListItem = list.addItem(newItemCreateInfo);
                items[index] = newListItem;

                //assign each column property
                for (var ColumnName in columnsTypes) {
                    if (columnsTypes.hasOwnProperty(ColumnName)) {
                        if (columnsTypes[ColumnName] !== 'TaxonomyFieldType' && columnsTypes[ColumnName] !== 'URL') {
                            items[index].set_item(ColumnName, item[ColumnName]);
                        }

                        else if (columnsTypes[ColumnName] === 'TaxonomyFieldType') {
                            setTaxonomyFieldValue(ctx, list, items[index], ColumnName, item[ColumnName])
                        }
                        else if (columnsTypes[ColumnName] !== 'URL') {

                        }
                        else if (columnsTypes[ColumnName] !== 'User') {

                        }
                    }
                }
                items[index].update();
            });

            ctx.executeQueryAsync(function () {
                props.successcallback(items);
            }, function (s, a) {
                props.failurecallback(a.get_message());
            });


        }, function (s, a) { deferred.reject(a.get_message()) });

    };

    //below functions updates the list item on the parameters send to this function using properties.
    //Paramaters:
    // webUrl - current site url,
    // ListTitle - List Name for Query
    //FieldValues - fields values to add in the SharePoint list
    // Successcallback  - success callback that takes response as parameter, failure call back.
    //failurecallback - failure call that takes parameter as error message
    var updateListItems = function (props) {
        var ctx = {};
        if (props.webUrl) {
            ctx = new SP.ClientContext(props.webUrl);
        }
        else {
            ctx = SP.ClientContext.get_current();
        }
        var list = ctx.get_web().get_lists().getByTitle(props.ListTitle);
        var columnsTypes = {}, columnNames = [], columns = [];
        if (props.FieldValues.length && props.FieldValues.length > 0) {
            var FieldValue = props.FieldValues[0];
            for (var ColumnName in FieldValue) {
                if (FieldValue.hasOwnProperty(ColumnName)) {
                    if (ColumnName !== 'ID')
                        columnNames.push(ColumnName);
                }
            }
        }
        else {
            props.successcallback('No values to save');
            return;
        }

        $.each(columnNames, function (index, columnName) {
            columns[index] = list.get_fields().getByInternalNameOrTitle(columnName);
            ctx.load(columns[index]);
        });

        ctx.executeQueryAsync(function () {
            $.each(columnNames, function (index, columnName) {
                columnsTypes[columnName] = columns[index].get_typeAsString();
            });
            // save all items to list at a time
            var items = [];
            $.each(props.FieldValues, function (index, item) {
                // var newItemCreateInfo = new SP.ListItemCreationInformation();
                var listItemToUpdate = list.getItemById(item.ID);
                items[index] = listItemToUpdate;

                //assign each column property
                for (var ColumnName in columnsTypes) {
                    if (columnsTypes.hasOwnProperty(ColumnName) && ColumnName !== 'ID') {
                        if (columnsTypes[ColumnName] !== 'TaxonomyFieldType' && columnsTypes[ColumnName] !== 'URL') {
                            items[index].set_item(ColumnName, item[ColumnName]);
                        }

                        else if (columnsTypes[ColumnName] === 'TaxonomyFieldType') {
                            setTaxonomyFieldValue(ctx, list, items[index], ColumnName, item[ColumnName])
                        }
                        else if (columnsTypes[ColumnName] !== 'URL') {

                        }
                        else if (columnsTypes[ColumnName] !== 'User') {

                        }
                    }
                }
                items[index].update();
                ctx.load(items[index]);
            });

            ctx.executeQueryAsync(function () {
                props.successcallback(items);
            }, function (s, a) {
                props.failurecallback(a.get_message());
            });


        }, function (s, a) { deferred.reject(a.get_message()) });

    };

    var deleteListItems = function (props) {
        LoadAndExecuteSodfunction('sp.js', function () {
            var context = new SP.ClientContext(props.SiteUrl);
            var list = context.get_web().get_lists().getByTitle(props.ListTitle);
            var item = list.getItemById(props.ItemID);
            item.recycle();
            context.executeQueryAsync(function () {
                props.successcallback();
            }, function (s, a) {
                props.failurecallback(a.get_message());
            });
        });
    };

    return {
        addListItems: addListItems,
        updateListItems: updateListItems,
        deleteListItems: deleteListItems,
        loadItems: loadItems,
        loadItem: loadItem,
        getIconUrl: getIconUrl,
        loadList: loadList,
        loadContentTypes: loadContentTypes
    };
}();


/**************************************REST Utilities *************************************/
// List Utilities
window.SPOUtilities.REST.List = window.SPOUtilities.REST.List || function () {

    //below function loads the list using REST services on the parameters send to this function using properties.
    //Paramaters:
    // SiteUrl - current site url,
    // ListTitle - List Name for Query
    //SelectValues - Include selected Columns in the Query
    //FilterCondition - Filter items to query selected items
    //SortCondition - Sort Condition to sort the items in query
    //Limit - No of items to limit the items using query
    // Successcallback  - success callback that takes response as parameter, failure call back.
    //failurecallback - failure call that takes parameter as error message
    var loadItems = function (props) {
        var url = props.SiteUrl + '/_api/web/lists/GetByTitle(\'' + props.ListTitle + '\')/items?';
        if (props.SelectValues && props.SelectValues.length > 0) {
            url = url + '$select=' + props.SelectValues;
        }
        if (props.FilterCondition && props.FilterCondition.trim() !== "") {
            url = url + '&$filter=' + props.FilterCondition;
        }
        if (props.SortCondition && props.SortCondition.trim() !== "") {
            url = url + '&$orderby=' + props.FilterCondition;
        }
        if (props.Limit && props.Limit > 0) {
            url = url + '&$top=' + props.Limit;
        }

        var info = {
            url: url,
            method: "GET",
            headers: {
                "Accept": "application/json; odata=verbose"
            },
            success: function (data) {
                var jsonObject = JSON.parse(data.body);
                if (props.successcallback && typeof (props.successcallback) === "function") {
                    props.successcallback(jsonObject.d.results);
                }
            },
            error: function (sender, args, errMsg) {
                props.failurecallback(errMsg);
            }
        };


        LoadAndExecuteSodfunction('sp.requestexecutor.js', function () {
            var executor = new SP.RequestExecutor(props.SiteUrl);
            executor.executeAsync(info);
        });
    };

    //below function loads file using URL on the parameters.
    //Paramaters:
    // SiteUrl - current site url,
    // docurl - Document URL for Query
    // Successcallback  - success callback that takes response as parameter, failure call back.
    //failurecallback - failure call that takes parameter as error message
    var loadfileByUrl = function (siteurl, docurl, successcallback, failurecallback) {
        LoadAndExecuteSodfunction('sp.requestexecutor.js', function () {
            var reqexecutor = new SP.RequestExecutor(siteurl);
            reqexecutor.executeAsync({
                url: siteurl + "/_api/web/getfilebyserverrelativeurl(@fileurl)/ListItemAllFields?@fileurl='" + docurl + "'",
                headers: {
                    'Accept': 'application/json;odata=verbose',
                    'Content-Type': 'application/json;odata=verbose'
                },
                success: function (data) {
                    successcallback(data)
                },
                error: function (s, a, errMsg) {
                    failurecallback(errMsg)
                }
            });
        });

    };


    return {
        loadItems: loadItems,
        loadfileByUrl: loadfileByUrl
    };
}();


window.SPOUtilities.REST.Library = window.SPOUtilities.REST.Library || function () {
    //Prive method for uploading a file in Library
    var doUpload = function (props) {
        //checkout file and upload
        var doUploadAction = function () {

            var info = {
                headers: {
                    "Accept": "application/json; odata=verbose",
                    "content-type": "application/json;odata=verbose",
                },
                url: props.webUrl + "/_api/web/lists/getbytitle('" + props.ListTitle + "')/RootFolder/Files/Add(url=@filename,overwrite=true)?@filename='" + encodeURIComponent(props.FileName) + "'",
                method: "POST",
                binaryStringRequestBody: true,
                body: props.body,
                success: function (data) {
                    var itemdata = JSON.parse(data.body).d;
                    if (props.successcallback && typeof (props.successcallback) === "function")
                        props.successcallback(itemdata);
                },
                error: function (sender, args, errMsg) {
                    if (props.failurecallback && typeof (props.failurecallback) === "function")
                        props.failurecallback(errMsg);

                }
            };
            LoadAndExecuteSodfunction("sp.requestexecutor.js", function () {
                var executor = new SP.RequestExecutor(props.siteUrl);
                executor.executeAsync(info);
            });
        };
        doUploadAction();
    };

    //below functions adds a file to the document library on the parameters send to this function using properties.
    //Paramaters:
    //
    // file - File object to Upload in Document Library
    // body - binady of the file object
    // Successcallback  - success callback that takes response as parameter, failure call back.
    // failurecallback - failure call that takes parameter as error message
    var uploadFile = function (props) {
        var newname = props.file.name;
        var newext = newname.substring(newname.lastIndexOf('.'));
        var reader = new FileReader();
        reader.onload = (function (theFile) {
            return function (e) {
                var body = null;
                if (FileReader.prototype.readAsBinaryString)
                    body = e.target.result;
                else
                    body = _arrayBufferToBase64(e.target.result);
                //uploadFile();
                props.body = body;
                doUpload(props);
            };
        })(props.file);
        if (reader.readAsBinaryString)
            reader.readAsBinaryString(props.file);
        else
            reader.readAsArrayBuffer(props.file);

    };

    //below functions Used to checkout the file and do the upload on the parameters send to this function using properties.
    //Paramaters:
    // SiteUrl - current site collection url,
    //webUrl - current site url
    //fileUrl - URL of the file to Checkout
    // Successcallback  - success callback that takes response as parameter, failure call back.
    //failurecallback - failure call that takes parameter as error message
    var checkOutFile = function (props) {
        var info = {
            headers: {
                "Accept": "application/json; odata=verbose",
                "content-type": "application/json;odata=verbose",
            },
            url: props.webUrl + "/_api/Web/GetFileByServerRelativeUrl(@fileurl)/checkout?@fileurl='" + props.fileUrl + "'",
            method: "POST",
            success: function (data) {
                var itemdata = JSON.parse(data.body).d;
                if (props.successcallback && typeof (props.successcallback) === "function")
                    props.successcallback(itemdata);
            },
            error: function (sender, args, errMsg) {
                if (props.failurecallback && typeof (props.failurecallback) === "function")
                    props.failurecallback(errMsg);
            }
        };
        LoadAndExecuteSodfunction("sp.requestexecutor.js", function () {
            var executor = new SP.RequestExecutor(props.siteUrl);
            executor.executeAsync(info);
        });

    };

    //below functions Used to checkin the file after the upload is completed, on the parameters send to this function using properties.
    //Paramaters:
    // SiteUrl - current site collection url,
    //webUrl - current site url
    //fileUrl - URL of the file to CheckIn
    //CheckinComment - Checkin Comments while checkin file
    //CheckinType - Type of File Checkin, Eg: Major version or minor version.
    // Successcallback  - success callback that takes response as parameter, failure call back.
    //failurecallback - failure call that takes parameter as error message
    var checkInFile = function (props) {
        var info = {
            headers: {
                "Accept": "application/json; odata=verbose",
                "content-type": "application/json;odata=verbose",
            },
            url: props.webUrl + "/_api/Web/GetFileByServerRelativeUrl(@fileurl)/CheckIn(comment='" + props.CheckinComment + "',checkintype=" + props.CheckinType + ")?@fileurl='" + props.fileUrl + "'",
            method: "POST",
            success: function (data) {
                var itemdata = JSON.parse(data.body).d;
                if (props.successcallback && typeof (props.successcallback) === "function")
                    props.successcallback(itemdata);
            },
            error: function (sender, args, errMsg) {
                if (props.failurecallback && typeof (props.failurecallback) === "function")
                    props.failurecallback(errMsg);
            }
        };
        LoadAndExecuteSodfunction("sp.requestexecutor.js", function () {
            var executor = new SP.RequestExecutor(props.siteUrl);
            executor.executeAsync(info);
        });

    };

    
    return {
        uploadFile: uploadFile,
        checkInFile: checkInFile,
        checkOutFile: checkOutFile
    };
}();