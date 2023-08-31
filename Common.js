/**********************************************************************************************************
* Common.js
* @author Enrique Romero
* @current version : 1.0
***********************************************************************************************************
* Version: 1.0
* Date: Enero, 2023
***********************************************************************************************************/

if (typeof (Common) === "undefined")
    Common = { __namespace: true };

var formContext;

Common.Api = {

    Retrieve: function (entityLogicalName, id, options, successCallback, errorCallback) {
        "use strict";
        Xrm.WebApi.retrieveRecord(entityLogicalName, id, options).then(successCallback, errorCallback);
    },

    UpdateRecord: function (entityLogicalName, id, options, successCallback, errorCallback) {
        "use strict";
        Xrm.WebApi.updateRecord(entityLogicalName, id, options).then(successCallback, errorCallback);
    },

    CreateRecord: function (entityLogicalName, data, successCallback, errorCallback) {
        "use strict";
        Xrm.WebApi.createRecord(entityLogicalName, data).then(successCallback, errorCallback);
    },

    RetrieveWithPromise: function (entityLogicalName, id, options) {
        "use strict";
        return Xrm.WebApi.retrieveRecord(entityLogicalName, id, options);
    },

    RetrieveMultipleRecords: function (entityLogicalName, options, successCallback, errorCallback) {
        "use strict";
        Xrm.WebApi.retrieveMultipleRecords(entityLogicalName, options).then(successCallback, errorCallback);
    },

    RetrieveMultipleRecordsWithPromise: function (entityLogicalName, options, successCallback, errorCallback) {
        "use strict";
        return Xrm.WebApi.retrieveMultipleRecords(entityLogicalName, options);
    },

    RetrieveMultipleRecordsFetchXML: function (entityLogicalName, FetchXML, successCallback, errorCallback) {
        "use strict";
        Xrm.WebApi.retrieveMultipleRecords(entityLogicalName, "?fetchXml=" + FetchXML).then(successCallback, errorCallback);
    },

    RetrieveMultipleRecordsFetchXMLWithPromise: function (entityLogicalName, FetchXML, successCallback, errorCallback) {
        "use strict";
        return Xrm.WebApi.retrieveMultipleRecords(entityLogicalName, "?fetchXml=" + FetchXML);
    },

    DeleteRecord: function (entityLogicalName, id, successCallback, errorCallback) {
        "use strict";
        Xrm.WebApi.deleteRecord(entityLogicalName, id).then(successCallback, errorCallback);
    },
    GetEnvironmentVariable: function (variableName, successCallBack, errorCallback) {
        "use strict";
        Common.Api.RetrieveMultipleRecords("environmentvariablevalue", "?$select=value&$filter=(EnvironmentVariableDefinitionId/schemaname eq '" + variableName + "') and (EnvironmentVariableDefinitionId/environmentvariabledefinitionid ne null)", successCallBack, errorCallback);
    },
    CallPowerAutomate: function (EnvironmentVariableName_URL, Parameter1) {
        "use strict";
        Common.Api.GetEnvironmentVariable(EnvironmentVariableName_URL, function (responseData) {
            if (responseData.entities.length > 0) {
                var data = JSON.stringify({
                    "Parameter1": Parameter1.slice(1, -1)
                });

                var xhr = new XMLHttpRequest();

                xhr.addEventListener("readystatechange", function () {
                    if (this.readyState === 4) {
                        OnResultPowerAutomate(this.responseText);
                    }
                });

                xhr.open("POST", responseData.entities[0].value);
                xhr.setRequestHeader("Content-Type", "application/json");

                xhr.send(data);

            }
        })
    }
};

Common.FormContext = {
    Init: function (executionContext, IsFormContext) {
        "use strict";
        if (IsFormContext !== undefined && IsFormContext === true)
            formContext = executionContext;
        else {
            formContext = executionContext.getFormContext();
        }
    },
    GetFormId() {
        "use strict";
        return formContext.ui.formSelector.getCurrentItem().getId();
    },
    getEntityName() {
        "use strict";
        return formContext.data.entity.getEntityName();
    },
    getFormContext() {
        "use strict";
        return formContext;
    },

    getRecordID() {
        "use strict";
        return formContext.data.entity.getId();
    },

    GetValue: function (field) {
        "use strict";
        var fieldValue = formContext.getAttribute(field);

        if (fieldValue) {
            return fieldValue.getValue();
        }
        return null;
    },

    SetValue: function (field, value) {
        "use strict";
        formContext.getAttribute(field).setValue(value);
    },

    getTab: function (tabName) {
        "use strict";
        return formContext.ui.tabs.get(tabName);
    },

    TabVisible: function (tabName, isVisible) {
        "use strict";
        var tabobj = Common.FormContext.getTab(tabName);
        tabobj.setVisible(isVisible);
    },

    SectionVisible: function (tabName, sectionName, isVisible) {
        "use strict";
        var tabobj = Common.FormContext.getTab(tabName);
        tabobj.sections.get(sectionName).setVisible(isVisible);
    },

    getSelectedTab: function (tabName) {
        "use strict";
        var tabobj = Common.FormContext.getTab(tabName);
        return tabobj.getDisplayState();
    },

    getUIControl: function (ControlName) {
        "use strict";
        return formContext.ui.controls.get(ControlName);
    },

    setFieldDisabled: function (FieldName, disable) {
        "use strict";
        formContext.getControl(FieldName).setDisabled(disable);
    },
    // RequiredLevel: none, required, recommended
    setFieldRequiredLevel: function (FieldName, RequiredLevel) {
        "use strict";
        formContext.getAttribute(FieldName).setRequiredLevel(RequiredLevel)
    },

    setFieldVisibility: function (FieldName, visible) {
        "use strict";
        formContext.getControl(FieldName).setVisible(visible);
    },

    setFieldRequiredAndVisible: function (FieldName, requiredAndVisible) {
        "use strict";
        if (requiredAndVisible) {
            Common.FormContext.setFieldRequiredLevel(FieldName, "required");
            Common.FormContext.setFieldVisibility(FieldName, true);
        } else {
            Common.FormContext.setFieldRequiredLevel(FieldName, "none");
            Common.FormContext.setFieldVisibility(FieldName, false);
        }
    },

    AddFilterToLookup: function (LogicalNameEntity, FieldName, FilterAttribute, ValuesOfAttribute, FilterString, MyCustomFunction) {
        "use strict";
        if (ValuesOfAttribute !== null && ValuesOfAttribute !== undefined)
            if (ValuesOfAttribute.length !== 0)
                var filter = Common.FormContext.BuildFilter(FilterAttribute, ValuesOfAttribute);
            else return;
        else if (FilterString !== undefined)
            filter = FilterString;
        else if (MyCustomFunction === undefined)
            return;

        if (MyCustomFunction === undefined) {
            formContext.getControl(FieldName).addPreSearch(function () {
                formContext.getControl(FieldName).addCustomFilter(filter, LogicalNameEntity);
            });
        }
        else
            formContext.getControl(FieldName).addPreSearch(MyCustomFunction);
    },
    AddPreSearch: function (FieldName, CustomFunction) {
        "use strict";
        formContext.getControl(FieldName).addPreSearch(CustomFunction);
    },
    RemovePreSearch: function (FieldName, MyFunction) {
        "use strict";
        formContext.getControl(FieldName).removePreSearch(MyFunction);
    },

    BuildFilter: function (FilterAttribute, ValuesOfAttribute) {
        "use strict";
        var returnValue = "<filter type='or'>";
        for (var i = 0; i < ValuesOfAttribute.length; i++) {
            returnValue += "<condition attribute='" + FilterAttribute + "' operator='eq' value='" + ValuesOfAttribute[i] + "'/>";
        }
        returnValue += "</filter>";
        return returnValue;
    },
    //level: ERROR, WARNING, INFO
    setFormNotification: function (message, level, uniqueId) {
        "use strict";
        formContext.ui.setFormNotification(message, level, uniqueId);
    },
    clearFormNotification: function (uniqueId) {
        "use strict";
        formContext.ui.clearFormNotification(uniqueId);
    },
    setFieldNotification(FieldName, message, uniqueId) {
        "use strict";
        formContext.getControl(FieldName).setNotification(message, uniqueId);
    },
    clearFieldNotification(FieldName, uniqueId) {
        "use strict";
        formContext.getControl(FieldName).clearNotification(uniqueId);
    },
    AddOnChange: function (FieldName, Function) {
        "use strict";
        formContext.getAttribute(FieldName).addOnChange(Function);
    },
    RefreshGrid: function (GridName) {
        "use strict";
        formContext.getControl(GridName).refresh();
    },
    /*
     *          alertStrings:
     *             - confirmButtonLabel: (Optional) String. The confirm button label. If you do not specify the button label, OK is used as the button label.
     *             - text: String. The message to be displayed in the alert dialog.
     *             - title: (Optional) String. The title of the alert dialog.
     * 
     * */
    OpenAlertDialog: function (alertStrings, alertOptions) {
        "use strict";
        Xrm.Navigation.openAlertDialog(alertStrings, alertOptions);
    },
    OpenAlertDialogWithMessage: function (Message) {
        "use strict";
        let alertStrings = new Object();
        alertStrings.text = Message;
        Common.FormContext.OpenAlertDialog(alertStrings);
    },
    RefreshRecord: function (save) {
        "use strict";
        formContext.data.refresh(save);
    },
    Save: function (saveOptions, successCallback, errorCallback) {
        "use strict";
        formContext.data.save(saveOptions).then(successCallback, errorCallback);
    },
    SaveWithPromise: function (saveOptions) {
        "use strict";
        return formContext.data.save(saveOptions);
    },
    SetSubmitMode: function (field, mode) {
        "use strict";
        formContext.getAttribute(field).setSubmitMode(mode)
    }
};
Common.Utilities = {

    CallAzureFunction: function (parameter, data, method) {
        "use strict";

        return new Promise((resolve, reject) => {
            var callAzureFunction = {
                Data: data,
                Parameter: parameter,
                Method: method,

                getMetadata: function () {
                    return {
                        boundParameter: null,
                        parameterTypes: {
                            Data: { typeName: "Edm.String", structuralProperty: 1 },
                            Parameter: { typeName: "Edm.String", structuralProperty: 1 },
                            Method: { typeName: "Edm.String", structuralProperty: 1 }
                        },
                        operationType: 0,
                        operationName: "urb_CallAzureFunction"
                    };
                }
            };


            window.parent.Xrm.WebApi.online.execute(callAzureFunction).then(
                resolve,
                reject
            );
        });
    },
    ShowMessageInDialog: function (Message) {
        "use strict";
        Xrm.Navigation.openAlertDialog({ text: Message });
    },
    OpenNavigateTo: function (executionContext, entityName,nameCustomPage, Title) {
        if(Common.FormContext.getFormContext() == undefined)

        var pageInput = {
            pageType: "custom",
            name: "erf_custompagecombinepricelist_2ae8b",
            entityName: "quote",
            recordId: Common.FormContext.getRecordID().slice(1,-1)
        };
        var navigationOptions = {
            target: 2,
            height: 600,
            width: 1200,
            title:"Products"

        };
        Xrm.Navigation.navigateTo(pageInput, navigationOptions)
            .then(
                function () {
                    // Called when page opens
                }
            ).catch(
                function (error) {
                    // Handle error
                }
            );
    }
};

Common.FormType = {

    GetFormType: function () {
        "use strict";
        return formContext.ui.getFormType();
    },

    Undefined: 0,

    Create: 1,

    Update: 2,

    ReadOnly: 3,

    Disabled: 4,

    BulkEdit: 6
};