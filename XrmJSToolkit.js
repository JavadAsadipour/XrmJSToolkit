/**
* Dynamics CRM Online, JavaScript Toolkit for JavaScript
* @author Javad Asadipour
* @current version : 1.0.0

* Credits:
*   The idea of this library was inspired by Jaimie Ji, David Berry
*
* Date: December, 2022
*
* What's new:
**********************************************************************************************************
*   Version: 1.0.0
*   Date: December, 2022
*       Sections Added:
*           XrmJSToolkit.Form
*           XrmJSToolkit.Api
*           XrmJSToolkit.Enumerables
*           XrmJSToolkit.App
*           XrmJSToolkit.User
**/


var XrmJSToolkit = (function () {

    var executionContext;

    var formContext;

    var globalContext;


    var initialize = function (executionContext) {

        executionContext = executionContext;

        formContext = executionContext.getFormContext();

        globalContext = Xrm.Utility.getGlobalContext();

    };


    var isArray = function (input) {

        return input.constructor.toString().indexOf("Array") !== -1;

    };


    var isString = function (input) {

        return input.constructor.toString().indexOf("String") !== -1;

    };

    //Enumerables Class

    var enumerables = (function () {

        const fieldNotificationLevels = {

            Recommendation: "RECOMMENDATION",

            Error: "ERROR"

        };


        const formNotificationLevels = {

            Info: "INFO",

            Warning: "WARNING",

            Error: "ERROR"

        };


        const formTypes = {

            Undefined: 0,

            Create: 1,

            Update: 2,

            ReadOnly: 3,

            Disabled: 4,

            BulkEdit: 6

        };


        return {

            FieldNotificationLevels: fieldNotificationLevels,

            FormNotificationLevels: formNotificationLevels,

            FormTypes: formTypes

        };

    })();

    //END Enumerables Class


    //Form Class

    var form = (function () {

        var hideField = function (fieldName) {

            if (isArray(fieldName)) {

                fieldName.forEach(

                    function (field) {

                        formContext.getControl(field).setVisible(false);

                    });

            } else if (isString(fieldName)) {

                formContext.getControl(fieldName).setVisible(false);

            }

        };


        let showField = function (fieldName) {

            if (isArray(fieldName)) {

                fieldName.forEach(

                    function (field) {

                        formContext.getControl(field).setVisible(true);

                    });

            } else if (isString(fieldName)) {

                formContext.getControl(fieldName).setVisible(true);

            }

        };


        let enableField = function (fieldName) {

            if (isArray(fieldName)) {

                fieldName.forEach(

                    function (field) {

                        formContext.getControl(field).setDisabled(false);

                    });

            } else if (isString(fieldName)) {

                formContext.getControl(fieldName).setDisabled(false);

            }

        };


        let disableField = function (fieldName) {

            if (isArray(fieldName)) {

                fieldName.forEach(

                    function (field) {

                        formContext.getControl(field).setDisabled(true);

                    });

            } else if (isString(fieldName)) {

                formContext.getControl(fieldName).setDisabled(true);

            }

        };


        let setFieldAsRequired = function (fieldName) {

            if (isArray(fieldName)) {

                fieldName.forEach(

                    function (field) {

                        updateRequirementLevel(field, "required");

                    });

            } else if (isString(fieldName)) {

                updateRequirementLevel(fieldName, "required");

            }

        };


        let setFieldAsOptional = function (fieldName) {

            if (isArray(fieldName)) {

                fieldName.forEach(

                    function (field) {

                        updateRequirementLevel(field, "none");

                    });

            } else if (isString(fieldName)) {

                updateRequirementLevel(fieldName, "none");

            }

        };


        let setFieldAsRecommended = function (fieldName) {

            if (isArray(fieldName)) {

                fieldName.forEach(

                    function (field) {

                        updateRequirementLevel(field, "recommended");

                    });

            } else if (isString(fieldName)) {

                updateRequirementLevel(fieldName, "recommended");

            }

        };


        let clearFieldValue = function (fieldName) {

            if (isArray(fieldName)) {

                fieldName.forEach(

                    function (field) {

                        formContext.getAttribute(field).setValue(null);

                    });

            } else if (isString(fieldName)) {

                formContext.getAttribute(fieldName).setValue(null);

            }

        };


        let removeOptionFromChoice = function (fieldName, optionValue) {

            if (isArray(fieldName)) {

                optionValue.forEach(

                    function (value) {

                        formContext.getControl(fieldName).removeOption(value);

                    });

            } else if (isString(fieldName)) {

                formContext.getControl(fieldName).removeOption(optionValue);

            }

        };


        let hideQuickForm = function (quickFormName) {

            if (isArray(quickFormName)) {

                quickFormName.forEach(

                    function (quickForm) {

                        formContext.ui.quickForms.get(quickForm).setVisible(false);

                    });

            } else if (isString(quickFormName)) {

                formContext.ui.quickForms.get(quickFormName).setVisible(false);

            }

        };


        let showQuickForm = function (quickFormName) {

            if (isArray(quickFormName)) {

                quickFormName.forEach(

                    function (quickForm) {

                        formContext.ui.quickForms.get(quickForm).setVisible(true);

                    });

            } else if (isString(quickFormName)) {

                formContext.ui.quickForms.get(quickFormName).setVisible(true);

            }

        };


        let getFieldValue = function (fieldName) {

            return formContext.getAttribute(fieldName).getValue();

        };


        let setFieldValue = function (fieldName, value) {

            formContext.getAttribute(fieldName).setValue(value);

        };


        var fieldIsNull = function (fieldName) {

            if (formContext.getAttribute(fieldName).getValue() == null)

                return true;

            return false;

        };


        let setLookupValue = function (fieldName, id, name, logicalName) {

            try {

                if (fieldName != null) {

                    let lookupValue = new Array();

                    lookupValue[0] = new Object();

                    lookupValue[0].id = id.replace("{", "").replace("}", "");

                    lookupValue[0].name = name;

                    lookupValue[0].entityType = logicalName;

                    if (lookupValue[0].id != null) {

                        formContext.getAttribute(fieldName).setValue(lookupValue);

                    }

                }

            } catch (e) {

                alert("Error in SetLookupValue: fieldName = " + fieldName + ", id = " + id + ", name = " + name + ", entityType = " + type + ", error = " + e);

            }

        };


        let disableAllControlsInForm = function () {

            formContext.ui.controls.forEach(function (control, i) {

                if (control && control.getDisabled && !control.getDisabled()) {

                    control.setDisabled(true);

                }

            });

        };


        let enableAllControlsInForm = function () {

            formContext.ui.controls.forEach(function (control, i) {

                if (control && control.getDisabled && control.getDisabled()) {

                    control.setDisabled(false);

                }

            });

        };


        let disableAllControlsInTab = function (tabName) {

            let tabControl = formContext.ui.tabs.get(tabName);

            if (tabControl != null) {

                formContext.ui.controls.forEach(
                    function (control) {

                        if (control.getParent() !== null && control.getParent().getParent() != null && control.getParent().getParent() === tabControl && control.getControlType() !== "subgrid") {

                            control.setDisabled(true);

                        }

                    });

            }

        };


        let enableAllControlsInTab = function (tabName) {

            let tabControl = formContext.ui.tabs.get(tabName);

            if (tabControl != null) {

                formContext.ui.controls.forEach(
                    function (control) {

                        if (control.getParent() !== null && control.getParent().getParent() != null && control.getParent().getParent() === tabControl && control.getControlType() !== "subgrid") {

                            control.setDisabled(false);

                        }

                    });

            }

        };


        let disableAllControlsInSection = function (tabName, sectionName) {

            let sectionControl = formContext.ui.tabs.get(tabName).sections.get(sectionName);

            sectionControl.controls.forEach(
                function (control) {

                    if (control.getControlType() !== "subgrid") {

                        control.setDisabled(true);

                    }

                });

        };


        let enableAllControlsInSection = function (tabName, sectionName) {

            let sectionControl = formContext.ui.tabs.get(tabName).sections.get(sectionName);

            sectionControl.controls.forEach(
                function (control) {

                    if (control.getControlType() !== "subgrid") {

                        control.setDisabled(false);

                    }

                });

        };


        let updateRequirementLevel = function (fieldName, levelName) {

            formContext.getAttribute(fieldName).setRequiredLevel(levelName);

        };


        let addFieldNotification = function (fieldName, message, level, uniqueId) {

            formContext?.getControl(fieldName)?.addNotification({

                messages: [message],

                notificationLevel: level,

                uniqueId: uniqueId,

            });

        };


        let clearFieldNotification = function (fieldName, uniqueId) {

            formContext?.getControl(fieldName)?.clearNotification(uniqueId);

        };


        let setFormNotification = function (message, level, uniqueId) {

            formContext?.ui.setFormNotification(message, level, uniqueId);

        };


        let clearFormNotification = function (uniqueId) {

            formContext?.ui.clearFormNotification(uniqueId);

        };


        let addFieldOnChange = function (fieldName, event) {

            formContext.getAttribute(fieldName).addOnChange(event);

        };


        let removeFieldOnChange = function (fieldName, event) {

            formContext.getAttribute(fieldName).removeOnChange(event);

        };


        let addSubgridOnLoad = function (subgridName, event) {

            formContext.getControl(subgridName).addOnLoad(event);

        };


        let removeSubgridOnLoad = function (subgridName, event) {

            formContext.getControl(subgridName).removeOnLoad(event);

        };


        let setFieldLabel = function (fieldName, label) {

            formContext.getControl(fieldName).setLabel(label);

        };


        let getCurrentRecordId = function () {

            return formContext.data.entity.getId().replace("{", "").replace("}", "");

        };


        let getCurrentForm = function () {

            return formContext.ui.formSelector.getCurrentItem();
        };

        let getCurrentFormType = function () {
            return formContext.ui.getFormType();
        }

        let getCurrentFormName = function () {

            return formContext.ui.formSelector.getCurrentItem().getLabel();
        };


        let getCurrentFormId = function () {

            return formContext.ui.formSelector.getCurrentItem().getId();
        };


        return {

            HideField: hideField,

            ShowField: showField,

            EnableField: enableField,

            DisableField: disableField,

            SetFieldAsRequired: setFieldAsRequired,

            SetFieldAsOptional: setFieldAsOptional,

            SetFieldAsRecommended: setFieldAsRecommended,

            ClearFieldValue: clearFieldValue,

            GetFieldValue: getFieldValue,

            SetFieldValue: setFieldValue,

            SetLookupValue: setLookupValue,

            FieldIsNull: fieldIsNull,

            SetFieldLabel: setFieldLabel,

            AddFieldNotification: addFieldNotification,

            ClearFieldNotification: clearFieldNotification,

            AddFieldOnChange: addFieldOnChange,

            RemoveFieldOnChange: removeFieldOnChange,

            RemoveOptionFromChoice: removeOptionFromChoice,

            EnableAllControlsInForm: enableAllControlsInForm,

            DisableAllControlsInForm: disableAllControlsInForm,

            EnableAllControlsInSection: enableAllControlsInSection,

            DisableAllControlsInSection: disableAllControlsInSection,

            EnableAllControlsInTab: enableAllControlsInTab,

            DisableAllControlsInTab: disableAllControlsInTab,

            SetFormNotification: setFormNotification,

            ClearFormNotification: clearFormNotification,

            ShowQuickForm: showQuickForm,

            HideQuickForm: hideQuickForm,

            GetCurrentForm: getCurrentForm,

            GetCurrentFormType: getCurrentFormType,

            GetCurrentFormName: getCurrentFormName,

            GetCurrentFormId: getCurrentFormId,

            AddSubgridOnLoad: addSubgridOnLoad,

            RemoveSubgridOnLoad: removeSubgridOnLoad,

            GetCurrentRecordId: getCurrentRecordId

        };

    })();

    //End Form Class


    //Api Class

    var api = (function () {

        let retrieveRecordAsync = async function (entityName, recordId, options) {

            let data = null;

            await Xrm.WebApi.retrieveRecord(entityName, recordId, options).then(
                function success(result) {

                    data = result;

                },

                function (error) {

                    console.log("ERROR in XrmJSToolkit RetrieveRecordAsync\r\n" + error.message);

                }
            );

            return data;

        };


        let retrieveMultipleRecordsAsync = async function (entityName, options, maxPageSize) {

            let data = null;

            await Xrm.WebApi.retrieveMultipleRecords(entityName, options, maxPageSize).then(
                function success(result) {

                    data = result.entities;

                },

                function (error) {

                    console.log("ERROR in XrmJSToolkit RetrieveMultipleRecordsAsync\r\n" + error.message);

                }
            );

            return data;

        };


        let createRecordAsync = async function (entityName, data) {

            var recordId = null;

            await Xrm.WebApi.createRecord(entityName, data).then(
                function success(result) {

                    recordId = result.id.replace("{", "").replace("}", "");

                },

                function (error) {

                    console.log("ERROR in XrmJSToolkit CreateRecordAsync\r\n" + error.message);

                }
            );

            return recordId;

        };


        let updateRecordAsync = async function (entityName, id, data) {

            var returnData = null;

            await Xrm.WebApi.updateRecord(entityName, id, data).then(
                function success(result) {

                    returnData = result;

                },

                function (error) {

                    console.log("ERROR in XrmJSToolkit UpdateRecordAsync\r\n" + error.message);

                }
            );

            return returnData;

        };


        let deleteRecordAsync = async function (entityName, id) {

            var returnData = null;

            await Xrm.WebApi.deleteRecord(entityName, id).then(
                function success(result) {

                    returnData = result;

                },

                function (error) {

                    console.log("ERROR in XrmJSToolkit DeleteRecordAsync\r\n" + error.message);

                }
            );

            return returnData;

        };


        let executeRequestAsync = async function (request) {

            let data = null;

            await Xrm.WebApi.online.execute(request).then(

                function success(result) {

                    data = result;

                },

                function (error) {

                    console.log("ERROR in XrmJSToolkit ExecuteRequestAsync\r\n" + error.message);

                }

            );

            return data;

        };


        let calculateRollupField = async function (entitySetName, recordId, rollupFieldName) {

            debugger;

            let url = Xrm.Page.context.getClientUrl() +

                "/api/data/v9.2/" + "CalculateRollupField(Target=@p1,FieldName=@p2)?" +

                "@p1={'@odata.id':'" + entitySetName + "(" + recordId + ")'}&" + "@p2='" + rollupFieldName + "'";

            let result = await makeRequest("GET", url);

        };


        let makeRequest = function (method, url) {

            return new Promise(function (resolve, reject) {

                let xhr = new XMLHttpRequest();

                xhr.open(method, url);

                xhr.onload = function () {

                    if (this.status >= 200 && this.status < 300) {

                        resolve(xhr.response);

                    } else {

                        reject({

                            status: this.status,

                            statusText: xhr.statusText

                        });

                    }

                };

                xhr.onerror = function () {

                    reject({

                        status: this.status,

                        statusText: xhr.statusText

                    });

                };

                xhr.send();

            });

        };

        return {

            RetrieveRecordAsync: retrieveRecordAsync,

            RetrieveMultipleRecordsAsync: retrieveMultipleRecordsAsync,

            CreateRecordAsync: createRecordAsync,

            UpdateRecordAsync: updateRecordAsync,

            DeleteRecordAsync: deleteRecordAsync,

            ExecuteRequestAsync: executeRequestAsync,

            CalculateRollupField: calculateRollupField

        };

    })();

    //End Api Class


    //App Class

    var app = (function () {

        let getCurrentApp = async function () {

            let appProperties = null;

            await globalContext.getCurrentAppProperties().then(
                function success(result) {

                    appProperties = result;

                },

                function (error) {

                    console.log("ERROR in XrmJSToolkit GetCurrentApp\r\n" + error.message);

                }
            );

            return appProperties;

        }

        return {

            GetCurrentApp: getCurrentApp

        };

    })();

    //End App Class


    //User Class

    var user = (function () {

        let hasCurrentUserRole = function (roleName) {

            let hasRole = false;

            let roles = Xrm.Utility.getGlobalContext().userSettings.roles;

            roles.forEach(role => {

                if (role.name === roleName) {

                    hasRole = true;

                }

            });

            return hasRole;

        }


        let getCurrentUserId = function () {

            return Xrm.Utility.getGlobalContext().userSettings.userId.replace("{", "").replace("}", "");

        }


        let getCurrentUserName = function () {

            return Xrm.Utility.getGlobalContext().userSettings.userName;

        }


        let getCurrentUserRoles = function () {

            return Xrm.Utility.getGlobalContext().userSettings.roles;

        }


        let getCurrentUserSettings = function () {

            return Xrm.Utility.getGlobalContext().userSettings;

        }


        let getCurrentUserBusinessUnitId = async function () {

            let userId = getCurrentUserId();

            let userInfo = await api.RetrieveRecordAsync("systemuser", userId, "?$select=fullname&$expand=businessunitid($select=name)");

            return userInfo.businessunitid.businessunitid.replace("{", "").replace("}", "");

        }


        let getCurrentUserBusinessUnitName = async function () {

            let userId = getCurrentUserId();

            let userInfo = await api.RetrieveRecordAsync("systemuser", userId, "?$select=fullname&$expand=businessunitid($select=name)");

            return userInfo.businessunitid.name;

        }


        return {

            GetCurrentUserId: getCurrentUserId,

            GetCurrentUserName: getCurrentUserName,

            GetCurrentUserRoles: getCurrentUserRoles,

            GetCurrentUserSettings: getCurrentUserSettings,

            HasCurrentUserRole: hasCurrentUserRole,

            GetCurrentUserBusinessUnitName: getCurrentUserBusinessUnitName,

            GetCurrentUserBusinessUnitId: getCurrentUserBusinessUnitId

        };

    })();

    //End User Class


    return {

        Initialize: initialize,

        Enums: enumerables,

        Form: form,

        Api: api,

        App: app,

        User: user

    };

})();
