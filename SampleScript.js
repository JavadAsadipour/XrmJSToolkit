//Define the main method as Async to be able to await for async functions

FormLoad = async function (executionContext) {

    XrmJSToolkit.Initialize(executionContext);

    formContext = executionContext.getFormContext();

    XRMForm = XrmJSToolkit.Form;

    XRMApi = XrmJSToolkit.Api;

    XRMEnums = XrmJSToolkit.Enumerables;

    XRMApp = XrmJSToolkit.App;

    XRMUser = XrmJSToolkit.User;


    //returns object
    //{
    //     appId: "f452c3df-2bd0-eb11-bacc-000d3a272ad5"
    //     displayName: "SYSTEM ADMIN"
    //     uniqueName: "stq_systemadmin"
    //     url: ""
    //     webResourceId: "953b9fac-1e5e-e611-80d6-00155ded156f"
    //     webResourceName: "msdyn_/Images/AppModule_Default_Icon.png"
    //     welcomePageId: undefined
    //     welcomePageName: undefine
    //}
    var currentApp = await XRMApp.GetCurrentApp();
    var appDisplayName = currentApp.displayName;


    //Check if the current user has a specific security role
    //returns boolean true/false
    var userHasBasicUserRole = XRMUser.HasCurrentUserRole("Basic User")


    //Hides a field, or a list of fields
    XRMForm.HideField("stq_name");
    XRMForm.HideField(["stq_name", "stq_column1", "stq_column2", "stq_column3"]);


    //Shows a field, or a list of fields
    XRMForm.ShowField("stq_name");
    XRMForm.ShowField(["stq_name", "stq_column1", "stq_column2", "stq_column3"]);


    //Disables a field, or a list of fields
    XRMForm.DisableField("stq_name");
    XRMForm.DisableField(["stq_name", "stq_column1", "stq_column2", "stq_column3"]);


    //Enables a field, or a list of fields
    XRMForm.EnableField("stq_name");
    XRMForm.EnableField(["stq_name", "stq_column1", "stq_column2", "stq_column3"]);


    //Make a field, or a list of fields Required
    XRMForm.SetFieldAsRequired("stq_name");
    XRMForm.SetFieldAsRequired(["stq_name", "stq_column1", "stq_column2", "stq_column3"]);


    //Make a field, or a list of fields Optional
    XRMForm.SetFieldAsOptional("stq_name");
    XRMForm.SetFieldAsOptional(["stq_name", "stq_column1", "stq_column2", "stq_column3"]);


    //Make a field, or a list of fields Recommended
    XRMForm.SetFieldAsRecommended("stq_name");
    XRMForm.SetFieldAsRecommended(["stq_name", "stq_column1", "stq_column2", "stq_column3"]);


    //Set a field, or a list of fields' value to null
    XRMForm.ClearFieldValue("stq_name");
    XRMForm.ClearFieldValue(["stq_name", "stq_column1", "stq_column2", "stq_column3"]);


    //Gets the value of a field
    var name = XRMForm.GetFieldValue("stq_name");
    var ownerLookup = XRMForm.GetFieldValue("ownerid");


    //Sets the value of a field (Works for all fields excluding Lookups. For Lookups use "SetLookupValue" method)
    XRMForm.SetFieldValue("stq_column1", "MyValue1");
    XRMForm.SetFieldValue("stq_column2", 123456);


    var dateData = new Date("2012-Aug-17 13:12");
    XRMForm.SetFieldValue("stq_datecolumn1", dateData);

    const choiceColumn1Enum = {
        Choice1: 956520001,
        Choice2: 956520002,
        Choice3: 956520003,
        Choice4: 956520004,
        Choice5: 956520005
    };
    XRMForm.SetFieldValue("stq_choicecolumn1", choiceColumn1Enum.Choice2);


    //Set a lookup value. The second parameter can be with or without {}.
    XRMForm.SetLookupValue("stq_parenttesttableid", "9279db88-e276-ed11-81ab-6045bd9054f7", "Sample2", "stq_testtable");
    XRMForm.SetLookupValue("stq_parenttesttableid", "{9279db88-e276-ed11-81ab-6045bd9054f7}", "Sample2", "stq_testtable");


    //Removes option(s) from a choice control
    XRMForm.RemoveOptionFromChoice("stq_choicecolumn1", choiceColumn1Enum.Choice3);
    XRMForm.RemoveOptionFromChoice("stq_choicecolumn1", [choiceColumn1Enum.Choice3, choiceColumn1Enum.Choice5]);


    //Returns true if the field does not have value
    var parentIsNull = XRMForm.FieldIsNull("stq_parenttesttableid");


    //Sets the fields display label
    XRMForm.SetFieldLabel("stq_column1", "My New Label");


    //Adds a notification to a field.
    XRMForm.AddFieldNotification("stq_column1", "My Custom Error Message", XRMEnums.FieldNotificationLevels.Error, "NotificationId1");
    XRMForm.AddFieldNotification("stq_column1", "My Custom Recommendation Message", XRMEnums.FieldNotificationLevels.Error, "NotificationId2");


    //Removes a notification from a field.
    XRMForm.ClearFieldNotification("stq_column1", "NotificationId1");
    XRMForm.ClearFieldNotification("stq_column1", "NotificationId2");


    //Sets a notification for the current form.
    XRMForm.SetFormNotification("Notification Message 1", XRMEnums.FormNotificationLevels.Error, "NotificationId1");
    XRMForm.SetFormNotification("Notification Message 2", XRMEnums.FormNotificationLevels.Warning, "NotificationId2");
    XRMForm.SetFormNotification("Notification Message 3", XRMEnums.FormNotificationLevels.Info, "NotificationId3");


    //Removes the notificiton from the current form.
    XRMForm.ClearFormNotification("NotificationId2");


    //Adds the Onchange event to a field
    XRMForm.AddFieldOnChange("stq_column1", OnColumn1Change);


    //Removes the Onchange event from a field
    XRMForm.RemoveFieldOnChange("stq_column1", OnColumn1Change);


    //Make all controls of a form ReadOnly
    XRMForm.DisableAllControlsInForm();


    //Enables all controls on a form
    XRMForm.EnableAllControlsInForm();


    //Disables all controls of a tab.
    XRMForm.DisableAllControlsInTab("Tab_General");


    //Enables all controls of a tab.
    XRMForm.EnableAllControlsInTab("Tab_General");


    //Disables all controls of a section.
    XRMForm.DisableAllControlsInSection("Tab_General", "Section_1");


    //Enables all controls of a section.
    XRMForm.EnableAllControlsInSection("Tab_General", "Section_2");


    //Hides a QuickView control
    XRMForm.HideQuickForm("ParentQuickView");


    //Shows a QuickView control
    XRMForm.ShowQuickForm("ParentQuickView");


    //Returns the current form object from the form selector
    var currentForm = XRMForm.GetCurrentForm();


    //Returns the current form name
    var currentFormName = XRMForm.GetCurrentFormName();


    //Returns the current form id
    var currentFormId = XRMForm.GetCurrentFormId();



    //Adds the OnLoad event to a subgrid
    XRMForm.AddSubgridOnLoad("Subgrid_1", OnSubGridLoad);


    //Removes the OnLoad event from a subgrid
    XRMForm.RemoveSubgridOnLoad("Subgrid_1", OnSubGridLoad);

    //Returns current record Id
    var currentRecordId = XRMForm.GetCurrentRecordId();
    alert(currentRecordId);


    //Creates a record with the recordData object
    var recordData = {
        "stq_name": "Sample 4",
        "stq_column2": "1111111",
        "stq_parenttesttableid@odata.bind": "/stq_testtables(9279db88-e276-ed11-81ab-6045bd9054f7)"
    }
    await XRMApi.CreateRecordAsync("stq_testtable", recordData);


    //Remove option(s) from a Choice field.
    XRMForm.RemoveOptionFromChoice("stq_choicecolumn1", [choiceColumn1Enum.Choice3, choiceColumn1Enum.Choice5]);
    XRMForm.RemoveOptionFromChoice("stq_choicecolumn1", choiceColumn1Enum.Choice4);

}

function OnColumn1Change() {

    alert("Value changed to: " + XRMForm.GetFieldValue("stq_column1"));


    XRMForm.RemoveFieldOnChange("stq_column1", OnColumn1Change);

}

function OnSubGridLoad() {

    alert("SubGridLoaded");

}

