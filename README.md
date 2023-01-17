# Project Description

**XrmJSToolkit** is a JavaScript library which can be used for JavaScript Development under the platform for **Microsoft Dynamics CRM 365 Online**. The library contains four major parts regarding functions:

- **Form:** Includes methods for form-based components, such as HideField, DisableAllControlsInSection.
  - **HideField**
  - **ShowField**
  - **EnableField**
  - **DisableField**
  - **SetFieldAsRequired**
  - **SetFieldAsOptional**
  - **SetFieldAsRecommended**
  - **ClearFieldValue**
  - **GetFieldValue**
  - **SetFieldValue**
  - **SetLookupValue**
  - **FieldIsNull**
  - **SetFieldLabel**
  - **AddFieldNotification**
  - **ClearFieldNotification**
  - **AddFieldOnChange**
  - **RemoveFieldOnChange**
  - **RemoveOptionFromChoice**
  - **EnableAllControlsInForm**
  - **DisableAllControlsInForm**
  - **EnableAllControlsInSection**
  - **DisableAllControlsInSection**
  - **EnableAllControlsInTab**
  - **DisableAllControlsInTab**
  - **SetFormNotification**
  - **ClearFormNotification**
  - **ShowQuickForm**
  - **HideQuickForm**
  - **GetCurrentForm**
  - **GetCurrentFormType**
  - **GetCurrentFormName**
  - **GetCurrentFormId**
  - **AddSubgridOnLoad**
  - **RemoveSubgridOnLoad**
  - **GetCurrentRecordId**
  
- **App:** Methods related to the current application.
  - **GetCurrentApp**

- **Api:** Methods in which it calls Dynamics Web API, such as CreateRecord.
  - **RetrieveRecordAsync**
  - **RetrieveMultipleRecordsAsync**
  - **CreateRecordAsync**
  - **UpdateRecordAsync**
  - **DeleteRecordAsync**
  - **ExecuteRequestAsync**
  - **CalculateRollupField**

- **User:** Contains functions related to the system users.
  - **GetCurrentUserId**
  - **GetCurrentUserName**
  - **GetCurrentUserRoles**
  - **GetCurrentUserSettings**
  - **HasCurrentUserRole**
  - **GetCurrentUserBusinessUnitName**
  - **GetCurrentUserBusinessUnitId**

- **Enums:** You can find common enumerators in this section.
  - **FieldNotificationLevels**
  - **FormNotificationLevels**
  - **FormTypes**

