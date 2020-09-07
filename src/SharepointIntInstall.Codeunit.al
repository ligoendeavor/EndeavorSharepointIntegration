codeunit 87002 "EDX09 Sharepoint Int. Install"
{
    Subtype = Install;

    trigger OnInstallAppPerCompany()
    var
        SPIntSetup: Record "EDX09 Sharepoint Int. Setup";
    begin
        if not SPIntSetup.get() then begin
            SPIntSetup."EDX09 Description" := 'MS Graph';
            SPIntSetup.Insert();
        end;
    end;


    // Evaluate(ClientID, '{7aff0926-311d-4c1f-b09f-ddd4502f8cae}');
    //                 SharepointIntSetup."EDX09 Description" := 'MSGraph';
    //                 SharepointIntSetup."EDX09 Service URL" := 'https://login.microsoftonline.com/e21dbd94-a418-4c71-bb92-a930e6009182/'; // Tenant ID
    //                 SharepointIntSetup."EDX09 Authorization URL Path" := 'oauth2/v2.0/authorize';
    //                 SharepointIntSetup."EDX09 Auth. Response Type" := 'code';
    //                 SharepointIntSetup."EDX09 Access Token URL Path" := 'oauth2/v2.0/token';
    //                 SharepointIntSetup."EDX09 Client ID" := ClientID;
    //                 SharepointIntSetup."EDX09 Scope" := 'https://graph.microsoft.com/.default';
    //                 SharepointIntSetup."EDX09 Redirect URL" := 'https://login.microsoftonline.com/common/oauth2/nativeclient';
    //                 SharepointIntSetup."EDX09 Sharepoint Base URL" := 'plmgroupcompany.sharepoint.com';
}