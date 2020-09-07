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
}