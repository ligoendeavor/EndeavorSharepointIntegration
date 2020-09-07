codeunit 87003 "EDX09 Sharepoint Int. Upgrade"
{
    Subtype = Upgrade;

    trigger OnUpgradePerCompany()
    var
        SPIntSetup: Record "EDX09 Sharepoint Int. Setup";
    begin
        if not SPIntSetup.get() then begin
            SPIntSetup."EDX09 Description" := 'MS Graph';
            SPIntSetup.Insert();
        end;
    end;
}