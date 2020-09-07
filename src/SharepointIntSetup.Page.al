page 87001 "EDX09 Sharepoint Int. Setup"
{
    PageType = Card;
    ApplicationArea = All;
    UsageCategory = Administration;
    SourceTable = "EDX09 Sharepoint Int. Setup";

    layout
    {
        area(Content)
        {
            group(SharepointIntegrationSetup)
            {
                Caption = 'Sharepoint Integration Setup';
                field("EDX09 Code"; "EDX09 Code")
                {
                    ApplicationArea = All;
                    Visible = false;
                }
                field("EDX09 Description"; "EDX09 Description")
                {
                    ApplicationArea = All;
                }
                field("EDX09 MS Graph URL Path"; "EDX09 MS Graph URL Path")
                {
                    ApplicationArea = All;
                }
                field("EDX09 Service URL"; "EDX09 Service URL")
                {
                    ApplicationArea = All;
                }
                field("EDX09 Sharepoint Base URL"; "EDX09 Sharepoint Base URL")
                {
                    ApplicationArea = All;
                }
                field("EDX09 Access Token URL Path"; "EDX09 Access Token URL Path")
                {
                    ApplicationArea = All;
                }
                field("EDX09 Authorization URL Path"; "EDX09 Authorization URL Path")
                {
                    ApplicationArea = All;
                }
                field("EDX09 Redirect URL"; "EDX09 Redirect URL")
                {
                    ApplicationArea = All;
                }
                field("EDX09 Scope"; "EDX09 Scope")
                {
                    ApplicationArea = All;
                }
                field("EDX09 Auth. Response Type"; "EDX09 Auth. Response Type")
                {
                    ApplicationArea = All;
                }
                field("EDX09 Client ID"; "EDX09 Client ID")
                {
                    ApplicationArea = All;
                }
                field(clitenSecret; ClientSecret)
                {
                    ApplicationArea = All;
                    Caption = 'Client Secret', Locked = true;
                    ToolTip = 'Specifies the secret of the API application.';

                    trigger OnValidate()
                    begin
                        SetSecret(ClientSecret);
                        Clear(ClientSecret);
                    end;
                }
            }
        }
    }

    var
        ClientSecret: Text;


}