table 87000 "EDX09 Sharepoint Int. Setup"
{
    Caption = 'Sharepoint Integration Setup';
    ReplicateData = false;

    fields
    {
        field(87001; "EDX09 Code"; Code[20])
        {
            Caption = 'Code';
            NotBlank = true;
        }
        field(87002; "EDX09 Description"; Text[250])
        {
            Caption = 'Description';
        }
        field(87003; "EDX09 Service URL"; Text[250])
        {
            Caption = 'Service URL';

            trigger OnValidate()
            var
                WebRequestHelper: Codeunit "Web Request Helper";
            begin
                if "EDX09 Service URL" <> '' then
                    WebRequestHelper.IsSecureHttpUrl("EDX09 Service URL");
            end;
        }
        field(87004; "EDX09 Redirect URL"; Text[250])
        {
            Caption = 'Redirect URL';
        }
        field(87005; "EDX09 Client ID"; Guid)
        {
            Caption = 'Client ID';
            DataClassification = EndUserIdentifiableInformation;
        }
        field(87006; "EDX09 Client Secret"; Guid)
        {
            Caption = 'Client Secret';
            DataClassification = EndUserIdentifiableInformation;
        }
        field(87007; "EDX09 Access Token"; Guid)
        {
            Caption = 'Access Token';
            DataClassification = EndUserIdentifiableInformation;
        }
        field(87008; "EDX09 Refresh Token"; Guid)
        {
            Caption = 'Refresh Token';
            DataClassification = EndUserIdentifiableInformation;
        }
        field(87009; "EDX09 Authorization URL Path"; Text[250])
        {
            Caption = 'Authorization URL Path';

            trigger OnValidate()
            begin
                CheckAndAppendURLPath("EDX09 Authorization URL Path");
            end;
        }
        field(87010; "EDX09 Access Token URL Path"; Text[250])
        {
            Caption = 'Access Token URL Path';

            trigger OnValidate()
            begin
                CheckAndAppendURLPath("EDX09 Access Token URL Path");
            end;
        }
        field(87011; "EDX09 Refresh Token URL Path"; Text[250])
        {
            Caption = 'Refresh Token URL Path';

            trigger OnValidate()
            begin
                CheckAndAppendURLPath("EDX09 Refresh Token URL Path");
            end;
        }
        field(87012; "EDX09 Scope"; Text[250])
        {
            Caption = 'Scope';
        }
        field(87013; "EDX09 Auth. Response Type"; Text[250])
        {
            Caption = 'Authorization Response Type';
        }
        field(87015; "EDX09 Sharepoint Base URL"; Text[250])
        {
            Caption = 'Sharepoint Base URL';
        }
        field(87016; "EDX09 MS Graph URL Path"; Text[250])
        {
            Caption = 'MS Graph URL Path';

            trigger OnValidate()
            begin
                CheckAndAppendURLPostPath("EDX09 MS Graph URL Path");
            end;
        }
        field(87017; "EDX09 Access Token Due Date"; DateTime)
        {
            Caption = 'Access Token Due DateTime';
            Editable = false;
        }
        field(87018; "EDX09 Enabled"; Boolean)
        {
            Caption = 'Enabled';
        }
    }

    keys
    {
        key(Key1; "EDX09 Code")
        {
            Clustered = true;
        }
    }

    var
        SPIntMgmt: Codeunit "EDX09 Sharepoint Int. Mgmt.";

    trigger OnDelete()
    begin
        DeleteToken("EDX09 Client ID");
        DeleteToken("EDX09 Client Secret");
        DeleteToken("EDX09 Access Token");
    end;

    local procedure DeleteToken(TokenKey: Guid)
    begin
        if not HasToken(TokenKey) then
            exit;

        IsolatedStorage.Delete(TokenKey, DataScope::Module);
    end;

    procedure HasToken(TokenKey: Guid): Boolean
    begin
        exit(not IsNullGuid(TokenKey) and IsolatedStorage.Contains(TokenKey, DataScope::Module));
    end;

    local procedure CheckAndAppendURLPath(var value: Text)
    begin
        if value <> '' then
            if value[1] <> '/' then
                value := '/' + value;
    end;

    local procedure CheckAndAppendURLPostPath(var value: Text)
    begin
        if value <> '' then
            if value[StrLen(value) - 1] <> '/' then
                value := value + '/';
    end;

    procedure SetClientID(NewClientID: Text)
    begin
        if IsNullGuid("EDX09 Client ID") then
            "EDX09 Client ID" := CreateGuid;

        IsolatedStorage.Set("EDX09 Client ID", NewClientID, DATASCOPE::Module);
    end;

    procedure GetClientID(): Text
    var
        Value: Text;
    begin
        IsolatedStorage.Get("EDX09 Client ID", DATASCOPE::Module, Value);
        exit(Value);
    end;

    procedure SetClientSecret(NewSecret: Text)
    begin
        if IsNullGuid("EDX09 Client Secret") then
            "EDX09 Client Secret" := CreateGuid;

        IsolatedStorage.Set("EDX09 Client Secret", NewSecret, DATASCOPE::Module);
    end;

    procedure GetClientSecret(): Text
    var
        Value: Text;
    begin
        IsolatedStorage.Get("EDX09 Client Secret", DATASCOPE::Module, Value);
        exit(Value);
    end;

    procedure SetAccessToken(NewAccessToken: Text)
    begin
        if IsNullGuid("EDX09 Access Token") then
            "EDX09 Access Token" := CreateGuid;

        IsolatedStorage.Set("EDX09 Access Token", NewAccessToken, DATASCOPE::Module);
    end;

    procedure GetAccessToken(): Text
    var
        Value: Text;
    begin
        IsolatedStorage.Get("EDX09 Access Token", DATASCOPE::Module, Value);
        exit(Value);
    end;

}

