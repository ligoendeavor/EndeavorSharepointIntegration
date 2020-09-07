page 87000 "EDX09 Sharepoint Test Page"
{
    PageType = List;
    ApplicationArea = All;
    UsageCategory = Lists;
    SourceTable = "Sales Invoice Header";

    layout
    {
        area(Content)
        {
            repeater(GroupName)
            {
                field("No."; "No.")
                {
                    ApplicationArea = All;
                }
                field("Sell-to Customer No."; "Sell-to Customer No.")
                {
                    ApplicationArea = All;
                }
                field("Sell-to Customer Name"; "Sell-to Customer Name")
                {
                    ApplicationArea = All;
                }
            }
        }

    }

    actions
    {
        area(Processing)
        {
            action(CreateOAuthSetup)
            {
                ApplicationArea = All;

                trigger OnAction()
                var
                    SharepointIntSetup: Record "EDX09 Sharepoint Int. Setup";
                    ClientID: Guid;
                begin
                    SharepointIntSetup.DeleteAll();

                    Evaluate(ClientID, '{7aff0926-311d-4c1f-b09f-ddd4502f8cae}');
                    SharepointIntSetup."EDX09 Description" := 'MSGraph';
                    SharepointIntSetup."EDX09 Service URL" := 'https://login.microsoftonline.com/e21dbd94-a418-4c71-bb92-a930e6009182/'; // Tenant ID
                    SharepointIntSetup."EDX09 Authorization URL Path" := 'oauth2/v2.0/authorize';
                    SharepointIntSetup."EDX09 Auth. Response Type" := 'code';
                    SharepointIntSetup."EDX09 Access Token URL Path" := 'oauth2/v2.0/token';
                    SharepointIntSetup."EDX09 Client ID" := ClientID;
                    SharepointIntSetup."EDX09 Scope" := 'https://graph.microsoft.com/.default';
                    //OAuthSetup."Client Secret" := '8Hd_-5O~TBICyYXh5fD99TNI5.-O93vmSp';
                    //SharepointIntSetup."EDXpm Secret" := '2~Ohe.B6iY.1je.ftiHcT64fm1MVTN~dOz';
                    SharepointIntSetup."EDX09 Redirect URL" := 'https://login.microsoftonline.com/common/oauth2/nativeclient';
                    SharepointIntSetup."EDX09 Sharepoint Base URL" := 'plmgroupcompany.sharepoint.com';
                    if not SharepointIntSetup.Insert() then
                        SharepointIntSetup.modify();
                end;
            }
            action(GetAccessToken)
            {
                ApplicationArea = All;

                trigger OnAction();
                var
                    RequestJsonContent: JsonObject;
                    RequestUrlContent: Text;
                    ResponseJson: Text;
                    RequestJson: Text;
                    HttpError: Text;
                    UrlString: Text;
                    JObject: JsonObject;
                    JsonString: Text;
                    SPIntSetup: Record "EDX09 Sharepoint Int. Setup" temporary;
                    AccessToken: Text;
                    RefreshToken: Text;
                    ExpireInSec: BigInteger;
                begin
                    SPIntSetup.Get();
                    CreateContentRequestForAccessToken(RequestUrlContent, SPIntSetup.GetSecret(), SPIntSetup."EDX09 Client ID", SPIntSetup."EDX09 Redirect URL", SPIntSetup."EDX09 Scope");
                    CreateRequestJSONForAccessRefreshTokenUrlEncoded(RequestJson, SPIntSetup."EDX09 Service URL", SPIntSetup."EDX09 Access Token URL Path", RequestUrlContent);

                    RequestAccessAndRefreshTokens(RequestJson, ResponseJson, AccessToken, RefreshToken, ExpireInSec, HttpError);
                    message(AccessToken);

                    //PutDocumentOnSP(AccessToken);

                end;
            }
        }
    }

    local procedure PutDocumentOnSP(AccessToken: Text)
    var
        URL: Text;
        vRequestContent: HttpContent;
        vContentHeaders: HttpHeaders;
        vHttpRequestMessage: HttpRequestMessage;
        vHttpClient: HttpClient;
        vHttpResponseMessage: HttpResponseMessage;
        ResponseText: Text;
    begin
        URL := 'https://graph.microsoft.com/v1.0/sites/plmgroupcompany.sharepoint.com,bfc455a2-f7ce-4a13-a8b0-e52a8fd333c7,99184950-e351-4948-b7da-7a36458b3b6e/drives/b!olXEv873E0qosOUqj9Mzx1BJGJlR40hJt9p6NkWLO26aWYcqXdnhRbjmXLxrYqCO/root:/TestAccountPL2_b9c2f7cb-56ec-ea11-a817-0/invoice/INV-01567-L2T7N6_9f942398-edec-ea11-a817/Testfile2.txt:/content';
        //URL := 'https://graph.microsoft.com/v1.0/sites/endeavor.sharepoint.com,6b35f2d6-c98e-487d-8bd4-e01e9a909e1f,3fb50d38-3d8d-4e89-ae87-da95dc36fc1c/lists/28f66938-5f57-43a8-99f7-9aca6e539514/drive/root:/Business%20Documents/Testfile2.txt:/content';

        vRequestContent.WriteFrom('Data inside file');
        vRequestContent.GetHeaders(vContentHeaders);
        vContentHeaders.Clear();
        vContentHeaders.Add('Content-Type', 'text/plain');
        vHttpClient.DefaultRequestHeaders().Add('Authorization', StrSubstNo('Bearer %1', AccessToken));

        vHttpRequestMessage.Method := 'PUT';
        vHttpRequestMessage.SetRequestUri(URL);
        vHttpRequestMessage.Content := vRequestContent;

        vHttpClient.Send(vHttpRequestMessage, vHttpResponseMessage);

        vHttpResponseMessage.Content().ReadAs(ResponseText);
        Message(ResponseText);
    end;

    local procedure CreateContentRequestForAccessToken(var UrlString: Text; ClientSecret: Text; ClientID: Text; RedirectURI: Text; Scope: Text)
    var
        //HttpUtility: DotNet HttpUtility;
        TypeHelper: Codeunit "Type Helper";
        RequestContent: Text;
        ClientIdText: Text;
    begin
        //RequestContent := StrSubstNo('grant_type=client_credentials&client_secret=%1&client_id=%2&redirect_uri=%3&scope=%4', ClientSecret, ClientID, TypeHelper.UrlEncode(RedirectURI), TypeHelper.UrlEncode(Scope));
        //RequestContent := TypeHelper.UrlEncode(RequestContent);
        //UrlString := RequestContent;
        //UrlString := 'client_id=377e1515-0efa-43e0-bed0-b3fa2f261c79&scope=https://graph.microsoft.com/.default&client_secret=o4HbRPa~FUz8~eAq0INy111Nd.~_XyIH1d&grant_type=client_credentials';
        UrlString += 'client_secret=' + TypeHelper.UrlEncode(ClientSecret);
        ClientIdText := '377e1515-0efa-43e0-bed0-b3fa2f261c79';
        UrlString += '&client_id=' + TypeHelper.UrlEncode(ClientIdText);
        //UrlString += '&redirect_uri=' + TypeHelper.UrlEncode(RedirectURI);
        UrlString += '&scope=' + TypeHelper.UrlEncode(Scope);
        UrlString += '&grant_type=client_credentials';
    end;

    local procedure CreateContentRequestJSONForAccessToken(var JObject: JsonObject; ClientSecret: Text; ClientID: Text; Scope: Text; RedirectURI: Text)
    begin
        JObject.Add('grant_type', 'client_credentials');
        JObject.Add('client_secret', ClientSecret);
        JObject.Add('client_id', ClientID);
        JObject.Add('scope', Scope);
        JObject.Add('redirect_uri', RedirectURI);
        //JObject.Add('code', AuthorizationCode);
    end;

    local procedure CreateRequestJSONForAccessRefreshTokenUrlEncoded(var JsonString: Text; ServiceURL: Text; URLRequestPath: Text; Content: Text)
    var
        JObject: JsonObject;
    begin
        if JObject.ReadFrom(JsonString) then;
        JObject.Add('ServiceURL', ServiceURL);
        JObject.Add('Method', 'POST');
        JObject.Add('URLRequestPath', URLRequestPath);
        JObject.Add('Accept', 'application/json');
        JObject.Add('Content-Type', 'application/x-www-form-urlencoded');
        JObject.Add('Content', Content);
        JObject.WriteTo(JsonString);
    end;

    local procedure CreateRequestJSONForAccessRefreshToken(var JsonString: Text; ServiceURL: Text; URLRequestPath: Text; var ContentJson: JsonObject)
    var
        JObject: JsonObject;
    begin
        if JObject.ReadFrom(JsonString) then;
        JObject.Add('ServiceURL', ServiceURL);
        JObject.Add('Method', 'POST');
        JObject.Add('URLRequestPath', URLRequestPath);
        JObject.Add('Content-Type', 'application/json');
        JObject.Add('Content', ContentJson);
        JObject.WriteTo(JsonString);
    end;

    local procedure RequestAccessAndRefreshTokens(RequestJson: Text; var ResponseJson: Text; var AccessToken: Text; var RefreshToken: Text; var ExpireInSec: BigInteger; var HttpError: Text): Boolean
    begin
        AccessToken := '';
        RefreshToken := '';
        ResponseJson := '';
        if InvokeHttpJSONRequest(RequestJson, ResponseJson, HttpError) then
            exit(ParseAccessAndRefreshTokens(ResponseJson, AccessToken, RefreshToken, ExpireInSec));
    end;

    local procedure CallWebService()
    var
        URL: Text;
        vRequestContent: HttpContent;
        vContentHeaders: HttpHeaders;
        vHttpRequestMessage: HttpRequestMessage;
        vHttpClient: HttpClient;
        vHttpResponseMessage: HttpResponseMessage;
        ResponseText: Text;
    begin
        URL := 'https://login.microsoftonline.com/3b2de1f1-4306-444d-bed8-e724ef455f1d/oauth2/v2.0/token';

        vRequestContent.WriteFrom('grant_type=client_credentials&client_id=377e1515-0efa-43e0-bed0-b3fa2f261c79&client_secret=o4HbRPa~FUz8~eAq0INy111Nd.~_XyIH1d&scope=https://graph.microsoft.com/.default');
        vRequestContent.GetHeaders(vContentHeaders);
        vContentHeaders.Clear();
        vContentHeaders.Add('Content-Type', 'application/x-www-form-urlencoded');

        vHttpRequestMessage.Method := 'POST';
        vHttpRequestMessage.SetRequestUri(URL);
        vHttpRequestMessage.Content := vRequestContent;

        vHttpClient.Send(vHttpRequestMessage, vHttpResponseMessage);

        vHttpResponseMessage.Content().ReadAs(ResponseText);
        Message(ResponseText);
    end;

    local procedure InvokeHttpJSONRequest(RequestJson: Text; var ResponseJson: Text; var HttpError: Text) Result: Boolean
    var
        Client: HttpClient;
        RequestMessage: HttpRequestMessage;
        ResponseMessage: HttpResponseMessage;
        ErrorMessage: Text;
    begin
        ResponseJson := '';
        HttpError := '';

        Client.Clear();
        Client.Timeout(60000);
        InitHttpRequest(RequestMessage, RequestJson);
        //InitHttpRequestMessage(RequestMessage, RequestJson);

        if not Client.Send(RequestMessage, ResponseMessage) then
            if ResponseMessage.IsBlockedByEnvironment() then
                ErrorMessage := StrSubstNo(EnvironmentBlocksErr, RequestMessage.GetRequestUri())
            else
                ErrorMessage := StrSubstNo(ConnectionErr, RequestMessage.GetRequestUri());

        if ErrorMessage <> '' then
            Error(ErrorMessage);

        exit(ProcessHttpResponseMessage(ResponseMessage, ResponseJson, HttpError));
    end;

    local procedure InitHttpRequest(var RequestMessage: HttpRequestMessage; RequestJson: Text)
    var
        JToken: JsonToken;
        ContentJToken: JsonToken;
        ContentJson: Text;
        RequestContent: HttpContent;
        ContentHeaders: HttpHeaders;
    begin

        if JToken.ReadFrom(RequestJson) then begin
            begin

                if JToken.SelectToken('Content', ContentJToken) then begin
                    ContentJToken.WriteTo(ContentJson);
                    RequestContent.WriteFrom('grant_type=client_credentials&client_id=377e1515-0efa-43e0-bed0-b3fa2f261c79&client_secret=o4HbRPa~FUz8~eAq0INy111Nd.~_XyIH1d&scope=https://graph.microsoft.com/.default');

                    //RequestContent.WriteFrom('grant_type=client_credentials&client_id=377e1515-0efa-43e0-bed0-b3fa2f261c79&client_secret=o4HbRPa~FUz8~eAq0INy111Nd.~_XyIH1d&scope=https://graph.microsoft.com/.default');
                    RequestContent.GetHeaders(ContentHeaders);
                    ContentHeaders.Clear();
                    ContentHeaders.Add('Content-Type', 'application/x-www-form-urlencoded');

                end;

                //URL := 'https://login.microsoftonline.com/3b2de1f1-4306-444d-bed8-e724ef455f1d/oauth2/v2.0/token';

                // RequestMessage.WriteFrom('grant_type=client_credentials&client_id=377e1515-0efa-43e0-bed0-b3fa2f261c79&client_secret=o4HbRPa~FUz8~eAq0INy111Nd.~_XyIH1d&scope=https://graph.microsoft.com/.default');
                // RequestMessage.GetHeaders(vContentHeaders);
                // RequestMessage.Clear();
                // RequestMessage.Add('Content-Type', 'application/x-www-form-urlencoded');

                RequestMessage.Method := 'POST';
                RequestMessage.SetRequestUri('https://login.microsoftonline.com/3b2de1f1-4306-444d-bed8-e724ef455f1d/oauth2/v2.0/token');
                RequestMessage.Content := RequestContent;
            end;

        end;
    end;

    local procedure ProcessHttpResponseMessage(var ResponseMessage: HttpResponseMessage; var ResponseJson: Text; var HttpError: Text) Result: Boolean
    var
        ResponseJObject: JsonObject;
        ContentJObject: JsonObject;
        JToken: JsonToken;
        ResponseText: Text;
        JsonResponse: Boolean;
        StatusCode: Integer;
        StatusReason: Text;
        StatusDetails: Text;
    begin
        Result := ResponseMessage.IsSuccessStatusCode();
        StatusCode := ResponseMessage.HttpStatusCode();
        StatusReason := ResponseMessage.ReasonPhrase();

        if ResponseMessage.Content().ReadAs(ResponseText) then
            JsonResponse := ContentJObject.ReadFrom(ResponseText);
        if JsonResponse then
            ResponseJObject.Add('Content', ContentJObject);

        if not Result then begin
            HttpError := StrSubstNo('HTTP error %1 (%2)', StatusCode, StatusReason);
            if JsonResponse then
                if ContentJObject.SelectToken('error_description', JToken) then begin
                    StatusDetails := JToken.AsValue().AsText();
                    HttpError += StrSubstNo('\%1', StatusDetails);
                end;
        end;

        SetHttpStatus(ResponseJObject, StatusCode, StatusReason, StatusDetails);
        //Message(StatusDetails);
        ResponseJObject.WriteTo(ResponseJson);
    end;

    local procedure SetHttpStatus(var JObject: JsonObject; StatusCode: Integer; StatusReason: Text; StatusDetails: Text)
    var
        JObject2: JsonObject;
    begin
        JObject2.Add('code', StatusCode);
        JObject2.Add('reason', StatusReason);
        if StatusDetails <> '' then
            JObject2.Add('details', StatusDetails);
        JObject.Add('Status', JObject2);
    end;

    local procedure ParseAccessAndRefreshTokens(ResponseJson: Text; var AccessToken: Text; var RefreshToken: Text; var ExpireInSec: BigInteger) Result: Boolean
    var
        JToken: JsonToken;
        NewAccessToken: Text;
        NewRefreshToken: Text;
    begin
        AccessToken := '';
        RefreshToken := '';
        ExpireInSec := 0;

        if JToken.ReadFrom(ResponseJson) then
            if JToken.SelectToken('Content', JToken) then
                foreach JToken in JToken.AsObject().Values() do
                    case JToken.Path() of
                        'Content.access_token':
                            NewAccessToken := JToken.AsValue().AsText();
                        'Content.refresh_token':
                            NewRefreshToken := JToken.AsValue().AsText();
                        'Content.expires_in':
                            ExpireInSec := JToken.AsValue().AsBigInteger();
                    end;
        // if (NewAccessToken = '') or (NewRefreshToken = '') then
        //     exit(false);

        AccessToken := NewAccessToken;
        RefreshToken := NewRefreshToken;
        exit(true);
    end;


    var
        EnvironmentBlocksErr: Label 'Environment blocks an outgoing HTTP request to ''%1''.', Comment = '%1 - url, e.g. https://microsoft.com';
        ConnectionErr: Label 'Connection to the remote service ''%1'' could not be established.', Comment = '%1 - url, e.g. https://microsoft.com';
}