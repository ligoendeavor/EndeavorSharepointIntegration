codeunit 87001 "EDX09 Sharepoint Int. Mgmt."
{
    procedure GetAccessToken(var AccessToken: Text)
    var
        SPIntSetup: Record "EDX09 Sharepoint Int. Setup";
        RequestUrlContent: Text;
        RequestJson: Text;
        ResponseJson: Text;
        RefreshToken: Text;
        ExpireInSec: BigInteger;
        HttpError: Text;
    begin
        SPIntSetup.Get();

        if ((SPIntSetup."EDX09 Access Token Due Date" <> 0DT) AND (SPIntSetup."EDX09 Access Token Due Date" < CurrentDateTime())) or
            ((SPIntSetup."EDX09 Access Token Due Date" = 0DT) or (IsNullGuid(SPIntSetup."EDX09 Access Token"))) then begin
            CreateContentRequestForAccessToken(RequestUrlContent, SPIntSetup.GetClientSecret(), format(SPIntSetup.GetClientID()), SPIntSetup."EDX09 Redirect URL", SPIntSetup."EDX09 Scope");
            CreateRequestJSONForAccessRefreshTokenUrlEncoded(RequestJson, SPIntSetup."EDX09 Service URL", SPIntSetup."EDX09 Access Token URL Path", RequestUrlContent);
            RequestAccessAndRefreshTokens(RequestJson, ResponseJson, AccessToken, RefreshToken, ExpireInSec, HttpError);

            if HttpError = '' then
                SaveToken(SPIntSetup, AccessToken, ExpireInSec)
            else
                LogActivity(SPIntSetup, '', 'GetAccessToken', HttpError, RequestJson, ResponseJson, true);
        end else begin
            // Check if we can use last fetched access token
            AccessToken := SPIntSetup.GetAccessToken();
        end;
    end;

    procedure PutDocumentOnSP(AccessToken: Text; SalesInvoiceVariant: Variant; DocumentStream: InStream; FileName: Text; SharepointSite: Text; SharepointDocLibrary: Text; SharepointFullRelativeUrl: Text)
    var
        URL: Text;
        vRequestContent: HttpContent;
        vContentHeaders: HttpHeaders;
        vHttpRequestMessage: HttpRequestMessage;
        //vHttpClient: HttpClient;
        vHttpResponseMessage: HttpResponseMessage;
        RequestText: Text;
        ResponseText: Text;
        siteID: Text;
        driveID: Text;
        folderPath: Text;
        HttpError: Text;
        ErrorMessage: Text;
    begin
        ErrorMessage := '';
        vHttpClient.Clear();
        vHttpClient.Timeout(60000);

        if (SharepointSite = '') or (SharepointDocLibrary = '') or (SharepointFullRelativeUrl = '') or (FileName = '')
        then begin
            if GuiAllowed then
                message(CanNotConstructURL);
            exit;
        end;

        URL := ConstructUploadURL(AccessToken, SharepointSite, SharepointDocLibrary, SharepointFullRelativeUrl, FileName);

        if URL <> '' then begin
            vRequestContent.WriteFrom(DocumentStream);
            vRequestContent.GetHeaders(vContentHeaders);
            vContentHeaders.Clear();
            vContentHeaders.Remove('Content-Type');
            vContentHeaders.Add('Content-Type', 'application/octet-stream');
            vHttpClient.DefaultRequestHeaders().Add('Authorization', StrSubstNo('Bearer %1', AccessToken));

            vHttpRequestMessage.Method := 'PUT';
            vHttpRequestMessage.SetRequestUri(URL);
            vHttpRequestMessage.Content := vRequestContent;

            if not vHttpClient.Send(vHttpRequestMessage, vHttpResponseMessage) then
                if vHttpResponseMessage.IsBlockedByEnvironment() then
                    ErrorMessage := StrSubstNo(EnvironmentBlocksErr, vHttpRequestMessage.GetRequestUri())
                else
                    ErrorMessage := StrSubstNo(ConnectionErr, vHttpRequestMessage.GetRequestUri());

            if ErrorMessage <> '' then
                Error(ErrorMessage);


            vHttpResponseMessage.Content().ReadAs(ResponseText);
            ProcessHttpResponseMessage(vHttpResponseMessage, ResponseText, HttpError);

            if HttpError <> '' then begin
                LogActivity(SalesInvoiceVariant, FileName, 'PutDocumentOnSP', HttpError, '', ResponseText, true);
            end;
        end;
    end;

    local procedure CreateContentRequestForAccessToken(var UrlString: Text; ClientSecret: Text; ClientID: Text; RedirectURI: Text; Scope: Text)
    var
        //HttpUtility: DotNet HttpUtility;
        TypeHelper: Codeunit "Type Helper";
        RequestContent: Text;
        ClientIdText: Text;
    begin
        UrlString += 'client_secret=' + TypeHelper.UrlEncode(ClientSecret);
        UrlString += '&client_id=' + TypeHelper.UrlEncode(ClientId);
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

    local procedure InvokeHttpJSONRequest(RequestJson: Text; var ResponseJson: Text; var HttpError: Text) Result: Boolean
    var
        Client: HttpClient;
        RequestMessage: HttpRequestMessage;
        ResponseMessage: HttpResponseMessage;
        ErrorMessage: Text;
    begin
        ResponseJson := '';
        HttpError := '';
        ErrorMessage := '';

        Client.Clear();
        Client.Timeout(60000);
        InitHttpRequest(RequestMessage, RequestJson);

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
        Content: text;
        SPIntSetup: Record "EDX09 Sharepoint Int. Setup";
    begin

        if JToken.ReadFrom(RequestJson) then begin
            begin
                SPIntSetup.Get();
                if JToken.SelectToken('Content', ContentJToken) then begin
                    ContentJToken.WriteTo(ContentJson);
                    RequestContent.WriteFrom('grant_type=client_credentials&client_id=' + SPIntSetup.GetClientID() + '&client_secret=' + SPIntSetup.GetClientSecret() + '&scope=' + SPIntSetup."EDX09 Scope");
                    RequestContent.GetHeaders(ContentHeaders);
                    ContentHeaders.Clear();
                    ContentHeaders.Add('Content-Type', 'application/x-www-form-urlencoded');
                end;

                RequestMessage.Method := 'POST';
                RequestMessage.SetRequestUri(SPIntSetup."EDX09 Service URL" + SPIntSetup."EDX09 Access Token URL Path");
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

    local procedure SaveToken(var SPIntSetup: Record "EDX09 Sharepoint Int. Setup"; AccessToken: Text; ExpireInSec: BigInteger)
    begin
        with SPIntSetup do begin
            "EDX09 Access Token Due Date" := CurrentDateTime() + ExpireInSec * 1000;
            SetAccessToken(AccessToken);
            Modify();
            Commit();
        end;
    end;

    local procedure ConstructUploadURL(AccessToken: Text; SharepointSite: Text; SharepointDocLibrary: Text; SharepointFullRelativeUrl: Text; FileName: Text): Text
    var
        SiteID: Text;
        DriveID: Text;
        Folders: Text;
        Length: Integer;
        SPIntSetup: Record "EDX09 Sharepoint Int. Setup";
    begin
        SiteID := GetSPSiteID(AccessToken, SharepointSite);
        DriveID := GetSPDriveID(AccessToken, SharepointSite, SharepointDocLibrary);
        Length := StrLen(SharepointDocLibrary);
        Folders := CopyStr(SharepointFullRelativeUrl, StrPos(SharepointFullRelativeUrl, SharepointDocLibrary) + Length);

        SPIntSetup.Get();
        if (SiteID <> '') and (DriveID <> '') then
            exit(SPIntSetup."EDX09 MS Graph URL Path" + SiteID + '/drives/' + DriveID + '/root:' + Folders + '/' + FileName + ':/content')
        else
            exit('');
    end;

    local procedure GetSPSiteID(AccessToken: Text; SharepointSite: Text): Text
    var
        RequestContent: HttpContent;
        ContentHeaders: HttpHeaders;
        HttpRequestMessage: HttpRequestMessage;
        HttpClient: HttpClient;
        HttpResponseMessage: HttpResponseMessage;
        ResponseText: Text;
        ResponseJson: Text;
        SPIntSetup: Record "EDX09 Sharepoint Int. Setup";
        DocumentSite: Text;
        Length: Integer;
        HttpError: Text;
        ErrorMessage: Text;
    begin
        ErrorMessage := '';

        SPIntSetup.Get();
        Length := StrLen('https://' + SPIntSetup."EDX09 Sharepoint Base URL" + '/sites/');
        DocumentSite := CopyStr(SharepointSite, StrPos(SharepointSite, 'https://' + SPIntSetup."EDX09 Sharepoint Base URL" + '/sites/') + Length);

        HttpClient.Clear();
        HttpClient.Timeout(60000);
        HttpClient.DefaultRequestHeaders().Add('Authorization', StrSubstNo('Bearer %1', AccessToken));

        HttpRequestMessage.Method := 'GET';
        HttpRequestMessage.SetRequestUri(SPIntSetup."EDX09 MS Graph URL Path" + SPIntSetup."EDX09 Sharepoint Base URL" + ':/sites/' + DocumentSite);

        if not HttpClient.Send(HttpRequestMessage, HttpResponseMessage) then
            if HttpResponseMessage.IsBlockedByEnvironment() then
                ErrorMessage := StrSubstNo(EnvironmentBlocksErr, HttpRequestMessage.GetRequestUri())
            else
                ErrorMessage := StrSubstNo(ConnectionErr, HttpRequestMessage.GetRequestUri());

        if ErrorMessage <> '' then
            Error(ErrorMessage);

        HttpResponseMessage.Content().ReadAs(ResponseText);
        ResponseJson := ResponseText;
        ProcessHttpResponseMessage(HttpResponseMessage, ResponseJson, HttpError);

        if HttpError <> '' then begin
            LogActivity(SPIntSetup.RecordId(), '', 'GetSPSiteID', HttpError, '', ResponseJson, false);
        end;

        exit(ParseSPSiteID(ResponseText));
    end;

    local procedure ParseSPSiteID(ResponseJson: Text): Text
    var
        JObject: JsonObject;
        JToken: JsonToken;
        NewAccessToken: Text;
        NewRefreshToken: Text;
    begin
        if JObject.ReadFrom(ResponseJson) then
            if JObject.SelectToken('id', JToken) then
                exit(JToken.AsValue().AsText());

        exit('');
    end;

    local procedure GetSPDriveID(AccessToken: Text; SharepointSite: Text; SharepointDocLibrary: Text): Text
    var
        RequestContent: HttpContent;
        ContentHeaders: HttpHeaders;
        HttpRequestMessage: HttpRequestMessage;
        HttpClient: HttpClient;
        HttpResponseMessage: HttpResponseMessage;
        ResponseText: Text;
        ResponseJson: Text;
        SPIntSetup: Record "EDX09 Sharepoint Int. Setup";
        DocumentSite: Text;
        Length: Integer;
        HttpError: Text;
        ErrorMessage: Text;
    begin
        ErrorMessage := '';

        SPIntSetup.Get();
        Length := StrLen('https://' + SPIntSetup."EDX09 Sharepoint Base URL" + '/sites/');
        DocumentSite := CopyStr(SharepointSite, StrPos(SharepointSite, 'https://' + SPIntSetup."EDX09 Sharepoint Base URL" + '/sites/') + Length);

        HttpClient.Clear();
        HttpClient.Timeout(60000);
        HttpClient.DefaultRequestHeaders().Add('Authorization', StrSubstNo('Bearer %1', AccessToken));

        HttpRequestMessage.Method := 'GET';
        HttpRequestMessage.SetRequestUri(SPIntSetup."EDX09 MS Graph URL Path" + SPIntSetup."EDX09 Sharepoint Base URL" + ':/sites/' + DocumentSite + ':/drives');

        if not HttpClient.Send(HttpRequestMessage, HttpResponseMessage) then
            if HttpResponseMessage.IsBlockedByEnvironment() then
                ErrorMessage := StrSubstNo(EnvironmentBlocksErr, HttpRequestMessage.GetRequestUri())
            else
                ErrorMessage := StrSubstNo(ConnectionErr, HttpRequestMessage.GetRequestUri());

        if ErrorMessage <> '' then
            Error(ErrorMessage);

        HttpResponseMessage.Content().ReadAs(ResponseText);
        ResponseJson := ResponseText;
        ProcessHttpResponseMessage(HttpResponseMessage, ResponseJson, HttpError);

        if HttpError <> '' then begin
            LogActivity(SPIntSetup.RecordId(), '', 'GetSPDriveID', HttpError, '', ResponseJson, false);
        end;

        exit(ParseSPDriveID(ResponseText, SharepointDocLibrary));
    end;

    local procedure ParseSPDriveID(ResponseJson: Text; SharepointDocLibrary: Text): Text
    var
        JObject: JsonObject;
        JsonDriveObject: JsonObject;
        JsonDrives: JsonArray;
        JToken: JsonToken;
        JsonDrive: JsonToken;
        NewAccessToken: Text;
        NewRefreshToken: Text;
    begin
        if JObject.ReadFrom(ResponseJson) then begin
            if JObject.Get('value', JToken) then begin
                JsonDrives := JToken.AsArray();
                foreach JsonDrive in JsonDrives do begin
                    JsonDriveObject := JsonDrive.AsObject();
                    if JsonDriveObject.Get('name', JToken) then
                        if UpperCase(JToken.AsValue().AsText()) = UpperCase(SharepointDocLibrary) then begin
                            if JsonDriveObject.Get('id', JToken) then begin
                                exit(JToken.AsValue().AsText());
                            end;
                        end;
                end;
            end;
        end;

        exit('');
    end;

    local procedure ParseUploadSession(ResponseJson: Text; var UploadURL: Text) Result: Boolean
    var
        JObject: JsonObject;
        JToken: JsonToken;
        NewAccessToken: Text;
        NewRefreshToken: Text;
    begin
        if JObject.ReadFrom(ResponseJson) then
            if JObject.SelectToken('uploadUrl', JToken) then
                UploadURL := JToken.AsValue().AsText();

        exit(true);
    end;

    procedure LogActivity(RelatedVariant: Variant; ContextCode: Text; ActivityDescription: Text; ActivityMessage: Text; RequestJson: Text; ResponseJson: Text; MaskContent: Boolean)
    var
        ActivityLog: Record "Activity Log";
        JObject: JsonObject;
        JToken: JsonToken;
        RequestJObject: JsonObject;
        ResponseJObject: JsonObject;
        Context: Text[30];
        Status: Option;
        JsonString: Text;
        DetailedInfoExists: Boolean;
    begin
        Context := CopyStr(StrSubstNo('%1 %2', ActivityLogContextTxt, ContextCode), 1, MaxStrLen(Context));
        Status := ActivityLog.Status::Failed;

        ActivityLog.LogActivity(RelatedVariant, Status, Context, ActivityDescription, ActivityMessage);
        if RequestJObject.ReadFrom(MaskHeaders(RequestJson)) then begin
            if MaskContent then
                if RequestJObject.SelectToken('Content', JToken) then
                    RequestJObject.Replace('Content', '***');
            JObject.Add('Request', RequestJObject);
            DetailedInfoExists := true;
        end;
        if ResponseJObject.ReadFrom(ResponseJson) then begin
            if MaskContent then
                if ResponseJObject.SelectToken('Content', JToken) then
                    ResponseJObject.Replace('Content', '***');
            JObject.Add('Response', ResponseJObject);
            DetailedInfoExists := true;
        end;
        if DetailedInfoExists then
            if JObject.WriteTo(JsonString) then
                if JsonString <> '' then
                    ActivityLog.SetDetailedInfoFromText(JsonString);

        Commit(); // need to prevent rollback to save the log
    end;

    local procedure MaskHeaders(RequestJson: Text) Result: Text;
    var
        JObject: JsonObject;
        JToken: JsonToken;
        KeyValue: Text;
    begin
        Result := RequestJson;
        if JObject.ReadFrom(RequestJson) then
            if JObject.SelectToken('Header', JToken) then begin
                foreach KeyValue in JToken.AsObject().Keys() do
                    JToken.AsObject().Replace(KeyValue, '***');
                JObject.WriteTo(Result);
            end;
    end;

    // local procedure PutLargeDocumentOnSP(AccessToken: Text; DocumentStream: InStream; FileName: Text)
    // var
    //     URL: Text;
    //     vRequestContent: HttpContent;
    //     vContentHeaders: HttpHeaders;
    //     vHttpRequestMessage: HttpRequestMessage;
    //     // vHttpClient: HttpClient;
    //     vHttpResponseMessage: HttpResponseMessage;
    //     ResponseText: Text;
    //     Base64Convert: Codeunit "Base64 Convert";
    //     UploadURL: Text;
    // begin
    //     URL := 'https://graph.microsoft.com/v1.0/sites/endeavor.sharepoint.com,6b35f2d6-c98e-487d-8bd4-e01e9a909e1f,3fb50d38-3d8d-4e89-ae87-da95dc36fc1c/lists/28f66938-5f57-43a8-99f7-9aca6e539514/drive/root:/Business%20Documents/' + FileName + ':/createUploadSession';
    //     //URL := 'https://graph.microsoft.com/v1.0/sites/endeavor.sharepoint.com,6b35f2d6-c98e-487d-8bd4-e01e9a909e1f,3fb50d38-3d8d-4e89-ae87-da95dc36fc1c/lists/28f66938-5f57-43a8-99f7-9aca6e539514/drive/root:/Business%20Documents/' + FileName + ':/content';

    //     vHttpClient.Clear();
    //     vHttpClient.Timeout(60000);
    //     vHttpClient.DefaultRequestHeaders().Add('Authorization', StrSubstNo('Bearer %1', AccessToken));

    //     vHttpRequestMessage.Method := 'POST';
    //     vHttpRequestMessage.SetRequestUri(URL);
    //     vHttpClient.Send(vHttpRequestMessage, vHttpResponseMessage);

    //     vHttpResponseMessage.Content().ReadAs(ResponseText);

    //     ParseUploadSession(ResponseText, UploadURL);
    //     UploadLargeFile(AccessToken, UploadURL, DocumentStream)

    //     //Message(ResponseText);
    // end;


    // local procedure UploadLargeFile(AccessToken: Text; UploadURL: Text; DocumentStream: InStream)
    // var
    //     iStartByte: BigInteger;
    //     iEndByte: BigInteger;
    //     iByteCount: BigInteger;
    //     iBytesLeft: BigInteger;
    //     iLengthOfStream: BigInteger;

    //     DotNetStream: Codeunit DotNet_Stream;
    //     DotNetByteArray: Codeunit DotNet_Array;
    //     DotNetChunkStream: Codeunit DotNet_Stream;
    //     ChunkOutStream: OutStream;
    //     ChunkInStream: InStream;


    //     tempBlob: Codeunit "Temp Blob";
    // begin
    //     DotNetStream.FromInStream(DocumentStream);
    //     iLengthOfStream := DotNetStream.Length();
    //     DotNetByteArray.ByteArray(4000000);

    //     iByteCount := 1;
    //     iStartByte := 0;
    //     iBytesLeft := iLengthOfStream;
    //     while iBytesLeft > 0 do begin
    //         iByteCount := DotNetStream.Read(DotNetByteArray, iStartByte, 102400/*327680*/);
    //         tempBlob.CreateOutStream(ChunkOutStream);
    //         DotNetChunkStream.FromOutStream(ChunkOutStream);
    //         DotNetChunkStream.Write(DotNetByteArray, 0, iByteCount);
    //         tempBlob.CreateInStream(ChunkInStream);
    //         iEndByte := iStartByte + iByteCount - 1;

    //         if iByteCount > 0 then begin

    //             SendFileChunk(AccessToken, UploadURL, ChunkInStream, iByteCount, iStartByte, iEndByte, iLengthOfStream);
    //             DotNetChunkStream.Dispose();
    //             iStartByte := iEndByte + 1;
    //             iBytesLeft := iLengthOfStream - iByteCount;
    //         end;
    //     end;
    // end;

    // local procedure GetRequestForFileChunk(): HttpHeaders
    // begin

    // end;

    // local procedure SendFileChunk(AccessToken: Text; UploadURL: Text; ChunkInStream: InStream; ByteCount: BigInteger; StartByte: BigInteger; EndByte: BigInteger; LengthOfStream: BigInteger)
    // var
    //     Client: HttpClient;
    //     vRequestContent: HttpContent;
    //     vContentHeaders: HttpHeaders;
    //     vHttpRequestMessage: HttpRequestMessage;
    //     newHttpRequestMessage: HttpRequestMessage;
    //     vHttpResponseMessage: HttpResponseMessage;
    //     ResponseText: Text;
    // begin
    //     Clear(Client);
    //     Clear(vHttpRequestMessage);
    //     Clear(vHttpResponseMessage);
    //     Clear(vContentHeaders);
    //     Client.Clear();
    //     Client.Timeout(1000);
    //     vRequestContent.Clear();

    //     vRequestContent.WriteFrom(ChunkInStream);

    //     vRequestContent.GetHeaders(vContentHeaders);
    //     vContentHeaders.Clear();
    //     vContentHeaders.Add('Content-Type', 'application/octet-stream');
    //     vContentHeaders.Add('Content-Length', StrSubstNo('%1', ByteCount));
    //     vContentHeaders.Add('Content-Range', StrSubstNo('bytes %1-%2/%3', StartByte, EndByte, LengthOfStream));
    //     Client.DefaultRequestHeaders().Add('Authorization', StrSubstNo('Bearer %1', AccessToken));

    //     vHttpRequestMessage.Method := 'PUT';
    //     vHttpRequestMessage.SetRequestUri(UploadURL);

    //     vHttpRequestMessage.Content := vRequestContent;

    //     newHttpRequestMessage := vHttpRequestMessage;

    //     Client.Send(newHttpRequestMessage, vHttpResponseMessage);
    //     vHttpResponseMessage.Content().ReadAs(ResponseText);
    // end;

    var
        EnvironmentBlocksErr: Label 'Environment blocks an outgoing HTTP request to ''%1''.', Comment = '%1 - url, e.g. https://microsoft.com';
        ConnectionErr: Label 'Connection to the remote service ''%1'' could not be established.', Comment = '%1 - url, e.g. https://microsoft.com';
        CanNotConstructURL: Label 'Can not construct URL to Sharepoint. PDF not uploaded';
        vHttpClient: HttpClient;
        ActivityLogContextTxt: Label 'Sharepoint Setup', Locked = true;
}