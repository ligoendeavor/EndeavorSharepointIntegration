codeunit 87000 "EDX09 Sharepoint Int. Events"
{

    [EventSubscriber(ObjectType::Codeunit, Codeunit::"Sales-Post", 'OnAfterPostSalesDoc', '', false, false)]
    local procedure OnAfterPostSalesDoc(var SalesHeader: Record "Sales Header"; var GenJnlPostLine: Codeunit "Gen. Jnl.-Post Line"; SalesShptHdrNo: Code[20]; RetRcpHdrNo: Code[20]; SalesInvHdrNo: Code[20]; SalesCrMemoHdrNo: Code[20]; CommitIsSuppressed: Boolean; InvtPickPutaway: Boolean)
    var
        outStreamReport: OutStream;
        inStreamReport: InStream;
        Parameters: Text;
        tempBlob: Codeunit "Temp Blob";
        Base64EncodedString: Text;
        AccessToken: Text;
        DocumentStream: InStream;
        DocumentRef: RecordRef;
        FieldRef: FieldRef;
        SalesInvoice: Record "Sales Invoice Header";
        SharepointURL: Text;
        SPMgmt: Codeunit "EDX09 Sharepoint Int. Mgmt.";
    begin
        exit;

        if SalesInvHdrNo <> '' then begin
            SalesInvoice.SetRange("No.", SalesInvHdrNo);
            if SalesInvoice.findset then begin
                DocumentRef.Get(SalesInvoice.RecordId);
                FieldRef := DocumentRef.FIELD(SalesInvoice.FieldNo("No."));
                FieldRef.SETFILTER(SalesInvHdrNo);

                tempBlob.CreateOutStream(outStreamReport, TextEncoding::UTF8);
                Report.SaveAs(Report::"Standard Sales - Invoice", '', ReportFormat::Pdf, outStreamReport, DocumentRef);

                SPMgmt.GetAccessToken(AccessToken);
                tempBlob.CreateInStream(DocumentStream);
                SPMgmt.PutDocumentOnSP(AccessToken, DocumentStream, StrSubstNo('%1.pdf', SalesInvHdrNo), SalesInvoice."EDXpm Sharepoint Site", SalesInvoice."EDXpm Sharepoint Doc Library", SalesInvoice."EDXpm Sharepoint Full Url");
            end;
        end;
    end;

    // [EventSubscriber(ObjectType::Codeunit, Codeunit::ReportManagement, 'OnAfterDocumentPrintReady', '', false, false)]
    // local procedure OnAfterDocumentPrintReady(ObjectType: Option "Report","Page"; ObjectId: Integer; ObjectPayload: JsonObject; DocumentStream: InStream; var Success: Boolean)
    // var
    //     ObjectNameToken: JsonToken;
    //     RequestJsonContent: JsonObject;
    //     RequestUrlContent: Text;
    //     ResponseJson: Text;
    //     RequestJson: Text;
    //     HttpError: Text;
    //     UrlString: Text;
    //     JObject: JsonObject;
    //     JsonString: Text;
    //     OAuthSetup: Record "OAuth 2.0 Setup";
    //     AccessToken: Text;
    //     RefreshToken: Text;
    //     ExpireInSec: BigInteger;
    //     ObjectName: Text;
    //     DocumentTypeToken: JsonToken;
    //     DocumentTypeParts: List of [Text];
    //     FileExtension: Text;
    //     tempBlob: Codeunit "Temp Blob";
    //     iStreamLength: BigInteger;
    //     BinaryReader: codeunit DotNet_BinaryReader;
    //     DotNetStream: Codeunit DotNet_Stream;
    // begin
    //     if Success then
    //         exit;

    //     if (ObjectType = ObjectType::Report) then begin

    //         ObjectPayload.Get('objectname', ObjectNameToken);
    //         ObjectPayload.Get('objectname', ObjectNameToken);
    //         ObjectName := ObjectNameToken.AsValue().AsText();
    //         ObjectPayload.Get('documenttype', DocumentTypeToken);

    //         // Step 4: Build the email message
    //         DocumentTypeParts := DocumentTypeToken.AsValue().AsText().Split('/');
    //         FileExtension := DocumentTypeParts.Get(DocumentTypeParts.Count);

    //         OAuthSetup.Get();
    //         CreateContentRequestForAccessToken(RequestUrlContent, OAuthSetup."EDXtmp Secret", OAuthSetup."Client ID", OAuthSetup."Redirect URL", OAuthSetup.Scope);
    //         CreateRequestJSONForAccessRefreshTokenUrlEncoded(RequestJson, OAuthSetup."Service URL", OAuthSetup."Access Token URL Path", RequestUrlContent);

    //         RequestAccessAndRefreshTokens(RequestJson, ResponseJson, AccessToken, RefreshToken, ExpireInSec, HttpError);
    //         //message(AccessToken);

    //         // Check if stream is larger than 4MB (chunk size dividable by 327680)
    //         DotNetStream.FromInStream(DocumentStream);
    //         iStreamLength := DotNetStream.Length();

    //         // if iStreamLength > 4000000 then
    //         //PutLargeDocumentOnSP(AccessToken, DocumentStream, ObjectName + '.' + FileExtension)
    //         // else
    //         PutDocumentOnSP(AccessToken, DocumentStream, ObjectName + '.' + FileExtension);
    //     end;
    // end;


}