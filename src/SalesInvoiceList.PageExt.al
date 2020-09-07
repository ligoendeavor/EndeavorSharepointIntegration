pageextension 87000 "EDX09 PostedSalesInvoicesExt" extends "Posted Sales Invoices"
{
    layout
    {

    }

    actions
    {
        addlast(processing)
        {
            action(UploadPDFToSP)
            {
                ApplicationArea = All;
                Caption = 'Upload PDF to Sharepoint';
                Image = SendAsPDF;

                trigger OnAction()
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
                    SharepointURL: Text;
                    SPMgmt: Codeunit "EDX09 Sharepoint Int. Mgmt.";
                begin
                    DocumentRef.Get(RecordId);
                    FieldRef := DocumentRef.FIELD(FieldNo("No."));
                    FieldRef.SETFILTER("No.");

                    tempBlob.CreateOutStream(outStreamReport, TextEncoding::UTF8);

                    Report.SaveAs(Report::"Standard Sales - Invoice", '', ReportFormat::Pdf, outStreamReport, DocumentRef);
                    //Report.SaveAs(Report::"Standard Sales - Invoice", Parameters, ReportFormat::Pdf, outStreamReport);

                    SPMgmt.GetAccessToken(AccessToken);
                    tempBlob.CreateInStream(DocumentStream);
                    if Confirm(UploadQuestion) then
                        SPMgmt.PutDocumentOnSP(AccessToken, DocumentStream, StrSubstNo('%1.pdf', "No."), "EDXpm Sharepoint Site", "EDXpm Sharepoint Doc Library", "EDXpm Sharepoint Full Url");
                end;
            }
        }
    }

    var
        UploadQuestion: Label 'Do you want to upload PDF to Sharepoint? Existing file will be replaced.';
}