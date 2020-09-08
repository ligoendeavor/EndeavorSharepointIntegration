pageextension 87000 "EDX09 PostedSalesInvoicesExt" extends "Posted Sales Invoices"
{
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
                    tempBlob: Codeunit "Temp Blob";
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

                    SPMgmt.GetAccessToken(AccessToken);
                    tempBlob.CreateInStream(DocumentStream);
                    if Confirm(UploadQuestion) then
                        SPMgmt.PutDocumentOnSP(AccessToken, RecordId(), DocumentStream, StrSubstNo('%1.pdf', "No."), "EDX Sharepoint Site", "EDX Sharepoint Doc Library", "EDX Sharepoint Full Url");
                end;
            }
        }
    }

    var
        UploadQuestion: Label 'Do you want to upload PDF to Sharepoint? Existing file will be replaced.';
}