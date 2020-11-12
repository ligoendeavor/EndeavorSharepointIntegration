//#pragma implicitwith disable
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
                Visible = ShowSharepointAction;

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
                    DocumentRef.Get(Rec.RecordId);
                    FieldRef := DocumentRef.FIELD(Rec.FieldNo("No."));
                    FieldRef.SETFILTER(Rec."No.");

                    tempBlob.CreateOutStream(outStreamReport, TextEncoding::UTF8);
                    Report.SaveAs(Report::"Standard Sales - Invoice", '', ReportFormat::Pdf, outStreamReport, DocumentRef);

                    SPMgmt.GetAccessToken(AccessToken);
                    tempBlob.CreateInStream(DocumentStream);
                    if Confirm(UploadQuestion) then
                        SPMgmt.PutDocumentOnSP(AccessToken, Rec.RecordId(), DocumentStream, StrSubstNo('%1.pdf', Rec."No."), Rec."EDX09 Sharepoint Site", Rec."EDX09 Sharepoint Doc Library", Rec."EDX09 Sharepoint Full Url");
                end;
            }
        }
    }

    trigger OnOpenPage()
    var
        SPIntMgmt: Codeunit "EDX09 Sharepoint Int. Mgmt.";
    begin
        ShowSharepointAction := SPIntMgmt.IsSharepointIntegrationEnabled();
    end;

    var
        UploadQuestion: Label 'Do you want to upload PDF to Sharepoint? Existing file will be replaced.';
        ShowSharepointAction: Boolean;
}
//#pragma implicitwith restore
