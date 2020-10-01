tableextension 87000 "EDX09 SalesHeaderExt" extends "Sales Header"
{
    fields
    {
        field(87000; "EDX09 Sharepoint Site"; Text[256])
        {
            Caption = 'Sharepoint Site';
        }
        field(87001; "EDX09 Sharepoint Doc Library"; Text[100])
        {
            Caption = 'Sharepoint Document Library';
        }
        field(87002; "EDX09 Sharepoint Full Url"; Text[256])
        {
            Caption = 'Sharepoint Full Relative Url';
        }
        field(87003; "EDX09 Sharepoint URL"; Text[256])
        {
            Caption = 'Sharepoint URL';
        }
    }

}