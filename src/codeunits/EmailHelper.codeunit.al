codeunit 90000 "SDH Email Helper"
{

    procedure SendEmail(Recipients: Text[250]; CCRecipients: Text[250]; BccRecipients: Text[250]; EmailSubject: Text[250]; EmailScenario: Enum "Email Scenario"; AttachmentBlob: Codeunit "Temp Blob"; AttachmentFileName: Text)
    begin
        TempEmailItem."From Address" := GetSenderEmail(UserId);
        TempEmailItem."From Name" := GetSenderName(UserId);
        TempEmailItem."Send to" := Recipients;
        TempEmailItem."Send CC" := CCRecipients;
        TempEmailItem."Send BCC" := BccRecipients;
        TempEmailItem.Subject := EmailSubject;
        TempEmailItem.SetBodyText('Demo Email Body.');
        TempEmailItem.Validate("Plaintext Formatted", true);
        AddEmailAttachments(TempEmailItem, AttachmentBlob, AttachmentFileName);
        TempEmailItem.Send(true, EmailScenario);
    end;

    local procedure GetSenderEmail(SenderUserID: Code[50]): Text[250]
    var
        UserSetup: Record "User Setup";
    begin
        if UserSetup.Get(SenderUserID) then
            exit(UserSetup."E-Mail");
    end;

    local procedure GetSenderName(SenderUserID: Code[50]): Text[100]
    var
        User: Record User;
    begin
        User.SetRange("User Name", SenderUserID);
        if User.FindFirst and (User."Full Name" <> '') then
            exit(User."Full Name");

        exit('');
    end;

    local procedure AddEmailAttachments(var TempEmailItem: Record "Email Item" temporary; AttachmentBlob: Codeunit "Temp Blob"; AttachmentFileName: Text)
    var
        AttachmentStream: InStream;
    begin
        AttachmentBlob.CreateInStream(AttachmentStream);
        TempEmailItem.AddAttachment(AttachmentStream, AttachmentFileName);
    end;

    var
        TempEmailItem: Record "Email Item" temporary;
}