report 90001 "SDH Export Multiple Sheets"
{
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = All;
    ProcessingOnly = true;
    Caption = 'Export Multiple Sheets To Excel';

    dataset
    {
        dataitem(CUstomer; Customer)
        {
            DataItemTableView = SORTING("No.") ORDER(Ascending);
            trigger OnPreDataItem()
            begin
                Customer.SetAutoCalcFields(Balance);
                CreateCustomerExcelHeader();
            end;

            trigger OnAfterGetRecord()
            begin
                CreateCustomerExcelBody();
            end;

            trigger OnPostDataItem()
            begin
                CreateExcelSheet(CopyStr(Customer.TableName, 1, 250), true);
            end;
        }
        dataitem(Vendor; Vendor)
        {
            DataItemTableView = SORTING("No.") ORDER(Ascending);
            trigger OnPreDataItem()
            begin
                Vendor.SetAutoCalcFields(Balance);
                CreateVendorExcelHeader();
            end;

            trigger OnAfterGetRecord()
            begin
                CreateVendorExcelBody();
            end;

            trigger OnPostDataItem()
            begin
                CreateExcelSheet(CopyStr(Vendor.TableName, 1, 250), false);
            end;
        }
    }

    trigger OnInitReport()
    begin
        TempExcelBuffer.DeleteAll();
    end;

    trigger OnPostReport()
    begin
        CreateExcelBook();
    end;

    local procedure CreateCustomerExcelHeader()
    begin
        TempExcelBuffer.AddColumn(Customer.FieldCaption("No."), False, '', True, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Customer.FieldCaption(Name), False, '', True, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Customer.FieldCaption("City"), False, '', True, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Customer.FieldCaption(Balance), False, '', True, false, false, '', TempExcelBuffer."Cell Type"::Text);
    end;

    local procedure CreateVendorExcelHeader()
    begin
        TempExcelBuffer.AddColumn(Vendor.FieldCaption("No."), False, '', True, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Vendor.FieldCaption(Name), False, '', True, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Vendor.FieldCaption("City"), False, '', True, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Vendor.FieldCaption(Balance), False, '', True, false, false, '', TempExcelBuffer."Cell Type"::Text);
    end;

    local procedure CreateCustomerExcelBody()
    begin
        TempExcelBuffer.NewRow();
        TempExcelBuffer.AddColumn(Customer."No.", False, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Customer.Name, False, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Customer.City, False, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Customer.Balance, False, '', false, false, false, '', TempExcelBuffer."Cell Type"::Number);
    end;

    local procedure CreateVendorExcelBody()
    begin
        TempExcelBuffer.NewRow();
        TempExcelBuffer.AddColumn(Vendor."No.", False, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Vendor.Name, False, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Vendor.City, False, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Vendor.Balance, False, '', false, false, false, '', TempExcelBuffer."Cell Type"::Number);
    end;

    local procedure CreateExcelBook()
    begin
        TempExcelBuffer.CloseBook();
        TempExcelBuffer.SetFriendlyFilename('Multi SHeet Export');
        TempExcelBuffer.OpenExcel();
    end;

    local procedure CreateExcelSheet(SheetName: Text[250]; NewBook: Boolean)
    begin
        if NewBook then
            TempExcelBuffer.CreateNewBook(SheetName)
        else
            TempExcelBuffer.SelectOrAddSheet(SheetName);
        TempExcelBuffer.WriteSheet(SheetName, CompanyName, UserId);
        TempExcelBuffer.DeleteAll();
        TempExcelBuffer.ClearNewRow();
    end;

    var
        TempExcelBuffer: Record "Excel Buffer" temporary;
}