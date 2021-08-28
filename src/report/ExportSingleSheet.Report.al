report 90000 "SDH Export Single Sheet"
{
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = All;
    ProcessingOnly = true;
    Caption = 'Export Item To Excel';

    dataset
    {
        dataitem(Item; Item)
        {
            DataItemTableView = SORTING("No.") ORDER(Ascending);
            trigger OnPreDataItem()
            begin
                Item.SetAutoCalcFields(Inventory);
                CreateExcelHeader();
            end;

            trigger OnAfterGetRecord()
            begin
                CreateExcelBody();
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

    local procedure CreateExcelHeader()
    begin
        TempExcelBuffer.AddColumn(Item.FieldCaption("No."), False, '', True, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Item.FieldCaption(Description), False, '', True, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Item.FieldCaption("Base Unit of Measure"), False, '', True, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Item.FieldCaption(Inventory), False, '', True, false, false, '', TempExcelBuffer."Cell Type"::Text);
    end;

    local procedure CreateExcelBody()
    begin
        TempExcelBuffer.NewRow();
        TempExcelBuffer.AddColumn(Item."No.", False, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Item.Description, False, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Item."Base Unit of Measure", False, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(Item.Inventory, False, '', false, false, false, '', TempExcelBuffer."Cell Type"::Number);
    end;

    local procedure CreateExcelBook()
    begin
        TempExcelBuffer.CreateNewBook('Item List');
        TempExcelBuffer.WriteSheet('Item List', CompanyName, UserId);
        TempExcelBuffer.CloseBook();
        TempExcelBuffer.SetFriendlyFilename('Item Export');
        TempExcelBuffer.OpenExcel();
    end;

    var
        TempExcelBuffer: Record "Excel Buffer" temporary;
}