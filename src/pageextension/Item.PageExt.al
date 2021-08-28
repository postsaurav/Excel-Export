pageextension 90000 "SDH Item" extends "Item List"
{
    actions
    {
        addfirst(processing)
        {
            action(ExportExcel)
            {
                ApplicationArea = All;
                Caption = 'Export To Excel';
                Promoted = true;
                PromotedCategory = Process;
                ToolTip = 'Exports Item List to Excel';
                Image = Excel;
                trigger OnAction()
                var
                    ExportExcel: Report "SDH Export Single Sheet";
                begin
                    ExportExcel.RunModal();
                end;
            }
        }
    }
}