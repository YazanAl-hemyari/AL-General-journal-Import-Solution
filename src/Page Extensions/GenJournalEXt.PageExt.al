pageextension 6000150 IBBGenJournalEXt extends "General Journal"
{
    layout
    {
        // Add changes to page layout here
    }

    actions
    {


        addfirst(Processing)
        {
            group(IBBImports)
            {
                Caption = 'Import';

                action("IBBImport")
                {
                    ApplicationArea = All;
                    Caption = 'Import Excel File';
                    ToolTip = 'Import Excel File';

                    Promoted = true;
                    PromotedCategory = Process;
                    PromotedIsBig = true;
                    PromotedOnly = true;
                    Image = ImportChartOfAccounts;



                    trigger OnAction()
                    var
                        ImportFile: Codeunit IBBExcelImportFile;
                    begin

                        // if rec."Journal Batch Name" <> '' then
                        //     Error(BatchISBlankMsg);
                        ImportFile.Run();


                    end;
                }
                action(IBBExportToExcel)
                {
                    Caption = 'Export to Excel';
                    ToolTip = 'Export to Excel';
                    ApplicationArea = All;
                    Promoted = true;
                    PromotedCategory = Process;
                    PromotedIsBig = true;
                    Image = Export;

                    trigger OnAction()
                    var
                        ExportFile: Codeunit IBBExcelImportFile;
                    begin
                        ;
                        ExportFile.ExportGenLine(Rec);
                    end;
                }
            }
        }




        // BatchISBlankMsg: Label 'Batch Name Can Not Be Blank';

    }
}
