codeunit 6000150 IBBExcelImportFile
{
    trigger OnRun()
    begin

        IBBReadExcelFile();
        ImportExcelData();
        //ExportGenLine(GenJournalLine);

    end;

    var


        TempExcelBuffer: Record "Excel Buffer" temporary;
        GenJournalLine: Record "Gen. Journal Line";
        //NameValueBufferOut: Record "Name/Value Buffer";
        FileManagement: Codeunit "File Management";
        FileName: Text;
        SheetName: Text;
        Instream: InStream;
        FromFile: Text;
        RowNo: Integer;
        //ColNo: Integer;
        LineNo: Integer;
        MaxRowNo: Integer;








        UploadExcelMsg: Label 'Please Choose a File to Import it ..';
        NoFileFoundMsg: Label 'No Excel File Was Detected !..';

        ImpSucessedMsg: Label 'Excel is Successfully Imported ..';


    procedure IBBReadExcelFile()
    begin

        UploadIntoStream(UploadExcelMsg, '', '', FromFile, Instream);

        if FromFile <> '' then begin
            FileName := FileManagement.GetFileName(FromFile);
            SheetName := TempExcelBuffer.SelectSheetsNameStream(Instream);
        end else
            Error(NoFileFoundMsg);

        TempExcelBuffer.Reset();
        TempExcelBuffer.DeleteAll();
        TempExcelBuffer.OpenBookStream(Instream, SheetName);
        TempExcelBuffer.ReadSheet();
        /*     IF NameValueBufferOut.FINDSET() THEN begin
                ;
                REPEAT
                    CLEAR(SheetName);
                    SheetName := NameValueBufferOut.Value;
                    TempExcelBuffer.OpenBookStream(Instream, SheetName);
                    TempExcelBuffer.ReadSheet();

                    MaxRowNo := 0;
                    TempExcelBuffer.RESET();
                    IF TempExcelBuffer.FindLast() THEN
                        MaxRowNo := TempExcelBuffer."Row No.";

                    //RowNo := 0;
                    //ColNo := 0;

                    If SheetName = 'Journal Line' THEN BEGIN
                        ;
                        ImportExcelData();
                    END;
                UNTIL NameValueBufferOut.NEXT() = 0;
            END; */
        Message('Total Count %1', TempExcelBuffer.Count);

    end;

    procedure ImportExcelData()
    var


    begin
        RowNo := 0;
        MaxRowNo := 0;
        LineNo := 0;
        GenJournalLine.Reset();
        if GenJournalLine.FindLast() then
            LineNo := GenJournalLine."Line No.";
        TempExcelBuffer.Reset();

        if TempExcelBuffer.FindLast() then begin
            ;
            MaxRowNo := TempExcelBuffer."Row No.";

        end;

        for RowNo := 2 to MaxRowNo do begin
            LineNo := LineNo + 10000;
            GenJournalLine.Init();
            GenJournalLine."Line No." := LineNo;
            Evaluate(GenJournalLine."Journal Batch Name", GetValueAtCell(RowNo, 1));
            Evaluate(GenJournalLine."Posting Date", GetValueAtCell(RowNo, 2));
            Evaluate(GenJournalLine."Document Type", GetValueAtCell(RowNo, 3));
            Evaluate(GenJournalLine."Document No.", GetValueAtCell(RowNo, 4));
            Evaluate(GenJournalLine."Account Type", GetValueAtCell(RowNo, 5));

            Evaluate(GenJournalLine."Account No.", GetValueAtCell(RowNo, 6));
            GenJournalLine.Validate(GenJournalLine."Account No.");

            Evaluate(GenJournalLine.Description, GetValueAtCell(RowNo, 7));

            Evaluate(GenJournalLine."Currency Code", GetValueAtCell(RowNo, 8));
            GenJournalLine.Validate(GenJournalLine."Currency Code");

            Evaluate(GenJournalLine."Gen. Posting Type", GetValueAtCell(RowNo, 9));
            GenJournalLine.Validate(GenJournalLine."Gen. Posting Type");

            Evaluate(GenJournalLine."Gen. Bus. Posting Group", GetValueAtCell(RowNo, 10));
            GenJournalLine.Validate(GenJournalLine."Gen. Bus. Posting Group");

            Evaluate(GenJournalLine."Gen. Prod. Posting Group", GetValueAtCell(RowNo, 11));
            GenJournalLine.Validate(GenJournalLine."Gen. Prod. Posting Group");

            Evaluate(GenJournalLine."Tax Liable", GetValueAtCell(RowNo, 12));

            Evaluate(GenJournalLine."Tax Area Code", GetValueAtCell(RowNo, 13));
            GenJournalLine.Validate(GenJournalLine."Tax Area Code");

            Evaluate(GenJournalLine."Tax Group Code", GetValueAtCell(RowNo, 14));
            GenJournalLine.Validate(GenJournalLine."Tax Group Code");

            Evaluate(GenJournalLine.Amount, GetValueAtCell(RowNo, 15));
            GenJournalLine.Validate(GenJournalLine.Amount);

            Evaluate(GenJournalLine."Bal. Account Type", GetValueAtCell(RowNo, 16));
            GenJournalLine.Validate(GenJournalLine."Bal. Account Type");

            Evaluate(GenJournalLine."Bal. Account No.", GetValueAtCell(RowNo, 17));
            GenJournalLine.Validate(GenJournalLine."Bal. Account No.");

            Evaluate(GenJournalLine."Bal. Gen. Posting Type", GetValueAtCell(RowNo, 18));
            GenJournalLine.Validate(GenJournalLine."Bal. Gen. Posting Type");

            Evaluate(GenJournalLine."Bal. Gen. Bus. Posting Group", GetValueAtCell(RowNo, 19));
            GenJournalLine.Validate(GenJournalLine."Bal. Gen. Bus. Posting Group");

            Evaluate(GenJournalLine."Bal. Gen. Prod. Posting Group", GetValueAtCell(RowNo, 20));
            GenJournalLine.Validate(GenJournalLine."Bal. Gen. Prod. Posting Group");

            Evaluate(GenJournalLine."Deferral Code", GetValueAtCell(RowNo, 21));
            GenJournalLine.Validate(GenJournalLine."Deferral Code");

            Evaluate(GenJournalLine.Correction, GetValueAtCell(RowNo, 22));
            Evaluate(GenJournalLine.Comment, GetValueAtCell(RowNo, 23));

            Evaluate(GenJournalLine."Shortcut Dimension 1 Code", GetValueAtCell(RowNo, 24));
            GenJournalLine.Validate(GenJournalLine."Shortcut Dimension 1 Code");

            Evaluate(GenJournalLine."Shortcut Dimension 2 Code", GetValueAtCell(RowNo, 25));
            GenJournalLine.Validate(GenJournalLine."Shortcut Dimension 2 Code");

            Evaluate(GenJournalLine."SCC Shortcut Dimension 3 Code", GetValueAtCell(RowNo, 26));
            GenJournalLine.Validate(GenJournalLine."SCC Shortcut Dimension 3 Code");

            Evaluate(GenJournalLine."PC Shortcut Dimension 4 Code", GetValueAtCell(RowNo, 27));
            GenJournalLine.Validate(GenJournalLine."PC Shortcut Dimension 4 Code");

            Evaluate(GenJournalLine."AC Shortcut Dimension 5 Code", GetValueAtCell(RowNo, 28));
            GenJournalLine.Validate(GenJournalLine."AC Shortcut Dimension 5 Code");

            Evaluate(GenJournalLine."SPC Shortcut Dimension 6 Code", GetValueAtCell(RowNo, 29));
            GenJournalLine.Validate(GenJournalLine."SPC Shortcut Dimension 6 Code");

            Evaluate(GenJournalLine."IBB Shortcut Dimension 7 Code", GetValueAtCell(RowNo, 29));
            GenJournalLine.Validate(GenJournalLine."SPC Shortcut Dimension 6 Code");

            Evaluate(GenJournalLine."IBB Shortcut Dimension 8 Code", GetValueAtCell(RowNo, 29));
            GenJournalLine.Validate(GenJournalLine."SPC Shortcut Dimension 6 Code");

            GenJournalLine.Insert();

        end;
        Message(ImpSucessedMsg);
    end;

    procedure ExportGenLine(var GenJournalLine: Record "Gen. Journal Line")
    var
        CustLedgerEntriesLbl: Label 'General Journal Line';
        ExcelFileNameLbl: Label 'Gen. Journal Line_%1_%2', Comment = '%1 = XML node name ; %2 = Parent XML node name';
    begin
        TempExcelBuffer.Reset();
        TempExcelBuffer.DeleteAll();
        TempExcelBuffer.NewRow();
        TempExcelBuffer.AddColumn(GenJournalLine.FieldCaption("Journal Batch Name"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GenJournalLine.FieldCaption("Posting Date"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GenJournalLine.FieldCaption("Document Type"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GenJournalLine.FieldCaption("Document No."), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GenJournalLine.FieldCaption("Account Type"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GenJournalLine.FieldCaption("Account No."), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GenJournalLine.FieldCaption(Description), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GenJournalLine.FieldCaption("Currency Code"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GenJournalLine.FieldCaption("Gen. Posting Type"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GenJournalLine.FieldCaption("Gen. Bus. Posting Group"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GenJournalLine.FieldCaption("Gen. Prod. Posting Group"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GenJournalLine.FieldCaption("Tax Liable"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GenJournalLine.FieldCaption("Tax Area Code"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GenJournalLine.FieldCaption("Tax Group Code"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GenJournalLine.FieldCaption(Amount), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GenJournalLine.FieldCaption("Bal. Account Type"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GenJournalLine.FieldCaption("Bal. Account No."), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GenJournalLine.FieldCaption("Bal. Gen. Posting Type"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GenJournalLine.FieldCaption("Bal. Gen. Bus. Posting Group"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GenJournalLine.FieldCaption("Bal. Gen. Prod. Posting Group"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GenJournalLine.FieldCaption("Deferral Code"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GenJournalLine.FieldCaption(Correction), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GenJournalLine.FieldCaption(Comment), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GenJournalLine.FieldCaption("Shortcut Dimension 1 Code"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GenJournalLine.FieldCaption("Shortcut Dimension 2 Code"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GenJournalLine.FieldCaption("SCC Shortcut Dimension 3 Code"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GenJournalLine.FieldCaption("PC Shortcut Dimension 4 Code"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GenJournalLine.FieldCaption("AC Shortcut Dimension 5 Code"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GenJournalLine.FieldCaption("SPC Shortcut Dimension 6 Code"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GenJournalLine.FieldCaption("IBB Shortcut Dimension 7 Code"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
        TempExcelBuffer.AddColumn(GenJournalLine.FieldCaption("IBB Shortcut Dimension 8 Code"), false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);


        if GenJournalLine.FindSet() then
            repeat
                TempExcelBuffer.NewRow();
                TempExcelBuffer.AddColumn(GenJournalLine."Journal Batch Name", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);

                TempExcelBuffer.AddColumn(GenJournalLine."Posting Date", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Date);
                TempExcelBuffer.AddColumn(GenJournalLine."Document Type", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(GenJournalLine."Document No.", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(GenJournalLine."Account Type", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(GenJournalLine."Account No.", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(GenJournalLine.Description, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(GenJournalLine."Currency Code", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(GenJournalLine."Gen. Posting Type", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(GenJournalLine."Gen. Bus. Posting Group", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(GenJournalLine."Gen. Prod. Posting Group", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(GenJournalLine."Tax Liable", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(GenJournalLine."Tax Area Code", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(GenJournalLine."Tax Group Code", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(GenJournalLine.Amount, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Number);
                TempExcelBuffer.AddColumn(GenJournalLine."Bal. Account Type", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(GenJournalLine."Bal. Account No.", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(GenJournalLine."Bal. Gen. Posting Type", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(GenJournalLine."Bal. Gen. Bus. Posting Group", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(GenJournalLine."Bal. Gen. Prod. Posting Group", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(GenJournalLine."Deferral Code", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(GenJournalLine.Correction, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(GenJournalLine.Comment, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(GenJournalLine."Shortcut Dimension 1 Code", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(GenJournalLine."Shortcut Dimension 2 Code", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(GenJournalLine."SCC Shortcut Dimension 3 Code", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(GenJournalLine."PC Shortcut Dimension 4 Code", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(GenJournalLine."AC Shortcut Dimension 5 Code", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(GenJournalLine."SPC Shortcut Dimension 6 Code", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(GenJournalLine."IBB Shortcut Dimension 7 Code", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
                TempExcelBuffer.AddColumn(GenJournalLine."IBB Shortcut Dimension 8 Code", false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
            until GenJournalLine.Next() = 0;
        TempExcelBuffer.CreateNewBook(CustLedgerEntriesLbl);
        TempExcelBuffer.WriteSheet(CustLedgerEntriesLbl, CompanyName, UserId);
        TempExcelBuffer.CloseBook();
        TempExcelBuffer.SetFriendlyFilename(StrSubstNo(ExcelFileNameLbl, CurrentDateTime, UserId));
        TempExcelBuffer.OpenExcel();
    end;


    procedure GetValueAtCell(RowNo: Integer; ColNo: Integer): Text
    begin

        TempExcelBuffer.Reset();
        If TempExcelBuffer.Get(RowNo, ColNo) then
            exit(TempExcelBuffer."Cell Value as Text")
        else
            exit('');
    end;


}