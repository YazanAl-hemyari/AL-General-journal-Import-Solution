tableextension 6000180 IBBGeneralJournal extends "Gen. Journal Line"
{
    fields
    {
        field(6000181; "SCC Shortcut Dimension 3 Code"; Code[20])
        {
            DataClassification = CustomerContent;
            CaptionClass = '1,2,3';
            Caption = 'Shortcut Dimension 3 Code';
            TableRelation = "Dimension Value".Code WHERE("Global Dimension No." = CONST(3),
                                                          Blocked = CONST(false));

            trigger OnValidate()
            begin
                ValidateShortcutDimCode(3, "SCC Shortcut Dimension 3 Code");
            end;
        }
        field(6000182; "PC Shortcut Dimension 4 Code"; Code[20])
        {
            DataClassification = CustomerContent;
            CaptionClass = '1,2,4';
            Caption = 'Shortcut Dimension 4 Code';
            TableRelation = "Dimension Value".Code WHERE("Global Dimension No." = CONST(4),
                                                          Blocked = CONST(false));

            trigger OnValidate()
            begin
                ValidateShortcutDimCode(4, "PC Shortcut Dimension 4 Code");
            end;
        }
        field(6000183; "AC Shortcut Dimension 5 Code"; Code[20])
        {
            DataClassification = CustomerContent;
            CaptionClass = '1,2,5';
            Caption = 'Shortcut Dimension 5 Code';
            TableRelation = "Dimension Value".Code WHERE("Global Dimension No." = CONST(5),
                                                          Blocked = CONST(false));

            trigger OnValidate()
            begin
                ValidateShortcutDimCode(5, "AC Shortcut Dimension 5 Code");
            end;
        }
        field(6000184; "SPC Shortcut Dimension 6 Code"; Code[20])
        {
            DataClassification = CustomerContent;
            CaptionClass = '1,2,6';
            Caption = 'Shortcut Dimension 6 Code';
            TableRelation = "Dimension Value".Code WHERE("Global Dimension No." = CONST(6),
                                                          Blocked = CONST(false));

            trigger OnValidate()
            begin
                ValidateShortcutDimCode(6, "SPC Shortcut Dimension 6 Code");
            end;
        }
        field(6000185; "IBB Shortcut Dimension 7 Code"; Code[20])
        {
            DataClassification = CustomerContent;
            CaptionClass = '1,2,7';
            Caption = 'Shortcut Dimension 7 Code';
            TableRelation = "Dimension Value".Code WHERE("Global Dimension No." = CONST(7),
                                                          Blocked = CONST(false));

            trigger OnValidate()
            begin
                ValidateShortcutDimCode(7, "IBB Shortcut Dimension 7 Code");
            end;
        }
        field(6000186; "IBB Shortcut Dimension 8 Code"; Code[20])
        {
            DataClassification = CustomerContent;
            CaptionClass = '1,2,8';
            Caption = 'Shortcut Dimension 8 Code';
            TableRelation = "Dimension Value".Code WHERE("Global Dimension No." = CONST(8),
                                                          Blocked = CONST(false));

            trigger OnValidate()
            begin
                ValidateShortcutDimCode(8, "IBB Shortcut Dimension 8 Code");
            end;
        }
    }


}