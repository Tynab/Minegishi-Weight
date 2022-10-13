Imports Microsoft.Office.Interop.Excel

Friend Module Util
    ''' <summary>
    ''' 運賃 (2トン車).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    Friend Sub Fare(xlApp As Application)
        PubDModVal(xlApp, "78", "D10", "400×200", 0.4, 15)
    End Sub

    ''' <summary>
    ''' 外周.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    Friend Sub Circumference(xlApp As Application)
        PubDVal(xlApp, "AG9", DtlDInp(vbTab & "40: "))
        PubDVal(xlApp, "AG10", DtlDInp(vbTab & "36: "))
        PubDVal(xlApp, "AG11", DtlDInp(vbTab & "31: "))
        PubDVal(xlApp, "AG12", DtlDInp(vbTab & "27: "))
        PubDVal(xlApp, "AG13", DtlDInp(vbTab & "22: "))
        PubDVal(xlApp, "AG14", DtlDInp(vbTab & "18: "))
        PubDVal(xlApp, "AG15", DtlDInp(vbTab & "13: "))
        PubDVal(xlApp, "AG16", DtlDInp(vbTab & " 9: "))
        PubDVal(xlApp, "AG17", DtlDInp(vbTab & " 4: "))
    End Sub

    ''' <summary>
    ''' 内周.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    Friend Sub InnerCircumference(xlApp As Application)
        PubDVal(xlApp, "AG18", DtlDInp(vbTab & "40: "))
        PubDVal(xlApp, "AG22", DtlDInp(vbTab & "36: "))
        PubDVal(xlApp, "AG24", DtlDInp(vbTab & "31: "))
        PubDVal(xlApp, "AG26", DtlDInp(vbTab & "27: "))
        PubDVal(xlApp, "AG28", DtlDInp(vbTab & "22: "))
        PubDVal(xlApp, "AG30", DtlDInp(vbTab & "18: "))
        PubDVal(xlApp, "AG32", DtlDInp(vbTab & "13: "))
        PubDVal(xlApp, "AG34", DtlDInp(vbTab & " 9: "))
        PubDVal(xlApp, "AG37", DtlDInp(vbTab & " 4: "))
    End Sub

    ''' <summary>
    ''' ストレート (D13) .
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    Friend Sub StrD13(xlApp As Application)
        PubDVal(xlApp, "AG59", DtlDInp(vbTab & "L-2500: "))
        PubDVal(xlApp, "AG55", DtlDInp(vbTab & "L-2000: "))
        PubDVal(xlApp, "AG58", DtlDInp(vbTab & "L-1500: "))
        PubDVal(xlApp, "AG56", DtlDInp(vbTab & "L-1200: "))
    End Sub

    ''' <summary>
    ''' ストレート (D10) .
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub StrD10(xlApp As Application, choosen As Double)
        If choosen = 1 Then
            PubDVal(xlApp, "AG76", DtlDInp(vbTab & "L-2000: "))
            PubDVal(xlApp, "AG77", DtlDInp(vbTab & "L-1500: "))
        End If
    End Sub

    ''' <summary>
    ''' コーナー (D13).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    Friend Sub CorD13(xlApp As Application)
        PubDModVal(xlApp, "61", "600×2420", 3.2, DtlDInp(vbTab & "600×2420: "))
        PubDVal(xlApp, "AG60", DtlDInp(vbTab & "600×2400: "))
        PubDModVal(xlApp, "62", "600×2000", 2.7, DtlDInp(vbTab & "600×2000: "))
        PubDVal(xlApp, "AG63", DtlDInp(vbTab & "600×1900: "))
        PubDModVal(xlApp, "67", "600×1050", 1.7, DtlDInp(vbTab & "600×1050: "))
        PubDVal(xlApp, "AG57", DtlDInp(vbTab & "600× 600: "))
        PubDVal(xlApp, "AG72", DtlDInp(vbTab & "455× 600: "))
        PubDModVal(xlApp, "71", "350×1650", 2.1, DtlDInp(vbTab & "350×1650: "))
    End Sub

    ''' <summary>
    ''' クランク.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub Crank(xlApp As Application, choosen As Double)
        If choosen = 1 Then
            PubDModVal(xlApp, "73", "D16", "750×460×750", 3.2, DtlDInp(vbTab & "D16 (750×460×750): "))
            PubDModVal(xlApp, "68", "600×460×600", 1.8, DtlDInp(vbTab & "D13 (600×460×600): "))
            PubDModVal(xlApp, "69", "500×460×500", 0.9, DtlDInp(vbTab & "D10 (500×460×500): "))
        End If
    End Sub

    ''' <summary>
    ''' コ型.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub UType(xlApp As Application, choosen As Double)
        If choosen = 1 Then
            PubDModVal(xlApp, "66", "D16", "750×460×750", 3.2, DtlDInp(vbTab & "D16 (750× 460×750): "))
            PubDVal(xlApp, "AG65", DtlDInp(vbTab & "D13 (600×1835×600): "))
            PubDVal(xlApp, "AG64", DtlDInp(vbTab & "D13 (600×1380×600): "))
            PubDModVal(xlApp, "70", "600×460×600", 1.8, DtlDInp(vbTab & "D13 (600× 460×600): "))
        End If
    End Sub

    ''' <summary>
    ''' フック (D10).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub Hook(xlApp As Application, choosen As Double)
        If choosen = 1 Then
            PubDModVal(xlApp, "81", "405×200", 0.4, "（立筋補強）", DtlDInp(vbTab & "405×200: "))
            PubDVal(xlApp, "AG80", DtlDInp(vbTab & "400×200: "))
            PubDVal(xlApp, "AG79", DtlDInp(vbTab & "390×200: "))
        End If
    End Sub

    ''' <summary>
    ''' スラブ曲 (D10).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="truck2Ton">2 ton truck.</param>
    Friend Sub SlabBndg(xlApp As Application, truck2Ton As Double)
        If Not truck2Ton = 1 Then
            PubDVal(xlApp, "AG82", DtlDInp(vbTab & "400×5600: "))
            PubDVal(xlApp, "AG83", DtlDInp(vbTab & "400×5100: "))
        End If
        PubDVal(xlApp, "AG84", DtlDInp(vbTab & "400×4600: "))
        PubDVal(xlApp, "AG85", DtlDInp(vbTab & "400×4100: "))
        PubDVal(xlApp, "AG86", DtlDInp(vbTab & "400×3600: "))
        PubDVal(xlApp, "AG87", DtlDInp(vbTab & "400×3100: "))
        PubDVal(xlApp, "AG88", DtlDInp(vbTab & "400×2600: "))
        PubDVal(xlApp, "AG89", DtlDInp(vbTab & "400×2100: "))
        PubDVal(xlApp, "AG94", DtlDInp(vbTab & "400×1820: "))
        PubDVal(xlApp, "AG90", DtlDInp(vbTab & "400×1600: "))
        PubDModVal(xlApp, "93", "400×1360", 1.1, DtlDInp(vbTab & "400×1360: "))
        PubDVal(xlApp, "AG91", DtlDInp(vbTab & "400×1100: "))
        PubDVal(xlApp, "AG92", DtlDInp(vbTab & "400× 600: "))
    End Sub

    ''' <summary>
    ''' スラブ直 (D10).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    ''' <param name="truck2Ton">2 ton truck.</param>
    Friend Sub SlabStr(xlApp As Application, choosen As Double, truck2Ton As Double)
        If choosen = 1 Then
            If Not truck2Ton = 1 Then
                PubDVal(xlApp, "AG113", DtlDInp(vbTab & "6000: "))
                PubDVal(xlApp, "AG114", DtlDInp(vbTab & "5500: "))
                PubDVal(xlApp, "AG115", DtlDInp(vbTab & "5000: "))
            End If
            PubDVal(xlApp, "AG116", DtlDInp(vbTab & "4500: "))
            PubDVal(xlApp, "AG117", DtlDInp(vbTab & "4000: "))
            PubDVal(xlApp, "AG118", DtlDInp(vbTab & "3500: "))
            PubDVal(xlApp, "AG119", DtlDInp(vbTab & "3000: "))
            PubDVal(xlApp, "AG120", DtlDInp(vbTab & "2500: "))
            PubDVal(xlApp, "AG121", DtlDInp(vbTab & "2000: "))
            PubDVal(xlApp, "AG122", DtlDInp(vbTab & "1500: "))
            PubDVal(xlApp, "AG123", DtlDInp(vbTab & "1000: "))
        End If
    End Sub

    ''' <summary>
    ''' スラブ補強曲 (D13).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    ''' <param name="truck2Ton">2 ton truck.</param>
    Friend Sub SlabReinfBndg(xlApp As Application, choosen As Double, truck2Ton As Double)
        If choosen = 1 Then
            If Not truck2Ton = 1 Then
                PubDModVal(xlApp, "125", "（スラブ補強）", DtlDInp(vbTab & "520×5480: "))
                PubDModVal(xlApp, "126", "（スラブ補強）", DtlDInp(vbTab & "520×4980: "))
            End If
            PubDModVal(xlApp, "127", "（スラブ補強）", DtlDInp(vbTab & "520×4480: "))
            PubDModVal(xlApp, "128", "（スラブ補強）", DtlDInp(vbTab & "520×3980: "))
            PubDModVal(xlApp, "129", "（スラブ補強）", DtlDInp(vbTab & "520×3480: "))
            PubDModVal(xlApp, "130", "（スラブ補強）", DtlDInp(vbTab & "520×2980: "))
            PubDModVal(xlApp, "131", "（スラブ補強）", DtlDInp(vbTab & "520×2480: "))
            PubDModVal(xlApp, "132", "（スラブ補強）", DtlDInp(vbTab & "520×1980: "))
            PubDModVal(xlApp, "133", "（スラブ補強）", DtlDInp(vbTab & "520×1480: "))
            PubDModVal(xlApp, "134", "（スラブ補強）", DtlDInp(vbTab & "520× 980: "))
            PubDModVal(xlApp, "135", "（スラブ補強）", DtlDInp(vbTab & "520× 480: "))
        End If
    End Sub

    ''' <summary>
    ''' スラブ補強直 (D13).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    ''' <param name="truck2Ton">2 ton truck.</param>
    Friend Sub SlabReinfStr(xlApp As Application, choosen As Double, truck2Ton As Double)
        If choosen = 1 Then
            If Not truck2Ton = 1 Then
                PubDModVal(xlApp, "102", "（スラブ補強）", DtlDInp(vbTab & "6000: "))
                PubDModVal(xlApp, "103", "（スラブ補強）", DtlDInp(vbTab & "5500: "))
                PubDModVal(xlApp, "104", "（スラブ補強）", DtlDInp(vbTab & "5000: "))
            End If
            PubDModVal(xlApp, "105", "（スラブ補強）", DtlDInp(vbTab & "4500: "))
            PubDModVal(xlApp, "106", "（スラブ補強）", DtlDInp(vbTab & "4000: "))
            PubDModVal(xlApp, "107", "（スラブ補強）", DtlDInp(vbTab & "3500: "))
            PubDModVal(xlApp, "108", "（スラブ補強）", DtlDInp(vbTab & "3000: "))
            PubDModVal(xlApp, "109", "（スラブ補強）", DtlDInp(vbTab & "2500: "))
            PubDModVal(xlApp, "110", "（スラブ補強）", DtlDInp(vbTab & "2000: "))
            PubDModVal(xlApp, "111", "（スラブ補強）", DtlDInp(vbTab & "1500: "))
            PubDModVal(xlApp, "112", "（スラブ補強）", DtlDInp(vbTab & "1000: "))
        End If
    End Sub

    ''' <summary>
    ''' 副資材リスト.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    Friend Sub Parts(xlApp As Application)
        DctVal(xlApp, "AJ3", Today)
        DctVal(xlApp, "AG169", 60)
        DctVal(xlApp, "AG167", 100)
        Dim name = $"{DtlSInp(vbTab & "邸名" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & ": ")}様"
        DctVal(xlApp, "J5", name)
        CType(xlApp.ActiveSheet, Worksheet).Name = name
        PubDVal(xlApp, "AG164", DtlDInp(vbTab & "ポリシート0.1 (本)" & vbTab & vbTab & vbTab & vbTab & vbTab & ": "))
        PubDVal(xlApp, "AG165", DtlDInpDesc(vbTab & "ホールダウンアンカー (本)", vbTab & vbTab & "[M16×900]" & vbTab))
        PubDVal(xlApp, "AG166", DtlDInpDesc(vbTab & "フリークランクアンカーボルト (本)", vbTab & "[M12×400]" & vbTab))
        PubDVal(xlApp, "AG168", DtlDInpDesc(vbTab & "サイコロブロック (個)", vbTab & vbTab & vbTab & "[H70]" & vbTab & vbTab))
        PubDVal(xlApp, "AG170", DtlDInpDesc(vbTab & "プラドーナッツ (個)", vbTab & vbTab & vbTab & "[10・13×50]" & vbTab))
        PubDVal(xlApp, "AG171", DtlDInp(vbTab & "ポリドーナッツ (個)" & vbTab & vbTab & vbTab & vbTab & vbTab & ": "))
        PubDVal(xlApp, "AG172", DtlDInp(vbTab & "ブルーシート・台木 (式)" & vbTab & vbTab & vbTab & vbTab & vbTab & ": "))
    End Sub
End Module
