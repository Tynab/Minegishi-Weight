﻿Imports Microsoft.Office.Interop.Excel

Friend Module Service
    ''' <summary>
    ''' Weight Minegishi.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    Friend Sub WtMinegishi(xlApp As Application)
        ' Fare
        Dim truck2Ton = HdrYNQ(vbTab & vbTab & "運賃 (2トン車): ")
        Fare(xlApp)
        ' Circumference
        PrefWarn(vbTab & vbTab & "外周")
        Circumference(xlApp)
        'Inner circumference
        PrefWarn(vbTab & vbTab & "内周")
        InnerCircumference(xlApp)
        ' Straight joint
        PubDVal(xlApp, "AG95", HdrDInp(vbTab & vbTab & "ストレート (D16[L-1500]): "))
        ' Straight D13
        PrefWarn(vbTab & vbTab & "ストレート (D13)")
        StrD13(xlApp)
        ' Straight D10
        StrD10(xlApp, HdrYNQ(vbTab & vbTab & "ストレート (D10): "))
        ' Corner joint
        PubDVal(xlApp, "AG96", HdrDInp(vbTab & vbTab & "コーナー (D16[750×750]): "))
        ' Corner D13
        PrefWarn(vbTab & vbTab & "コーナー (D13)")
        CorD13(xlApp)
        ' Corner 155 degree
        PubDVal(xlApp, "AG75", HdrDInp(vbTab & vbTab & "コーナー (D13[曲 155°]): "))
        ' Crank
        Crank(xlApp, HdrYNQ(vbTab & vbTab & "クランク: "))
        ' U type
        UType(xlApp, HdrYNQ(vbTab & vbTab & "コ型: "))
        ' Hook
        Hook(xlApp, HdrYNQ(vbTab & vbTab & "フック (D10): "))
        ' Slab bending
        PrefWarn(vbTab & vbTab & "スラブ曲 (D10)")
        SlabBndg(xlApp, truck2Ton)
        ' Slab straight
        SlabStr(xlApp, HdrYNQ(vbTab & vbTab & "スラブ直 (D10): "), truck2Ton)
        ' Slab reinforcement bending
        SlabReinfBndg(xlApp, HdrYNQ(vbTab & vbTab & "スラブ補強曲 (D13): "), truck2Ton)
        ' Slab reinforcement straight
        SlabReinfStr(xlApp, HdrYNQ(vbTab & vbTab & "スラブ補強直 (D13): "), truck2Ton)
        ' Parts
        PrefWarn(vbTab & vbTab & "副資材リスト")
        Parts(xlApp)
        ' Free row 71
        ' Free row 97
    End Sub
End Module
