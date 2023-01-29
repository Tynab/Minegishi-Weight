# MINEGISHI WEIGHT SOLUTION
Solution to help 西山 team of エマール group transfer data faster for 峰岸重量 from 文化シャッター partner.

## MASK
<p align="center">
<img src="https://raw.githubusercontent.com/Tynab/Minegishi-Weight/main/pic/0.png"></img>
</p>

## CODE DEMO
```vb
''' <summary>
''' 運賃 (2トン車).
''' </summary>
''' <param name="xlApp">Excel Application.</param>
Friend Sub Fare(xlApp As Application)
    PubDModVal(xlApp, "78", "D10", "400×200", 0.4, 15)
End Sub
```