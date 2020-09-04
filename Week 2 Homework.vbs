{\rtf1\ansi\ansicpg1252\cocoartf2513
\cocoatextscaling0\cocoaplatform0{\fonttbl\f0\fswiss\fcharset0 Helvetica;}
{\colortbl;\red255\green255\blue255;}
{\*\expandedcolortbl;;}
\paperw11900\paperh16840\margl1440\margr1440\vieww10800\viewh8400\viewkind0
\pard\tx566\tx1133\tx1700\tx2267\tx2834\tx3401\tx3968\tx4535\tx5102\tx5669\tx6236\tx6803\pardirnatural\partightenfactor0

\f0\fs24 \cf0 Sub stockpractice()\
\
    ' Preparing a loop to run through all sheets\
    Dim Ws As Integer\
    For Ws = 1 To Sheets.Count\
        Sheets(Ws).Activate\
         \
        ' Setting up the summary table\
        ' -------------------------\
        Dim summary_row As Integer\
        Dim total_stock_volume As Double\
        summary_row = 2\
        Range("I1").Value = "Ticker"\
        Range("J1").Value = "Yearly Change"\
        Range("K1").Value = "% change"\
        Range("L1").Value = "Total stock volume"\
        \
        'setting up 'greatest' table\
        ' ---------------------\
        Range("N1").Value = "Greatest ..."\
        Range("N2").Value = "Greatest % increase"\
        Range("N3").Value = "Greatest % decrease"\
        Range("N4").Value = "Greatest Total Volume"\
        Range("O1").Value = "Ticker"\
        Range("P1").Value = "Value"\
        Columns("N").AutoFit\
        \
        ' count the last row\
        Dim lastrow As Long\
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row\
        \
        ' LOOP FOR MAIN SUMMARY TABLE\
        '============================\
        \
        ' loop through all stock data\
        For i = 2 To lastrow\
        \
            ' volume definition\
            volume = Cells(i, 7).Value + 0\
        \
            ' setting the parameters for ticker\
            ticker = Cells(i, 1).Value\
            \
           ' establishing the open price for each ticker by checking if the previous cell has a different reference\
            If ticker <> Cells(i - 1, 1).Value Then\
                open_price = Cells(i, 3).Value + 0\
            End If\
            \
            ' check if we are still on the same ticker and if not, then...\
            If ticker <> Cells(i + 1, 1).Value Then\
                \
                ' select the ticker\
                ticker = Cells(i, 1).Value\
                \
                ' add it to the summary table\
                Range("I" & summary_row).Value = ticker\
                \
                ' set the close price\
                close_price = Cells(i, 6).Value + 0\
                \
                ' calculate the yearly change in the summary table\
                Range("J" & summary_row).Value = close_price - open_price\
                \
                    ' adding the conditional colour formatting\
                    If Range("J" & summary_row).Value > 0 Then\
                        Range("J" & summary_row).Interior.ColorIndex = 4\
                    Else\
                        Range("J" & summary_row).Interior.ColorIndex = 3\
                    End If\
                \
                ' calculate the percentage change\
                If open_price > 0 And close_price > 0 Then\
                    Range("K" & summary_row).Value = (close_price - open_price) / open_price\
                    Range("K" & summary_row).NumberFormat = "0.00%"\
                    Else\
                    Range("K" & summary_row).Value = 0\
                End If\
                \
                ' add to the total stock volume\
                If volume > 0 Then\
                    total_stock_volume = total_stock_volume + volume\
                    Else\
                    \
                End If\
                \
                ' print total stock volume\
                Range("L" & summary_row).Value = total_stock_volume\
                \
                ' add one to the summary row\
                summary_row = summary_row + 1\
                \
                ' Reset Total stock volume\
                total_stock_volume = 0 + 0\
                \
            ' if it is the same ticker value\
            Else\
            \
                ' Add to the total stock volume\
                  total_stock_volume = total_stock_volume + volume\
                  \
            End If\
            \
        Next i\
        \
        ' CALCULATING THE VALUE SUMMARY TABLE\
        ' ===========================\
        \
        ' for each new worksheet, resetting the default value\
        highest_yearly_change = Range("K2").Value\
        lowest_yearly_change = Range("K2").Value\
        greatest_volume = Range("L2").Value\
        \
        'finding the last row of each summary table\
        last_row = Cells(Rows.Count, 10).End(xlUp).Row\
        \
        ' setting a loop to run through the summary table after it has been calculated\
        For x = 2 To last_row\
        \
            ' calculating the greatest % increase\
            If Cells(x, 11).Value > highest_yearly_change Then\
                highest_yearly_change = Cells(x, 11).Value\
                highest_ticker = Cells(x, 9).Value\
            End If\
            \
            ' calculating the lowest % increase\
            If Cells(x, 11).Value < lowest_yearly_change Then\
                lowest_yearly_change = Cells(x, 11).Value\
                lowest_ticker = Cells(x, 9).Value\
            End If\
            \
            ' calculating the greatest volume\
            If Cells(x, 12).Value > greatest_volume Then\
                greatest_volume = Cells(x, 12).Value\
                volume_ticker = Cells(x, 9).Value\
            End If\
            \
        Next x\
                \
            ' filling in the value table\
            Range("O2").Value = highest_ticker\
            Range("P2").Value = highest_yearly_change\
            Range("P2").NumberFormat = "0.00%"\
            Range("O3").Value = lowest_ticker\
            Range("P3").Value = lowest_yearly_change\
            Range("P3").NumberFormat = "0.00%"\
            Range("O4").Value = volume_ticker\
            Range("P4").Value = greatest_volume\
        \
    Next Ws\
    \
End Sub\
\
\
        \
        \
            \
        \
            \
        \
        \
    \
    \
\
}