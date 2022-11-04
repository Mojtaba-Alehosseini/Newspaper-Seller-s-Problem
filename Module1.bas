Attribute VB_Name = "Module1"
Sub RunSimulation()
Attribute RunSimulation.VB_ProcData.VB_Invoke_Func = "R\n14"

    ' Input Variables
    ' NPP=Newspaper Purchase , DS=Days Simulation
    SellingPrice = Range("c4") / 100
    PurchasePrice = Range("c5") / 100
    ScrapPrice = Range("c6") / 100
    NOofNPP = Range("d9")
    NOofDS = Range("c11")
    Repeat = Range("c12")
    
    ' Type of Newsday , Pb=Probability
    GoodPb = Range("t6")
    FairPb = Range("t7")
    PoorPb = Range("t8")
    
    ' Type of Distribution of Newspapers Demanded
    Dim i As Integer
    Dim CPGood(1 To 7) As Double
    Dim CPFair(1 To 7) As Double
    Dim CPPoor(1 To 7) As Double
    Dim DNDemand(1 To 7) As Integer
    ' CP = Cumulative Probabilities & DN = Distribution of Newspapers
    ' nemidunm chejuri bedune for meghdar bedam beheshun, nemishod!
    For i = 1 To 7
        CPGood(i) = Cells(i + 6, "n")
    Next i
    For i = 1 To 7
        CPFair(i) = Cells(i + 6, "o")
    Next i
    For i = 1 To 7
        CPPoor(i) = Cells(i + 6, "p")
    Next i
    For i = 1 To 7
        DNDemand(i) = Cells(i + 6, "j")
    Next i
    
        
    ' Repeat Loop
    
    Dim x As Integer, y As Integer
    Dim RandomDigit As Double
    Dim TypeofNewsday As String
    Dim Demand As Integer
    Dim SumDailyProfit As Double
    Dim RTAvgProfit As Double
    Dim SumRTAvgProfit As Double
    Dim AvgProfit As Double
    RTAvgProfit = 0
    SumRTAvgProfit = 0
    AvgProfit = 0
    
    
    For x = 1 To Repeat
        SumDailyProfit = 0
        RTAvgProfit = 0
        Cells(x + 20, "l") = x
        Cells(x + 20, "m") = "Loading...!"
        Cells(17, "l") = x
        Cells(17, "m") = "Loading...!"
        
        For y = 1 To NOofDS
            ' Days Column
            Cells(y + 20, "a") = y
            Cells(17, "i") = y
            ' Random Digit for Type of Newsdaye column
            RandomDigit = Rnd()
            Cells(y + 20, "b") = RandomDigit
            ' Type of Newsday Column
            If RandomDigit <= GoodPb Then
                TypeofNewsday = "Good"
            Else
                If RandomDigit <= FairPb Then
                    TypeofNewsday = "Fair"
                Else
                    If RandomDigit <= PoorPb Then
                        TypeofNewsday = "Poor"
                    End If
                End If
            End If
            Cells(y + 20, "c") = TypeofNewsday
            
            ' Random Digit for Demand Column
            RandomDigit = Rnd()
            Cells(y + 20, "d") = RandomDigit
            ' Demand
            Dim k As Integer
            k = 1
            If TypeofNewsday = "Good" Then
                 Do While k <= 7
                    If RandomDigit <= CPGood(k) Then
                        Demand = DNDemand(k)
                        Exit Do
                    End If
                    k = k + 1
                 Loop
             End If
            If TypeofNewsday = "Fair" Then
                 Do While k <= 7
                    If RandomDigit <= CPFair(k) Then
                        Demand = DNDemand(k)
                        Exit Do
                    End If
                    k = k + 1
                 Loop
             End If
            If TypeofNewsday = "Poor" Then
                 Do While k <= 7
                    If RandomDigit <= CPPoor(k) Then
                        Demand = DNDemand(k)
                        Exit Do
                    End If
                    k = k + 1
                 Loop
             End If
             Cells(y + 20, "e") = Demand
             
             ' Revenue from Sales
             Dim Revenue As Double
             If Demand > NOofNPP Then
                Revenue = NOofNPP * SellingPrice
                Else
                   Revenue = Demand * SellingPrice
             End If
             Cells(y + 20, "f") = Revenue
             
             
             ' Lost Profit from Excess Demand
             Dim LostProfit As Double
             If Demand > NOofNPP Then
                LostProfit = (Demand - NOofNPP) * (SellingPrice - PurchasePrice)
                Else
                LostProfit = 0
             End If
             Cells(y + 20, "g") = LostProfit
             
             ' Salvage from Sale of Scrap
             Dim Scrap As Double
             If Demand < NOofNPP Then
                Scrap = (NOofNPP - Demand) * ScrapPrice
                Else
                Scrap = 0
             End If
             Cells(y + 20, "h") = Scrap
             
             ' Daily Cost
             Dim DailyCost As Double
             DailyCost = NOofNPP * PurchasePrice
             Cells(y + 20, "i") = DailyCost
             
             ' Daily Profit
             Dim DailyProfit As Double
             DailyProfit = Revenue + Scrap - DailyCost - LostProfit
             Cells(y + 20, "j") = DailyProfit
             Cells(17, "j") = DailyProfit
             
             ' Sum Daily Profit , baraye mohasebe miangin
             SumDailyProfit = SumDailyProfit + DailyProfit
             
        
            
        Next y
        
        ' Repeat Table
        RTAvgProfit = SumDailyProfit / NOofDS
        Cells(x + 20, "m") = RTAvgProfit
        Cells(17, "m") = RTAvgProfit
        SumRTAvgProfit = SumRTAvgProfit + RTAvgProfit


    Next x
    
    Range("i17", "m17").Clear
    
    
    ' Profit Table
    AvgProfit = SumRTAvgProfit / Repeat
    
    ' dige hal nadashtam ba halhe benevisam charta if gozashtm :)
    ' az 50 chon gofte bud faghat mitune tu baste haye 10 tayi bekhare va hadde aqal ham 5 baste bayad bekhare
    ' ta 100 chon tu jadvale demand max ruzi 100 ta neveshte shode
    
    If NOofNPP = 50 Then
        AvgProfitmulRepeat = AvgProfit * Repeat
        HistoryofAvgProfitmulRepeat = Range("q21").Value * Range("r21").Value
        TotalRepeat = Repeat + Range("q21").Value
        Cells(21, "r") = (AvgProfitmulRepeat + HistoryofAvgProfitmulRepeat) / TotalRepeat
        Cells(21, "q") = TotalRepeat
    End If
    
    If NOofNPP = 60 Then
        AvgProfitmulRepeat = AvgProfit * Repeat
        HistoryofAvgProfitmulRepeat = Range("q22").Value * Range("r22").Value
        TotalRepeat = Repeat + Range("q22").Value
        Cells(22, "r") = (AvgProfitmulRepeat + HistoryofAvgProfitmulRepeat) / TotalRepeat
        Cells(22, "q") = TotalRepeat
    End If
    
    If NOofNPP = 70 Then
        AvgProfitmulRepeat = AvgProfit * Repeat
        HistoryofAvgProfitmulRepeat = Range("q23").Value * Range("r23").Value
        TotalRepeat = Repeat + Range("q23").Value
        Cells(23, "r") = (AvgProfitmulRepeat + HistoryofAvgProfitmulRepeat) / TotalRepeat
        Cells(23, "q") = TotalRepeat
    End If
    
    If NOofNPP = 80 Then
        AvgProfitmulRepeat = AvgProfit * Repeat
        HistoryofAvgProfitmulRepeat = Range("q24").Value * Range("r24").Value
        TotalRepeat = Repeat + Range("q24").Value
        Cells(24, "r") = (AvgProfitmulRepeat + HistoryofAvgProfitmulRepeat) / TotalRepeat
        Cells(24, "q") = TotalRepeat
    End If
    
    If NOofNPP = 90 Then
        AvgProfitmulRepeat = AvgProfit * Repeat
        HistoryofAvgProfitmulRepeat = Range("q25").Value * Range("r25").Value
        TotalRepeat = Repeat + Range("q25").Value
        Cells(25, "r") = (AvgProfitmulRepeat + HistoryofAvgProfitmulRepeat) / TotalRepeat
        Cells(25, "q") = TotalRepeat
    End If
    
    If NOofNPP = 100 Then
        AvgProfitmulRepeat = AvgProfit * Repeat
        HistoryofAvgProfitmulRepeat = Range("q26").Value * Range("r26").Value
        TotalRepeat = Repeat + Range("q26").Value
        Cells(26, "r") = (AvgProfitmulRepeat + HistoryofAvgProfitmulRepeat) / TotalRepeat
        Cells(26, "q") = TotalRepeat
    End If
    
    
End Sub
