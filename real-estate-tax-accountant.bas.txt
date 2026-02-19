Option Explicit

' ===========================================================
' 부동산 세금 계산기 2026 - VBA 모듈
' ===========================================================


' 공통 메시지박스
Sub ShowInfo(msg As String)
    MsgBox msg, vbInformation, "부동산 세금 계산기 2026"
End Sub

' 시트 이동
Sub GoToYangdo()
    Sheets("양도소득세").Select
    Sheets("양도소득세").Range("C7").Select
End Sub

Sub GoToJongbu()
    Sheets("종합부동산세").Select
    Sheets("종합부동산세").Range("C7").Select
End Sub

Sub GoToJeungYeo()
    Sheets("증여세").Select
    Sheets("증여세").Range("C7").Select
End Sub

Sub GoToSangsok()
    Sheets("상속세").Select
    Sheets("상속세").Range("C7").Select
End Sub

Sub GoToImde()
    Sheets("임대소득세").Select
    Sheets("임대소득세").Range("C7").Select
End Sub

Sub GoToMain()
    Sheets("메인").Select
End Sub

Sub GoToRateTable()
    Sheets("세율표").Select
End Sub

' ===========================================================
' 누진세율 공통 함수 (2026 소득세 기본세율)
' ===========================================================
Sub GetProgressiveTax(ByVal base As Double, ByRef rate As Double, ByRef ded As Double)
    If base <= 14000000 Then
        rate = 0.06: ded = 0
    ElseIf base <= 50000000 Then
        rate = 0.15: ded = 1260000
    ElseIf base <= 88000000 Then
        rate = 0.24: ded = 5760000
    ElseIf base <= 150000000 Then
        rate = 0.35: ded = 15440000
    ElseIf base <= 300000000 Then
        rate = 0.38: ded = 19940000
    ElseIf base <= 500000000 Then
        rate = 0.40: ded = 25940000
    ElseIf base <= 1000000000 Then
        rate = 0.42: ded = 35940000
    Else
        rate = 0.45: ded = 65940000
    End If
End Sub

' ===========================================================
' 1. 양도소득세 계산
' ===========================================================
Sub CalcYangdo()
    Dim ws As Worksheet
    Set ws = Sheets("양도소득세")

    Dim yangdoAmt As Double
    Dim chwidukAmt As Double
    Dim gyungbi As Double
    Dim boyuYears As Double
    Dim jugeoyears As Double
    Dim houseCnt As Integer
    Dim isAdjust As Boolean
    Dim isOneHouse As Boolean
    Dim isNonBiz As Boolean
    Dim isUnreg As Boolean

    yangdoAmt = Val(ws.Range("C7").Value)
    chwidukAmt = Val(ws.Range("C8").Value)
    gyungbi = Val(ws.Range("C9").Value)
    boyuYears = Val(ws.Range("C10").Value)
    jugeoyears = Val(ws.Range("C11").Value)
    houseCnt = Val(ws.Range("C12").Value)
    isAdjust = (ws.Range("C13").Value = "예")
    isOneHouse = (ws.Range("C14").Value = "예")
    isNonBiz = (ws.Range("C15").Value = "예")
    isUnreg = (ws.Range("C16").Value = "예")

    If yangdoAmt <= 0 Or chwidukAmt <= 0 Then
        ShowInfo "양도가액과 취득가액을 올바르게 입력해주세요."
        Exit Sub
    End If

    ' 양도차익
    Dim yangdoChaiik As Double
    yangdoChaiik = yangdoAmt - chwidukAmt - gyungbi

    ' 비과세 판단
    Dim isNonTax As Boolean
    isNonTax = False
    If isOneHouse And boyuYears >= 2 And houseCnt = 1 Then
        If isAdjust Then
            If jugeoyears >= 2 And yangdoAmt <= 1200000000 Then
                isNonTax = True
            End If
        Else
            If yangdoAmt <= 1200000000 Then
                isNonTax = True
            End If
        End If
    End If

    ' 고가주택 과세 비율 (12억 초과)
    Dim taxableRatio As Double
    taxableRatio = 1
    If isOneHouse And yangdoAmt > 1200000000 And boyuYears >= 2 Then
        taxableRatio = (yangdoAmt - 1200000000) / yangdoAmt
        isNonTax = False
    End If

    Dim taxableGain As Double
    If isNonTax Then
        taxableGain = 0
    Else
        taxableGain = yangdoChaiik * taxableRatio
    End If

    ' 장기보유특별공제율
    Dim ltdRate As Double
    ltdRate = 0

    Dim holdRate As Double
    Dim liveRate As Double
    holdRate = 0
    liveRate = 0

    If isUnreg Then
        ltdRate = 0
    ElseIf isNonTax Then
        ltdRate = 0
    ElseIf isOneHouse And boyuYears >= 3 Then
        If boyuYears >= 10 Then
            holdRate = 0.4
        Else
            holdRate = Int(boyuYears) * 0.04
        End If
        If jugeoyears >= 2 Then
            If jugeoyears >= 10 Then
                liveRate = 0.4
            Else
                liveRate = Int(jugeoyears) * 0.04
            End If
        Else
            liveRate = 0
            holdRate = Int(boyuYears) * 0.02
        End If
        ltdRate = holdRate + liveRate
        If ltdRate > 0.8 Then ltdRate = 0.8
    ElseIf boyuYears >= 3 Then
        ltdRate = Int(boyuYears) * 0.02
        If ltdRate > 0.3 Then ltdRate = 0.3
    End If

    Dim ltdAmt As Double
    ltdAmt = taxableGain * ltdRate

    ' 양도소득금액
    Dim yangdoSoduk As Double
    yangdoSoduk = taxableGain - ltdAmt

    ' 기본공제 250만원
    Dim basicDeduct As Double
    If isUnreg Then
        basicDeduct = 0
    Else
        basicDeduct = 2500000
    End If
    If yangdoSoduk < basicDeduct Then basicDeduct = yangdoSoduk

    ' 과세표준
    Dim taxBase As Double
    taxBase = yangdoSoduk - basicDeduct
    If taxBase < 0 Then taxBase = 0

    ' 세율
    Dim taxRate As Double
    Dim dedAmt As Double
    taxRate = 0
    dedAmt = 0

    If isUnreg Then
        taxRate = 0.7: dedAmt = 0
    ElseIf boyuYears < 1 Then
        taxRate = 0.7: dedAmt = 0
    ElseIf boyuYears < 2 Then
        taxRate = 0.6: dedAmt = 0
    Else
        If isNonBiz Then
            Call GetProgressiveTax(taxBase, taxRate, dedAmt)
            taxRate = taxRate + 0.1
        Else
            Call GetProgressiveTax(taxBase, taxRate, dedAmt)
        End If
    End If

    ' 산출세액
    Dim sanChuTax As Double
    sanChuTax = taxBase * taxRate - dedAmt
    If sanChuTax < 0 Then sanChuTax = 0

    ' 지방소득세
    Dim localTax As Double
    localTax = sanChuTax * 0.1

    ' 총 납부세액
    Dim totalTax As Double
    totalTax = sanChuTax + localTax

    ' 결과 출력
    ws.Range("G7").Value = yangdoChaiik
    ws.Range("G8").Value = IIf(isNonTax, "비과세", "과세")
    ws.Range("G9").Value = ltdRate
    ws.Range("G10").Value = ltdAmt
    ws.Range("G11").Value = yangdoSoduk
    ws.Range("G12").Value = basicDeduct
    ws.Range("G13").Value = taxBase
    ws.Range("G14").Value = taxRate
    ws.Range("G15").Value = sanChuTax
    ws.Range("G16").Value = localTax
    ws.Range("G17").Value = totalTax

    If yangdoChaiik > 0 Then
        ws.Range("G18").Value = totalTax / yangdoChaiik
    Else
        ws.Range("G18").Value = 0
    End If

    With ws.Range("G17")
        .Font.Bold = True
        .Font.Size = 14
        If totalTax > 100000000 Then
            .Font.Color = RGB(200, 0, 0)
        ElseIf totalTax > 0 Then
            .Font.Color = RGB(30, 100, 200)
        Else
            .Font.Color = RGB(0, 150, 0)
        End If
    End With

    ShowInfo "계산이 완료되었습니다!" & Chr(13) & Chr(13) & _
             "▶ 양도차익: " & Format(yangdoChaiik, "#,##0") & "원" & Chr(13) & _
             "▶ 장기보유특별공제: " & Format(ltdRate * 100, "0.0") & "%" & Chr(13) & _
             "▶ 산출세액: " & Format(sanChuTax, "#,##0") & "원" & Chr(13) & _
             "▶ 지방소득세: " & Format(localTax, "#,##0") & "원" & Chr(13) & _
             "======================" & Chr(13) & _
             "▶ 총 납부세액: " & Format(totalTax, "#,##0") & "원"
End Sub

Sub ResetYangdo()
    Dim ws As Worksheet
    Set ws = Sheets("양도소득세")
    ws.Range("C7:C16").ClearContents
    ws.Range("G7:G18").ClearContents
    ws.Range("C12").Value = 1
    ws.Range("C13").Value = "아니오"
    ws.Range("C14").Value = "예"
    ws.Range("C15").Value = "아니오"
    ws.Range("C16").Value = "아니오"
End Sub

' ===========================================================
' 2. 종합부동산세 계산
' ===========================================================
Sub CalcJongbu()
    Dim ws As Worksheet
    Set ws = Sheets("종합부동산세")

    Dim publicPrice As Double
    Dim houseCnt As Integer
    Dim isOne As Boolean
    Dim age As Integer
    Dim holdYears As Integer
    Dim isCorp As Boolean

    publicPrice = Val(ws.Range("C7").Value)
    houseCnt = Val(ws.Range("C8").Value)
    isOne = (ws.Range("C9").Value = "예")
    age = Val(ws.Range("C10").Value)
    holdYears = Val(ws.Range("C11").Value)
    isCorp = (ws.Range("C12").Value = "예")

    If publicPrice <= 0 Then
        ShowInfo "공시가격 합계를 입력해주세요."
        Exit Sub
    End If

    ' 공제금액
    Dim deductAmt As Double
    If isCorp Then
        deductAmt = 0
    ElseIf isOne And houseCnt = 1 Then
        deductAmt = 1200000000
    Else
        deductAmt = 900000000
    End If

    ' 과세표준 (공정시장가액비율 60%)
    Dim taxBase As Double
    taxBase = (publicPrice - deductAmt) * 0.6

    If taxBase <= 0 Then
        ws.Range("G7").Value = publicPrice
        ws.Range("G8").Value = deductAmt
        ws.Range("G9").Value = 0
        ws.Range("G10").Value = 0
        ws.Range("G11").Value = 0
        ws.Range("G12").Value = 0
        ws.Range("G13").Value = 0
        ws.Range("G14").Value = 0
        ws.Range("G15").Value = 0
        ws.Range("G16").Value = 0
        ShowInfo "공시가격이 공제금액 이하입니다. 종합부동산세 과세 대상이 아닙니다."
        Exit Sub
    End If

    ' 세율
    Dim taxRate As Double
    Dim ded As Double
    taxRate = 0
    ded = 0

    If isCorp Then
        taxRate = 0.05: ded = 0
    ElseIf houseCnt >= 3 Then
        If taxBase <= 300000000 Then
            taxRate = 0.005: ded = 0
        ElseIf taxBase <= 600000000 Then
            taxRate = 0.007: ded = 600000
        ElseIf taxBase <= 1200000000 Then
            taxRate = 0.01: ded = 2400000
        ElseIf taxBase <= 2500000000 Then
            taxRate = 0.02: ded = 14400000
        ElseIf taxBase <= 5000000000# Then
            taxRate = 0.03: ded = 39400000
        Else
            taxRate = 0.05: ded = 139400000
        End If
    Else
        If taxBase <= 300000000 Then
            taxRate = 0.005: ded = 0
        ElseIf taxBase <= 600000000 Then
            taxRate = 0.007: ded = 600000
        ElseIf taxBase <= 1200000000 Then
            taxRate = 0.01: ded = 2400000
        ElseIf taxBase <= 2500000000 Then
            taxRate = 0.014: ded = 7200000
        ElseIf taxBase <= 9400000000# Then
            taxRate = 0.02: ded = 22200000
        Else
            taxRate = 0.027: ded = 87999000
        End If
    End If

    ' 산출세액
    Dim sanChu As Double
    sanChu = taxBase * taxRate - ded
    If sanChu < 0 Then sanChu = 0

    ' 재산세 공제 (추정)
    Dim propTaxDeduct As Double
    propTaxDeduct = publicPrice * 0.0014

    ' 고령자 + 장기보유 세액공제
    Dim ageRate As Double
    Dim holdRate2 As Double
    Dim seAekRate As Double
    ageRate = 0
    holdRate2 = 0
    seAekRate = 0

    If isOne And houseCnt = 1 Then
        If age >= 70 Then
            ageRate = 0.4
        ElseIf age >= 65 Then
            ageRate = 0.3
        ElseIf age >= 60 Then
            ageRate = 0.2
        End If

        If holdYears >= 15 Then
            holdRate2 = 0.5
        ElseIf holdYears >= 10 Then
            holdRate2 = 0.4
        ElseIf holdYears >= 5 Then
            holdRate2 = 0.2
        End If

        seAekRate = ageRate + holdRate2
        If seAekRate > 0.8 Then seAekRate = 0.8
    End If

    Dim taxAmt As Double
    taxAmt = (sanChu - propTaxDeduct) * (1 - seAekRate)
    If taxAmt < 0 Then taxAmt = 0

    ' 농어촌특별세 20%
    Dim ruralTax As Double
    ruralTax = taxAmt * 0.2

    Dim totalTax As Double
    totalTax = taxAmt + ruralTax

    ws.Range("G7").Value = publicPrice
    ws.Range("G8").Value = deductAmt
    ws.Range("G9").Value = taxBase
    ws.Range("G10").Value = taxRate
    ws.Range("G11").Value = sanChu
    ws.Range("G12").Value = propTaxDeduct
    ws.Range("G13").Value = seAekRate
    ws.Range("G14").Value = taxAmt
    ws.Range("G15").Value = ruralTax
    ws.Range("G16").Value = totalTax

    ShowInfo "계산이 완료되었습니다!" & Chr(13) & Chr(13) & _
             "▶ 과세표준: " & Format(taxBase, "#,##0") & "원" & Chr(13) & _
             "▶ 세율: " & Format(taxRate * 100, "0.0") & "%" & Chr(13) & _
             "▶ 고령자+장기보유 공제율: " & Format(seAekRate * 100, "0") & "%" & Chr(13) & _
             "▶ 종부세: " & Format(taxAmt, "#,##0") & "원" & Chr(13) & _
             "▶ 농어촌특별세: " & Format(ruralTax, "#,##0") & "원" & Chr(13) & _
             "======================" & Chr(13) & _
             "▶ 총 납부세액: " & Format(totalTax, "#,##0") & "원"
End Sub

Sub ResetJongbu()
    Sheets("종합부동산세").Range("C7:C12").ClearContents
    Sheets("종합부동산세").Range("G7:G16").ClearContents
    Sheets("종합부동산세").Range("C8").Value = 1
    Sheets("종합부동산세").Range("C9").Value = "예"
    Sheets("종합부동산세").Range("C12").Value = "아니오"
End Sub

' ===========================================================
' 3. 증여세 계산
' ===========================================================
Sub CalcJeungYeo()
    Dim ws As Worksheet
    Set ws = Sheets("증여세")

    Dim giftAmt As Double
    Dim relation As String
    Dim prevGift As Double
    Dim isMinor As Boolean
    Dim debtAmt As Double

    giftAmt = Val(ws.Range("C7").Value)
    relation = ws.Range("C8").Value
    prevGift = Val(ws.Range("C9").Value)
    isMinor = (ws.Range("C10").Value = "예")
    debtAmt = Val(ws.Range("C11").Value)

    If giftAmt <= 0 Then
        ShowInfo "증여재산가액을 입력해주세요."
        Exit Sub
    End If

    ' 증여재산공제
    Dim deductAmt As Double
    Select Case relation
        Case "배우자"
            deductAmt = 600000000
        Case "직계존속->직계비속"
            If isMinor Then
                deductAmt = 20000000
            Else
                deductAmt = 50000000
            End If
        Case "직계비속->직계존속"
            deductAmt = 50000000
        Case "기타 친족"
            deductAmt = 10000000
        Case Else
            deductAmt = 0
    End Select

    ' 순수 증여분
    Dim netGift As Double
    netGift = giftAmt - debtAmt
    If netGift < 0 Then netGift = 0

    ' 과세표준
    Dim taxBase As Double
    taxBase = netGift + prevGift - deductAmt
    If taxBase < 0 Then taxBase = 0

    ' 세율
    Dim taxRate As Double
    Dim ded As Double
    taxRate = 0
    ded = 0

    If taxBase <= 100000000 Then
        taxRate = 0.1: ded = 0
    ElseIf taxBase <= 500000000 Then
        taxRate = 0.2: ded = 10000000
    ElseIf taxBase <= 1000000000 Then
        taxRate = 0.3: ded = 60000000
    ElseIf taxBase <= 3000000000# Then
        taxRate = 0.4: ded = 160000000
    Else
        taxRate = 0.5: ded = 460000000
    End If

    Dim totalCalc As Double
    totalCalc = taxBase * taxRate - ded

    ' 전증여 세액 차감
    Dim prevTaxBase As Double
    Dim prevTax As Double
    prevTaxBase = prevGift - deductAmt
    prevTax = 0

    If prevTaxBase > 0 Then
        Dim prevRate As Double
        Dim prevDed As Double
        prevRate = 0
        prevDed = 0
        If prevTaxBase <= 100000000 Then
            prevRate = 0.1: prevDed = 0
        ElseIf prevTaxBase <= 500000000 Then
            prevRate = 0.2: prevDed = 10000000
        ElseIf prevTaxBase <= 1000000000 Then
            prevRate = 0.3: prevDed = 60000000
        ElseIf prevTaxBase <= 3000000000# Then
            prevRate = 0.4: prevDed = 160000000
        Else
            prevRate = 0.5: prevDed = 460000000
        End If
        prevTax = prevTaxBase * prevRate - prevDed
    End If

    Dim sanChu As Double
    sanChu = totalCalc - prevTax
    If sanChu < 0 Then sanChu = 0

    ' 신고세액공제 3%
    Dim reportDeduct As Double
    reportDeduct = sanChu * 0.03

    Dim finalTax As Double
    finalTax = sanChu - reportDeduct

    ' 증여 취득세 3.5%
    Dim acquiTax As Double
    acquiTax = netGift * 0.035

    ws.Range("G7").Value = giftAmt
    ws.Range("G8").Value = debtAmt
    ws.Range("G9").Value = netGift
    ws.Range("G10").Value = deductAmt
    ws.Range("G11").Value = prevGift
    ws.Range("G12").Value = taxBase
    ws.Range("G13").Value = taxRate
    ws.Range("G14").Value = sanChu
    ws.Range("G15").Value = reportDeduct
    ws.Range("G16").Value = finalTax
    ws.Range("G17").Value = acquiTax
    ws.Range("G18").Value = finalTax + acquiTax

    ShowInfo "계산이 완료되었습니다!" & Chr(13) & Chr(13) & _
             "▶ 순증여재산가액: " & Format(netGift, "#,##0") & "원" & Chr(13) & _
             "▶ 증여재산공제: " & Format(deductAmt, "#,##0") & "원" & Chr(13) & _
             "▶ 과세표준: " & Format(taxBase, "#,##0") & "원" & Chr(13) & _
             "▶ 산출세액: " & Format(sanChu, "#,##0") & "원" & Chr(13) & _
             "▶ 신고공제(-3%): " & Format(reportDeduct, "#,##0") & "원" & Chr(13) & _
             "▶ 납부 증여세: " & Format(finalTax, "#,##0") & "원" & Chr(13) & _
             "▶ 증여 취득세(3.5%): " & Format(acquiTax, "#,##0") & "원" & Chr(13) & _
             "======================" & Chr(13) & _
             "▶ 합계 납부세액: " & Format(finalTax + acquiTax, "#,##0") & "원"
End Sub

Sub ResetJeungYeo()
    Sheets("증여세").Range("C7:C11").ClearContents
    Sheets("증여세").Range("G7:G18").ClearContents
    Sheets("증여세").Range("C8").Value = "직계존속->직계비속"
    Sheets("증여세").Range("C10").Value = "아니오"
End Sub

' ===========================================================
' 4. 상속세 계산
' ===========================================================
Sub CalcSangsok()
    Dim ws As Worksheet
    Set ws = Sheets("상속세")

    Dim totalAsset As Double
    Dim debt As Double
    Dim funeral As Double
    Dim preGift As Double
    Dim hasSpouse As Boolean
    Dim childCnt As Integer
    Dim minorAmt As Double
    Dim disabledAmt As Double
    Dim financialAmt As Double
    Dim sameHouseAmt As Double

    totalAsset = Val(ws.Range("C7").Value)
    debt = Val(ws.Range("C8").Value)
    funeral = Val(ws.Range("C9").Value)
    preGift = Val(ws.Range("C10").Value)
    hasSpouse = (ws.Range("C11").Value = "예")
    childCnt = Val(ws.Range("C12").Value)
    minorAmt = Val(ws.Range("C13").Value)
    disabledAmt = Val(ws.Range("C14").Value)
    financialAmt = Val(ws.Range("C15").Value)
    sameHouseAmt = Val(ws.Range("C16").Value)

    If totalAsset <= 0 Then
        ShowInfo "상속재산 합계를 입력해주세요."
        Exit Sub
    End If

    ' 장례비 한도
    Dim maxFuneral As Double
    maxFuneral = 15000000
    If funeral > maxFuneral Then funeral = maxFuneral

    ' 과세가액
    Dim taxableAmt As Double
    taxableAmt = totalAsset - debt - funeral + preGift
    If taxableAmt < 0 Then taxableAmt = 0

    ' 기초공제 2억
    Dim basicDeduct As Double
    basicDeduct = 200000000

    ' 인적공제
    Dim personalDeduct As Double
    personalDeduct = childCnt * 50000000
    personalDeduct = personalDeduct + minorAmt * 10000000
    personalDeduct = personalDeduct + disabledAmt * 10000000

    ' 일괄공제 5억 vs 기초+인적
    Dim totalPersonal As Double
    totalPersonal = basicDeduct + personalDeduct

    Dim useDeduct As Double
    If totalPersonal > 500000000 Then
        useDeduct = totalPersonal
    Else
        useDeduct = 500000000
    End If

    ' 배우자공제 (최소 5억)
    Dim spouseDeduct As Double
    spouseDeduct = 0
    If hasSpouse Then
        spouseDeduct = 500000000
    End If

    ' 금융재산공제 (20%, 최대 2억)
    Dim finDeduct As Double
    finDeduct = financialAmt * 0.2
    If finDeduct > 200000000 Then finDeduct = 200000000

    ' 동거주택공제 (80%, 최대 6억)
    Dim houseDeduct As Double
    houseDeduct = sameHouseAmt * 0.8
    If houseDeduct > 600000000 Then houseDeduct = 600000000

    ' 총 공제
    Dim totalDeduct As Double
    totalDeduct = useDeduct + spouseDeduct + finDeduct + houseDeduct

    ' 과세표준
    Dim taxBase As Double
    taxBase = taxableAmt - totalDeduct
    If taxBase < 0 Then taxBase = 0

    ' 세율
    Dim taxRate As Double
    Dim ded As Double
    taxRate = 0
    ded = 0

    If taxBase <= 100000000 Then
        taxRate = 0.1: ded = 0
    ElseIf taxBase <= 500000000 Then
        taxRate = 0.2: ded = 10000000
    ElseIf taxBase <= 1000000000 Then
        taxRate = 0.3: ded = 60000000
    ElseIf taxBase <= 3000000000# Then
        taxRate = 0.4: ded = 160000000
    Else
        taxRate = 0.5: ded = 460000000
    End If

    Dim sanChu As Double
    sanChu = taxBase * taxRate - ded
    If sanChu < 0 Then sanChu = 0

    ws.Range("G7").Value = totalAsset
    ws.Range("G8").Value = debt + funeral
    ws.Range("G9").Value = preGift
    ws.Range("G10").Value = taxableAmt
    ws.Range("G11").Value = useDeduct
    ws.Range("G12").Value = spouseDeduct
    ws.Range("G13").Value = finDeduct
    ws.Range("G14").Value = houseDeduct
    ws.Range("G15").Value = totalDeduct
    ws.Range("G16").Value = taxBase
    ws.Range("G17").Value = taxRate
    ws.Range("G18").Value = sanChu

    ShowInfo "계산이 완료되었습니다!" & Chr(13) & Chr(13) & _
             "▶ 상속세 과세가액: " & Format(taxableAmt, "#,##0") & "원" & Chr(13) & _
             "▶ 총 상속공제: " & Format(totalDeduct, "#,##0") & "원" & Chr(13) & _
             "▶ 과세표준: " & Format(taxBase, "#,##0") & "원" & Chr(13) & _
             "▶ 세율: " & Format(taxRate * 100, "0") & "%" & Chr(13) & _
             "======================" & Chr(13) & _
             "▶ 산출 상속세: " & Format(sanChu, "#,##0") & "원"
End Sub

Sub ResetSangsok()
    Sheets("상속세").Range("C7:C16").ClearContents
    Sheets("상속세").Range("G7:G18").ClearContents
    Sheets("상속세").Range("C11").Value = "예"
End Sub

' ===========================================================
' 5. 임대소득세 계산
' ===========================================================
Sub CalcImde()
    Dim ws As Worksheet
    Set ws = Sheets("임대소득세")

    Dim monthlyRent As Double
    Dim deposit As Double
    Dim houseCnt As Integer
    Dim otherIncome As Double
    Dim isReg As Boolean
    Dim pubPrice As Double

    monthlyRent = Val(ws.Range("C7").Value)
    deposit = Val(ws.Range("C8").Value)
    houseCnt = Val(ws.Range("C9").Value)
    otherIncome = Val(ws.Range("C10").Value)
    isReg = (ws.Range("C11").Value = "예")
    pubPrice = Val(ws.Range("C12").Value)

    ' 과세 여부 판단
    Dim isTaxable As Boolean
    isTaxable = False
    If houseCnt >= 3 Then isTaxable = True
    If houseCnt = 2 And monthlyRent > 0 Then isTaxable = True
    If houseCnt = 1 And pubPrice > 1200000000 And monthlyRent > 0 Then isTaxable = True

    If Not isTaxable Then
        ShowInfo "현재 입력 기준으로 주택임대소득 과세대상이 아닙니다." & Chr(13) & _
                 "(1주택: 기준시가 12억 이하 비과세, 2주택 이하 전세 비과세)"
        Exit Sub
    End If

    ' 간주임대료 (3주택 이상, 보증금 3억 초과)
    Dim ganjuRent As Double
    ganjuRent = 0
    If houseCnt >= 3 And deposit > 300000000 Then
        ganjuRent = (deposit - 300000000) * 0.6 * 0.031
    End If

    ' 총 임대수입
    Dim totalRent As Double
    totalRent = monthlyRent + ganjuRent

    ' 필요경비율
    Dim expenseRate As Double
    If isReg Then
        expenseRate = 0.6
    Else
        expenseRate = 0.5
    End If

    ' 기본공제
    Dim basicDeduct As Double
    If isReg Then
        basicDeduct = 4000000
    Else
        basicDeduct = 2000000
    End If

    ' 분리과세 계산
    Dim sepBase As Double
    Dim sepTaxAmt As Double
    sepBase = totalRent * (1 - expenseRate) - basicDeduct
    If sepBase < 0 Then sepBase = 0
    sepTaxAmt = sepBase * 0.14

    ' 종합과세 계산
    Dim compBase As Double
    Dim compRate As Double
    Dim compDed As Double
    compBase = totalRent * (1 - expenseRate) + otherIncome
    compRate = 0
    compDed = 0
    Call GetProgressiveTax(compBase, compRate, compDed)

    Dim compTaxTotal As Double
    compTaxTotal = compBase * compRate - compDed

    Dim otherTax As Double
    otherTax = 0
    If otherIncome > 0 Then
        Dim otherRate As Double
        Dim otherDed As Double
        otherRate = 0
        otherDed = 0
        Call GetProgressiveTax(otherIncome, otherRate, otherDed)
        otherTax = otherIncome * otherRate - otherDed
    End If

    Dim compTaxOnRent As Double
    compTaxOnRent = compTaxTotal - otherTax
    If compTaxOnRent < 0 Then compTaxOnRent = 0

    ' 지방소득세
    Dim sepLocalTax As Double
    Dim compLocalTax As Double
    sepLocalTax = sepTaxAmt * 0.1
    compLocalTax = compTaxOnRent * 0.1

    ' 출력
    ws.Range("G7").Value = monthlyRent
    ws.Range("G8").Value = ganjuRent
    ws.Range("G9").Value = totalRent
    ws.Range("G10").Value = expenseRate
    ws.Range("G11").Value = sepBase
    ws.Range("G12").Value = sepTaxAmt
    ws.Range("G13").Value = sepLocalTax
    ws.Range("G14").Value = sepTaxAmt + sepLocalTax
    ws.Range("G15").Value = compBase
    ws.Range("G16").Value = compTaxOnRent
    ws.Range("G17").Value = compLocalTax
    ws.Range("G18").Value = compTaxOnRent + compLocalTax

    Dim sepTotal As Double
    Dim compTotal As Double
    sepTotal = sepTaxAmt + sepLocalTax
    compTotal = compTaxOnRent + compLocalTax

    Dim recommend As String
    If totalRent <= 20000000 Then
        If sepTotal < compTotal Then
            recommend = "★ 분리과세가 유리합니다! (절세: " & Format(compTotal - sepTotal, "#,##0") & "원)"
        Else
            recommend = "★ 종합과세가 유리합니다! (절세: " & Format(sepTotal - compTotal, "#,##0") & "원)"
        End If
    Else
        recommend = "총 임대수입 2,000만원 초과 - 종합과세 의무 적용"
    End If

    ShowInfo "계산이 완료되었습니다!" & Chr(13) & Chr(13) & _
             "▶ 총 임대수입(간주임대료 포함): " & Format(totalRent, "#,##0") & "원" & Chr(13) & _
             "▶ 간주임대료: " & Format(ganjuRent, "#,##0") & "원" & Chr(13) & Chr(13) & _
             "[ 분리과세 선택 시 ] " & Format(sepTotal, "#,##0") & "원" & Chr(13) & _
             "[ 종합과세 선택 시 ] " & Format(compTotal, "#,##0") & "원" & Chr(13) & Chr(13) & _
             recommend
End Sub

Sub ResetImde()
    Sheets("임대소득세").Range("C7:C12").ClearContents
    Sheets("임대소득세").Range("G7:G18").ClearContents
    Sheets("임대소득세").Range("C9").Value = 1
    Sheets("임대소득세").Range("C11").Value = "아니오"
End Sub

' ===========================================================
' 전체 초기화
' ===========================================================
Sub ResetAll()
    Dim ans As Integer
    ans = MsgBox("모든 시트의 입력값과 계산 결과를 초기화하시겠습니까?", _
                 vbYesNo + vbQuestion, "초기화 확인")
    If ans = vbYes Then
        Call ResetYangdo
        Call ResetJongbu
        Call ResetJeungYeo
        Call ResetSangsok
        Call ResetImde
        ShowInfo "모든 데이터가 초기화되었습니다."
    End If
End Sub
