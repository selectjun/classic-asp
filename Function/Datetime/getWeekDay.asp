<%
    ' 요일을 반환(한국어)
    ' @param    day 요일 숫자 정보
    ' @return       요일(한국어)
    Function getWeekDay(day)
        Dim weekDay : weekDay = ""
        Select Case weekDay(day)
            case "1" weekDay = "일"
            case "2" weekDay = "월"
            case "3" weekDay = "화"
            case "4" weekDay = "수"
            case "5" weekDay = "목"
            case "6" weekDay = "금"
            case "7" weekDay = "토"
        End Select
        getWeekDay = weekDate
    End Function
%>