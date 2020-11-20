<%
    ' 배열 검색
    ' @param    Array   source  검색 목록
    ' @param    Object  target  검색 대상
    ' @return   Boolean         검색 대상 존재 여부
    Function InArray(source, target)
        dim i, isFind
        isFind = False
        For i = 0 to UBound(source)
            If source(i) = target Then
                isFind = True
                Exit For
            End If
        Next
        InArray = isFind
    End Function
%>