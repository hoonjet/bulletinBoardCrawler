Attribute VB_Name = "Module1"


Sub crawlData()

    Dim target As String
    Dim targetURL As String
    Dim count As Integer
    Dim offset As Integer
    Dim isEnd As Boolean
    
    
    offset = 10 'number rows of offsetting data
    
    target = Worksheets("SettingPage").Cells(1, 1).Value 'Formatting target URL
    pageNum = Worksheets("SettingPage").Cells(2, 1).Value
    
    'Notice
    noticeTextStart = Worksheets("SettingPage").Cells(3, 1).Value
    noticeTextEnd = Worksheets("SettingPage").Cells(3, 2).Value
    
    'Block of text required to crawl per article
    subjectTextStart = Worksheets("SettingPage").Cells(4, 1).Value
    subjectTextEnd = Worksheets("SettingPage").Cells(4, 2).Value
    
    'Index
    indexTextStart = Worksheets("SettingPage").Cells(5, 1).Value
    indexTextEnd = Worksheets("SettingPage").Cells(5, 2).Value
    
    'article title
    titleTextStart = Worksheets("SettingPage").Cells(6, 1).Value
    titleTextMid = Worksheets("SettingPage").Cells(6, 2).Value
    titleTextEnd = Worksheets("SettingPage").Cells(6, 3).Value
    
   'number of replies
    replyTextStart = Worksheets("SettingPage").Cells(7, 1).Value
    replyTextEnd = Worksheets("SettingPage").Cells(7, 2).Value
    
    'date
    dateTextStart = Worksheets("SettingPage").Cells(8, 1).Value
    dateTextEnd = Worksheets("SettingPage").Cells(8, 2).Value
    
    'view
    viewTextStart = Worksheets("SettingPage").Cells(9, 1).Value
    viewTextEnd = Worksheets("SettingPage").Cells(9, 2).Value
    
    'recommendation
    recommTextStart = Worksheets("SettingPage").Cells(10, 1).Value
    recommTextEnd = Worksheets("SettingPage").Cells(10, 2).Value
        
    
    'Title (of book)
    nameTextStart = Worksheets("SettingPage").Cells(11, 1).Value
    nameTextEnd = Worksheets("SettingPage").Cells(11, 2).Value
        
        
    'favorite
    favorTextStart = Worksheets("SettingPage").Cells(12, 1).Value
    favorTextEnd = Worksheets("SettingPage").Cells(12, 2).Value
    
    count = 1
    
    'Make a new with name of current date and time
    
    SheetName = Format(Now(), "yymmdd_hhmmss")
    Sheets("Format").Copy After:=Sheets(3)
    Sheets(4).Name = SheetName
    
    i = 1 'Page Counter
    
    Do While isEnd = False
              
    'Crawling data
    
          
        targetURL = target & "/page/" & i 'URL to crop (just for 1st page)
        Set XMLHTTP = CreateObject("MSXML2.serverXMLHTTP")
        XMLHTTP.Open "GET", targetURL, False
        XMLHTTP.setRequestHeader "Content-Type", "text/xml"
        XMLHTTP.send
            
        crawledText = XMLHTTP.ResponseText 'Total information (Crawled from website)
               
        nameIndexStart = InStr(crawledText, nameTextStart) 'Get the favorite information first
        nameIndexEnd = InStr(crawledText, nameTextEnd)
        nameInfo = Mid(crawledText, nameIndexStart + Len(nameTextStart), nameIndexEnd - nameIndexStart - Len(nameTextStart))
        
        favorIndexStart = InStr(crawledText, favorTextStart) 'Get the favorite information first
        favorIndexEnd = InStr(crawledText, favorTextEnd)
        favorInfo = Mid(crawledText, favorIndexStart + Len(favorTextStart), favorIndexEnd - favorIndexStart - Len(favorTextStart))

        
        Do While InStr(crawledText, noticeTextStart) <> 0 'Crop out unncessary part (e.g. Notice)
            noticeIndexEnd = InStr(crawledText, noticeTextEnd)
            crawledText = Right(crawledText, Len(crawledText) - noticeIndexEnd - Len(noticeTextEnd))
        Loop
        
        
        
        Do While InStr(crawledText, subjectTextStart) <> 0
            subjectIndexStart = InStr(crawledText, subjectTextStart)
            subjectIndexEnd = InStr(crawledText, subjectTextEnd)
            
            'Crop out information needed
            workingText = Mid(crawledText, subjectIndexStart + Len(subjectTextStart), subjectIndexEnd - subjectIndexStart - Len(subjectTextStart))
           
            
            'Find index numbers
            indexIndexStart = InStr(workingText, indexTextStart)
            indexIndexEnd = InStr(workingText, indexTextEnd)
            indexInfo = Mid(workingText, indexIndexStart + Len(indexTextStart), indexIndexEnd - indexIndexStart - Len(indexTextStart))
            Worksheets(SheetName).Cells(count + offset, 1).Value = indexInfo
            If indexInfo = 1 Then 'if reached to the end of the list (index = 1)
                isEnd = True
            End If
            workingText = Mid(workingText, indexIndexEnd + Len(indexTextEnd), Len(workingText))
            
            
            
            'Find Titles
            titleIndexStart = InStr(workingText, titleTextStart)
            titleIndexEnd = InStr(workingText, titleTextEnd)
            titleSubText = Mid(workingText, titleIndexStart + Len(titleTextStart), titleIndexEnd - titleIndexStart - Len(titleTextStart))
            titleIndexMid = InStr(titleSubText, titleTextMid)
            titleInfo = Mid(titleSubText, titleIndexMid + Len(titleTextMid), Len(titleSubText))
            Worksheets(SheetName).Cells(count + offset, 2).Value = titleInfo
            workingText = Mid(workingText, titleIndexEnd + Len(titleTextEnd), Len(workingText))
            
            
            'Find number of replies
            replyIndexStart = InStr(workingText, replyTextStart)
            If replyIndexStart <> 0 Then
                replyIndexEnd = InStr(workingText, replyTextEnd)
                replyInfo = Mid(workingText, replyIndexStart + Len(replyTextStart), replyIndexEnd - replyIndexStart - Len(replyTextStart))
                Worksheets(SheetName).Cells(count + offset, 3).Value = replyInfo
                workingText = Mid(workingText, replyIndexEnd + Len(replyTextEnd), Len(workingText))
            Else 'In case when there is no reply
                Worksheets(SheetName).Cells(count + offset, 3).Value = 0
            End If
            
            'Find date information
            dateIndexStart = InStr(workingText, dateTextStart)
            workingText = Mid(workingText, dateIndexStart, Len(workingText))
            dateIndexStart = InStr(workingText, dateTextStart)
            dateIndexEnd = InStr(workingText, dateTextEnd)
            dateInfo = Mid(workingText, dateIndexStart + Len(dateTextStart), dateIndexEnd - dateIndexStart - Len(dateTextStart))
            
                'Change string to date
                If InStr(dateInfo, ".") = 0 Then
                    dateInfo = Date 'if uploaded within 24 hours, consider it as 1 day
                Else
                    numYear = Left(dateInfo, 2) + 2000
                    numMonth = Mid(dateInfo, 4, 2)
                    numDay = Right(dateInfo, 2)
                    dateInfo = DateSerial(numYear, numMonth, numDay)
                End If
                           
            Worksheets(SheetName).Cells(count + offset, 4).Value = dateInfo
            workingText = Mid(workingText, dateIndexEnd + Len(dateTextEnd), Len(workingText))
           
            'Find number of views information
            viewIndexStart = InStr(workingText, viewTextStart)
            workingText = Mid(workingText, viewIndexStart, Len(workingText))
            viewIndexStart = InStr(workingText, viewTextStart)
            viewIndexEnd = InStr(workingText, viewTextEnd)
            viewInfo = Mid(workingText, viewIndexStart + Len(viewTextStart), viewIndexEnd - viewIndexStart - Len(viewTextStart))
            viewInfo = Right(viewInfo, Len(viewInfo) - 1) 'remove enter
            Worksheets(SheetName).Cells(count + offset, 5).Value = viewInfo
            workingText = Mid(workingText, viewIndexEnd + Len(viewTextEnd), Len(workingText))


            'Find number of recommendations
            recommIndexStart = InStr(workingText, recommTextStart)
            workingText = Mid(workingText, recommIndexStart, Len(workingText))
            recommIndexStart = InStr(workingText, recommTextStart)
            recommIndexEnd = InStr(workingText, recommTextEnd)
            recommInfo = Mid(workingText, recommIndexStart + Len(recommTextStart), recommIndexEnd - recommIndexStart - Len(recommTextStart))
            Worksheets(SheetName).Cells(count + offset, 6).Value = recommInfo
            workingText = Mid(workingText, recommIndexEnd + Len(recommTextEnd), Len(workingText))

                                    
            crawledText = Right(crawledText, Len(crawledText) - subjectIndexEnd)
            count = count + 1
        Loop
        i = i + 1
    Loop
   
    
    
    'Performing calculation
    For i = 1 To count - 1
        
        'number of recommendation per views
        Worksheets(SheetName).Cells(i + offset, 8).Value = Worksheets(SheetName).Cells(i + offset, 6).Value / Worksheets(SheetName).Cells(i + offset, 5).Value
        Worksheets(SheetName).Columns(8).NumberFormat = "0.0%"
        'number of replies per views
        Worksheets(SheetName).Cells(i + offset, 9).Value = Worksheets(SheetName).Cells(i + offset, 3).Value / Worksheets(SheetName).Cells(i + offset, 5).Value
        Worksheets(SheetName).Columns(9).NumberFormat = "#,###0.000"
               
        
        'Date posted and Views per date posted
        Worksheets(SheetName).Cells(i + offset, 10).Value = Date - Worksheets(SheetName).Cells(i + offset, 4).Value
        
        If Worksheets(SheetName).Cells(i + offset, 10).Value <> 0 Then
           Worksheets(SheetName).Cells(i + offset, 11).Value = Worksheets(SheetName).Cells(i + offset, 5).Value / Worksheets(SheetName).Cells(i + offset, 10).Value
        Else
           Worksheets(SheetName).Cells(i + offset, 11).Value = 0
        End If
        Worksheets(SheetName).Columns(13).NumberFormat = "#,#0.0"
        
        'Number of views / number of views of the 1st article
        Worksheets(SheetName).Cells(i + offset, 12).Value = Worksheets(SheetName).Cells(i + offset, 5).Value / Worksheets(SheetName).Cells(count + offset - 1, 5).Value
        Worksheets(SheetName).Columns(12).NumberFormat = "0.00%"
        
        'Number of views / number of views of previous article
        If i <> count - 1 Then
        Worksheets(SheetName).Cells(i + offset, 13).Value = Worksheets(SheetName).Cells(i + offset, 5).Value / Worksheets(SheetName).Cells(i + offset + 1, 5).Value
        Else
        Worksheets(SheetName).Cells(i + offset, 13).Value = 1
        End If
        Worksheets(SheetName).Columns(13).NumberFormat = "0.0%"
        
    
    Next i
    
   'Overall calculation
    Worksheets(SheetName).Range("B1").Value = nameInfo ' Title
    Worksheets(SheetName).Range("B2").Value = count - 1 ' Number of articles
    Worksheets(SheetName).Range("B2").NumberFormat = "#0"
    Worksheets(SheetName).Range("D2").Value = Date ' Date of analysis
    
    Worksheets(SheetName).Range("B3").Value = Application.Sum(Worksheets(SheetName).Range(Cells(offset + 1, 5), Worksheets(SheetName).Cells(count + offset, 5))) ' Total View
    Worksheets(SheetName).Range("B3").NumberFormat = "#0"
    Worksheets(SheetName).Range("D3").Value = Application.Average(Worksheets(SheetName).Range(Cells(offset + 1, 5), Worksheets(SheetName).Cells(count + offset, 5))) ' Average View
    Worksheets(SheetName).Range("D3").NumberFormat = "#,#0.0"
      
    Worksheets(SheetName).Range("B4").Value = Application.Sum(Range(Worksheets(SheetName).Cells(offset + 1, 6), Worksheets(SheetName).Cells(count + offset, 6))) ' Total recommendation
    Worksheets(SheetName).Range("B4").NumberFormat = "#0"
    Worksheets(SheetName).Range("D4").Value = Application.Average(Worksheets(SheetName).Range(Cells(offset + 1, 6), Worksheets(SheetName).Cells(count + offset, 6))) ' Average recommendation
    Worksheets(SheetName).Range("D4").NumberFormat = "#,#0.0"
    
    
    Worksheets(SheetName).Range("B5").Value = Date - Worksheets(SheetName).Cells(count + offset - 1, 4).Value 'Days since 1st posting
    Worksheets(SheetName).Range("B5").NumberFormat = "#0"
    If Worksheets(SheetName).Range("B5").Value <> 0 Then
        Worksheets(SheetName).Range("D5").Value = Worksheets(SheetName).Range("B2").Value / Worksheets(SheetName).Range("B5").Value 'Average uploads per day
        Else
        Worksheets(SheetName).Range("D5").Value = 0
    End If
    
    Worksheets(SheetName).Range("D5").NumberFormat = "#,##0.00"
        
    
    Worksheets(SheetName).Range("B6").Value = Worksheets(SheetName).Range("B4").Value / Worksheets(SheetName).Range("B3").Value   'recommdation rate (recommdnation/view)
    Worksheets(SheetName).Range("B6").NumberFormat = "0.00%"
    Worksheets(SheetName).Range("B7").Value = favorInfo 'Favorite
    Worksheets(SheetName).Range("B7").NumberFormat = "#0"
    
    
    'Rename Worksheet Name
    Worksheets(SheetName).Name = nameInfo & "_" & SheetName
    
    
    'Make overall summary file
    Dim numWorks As Integer
    i = 0
    summaryOffsets = 9 'Number of columns required in Summary sheet per work +1
    
    Do While Not (IsEmpty(Worksheets("summary").Cells(2, 2 + i * summaryOffsets).Value))
        i = i + 1
    Loop
    
    numWorks = i
    
    
    Dim isNew As Boolean
    
    
    i = 4 'Only add summary for the
            
        If numWorks = 0 Then 'first summary chart
            Worksheets("Summary").Cells(2, 2).Value = Sheets(i).Cells(1, 2).Value
            Worksheets("Summary").Cells(5, 1).Value = Sheets(i).Cells(2, 4).Value
            Worksheets("Summary").Cells(5, 2).Value = Sheets(i).Cells(2, 2).Value
            Worksheets("Summary").Cells(5, 3).Value = Sheets(i).Cells(3, 2).Value
            Worksheets("Summary").Cells(5, 4).Value = Sheets(i).Cells(4, 2).Value
            Worksheets("Summary").Cells(5, 5).Value = Sheets(i).Cells(3, 4).Value
            Worksheets("Summary").Columns(5).NumberFormat = "#,##0.00"
            Worksheets("Summary").Cells(5, 6).Value = Sheets(i).Cells(4, 4).Value
            Worksheets("Summary").Columns(6).NumberFormat = "#,##0.00"
            Worksheets("Summary").Cells(5, 7).Value = Sheets(i).Cells(7, 2).Value
            
            
            numWorks = numWorks + 1
            
        Else
            
            isNew = True
            Dim titleIndex As Integer
            titleIndex = 1
                        
            For j = 1 To numWorks
                
                If Worksheets("Summary").Cells(2, (j - 1) * summaryOffsets + 2).Value = Sheets(i).Cells(1, 2).Value Then
                    isNew = False
                    titleIndex = j
                End If
                
            Next j
                
            If isNew Then
            
                    numWorks = numWorks + 1
                    'Make new format
                    Worksheets("Summary").Cells(2, 1 + (numWorks - 1) * summaryOffsets).Value = Worksheets("Summary").Cells(2, 1).Value
                    Worksheets("Summary").Cells(4, 1 + (numWorks - 1) * summaryOffsets).Value = Worksheets("Summary").Cells(4, 1).Value
                    Worksheets("Summary").Cells(4, 2 + (numWorks - 1) * summaryOffsets).Value = Worksheets("Summary").Cells(4, 2).Value
                    Worksheets("Summary").Cells(4, 3 + (numWorks - 1) * summaryOffsets).Value = Worksheets("Summary").Cells(4, 3).Value
                    Worksheets("Summary").Cells(4, 4 + (numWorks - 1) * summaryOffsets).Value = Worksheets("Summary").Cells(4, 4).Value
                    Worksheets("Summary").Cells(4, 5 + (numWorks - 1) * summaryOffsets).Value = Worksheets("Summary").Cells(4, 5).Value
                    Worksheets("Summary").Cells(4, 6 + (numWorks - 1) * summaryOffsets).Value = Worksheets("Summary").Cells(4, 6).Value
                    Worksheets("Summary").Cells(4, 7 + (numWorks - 1) * summaryOffsets).Value = Worksheets("Summary").Cells(4, 7).Value
                                        
                    Worksheets("Summary").Cells(2, 2 + (numWorks - 1) * summaryOffsets).Value = Sheets(i).Cells(1, 2).Value
                    Worksheets("Summary").Cells(5, 1 + (numWorks - 1) * summaryOffsets).Value = Sheets(i).Cells(2, 4).Value
                    Worksheets("Summary").Cells(5, 2 + (numWorks - 1) * summaryOffsets).Value = Sheets(i).Cells(2, 2).Value
                    Worksheets("Summary").Cells(5, 3 + (numWorks - 1) * summaryOffsets).Value = Sheets(i).Cells(3, 2).Value
                    Worksheets("Summary").Cells(5, 4 + (numWorks - 1) * summaryOffsets).Value = Sheets(i).Cells(4, 2).Value
                    Worksheets("Summary").Cells(5, 5 + (numWorks - 1) * summaryOffsets).Value = Sheets(i).Cells(3, 4).Value
                    Worksheets("Summary").Cells(5, 6 + (numWorks - 1) * summaryOffsets).Value = Sheets(i).Cells(4, 4).Value
                    Worksheets("Summary").Cells(5, 7 + (numWorks - 1) * summaryOffsets).Value = Sheets(i).Cells(7, 2).Value
                    
                    Worksheets("Summary").Columns(5 + (numWorks - 1)).NumberFormat = "#,##0.00"
                    Worksheets("Summary").Columns(6 + (numWorks - 1)).NumberFormat = "#,##0.00"
                    
            Else
                                  
                'Existing work
                Dim rng As Range ' insert a new row
                Set rng = Worksheets("Summary").Range(Worksheets("Summary").Cells(5, (titleIndex - 1) * summaryOffsets + 1), Worksheets("Summary").Cells(5, (titleIndex - 1) * summaryOffsets + summaryOffsets))
                rng.Insert Shift:=xlDown
                
                Worksheets("Summary").Cells(2, 2 + (titleIndex - 1) * summaryOffsets).Value = Sheets(i).Cells(1, 2).Value
                Worksheets("Summary").Cells(5, 1 + (titleIndex - 1) * summaryOffsets).Value = Sheets(i).Cells(2, 4).Value
                Worksheets("Summary").Cells(5, 2 + (titleIndex - 1) * summaryOffsets).Value = Sheets(i).Cells(2, 2).Value
                Worksheets("Summary").Cells(5, 3 + (titleIndex - 1) * summaryOffsets).Value = Sheets(i).Cells(3, 2).Value
                Worksheets("Summary").Cells(5, 4 + (titleIndex - 1) * summaryOffsets).Value = Sheets(i).Cells(4, 2).Value
                Worksheets("Summary").Cells(5, 5 + (titleIndex - 1) * summaryOffsets).Value = Sheets(i).Cells(3, 4).Value
                Worksheets("Summary").Cells(5, 6 + (titleIndex - 1) * summaryOffsets).Value = Sheets(i).Cells(4, 4).Value
                Worksheets("Summary").Cells(5, 7 + (titleIndex - 1) * summaryOffsets).Value = Sheets(i).Cells(7, 2).Value
                
               
                                    
                   
                    
            End If
            
        
        End If
      Worksheets("Summary").Columns(5 + (titleIndex - 1)).NumberFormat = "#,#0.0"
      Worksheets("Summary").Columns(6 + (titleIndex - 1)).NumberFormat = "#,#0.0"
    
    

End Sub
