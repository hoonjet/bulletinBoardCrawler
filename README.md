# bulletinBoardCrawler
Simple webcrawler example coded with VBA: Programmed with Excel 2016, tested with FireFox 55.0.3 & Windows 10

- Crawl web information (Bulletin boards), performing statistical analysis 
- Designed for bullentin boards in www.munpia.com, but could be easily modified for other sites
- Core codes (web crawlin part)

        'Use XmlHttpRequest
        'XmlHttpRequest object is used to make HTTP requests in VBA
        'please refer: https://codingislove.com/http-requests-excel-vba/

        targetURL = target & "/page/" & i 'URL format for targetted webpage (i is page number in For Loop)        
        Set XMLHTTP = CreateObject("MSXML2.serverXMLHTTP")
        XMLHTTP.Open "GET", targetURL, False
        XMLHTTP.setRequestHeader "Content-Type", "text/xml" 'Crawl data (source code)
        XMLHTTP.send            
        crawledText = XMLHTTP.ResponseText 'Total information (Crawled from website)
        


- File infos
   1. Crawler.xlsm : Excel file with VBA macro and other documents (VBA must be enabled)
   2. WebCrawler.bas : Exported basic file
